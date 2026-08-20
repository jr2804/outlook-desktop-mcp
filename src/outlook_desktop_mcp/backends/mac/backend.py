"""macOS AppleScript backend — automates Microsoft Outlook for Mac via osascript.

All methods build AppleScript, run it through the bridge, and return pydantic
models. Handled failures raise :class:`BackendError`.

Signature parity notes:
- ``account`` is accepted and ignored (AppleScript always targets the
  currently active account).
- ``list_folders`` supports the drill-down model on a best-effort basis:
  Outlook's AppleScript dictionary exposes ``mail folders`` as a flat
  collection without parent links, so a ``folder`` argument filters by the
  last path segment and ``max_depth`` beyond 1 cannot be honored.
"""

from __future__ import annotations

import contextlib
import logging
import os
import re
from datetime import datetime, timedelta

from outlook_desktop_mcp.backends.base import Backend, BackendError
from outlook_desktop_mcp.backends.mac.applescript_helpers import (
    DELIM,
    RECORD_DELIM,
    escape,
    format_date,
    resolve_folder_ref,
)
from outlook_desktop_mcp.backends.mac.bridge import AppleScriptBridge
from outlook_desktop_mcp.models import (
    AttachmentInfo,
    AttachmentSavedResult,
    DraftSavedResult,
    EmailFull,
    EmailSummary,
    EventCreatedResult,
    EventFull,
    EventSummary,
    EventUpdatedResult,
    FolderInfo,
    ItemStatusResult,
    MeetingSentResult,
    MovedResult,
    OofSetResult,
    ReplyDraftSavedResult,
    ReplySentResult,
    SentResult,
    TaskCreatedResult,
    TaskFull,
    TaskSummary,
)

logger = logging.getLogger("outlook_desktop_mcp.backends.mac.backend")

_UI_MESSAGE_LIST_PATH = (
    'tell application "System Events"\n'
    '    tell process "Microsoft Outlook"\n'
    "        tell window 1\n"
    "            tell splitter group 1\n"
    "                tell splitter group 1\n"
    "                    tell splitter group 1\n"
    "                        tell group 1\n"
    "                            tell scroll area 1\n"
    "                                tell table 1\n"
)

_UI_MESSAGE_LIST_END = (
    "                                end tell\n"
    "                            end tell\n"
    "                        end tell\n"
    "                    end tell\n"
    "                end tell\n"
    "            end tell\n"
    "        end tell\n"
    "    end tell\n"
    "end tell"
)

# Locale-dependent status tokens for UI scraping
_UNREAD_TOKENS = {"Ulest", "Unread", "Non lu", "Nicht gelesen", "No leído", "未読", "未读"}
_ATTACHMENT_TOKENS = {
    "Har filer",
    "Has attachments",
    "Contient des fichiers",
    "Hat Anlagen",
    "Tiene archivos adjuntos",
    "添付ファイルあり",
    "有附件",
}
_SKIP_PREFIXES = ("Merket som", "Marked as", "Marqué comme", "Markiert als", "Marcado como", "A ")
_CATEGORY_TOKENS = {"Kategorisert", "Categorized", "Catégorisé", "Kategorisiert", "Categorizado"}


def _truncate(text: str, max_length: int = 5000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def _clean(value: str) -> str:
    """Replace AppleScript's 'missing value' with empty string."""
    v = value.strip()
    return "" if v == "missing value" else v


def _parse_iso(date_str: str) -> datetime:
    return datetime.fromisoformat(date_str)


class AppleScriptBackend(Backend):
    """Outlook for Mac, driven through osascript."""

    name = "applescript"

    def __init__(self) -> None:
        self.bridge = AppleScriptBridge()

    def start(self) -> None:
        import asyncio  # noqa: PLC0415

        asyncio.run(self.bridge.start())

    def stop(self) -> None:
        import asyncio  # noqa: PLC0415

        asyncio.run(self.bridge.stop())

    # --- email ---

    async def compose_email(
        self,
        to: str,
        subject: str,
        body: str,
        cc: str,
        bcc: str,
        html_body: str,
        account: str,
        send: bool,
    ) -> SentResult | DraftSavedResult:
        to_lines = self._recipient_lines(to, "to recipient")
        cc_lines = self._recipient_lines(cc, "cc recipient") if cc else ""
        bcc_lines = self._recipient_lines(bcc, "bcc recipient") if bcc else ""

        content_prop = f'html content:"{escape(html_body)}"' if html_body else f'content:"{escape(body)}"'
        action = "send newMsg" if send else 'save newMsg in folder "Drafts"'

        script = f'''tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {{subject:"{escape(subject)}", {content_prop}}}
    {to_lines}{cc_lines}{bcc_lines}
    {action}
    return id of newMsg as text
end tell'''

        raw = await self.bridge.run(script)
        if send:
            return SentResult(subject=subject, to=to)
        return DraftSavedResult(subject=subject, to=to, entry_id=raw)

    async def list_emails(
        self,
        folder: str,
        count: int,
        unread_only: bool,
        start_date: str,
        end_date: str,
        account: str,
    ) -> list[EmailSummary]:
        folder_ref = resolve_folder_ref(folder)
        unread_filter = " whose is read is false" if unread_only else ""
        date_filter = ""
        if start_date:
            start = _parse_iso(start_date)
            date_filter = f" whose time received >= {format_date(start)}"
            if end_date:
                end = _parse_iso(end_date)
                date_filter = f" whose time received >= {format_date(start)} and time received <= {format_date(end)}"
        elif end_date:
            end = _parse_iso(end_date)
            date_filter = f" whose time received <= {format_date(end)}"
        # Combine filters (only one whose-clause allowed; date filter wins, read
        # status is re-checked below if needed)
        whose = date_filter or unread_filter

        script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set allMsgs to messages of folderRef{whose}
    set msgCount to count of allMsgs
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    set output to ""
    repeat with i from 1 to maxCount
        set m to item i of allMsgs
        set mid to id of m
        set msubject to subject of m
        set msender to ""
        try
            set msender to address of sender of m
        end try
        set msenderName to ""
        try
            set msenderName to name of sender of m
        end try
        set mtime to time received of m as string
        set misread to is read of m
        set mattcount to 0
        try
            set mattcount to count of attachments of m
        end try
        set output to output & (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

        raw = await self.bridge.run(script)
        results = self._parse_email_summaries(raw)
        if unread_only and date_filter:
            results = [r for r in results if r.unread]

        # Fallback: New Outlook for Mac keeps Exchange messages outside the
        # AppleScript-visible mailbox. If nothing was found for the inbox,
        # read the visible message list via UI scripting.
        if not results and folder.lower().strip() == "inbox":
            with contextlib.suppress(Exception):
                results = await self._ui_list_messages(count)

        return results

    async def read_email(
        self,
        entry_id: str,
        subject_search: str,
        folder: str,
        account: str,
    ) -> EmailFull:
        if entry_id:
            script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    {self._email_fields_script("m")}
    return (id of m as text) & "{DELIM}" & (subject of m) & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{DELIM}" & mto & "{DELIM}" & mcc & "{DELIM}" & mbody
end tell'''
        elif subject_search:
            folder_ref = resolve_folder_ref(folder)
            safe_query = escape(subject_search)
            script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set matchMsgs to messages of folderRef whose subject contains "{safe_query}"
    if (count of matchMsgs) = 0 then return "NOT_FOUND"
    set m to item 1 of matchMsgs
    {self._email_fields_script("m")}
    return (id of m as text) & "{DELIM}" & (subject of m) & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{DELIM}" & mto & "{DELIM}" & mcc & "{DELIM}" & mbody
end tell'''
        else:
            raise BackendError("Provide either entry_id or subject_search")

        raw = await self.bridge.run(script)
        if raw == "NOT_FOUND":
            raise BackendError(f"No email found matching '{subject_search}'")

        parts = raw.split(DELIM, 9)
        if len(parts) < 10:
            raise BackendError("Failed to parse email data")

        att_count = int(parts[6].strip()) if parts[6].strip().isdigit() else 0
        return EmailFull(
            entry_id=parts[0].strip(),
            subject=parts[1].strip() or "(no subject)",
            sender=parts[2].strip(),
            sender_name=parts[3].strip(),
            received_time=_clean(parts[4]),
            unread=parts[5].strip().lower() != "true",
            has_attachments=att_count > 0,
            attachment_count=att_count,
            to=parts[7].strip(),
            cc=parts[8].strip(),
            body=_truncate(_clean(parts[9])),
        )

    async def mark_as_read(self, entry_id: str, account: str) -> ItemStatusResult:
        script = f"""tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set is read of m to true
    return subject of m
end tell"""
        subject = await self.bridge.run(script)
        return ItemStatusResult(status="marked_read", subject=subject, entry_id=entry_id)

    async def mark_as_unread(self, entry_id: str, account: str) -> ItemStatusResult:
        script = f"""tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set is read of m to false
    return subject of m
end tell"""
        subject = await self.bridge.run(script)
        return ItemStatusResult(status="marked_unread", subject=subject, entry_id=entry_id)

    async def move_email(self, entry_id: str, target_folder: str, account: str) -> MovedResult:
        dest_ref = resolve_folder_ref(target_folder)
        script = f"""tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set msubject to subject of m
    move m to {dest_ref}
    return msubject
end tell"""
        subject = await self.bridge.run(script)
        return MovedResult(subject=subject, target_folder=target_folder)

    async def reply_email(
        self,
        entry_id: str,
        body: str,
        reply_all: bool,
        html_body: str,
        account: str,
        send: bool,
    ) -> ReplySentResult | ReplyDraftSavedResult:
        reply_cmd = "reply all to" if reply_all else "reply to"
        action = "send replyMsg" if send else 'save replyMsg in folder "Drafts"'
        if html_body:
            set_content = f'set html content of replyMsg to "{escape(html_body)}" & return & return & html content of replyMsg'
        else:
            set_content = f'set content of replyMsg to "{escape(body)}" & return & return & content of replyMsg'
        script = f"""tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set msubject to subject of m
    set replyMsg to {reply_cmd} m
    {set_content}
    {action}
    return msubject
end tell"""
        subject = await self.bridge.run(script)
        if send:
            return ReplySentResult(subject=subject, reply_all=reply_all)
        return ReplyDraftSavedResult(subject=subject, reply_all=reply_all)

    async def list_folders(self, folder: str, max_depth: int, account: str) -> list[FolderInfo]:
        script = f'''tell application "Microsoft Outlook"
    set allFolders to mail folders
    set output to ""
    repeat with f in allFolders
        set fname to name of f
        set fcount to count of messages of f
        set funread to unread count of f
        set output to output & fname & "{DELIM}" & (fcount as text) & "{DELIM}" & (funread as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

        raw = await self.bridge.run(script)
        results: list[FolderInfo] = []
        for raw_record in (raw or "").split(RECORD_DELIM):
            record = raw_record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 3:
                continue
            results.append(
                FolderInfo(
                    name=parts[0].strip(),
                    full_path=parts[0].strip(),
                    item_count=int(parts[1].strip()) if parts[1].strip().isdigit() else 0,
                    unread_count=int(parts[2].strip()) if parts[2].strip().isdigit() else 0,
                )
            )

        # Best-effort drill-down: filter the flat collection by the last path
        # segment (the dictionary exposes no parent links for real traversal).
        if folder:
            last_segment = folder.strip().split("/")[-1].lower()
            filtered = [f for f in results if f.name.lower() == last_segment]
            if not filtered:
                raise BackendError(f"Folder '{folder}' not found")
            for f in filtered:
                f.full_path = folder
            return filtered

        return results

    async def search_emails(
        self,
        query: str,
        folder: str,
        count: int,
        start_date: str,
        end_date: str,
        account: str,
    ) -> list[EmailSummary]:
        folder_ref = resolve_folder_ref(folder)
        safe_query = escape(query)

        script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set matchMsgs to messages of folderRef whose subject contains "{safe_query}"
    set msgCount to count of matchMsgs
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    set output to ""
    repeat with i from 1 to maxCount
        set m to item i of matchMsgs
        set mid to id of m
        set msubject to subject of m
        set msender to ""
        try
            set msender to address of sender of m
        end try
        set msenderName to ""
        try
            set msenderName to name of sender of m
        end try
        set mtime to time received of m as string
        set misread to is read of m
        set mattcount to 0
        try
            set mattcount to count of attachments of m
        end try
        set output to output & (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

        raw = await self.bridge.run(script)
        results = self._parse_email_summaries(raw)

        # Date filtering in Python (AppleScript whose-clauses with dates and
        # subject conditions combined are unreliable in Outlook for Mac).
        if start_date or end_date:
            start = _parse_iso(start_date) if start_date else None
            end = _parse_iso(end_date) if end_date else (datetime.now() if start_date else None)
            filtered = []
            for r in results:
                iso = self._to_iso(r.received_time)
                if iso is None:
                    filtered.append(r)  # keep unparseable timestamps
                    continue
                dt = datetime.fromisoformat(iso)
                if start and dt < start:
                    continue
                if end and dt > end:
                    continue
                filtered.append(r)
            results = filtered

        return results

    # --- calendar ---

    async def list_events(
        self,
        start_date: str,
        end_date: str,
        count: int,
        account: str,
        folder: str,
    ) -> list[EventSummary]:
        start = _parse_iso(start_date) if start_date else datetime.now()
        end = _parse_iso(end_date) if end_date else start + timedelta(days=7)

        # Fetch more than needed, filter by date in Python since AppleScript
        # whose-clause date filtering can be unreliable in Outlook for Mac.
        fetch_limit = count * 3

        script = f'''tell application "Microsoft Outlook"
    set evts to calendar events
    set evtCount to count of evts
    set maxFetch to {fetch_limit}
    if evtCount < maxFetch then set maxFetch to evtCount
    set output to ""
    repeat with i from 1 to maxFetch
        set e to item i of evts
        set eid to id of e
        set esubject to subject of e
        set estart to start time of e as string
        set eend to end time of e as string
        set elocation to ""
        try
            set elocation to location of e
        end try
        set eorganizer to ""
        try
            set eorganizer to organizer of e
        end try
        set eallday to all day flag of e
        set output to output & (eid as text) & "{DELIM}" & esubject & "{DELIM}" & estart & "{DELIM}" & eend & "{DELIM}" & elocation & "{DELIM}" & eorganizer & "{DELIM}" & (eallday as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

        raw = await self.bridge.run(script)
        results = self._parse_event_summaries(raw)

        # Python-side date filtering
        filtered = []
        for e in results:
            iso = self._to_iso(e.start)
            if iso is None:
                filtered.append(e)
                continue
            dt = datetime.fromisoformat(iso)
            if start <= dt <= end:
                filtered.append(e)
        return filtered[:count]

    async def get_event(self, entry_id: str, account: str) -> EventFull:
        script = f'''tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    set eid to id of e
    set esubject to subject of e
    set estart to start time of e as string
    set eend to end time of e as string
    set elocation to ""
    try
        set elocation to location of e
    end try
    set eorganizer to ""
    try
        set eorganizer to organizer of e
    end try
    set eallday to all day flag of e
    set ebody to ""
    try
        set ebody to plain text content of e
    end try
    set eattendees to ""
    try
        set attList to attendees of e
        repeat with a in attList
            set eattendees to eattendees & address of a & "; "
        end repeat
    end try
    return (eid as text) & "{DELIM}" & esubject & "{DELIM}" & estart & "{DELIM}" & eend & "{DELIM}" & elocation & "{DELIM}" & eorganizer & "{DELIM}" & (eallday as text) & "{DELIM}" & ebody & "{DELIM}" & eattendees
end tell'''

        raw = await self.bridge.run(script)
        parts = raw.split(DELIM, 8)
        if len(parts) < 9:
            raise BackendError("Failed to parse event data")

        return EventFull(
            entry_id=parts[0].strip(),
            subject=parts[1].strip() or "(no subject)",
            start=parts[2].strip(),
            end=parts[3].strip(),
            location=_clean(parts[4]),
            organizer=_clean(parts[5]),
            all_day=parts[6].strip().lower() == "true",
            body=_truncate(_clean(parts[7])),
            attendees=parts[8].strip(),
        )

    async def create_event(
        self,
        subject: str,
        start: str,
        end: str,
        location: str,
        body: str,
        all_day: bool,
        reminder_minutes: int,
        account: str,
        folder: str,
    ) -> EventCreatedResult:
        start_dt = _parse_iso(start)
        end_dt = _parse_iso(end)

        props = f'subject:"{escape(subject)}", start time:{format_date(start_dt)}, end time:{format_date(end_dt)}'
        if location:
            props += f', location:"{escape(location)}"'
        if body:
            props += f', content:"{escape(body)}"'
        if all_day:
            props += ", all day flag:true"

        script = f'''tell application "Microsoft Outlook"
    set newEvt to make new calendar event with properties {{{props}}}
    return (id of newEvt as text) & "{DELIM}" & (subject of newEvt) & "{DELIM}" & (start time of newEvt as string) & "{DELIM}" & (end time of newEvt as string)
end tell'''

        raw = await self.bridge.run(script)
        parts = raw.split(DELIM)
        return EventCreatedResult(
            entry_id=parts[0].strip() if len(parts) > 0 else "",
            subject=parts[1].strip() if len(parts) > 1 else subject,
            start=parts[2].strip() if len(parts) > 2 else start,
            end=parts[3].strip() if len(parts) > 3 else end,
        )

    async def create_meeting(
        self,
        subject: str,
        start: str,
        end: str,
        required_attendees: str,
        location: str,
        body: str,
        optional_attendees: str,
        account: str,
    ) -> MeetingSentResult:
        start_dt = _parse_iso(start)
        end_dt = _parse_iso(end)

        props = f'subject:"{escape(subject)}", start time:{format_date(start_dt)}, end time:{format_date(end_dt)}'
        if location:
            props += f', location:"{escape(location)}"'
        if body:
            props += f', content:"{escape(body)}"'

        attendee_lines = ""
        for raw_addr in required_attendees.split(";"):
            addr = raw_addr.strip()
            if addr:
                attendee_lines += f'make new required attendee at newEvt with properties {{email address:{{address:"{escape(addr)}"}}}}\n'
        if optional_attendees:
            for raw_addr in optional_attendees.split(";"):
                addr = raw_addr.strip()
                if addr:
                    attendee_lines += f'make new optional attendee at newEvt with properties {{email address:{{address:"{escape(addr)}"}}}}\n'

        script = f"""tell application "Microsoft Outlook"
    set newEvt to make new calendar event with properties {{{props}}}
    {attendee_lines}
    return (id of newEvt as text)
end tell"""

        await self.bridge.run(script)
        return MeetingSentResult(
            subject=subject,
            required_attendees=required_attendees,
            optional_attendees=optional_attendees or None,
        )

    async def update_event(
        self,
        entry_id: str,
        subject: str,
        start: str,
        end: str,
        location: str,
        body: str,
        account: str,
    ) -> EventUpdatedResult:
        set_lines = ""
        if subject:
            set_lines += f'set subject of e to "{escape(subject)}"\n'
        if start:
            set_lines += f"set start time of e to {format_date(_parse_iso(start))}\n"
        if end:
            set_lines += f"set end time of e to {format_date(_parse_iso(end))}\n"
        if location:
            set_lines += f'set location of e to "{escape(location)}"\n'
        if body:
            set_lines += f'set content of e to "{escape(body)}"\n'

        if not set_lines:
            raise BackendError("No fields to update")

        script = f'''tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    {set_lines}
    return (id of e as text) & "{DELIM}" & (subject of e) & "{DELIM}" & (start time of e as string) & "{DELIM}" & (end time of e as string) & "{DELIM}" & (location of e)
end tell'''

        raw = await self.bridge.run(script)
        parts = raw.split(DELIM)
        return EventUpdatedResult(
            entry_id=parts[0].strip() if len(parts) > 0 else entry_id,
            subject=parts[1].strip() if len(parts) > 1 else "",
            start=parts[2].strip() if len(parts) > 2 else "",
            end=parts[3].strip() if len(parts) > 3 else "",
            location=_clean(parts[4]) if len(parts) > 4 else "",
        )

    async def delete_event(self, entry_id: str, account: str) -> ItemStatusResult:
        script = f"""tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    set esubject to subject of e
    delete e
    return esubject
end tell"""
        subject = await self.bridge.run(script)
        return ItemStatusResult(status="deleted", subject=subject, entry_id=entry_id)

    async def search_events(
        self,
        query: str,
        start_date: str,
        end_date: str,
        count: int,
        account: str,
        folder: str,
    ) -> list[EventSummary]:
        safe_query = escape(query)

        script = f'''tell application "Microsoft Outlook"
    set evts to calendar events whose subject contains "{safe_query}"
    set evtCount to count of evts
    set maxCount to {count}
    if evtCount < maxCount then set maxCount to evtCount
    set output to ""
    repeat with i from 1 to maxCount
        set e to item i of evts
        set eid to id of e
        set esubject to subject of e
        set estart to start time of e as string
        set eend to end time of e as string
        set elocation to ""
        try
            set elocation to location of e
        end try
        set eorganizer to ""
        try
            set eorganizer to organizer of e
        end try
        set eallday to all day flag of e
        set output to output & (eid as text) & "{DELIM}" & esubject & "{DELIM}" & estart & "{DELIM}" & eend & "{DELIM}" & elocation & "{DELIM}" & eorganizer & "{DELIM}" & (eallday as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

        raw = await self.bridge.run(script)
        return self._parse_event_summaries(raw)

    # --- tasks ---

    async def list_tasks(
        self,
        include_completed: bool,
        count: int,
        account: str,
    ) -> list[TaskSummary]:
        completed_filter = "" if include_completed else " whose todo flag is not completed"

        script = f'''tell application "Microsoft Outlook"
    set taskList to tasks{completed_filter}
    set taskCount to count of taskList
    set maxCount to {count}
    if taskCount < maxCount then set maxCount to taskCount
    set output to ""
    repeat with i from 1 to maxCount
        set t to item i of taskList
        set tid to id of t
        set tname to name of t
        set tdue to ""
        try
            set tdue to due date of t as string
        end try
        set tflag to todo flag of t
        set tpriority to priority of t
        set output to output & (tid as text) & "{DELIM}" & tname & "{DELIM}" & tdue & "{DELIM}" & (tflag as text) & "{DELIM}" & (tpriority as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

        raw = await self.bridge.run(script)
        results: list[TaskSummary] = []
        for raw_record in (raw or "").split(RECORD_DELIM):
            record = raw_record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 5:
                continue
            results.append(
                TaskSummary(
                    entry_id=parts[0].strip(),
                    subject=parts[1].strip() or "(no subject)",
                    due_date=_clean(parts[2]) or None,
                    complete=parts[3].strip() == "completed",
                    priority=parts[4].strip(),
                )
            )
        return results

    async def get_task(self, entry_id: str, account: str) -> TaskFull:
        script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set tid to id of t
    set tname to name of t
    set tdue to ""
    try
        set tdue to due date of t as string
    end try
    set tflag to todo flag of t
    set tpriority to priority of t
    set tbody to ""
    try
        set tbody to plain text content of t
    end try
    set tstartdate to ""
    try
        set tstartdate to start date of t as string
    end try
    return (tid as text) & "{DELIM}" & tname & "{DELIM}" & tdue & "{DELIM}" & (tflag as text) & "{DELIM}" & (tpriority as text) & "{DELIM}" & tbody & "{DELIM}" & tstartdate
end tell'''

        raw = await self.bridge.run(script)
        parts = raw.split(DELIM, 6)
        if len(parts) < 7:
            raise BackendError("Failed to parse task data")

        return TaskFull(
            entry_id=parts[0].strip(),
            subject=parts[1].strip() or "(no subject)",
            due_date=_clean(parts[2]) or None,
            complete=parts[3].strip() == "completed",
            priority=parts[4].strip(),
            body=_truncate(_clean(parts[5])),
            start_date=_clean(parts[6]) or None,
        )

    async def create_task(
        self,
        subject: str,
        body: str,
        due_date: str,
        importance: str,
        reminder_minutes: int,
        account: str,
    ) -> TaskCreatedResult:
        imp_map = {"low": "priority low", "normal": "priority normal", "high": "priority high"}
        imp_val = imp_map.get(importance.lower(), "priority normal")

        props = f'name:"{escape(subject)}", priority:{imp_val}'
        if due_date:
            props += f", due date:{format_date(_parse_iso(due_date))}"
        if body:
            props += f', content:"{escape(body)}"'

        script = f'''tell application "Microsoft Outlook"
    set newTask to make new task with properties {{{props}}}
    return (id of newTask as text) & "{DELIM}" & (name of newTask)
end tell'''

        raw = await self.bridge.run(script)
        parts = raw.split(DELIM)
        return TaskCreatedResult(
            entry_id=parts[0].strip() if len(parts) > 0 else "",
            subject=parts[1].strip() if len(parts) > 1 else subject,
            due_date=due_date or None,
        )

    async def complete_task(self, entry_id: str, account: str) -> ItemStatusResult:
        script = f"""tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set todo flag of t to completed
    return name of t
end tell"""
        name = await self.bridge.run(script)
        return ItemStatusResult(status="completed", subject=name, entry_id=entry_id)

    async def delete_task(self, entry_id: str, account: str) -> ItemStatusResult:
        script = f"""tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set tname to name of t
    delete t
    return tname
end tell"""
        name = await self.bridge.run(script)
        return ItemStatusResult(status="deleted", subject=name, entry_id=entry_id)

    # --- attachments ---

    async def list_attachments(self, entry_id: str, account: str) -> list[AttachmentInfo]:
        script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    set output to ""
    repeat with i from 1 to attCount
        set a to item i of attList
        set aname to name of a
        set asize to file size of a
        set output to output & (i as text) & "{DELIM}" & aname & "{DELIM}" & (asize as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

        raw = await self.bridge.run(script)
        results: list[AttachmentInfo] = []
        for raw_record in (raw or "").split(RECORD_DELIM):
            record = raw_record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 3:
                continue
            results.append(
                AttachmentInfo(
                    index=int(parts[0].strip()) if parts[0].strip().isdigit() else 0,
                    filename=parts[1].strip(),
                    size=int(parts[2].strip()) if parts[2].strip().isdigit() else 0,
                )
            )
        return results

    async def save_attachment(
        self,
        entry_id: str,
        attachment_index: int,
        save_directory: str,
        account: str,
    ) -> AttachmentSavedResult:
        if not save_directory:
            save_directory = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(save_directory, exist_ok=True)
        save_dir_posix = save_directory

        script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    if attCount < {attachment_index} then return "ERROR:Only " & attCount & " attachment(s)"
    set a to item {attachment_index} of attList
    set aname to name of a
    set savePath to "{escape(save_dir_posix)}/" & aname
    save a in savePath
    return aname & "{DELIM}" & savePath
end tell'''

        raw = await self.bridge.run(script)
        if raw.startswith("ERROR:"):
            raise BackendError(raw.removeprefix("ERROR:"))

        parts = raw.split(DELIM)
        filename = parts[0].strip() if len(parts) > 0 else "unknown"
        return AttachmentSavedResult(
            filename=filename,
            path=os.path.join(save_directory, filename),
        )

    # --- out of office ---

    async def set_out_of_office(self, enabled: bool, message: str, account: str) -> OofSetResult:
        enabled_str = "true" if enabled else "false"
        msg_arg = f', out of office message "{escape(message)}"' if message else ""

        script = f"""tell application "Microsoft Outlook"
    set out of office state to {enabled_str}{msg_arg}
    return out of office state as text
end tell"""

        result = await self.bridge.run(script)
        status = "on" if "true" in result.lower() else "off"
        return OofSetResult(out_of_office=status == "on", status=status)

    async def _ui_list_messages(self, count: int = 10) -> list[EmailSummary]:
        """Read visible inbox messages via UI scripting (System Events).

        Fallback for New Outlook for Mac where AppleScript's inbox keyword
        only sees the empty local mailbox.
        """
        script = (
            _UI_MESSAGE_LIST_PATH + f"                                    set rowList to rows\n"
            f"                                    set rowCount to count of rowList\n"
            f"                                    set maxRows to rowCount\n"
            f"                                    if maxRows > {count} then set maxRows to {count}\n"
            f'                                    set output to ""\n'
            f"                                    repeat with i from 1 to maxRows\n"
            f"                                        set r to row i\n"
            f"                                        try\n"
            f"                                            set cellDesc to description of UI element 1 of r\n"
            f'                                            set output to output & cellDesc & "{RECORD_DELIM}"\n'
            f"                                        end try\n"
            f"                                    end repeat\n"
            f"                                    return output\n" + _UI_MESSAGE_LIST_END
        )

        raw = await self.bridge.run(script)
        if not raw:
            return []

        results = []
        for idx, raw_record in enumerate(raw.split(RECORD_DELIM), start=1):
            record = raw_record.strip()
            if not record:
                continue
            # Cell description format uses `,` + 4+ spaces as major field
            # separators, while in-content commas have 0-1 trailing spaces.
            fields = [f.strip() for f in re.split(r",\s{4,}", record)]

            is_unread = False
            has_attachment = False
            cleaned = []
            for f in fields:
                if not f:
                    continue
                if f in _UNREAD_TOKENS:
                    is_unread = True
                    continue
                if any(tok in f for tok in _ATTACHMENT_TOKENS):
                    has_attachment = True
                    continue
                if f in _CATEGORY_TOKENS or any(f.startswith(p) for p in _SKIP_PREFIXES):
                    continue
                cleaned.append(f)

            sender_subject = cleaned[0] if cleaned else ""
            time_str = cleaned[1].rstrip(",").strip() if len(cleaned) > 1 else ""

            # Remove thread/unread count prefixes like "2 messages, "
            ss = re.sub(r"^\d+\s+[\w\s]+,\s*", "", sender_subject)
            comma_pos = ss.find(", ")
            if comma_pos > 0:
                sender = ss[:comma_pos].strip()
                subject = ss[comma_pos + 2 :].strip()
            else:
                sender = ""
                subject = ss.strip()

            results.append(
                EmailSummary(
                    entry_id=f"ui-{idx}",
                    subject=subject or "(could not parse subject)",
                    sender="",
                    sender_name=sender,
                    received_time=time_str,
                    unread=is_unread,
                    has_attachments=has_attachment,
                    attachment_count=1 if has_attachment else 0,
                )
            )

        return results

    # --- internal helpers ---

    @staticmethod
    def _recipient_lines(addresses: str, kind: str) -> str:
        """Build AppleScript lines adding recipients of one kind to newMsg."""
        lines = ""
        for raw_addr in addresses.split(";"):
            addr = raw_addr.strip()
            if addr:
                lines += f'make new {kind} at newMsg with properties {{email address:{{address:"{escape(addr)}"}}}}\n'
        return lines

    @staticmethod
    def _email_fields_script(var: str) -> str:
        """AppleScript lines collecting sender/time/recipient/body fields."""
        return f"""set msender to ""
    try
        set msender to address of sender of {var}
    end try
    set msenderName to ""
    try
        set msenderName to name of sender of {var}
    end try
    set mtime to time received of {var} as string
    set misread to is read of {var}
    set mattcount to 0
    try
        set mattcount to count of attachments of {var}
    end try
    set mto to ""
    try
        set recips to to recipients of {var}
        repeat with r in recips
            set mto to mto & address of r & "; "
        end repeat
    end try
    set mcc to ""
    try
        set recips to cc recipients of {var}
        repeat with r in recips
            set mcc to mcc & address of r & "; "
        end repeat
    end try
    set mbody to ""
    try
        set mbody to plain text content of {var}
    end try"""

    @staticmethod
    def _parse_email_summaries(raw: str) -> list[EmailSummary]:

        results: list[EmailSummary] = []
        for raw_record in (raw or "").split(RECORD_DELIM):
            record = raw_record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 7:
                continue
            att_count = int(parts[6]) if parts[6].strip().isdigit() else 0
            results.append(
                EmailSummary(
                    entry_id=parts[0].strip(),
                    subject=parts[1].strip() or "(no subject)",
                    sender=parts[2].strip(),
                    sender_name=parts[3].strip(),
                    received_time=_clean(parts[4]),
                    unread=parts[5].strip().lower() != "true",  # is_read -> unread
                    has_attachments=att_count > 0,
                    attachment_count=att_count,
                )
            )
        return results

    @staticmethod
    def _parse_event_summaries(raw: str) -> list[EventSummary]:
        results: list[EventSummary] = []
        for raw_record in (raw or "").split(RECORD_DELIM):
            record = raw_record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 7:
                continue
            results.append(
                EventSummary(
                    entry_id=parts[0].strip(),
                    subject=parts[1].strip() or "(no subject)",
                    start=parts[2].strip(),
                    end=parts[3].strip(),
                    location=_clean(parts[4]),
                    organizer=_clean(parts[5]),
                    all_day=parts[6].strip().lower() == "true",
                )
            )
        return results

    @staticmethod
    def _to_iso(locale_time: str) -> str | None:
        """Best-effort conversion of an AppleScript date string to ISO 8601."""
        from outlook_desktop_mcp.backends.mac.applescript_helpers import parse_date  # noqa: PLC0415

        iso = parse_date(locale_time)
        try:
            datetime.fromisoformat(iso)
            return iso
        except ValueError:
            return None
