"""Windows COM backend — automates classic Outlook Desktop via pywin32.

All methods run their COM work on the bridge thread via ``bridge.call`` and
return pydantic models. Handled failures raise :class:`BackendError`.
"""

from __future__ import annotations

import asyncio
import logging
import os
import re
from datetime import datetime, timedelta
from typing import Any, cast

from outlook_desktop_mcp.backends.base import Backend, BackendError
from outlook_desktop_mcp.backends.win._types import (
    Appointment,
    Folder,
    MailItem,
    Namespace,
    Store,
    TaskItem,
)
from outlook_desktop_mcp.backends.win._types import (
    Outlook as OutlookCOM,
)
from outlook_desktop_mcp.backends.win.bridge import OutlookBridge
from outlook_desktop_mcp.backends.win.errors import format_com_error
from outlook_desktop_mcp.backends.win.formatting import (
    format_email_full,
    format_email_summary,
    format_event_full,
    format_event_summary,
    format_task_full,
    format_task_summary,
)
from outlook_desktop_mcp.models import (
    AccountInfo,
    AttachmentInfo,
    AttachmentSavedResult,
    CategoriesSetResult,
    CategoryInfo,
    DraftSavedResult,
    EmailFull,
    EmailSummary,
    EventCreatedResult,
    EventFull,
    EventSummary,
    EventUpdatedResult,
    FolderInfo,
    ItemStatusResult,
    MeetingResponseResult,
    MeetingSentResult,
    MovedResult,
    OofSetResult,
    OofStatus,
    ReplyDraftSavedResult,
    ReplySentResult,
    RuleInfo,
    RuleToggledResult,
    SentResult,
    TaskCreatedResult,
    TaskFull,
    TaskSummary,
)
from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_TASKS,
    OL_MAIL_ITEM,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_OPTIONAL,
    OL_REQUIRED,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_RESPONSE_TENTATIVE,
    OL_TASK_COMPLETE,
    OL_TASK_ITEM,
)

logger = logging.getLogger("outlook_desktop_mcp.backends.win.backend")

# Outlook item Class constants (olObjectClass — distinct from olItemType)
_OL_CLASS_MAIL = 43
_OL_CLASS_APPOINTMENT = 26
_OL_CLASS_TASK = 48


def _safe_dasl(query: str) -> str:
    """Sanitize a string for use in a DASL LIKE filter value.

    Escapes SQL wildcards (% and _) so user input is treated as literals,
    then escapes quote characters required by DASL syntax.
    """
    query = query.replace("%", "[%]").replace("_", "[_]")
    return query.replace("'", "''").replace('"', '""')


def _parse_date(date_str: str) -> datetime:
    """Parse ISO 8601 date string like '2026-02-25 14:00' or '2026-02-25T14:00:00'."""
    return datetime.fromisoformat(date_str)


def _require_class(item: Any, expected_class: int, label: str) -> None:
    """Raise BackendError if the item is not of the expected Outlook class."""
    if item.Class != expected_class:
        raise BackendError(f"Entry ID does not refer to a {label}.")


def _resolve_store(namespace: Namespace, account: str = "") -> Store | None:
    """Resolve an account name to an Outlook Store object.

    If account is empty, returns DefaultStore. Otherwise does a
    case-insensitive substring match on Store.DisplayName.
    """
    if not account:
        return namespace.DefaultStore

    account_lower = account.lower().strip()
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i + 1)
        if account_lower in store.DisplayName.lower():
            return store

    return None


def _require_store(namespace: Namespace, account: str = "") -> Store:
    """Resolve store, raising BackendError if not found."""
    store = _resolve_store(namespace, account)
    if store is None:
        raise BackendError(f"Account '{account}' not found. Use list_accounts to see available accounts.")
    return store


def _walk_folders(parent: Folder, name_lower: str) -> Folder | None:
    """Recursively search subfolders of parent for a folder matching name_lower."""
    for i in range(parent.Folders.Count):
        try:
            f: Folder = parent.Folders.Item(i + 1)
            if f.Name.lower() == name_lower:
                return f
            found = _walk_folders(f, name_lower)
            if found:
                return found
        except Exception:  # noqa: S112
            continue
    return None


def _resolve_folder(namespace: Namespace, folder_name: str, store: Store | None = None) -> Folder | None:
    """Resolve a folder name to an Outlook MAPIFolder object.

    Resolution order:
    1. Slash-delimited path (e.g. "Inbox/Receipts") — traverse segment by segment
    2. Built-in Outlook folder enum (inbox, sent, deleted, etc.)
    3. Root-level folder name match (fast path)
    4. Recursive depth-first search of entire folder tree (fallback)
    """
    folder_name = folder_name.strip()
    store = store or namespace.DefaultStore

    if "/" in folder_name:
        parts = [p.strip() for p in folder_name.split("/")]
        current: Folder | None = _resolve_folder(namespace, parts[0], store)
        if current is None:
            return None
        for part in parts[1:]:
            part_lower = part.lower()
            found: Folder | None = None
            for i in range(current.Folders.Count):
                try:
                    f: Folder = current.Folders.Item(i + 1)
                    if f.Name.lower() == part_lower:
                        found = f
                        break
                except Exception:  # noqa: S112
                    continue
            if found is None:
                return None
            current = found
        return current

    folder_lower = folder_name.lower()

    if folder_lower in FOLDER_NAME_TO_ENUM:
        return store.GetDefaultFolder(FOLDER_NAME_TO_ENUM[folder_lower])

    root = store.GetRootFolder()
    for i in range(root.Folders.Count):
        try:
            f = root.Folders.Item(i + 1)
            if f.Name.lower() == folder_lower:
                return f
        except Exception:  # noqa: S112
            continue

    return _walk_folders(root, folder_lower)


class ComBackend(Backend):
    """Outlook Classic on Windows, driven through the COM bridge thread."""

    name = "com"
    supports_accounts = True
    supports_rules = True
    supports_categories = True
    supports_meeting_response = True
    supports_oof_status_query = True

    def __init__(self) -> None:
        self.bridge = OutlookBridge()

    def start(self) -> None:
        asyncio.run(self.bridge.start())

    def stop(self) -> None:
        asyncio.run(self.bridge.stop())

    def format_unexpected_error(self, action: str, e: Exception) -> str:
        """Surface COM HRESULT detail (Windows-only override)."""
        return f"{action}: {format_com_error(e)}"

    # --- accounts ---

    async def list_accounts(self) -> list[AccountInfo]:
        def _list(outlook: OutlookCOM, namespace: Namespace) -> list[AccountInfo]:
            default_id = namespace.DefaultStore.StoreID
            return [
                AccountInfo(
                    display_name=store.DisplayName,
                    store_id=store.StoreID,
                    is_default=store.StoreID == default_id,
                )
                for i in range(namespace.Stores.Count)
                for store in [namespace.Stores.Item(i + 1)]
            ]

        return await self.bridge.call(_list)

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
        def _compose(outlook: OutlookCOM, namespace: Namespace) -> SentResult | DraftSavedResult:
            store = _require_store(namespace, account)
            mail: MailItem = outlook.CreateItem(OL_MAIL_ITEM)
            for acc in outlook.Session.Accounts:
                if acc.DeliveryStore.StoreID == store.StoreID:
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))  # SendUsingAccount
                    break
            mail.To = to
            mail.Subject = subject
            mail.Body = body
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc
            if html_body:
                mail.HTMLBody = html_body
            if send:
                mail.Send()
                return SentResult(subject=mail.Subject, to=to)
            mail.Save()
            return DraftSavedResult(subject=mail.Subject, to=to, entry_id=mail.EntryID)

        return await self.bridge.call(_compose)

    async def list_emails(
        self,
        folder: str,
        count: int,
        unread_only: bool,
        start_date: str,
        end_date: str,
        account: str,
    ) -> list[EmailSummary]:
        def _list(outlook: OutlookCOM, namespace: Namespace) -> list[EmailSummary]:
            effective_count = min(max(1, count), 200)
            store = _require_store(namespace, account)
            target = _resolve_folder(namespace, folder, store)
            if not target:
                raise BackendError(f"Folder '{folder}' not found")

            items = target.Items
            items.Sort("[ReceivedTime]", True)

            restrictions = []
            if unread_only:
                restrictions.append("[UnRead] = True")
            if start_date:
                start = _parse_date(start_date)
                restrictions.append(f"[ReceivedTime] >= '{start.strftime('%m/%d/%Y %H:%M')}'")
            if end_date:
                end = _parse_date(end_date)
                restrictions.append(f"[ReceivedTime] <= '{end.strftime('%m/%d/%Y %H:%M')}'")
            elif start_date:
                restrictions.append(f"[ReceivedTime] <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'")

            if restrictions:
                items = items.Restrict(" AND ".join(restrictions))

            results = []
            limit = min(effective_count, items.Count)
            for i in range(limit):
                try:
                    results.append(EmailSummary.model_validate(format_email_summary(items.Item(i + 1))))
                except Exception:  # noqa: S112
                    continue
            return results

        return await self.bridge.call(_list)

    async def read_email(
        self,
        entry_id: str,
        subject_search: str,
        folder: str,
        account: str,
    ) -> EmailFull:
        def _read(outlook: OutlookCOM, namespace: Namespace) -> EmailFull:
            if entry_id:
                return EmailFull.model_validate(format_email_full(namespace.GetItemFromID(entry_id)))

            if not subject_search:
                raise BackendError("Provide either entry_id or subject_search")

            store = _require_store(namespace, account)
            target = _resolve_folder(namespace, folder, store)
            if not target:
                raise BackendError(f"Folder '{folder}' not found")

            safe_query = _safe_dasl(subject_search)
            filter_str = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"
            items = target.Items.Restrict(filter_str)
            items.Sort("[ReceivedTime]", True)
            if items.Count == 0:
                raise BackendError(f"No email found matching '{subject_search}'")

            return EmailFull.model_validate(format_email_full(items.Item(1)))

        return await self.bridge.call(_read)

    async def mark_as_read(self, entry_id: str, account: str) -> ItemStatusResult:
        def _mark(outlook: OutlookCOM, namespace: Namespace) -> ItemStatusResult:
            item = self._get_item(namespace, entry_id, account)
            _require_class(item, _OL_CLASS_MAIL, "mail item")
            subject = item.Subject
            item.UnRead = False
            item.Save()
            return ItemStatusResult(status="marked_read", subject=subject, entry_id=entry_id)

        return await self.bridge.call(_mark)

    async def mark_as_unread(self, entry_id: str, account: str) -> ItemStatusResult:
        def _mark(outlook: OutlookCOM, namespace: Namespace) -> ItemStatusResult:
            item = self._get_item(namespace, entry_id, account)
            _require_class(item, _OL_CLASS_MAIL, "mail item")
            subject = item.Subject
            item.UnRead = True
            item.Save()
            return ItemStatusResult(status="marked_unread", subject=subject, entry_id=entry_id)

        return await self.bridge.call(_mark)

    async def move_email(self, entry_id: str, target_folder: str, account: str) -> MovedResult:
        def _move(outlook: OutlookCOM, namespace: Namespace) -> MovedResult:
            item = namespace.GetItemFromID(entry_id)
            _require_class(item, _OL_CLASS_MAIL, "mail item")
            subject = item.Subject

            store = _require_store(namespace, account)
            dest = _resolve_folder(namespace, target_folder, store)
            if not dest:
                raise BackendError(f"Target folder '{target_folder}' not found. Use list_folders to see available folders.")

            item.Move(dest)
            return MovedResult(subject=subject, target_folder=target_folder)

        return await self.bridge.call(_move)

    async def reply_email(
        self,
        entry_id: str,
        body: str,
        reply_all: bool,
        account: str,
        send: bool,
    ) -> ReplySentResult | ReplyDraftSavedResult:
        def _reply(
            outlook: OutlookCOM,
            namespace: Namespace,
        ) -> ReplySentResult | ReplyDraftSavedResult:
            item = self._get_item(namespace, entry_id, account)
            _require_class(item, _OL_CLASS_MAIL, "mail item")
            subject = item.Subject
            reply_item = item.ReplyAll() if reply_all else item.Reply()
            if html_body:
                reply_item.HTMLBody = html_body + "<br><br>" + reply_item.HTMLBody
            else:
                reply_item.Body = body + "\n\n" + reply_item.Body
            if send:
                reply_item.Send()
                return ReplySentResult(subject=subject, reply_all=reply_all)
            reply_item.Save()
            return ReplyDraftSavedResult(subject=subject, reply_all=reply_all, entry_id=reply_item.EntryID)

        return await self.bridge.call(_reply)

    async def list_folders(self, folder: str, max_depth: int, account: str) -> list[FolderInfo]:
        def _list(outlook: OutlookCOM, namespace: Namespace) -> list[FolderInfo]:
            effective_max_depth = min(max(1, max_depth), 10)
            store = _require_store(namespace, account)

            if folder:
                start = _resolve_folder(namespace, folder, store)
                if not start:
                    raise BackendError(f"Folder '{folder}' not found")
                base_path = folder
            else:
                start = store.GetRootFolder()
                base_path = ""

            def walk(f: Folder, depth: int, path_prefix: str) -> FolderInfo:
                current_path = f"{path_prefix}/{f.Name}" if path_prefix else f.Name
                info = FolderInfo(
                    name=f.Name,
                    full_path=current_path,
                    item_count=f.Items.Count,
                    unread_count=f.UnReadItemCount,
                )
                if depth < effective_max_depth:
                    for i in range(f.Folders.Count):
                        try:
                            child: Folder = f.Folders.Item(i + 1)
                            info.subfolders.append(walk(child, depth + 1, current_path))
                        except Exception:  # noqa: S112
                            continue
                return info

            folders: list[FolderInfo] = []
            for i in range(start.Folders.Count):
                try:
                    child: Folder = start.Folders.Item(i + 1)
                    folders.append(walk(child, 1, base_path))
                except Exception:  # noqa: S112
                    continue
            return folders

        return await self.bridge.call(_list)

    async def search_emails(
        self,
        query: str,
        folder: str,
        count: int,
        start_date: str,
        end_date: str,
        account: str,
    ) -> list[EmailSummary]:
        def _search(outlook: OutlookCOM, namespace: Namespace) -> list[EmailSummary]:
            effective_count = min(max(1, count), 200)
            store = _require_store(namespace, account)
            target = _resolve_folder(namespace, folder, store)
            if not target:
                raise BackendError(f"Folder '{folder}' not found")

            safe_query = _safe_dasl(query)
            dasl_parts = [f"(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"]
            if start_date:
                start = _parse_date(start_date)
                dasl_parts.append(f"\"urn:schemas:httpmail:datereceived\" >= '{start.strftime('%m/%d/%Y %H:%M')}'")
            if end_date:
                end = _parse_date(end_date)
                dasl_parts.append(f"\"urn:schemas:httpmail:datereceived\" <= '{end.strftime('%m/%d/%Y %H:%M')}'")
            elif start_date:
                dasl_parts.append(f"\"urn:schemas:httpmail:datereceived\" <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'")

            filter_str = "@SQL=" + " AND ".join(dasl_parts)
            items = target.Items.Restrict(filter_str)
            items.Sort("[ReceivedTime]", True)

            results = []
            limit = min(effective_count, items.Count)
            for i in range(limit):
                try:
                    results.append(EmailSummary.model_validate(format_email_summary(items.Item(i + 1))))
                except Exception:  # noqa: S112
                    continue
            return results

        return await self.bridge.call(_search)

    # --- calendar ---

    async def list_events(
        self,
        start_date: str,
        end_date: str,
        count: int,
        account: str,
        folder: str,
    ) -> list[EventSummary]:
        def _list(outlook: OutlookCOM, namespace: Namespace) -> list[EventSummary]:
            effective_count = min(max(1, count), 200)
            store = _require_store(namespace, account)
            calendar = self._resolve_calendar(namespace, folder, store)
            items = calendar.Items

            # CRITICAL ORDER: Sort BEFORE IncludeRecurrences BEFORE Restrict
            items.Sort("[Start]")
            items.IncludeRecurrences = True

            start = _parse_date(start_date) if start_date else datetime.now()
            end = _parse_date(end_date) if end_date else start + timedelta(days=7)

            restrict = f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M')}' AND [Start] <= '{end.strftime('%m/%d/%Y %H:%M')}'"
            filtered = items.Restrict(restrict)

            results = []
            n = 0
            for item in filtered:
                n += 1
                try:
                    results.append(EventSummary.model_validate(format_event_summary(item)))
                except Exception:  # noqa: S112
                    continue
                if n >= effective_count:
                    break
            return results

        return await self.bridge.call(_list)

    async def get_event(self, entry_id: str, account: str) -> EventFull:
        def _get(outlook: OutlookCOM, namespace: Namespace) -> EventFull:
            item = self._get_item(namespace, entry_id, account)
            return EventFull.model_validate(format_event_full(item))

        return await self.bridge.call(_get)

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
        def _create(outlook: OutlookCOM, namespace: Namespace) -> EventCreatedResult:
            appt: Appointment = outlook.CreateItem(OL_APPOINTMENT_ITEM)
            if account:
                store = _require_store(namespace, account)
                cal = self._resolve_calendar(namespace, folder, store)
                appt = cast(Appointment, appt.Move(cal))
            elif folder:
                cal = self._resolve_calendar(namespace, folder, namespace.DefaultStore)
                appt = cast(Appointment, appt.Move(cal))
            appt.Subject = subject
            appt.Start = start
            appt.End = end
            if location:
                appt.Location = location
            if body:
                appt.Body = body
            appt.AllDayEvent = all_day
            if not all_day and reminder_minutes > 0:
                appt.ReminderSet = True
                appt.ReminderMinutesBeforeStart = reminder_minutes
            else:
                appt.ReminderSet = False
            appt.Save()
            return EventCreatedResult(
                subject=appt.Subject,
                start=str(appt.Start),
                end=str(appt.End),
                entry_id=appt.EntryID,
            )

        return await self.bridge.call(_create)

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
        def _create(outlook: OutlookCOM, namespace: Namespace) -> MeetingSentResult:
            appt: Appointment = outlook.CreateItem(OL_APPOINTMENT_ITEM)
            if account:
                store = _require_store(namespace, account)
                for acc in outlook.Session.Accounts:
                    if acc.DeliveryStore.StoreID == store.StoreID:
                        appt._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                        break
            appt.Subject = subject
            appt.Start = start
            appt.End = end
            appt.MeetingStatus = OL_MEETING
            if location:
                appt.Location = location
            if body:
                appt.Body = body

            for raw_addr in required_attendees.split(";"):
                addr = raw_addr.strip()
                if addr:
                    recip = appt.Recipients.Add(addr)
                    recip.Type = OL_REQUIRED

            if optional_attendees:
                for raw_addr in optional_attendees.split(";"):
                    addr = raw_addr.strip()
                    if addr:
                        recip = appt.Recipients.Add(addr)
                        recip.Type = OL_OPTIONAL

            appt.Recipients.ResolveAll()
            appt.Send()
            return MeetingSentResult(
                subject=subject,
                required_attendees=required_attendees,
                optional_attendees=optional_attendees or None,
            )

        return await self.bridge.call(_create)

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
        def _update(outlook: OutlookCOM, namespace: Namespace) -> EventUpdatedResult:
            item = cast(Appointment, self._get_item(namespace, entry_id, account))
            _require_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item")
            if subject:
                item.Subject = subject
            if start:
                item.Start = start
            if end:
                item.End = end
            if location:
                item.Location = location
            if body:
                item.Body = body
            item.Save()
            return EventUpdatedResult(
                subject=item.Subject,
                start=str(item.Start),
                end=str(item.End),
                location=item.Location or "",
                entry_id=item.EntryID,
            )

        return await self.bridge.call(_update)

    async def delete_event(self, entry_id: str, account: str) -> ItemStatusResult:
        def _delete(outlook: OutlookCOM, namespace: Namespace) -> ItemStatusResult:
            item = self._get_item(namespace, entry_id, account)
            _require_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item")
            subject = item.Subject

            # If this is a meeting we organized, cancel it (sends notices)
            if item.MeetingStatus == OL_MEETING:
                item.MeetingStatus = OL_MEETING_CANCELED
                item.Send()
                return ItemStatusResult(
                    status="meeting_canceled",
                    subject=subject,
                    note="Cancellation sent to attendees",
                )

            item.Delete()
            return ItemStatusResult(status="deleted", subject=subject)

        return await self.bridge.call(_delete)

    async def respond_to_meeting(
        self,
        entry_id: str,
        response: str,
        account: str,
    ) -> MeetingResponseResult:
        response_map = {
            "accept": OL_RESPONSE_ACCEPTED,
            "decline": OL_RESPONSE_DECLINED,
            "tentative": OL_RESPONSE_TENTATIVE,
        }
        response_lower = response.lower().strip()
        if response_lower not in response_map:
            raise BackendError(f"response must be 'accept', 'decline', or 'tentative'. Got: '{response}'")

        def _respond(outlook: OutlookCOM, namespace: Namespace) -> MeetingResponseResult:
            item = self._get_item(namespace, entry_id, account)
            _require_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item")
            subject = item.Subject
            response_item = item.Respond(response_map[response_lower])
            response_item.Send()
            return MeetingResponseResult(response=response_lower, subject=subject)

        return await self.bridge.call(_respond)

    async def search_events(
        self,
        query: str,
        start_date: str,
        end_date: str,
        count: int,
        account: str,
        folder: str,
    ) -> list[EventSummary]:
        def _search(outlook: OutlookCOM, namespace: Namespace) -> list[EventSummary]:
            effective_count = min(max(1, count), 200)
            store = _require_store(namespace, account)
            calendar = self._resolve_calendar(namespace, folder, store)
            items = calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True

            start = _parse_date(start_date) if start_date else datetime.now() - timedelta(days=30)
            end = _parse_date(end_date) if end_date else datetime.now() + timedelta(days=30)

            restrict = f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M')}' AND [Start] <= '{end.strftime('%m/%d/%Y %H:%M')}'"
            filtered = items.Restrict(restrict)

            query_lower = query.lower()
            results = []
            for item in filtered:
                if query_lower in (item.Subject or "").lower():
                    try:
                        results.append(EventSummary.model_validate(format_event_summary(item)))
                    except Exception:  # noqa: S112
                        continue
                    if len(results) >= effective_count:
                        break
            return results

        return await self.bridge.call(_search)

    # --- tasks ---

    async def list_tasks(
        self,
        include_completed: bool,
        count: int,
        account: str,
    ) -> list[TaskSummary]:
        def _list(outlook: OutlookCOM, namespace: Namespace) -> list[TaskSummary]:
            effective_count = min(max(1, count), 200)
            store = _require_store(namespace, account)
            folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
            items = folder.Items
            items.Sort("[DueDate]")

            if not include_completed:
                items = items.Restrict("[Complete] = False")

            results = []
            limit = min(effective_count, items.Count)
            for i in range(limit):
                try:
                    results.append(TaskSummary.model_validate(format_task_summary(items.Item(i + 1))))
                except Exception:  # noqa: S112
                    continue
            return results

        return await self.bridge.call(_list)

    async def get_task(self, entry_id: str, account: str) -> TaskFull:
        def _get(outlook: OutlookCOM, namespace: Namespace) -> TaskFull:
            item = self._get_item(namespace, entry_id, account)
            return TaskFull.model_validate(format_task_full(item))

        return await self.bridge.call(_get)

    async def create_task(
        self,
        subject: str,
        body: str,
        due_date: str,
        importance: str,
        reminder_minutes: int,
        account: str,
    ) -> TaskCreatedResult:
        def _create(outlook: OutlookCOM, namespace: Namespace) -> TaskCreatedResult:
            task: TaskItem = outlook.CreateItem(OL_TASK_ITEM)
            if account:
                store = _require_store(namespace, account)
                tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
                task = cast(TaskItem, task.Move(tasks_folder))
            task.Subject = subject
            if body:
                task.Body = body
            if due_date:
                task.DueDate = due_date
            imp_map = {"low": 0, "normal": 1, "high": 2}
            task.Importance = imp_map.get(importance.lower(), 1)
            if reminder_minutes > 0:
                task.ReminderSet = True
                task.ReminderMinutesBeforeStart = reminder_minutes
            else:
                task.ReminderSet = False
            task.Save()
            return TaskCreatedResult(
                subject=task.Subject,
                entry_id=task.EntryID,
                due_date=str(task.DueDate) if due_date else None,
            )

        return await self.bridge.call(_create)

    async def complete_task(self, entry_id: str, account: str) -> ItemStatusResult:
        def _complete(outlook: OutlookCOM, namespace: Namespace) -> ItemStatusResult:
            item = cast(TaskItem, self._get_item(namespace, entry_id, account))
            _require_class(item, _OL_CLASS_TASK, "task item")
            item.Status = OL_TASK_COMPLETE
            item.PercentComplete = 100
            item.Save()
            return ItemStatusResult(status="completed", subject=item.Subject, entry_id=entry_id)

        return await self.bridge.call(_complete)

    async def delete_task(self, entry_id: str, account: str) -> ItemStatusResult:
        def _delete(outlook: OutlookCOM, namespace: Namespace) -> ItemStatusResult:
            item = cast(TaskItem, self._get_item(namespace, entry_id, account))
            _require_class(item, _OL_CLASS_TASK, "task item")
            subject = item.Subject
            item.Delete()
            return ItemStatusResult(status="deleted", subject=subject, entry_id=entry_id)

        return await self.bridge.call(_delete)

    # --- attachments ---

    async def list_attachments(self, entry_id: str, account: str) -> list[AttachmentInfo]:
        def _list(outlook: OutlookCOM, namespace: Namespace) -> list[AttachmentInfo]:
            item = self._get_item(namespace, entry_id, account)
            return [
                AttachmentInfo(index=i + 1, filename=att.FileName, size=att.Size)
                for i in range(item.Attachments.Count)
                for att in [item.Attachments.Item(i + 1)]
            ]

        return await self.bridge.call(_list)

    async def save_attachment(
        self,
        entry_id: str,
        attachment_index: int,
        save_directory: str,
        account: str,
    ) -> AttachmentSavedResult:
        def _save(outlook: OutlookCOM, namespace: Namespace) -> AttachmentSavedResult:
            item = self._get_item(namespace, entry_id, account)
            if attachment_index < 1 or item.Attachments.Count < attachment_index:
                raise BackendError(f"Only {item.Attachments.Count} attachment(s), requested index {attachment_index}")

            att = item.Attachments.Item(attachment_index)
            if not save_directory:
                save_directory = os.path.join(os.path.expanduser("~"), "Downloads")

            save_directory = os.path.realpath(save_directory)
            os.makedirs(save_directory, exist_ok=True)

            # Strip path separators and dangerous characters from filename
            safe_name = os.path.basename(att.FileName)
            safe_name = re.sub(r"[^\w\.\-_ ]", "_", safe_name)
            if not safe_name:
                safe_name = "attachment"

            save_path = os.path.join(save_directory, safe_name)

            # Ensure final path is still inside the intended directory
            real_path = os.path.realpath(save_path)
            if not real_path.startswith(save_directory + os.sep) and real_path != save_directory:
                raise BackendError("Attachment filename would escape the target directory.")

            att.SaveAsFile(save_path)
            return AttachmentSavedResult(filename=safe_name, path=save_path, size=att.Size)

        return await self.bridge.call(_save)

    # --- categories ---

    async def list_categories(self, account: str) -> list[CategoryInfo]:
        def _list(outlook: OutlookCOM, namespace: Namespace) -> list[CategoryInfo]:
            # Categories are profile-wide, not per-store; account accepted for consistency
            return [CategoryInfo(name=cat.Name, color=cat.Color) for i in range(namespace.Categories.Count) for cat in [namespace.Categories.Item(i + 1)]]

        return await self.bridge.call(_list)

    async def set_category(
        self,
        entry_id: str,
        categories: str,
        account: str,
    ) -> CategoriesSetResult:
        def _set(outlook: OutlookCOM, namespace: Namespace) -> CategoriesSetResult:
            item = self._get_item(namespace, entry_id, account)
            item.Categories = categories
            item.Save()
            return CategoriesSetResult(subject=item.Subject, categories=item.Categories or "")

        return await self.bridge.call(_set)

    # --- rules ---

    async def list_rules(self, account: str) -> list[RuleInfo]:
        def _list(outlook: OutlookCOM, namespace: Namespace) -> list[RuleInfo]:
            store = _require_store(namespace, account)
            rules = store.GetRules()
            return [RuleInfo(index=i + 1, name=rule.Name, enabled=bool(rule.Enabled)) for i in range(rules.Count) for rule in [rules.Item(i + 1)]]

        return await self.bridge.call(_list)

    async def toggle_rule(self, rule_name: str, enabled: bool, account: str) -> RuleToggledResult:
        def _toggle(outlook: OutlookCOM, namespace: Namespace) -> RuleToggledResult:
            store = _require_store(namespace, account)
            rules = store.GetRules()
            for i in range(rules.Count):
                rule = rules.Item(i + 1)
                if rule.Name == rule_name:
                    logger.warning("toggle_rule: setting rule '%s' enabled=%s", rule_name, enabled)
                    rule.Enabled = enabled
                    rules.Save()
                    return RuleToggledResult(status="enabled" if enabled else "disabled", rule=rule_name)
            raise BackendError(f"Rule '{rule_name}' not found. Use list_rules to see available rules.")

        return await self.bridge.call(_toggle)

    # --- out of office ---

    async def get_out_of_office(self, account: str) -> OofStatus:
        def _get(outlook: OutlookCOM, namespace: Namespace) -> OofStatus:
            store = _require_store(namespace, account)
            prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
            try:
                oof_state = bool(store.PropertyAccessor.GetProperty(prop_tag))
                return OofStatus(out_of_office=oof_state, status="on" if oof_state else "off")
            except Exception:
                return OofStatus(
                    out_of_office=None,
                    status="unknown",
                    note="Could not read OOF property. Check Outlook settings directly.",
                )

        return await self.bridge.call(_get)

    async def set_out_of_office(self, enabled: bool, message: str, account: str) -> OofSetResult:
        def _set(outlook: OutlookCOM, namespace: Namespace) -> OofSetResult:
            store = _require_store(namespace, account)
            prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
            store.PropertyAccessor.SetProperty(prop_tag, enabled)
            note = ""
            if enabled and message:
                note = "OOF enabled, but the auto-reply message cannot be set via COM. Configure it in Outlook: File > Automatic Replies."
            return OofSetResult(
                out_of_office=enabled,
                status="on" if enabled else "off",
                note=note,
            )

        return await self.bridge.call(_set)

    # --- shared COM helpers (run on the bridge thread) ---

    @staticmethod
    def _get_item(namespace: Namespace, entry_id: str, account: str) -> Any:
        """Fetch an item by EntryID, optionally scoped to a store."""
        if account:
            store = _require_store(namespace, account)
            return namespace.GetItemFromID(entry_id, store.StoreID)
        return namespace.GetItemFromID(entry_id)

    @staticmethod
    def _resolve_calendar(namespace: Namespace, folder: str, store: Store) -> Folder:
        """Resolve the calendar folder: named path or the store's default."""
        if folder:
            calendar = _resolve_folder(namespace, folder, store)
            if not calendar:
                raise BackendError(f"Calendar folder '{folder}' not found")
            return calendar
        return store.GetDefaultFolder(OL_FOLDER_CALENDAR)
