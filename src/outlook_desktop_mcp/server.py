"""Unified Outlook Desktop MCP Server
=====================================
Single MCP tool surface for Outlook Desktop on Windows (COM) and macOS
(AppleScript). Platform behavior lives in ``outlook_desktop_mcp.backends``;
this module owns tool signatures, docstrings (the LLM-facing contract),
input validation, and the JSON response contract (pydantic models from
``outlook_desktop_mcp.models``).

The backend is injected via :func:`set_backend` — platform selection happens
only in ``entrypoint.py``.

Entry point: python -m outlook_desktop_mcp
"""

from __future__ import annotations

import json
import logging
import sys
from collections.abc import Awaitable, Callable
from typing import TypeVar

from fastmcp import FastMCP
from fastmcp.tools.base import Tool
from pydantic import BaseModel

from outlook_desktop_mcp.backends.base import Backend, BackendError
from outlook_desktop_mcp.instructions import build_instructions
from outlook_desktop_mcp.models import Error
from outlook_desktop_mcp.platform import Platform

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")

# --- MCP Server ---
#
# ``mcp`` is constructed once per platform: FastMCP freezes ``instructions``
# at construction, so we rebuild it inside :func:`set_backend` using the
# platform-specific instructions from :mod:`outlook_desktop_mcp.instructions`.
# All @mcp.tool() decorators below register against the *template* instance
# and the registered handlers are re-applied to the live instance in
# :func:`set_backend`.

_template_mcp = FastMCP("outlook-desktop-mcp")

# Exposed as ``mcp`` so existing tool decorators continue to bind to a
# recognisable FastMCP instance. :func:`main` always uses the live,
# per-platform instance installed by :func:`set_backend`.
mcp = _template_mcp

# --- Backend injection (platform choice happens in entrypoint.py only) ---

backend: Backend | None = None


def set_backend(b: Backend, platform: Platform) -> None:
    """Install the platform backend and register capability-gated tools.

    Rebuilds the FastMCP instance with the platform-specific instructions
    string and re-registers every tool already bound to the template
    instance.
    """
    global backend, mcp  # noqa: PLW0603
    backend = b

    live_mcp = FastMCP("outlook-desktop-mcp", instructions=build_instructions(platform))

    # Re-register every tool that was bound to the template at import time.
    for tool in _template_mcp._local_provider._components.values():
        if isinstance(tool, Tool):
            live_mcp.add_tool(tool)

    mcp = live_mcp

    if b.supports_accounts:
        mcp.tool()(list_accounts)
    if b.supports_meeting_response:
        mcp.tool()(respond_to_meeting)
    if b.supports_categories:
        mcp.tool()(list_categories)
        mcp.tool()(set_category)
    if b.supports_rules:
        mcp.tool()(list_rules)
        mcp.tool()(toggle_rule)
    if b.supports_oof_status_query:
        mcp.tool()(get_out_of_office)
    logger.info("Backend set: %s (platform=%s)", b.name, platform)


def _require_backend() -> Backend:
    if backend is None:
        raise BackendError("No backend configured (call set_backend first)")
    return backend


# --- Response helpers ---


def _dump(result: BaseModel | list[BaseModel]) -> str:
    """Serialize a model or model list to the tool JSON contract."""
    if isinstance(result, BaseModel):
        return result.model_dump_json(indent=2)
    return json.dumps([r.model_dump() for r in result], indent=2)


M = TypeVar("M", bound=BaseModel)


async def _run[M: BaseModel](action: str, call: Callable[[], Awaitable[M | list[BaseModel]]]) -> str:
    """Execute a backend call, converting failures into Error JSON."""
    try:
        return _dump(await call())
    except BackendError as e:
        return Error(error=str(e)).model_dump_json(indent=2)
    except Exception as e:  # noqa: BLE001 - tool boundary
        return Error(error=_require_backend().format_unexpected_error(action, e)).model_dump_json(indent=2)


def _validate_subject(subject: str | None, *, allow_skip: bool = False) -> str | None:
    """Return an Error JSON string if subject is blank, else None.

    Args:
        subject: The subject value to check.
        allow_skip: If True, an empty string means "do not change" (for update_*
            tools) and is accepted; whitespace-only is still rejected.
    """
    if subject is None:
        return Error(error="subject is required and must not be blank").model_dump_json(indent=2)
    if not subject and allow_skip:
        return None
    if not subject or not subject.strip():
        msg = "subject is required and must not be blank" if not allow_skip else "subject, if provided, must not be blank"
        return Error(error=msg).model_dump_json(indent=2)
    return None


# =====================================================================
# EMAIL TOOLS
# =====================================================================


@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately. The email will appear in the
    user's Sent Items folder after sending.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.
        account: Optional (Windows only). Account display name (or substring)
            to send from. Default: primary account. Ignored on macOS.

    Returns:
        JSON object with status "sent", subject, and recipients, or an error.
    """
    return await _run(
        "Error sending email",
        lambda: _require_backend().compose_email(to, subject, body, cc, bcc, html_body, account, send=True),
    )


@mcp.tool()
async def draft_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
) -> str:
    """Create an email draft without sending it.

    Works like send_email, but saves the message to the Drafts folder
    instead of sending immediately. Useful when the user wants to review
    a message before it goes out.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
        subject: The email subject line.
        body: The plain-text body of the email.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook
            renders the email as HTML.
        account: Optional (Windows only). Account display name (or substring)
            to draft from. Ignored on macOS.

    Returns:
        JSON object with status "draft_saved", subject, recipients, and the
        draft entry_id, or an error.
    """
    return await _run(
        "Error creating draft",
        lambda: _require_backend().compose_email(to, subject, body, cc, bcc, html_body, account, send=False),
    )


@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender, sender_name,
    received_time, unread status, and attachment info.

    Use the entry_id from results to read full content with read_email,
    or to perform actions like mark_as_read, move_email, or reply_email.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", "archive", or any custom folder name visible in
            list_folders output.
        count: Maximum number of emails to return. Default 10, max recommended 50.
        unread_only: If true, only return unread emails. Default false.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional (Windows only). Account display name (or substring)
            to target. Ignored on macOS.

    Returns:
        JSON array of email summary objects.
    """
    return await _run(
        "Error listing emails",
        lambda: _require_backend().list_emails(folder, count, unread_only, start_date, end_date, account),
    )


@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
    account: str = "",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (preferred, exact match) OR
    subject_search (finds most recent match by subject substring).

    Args:
        entry_id: The Outlook EntryID of the email. Most reliable way to
            identify a specific email. Get this from list_emails or
            search_emails results.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search when using subject_search. Ignored when
            entry_id is provided. Default "inbox".
        account: Optional (Windows only). Account display name (or substring)
            to target. Ignored on macOS.

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info).
    """
    return await _run(
        "Error reading email",
        lambda: _require_backend().read_email(entry_id, subject_search, folder, account),
    )


@mcp.tool()
async def mark_as_read(entry_id: str, account: str = "") -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: The Outlook EntryID of the email. Get this from list_emails
            or search_emails results.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object confirming the read status, or an error.
    """
    return await _run(
        "Error marking email as read",
        lambda: _require_backend().mark_as_read(entry_id, account),
    )


@mcp.tool()
async def mark_as_unread(entry_id: str, account: str = "") -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: The Outlook EntryID of the email. Get this from list_emails
            or search_emails results.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object confirming the unread status, or an error.
    """
    return await _run(
        "Error marking email as unread",
        lambda: _require_backend().mark_as_unread(entry_id, account),
    )


@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
    account: str = "",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid. Common use: archiving emails after processing.

    Args:
        entry_id: The Outlook EntryID of the email to move.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "archive", "inbox", "sent", "deleted"/
            "trash", "drafts", "junk"/"spam", or any custom folder name.
        account: Optional (Windows only). Account display name (or substring)
            to resolve the target folder in. Ignored on macOS.

    Returns:
        JSON object confirming the move, or an error.
    """
    return await _run(
        "Error moving email",
        lambda: _require_backend().move_email(entry_id, target_folder, account),
    )


@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html_body: str = "",
    account: str = "",
) -> str:
    """Reply to an email in Outlook.

    Creates and sends a reply, preserving the original message thread.
    Use reply_all=True to reply to all recipients (sender + CC list).

    Args:
        entry_id: The Outlook EntryID of the email to reply to.
        body: The plain-text reply message text. Prepended above the original
            message in the email thread. If html_body is also provided, both
            are set and Outlook will prefer the HTML version.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.
        html_body: Optional. HTML-formatted reply body. When provided, Outlook
            renders the reply as HTML. The plain-text body serves as fallback.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object confirming the reply was sent, or an error.
    """
    return await _run(
        "Error replying to email",
        lambda: _require_backend().reply_email(entry_id, body, reply_all, html_body, account, send=True),
    )


@mcp.tool()
async def draft_reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html_body: str = "",
    account: str = "",
) -> str:
    """Create a reply draft without sending it.

    Works like reply_email, but saves the reply to the Drafts folder
    instead of sending immediately. The reply preserves the original
    message thread, allowing the user to review it before sending.

    Args:
        entry_id: The Outlook EntryID of the email to reply to.
        body: The plain-text reply message text. Prepended above the original
            message in the email thread. If html_body is also provided, both
            are set and Outlook will prefer the HTML version.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.
        html_body: Optional. HTML-formatted reply body. When provided, Outlook
            renders the reply as HTML. The plain-text body serves as fallback.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object confirming the draft was saved, or an error.
    """
    return await _run(
        "Error creating reply draft",
        lambda: _require_backend().reply_email(entry_id, body, reply_all, html_body, account, send=False),
    )


@mcp.tool()
async def list_folders(folder: str = "", max_depth: int = 3, account: str = "") -> str:
    """List mail folders in the user's Outlook mailbox.

    When called with no folder argument, lists top-level folders. Provide a
    folder name to drill into its subfolders — use this to browse the full
    folder tree step by step (e.g. first call with no folder to see top-level,
    then call with folder="Inbox" to see Inbox children, then
    folder="Inbox/Projects" to go deeper).

    Folder names from this output can be used directly in list_emails,
    move_email, search_emails, etc. Use slash-delimited paths for nested
    folders (e.g. "Inbox/Receipts/2026").

    NOTE: On macOS, Outlook's AppleScript dictionary exposes folders as a
    flat collection without parent links, so drill-down filters by the last
    path segment and depth beyond 1 cannot be honored.

    Args:
        folder: Optional. Folder to list children of. Supports folder names
            ("Inbox"), slash paths ("Inbox/Receipts"), or built-in names
            ("sent", "drafts"). When empty, lists from the mailbox root.
        max_depth: How many levels deep to recurse below the starting folder.
            Default 3. Set to 1 to see only immediate children. (Windows only;
            ignored on macOS.)
        account: Optional (Windows only). Account display name (or substring)
            to target. Ignored on macOS.

    Returns:
        JSON array of folder objects with name, full_path, item_count,
        unread_count, and subfolders (if any).
    """
    return await _run(
        "Error listing folders",
        lambda: _require_backend().list_folders(folder, max_depth, account),
    )


@mcp.tool()
async def search_emails(
    query: str,
    folder: str = "inbox",
    count: int = 10,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    """Search for emails in Outlook using text search.

    Searches email subjects (and bodies on Windows) using Outlook's filtering.
    Results are sorted by received time (newest first). Each result includes
    entry_id for further operations.

    Args:
        query: The search term (case-insensitive substring match).
            Examples: "budget report", "meeting notes", "quarterly".
        folder: Folder to search in. Default "inbox". Supports same names as
            list_emails.
        count: Maximum results to return. Default 10.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional (Windows only). Account display name (or substring)
            to target. Ignored on macOS.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    return await _run(
        "Error searching emails",
        lambda: _require_backend().search_emails(query, folder, count, start_date, end_date, account),
    )


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
    account: str = "",
    folder: str = "",
) -> str:
    """List upcoming calendar events from Outlook.

    Returns a JSON array of event summaries within a date range, sorted by
    start time. Includes recurring event occurrences. Each summary has
    entry_id, subject, start, end, duration, location, organizer, attendees,
    and status info.

    Use entry_id from results with get_event, update_event, delete_event,
    or respond_to_meeting (Windows).

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return. Default 20.
        account: Optional (Windows only). Account display name (or substring)
            to target. Ignored on macOS.
        folder: Optional (Windows only). Calendar folder name or path (e.g.
            "Team-Kalender", "Calendar/Team-Kalender"). Default: primary
            calendar. Ignored on macOS.

    Returns:
        JSON array of event summary objects.
    """
    return await _run(
        "Error listing events",
        lambda: _require_backend().list_events(start_date, end_date, count, account, folder),
    )


@mcp.tool()
async def get_event(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, recurrence status, reminders, and response status.

    Args:
        entry_id: The Outlook EntryID of the event. Get this from list_events
            or search_events results.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object with full event details.
    """
    return await _run(
        "Error reading event",
        lambda: _require_backend().get_event(entry_id, account),
    )


@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
    account: str = "",
    folder: str = "",
) -> str:
    """Create a personal calendar appointment (no attendees).

    Creates and saves an appointment on the user's calendar. This is a
    personal event — no meeting invitations are sent. Use create_meeting
    instead if you need to invite attendees.

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date:
            "2026-02-25".
        end: End time in ISO 8601 format. For all-day events, use the next
            day: "2026-02-26".
        location: Optional. Event location (e.g. "Conference Room A",
            "Microsoft Teams Meeting").
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.
        account: Optional (Windows only). Account display name (or substring)
            to create the event in. Ignored on macOS.
        folder: Optional (Windows only). Calendar folder name or path.
            Default: primary calendar. Ignored on macOS.

    Returns:
        JSON object with status "created", subject, start, end, entry_id,
        or an error.
    """
    if err := _validate_subject(subject):
        return err
    return await _run(
        "Error creating event",
        lambda: _require_backend().create_event(subject, start, end, location, body, all_day, reminder_minutes, account, folder),
    )


@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    account: str = "",
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and immediately sends meeting requests to
    all specified attendees. The meeting will appear on the organizer's
    calendar and attendees will receive an invitation they can accept,
    decline, or tentatively accept.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location (e.g. "Teams", "Room 301").
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.
        account: Optional (Windows only). Account display name (or substring)
            to send from. Ignored on macOS.

    Returns:
        JSON object confirming the meeting was created and invitations sent.
    """
    if err := _validate_subject(subject):
        return err
    return await _run(
        "Error creating meeting",
        lambda: _require_backend().create_meeting(
            subject,
            start,
            end,
            required_attendees,
            location,
            body,
            optional_attendees,
            account,
        ),
    )


@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
    account: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged. For meetings
    you organize, attendees will receive an update notification.

    Args:
        entry_id: The Outlook EntryID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object with the updated event details, or an error.
    """
    if err := _validate_subject(subject, allow_skip=True):
        return err
    return await _run(
        "Error updating event",
        lambda: _require_backend().update_event(entry_id, subject, start, end, location, body, account),
    )


@mcp.tool()
async def delete_event(entry_id: str, account: str = "") -> str:
    """Delete a calendar event or cancel a meeting.

    For personal appointments, the event is simply deleted. For meetings
    you organized (Windows), this cancels the meeting and sends cancellation
    notices to all attendees. For meetings you received, this removes the
    event from your calendar.

    Args:
        entry_id: The Outlook EntryID of the event to delete/cancel.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object confirming the deletion/cancellation, or an error.
    """
    return await _run(
        "Error deleting event",
        lambda: _require_backend().delete_event(entry_id, account),
    )


@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
    account: str = "",
    folder: str = "",
) -> str:
    """Search for calendar events by keyword.

    Searches event subjects within a date range. Results are sorted by
    start time. Includes recurring event occurrences.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return. Default 10.
        account: Optional (Windows only). Account display name (or substring)
            to target. Ignored on macOS.
        folder: Optional (Windows only). Calendar folder name or path.
            Ignored on macOS.

    Returns:
        JSON array of matching event summaries.
    """
    return await _run(
        "Error searching events",
        lambda: _require_backend().search_events(query, start_date, end_date, count, account, folder),
    )


# =====================================================================
# TASK TOOLS
# =====================================================================


@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
    account: str = "",
) -> str:
    """List tasks from the Outlook Tasks folder.

    Returns a JSON array of task summaries sorted by due date. Each task
    includes entry_id, subject, status, percent_complete, due_date,
    importance, and categories.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.
        account: Optional (Windows only). Account display name (or substring)
            to target. Ignored on macOS.

    Returns:
        JSON array of task summary objects.
    """
    return await _run(
        "Error listing tasks",
        lambda: _require_backend().list_tasks(include_completed, count, account),
    )


@mcp.tool()
async def get_task(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The Outlook EntryID of the task.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object with full task details including body.
    """
    return await _run(
        "Error reading task",
        lambda: _require_backend().get_task(entry_id, account),
    )


@mcp.tool()
async def create_task(
    subject: str,
    body: str = "",
    due_date: str = "",
    importance: str = "normal",
    reminder_minutes: int = 0,
    account: str = "",
) -> str:
    """Create a new task in Outlook.

    Args:
        subject: The task title.
        body: Optional. Task description or notes.
        due_date: Optional. Due date in ISO 8601 format (e.g. "2026-03-01").
        importance: Optional. "low", "normal" (default), or "high".
        reminder_minutes: Optional. Minutes before due date to remind.
            Default 0 (no reminder). (Windows only; ignored on macOS.)
        account: Optional (Windows only). Account display name (or substring)
            to create the task in. Ignored on macOS.

    Returns:
        JSON object with status "created", subject, entry_id, and due_date.
    """
    if err := _validate_subject(subject):
        return err
    return await _run(
        "Error creating task",
        lambda: _require_backend().create_task(subject, body, due_date, importance, reminder_minutes, account),
    )


@mcp.tool()
async def complete_task(entry_id: str, account: str = "") -> str:
    """Mark a task as complete.

    Sets the task status to complete and percent to 100%.

    Args:
        entry_id: The Outlook EntryID of the task.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object confirming completion.
    """
    return await _run(
        "Error completing task",
        lambda: _require_backend().complete_task(entry_id, account),
    )


@mcp.tool()
async def delete_task(entry_id: str, account: str = "") -> str:
    """Delete a task from Outlook.

    Args:
        entry_id: The Outlook EntryID of the task to delete.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object confirming the deletion.
    """
    return await _run(
        "Error deleting task",
        lambda: _require_backend().delete_task(entry_id, account),
    )


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================


@mcp.tool()
async def list_attachments(entry_id: str, account: str = "") -> str:
    """List all attachments on an email or calendar event.

    Args:
        entry_id: The EntryID of the email or event to check for attachments.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON array of attachment objects with index, filename, and size.
    """
    return await _run(
        "Error listing attachments",
        lambda: _require_backend().list_attachments(entry_id, account),
    )


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
    account: str = "",
) -> str:
    """Save an attachment from an email or event to disk.

    Downloads the specified attachment to a local directory.

    Args:
        entry_id: The EntryID of the email or event containing the attachment.
        attachment_index: Which attachment to save (1-based index). Default 1
            (first attachment). Use list_attachments to see available indices.
        save_directory: Directory to save the file to. Default: user's
            Downloads folder.
        account: Optional (Windows only). Account display name (or substring).
            Only needed if entry_id is ambiguous across stores.

    Returns:
        JSON object with the saved file path, or an error.
    """
    return await _run(
        "Error saving attachment",
        lambda: _require_backend().save_attachment(entry_id, attachment_index, save_directory, account),
    )


# =====================================================================
# OUT OF OFFICE TOOLS
# =====================================================================


@mcp.tool()
async def set_out_of_office(
    enabled: bool = True,
    message: str = "",
    account: str = "",
) -> str:
    """Enable or disable the Out of Office (auto-reply) status.

    On macOS, optionally sets the auto-reply message body. On Windows, the
    Outlook Object Model cannot set the message text — if 'message' is
    provided there, only the OOF state is changed and the message must be
    configured in Outlook's UI (File > Automatic Replies).

    Args:
        enabled: True to enable Out of Office, False to disable it.
        message: Optional. The auto-reply message body (macOS only; ignored
            by the COM API on Windows).
        account: Optional (Windows only). Account display name (or substring)
            to target. Ignored on macOS.

    Returns:
        JSON object confirming the new OOF status, or an error.
    """
    return await _run(
        "Error setting OOF status",
        lambda: _require_backend().set_out_of_office(enabled, message, account),
    )


# =====================================================================
# WINDOWS-ONLY TOOLS
# Registered conditionally in set_backend() based on backend capabilities.
# =====================================================================


async def list_accounts() -> str:
    """List all Outlook accounts (stores) configured in the profile.

    Returns a JSON array of account objects with display_name, store_id,
    and is_default. Use the display_name (or a unique substring) as the
    'account' parameter in other tools to target a specific account.

    Returns:
        JSON array of account objects.
    """
    return await _run(
        "Error listing accounts",
        lambda: _require_backend().list_accounts(),
    )


async def respond_to_meeting(
    entry_id: str,
    response: str,
    account: str = "",
) -> str:
    """Respond to a meeting invitation (accept, decline, or tentative).

    Sends your response to the meeting organizer. The meeting will be
    added to (or updated on) your calendar accordingly.

    Args:
        entry_id: The Outlook EntryID of the meeting to respond to.
            Get this from list_events or search_events.
        response: Your response. Must be one of: "accept", "decline",
            or "tentative".
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object confirming your response, or an error.
    """
    return await _run(
        "Error responding to meeting",
        lambda: _require_backend().respond_to_meeting(entry_id, response, account),
    )


async def list_categories(account: str = "") -> str:
    """List all available Outlook categories.

    Returns the color categories configured in the user's Outlook profile.
    These can be applied to emails, events, tasks, and other items.

    Args:
        account: Optional. Account display name (or substring). Categories
            are profile-wide; the parameter is accepted for consistency.

    Returns:
        JSON array of category objects with name and color index.
    """
    return await _run(
        "Error listing categories",
        lambda: _require_backend().list_categories(account),
    )


async def set_category(
    entry_id: str,
    categories: str,
    account: str = "",
) -> str:
    """Set categories on an email, event, or task.

    Replaces any existing categories on the item. Use comma-separated
    values for multiple categories.

    Args:
        entry_id: The EntryID of the item to categorize.
        categories: Category name(s), comma-separated. Example:
            "Important" or "Work, Follow-up". Use an empty string to
            clear all categories.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object confirming the categories applied.
    """
    return await _run(
        "Error setting categories",
        lambda: _require_backend().set_category(entry_id, categories, account),
    )


async def list_rules(account: str = "") -> str:
    """List all mail rules in Outlook.

    Returns the configured inbox rules with their names and enabled status.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account.

    Returns:
        JSON array of rule objects with name, enabled status, and index.
    """
    return await _run(
        "Error listing rules",
        lambda: _require_backend().list_rules(account),
    )


async def toggle_rule(
    rule_name: str,
    enabled: bool,
    account: str = "",
) -> str:
    """Enable or disable a mail rule by name.

    CAUTION: This modifies live mail rules immediately. Confirm the rule name
    with list_rules before calling.

    Args:
        rule_name: The exact name of the rule to toggle. Use list_rules
            to see available rule names.
        enabled: True to enable the rule, False to disable it.
        account: Optional. Account display name (or substring) to target.
            Default: primary account.

    Returns:
        JSON object confirming the new rule status, or an error.
    """
    return await _run(
        "Error toggling rule",
        lambda: _require_backend().toggle_rule(rule_name, enabled, account),
    )


async def get_out_of_office(account: str = "") -> str:
    """Check the current Out of Office (auto-reply) status.

    Returns whether Out of Office is currently enabled.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account.

    Returns:
        JSON object with the OOF status.
    """
    return await _run(
        "Error checking OOF status",
        lambda: _require_backend().get_out_of_office(account),
    )


# =====================================================================
# Entry point
# =====================================================================


def main() -> None:
    if backend is None:
        raise RuntimeError("No backend configured. Use outlook_desktop_mcp.entrypoint, which selects the platform backend.")
    logger.info("Starting Outlook Desktop MCP server (backend: %s)...", backend.name)
    backend.start()
    logger.info("Backend ready. Starting MCP stdio transport...")
    try:
        mcp.run(transport="stdio")
    finally:
        backend.stop()
