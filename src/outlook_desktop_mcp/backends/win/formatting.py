"""Helpers for extracting and formatting Outlook item data."""

import os
import re
from collections.abc import Callable
from typing import Any

from outlook_desktop_mcp.tools._folder_constants import (
    BUSY_STATUS_NAMES,
    IMPORTANCE_NAMES,
    MEETING_STATUS_NAMES,
    RESPONSE_NAMES,
    TASK_STATUS_NAMES,
)


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", "", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def format_email_full(item: Any, body_max_length: int = 5000) -> dict:
    """Extract full email details including body."""
    result = format_email_summary(item)
    result["to"] = item.To or ""
    result["cc"] = item.CC or ""
    result["body"] = truncate(item.Body or "", body_max_length)
    return result


def format_email_summary(item: Any) -> dict:
    """Extract key fields from an Outlook MailItem into a dict."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "sender": getattr(item, "SenderEmailAddress", "unknown"),
        "sender_name": getattr(item, "SenderName", "unknown"),
        "received_time": str(item.ReceivedTime),
        "unread": bool(item.UnRead),
        "has_attachments": bool(item.Attachments.Count > 0),
        "attachment_count": item.Attachments.Count,
    }


def format_event_full(item: Any, body_max_length: int = 5000) -> dict:
    """Full event details including body."""
    result = format_event_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["reminder_minutes"] = item.ReminderMinutesBeforeStart if item.ReminderSet else None
    result["categories"] = item.Categories or ""
    result["response_status"] = RESPONSE_NAMES.get(item.ResponseStatus, "unknown")
    return result


# --- Calendar formatting ---


def _address_info_enabled() -> bool:
    """Whether it is safe to read Object-Model-Guard protected properties.

    ``Organizer``/``RequiredAttendees``/``OptionalAttendees`` are *address
    information*: reading them pops a modal "A program is trying to access email
    address information" prompt unless an admin policy auto-approves it. With
    nobody at the desk (cron, headless sync) that prompt blocks the COM thread
    forever — the call never returns and never raises, so ``try/except`` alone
    cannot save it.

    Set ``OUTLOOK_MCP_ADDRESS_FIELDS=1`` to opt back in once the machine has the
    Object Model Guard suppression policy applied. Default: skip these fields.
    """
    return os.environ.get("OUTLOOK_MCP_ADDRESS_FIELDS", "").strip().lower() in {"1", "true", "yes"}


def _guarded(getter: Callable[[], Any], default: Any) -> Any:
    """Read a COM property that may be blocked by Outlook's Object Model Guard.

    A *denied* read raises ``com_error(-2147467259, 'Unspecified error')``; an
    *unanswered* read hangs. These fields are informational only, so never let
    one of them abort or stall the whole item — that failure mode silently
    emptied ``list_events`` and made the absence sync re-create every vacation
    day as a duplicate.
    """
    if not _address_info_enabled():
        return default
    try:
        return getter()
    except Exception:  # noqa: BLE001 - guard-blocked field, fall back to default
        return default


def format_event_summary(item: Any) -> dict:
    """Extract key fields from an Outlook AppointmentItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "start": str(item.Start),
        "end": str(item.End),
        "duration": item.Duration,
        "location": item.Location or "",
        "organizer": _guarded(lambda: item.Organizer or "", ""),
        "is_recurring": bool(item.IsRecurring),
        "all_day": bool(item.AllDayEvent),
        "busy_status": BUSY_STATUS_NAMES.get(item.BusyStatus, "unknown"),
        "meeting_status": MEETING_STATUS_NAMES.get(item.MeetingStatus, "unknown"),
        "required_attendees": _guarded(lambda: item.RequiredAttendees or "", ""),
        "optional_attendees": _guarded(lambda: item.OptionalAttendees or "", ""),
    }


def format_task_full(item: Any, body_max_length: int = 5000) -> dict:
    """Full task details including body."""
    result = format_task_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["date_completed"] = str(item.DateCompleted) if item.Complete else None
    return result


def truncate(text: str, max_length: int = 2000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


# --- Task formatting ---


def format_task_summary(item: Any) -> dict:
    """Extract key fields from an Outlook TaskItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "status": TASK_STATUS_NAMES.get(item.Status, "unknown"),
        "percent_complete": item.PercentComplete,
        "due_date": str(item.DueDate) if str(item.DueDate) != "01/01/4501" else None,
        "start_date": str(item.StartDate) if str(item.StartDate) != "01/01/4501" else None,
        "importance": IMPORTANCE_NAMES.get(item.Importance, "normal"),
        "complete": bool(item.Complete),
        "categories": item.Categories or "",
        "owner": item.Owner or "",
    }
