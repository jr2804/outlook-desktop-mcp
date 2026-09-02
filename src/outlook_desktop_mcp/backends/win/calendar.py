"""Calendar/meeting operations on Windows via COM."""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime, timedelta
from typing import Any, cast

from outlook_desktop_mcp.backends.base import BackendError
from outlook_desktop_mcp.backends.win._types import Appointment, Folder, Namespace, Store
from outlook_desktop_mcp.backends.win.formatting import (
    format_event_full,
    format_event_summary,
)
from outlook_desktop_mcp.backends.win.helpers import (
    _OL_CLASS_APPOINTMENT,
    _jet_datetime,
    _parse_date,
    _require_class,
    _require_store,
    _resolve_folder,
)
from outlook_desktop_mcp.models import (
    EventCreatedResult,
    EventFull,
    EventSummary,
    EventUpdatedResult,
    ItemStatusResult,
    MeetingResponseResult,
    MeetingSentResult,
)
from outlook_desktop_mcp.tools._folder_constants import (
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_OPTIONAL,
    OL_REQUIRED,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_RESPONSE_TENTATIVE,
)


async def list_events(
    bridge: Any,
    start_date: str,
    end_date: str,
    count: int,
    account: str,
    folder: str,
) -> list[EventSummary]:
    def _list(outlook: Any, namespace: Namespace) -> list[EventSummary]:
        effective_count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendar = _resolve_calendar(namespace, folder, store)
        items = calendar.Items

        # CRITICAL ORDER: Sort BEFORE IncludeRecurrences BEFORE Restrict
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now()
        end = _parse_date(end_date) if end_date else start + timedelta(days=7)

        restrict = f"[Start] >= '{_jet_datetime(start)}' AND [Start] <= '{_jet_datetime(end)}'"
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

    return await bridge.call(_list)


async def get_event(bridge: Any, get_item: Callable, entry_id: str, account: str) -> EventFull:
    def _get(outlook: Any, namespace: Namespace) -> EventFull:
        item = get_item(namespace, entry_id, account)
        return EventFull.model_validate(format_event_full(item))

    return await bridge.call(_get)


async def create_event(
    bridge: Any,
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
    def _create(outlook: Any, namespace: Namespace) -> EventCreatedResult:
        appt: Appointment = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        if account:
            store = _require_store(namespace, account)
            cal = _resolve_calendar(namespace, folder, store)
            appt = cast(Appointment, appt.Move(cal))
        elif folder:
            cal = _resolve_calendar(namespace, folder, namespace.DefaultStore)
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

    return await bridge.call(_create)


async def create_meeting(
    bridge: Any,
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str,
    body: str,
    optional_attendees: str,
    account: str,
) -> MeetingSentResult:
    def _create(outlook: Any, namespace: Namespace) -> MeetingSentResult:
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

    return await bridge.call(_create)


async def update_event(
    bridge: Any,
    get_item: Callable,
    entry_id: str,
    subject: str,
    start: str,
    end: str,
    location: str,
    body: str,
    account: str,
) -> EventUpdatedResult:
    def _update(outlook: Any, namespace: Namespace) -> EventUpdatedResult:
        item = cast(Appointment, get_item(namespace, entry_id, account))
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

    return await bridge.call(_update)


async def delete_event(bridge: Any, get_item: Callable, entry_id: str, account: str) -> ItemStatusResult:
    def _delete(outlook: Any, namespace: Namespace) -> ItemStatusResult:
        item = get_item(namespace, entry_id, account)
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

    return await bridge.call(_delete)


async def respond_to_meeting(
    bridge: Any,
    get_item: Callable,
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

    def _respond(outlook: Any, namespace: Namespace) -> MeetingResponseResult:
        item = get_item(namespace, entry_id, account)
        _require_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item")
        subject = item.Subject
        response_item = item.Respond(response_map[response_lower])
        response_item.Send()
        return MeetingResponseResult(response=response_lower, subject=subject)

    return await bridge.call(_respond)


async def search_events(
    bridge: Any,
    query: str,
    start_date: str,
    end_date: str,
    count: int,
    account: str,
    folder: str,
) -> list[EventSummary]:
    def _search(outlook: Any, namespace: Namespace) -> list[EventSummary]:
        effective_count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendar = _resolve_calendar(namespace, folder, store)
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now() - timedelta(days=30)
        end = _parse_date(end_date) if end_date else datetime.now() + timedelta(days=30)

        restrict = f"[Start] >= '{_jet_datetime(start)}' AND [Start] <= '{_jet_datetime(end)}'"
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

    return await bridge.call(_search)


def _resolve_calendar(namespace: Namespace, folder: str, store: Store) -> Folder:
    """Resolve the calendar folder: named path or the store's default."""
    if folder:
        calendar = _resolve_folder(namespace, folder, store)
        if not calendar:
            raise BackendError(f"Calendar folder '{folder}' not found")
        return calendar
    return store.GetDefaultFolder(OL_FOLDER_CALENDAR)
