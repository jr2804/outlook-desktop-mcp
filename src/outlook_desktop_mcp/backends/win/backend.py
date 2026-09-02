"""Windows COM backend — automates classic Outlook Desktop via pywin32.

All methods run their COM work on the bridge thread via ``bridge.call`` and
return pydantic models. Handled failures raise :class:`BackendError`.

The implementation lives in domain modules (email, calendar, tasks,
attachments, categories, rules, oof); this module is the facade that wires
them into the ``Backend`` contract.
"""

from __future__ import annotations

import asyncio
from typing import Any

from outlook_desktop_mcp.backends.base import Backend
from outlook_desktop_mcp.backends.win._types import Namespace
from outlook_desktop_mcp.backends.win.attachments import (
    list_attachments as _list_attachments,
)
from outlook_desktop_mcp.backends.win.attachments import (
    save_attachment as _save_attachment,
)
from outlook_desktop_mcp.backends.win.bridge import OutlookBridge
from outlook_desktop_mcp.backends.win.calendar import (
    create_event as _create_event,
)
from outlook_desktop_mcp.backends.win.calendar import (
    create_meeting as _create_meeting,
)
from outlook_desktop_mcp.backends.win.calendar import (
    delete_event as _delete_event,
)
from outlook_desktop_mcp.backends.win.calendar import (
    get_event as _get_event,
)
from outlook_desktop_mcp.backends.win.calendar import (
    list_events as _list_events,
)
from outlook_desktop_mcp.backends.win.calendar import (
    respond_to_meeting as _respond_to_meeting,
)
from outlook_desktop_mcp.backends.win.calendar import (
    search_events as _search_events,
)
from outlook_desktop_mcp.backends.win.calendar import (
    update_event as _update_event,
)
from outlook_desktop_mcp.backends.win.categories import (
    list_categories as _list_categories,
)
from outlook_desktop_mcp.backends.win.categories import (
    set_category as _set_category,
)
from outlook_desktop_mcp.backends.win.email import (
    compose_email as _compose_email,
)
from outlook_desktop_mcp.backends.win.email import (
    list_accounts as _list_accounts,
)
from outlook_desktop_mcp.backends.win.email import (
    list_emails as _list_emails,
)
from outlook_desktop_mcp.backends.win.email import (
    list_folders as _list_folders,
)
from outlook_desktop_mcp.backends.win.email import (
    mark_as_read as _mark_as_read,
)
from outlook_desktop_mcp.backends.win.email import (
    mark_as_unread as _mark_as_unread,
)
from outlook_desktop_mcp.backends.win.email import (
    move_email as _move_email,
)
from outlook_desktop_mcp.backends.win.email import (
    read_email as _read_email,
)
from outlook_desktop_mcp.backends.win.email import (
    reply_email as _reply_email,
)
from outlook_desktop_mcp.backends.win.email import (
    search_emails as _search_emails,
)
from outlook_desktop_mcp.backends.win.errors import format_com_error
from outlook_desktop_mcp.backends.win.helpers import _require_store
from outlook_desktop_mcp.backends.win.oof import (
    get_out_of_office as _get_out_of_office,
)
from outlook_desktop_mcp.backends.win.oof import (
    set_out_of_office as _set_out_of_office,
)
from outlook_desktop_mcp.backends.win.rules import (
    list_rules as _list_rules,
)
from outlook_desktop_mcp.backends.win.rules import (
    toggle_rule as _toggle_rule,
)
from outlook_desktop_mcp.backends.win.tasks import (
    complete_task as _complete_task,
)
from outlook_desktop_mcp.backends.win.tasks import (
    create_task as _create_task,
)
from outlook_desktop_mcp.backends.win.tasks import (
    delete_task as _delete_task,
)
from outlook_desktop_mcp.backends.win.tasks import (
    get_task as _get_task,
)
from outlook_desktop_mcp.backends.win.tasks import (
    list_tasks as _list_tasks,
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

    def format_unexpected_error(self, action: str, e: Exception) -> str:  # noqa: PLR6301
        """Surface COM HRESULT detail (Windows-only override)."""
        return f"{action}: {format_com_error(e)}"

    # --- accounts ---

    async def list_accounts(self) -> list[AccountInfo]:
        return await _list_accounts(self.bridge)

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
        return await _compose_email(self.bridge, to, subject, body, cc, bcc, html_body, account, send)

    async def list_emails(
        self,
        folder: str,
        count: int,
        unread_only: bool,
        start_date: str,
        end_date: str,
        account: str,
    ) -> list[EmailSummary]:
        return await _list_emails(self.bridge, folder, count, unread_only, start_date, end_date, account)

    async def read_email(
        self,
        entry_id: str,
        subject_search: str,
        folder: str,
        account: str,
    ) -> EmailFull:
        return await _read_email(self.bridge, entry_id, subject_search, folder, account)

    async def mark_as_read(self, entry_id: str, account: str) -> ItemStatusResult:
        return await _mark_as_read(self.bridge, self._get_item, entry_id, account)

    async def mark_as_unread(self, entry_id: str, account: str) -> ItemStatusResult:
        return await _mark_as_unread(self.bridge, self._get_item, entry_id, account)

    async def move_email(self, entry_id: str, target_folder: str, account: str) -> MovedResult:
        return await _move_email(self.bridge, entry_id, target_folder, account)

    async def reply_email(
        self,
        entry_id: str,
        body: str,
        reply_all: bool,
        html_body: str,
        account: str,
        send: bool,
    ) -> ReplySentResult | ReplyDraftSavedResult:
        return await _reply_email(self.bridge, self._get_item, entry_id, body, reply_all, html_body, account, send)

    async def list_folders(self, folder: str, max_depth: int, account: str) -> list[FolderInfo]:
        return await _list_folders(self.bridge, folder, max_depth, account)

    async def search_emails(
        self,
        query: str,
        folder: str,
        count: int,
        start_date: str,
        end_date: str,
        account: str,
    ) -> list[EmailSummary]:
        return await _search_emails(self.bridge, query, folder, count, start_date, end_date, account)

    # --- calendar ---

    async def list_events(
        self,
        start_date: str,
        end_date: str,
        count: int,
        account: str,
        folder: str,
    ) -> list[EventSummary]:
        return await _list_events(self.bridge, start_date, end_date, count, account, folder)

    async def get_event(self, entry_id: str, account: str) -> EventFull:
        return await _get_event(self.bridge, self._get_item, entry_id, account)

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
        return await _create_event(self.bridge, subject, start, end, location, body, all_day, reminder_minutes, account, folder)

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
        return await _create_meeting(self.bridge, subject, start, end, required_attendees, location, body, optional_attendees, account)

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
        return await _update_event(self.bridge, self._get_item, entry_id, subject, start, end, location, body, account)

    async def delete_event(self, entry_id: str, account: str) -> ItemStatusResult:
        return await _delete_event(self.bridge, self._get_item, entry_id, account)

    async def respond_to_meeting(self, entry_id: str, response: str, account: str) -> MeetingResponseResult:
        return await _respond_to_meeting(self.bridge, self._get_item, entry_id, response, account)

    async def search_events(
        self,
        query: str,
        start_date: str,
        end_date: str,
        count: int,
        account: str,
        folder: str,
    ) -> list[EventSummary]:
        return await _search_events(self.bridge, query, start_date, end_date, count, account, folder)

    # --- tasks ---

    async def list_tasks(
        self,
        include_completed: bool,
        count: int,
        account: str,
    ) -> list[TaskSummary]:
        return await _list_tasks(self.bridge, include_completed, count, account)

    async def get_task(self, entry_id: str, account: str) -> TaskFull:
        return await _get_task(self.bridge, self._get_item, entry_id, account)

    async def create_task(
        self,
        subject: str,
        body: str,
        due_date: str,
        importance: str,
        reminder_minutes: int,
        account: str,
    ) -> TaskCreatedResult:
        return await _create_task(self.bridge, subject, body, due_date, importance, reminder_minutes, account)

    async def complete_task(self, entry_id: str, account: str) -> ItemStatusResult:
        return await _complete_task(self.bridge, self._get_item, entry_id, account)

    async def delete_task(self, entry_id: str, account: str) -> ItemStatusResult:
        return await _delete_task(self.bridge, self._get_item, entry_id, account)

    # --- attachments ---

    async def list_attachments(self, entry_id: str, account: str) -> list[AttachmentInfo]:
        return await _list_attachments(self.bridge, self._get_item, entry_id, account)

    async def save_attachment(
        self,
        entry_id: str,
        attachment_index: int,
        save_directory: str,
        account: str,
    ) -> AttachmentSavedResult:
        return await _save_attachment(self.bridge, self._get_item, entry_id, attachment_index, save_directory, account)

    # --- categories ---

    async def list_categories(self, account: str) -> list[CategoryInfo]:
        return await _list_categories(self.bridge, account)

    async def set_category(self, entry_id: str, categories: str, account: str) -> CategoriesSetResult:
        return await _set_category(self.bridge, self._get_item, entry_id, categories, account)

    # --- rules ---

    async def list_rules(self, account: str) -> list[RuleInfo]:
        return await _list_rules(self.bridge, account)

    async def toggle_rule(self, rule_name: str, enabled: bool, account: str) -> RuleToggledResult:
        return await _toggle_rule(self.bridge, rule_name, enabled, account)

    # --- out of office ---

    async def get_out_of_office(self, account: str) -> OofStatus:
        return await _get_out_of_office(self.bridge, account)

    async def set_out_of_office(self, enabled: bool, message: str, account: str) -> OofSetResult:
        return await _set_out_of_office(self.bridge, enabled, message, account)

    # --- shared COM helpers (run on the bridge thread) ---

    @staticmethod
    def _get_item(namespace: Namespace, entry_id: str, account: str) -> Any:
        """Fetch an item by EntryID, optionally scoped to a store."""
        if account:
            store = _require_store(namespace, account)
            return namespace.GetItemFromID(entry_id, store.StoreID)
        return namespace.GetItemFromID(entry_id)
