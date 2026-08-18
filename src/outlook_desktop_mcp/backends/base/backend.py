"""Backend contract for platform-specific Outlook automation.

A backend implements every operation as an ``async`` method returning pydantic
models from :mod:`outlook_desktop_mcp.models`. Handled failures raise
:class:`BackendError`; unexpected exceptions propagate to the server layer,
which converts them into ``Error`` JSON.

The ``account`` parameter is part of every signature for parity. Backends that
cannot select accounts (AppleScript) accept and ignore it.
"""

from __future__ import annotations

from abc import ABC, abstractmethod

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


class BackendError(Exception):
    """A handled, client-safe backend failure (folder not found, wrong type…)."""


class Backend(ABC):
    """Platform backend interface. One instance per server process."""

    name: str = "abstract"

    # Capability flags — the server only registers matching tools.
    supports_accounts: bool = False
    supports_rules: bool = False
    supports_categories: bool = False
    supports_meeting_response: bool = False
    supports_oof_status_query: bool = False

    # --- lifecycle ---

    def start(self) -> None:  # noqa: B027 - optional override
        """Prepare the backend (COM bridge thread, subprocess…)."""

    def stop(self) -> None:  # noqa: B027 - optional override
        """Release backend resources."""

    # --- email ---

    @abstractmethod
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
    ) -> SentResult | DraftSavedResult: ...

    @abstractmethod
    async def list_emails(
        self,
        folder: str,
        count: int,
        unread_only: bool,
        start_date: str,
        end_date: str,
        account: str,
    ) -> list[EmailSummary]: ...

    @abstractmethod
    async def read_email(
        self,
        entry_id: str,
        subject_search: str,
        folder: str,
        account: str,
    ) -> EmailFull: ...

    @abstractmethod
    async def mark_as_read(self, entry_id: str, account: str) -> ItemStatusResult: ...

    @abstractmethod
    async def mark_as_unread(self, entry_id: str, account: str) -> ItemStatusResult: ...

    @abstractmethod
    async def move_email(self, entry_id: str, target_folder: str, account: str) -> MovedResult: ...

    @abstractmethod
    async def reply_email(
        self,
        entry_id: str,
        body: str,
        reply_all: bool,
        account: str,
        send: bool,
    ) -> ReplySentResult | ReplyDraftSavedResult: ...

    @abstractmethod
    async def list_folders(self, folder: str, max_depth: int, account: str) -> list[FolderInfo]: ...

    @abstractmethod
    async def search_emails(
        self,
        query: str,
        folder: str,
        count: int,
        start_date: str,
        end_date: str,
        account: str,
    ) -> list[EmailSummary]: ...

    # --- calendar ---

    @abstractmethod
    async def list_events(
        self,
        start_date: str,
        end_date: str,
        count: int,
        account: str,
        folder: str,
    ) -> list[EventSummary]: ...

    @abstractmethod
    async def get_event(self, entry_id: str, account: str) -> EventFull: ...

    @abstractmethod
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
    ) -> EventCreatedResult: ...

    @abstractmethod
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
    ) -> MeetingSentResult: ...

    @abstractmethod
    async def update_event(
        self,
        entry_id: str,
        subject: str,
        start: str,
        end: str,
        location: str,
        body: str,
        account: str,
    ) -> EventUpdatedResult: ...

    @abstractmethod
    async def delete_event(self, entry_id: str, account: str) -> ItemStatusResult: ...

    @abstractmethod
    async def search_events(
        self,
        query: str,
        start_date: str,
        end_date: str,
        count: int,
        account: str,
        folder: str,
    ) -> list[EventSummary]: ...

    # --- tasks ---

    @abstractmethod
    async def list_tasks(
        self,
        include_completed: bool,
        count: int,
        account: str,
    ) -> list[TaskSummary]: ...

    @abstractmethod
    async def get_task(self, entry_id: str, account: str) -> TaskFull: ...

    @abstractmethod
    async def create_task(
        self,
        subject: str,
        body: str,
        due_date: str,
        importance: str,
        reminder_minutes: int,
        account: str,
    ) -> TaskCreatedResult: ...

    @abstractmethod
    async def complete_task(self, entry_id: str, account: str) -> ItemStatusResult: ...

    @abstractmethod
    async def delete_task(self, entry_id: str, account: str) -> ItemStatusResult: ...

    # --- attachments ---

    @abstractmethod
    async def list_attachments(self, entry_id: str, account: str) -> list[AttachmentInfo]: ...

    @abstractmethod
    async def save_attachment(
        self,
        entry_id: str,
        attachment_index: int,
        save_directory: str,
        account: str,
    ) -> AttachmentSavedResult: ...

    # --- out of office ---

    @abstractmethod
    async def set_out_of_office(self, enabled: bool, message: str, account: str) -> OofSetResult: ...

    # --- Windows-only capabilities (default: unsupported) ---

    async def list_accounts(self) -> list[AccountInfo]:
        raise BackendError("list_accounts is not supported on this platform")

    async def respond_to_meeting(self, entry_id: str, response: str, account: str) -> MeetingResponseResult:
        raise BackendError("respond_to_meeting is not supported on this platform")

    async def list_categories(self, account: str) -> list[CategoryInfo]:
        raise BackendError("list_categories is not supported on this platform")

    async def set_category(self, entry_id: str, categories: str, account: str) -> CategoriesSetResult:
        raise BackendError("set_category is not supported on this platform")

    async def list_rules(self, account: str) -> list[RuleInfo]:
        raise BackendError("list_rules is not supported on this platform")

    async def toggle_rule(self, rule_name: str, enabled: bool, account: str) -> RuleToggledResult:
        raise BackendError("toggle_rule is not supported on this platform")

    async def get_out_of_office(self, account: str) -> OofStatus:
        raise BackendError("get_out_of_office is not supported on this platform")

    # --- Unexpected-error formatting (platform-specific override lives in
    # the concrete backend; default returns a generic client-safe message).
    def format_unexpected_error(self, action: str, e: Exception) -> str:
        """Render an unexpected (non-``BackendError``) exception for the client.

        Concrete backends override this to expose platform-native diagnostic
        detail (e.g. COM HRESULT). The default keeps the message generic
        because it cannot assume what kind of platform exception leaked.
        """
        return f"{action}: An unexpected error occurred."
