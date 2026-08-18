"""Pydantic response models — single source of truth for MCP tool output shapes.

Every tool returns JSON derived from these models:
- success objects carry a ``status`` field (or are arrays of info models)
- failures are represented by ``Error``
"""

from __future__ import annotations

from pydantic import BaseModel, Field

# --- Errors ---


class Error(BaseModel):
    """Uniform failure payload for every tool."""

    error: str


# --- Accounts / Folders ---


class AccountInfo(BaseModel):
    display_name: str
    store_id: str
    is_default: bool


class FolderInfo(BaseModel):
    name: str
    full_path: str = ""
    item_count: int = 0
    unread_count: int = 0
    subfolders: list[FolderInfo] = Field(default_factory=list)


# --- Email ---


class EmailSummary(BaseModel):
    entry_id: str
    subject: str = "(no subject)"
    sender: str = ""
    sender_name: str = ""
    received_time: str = ""
    unread: bool = False
    has_attachments: bool = False
    attachment_count: int = 0


class EmailFull(EmailSummary):
    to: str = ""
    cc: str = ""
    body: str = ""


class SentResult(BaseModel):
    status: str = "sent"
    subject: str
    to: str = ""


class DraftSavedResult(BaseModel):
    status: str = "draft_saved"
    subject: str
    to: str = ""
    entry_id: str = ""


class ReplySentResult(BaseModel):
    status: str = "sent"
    subject: str
    reply_all: bool = False


class ReplyDraftSavedResult(BaseModel):
    status: str = "draft_saved"
    subject: str
    reply_all: bool = False
    entry_id: str = ""


class MovedResult(BaseModel):
    status: str = "moved"
    subject: str
    target_folder: str


# --- Calendar ---


class EventSummary(BaseModel):
    entry_id: str
    subject: str = "(no subject)"
    start: str = ""
    end: str = ""
    duration: int | None = None
    location: str = ""
    organizer: str = ""
    is_recurring: bool = False
    all_day: bool = False
    busy_status: str = "unknown"
    meeting_status: str = "unknown"
    required_attendees: str = ""
    optional_attendees: str = ""


class EventFull(EventSummary):
    body: str = ""
    reminder_set: bool = False
    reminder_minutes: int | None = None
    categories: str = ""
    response_status: str = "unknown"
    attendees: str = ""


class EventCreatedResult(BaseModel):
    status: str = "created"
    subject: str
    start: str
    end: str
    entry_id: str = ""


class EventUpdatedResult(BaseModel):
    status: str = "updated"
    subject: str
    start: str = ""
    end: str = ""
    location: str = ""
    entry_id: str = ""


class MeetingSentResult(BaseModel):
    status: str = "sent"
    subject: str
    required_attendees: str
    optional_attendees: str | None = None


class MeetingResponseResult(BaseModel):
    status: str = "responded"
    response: str
    subject: str


# --- Tasks ---


class TaskSummary(BaseModel):
    entry_id: str
    subject: str = "(no subject)"
    status: str = ""
    percent_complete: int | None = None
    due_date: str | None = None
    start_date: str | None = None
    importance: str = ""
    complete: bool = False
    categories: str = ""
    owner: str = ""
    priority: str = ""


class TaskFull(TaskSummary):
    body: str = ""
    reminder_set: bool = False
    date_completed: str | None = None


class TaskCreatedResult(BaseModel):
    status: str = "created"
    subject: str
    entry_id: str = ""
    due_date: str | None = None


# --- Attachments ---


class AttachmentInfo(BaseModel):
    index: int
    filename: str
    size: int = 0


class AttachmentSavedResult(BaseModel):
    status: str = "saved"
    filename: str
    path: str
    size: int | None = None


# --- Categories / Rules ---


class CategoryInfo(BaseModel):
    name: str
    color: int = 0


class CategoriesSetResult(BaseModel):
    status: str = "categories_set"
    subject: str
    categories: str = ""


class RuleInfo(BaseModel):
    index: int
    name: str
    enabled: bool


class RuleToggledResult(BaseModel):
    status: str  # "enabled" | "disabled"
    rule: str


# --- Generic item status (mark read/unread, complete, delete, cancel) ---


class ItemStatusResult(BaseModel):
    status: str
    subject: str
    entry_id: str = ""
    note: str = ""


# --- Out of Office ---


class OofStatus(BaseModel):
    out_of_office: bool | None = None
    status: str  # "on" | "off" | "unknown"
    note: str = ""


class OofSetResult(BaseModel):
    out_of_office: bool
    status: str  # "on" | "off"
    note: str = ""
