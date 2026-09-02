"""Shared helper functions for the Windows COM backend."""

from __future__ import annotations

import dataclasses
from datetime import UTC, datetime
from typing import Any

from outlook_desktop_mcp.backends.base import BackendError
from outlook_desktop_mcp.backends.win._types import Folder, Namespace, Store
from outlook_desktop_mcp.tools._folder_constants import FOLDER_NAME_TO_ENUM

# Outlook item class enums (OlObjectClass).
_OL_CLASS_MAIL = 43
_OL_CLASS_APPOINTMENT = 26
_OL_CLASS_TASK = 48


@dataclasses.dataclass
class _DateWindow:
    """A [lo, hi] bounds pair for Python-side date filtering.

    `lo` and `hi` are timezone-aware datetimes (UTC) or None when the bound is
    open. `hi` is inclusive; `lo` is inclusive.
    """

    lo: datetime | None
    hi: datetime | None


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


def _parse_date_window(start_date: str, end_date: str) -> _DateWindow:
    """Build a timezone-aware window from ISO date(-time) strings.

    Input dates are assumed to be local time (naive) and are converted to UTC.
    `start_date` maps to the window's lower bound (inclusive), `end_date` to the
    upper bound (inclusive).
    """
    local = datetime.now().astimezone()

    def _to_utc(value: str) -> datetime:
        dt = datetime.fromisoformat(value)
        # Naive input -> interpret as local time, then convert to UTC.
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=local.tzinfo)
        return dt.astimezone(UTC)

    lo = _to_utc(start_date) if start_date else None
    hi = _to_utc(end_date) if end_date else None
    if not end_date and start_date:
        # A start date without an end date means "through now".
        hi = datetime.now(UTC)
    return _DateWindow(lo=lo, hi=hi)


def _item_received_utc(item: Any) -> datetime | None:
    """Return an email's ReceivedTime as an aware UTC datetime, or None."""
    try:
        return _com_datetime_utc(item.ReceivedTime)
    except Exception:  # noqa: S112 - missing attribute / COM transient
        return None


def _item_start_utc(item: Any) -> datetime | None:
    """Return an appointment's Start as an aware UTC datetime, or None."""
    try:
        return _com_datetime_utc(item.Start)
    except Exception:  # noqa: S112 - missing attribute / COM transient
        return None


def _com_datetime_utc(raw: Any) -> datetime | None:
    """Normalize a COM date value (native datetime or ISO string) to aware UTC."""
    if raw is None:
        return None
    if isinstance(raw, str):
        try:
            raw = datetime.fromisoformat(raw)
        except ValueError:
            return None
    if raw.tzinfo is None:
        raw = raw.replace(tzinfo=datetime.now().astimezone().tzinfo)
    return raw.astimezone(UTC)


def _within_window(dt: datetime, window: _DateWindow) -> bool:
    """Return True if the aware UTC *dt* is inside [lo, hi] (both inclusive)."""
    if window.lo is not None and dt < window.lo:
        return False
    return not (window.hi is not None and dt > window.hi)


def _require_class(item: Any, expected_class: int, label: str) -> None:
    """Raise BackendError if the item is not of the expected Outlook class."""
    if item.Class != expected_class:
        raise BackendError(f"Entry ID does not refer to a {label}.")


def _require_store(namespace: Namespace, account: str = "") -> Store:
    """Resolve store, raising BackendError if not found."""
    store = _resolve_store(namespace, account)
    if store is None:
        raise BackendError(f"Account '{account}' not found. Use list_accounts to see available accounts.")
    return store


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
