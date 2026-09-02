"""Email operations on Windows via COM."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from outlook_desktop_mcp.backends.base import BackendError
from outlook_desktop_mcp.backends.win._types import Folder, MailItem, Namespace
from outlook_desktop_mcp.backends.win.formatting import (
    format_email_full,
    format_email_summary,
)
from outlook_desktop_mcp.backends.win.helpers import (
    _OL_CLASS_MAIL,
    _item_received_utc,
    _parse_date_window,
    _require_class,
    _require_store,
    _resolve_folder,
    _safe_dasl,
    _within_window,
)
from outlook_desktop_mcp.models import (
    AccountInfo,
    DraftSavedResult,
    EmailFull,
    EmailSummary,
    FolderInfo,
    ItemStatusResult,
    MovedResult,
    ReplyDraftSavedResult,
    ReplySentResult,
    SentResult,
)
from outlook_desktop_mcp.tools._folder_constants import OL_MAIL_ITEM


async def list_accounts(bridge: Any) -> list[AccountInfo]:
    def _list(outlook: Any, namespace: Namespace) -> list[AccountInfo]:
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

    return await bridge.call(_list)


async def compose_email(  # noqa: PLC0415
    bridge: Any,
    to: str,
    subject: str,
    body: str,
    cc: str,
    bcc: str,
    html_body: str,
    account: str,
    send: bool,
) -> SentResult | DraftSavedResult:
    def _compose(outlook: Any, namespace: Namespace) -> SentResult | DraftSavedResult:
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

    return await bridge.call(_compose)


async def list_emails(  # noqa: PLC0415
    bridge: Any,
    folder: str,
    count: int,
    unread_only: bool,
    start_date: str,
    end_date: str,
    account: str,
) -> list[EmailSummary]:
    def _list(outlook: Any, namespace: Namespace) -> list[EmailSummary]:
        effective_count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            raise BackendError(f"Folder '{folder}' not found")
        items = target.Items
        items.Sort("[ReceivedTime]", True)
        # Outlook's date Restrict is unreliable here (silently returns the whole
        # folder / drops sibling filters), so restrict only by the filters that
        # are proven to work (unread) and post-filter by date in Python.
        restrictions = []
        if unread_only:
            restrictions.append("[UnRead] = True")
        if restrictions:
            items = items.Restrict(" AND ".join(restrictions))
        results = []
        # Parse the date window (local-naive from input, treated as local time).
        window = _parse_date_window(start_date, end_date)
        total = items.Count
        # Walk newest→oldest until we've collected the target count, the folder
        # is exhausted, or we sink below the window's lower bound.
        for i in range(total):
            if len(results) >= effective_count:
                break
            try:
                item = items.Item(i + 1)
            except Exception:  # noqa: S112 - item vanished mid-iteration
                continue
            dt = _item_received_utc(item)
            if dt is None:
                continue
            if not _within_window(dt, window):
                if window.lo is not None and dt < window.lo:
                    break  # newer items already collected; rest are older
                continue  # too new (above hi) — keep walking
            try:
                results.append(EmailSummary.model_validate(format_email_summary(item)))
            except Exception:  # noqa: S112
                continue
        return results

    return await bridge.call(_list)


async def read_email(
    bridge: Any,
    entry_id: str,
    subject_search: str,
    folder: str,
    account: str,
) -> EmailFull:
    def _read(outlook: Any, namespace: Namespace) -> EmailFull:
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

    return await bridge.call(_read)


async def mark_as_read(bridge: Any, get_item: Callable, entry_id: str, account: str) -> ItemStatusResult:
    def _mark(outlook: Any, namespace: Namespace) -> ItemStatusResult:
        item = get_item(namespace, entry_id, account)
        _require_class(item, _OL_CLASS_MAIL, "mail item")
        subject = item.Subject
        item.UnRead = False
        item.Save()
        return ItemStatusResult(status="marked_read", subject=subject, entry_id=entry_id)

    return await bridge.call(_mark)


async def mark_as_unread(bridge: Any, get_item: Callable, entry_id: str, account: str) -> ItemStatusResult:
    def _mark(outlook: Any, namespace: Namespace) -> ItemStatusResult:
        item = get_item(namespace, entry_id, account)
        _require_class(item, _OL_CLASS_MAIL, "mail item")
        subject = item.Subject
        item.UnRead = True
        item.Save()
        return ItemStatusResult(status="marked_unread", subject=subject, entry_id=entry_id)

    return await bridge.call(_mark)


async def move_email(bridge: Any, entry_id: str, target_folder: str, account: str) -> MovedResult:
    def _move(outlook: Any, namespace: Namespace) -> MovedResult:
        item = namespace.GetItemFromID(entry_id)
        _require_class(item, _OL_CLASS_MAIL, "mail item")
        subject = item.Subject
        store = _require_store(namespace, account)
        dest = _resolve_folder(namespace, target_folder, store)
        if not dest:
            raise BackendError(f"Target folder '{target_folder}' not found. Use list_folders to see available folders.")
        item.Move(dest)
        return MovedResult(subject=subject, target_folder=target_folder)

    return await bridge.call(_move)


async def reply_email(  # noqa: PLC0415
    bridge: Any,
    get_item: Callable,
    entry_id: str,
    body: str,
    reply_all: bool,
    html_body: str,
    account: str,
    send: bool,
) -> ReplySentResult | ReplyDraftSavedResult:
    def _reply(outlook: Any, namespace: Namespace) -> ReplySentResult | ReplyDraftSavedResult:
        item = get_item(namespace, entry_id, account)
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

    return await bridge.call(_reply)


async def list_folders(bridge: Any, folder: str, max_depth: int, account: str) -> list[FolderInfo]:
    def _list(outlook: Any, namespace: Namespace) -> list[FolderInfo]:
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

    return await bridge.call(_list)


async def search_emails(  # noqa: PLC0415
    bridge: Any,
    query: str,
    folder: str,
    count: int,
    start_date: str,
    end_date: str,
    account: str,
) -> list[EmailSummary]:
    def _search(outlook: Any, namespace: Namespace) -> list[EmailSummary]:
        effective_count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            raise BackendError(f"Folder '{folder}' not found")
        safe_query = _safe_dasl(query)
        filter_str = f"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)
        results = []
        # Date filtering is done in Python: adding a datereceived comparison to
        # the Restrict silently drops the whole filter (including the subject
        # match) on this Outlook build, so restrict by subject only and apply
        # the window here, walking newest→oldest until collected or past lo.
        window = _parse_date_window(start_date, end_date)
        total = items.Count
        for i in range(total):
            if len(results) >= effective_count:
                break
            try:
                item = items.Item(i + 1)
            except Exception:  # noqa: S112 - item vanished mid-iteration
                continue
            dt = _item_received_utc(item)
            if dt is None:
                continue
            if not _within_window(dt, window):
                if window.lo is not None and dt < window.lo:
                    break  # newer items already collected; rest are older
                continue  # too new (above hi) — keep walking
            try:
                results.append(EmailSummary.model_validate(format_email_summary(item)))
            except Exception:  # noqa: S112
                continue
        return results

    return await bridge.call(_search)
