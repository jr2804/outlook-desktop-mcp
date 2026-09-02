"""Attachment operations on Windows via COM."""

from __future__ import annotations

import os
import re
from collections.abc import Callable
from typing import Any

from outlook_desktop_mcp.backends.base import BackendError
from outlook_desktop_mcp.backends.win._types import Namespace
from outlook_desktop_mcp.models import AttachmentInfo, AttachmentSavedResult


async def list_attachments(bridge: Any, get_item: Callable, entry_id: str, account: str) -> list[AttachmentInfo]:
    def _list(outlook: Any, namespace: Namespace) -> list[AttachmentInfo]:
        item = get_item(namespace, entry_id, account)
        return [
            AttachmentInfo(index=i + 1, filename=att.FileName, size=att.Size) for i in range(item.Attachments.Count) for att in [item.Attachments.Item(i + 1)]
        ]

    return await bridge.call(_list)


async def save_attachment(
    bridge: Any,
    get_item: Callable,
    entry_id: str,
    attachment_index: int,
    save_directory: str,
    account: str,
) -> AttachmentSavedResult:
    def _save(outlook: Any, namespace: Namespace) -> AttachmentSavedResult:
        item = get_item(namespace, entry_id, account)
        if attachment_index < 1 or item.Attachments.Count < attachment_index:
            raise BackendError(f"Only {item.Attachments.Count} attachment(s), requested index {attachment_index}")

        att = item.Attachments.Item(attachment_index)
        directory = save_directory or os.path.join(os.path.expanduser("~"), "Downloads")

        directory = os.path.realpath(directory)
        os.makedirs(directory, exist_ok=True)

        # Strip path separators and dangerous characters from filename
        safe_name = os.path.basename(att.FileName)
        safe_name = re.sub(r"[^\w\.\-_ ]", "_", safe_name)
        if not safe_name:
            safe_name = "attachment"

        save_path = os.path.join(directory, safe_name)

        # Ensure final path is still inside the intended directory
        real_path = os.path.realpath(save_path)
        if not real_path.startswith(directory + os.sep) and real_path != directory:
            raise BackendError("Attachment filename would escape the target directory.")

        att.SaveAsFile(save_path)
        return AttachmentSavedResult(filename=safe_name, path=save_path, size=att.Size)

    return await bridge.call(_save)
