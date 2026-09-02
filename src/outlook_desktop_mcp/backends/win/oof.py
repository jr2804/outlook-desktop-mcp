"""Out-of-office operations on Windows via COM."""

from __future__ import annotations

from typing import Any

from outlook_desktop_mcp.backends.win._types import Namespace
from outlook_desktop_mcp.backends.win.helpers import _require_store
from outlook_desktop_mcp.models import OofSetResult, OofStatus


async def get_out_of_office(bridge: Any, account: str) -> OofStatus:
    def _get(outlook: Any, namespace: Namespace) -> OofStatus:
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

    return await bridge.call(_get)


async def set_out_of_office(bridge: Any, enabled: bool, message: str, account: str) -> OofSetResult:
    def _set(outlook: Any, namespace: Namespace) -> OofSetResult:
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

    return await bridge.call(_set)
