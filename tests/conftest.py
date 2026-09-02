"""Shared pytest fixtures for outlook-desktop-mcp.

Cross-platform note
--------------------
The real backends only work with a running Outlook (COM on Windows,
AppleScript on macOS). On platforms where no real backend can run (e.g.
Linux/WSL CI), we install a stub backend whose methods fail *immediately*
with a clear error instead of blocking for a bridge timeout. This keeps the
unit tests (e.g. the blank-subject guard tests) fast and deterministic
everywhere while still exercising all logic *before* the backend boundary.
"""

import os
import sys
from collections.abc import Iterator
from typing import Any, Never

import pytest
import pythoncom
import win32com.client
from _pytest.config import Config

from outlook_desktop_mcp import server
from outlook_desktop_mcp.backends.base import Backend
from outlook_desktop_mcp.backends.mac import AppleScriptBackend
from outlook_desktop_mcp.backends.win import ComBackend
from outlook_desktop_mcp.platform import Platform, current_platform

# Test modules that drive a live Outlook COM / AppleScript session. They are
# manual validation scripts (by the upstream author) that require Windows/macOS
# and a running Outlook. They are skipped on platforms without COM.
_COM_ONLY_MODULES = {
    "tests.test_email_com",
    "tests.test_calendar_com",
    "tests.test_extras_com",
}

_UNAVAILABLE_MSG = "Outlook backend is unavailable on this platform (expected on Linux/WSL). Integration tests require Windows/macOS."


class _UnavailableBackend(Backend):
    """Backend stub that fails fast (no Outlook available on this platform)."""

    name = "unavailable-stub"

    async def compose_email(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_emails(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def read_email(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def mark_as_read(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def mark_as_unread(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def move_email(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def reply_email(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_folders(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def search_emails(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_events(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def get_event(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def create_event(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def create_meeting(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def update_event(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def delete_event(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def search_events(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_tasks(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def get_task(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def create_task(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def complete_task(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def delete_task(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_attachments(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def save_attachment(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def set_out_of_office(self, *args, **kwargs) -> Never:
        raise RuntimeError(_UNAVAILABLE_MSG)


def _install_backend() -> None:
    """Install the real backend on Windows/macOS, a stub elsewhere."""
    try:
        platform = current_platform()
    except RuntimeError:
        # Unsupported host (e.g. Linux CI): fall back to the unavailable stub.
        server.set_backend(_UnavailableBackend(), platform=Platform.WINDOWS)
        return
    if platform == Platform.DARWIN:
        server.set_backend(AppleScriptBackend(), platform=platform)
    elif platform == Platform.WINDOWS:
        server.set_backend(ComBackend(), platform=platform)
    else:
        server.set_backend(_UnavailableBackend(), platform=Platform.WINDOWS)


_install_backend()


# Tests that MUTATE the live Outlook mailbox (send/create/update/delete/move/
# mark/set). They are skipped by default so parallel work in Outlook is not
# disturbed; set OUTLOOK_MCP_WRITE_TESTS=1 (or true) to include them.
# Read-only tests (list/read/search/get) always run.
_WRITE_TEST_NAMES = {
    "test_send_email",
    "test_mark_read_unread",
    "test_move_to_archive",
    "test_create_appointment",
    "test_create_meeting",
    "test_update_event",
    "test_delete_event",
    "test_create_task",
    "test_complete_task",
    "test_delete_task",
    "test_set_category",
}


@pytest.fixture(autouse=True)
def _ensure_backend_installed() -> None:
    """Guarantee a backend is installed for every test (idempotent)."""
    if server.backend is None:
        _install_backend()


@pytest.fixture(scope="session")
def namespace(outlook: Any) -> Any:
    """Session-scoped MAPI namespace bound to the live Outlook session."""
    return outlook.GetNamespace("MAPI")


# --- COM fixtures for the *_com_test validation scripts ---


@pytest.fixture(scope="session")
def outlook() -> Iterator[Any]:
    """Session-scoped live Outlook Application COM object (Windows only)."""
    if sys.platform != "win32":
        pytest.skip("Requires a live Outlook COM session (Windows only).")

    pythoncom.CoInitialize()
    try:
        yield win32com.client.Dispatch("Outlook.Application")
    finally:
        pythoncom.CoUninitialize()


def pytest_collection_modifyitems(config: Config, items: list[Any]) -> None:  # noqa: ANN001
    """Skip COM modules off-Windows and mutating tests unless opted in."""
    del config
    skip_com = pytest.mark.skip(reason="Requires a live Outlook COM session (Windows/macOS only).")
    skip_write = pytest.mark.skip(reason="Mutates the live Outlook mailbox; set OUTLOOK_MCP_WRITE_TESTS=1 to run.")
    write_enabled = _write_tests_enabled()
    for item in items:
        if sys.platform != "win32" and item.module.__name__ in _COM_ONLY_MODULES:
            item.add_marker(skip_com)
        elif not write_enabled and item.name.split("[")[0] in _WRITE_TEST_NAMES:
            item.add_marker(skip_write)


def _write_tests_enabled() -> bool:
    """Mutating live-Outlook tests opt-in via OUTLOOK_MCP_WRITE_TESTS."""
    return os.environ.get("OUTLOOK_MCP_WRITE_TESTS", "").strip().lower() in ("1", "true", "yes")
