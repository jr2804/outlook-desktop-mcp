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
import sys

import pytest

import outlook_desktop_mcp.server as server
from outlook_desktop_mcp.backends.base import Backend

# Test modules that drive a live Outlook COM / AppleScript session. They are
# manual validation scripts (by the upstream author) that require Windows/macOS
# and a running Outlook. They are skipped on platforms without COM.
_COM_ONLY_MODULES = {
    "tests.phase1_com_test",
    "tests.calendar_com_test",
    "tests.extras_com_test",
}

_UNAVAILABLE_MSG = (
    "Outlook backend is unavailable on this platform "
    "(expected on Linux/WSL). Integration tests require Windows/macOS."
)


class _UnavailableBackend(Backend):
    """Backend stub that fails fast (no Outlook available on this platform)."""

    name = "unavailable-stub"

    async def compose_email(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_emails(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def read_email(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def mark_as_read(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def mark_as_unread(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def move_email(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def reply_email(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_folders(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def search_emails(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_events(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def get_event(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def create_event(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def create_meeting(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def update_event(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def delete_event(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def search_events(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_tasks(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def get_task(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def create_task(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def complete_task(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def delete_task(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def list_attachments(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def save_attachment(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)

    async def set_out_of_office(self, *args, **kwargs):
        raise RuntimeError(_UNAVAILABLE_MSG)


def _install_backend():
    """Install the real backend on Windows/macOS, a stub elsewhere."""
    from outlook_desktop_mcp.platform import Platform, current_platform

    platform = current_platform()
    if platform == Platform.DARWIN:
        from outlook_desktop_mcp.backends.mac import AppleScriptBackend
        server.set_backend(AppleScriptBackend(), platform=platform)
    elif platform == Platform.WINDOWS:
        from outlook_desktop_mcp.backends.win import ComBackend
        server.set_backend(ComBackend(), platform=platform)
    else:
        server.set_backend(_UnavailableBackend(), platform=Platform.WINDOWS)


_install_backend()


@pytest.fixture(autouse=True)
def _ensure_backend_installed():
    """Guarantee a backend is installed for every test (idempotent)."""
    if server.backend is None:
        _install_backend()
    yield


def pytest_collection_modifyitems(config, items):
    """Skip Windows/macOS-only COM test modules when COM is unavailable."""
    if sys.platform == "win32":
        return
    skip_com = pytest.mark.skip(
        reason="Requires a live Outlook COM session (Windows/macOS only)."
    )
    for item in items:
        if item.module.__name__ in _COM_ONLY_MODULES:
            item.add_marker(skip_com)
