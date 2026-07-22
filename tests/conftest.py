"""Shared pytest fixtures for outlook-desktop-mcp tests.

Cross-platform note
--------------------
The Outlook COM bridge only works on Windows (and AppleScript on macOS).
On platforms where the real bridge cannot run (e.g. Linux/WSL CI), we stub
``bridge.call`` so that any code path reaching it fails *immediately* with a
clear error instead of blocking for the bridge's 60s COM timeout. This keeps
the unit tests (e.g. the blank-subject guard tests) fast and deterministic
everywhere while still exercising all logic *before* the COM boundary.
"""
import sys

import pytest

import outlook_desktop_mcp.server as server

# Test modules that drive a live Outlook COM / AppleScript session. They are
# manual validation scripts (by the upstream author) that require Windows/macOS
# and a running Outlook. They are skipped on platforms without COM.
_COM_ONLY_MODULES = {
    "tests.phase1_com_test",
    "tests.calendar_com_test",
    "tests.extras_com_test",
}


def _make_linux_bridge_stub():
    """Return a ``bridge.call`` replacement that fails fast (no COM available)."""

    async def _stub_call(func, *args, **kwargs):
        raise RuntimeError(
            "Outlook COM bridge is unavailable on this platform "
            "(expected on Linux/WSL). COM integration tests require Windows/macOS."
        )

    return _stub_call


@pytest.fixture(autouse=True)
def _stub_com_bridge_on_non_windows():
    """On non-Windows, replace bridge.call with a fast-failing stub.

    This is autouse so every test benefits without opting in. On Windows the
    real bridge is left intact and integration tests run against live Outlook.
    """
    if sys.platform == "win32":
        yield
        return

    original = server.bridge.call
    server.bridge.call = _make_linux_bridge_stub()
    try:
        yield
    finally:
        server.bridge.call = original


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
