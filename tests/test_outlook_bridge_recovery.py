"""Unit tests for ``OutlookBridge`` reconnect-once + idle TTL release.

These tests avoid ``start()`` so no real COM thread or running Outlook is
required. The recovery helpers are exercised against mock Outlook /
Namespace objects that just succeed, raise, or raise-after-reconnect.
"""

from __future__ import annotations

import sys
import time
from typing import Any

import pytest

# The bridge module imports `pythoncom` at module top; gate that import on
# Windows because pywin32 is Windows-only. Tests still run on Linux CI.
if sys.platform != "win32":
    pytest.skip("Outlook COM bridge is Windows-only", allow_module_level=True)


from outlook_desktop_mcp.backends.win.bridge import (  # noqa: E402  (after skip)
    DEFAULT_IDLE_TIMEOUT,
    OutlookBridge,
)


class _FakeComError(Exception):
    """Stand-in for pywintypes.com_error — any exception counts as a disconnect signal."""


class _FakeOutlook:
    def __init__(self, namespace: Any) -> None:
        self._namespace = namespace


def test_first_call_dispatches_and_returns_value() -> None:
    bridge = _make_bridge()
    calls: list[tuple[Any, Any]] = []

    def func(outlook: Any, namespace: Any, *, marker: int) -> int:
        calls.append((id(outlook), id(namespace)))
        return marker * 2

    result = bridge._run_with_recovery(func, (), {"marker": 7})

    assert result == 14
    assert bridge._outlook is not None
    assert bridge._namespace is not None
    # First call dispatches once and returns; helper invoked once.
    assert len(calls) == 1


def test_disconnect_reconnects_and_retries_succeeding() -> None:
    """A failed call triggers one reconnect+retry; if retry returns, that's the answer."""
    bridge = _make_bridge()
    # Seed the bridge with a "dead" outlook: the cached object exists so the
    # lazy `_outlook is None` branch is skipped, and the first call will fail.
    bridge._outlook = _FakeOutlook(namespace=None)
    bridge._namespace = None
    invocations: list[int] = []
    object.__setattr__(bridge, "_connect", lambda: _apply_connect_stub(bridge, "fake-ns"))

    def func(outlook: Any, namespace: Any, *, n: int) -> str:
        invocations.append(n)
        if len(invocations) == 1:
            # Simulate the dead RCW: first attempt throws.
            raise _FakeComError("RPC_E_DISCONNECTED proxy dead")
        return f"ok-{n}"

    result = bridge._run_with_recovery(func, (), {"n": 5})

    assert result == "ok-5"
    assert len(invocations) == 2  # one failed, one succeeded after reconnect
    assert bridge._namespace == "fake-ns"


def test_disconnect_retry_failure_surfaces_second_error() -> None:
    """If retry also fails, callers see the second error (not the first)."""
    bridge = _make_bridge()
    bridge._outlook = _FakeOutlook(namespace=None)
    invocations: list[int] = []
    object.__setattr__(bridge, "_connect", lambda: _apply_connect_stub(bridge, "fake-ns"))

    def func(outlook: Any, namespace: Any) -> None:
        invocations.append(1)
        raise _FakeComError(f"boom-{len(invocations)}")

    with pytest.raises(_FakeComError, match="boom-2"):
        bridge._run_with_recovery(func, (), {})

    assert len(invocations) == 2  # one try + one retry


def test_idle_ttl_releases_outlook() -> None:
    """After ``idle_timeout`` seconds of inactivity the cached outlook is dropped."""
    bridge = _make_bridge(idle_timeout=0.001)  # 1 ms — expire immediately
    bridge._outlook = _FakeOutlook(namespace="old")
    bridge._namespace = "old"
    bridge._last_used_ts = bridge._last_used_ts - 1.0  # long ago

    bridge._check_idle_release()

    assert bridge._outlook is None
    assert bridge._namespace is None


def test_idle_ttl_does_not_release_when_fresh() -> None:
    """Recently-used outlook survives the idle check."""
    bridge = _make_bridge(idle_timeout=3600.0)
    bridge._outlook = _FakeOutlook(namespace="fresh")
    bridge._namespace = "fresh"

    bridge._check_idle_release()

    assert bridge._outlook is not None
    assert bridge._namespace is not None


def test_successful_call_resets_idle_timestamp() -> None:
    bridge = _make_bridge(idle_timeout=3600.0)
    bridge._outlook = _FakeOutlook(namespace="ns")
    bridge._namespace = "ns"
    bridge._last_used_ts = 0.0  # ancient

    bridge._run_with_recovery(lambda outlook, ns: "ok", (), {})
    # If the call resets the timer, _last_used_ts must be in the future.
    assert bridge._last_used_ts > 0


def test_non_disconnect_exception_still_triggers_one_reconnect() -> None:
    """Spec: any failure triggers one reconnect-and-retry (broad net)."""
    bridge = _make_bridge()
    invocations: list[int] = []
    object.__setattr__(bridge, "_connect", lambda: _apply_connect_stub(bridge, "fake-ns"))

    def func(outlook: Any, namespace: Any) -> str:
        invocations.append(1)
        if len(invocations) == 1:
            raise ValueError("not-a-com-error but we still retry once")
        return "recovered"

    result = bridge._run_with_recovery(func, (), {})

    assert result == "recovered"
    assert len(invocations) == 2


def _apply_connect_stub(bridge: OutlookBridge, namespace: Any = "fake-ns") -> None:
    """Test stand-in for ``_connect`` that never touches real COM.

    Monkey-patched onto ``OutlookBridge`` via ``object.__setattr__`` to bypass
    strict type-checking on attribute assignment; the real Win bridge assigns
    the cached Outlook on the COM thread, tests replace it with a stub.
    """
    bridge._outlook = _FakeOutlook(namespace=namespace)
    bridge._namespace = namespace
    bridge._last_used_ts = time.monotonic()


def _make_bridge(idle_timeout: float = DEFAULT_IDLE_TIMEOUT) -> OutlookBridge:
    """Create a bridge instance whose COM thread never starts."""
    return OutlookBridge(idle_timeout=idle_timeout)


def test_default_idle_timeout_matches_documented_value() -> None:
    """Pinned at 1h to match the user-facing operator docs."""
    assert DEFAULT_IDLE_TIMEOUT == 3600.0
