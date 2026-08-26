"""Unit tests for locale-aware Jet/DASL date formatting in the Windows backend.

Regression tests for the month/day swap bug: Outlook's Jet parser reads slash
dates in the user's locale order, so a fixed ``%m/%d/%Y`` silently swapped
month and day on day-first locales (e.g. German) whenever day <= 12.
"""

from __future__ import annotations

import sys
from datetime import datetime, timezone

from outlook_desktop_mcp.backends.win.backend import (
    _dasl_utc,
    _jet_datetime,
    _locale_date_order,
)


def test_jet_datetime_explicit_orders() -> None:
    """The same instant renders per short-date order (Jan 6 must stay Jan 6)."""
    dt = datetime(2026, 1, 6, 9, 5)
    assert _jet_datetime(dt, order=0) == "01/06/2026 09:05"  # MDY
    assert _jet_datetime(dt, order=1) == "06/01/2026 09:05"  # DMY
    assert _jet_datetime(dt, order=2) == "2026/01/06 09:05"  # YMD


def test_locale_date_order_valid() -> None:
    """Detector returns a known order; non-Windows hosts fall back to MDY."""
    if sys.platform == "win32":
        assert _locale_date_order() in (0, 1, 2)
    else:
        assert _locale_date_order() == 0


def test_dasl_utc_naive_becomes_zulu() -> None:
    """Naive datetimes are treated as local and converted to a UTC literal."""
    out = _dasl_utc(datetime(2026, 1, 6, 12, 0))
    assert out.endswith("Z")
    assert "T" in out
    assert "-" in out and ":" in out


def test_dasl_utc_tz_aware_roundtrip() -> None:
    dt = datetime(2026, 1, 6, 12, 0, tzinfo=timezone.utc)
    assert _dasl_utc(dt) == "2026-01-06T12:00:00Z"
