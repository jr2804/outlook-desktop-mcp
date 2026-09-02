"""Unit tests for locale-aware Jet/DASL date formatting in the Windows backend.

Regression tests for the month/day swap bug: Outlook's Jet parser reads slash
dates in the user's locale order, so a fixed ``%m/%d/%Y`` silently swapped
month and day on day-first locales (e.g. German) whenever day <= 12.
"""

from __future__ import annotations

import sys
from datetime import UTC, datetime, timedelta

from outlook_desktop_mcp.backends.win.helpers import (
    _dasl_utc,
    _DateWindow,
    _jet_datetime,
    _locale_date_order,
    _parse_date_window,
    _within_window,
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
    assert "-" in out
    assert ":" in out


def test_dasl_utc_tz_aware_roundtrip() -> None:
    dt = datetime(2026, 1, 6, 12, 0, tzinfo=UTC)
    assert _dasl_utc(dt) == "2026-01-06T12:00:00Z"


def test_parse_date_window_naive_local_to_utc() -> None:
    """Naive date strings are interpreted as local time and converted to UTC."""
    w = _parse_date_window("2026-01-05", "2026-01-09")
    assert w.lo is not None
    assert w.lo.tzinfo is not None
    assert w.hi is not None
    assert w.hi.tzinfo is not None
    assert w.lo < w.hi


def test_parse_date_window_open_high_when_only_start() -> None:
    """A start date with no end date closes the window at 'now'."""
    w = _parse_date_window("2026-01-05", "")
    assert w.lo is not None
    assert w.hi is not None  # defaults to now


def test_parse_date_window_both_empty_is_open() -> None:

    w = _parse_date_window("", "")
    assert w.lo is None
    assert w.hi is None


def test_within_window_boundaries_inclusive() -> None:

    w = _parse_date_window("2026-01-05", "2026-01-09")
    assert w.lo is not None
    assert w.hi is not None
    assert _within_window(w.lo, w)  # lower bound inclusive
    assert _within_window(w.hi, w)  # upper bound inclusive
    assert _within_window(w.lo + timedelta(minutes=1), w)
    assert not _within_window(w.lo - timedelta(minutes=1), w)
    assert not _within_window(w.hi + timedelta(minutes=1), w)
    assert _within_window(datetime(2026, 1, 7, 12, 0, tzinfo=UTC), w)


def test_within_window_open_bounds() -> None:

    w = _DateWindow(lo=None, hi=None)
    assert _within_window(datetime(2026, 1, 7, tzinfo=UTC), w)


def test_date_window_comparisons_are_utc_aware() -> None:
    """The window bounds and candidate items are compared in UTC regardless of
    the caller's local timezone, so naive input and aware items line up.
    """
    # 2026-01-05 00:00 local == 2026-01-04T23:00:00Z for UTC+1
    w = _parse_date_window("2026-01-05", "2026-01-05 23:59")
    # A candidate item stamped 2026-01-05 00:00 UTC is just above lo (23:00 UTC).
    assert _within_window(datetime(2026, 1, 5, 0, 0, tzinfo=UTC), w)
    # A candidate stamped 2026-01-04 21:59 UTC is strictly below lo.
    assert not _within_window(datetime(2026, 1, 4, 21, 59, tzinfo=UTC), w)
