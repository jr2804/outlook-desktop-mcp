"""Unit tests for the Windows backend date-window filtering helpers.
Covers window parsing, inclusive UTC-boundary comparison, and COM date
normalization used by email and calendar searches.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta

from outlook_desktop_mcp.backends.win.helpers import (
    _DateWindow,
    _item_start_utc,
    _parse_date_window,
    _within_window,
)


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


def test_item_start_utc_normalizes_to_aware_utc() -> None:
    """Start is normalized to an aware UTC datetime (naive treated as local)."""

    class _Fake:
        def __init__(self, start: object) -> None:
            self.Start = start

    # Naive value -> aware UTC.
    dt = _item_start_utc(_Fake(datetime(2026, 1, 5, 12, 0)))
    assert dt is not None
    assert dt.tzinfo is not None
    assert dt.tzinfo.utcoffset(dt) == UTC.utcoffset(datetime(2026, 1, 5, tzinfo=UTC))

    # ISO string input.
    dt2 = _item_start_utc(_Fake("2026-01-05T12:00:00"))
    assert dt2 is not None
    assert dt2.tzinfo is not None

    # None and invalid values return None.
    assert _item_start_utc(_Fake(None)) is None
    assert _item_start_utc(_Fake("not-a-date")) is None
