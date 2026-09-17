"""Regression tests for Object Model Guard handling in calendar formatting.

Outlook's Object Model Guard blocks *address information* reads
(``Organizer``/``RequiredAttendees``/``OptionalAttendees``) when no admin policy
auto-approves them. A denied read raises ``com_error(-2147467259, 'Unspecified
error')``; an unanswered prompt blocks the COM thread indefinitely.

Both failure modes used to abort/empty ``list_events``, which made the HR absence
sync believe no event existed and re-create every vacation day as a duplicate.
"""

from __future__ import annotations

import os

import pytest

from outlook_desktop_mcp.backends.win import formatting


class _GuardBlocked:
    """COM property that raises the way Object Model Guard denial does."""

    def __init__(self) -> None:
        self.message = "Unspecified error"

    def __str__(self) -> str:  # pragma: no cover - repr convenience
        return "com_error(-2147467259, 'Unspecified error')"

    def raise_it(self):
        raise RuntimeError(str(self))


class FakeAppointment:
    """Minimal AppointmentItem stand-in (no COM)."""

    EntryID = "EID-1"
    Subject = "Urlaub JR"
    Start = "2026-07-31 00:00:00+00:00"
    End = "2026-08-01 00:00:00+00:00"
    Duration = 1440
    Location = ""
    IsRecurring = False
    AllDayEvent = True
    BusyStatus = 3
    MeetingStatus = 0

    @property
    def Organizer(self):
        raise RuntimeError("com_error(-2147467259, 'Unspecified error')")

    @property
    def RequiredAttendees(self):
        raise RuntimeError("com_error(-2147467259, 'Unspecified error')")

    @property
    def OptionalAttendees(self):
        raise RuntimeError("com_error(-2147467259, 'Unspecified error')")


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    monkeypatch.delenv("OUTLOOK_MCP_ADDRESS_FIELDS", raising=False)


def test_guarded_fields_default_off_avoids_blocked_reads():
    """Default: protected fields are never read, so a prompt cannot block us."""
    summary = formatting.format_event_summary(FakeAppointment())
    assert summary["organizer"] == ""
    assert summary["required_attendees"] == ""
    assert summary["optional_attendees"] == ""


def test_core_fields_survive_blocked_address_fields():
    """The identity of the event must survive a guard-blocked field."""
    summary = formatting.format_event_summary(FakeAppointment())
    assert summary["subject"] == "Urlaub JR"
    assert summary["start"].startswith("2026-07-31")
    assert summary["entry_id"] == "EID-1"


def test_guarded_falls_back_when_opt_in_read_is_denied(monkeypatch):
    """With the opt-in flag set, a denied read degrades to the default."""
    monkeypatch.setenv("OUTLOOK_MCP_ADDRESS_FIELDS", "1")
    summary = formatting.format_event_summary(FakeAppointment())
    assert summary["organizer"] == ""
    assert summary["required_attendees"] == ""


def test_guarded_returns_value_when_read_is_allowed(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_ADDRESS_FIELDS", "1")

    class Allowed(FakeAppointment):
        Organizer = "Jan Reimes"
        RequiredAttendees = "Max Havenith"
        OptionalAttendees = ""

    summary = formatting.format_event_summary(Allowed())
    assert summary["organizer"] == "Jan Reimes"
    assert summary["required_attendees"] == "Max Havenith"


def test_address_info_flag_parsing(monkeypatch):
    for value in ("1", "true", "TRUE", "yes", " Yes "):
        monkeypatch.setenv("OUTLOOK_MCP_ADDRESS_FIELDS", value)
        assert formatting._address_info_enabled() is True
    for value in ("", "0", "false", "no"):
        monkeypatch.setenv("OUTLOOK_MCP_ADDRESS_FIELDS", value)
        assert formatting._address_info_enabled() is False


def test_guarded_does_not_swallow_non_callable_errors():
    """A genuinely broken getter still yields the default, never a crash."""
    assert formatting._guarded(lambda: 1 / 0, "fallback") == "fallback"
