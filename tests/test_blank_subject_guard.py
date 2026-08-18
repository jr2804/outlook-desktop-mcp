"""
Unit tests for subject validation in create_event / create_meeting / create_task / update_event.

These tests verify that blank/whitespace-only subjects are rejected at the tool boundary,
preventing blank events from being created in Outlook. They do not require a running
Outlook instance — they test the validation guard before bridge.call() is invoked.

Run with:
    cd D:/UserData/hermes/packages/outlook-desktop-mcp
    uv run pytest tests/test_blank_subject_guard.py -v
    # OR if pytest isn't installed:
    uv run python tests/test_blank_subject_guard.py
"""

import asyncio
import json
import os
import sys

# Make src/ importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))


async def test_create_event_blank_subject_rejected() -> None:
    """create_event with subject='' should be rejected with a JSON error."""
    from outlook_desktop_mcp import server  # noqa: PLC0415

    result = await server.create_event(subject="", start="2026-05-15 09:00", end="2026-05-15 18:00")
    data = json.loads(result)
    assert "error" in data, f"Expected error, got: {result!r}"
    assert "blank" in data["error"].lower(), f"Error should mention blank, got: {data['error']!r}"
    print(f"  PASS  create_event(blank subject) -> {data['error']!r}")


async def test_create_event_whitespace_subject_rejected() -> None:
    """create_event with subject='   ' should be rejected."""
    from outlook_desktop_mcp import server  # noqa: PLC0415

    result = await server.create_event(subject="   ", start="2026-05-15 09:00", end="2026-05-15 18:00")
    data = json.loads(result)
    assert "error" in data, f"Expected error, got: {result!r}"
    print(f"  PASS  create_event(whitespace subject) -> {data['error']!r}")


async def test_create_event_valid_subject_proceeds() -> None:
    """create_event with a real subject should NOT be rejected by the guard.

    (This will fail at bridge.call() because there's no real Outlook, but the error
    message must NOT be the blank-subject error — it should be a COM/bridge error.)
    """
    from outlook_desktop_mcp import server  # noqa: PLC0415

    result = await server.create_event(subject="Valid test event", start="2026-05-15 09:00", end="2026-05-15 18:00")
    # Should NOT be the blank-subject error
    if result.startswith("{"):
        data = json.loads(result)
        if "error" in data:
            assert "blank" not in data["error"].lower(), f"Should not be blank-subject error: {data['error']!r}"
    print(f"  PASS  create_event(valid subject) -> bridge proceeds (result={result[:60]!r})")


async def test_create_meeting_blank_subject_rejected() -> None:
    """create_meeting with subject='' should be rejected."""
    from outlook_desktop_mcp import server  # noqa: PLC0415

    result = await server.create_meeting(subject="", start="2026-05-15 09:00", end="2026-05-15 18:00", required_attendees="test@example.com")
    data = json.loads(result)
    assert "error" in data, f"Expected error, got: {result!r}"
    assert "blank" in data["error"].lower()
    print(f"  PASS  create_meeting(blank subject) -> {data['error']!r}")


async def test_create_task_blank_subject_rejected() -> None:
    """create_task with subject='' should be rejected."""
    from outlook_desktop_mcp import server  # noqa: PLC0415

    result = await server.create_task(subject="")
    data = json.loads(result)
    assert "error" in data, f"Expected error, got: {result!r}"
    assert "blank" in data["error"].lower()
    print(f"  PASS  create_task(blank subject) -> {data['error']!r}")


async def test_create_task_whitespace_subject_rejected() -> None:
    r"""create_task with subject='\t\n  ' should be rejected."""
    from outlook_desktop_mcp import server  # noqa: PLC0415

    result = await server.create_task(subject="\t\n  ")
    data = json.loads(result)
    assert "error" in data
    print(f"  PASS  create_task(whitespace subject) -> {data['error']!r}")


async def test_update_event_whitespace_subject_rejected() -> None:
    """update_event with explicit whitespace subject should be rejected.

    Note: subject='' (the default) means "don't change", and is allowed.
    But subject='   ' (explicitly set to blank) should be rejected.
    """
    from outlook_desktop_mcp import server  # noqa: PLC0415

    result = await server.update_event(entry_id="00000000", subject="   ")
    data = json.loads(result)
    assert "error" in data, f"Expected error, got: {result!r}"
    print(f"  PASS  update_event(whitespace subject) -> {data['error']!r}")


async def test_update_event_no_subject_allowed() -> None:
    """update_event with subject='' (default = "skip") should NOT be rejected by blank guard.

    The error you see should be a different one (entry not found, COM error, etc.) —
    NOT the blank-subject error.
    """
    from outlook_desktop_mcp import server  # noqa: PLC0415

    result = await server.update_event(entry_id="00000000", subject="")
    if result.startswith("{"):
        data = json.loads(result)
        if "error" in data:
            assert "blank" not in data["error"].lower(), f"Empty string = 'skip' should be allowed, got: {data['error']!r}"
    print(f"  PASS  update_event(empty=skip) -> guard passes (result={result[:60]!r})")


async def main() -> None:
    tests = [
        test_create_event_blank_subject_rejected,
        test_create_event_whitespace_subject_rejected,
        test_create_event_valid_subject_proceeds,
        test_create_meeting_blank_subject_rejected,
        test_create_task_blank_subject_rejected,
        test_create_task_whitespace_subject_rejected,
        test_update_event_whitespace_subject_rejected,
        test_update_event_no_subject_allowed,
    ]
    passed = 0
    failed = 0
    for test in tests:
        print(f"\n{test.__name__}:")
        try:
            await test()
            passed += 1
        except AssertionError as e:
            print(f"  FAIL: {e}")
            failed += 1
        except Exception as e:
            print(f"  ERROR: {e}")
            failed += 1
    print(f"\n{'=' * 60}")
    print(f"Results: {passed}/{passed + failed} passed")
    print(f"{'=' * 60}")
    return failed == 0


if __name__ == "__main__":
    success = asyncio.run(main())
    sys.exit(0 if success else 1)
