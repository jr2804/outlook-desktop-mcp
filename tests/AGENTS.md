# AGENTS.md — tests

## Purpose

Test suite for outlook-desktop-mcp. Two layers: protocol/MCP tests that run
without Outlook, and COM integration tests that require a live Outlook desktop app.

## Ownership

- `conftest.py` — shared fixtures.
- `ref/MSOUTL.OLB` — COM type library reference (local, not committed).
- Naming: `test_*_com.py` = COM integration, `test_*_mcp.py` = MCP protocol, `test_*.py` = unit.

## Local Contracts

- COM tests require a live Outlook; they skip/no-op on Linux — never assert success when Outlook is unavailable.
- Mutating COM tests (send/create/update/delete/move/mark/set — see `_WRITE_TEST_NAMES` in `conftest.py`) are skipped by default so parallel Outlook work is not disturbed. Opt in with `OUTLOOK_MCP_WRITE_TESTS=1` (or `true`). Read-only tests (list/read/search/get) always run.
- Async: `asyncio_mode = "auto"` (set in `pyproject.toml`) — do not add per-test asyncio markers.
- New tools need coverage — see `CONTRIBUTING.md` "Adding New Tools".

## Verification

- `uv run pytest` (repo root) — full suite.
- `outlook-desktop-mcp.cmd test` — runs `tests/test_email_com.py`.

## Child DOX Index

Leaf directory — no nested AGENTS.md files.
