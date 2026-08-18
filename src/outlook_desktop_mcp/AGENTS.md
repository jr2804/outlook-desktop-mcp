# AGENTS.md — src/outlook_desktop_mcp

## Purpose

Core package of outlook-desktop-mcp: exposes a running Outlook Desktop as an
MCP server. One codebase, two backends — COM (Windows) and AppleScript (macOS).

## Ownership

- `entrypoint.py` — platform routing. The only module allowed to switch on
  `sys.platform` (via `current_platform()`); picks the backend and injects it
  via `server.set_backend(backend, platform=...)`.
- `server.py` — unified MCP tool surface (33 tools on Windows, 26 on macOS):
  signatures, docstrings, validation, JSON contract. Owns the FastMCP
  instance and the `_run` error-conversion boundary.
- `models.py` — pydantic response models; single source of truth for tool
  output shapes (`status` on success, `error` on failure).
- `platform.py` — `Platform` StrEnum + `current_platform()` helper.
- `instructions.py` — `build_instructions(platform)` composing the
  LLM-facing system prompt per platform.
- `backends/base/` — `Backend` ABC + `BackendError` + capability flags +
  `BridgeBase` ABC.
- `backends/win/` — Windows-only: COM bridge, `ComBackend`, COM type stubs,
  COM error formatting, Outlook item formatting helpers.
- `backends/mac/` — macOS-only: AppleScript bridge, `AppleScriptBackend`,
  AppleScript helpers.
- `tools/` — constants only (`_folder_constants.py`); shared by both backends.

## Local Contracts

- Platform routing: `entrypoint.py` selects the backend; nothing else branches
  on OS. Windows-only tools are registered in `server.set_backend()` based on
  backend capability flags.
- Tool surface: each tool is an `async def` with `@mcp.tool()`; the docstring
  is the LLM-facing contract — keep it detailed and current.
- Return contract: backends return pydantic models from `models.py` (or raise
  `BackendError`); `server._run()` serializes to JSON — success objects carry a
  `status` field, failures an `error` field. No plain-string returns anywhere.
- Unified signatures: tools take the superset of parameters; `account` (and
  some `folder`/date params) are documented as Windows-only/ignored-on-macOS.
- Bridge functions receive `(outlook, namespace)` as first args (COM) or the
  bridge object (macOS).
- Constants: import from `tools/_folder_constants.py`; never hard-code Outlook
  enum values elsewhere.
- Errors: `BackendError` for handled failures; unexpected exceptions are
  rendered client-safe via `Backend.format_unexpected_error()` (overridden by
  `ComBackend` to surface COM HRESULT detail). `server._run()` is the single
  conversion boundary.
- Platform independence: **no platform-specific code outside `backends/`**.
  `utils/` was removed; its contents moved into the relevant backend package.

## Work Guidance

Adding a tool: 1) backend function, 2) `@mcp.tool()` async handler with detailed
docstring, 3) test coverage, 4) update `instructions.py` if a capability
category changes. See `CONTRIBUTING.md`.

## Verification

- `uv run pytest` — protocol-level tests.
- `outlook-desktop-mcp.cmd test` — COM validation (Windows, live Outlook).

## Child DOX Index

| Path | Scope |
|------|-------|
| `backends/base/` | `Backend` ABC, `BackendError`, `BridgeBase` ABC, capability flags |
| `backends/win/` | Windows COM backend, bridge, type stubs, error/formatting helpers |
| `backends/mac/` | macOS AppleScript backend, bridge, helpers |
| `tools/` | Leaf — `_folder_constants.py`; no child AGENTS.md |
