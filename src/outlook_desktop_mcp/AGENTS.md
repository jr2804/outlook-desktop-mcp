# AGENTS.md — src/outlook_desktop_mcp

## Purpose

Core package of outlook-desktop-mcp: exposes a running Outlook Desktop as an
MCP server. One codebase, two backends — COM (Windows) and AppleScript (macOS).

## Ownership

- `entrypoint.py` — platform routing. The only module allowed to switch on `sys.platform`.
- `server.py` — Windows COM MCP tool surface (30 tools).
- `server_mac.py` — macOS AppleScript MCP tool surface (23 tools).
- `com_bridge.py` / `applescript_bridge.py` — backend transport; own all direct Outlook interaction.
- `tools/`, `utils/`, `_types.py` — constants, helpers, COM type stubs.

## Local Contracts

- Platform routing: `entrypoint.py` dispatches `darwin → server_mac`, else `server`; tool modules never branch on OS.
- Tool surface: each tool is an `async def` with `@mcp.tool()`; the docstring is the LLM-facing contract — keep it detailed and current.
- Bridge functions receive `(outlook, namespace)` as first args (COM) or the bridge object (macOS).
- Constants: import from `tools/_folder_constants.py`; never hard-code Outlook enum values elsewhere.
- Errors: convert backend exceptions to client-safe text via `utils/errors.py`.

## Work Guidance

Adding a tool: 1) backend function, 2) `@mcp.tool()` async handler with detailed
docstring, 3) test coverage, 4) update the `instructions` string if a capability
category changes. See `CONTRIBUTING.md`.

## Verification

- `uv run pytest` — protocol-level tests.
- `outlook-desktop-mcp.cmd test` — COM validation (Windows, live Outlook).

## Child DOX Index

| Path | Scope |
|------|-------|
| `tools/` | Leaf — `_folder_constants.py`; no child AGENTS.md |
| `utils/` | Leaf — helpers; no child AGENTS.md |
