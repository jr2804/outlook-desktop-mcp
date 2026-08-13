# AGENTS.md — ONBOARDING

Read this when starting a new session. After first read, only revisit when
project structure or tooling changes significantly.

## Project

Outlook Desktop as an MCP server — Windows (COM) and macOS (AppleScript).
No Graph API, no Entra app registration, no OAuth — local Outlook only.
Full docs: `README.md`.

## Quick start

```bash
# install (dev)
uv sync

# test — protocol level, no Outlook required
uv run pytest

# COM validation (Windows, live Outlook required)
outlook-desktop-mcp.cmd test

# run MCP server (stdio)
outlook-desktop-mcp.cmd mcp
```

## Entry points (read these first)

| File | Why |
| ------ | ----- |
| `AGENTS.md` | Root rail — rules + `.agents/` index |
| `.agents/POLICIES.md` | Boundaries, priorities, verification |
| `.agents/FILES.md` | Source-of-truth locations |

## Where to dig deeper

- `src/outlook_desktop_mcp/entrypoint.py` — platform routing (`darwin` → `server_mac`, else `server`)
- `src/outlook_desktop_mcp/server.py` — Windows COM MCP server (30 tools)
- `src/outlook_desktop_mcp/server_mac.py` — macOS AppleScript MCP server (23 tools)
- `src/outlook_desktop_mcp/com_bridge.py` / `applescript_bridge.py` — transport layer
- `docs/` — user-facing documentation
- Source-tree AGENTS.md files — local contracts (see root `AGENTS.md` index)

## Available tools

If MCP/CLI tools are available for this workspace:

- **bd** / beads — issue tracking (installed via mise)
- **rg** / **tree-cli** — file listing, on-demand tree views
- **uv** / **mise** — environment and tool management
