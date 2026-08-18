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

- `src/outlook_desktop_mcp/entrypoint.py` — platform routing (`current_platform()` → `backends/{win,mac}`)
- `src/outlook_desktop_mcp/server.py` — unified MCP tool surface (33 Win / 26 Mac tools)
- `src/outlook_desktop_mcp/platform.py` — `Platform` StrEnum
- `src/outlook_desktop_mcp/instructions.py` — `build_instructions(platform)`
- `src/outlook_desktop_mcp/backends/{win,mac}/` — platform-specific bridges and helpers
- `docs/` — user-facing documentation
- Source-tree AGENTS.md files — local contracts (see root `AGENTS.md` index)

## Available tools

If MCP/CLI tools are available for this workspace:

- **bd** / beads — issue tracking (installed via mise)
- **rg** / **tree-cli** — file listing, on-demand tree views
- **uv** / **mise** — environment and tool management
