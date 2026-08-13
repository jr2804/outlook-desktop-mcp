# AGENTS.md — MAINTENANCE

How to keep `.agents/` files current.

## Process

1. When a rule changes, update the owning file — never another file.
2. When adding a new `.agents/` file or directory, add it to the index table
   in root `AGENTS.md`.
3. After any `.agents/` change, check the index for stale entries.

## Principles

- **Single source of truth.** Each rule lives in exactly one file. The tier,
  pointer, rationale, and size rules are in **DOX authoring**
  (`.agents/POLICIES.md`) — consult it on every DOX change.
- **AGENTS.md is an index.** It points to `.agents/` files, does not replace them.

## File update triggers

| File / Dir | Update when |
| ------------ | ------------- |
| `ONBOARDING.md` | Project structure, tooling, or entry points change |
| `POLICIES.md` | Boundaries, priorities, or verification change |
| `FILES.md` | Path constants, config keys, or naming conventions change |
| `HISTORY.md` | Notable decision made or resolved |
| `MAINTENANCE.md` | Maintenance procedures, triggers, or principles change |
| `plans/` | New feature implementation starts |
| `history/` | HISTORY.md overflow or completed plans archived |
