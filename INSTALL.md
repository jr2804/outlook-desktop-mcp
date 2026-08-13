# INSTALL — Scaffold a project with AGENTS.md

Feed this file to your coding agent to set up the DOX hierarchy.

> **⚠ Note on `.agents/` directory**
> The `.agents/` subdirectory is a common convention — tools, skills, and
> agent runtimes may already create one in your project. It is also
> frequently listed in `.gitignore` patterns (especially when added by
> Copilot, Cursor, or similar tooling). **You must add git exclusion
> rules** in your project's `.gitignore` so that the DOX markdown files
> under `.agents/` are tracked by version control:
>
> ```gitignore
> # Keep .agents/ DOX .md files tracked (override any blanket ignore)
> /.agents/*
> !/.agents/*.md
> !/.agents/history/
> !/.agents/history/**
>
> # Implementation plans — mutable working artifacts, never committed
> /.agents/plans/
> ```
>
> The first two `!` lines un-ignore `.md` files directly in `.agents/` and
> the `.agents/history/` subdirectory (durable decision records). The
> `/.agents/plans/` line explicitly excludes implementation plans.
>
> **If your project's `.gitignore` already has a blanket ignore for this
> directory** (e.g. `/.agents/`, `/.agents`, or just `.agents`), the
> first line is redundant — you only need the `!` lines:
>
> ```gitignore
> # Keep .agents/ DOX files tracked (override existing ignore)
> !/.agents/*.md
> !/.agents/history/
> !/.agents/history/**
>
> # Implementation plans — mutable working artifacts, never committed
> /.agents/plans/
> ```
>
> **Conversely, if your project already has explicit inclusion rules**
> for `.agents/` (e.g. `!.agents/` + `!.agents/**`, or a
> project-specific pattern that already covers `.md` files), then no
> gitignore change is needed at all — skip the whole block.
>
> In all cases, the goal is the same: only DOX `.md` files and
> `.agents/history/` are tracked; implementation plans and anything else
> tools may drop in `.agents/` (configs, binaries, lock files) stay
> ignored.

## Files to create

### Root `AGENTS.md`

Copy AGENTS.md from agents-scaffold. Fill sections 2–4:

**AGENTS.md §4 (Project rules):** 5-10 essential, frequently-broken rules that
must be injected on every LLM call. Examples: path conventions, import rules,
naming patterns, required config keys. Everything non-essential goes into
`.agents/` files.

### `.agents/ONBOARDING.md`

One-line project description. Entry points. Build/test commands. Reference
to tools (codegraph, grepai, repos, etc. if available).

### `.agents/POLICIES.md`

Keep the generic boundaries/priorities/verification template. Add
project-specific policies under the correct section.

### `.agents/FILES.md`

Map each category of definition to its source file. Example:

| What            | Where                   |
| --------------- | ----------------------- |
| Path constants  | `src/project/paths.py`  |
| Config defaults | `src/project/config.py` |

### `.agents/HISTORY.md`

Start empty. Add entries as decisions are made.

### `.agents/MAINTENANCE.md`

Keep generic template. Only change update triggers if project has unique needs.

## Update mode (when files already exist)

If `AGENTS.md` and `.agents/` already exist, this is an **update**, not a
first-time install. Do not bail out — the template may have changed.

### What to do

1. **Fetch the latest `INSTALL.md`** from the same remote URL. Use it as the
   source of truth for what the current template looks like.
2. **Diff each file structurally** — do not just check existence. Compare
   sections, rules, and entries against what is on disk.
3. **Apply additions** — new sections in `POLICIES.md`, new files/dirs in
   `.agents/`, new gitignore patterns.
4. **Apply rewordings** — if a rule's wording changed in the template (e.g.
   "No history" → "No history in AGENTS.md files"), update it.
5. **Preserve project-specific customizations** — project rules in
   `AGENTS.md`, `FILES.md` entries, `HISTORY.md` entries, and any
   project-specific policies added under the correct sections in
   `POLICIES.md`. These are the user's content, not template content.
6. **Report what changed** — list each file and whether it was added,
   updated, or left unchanged.

### What NOT to do

- Do NOT delete or overwrite project-specific content (project rules,
  FILES.md entries, HISTORY entries, custom policies).
- Do NOT skip the update just because files exist — the template evolves.
- Do NOT re-add content that was intentionally removed by the user.

## Migration instructions

### Review existing AGENTS.md

1. **Move project description →** `README.md` or `docs/`. Keep max 1 sentence
   in AGENTS.md.
2. **Delete tree views** — generate on demand with `rg --files | tree-cli --fromfile`.
3. **Delete history entries** — move important ones to `.agents/HISTORY.md`
   with git refs. Routine history stays in git log.
4. **Delete TODO lists** — use an issue tracker or beads.
5. **Move non-essential rules →** `.agents/POLICIES.md` or subtree AGENTS.md.
   Only inject what is needed on every LLM call.
6. **Move path/constant definitions →** `.agents/FILES.md`.
7. **Add ⛔ No Patching section** at end of root AGENTS.md.

### Check for tool patches

Search AGENTS.md for content inserted by tools (beads, linters, CI bots).
If found:

- Move legitimate rules to the correct `.agents/` file.
- Delete everything after the ⛔ No Patching block.

### Create subtree AGENTS.md files

For directories with dense local contracts:

```text
src/api/AGENTS.md     → API-specific rules
tests/AGENTS.md       → Testing conventions
```

These are demand-loaded on entry and keep root AGENTS.md lean.
