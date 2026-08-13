# AGENTS.md — FILES

Single source of truth for paths, config keys, and naming conventions.
Kept compact — agents hallucinate less when they know where definitions live.

## Pattern

- One file owns each class of definition (paths, config defaults, enums).
- Import from that file. Never hard-code values in other modules.
- Variables that address files get `_file` suffix; directories get `_dir`.

## Project-specific sources of truth

| What | Where | Key names |
| ------ | ------- | ----------- |
| Outlook folder enums & name→enum map | `src/outlook_desktop_mcp/tools/_folder_constants.py` | `OL_FOLDER_*`, `FOLDER_NAME_TO_ENUM`, `OL_MAIL_ITEM`, `OL_APPOINTMENT_ITEM`, `OL_TASK_ITEM` |
| COM type stubs (Protocols) | `src/outlook_desktop_mcp/_types.py` | `Outlook`, `Namespace`, `Folder`, `MailItem`, `Appointment`, `TaskItem`, `Store` |
| Error formatting | `src/outlook_desktop_mcp/utils/errors.py` | `format_com_error` |
| Summary/full formatting | `src/outlook_desktop_mcp/utils/formatting.py` | `format_email_summary`, `format_event_summary`, `format_task_summary` |
| AppleScript helpers | `src/outlook_desktop_mcp/utils/applescript_helpers.py` | `escape`, `format_date`, `parse_date`, `FOLDER_MAP`, `resolve_folder_ref` |
| Package config, deps, pytest opts | `pyproject.toml` | name/version, `dependencies`, `[project.scripts]`, `[tool.pytest.ini_options]` |
| Platform routing | `src/outlook_desktop_mcp/entrypoint.py` | `sys.platform` dispatch |
