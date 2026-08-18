"""Landing-page instructions for the MCP server, composed per platform.

This text is shown to clients (e.g. Claude) when the MCP server is loaded.
It must remain accurate and stable — when the tool list changes, update it
here as well.
"""

from __future__ import annotations

from outlook_desktop_mcp.platform import Platform

_BASE_INSTRUCTIONS = """\
This MCP server gives you full access to Microsoft Outlook Desktop on \
{platform_label}. It can send emails, read inbox messages, search across \
folders, mark messages as read/unread, move messages between folders \
(including archive), reply to emails, manage calendar events and meetings, \
manage tasks, and list the complete folder hierarchy.

All operations use the locally running Outlook app — no Microsoft Graph \
API, no Entra app registration, no OAuth tokens needed. The user's existing \
Outlook session handles all authentication.

PREREQUISITE: Outlook must be running. {prerequisite}

NOTE: entry_id values are platform-specific (hex strings on Windows, numeric \
IDs on macOS). Always take them from list/search results on the same machine.

The 'account' parameter selects a mailbox on Windows; on macOS it is \
accepted but ignored (AppleScript always targets the active account).

AVAILABLE TOOL CATEGORIES:
- Email: send, list, read, search, reply, mark read/unread, move, attachments
- Drafts: draft_email (save to Drafts), draft_reply_email (draft a reply)
- Calendar: list events, create appointments/meetings, update, delete, search events
- Tasks: create, list, complete, delete to-do items
- Out of Office: enable/disable auto-reply
- Folders: list folder hierarchy with item counts
{windows_only}\
"""

_PLATFORM_LABELS = {
    Platform.DARWIN: "macOS (AppleScript)",
    Platform.WINDOWS: "Windows (COM)",
}

_PREREQUISITES = {
    Platform.DARWIN: "Microsoft Outlook for Mac must be running.",
    Platform.WINDOWS: ("On Windows, classic OUTLOOK.EXE only (the new/modern Outlook olk.exe is NOT supported)."),
}

_WINDOWS_ONLY = """\
Windows only (not registered on macOS):
- Accounts: list_accounts
- Meeting responses: respond_to_meeting
- Categories: list and set color categories
- Rules: list and manage mail rules
- Out of Office status query: get_out_of_office
"""


def build_instructions(platform: Platform) -> str:
    """Compose the instructions string shown to clients on tool listing."""
    return _BASE_INSTRUCTIONS.format(
        platform_label=_PLATFORM_LABELS[platform],
        prerequisite=_PREREQUISITES[platform],
        windows_only=_WINDOWS_ONLY if platform == Platform.WINDOWS else "",
    ).rstrip()
