"""macOS-only backend package: AppleScript bridge + :class:`AppleScriptBackend`."""

from outlook_desktop_mcp.backends.mac.backend import AppleScriptBackend
from outlook_desktop_mcp.backends.mac.bridge import AppleScriptBridge

__all__ = ["AppleScriptBackend", "AppleScriptBridge"]
