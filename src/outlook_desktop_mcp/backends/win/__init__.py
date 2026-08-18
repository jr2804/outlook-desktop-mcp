"""Windows-only backend package: COM bridge + :class:`ComBackend`."""

from outlook_desktop_mcp.backends.win.backend import ComBackend
from outlook_desktop_mcp.backends.win.bridge import OutlookBridge

__all__ = ["ComBackend", "OutlookBridge"]
