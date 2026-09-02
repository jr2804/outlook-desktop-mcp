"""Platform-aware entry point for outlook-desktop-mcp.

This is the ONLY module that calls :func:`current_platform` to select a
backend; everything else stays platform-independent. The backend is
installed into the unified server via ``server.set_backend`` before
``server.main()`` runs the MCP stdio transport.
"""

from __future__ import annotations

from outlook_desktop_mcp import server
from outlook_desktop_mcp.backends.mac import AppleScriptBackend
from outlook_desktop_mcp.platform import Platform, current_platform


def main() -> None:
    """Select the platform backend, inject it, and run the MCP server."""
    platform = current_platform()

    if platform == Platform.DARWIN:
        backend = AppleScriptBackend()
    else:  # Platform.WINDOWS
        from outlook_desktop_mcp.backends.win import ComBackend  # noqa: PLC0415

        backend = ComBackend()

    server.set_backend(backend, platform=platform)
    server.main()


if __name__ == "__main__":
    main()
