"""Platform identifiers used across the package.

Centralised so ``sys.platform`` is the only place that switches on OS —
every other consumer compares against :class:`Platform` values.
"""

from __future__ import annotations

import sys
from enum import StrEnum


class Platform(StrEnum):
    """Supported target platforms."""

    DARWIN = "darwin"
    WINDOWS = "win32"


def current_platform() -> Platform:
    """Map the current ``sys.platform`` to a :class:`Platform` value."""
    if sys.platform.startswith("darwin"):
        return Platform.DARWIN
    if sys.platform.startswith("win"):
        return Platform.WINDOWS
    raise RuntimeError(f"Unsupported platform: {sys.platform!r}. outlook-desktop-mcp only runs on Windows or macOS.")
