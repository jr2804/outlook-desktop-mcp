"""Common base class for platform-specific Outlook bridges.

A bridge owns the lifecycle of a single Outlook automation transport (a
COM thread on Windows, an osascript subprocess on macOS) and exposes an
async ``call`` / ``run`` method that backends use to execute work.

Subclasses implement :meth:`start`, :meth:`stop`, and the async ``run``
method. Backends must only import the concrete subclass from
``outlook_desktop_mcp.backends.win`` or ``outlook_desktop_mcp.backends.mac``.
"""

from __future__ import annotations

from abc import ABC, abstractmethod


class BridgeBase(ABC):
    """Abstract base class for Outlook automation bridges."""

    name: str = "abstract"

    @abstractmethod
    async def start(self) -> None:
        """Initialize the bridge and verify Outlook is reachable."""

    @abstractmethod
    async def stop(self) -> None:
        """Release any bridge-held resources."""

    # Convenience helpers shared by all bridges ------------------------------

    def is_running(self) -> bool:
        """Return True if the bridge has been started and not stopped.

        Subclasses set :attr:`_running` after :meth:`start` completes.
        """
        return getattr(self, "_running", False)
