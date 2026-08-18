"""Platform-independent backend base: :class:`Backend` and :class:`BackendError`.

Re-exports for backwards-compatible imports::

    from outlook_desktop_mcp.backends.base import Backend, BackendError

are still valid because this package re-exports them below.
"""

from outlook_desktop_mcp.backends.base.backend import Backend, BackendError

__all__ = ["Backend", "BackendError"]
