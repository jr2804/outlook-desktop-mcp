"""COM Threading Bridge
====================
Runs all Outlook COM calls on a dedicated STA (Single-Threaded Apartment)
thread so the async MCP event loop never touches COM objects directly.

Every COM function passed to ``bridge.call()`` receives ``(outlook, namespace, ...)``
as its first two arguments — the live COM objects that only exist on the
COM thread.
"""

from __future__ import annotations

import asyncio
import logging
import queue
import threading
from collections.abc import Callable
from typing import Any

from outlook_desktop_mcp.backends.base.bridge import BridgeBase

logger = logging.getLogger("outlook_desktop_mcp.backends.win.bridge")

COM_OPERATION_TIMEOUT = 60  # seconds
COM_INIT_TIMEOUT = 15  # seconds


class OutlookBridge(BridgeBase):
    """Manages a dedicated COM thread for Outlook operations."""

    name = "com"

    def __init__(self) -> None:
        self._thread: threading.Thread | None = None
        self._request_queue: queue.Queue = queue.Queue()
        self._outlook = None
        self._namespace = None
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._init_error: Exception | None = None
        self._running = False

    async def start(self) -> None:
        """Start the COM thread. Call once at server startup."""
        self._thread = threading.Thread(target=self._com_thread_main, daemon=True, name="outlook-com")
        self._thread.start()
        if not self._ready.wait(timeout=COM_INIT_TIMEOUT):
            if self._init_error:
                raise self._init_error
            raise RuntimeError(f"Outlook COM thread failed to initialize within {COM_INIT_TIMEOUT}s. Is Outlook Desktop (Classic) running?")
        self._running = True

    async def call(self, func: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
        """Schedule a function to run on the COM thread and await its result.

        The function signature must be: ``func(outlook, namespace, *args, **kwargs)``.
        """
        result_event = threading.Event()
        result_holder: dict = {}
        self._request_queue.put((func, args, kwargs, result_event, result_holder))

        loop = asyncio.get_running_loop()
        signaled = await loop.run_in_executor(None, lambda: result_event.wait(timeout=COM_OPERATION_TIMEOUT))
        if not signaled:
            raise TimeoutError(
                f"Outlook COM operation timed out after {COM_OPERATION_TIMEOUT} seconds. Outlook may be waiting for user input (e.g., a dialog box)."
            )

        if "error" in result_holder:
            raise result_holder["error"]
        return result_holder.get("value")

    async def stop(self) -> None:
        """Signal the COM thread to shut down."""
        self._shutdown.set()
        if self._thread:
            self._thread.join(timeout=5)
        self._running = False

    def _com_thread_main(self) -> None:
        """Main loop for the COM thread."""
        # ty: ignore[unresolved-import]
        import pythoncom  # noqa: PLC0415

        # ty: ignore[unresolved-import]
        import win32com.client  # noqa: PLC0415

        pythoncom.CoInitialize()
        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
            store_name = self._namespace.DefaultStore.DisplayName
            user_name = self._namespace.CurrentUser.Name
            logger.debug("COM thread ready. Store: %s, User: %s", store_name, user_name)
            self._ready.set()

            while not self._shutdown.is_set():
                try:
                    func, args, kwargs, result_event, result_holder = self._request_queue.get(timeout=0.5)
                except queue.Empty:
                    continue
                try:
                    result_holder["value"] = func(self._outlook, self._namespace, *args, **kwargs)
                except Exception as e:
                    result_holder["error"] = e
                finally:
                    result_event.set()
        except Exception as e:
            self._init_error = e
            self._ready.set()  # Unblock the caller so they see the error
            logger.error("COM thread init failed: %s", e)
        finally:
            pythoncom.CoUninitialize()
