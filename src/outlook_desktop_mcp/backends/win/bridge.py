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
import contextlib
import logging
import queue
import threading
import time
from collections.abc import Callable
from typing import Any

import pythoncom
import win32com.client

from outlook_desktop_mcp.backends.base.bridge import BridgeBase

COM_OPERATION_TIMEOUT = 60  # seconds
COM_INIT_TIMEOUT = 15  # seconds

# Default idle time after which the cached Outlook RCW is released. A new
# Dispatch starts on the next access. Keeping this on the higher side is fine:
# re-dispatch is cheap and we only do it after a full hour of inactivity.
DEFAULT_IDLE_TIMEOUT = 3600.0

logger = logging.getLogger("outlook_desktop_mcp.backends.win.bridge")


class OutlookBridge(BridgeBase):
    """Manages a dedicated COM thread for Outlook operations."""

    name = "com"

    def __init__(self, idle_timeout: float = DEFAULT_IDLE_TIMEOUT) -> None:
        self._thread: threading.Thread | None = None
        self._request_queue: queue.Queue = queue.Queue()
        self._outlook: Any = None
        self._namespace: Any = None
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._init_error: Exception | None = None
        self._running = False
        self._idle_timeout = float(idle_timeout)
        self._last_used_ts: float = time.monotonic()

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

        The function is invoked with ``(outlook, namespace, *args, **kwargs)`` on
        the COM thread.
        """
        result_event = threading.Event()
        result_holder: dict[str, Any] = {}

        def _run() -> None:
            self._check_idle_release()
            try:
                result_holder["value"] = self._run_with_recovery(func, args, kwargs)
                result_holder["ok"] = True
            except Exception as e:
                result_holder["error"] = e
            finally:
                result_event.set()

        await asyncio.get_event_loop().run_in_executor(None, self._request_queue.put, _run)

        try:
            await asyncio.wait_for(_wait_for_event(result_event), timeout=COM_OPERATION_TIMEOUT)
        except TimeoutError:
            raise TimeoutError(f"COM operation timed out after {COM_OPERATION_TIMEOUT}s") from None

        if "error" in result_holder:
            raise result_holder["error"]
        return result_holder.get("value")

    async def stop(self) -> None:
        """Signal the COM thread to shut down."""
        self._shutdown.set()
        if self._thread:
            self._thread.join(timeout=5)
        self._running = False

    # ------------------------------------------------------------------ COM thread

    def _com_thread_main(self) -> None:
        """Main loop for the COM thread."""
        pythoncom.CoInitialize()
        try:
            try:
                self._connect()
                store_name = self._namespace.DefaultStore.DisplayName
                user_name = self._namespace.CurrentUser.Name
                logger.debug("COM thread ready. Store: %s, User: %s", store_name, user_name)
            except Exception as e:
                self._init_error = e
                self._ready.set()  # Unblock the caller so they see the error
                logger.error("COM thread init failed: %s", e)
                return
            self._ready.set()

            while not self._shutdown.is_set():
                try:
                    request = self._request_queue.get(timeout=0.5)
                except queue.Empty:
                    continue
                # Exceptions are already reported to the caller via the
                # ``result_holder`` populated inside ``_run_with_recovery``;
                # we never want them to escape the COM thread loop.
                with contextlib.suppress(Exception):
                    request()
        finally:
            pythoncom.CoUninitialize()

    def _connect(self) -> None:
        """Dispatch Outlook and cache both the Application and Namespace."""
        self._outlook = win32com.client.Dispatch("Outlook.Application")
        self._namespace = self._outlook.GetNamespace("MAPI")
        self._last_used_ts = time.monotonic()

    def _release(self, reason: str) -> None:
        """Drop the cached COM references so the next call re-dispatches."""
        if self._outlook is None:
            return
        logger.info("releasing_outlook_com_object reason=%s", reason)
        self._outlook = None
        self._namespace = None

    def _check_idle_release(self) -> None:
        if self._outlook is None:
            return
        if time.monotonic() - self._last_used_ts >= self._idle_timeout:
            self._release(reason="idle_ttl_expired")

    def _run_with_recovery(
        self,
        func: Callable[..., Any],
        args: tuple[Any, ...],
        kwargs: dict[str, Any],
    ) -> Any:
        """Run a closure with at most one reconnect-and-retry on any failure.

        The cached ``_outlook``/``_namespace`` may be stale (Outlook restarted,
        disconnected RCW after long idle). On any failure we drop the cache,
        re-dispatch, and try the closure once more. The second failure is the
        one callers see — no silent retry loops.
        """
        if self._outlook is None:
            self._connect()

        try:
            value = func(self._outlook, self._namespace, *args, **kwargs)
        except Exception as first_error:
            logger.warning("com_call_failed_reconnecting error=%s", first_error)
            self._release(reason="call_failure")
            try:
                self._connect()
            except Exception as reconnect_error:
                logger.error("com_reconnect_failed error=%s", reconnect_error)
                # If reconnect itself fails, surface the original failure cause —
                # that's the diagnostic the caller (and user) actually cares about.
                raise first_error from None
            try:
                value = func(self._outlook, self._namespace, *args, **kwargs)
            except Exception as second_error:
                logger.error("com_call_failed_after_reconnect error=%s", second_error)
                raise second_error from None

        self._last_used_ts = time.monotonic()
        return value


async def _wait_for_event(event: threading.Event) -> None:
    """Await a threading.Event without blocking the event loop."""
    while not event.is_set():
        await asyncio.sleep(0.01)
