"""AppleScript Bridge
==================
Runs Outlook automation via ``osascript`` subprocess calls on macOS.
Each call is stateless — no persistent COM-like objects.

Every tool builds an AppleScript string and passes it to ``bridge.run()``.
"""
from __future__ import annotations

import asyncio
import logging

from outlook_desktop_mcp.backends.base.bridge import BridgeBase

logger = logging.getLogger("outlook_desktop_mcp.backends.mac.bridge")

STARTUP_TIMEOUT = 10
SCRIPT_TIMEOUT = 30


class AppleScriptBridge(BridgeBase):
    """Manages AppleScript execution for Outlook on macOS."""

    name = "applescript"

    def __init__(self) -> None:
        self._version: str | None = None
        self._running = False

    async def start(self) -> None:
        """Verify Outlook is running and accessible. Call once at server startup."""
        try:
            self._version = await self.run(
                'tell application "Microsoft Outlook" to get version',
                timeout=STARTUP_TIMEOUT,
            )
            logger.info("AppleScript bridge ready. Outlook version: %s", self._version)
            self._running = True
        except Exception as e:
            raise RuntimeError(
                f"Cannot connect to Microsoft Outlook via AppleScript. "
                f"Is Outlook running? Error: {e}"
            ) from e

    async def run(self, script: str, timeout: float = SCRIPT_TIMEOUT) -> str:
        """Execute an AppleScript and return stdout as a string.

        Raises RuntimeError on non-zero exit or timeout.
        """
        proc = await asyncio.create_subprocess_exec(
            "osascript", "-e", script,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        try:
            stdout, stderr = await asyncio.wait_for(
                proc.communicate(), timeout=timeout
            )
        except asyncio.TimeoutError:
            proc.kill()
            await proc.communicate()
            raise RuntimeError(f"AppleScript timed out after {timeout}s")

        if proc.returncode != 0:
            err = stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"AppleScript error: {err}")

        return stdout.decode("utf-8", errors="replace").strip()

    async def run_lines(self, script: str, timeout: float = SCRIPT_TIMEOUT) -> list[str]:
        """Execute an AppleScript and return output split into non-empty lines."""
        result = await self.run(script, timeout=timeout)
        return [line for line in result.split("\n") if line.strip()]

    async def stop(self) -> None:
        """No-op — AppleScript has no persistent resources to release."""
        self._running = False
