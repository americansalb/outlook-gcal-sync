"""Python-to-AppleScript interface via osascript subprocess."""

import logging
import subprocess
import time

logger = logging.getLogger("outlook_gcal_sync.outlook.bridge")


class AppleScriptError(Exception):
    """Raised when an AppleScript command fails."""

    def __init__(self, message: str, stderr: str = "", return_code: int = -1):
        super().__init__(message)
        self.stderr = stderr
        self.return_code = return_code


class OutlookNotRunningError(AppleScriptError):
    """Raised when Outlook is not running."""


def run_applescript(script: str, timeout: int = 60) -> str:
    """Execute an AppleScript string via osascript and return stdout.

    Args:
        script: The AppleScript source code to execute.
        timeout: Maximum seconds to wait for execution.

    Returns:
        The stdout output from osascript.

    Raises:
        AppleScriptError: If the script fails.
        OutlookNotRunningError: If Outlook is not running.
    """
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        raise AppleScriptError(f"AppleScript timed out after {timeout}s")

    if result.returncode != 0:
        stderr = result.stderr.strip()
        if "Application isn't running" in stderr or "-600" in stderr:
            raise OutlookNotRunningError(
                "Microsoft Outlook is not running. Please launch Outlook first.",
                stderr=stderr,
                return_code=result.returncode,
            )
        raise AppleScriptError(
            f"AppleScript failed (code {result.returncode}): {stderr}",
            stderr=stderr,
            return_code=result.returncode,
        )

    return result.stdout.strip()


def run_applescript_with_retry(
    script: str, timeout: int = 60, max_retries: int = 3, backoff_base: float = 2.0,
) -> str:
    """Execute AppleScript with retry and exponential backoff.

    Retries on transient errors like Outlook being busy (-1708).
    Does NOT retry if Outlook is not running.
    """
    last_error = None
    for attempt in range(max_retries):
        try:
            return run_applescript(script, timeout=timeout)
        except OutlookNotRunningError:
            raise  # Don't retry if Outlook isn't running
        except AppleScriptError as e:
            last_error = e
            if attempt < max_retries - 1:
                wait = backoff_base ** attempt
                logger.warning(
                    "AppleScript attempt %d/%d failed: %s. Retrying in %.1fs...",
                    attempt + 1, max_retries, e, wait,
                )
                time.sleep(wait)
    raise last_error  # type: ignore[misc]


def escape_applescript_string(value: str) -> str:
    """Escape a string for safe embedding in AppleScript string literals.

    Handles backslashes, double quotes, and newlines.
    """
    return (
        value
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", "\\n")
        .replace("\r", "\\r")
        .replace("\t", "\\t")
    )


def check_outlook_running() -> bool:
    """Check if Microsoft Outlook is currently running."""
    try:
        result = run_applescript(
            'tell application "System Events" to return '
            '(name of every process) contains "Microsoft Outlook"',
            timeout=10,
        )
        return result.lower() == "true"
    except AppleScriptError:
        return False


def list_calendars() -> list[str]:
    """List all calendar names in Outlook."""
    script = 'tell application "Microsoft Outlook" to return name of every calendar'
    result = run_applescript(script)
    if not result:
        return []
    return [name.strip() for name in result.split(", ")]
