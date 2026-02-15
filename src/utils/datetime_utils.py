"""Datetime parsing and timezone utilities."""

import subprocess
from datetime import date, datetime, timezone
from zoneinfo import ZoneInfo


_local_tz: ZoneInfo | None = None


def get_local_timezone() -> ZoneInfo:
    """Get the Mac's local timezone as a ZoneInfo object. Cached after first call."""
    global _local_tz
    if _local_tz is not None:
        return _local_tz

    try:
        result = subprocess.run(
            ["systemsetup", "-gettimezone"],
            capture_output=True, text=True, timeout=5,
        )
        # Output: "Time Zone: America/New_York"
        tz_name = result.stdout.strip().split(": ", 1)[-1]
        _local_tz = ZoneInfo(tz_name)
    except Exception:
        # Fallback: derive from system datetime
        _local_tz = datetime.now().astimezone().tzinfo  # type: ignore[assignment]
    return _local_tz


def parse_applescript_components(
    year: int, month: int, day: int,
    hour: int = 0, minute: int = 0, second: int = 0,
) -> datetime:
    """Build a timezone-aware datetime from AppleScript numeric components.

    AppleScript dates are in the local timezone.
    """
    tz = get_local_timezone()
    return datetime(year, month, day, hour, minute, second, tzinfo=tz)


def to_google_datetime(dt: datetime) -> str:
    """Format a datetime for Google Calendar API (RFC 3339)."""
    return dt.isoformat()


def to_google_date(d: date) -> str:
    """Format a date for Google Calendar all-day events."""
    return d.isoformat()


def from_google_datetime(dt_str: str, tz_name: str | None = None) -> datetime:
    """Parse a Google Calendar API datetime string into a timezone-aware datetime."""
    from dateutil.parser import isoparse
    dt = isoparse(dt_str)
    if dt.tzinfo is None:
        tz = ZoneInfo(tz_name) if tz_name else get_local_timezone()
        dt = dt.replace(tzinfo=tz)
    return dt


def from_google_date(date_str: str) -> date:
    """Parse a Google Calendar all-day event date string."""
    return date.fromisoformat(date_str)


def to_applescript_date_components(dt: datetime) -> dict:
    """Convert a datetime to components for AppleScript date construction.

    Returns dict with year, month, day, hour, minute, second in local timezone.
    """
    local_dt = dt.astimezone(get_local_timezone())
    return {
        "year": local_dt.year,
        "month": local_dt.month,
        "day": local_dt.day,
        "hour": local_dt.hour,
        "minute": local_dt.minute,
        "second": local_dt.second,
    }
