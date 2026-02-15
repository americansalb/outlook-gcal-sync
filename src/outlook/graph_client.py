"""Microsoft Graph Calendar API client using raw REST calls."""

import logging
import re
from datetime import date, datetime, timedelta
from html.parser import HTMLParser

import requests

from src.sync.models import UnifiedEvent
from src.utils.datetime_utils import get_local_timezone

logger = logging.getLogger("outlook_gcal_sync.microsoft.client")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class _HTMLTextExtractor(HTMLParser):
    """Strip HTML tags and return plain text."""

    def __init__(self):
        super().__init__()
        self._parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self._parts.append(data)

    def get_text(self) -> str:
        return "".join(self._parts).strip()


def _strip_html(html: str) -> str:
    """Convert HTML body to plain text."""
    if not html:
        return ""
    extractor = _HTMLTextExtractor()
    extractor.feed(html)
    return extractor.get_text()


class GraphCalendarClient:
    """Thin wrapper around Microsoft Graph Calendar API endpoints."""

    def __init__(self, access_token: str):
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        })

    def _get(self, url: str, params: dict | None = None) -> dict:
        resp = self.session.get(url, params=params)
        resp.raise_for_status()
        return resp.json()

    def _post(self, url: str, json_body: dict) -> dict:
        resp = self.session.post(url, json=json_body)
        resp.raise_for_status()
        return resp.json()

    def _patch(self, url: str, json_body: dict) -> dict:
        resp = self.session.patch(url, json=json_body)
        resp.raise_for_status()
        return resp.json()

    def _delete(self, url: str) -> None:
        resp = self.session.delete(url)
        resp.raise_for_status()

    def list_calendars(self) -> list[dict]:
        """List all calendars for the authenticated user.

        Returns:
            List of calendar dicts with 'id', 'name', 'isDefaultCalendar', etc.
        """
        url = f"{GRAPH_BASE}/me/calendars"
        all_calendars = []
        while url:
            data = self._get(url)
            all_calendars.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        logger.info("Listed %d calendars from Microsoft Graph.", len(all_calendars))
        return all_calendars

    def resolve_calendar_id(self, calendar_name: str) -> str:
        """Resolve a calendar name to its Graph ID.

        If calendar_name is empty, returns the default calendar ID.

        Raises:
            ValueError: If no matching calendar is found.
        """
        calendars = self.list_calendars()
        if not calendar_name:
            for cal in calendars:
                if cal.get("isDefaultCalendar"):
                    return cal["id"]
            # Fallback to first calendar
            if calendars:
                return calendars[0]["id"]
            raise ValueError("No calendars found in Microsoft account.")

        for cal in calendars:
            if cal.get("name", "").lower() == calendar_name.lower():
                return cal["id"]

        available = [c.get("name", "?") for c in calendars]
        raise ValueError(
            f"Calendar '{calendar_name}' not found. Available: {', '.join(available)}"
        )

    def list_events(
        self,
        calendar_id: str,
        start: datetime,
        end: datetime,
    ) -> list[dict]:
        """List events using calendarView (auto-expands recurring events).

        Args:
            calendar_id: The Graph calendar ID.
            start: Start of date range.
            end: End of date range.

        Returns:
            List of event dicts from Graph API.
        """
        url = f"{GRAPH_BASE}/me/calendars/{calendar_id}/calendarView"
        params = {
            "startDateTime": start.isoformat(),
            "endDateTime": end.isoformat(),
            "$top": "250",
            "$select": "id,subject,body,start,end,location,isAllDay,isReminderOn,reminderMinutesBeforeStart,type,seriesMasterId",
        }

        all_events = []
        while True:
            data = self._get(url, params=params)
            all_events.extend(data.get("value", []))
            next_link = data.get("@odata.nextLink")
            if not next_link:
                break
            url = next_link
            params = None  # nextLink includes all params

        logger.info("Listed %d events from Microsoft Graph calendar.", len(all_events))
        return all_events

    def create_event(self, calendar_id: str, event_body: dict) -> dict:
        """Create a new calendar event.

        Returns:
            The created event dict (includes 'id').
        """
        url = f"{GRAPH_BASE}/me/calendars/{calendar_id}/events"
        result = self._post(url, event_body)
        logger.info("Created Outlook event '%s' (ID: %s)", result.get("subject", ""), result.get("id", ""))
        return result

    def update_event(self, event_id: str, event_body: dict) -> dict:
        """Update an existing event (PATCH — partial update).

        Returns:
            The updated event dict.
        """
        url = f"{GRAPH_BASE}/me/events/{event_id}"
        result = self._patch(url, event_body)
        logger.info("Updated Outlook event '%s' (ID: %s)", result.get("subject", ""), event_id)
        return result

    def delete_event(self, event_id: str) -> None:
        """Delete an event."""
        url = f"{GRAPH_BASE}/me/events/{event_id}"
        self._delete(url)
        logger.info("Deleted Outlook event (ID: %s)", event_id)


# ── Converters ───────────────────────────────────────────────────────────────


def graph_event_to_unified(event: dict) -> UnifiedEvent:
    """Convert a Microsoft Graph event dict to a UnifiedEvent.

    Graph API datetime format:
      Timed: {"dateTime": "2026-02-14T09:00:00.0000000", "timeZone": "Eastern Standard Time"}
      All-day: {"dateTime": "2026-02-14T00:00:00.0000000", "timeZone": "UTC"}

    Args:
        event: A Graph API event resource dict.

    Returns:
        A UnifiedEvent with source="outlook".
    """
    from zoneinfo import ZoneInfo

    is_all_day = event.get("isAllDay", False)
    start_raw = event.get("start", {})
    end_raw = event.get("end", {})

    if is_all_day:
        # All-day events: extract date only
        start_date_str = start_raw.get("dateTime", "")[:10]
        end_date_str = end_raw.get("dateTime", "")[:10]
        start_date = date.fromisoformat(start_date_str)
        end_date = date.fromisoformat(end_date_str)

        tz = get_local_timezone()
        start_time = datetime(start_date.year, start_date.month, start_date.day, tzinfo=tz)
        end_time = datetime(end_date.year, end_date.month, end_date.day, tzinfo=tz)
    else:
        start_time = _parse_graph_datetime(start_raw)
        end_time = _parse_graph_datetime(end_raw)
        start_date = None
        end_date = None

    # Body — Graph returns HTML by default
    body = event.get("body", {})
    description = ""
    if body.get("contentType") == "text":
        description = body.get("content", "")
    elif body.get("contentType") == "html":
        description = _strip_html(body.get("content", ""))

    # Location
    location = event.get("location", {}).get("displayName", "") or ""

    # Reminders
    is_reminder_on = event.get("isReminderOn", False)
    reminder_minutes = event.get("reminderMinutesBeforeStart", 0) if is_reminder_on else 0

    # Recurring
    is_recurring = event.get("type", "singleInstance") != "singleInstance"

    return UnifiedEvent(
        source="outlook",
        source_id=event["id"],
        title=event.get("subject", ""),
        description=description,
        location=location,
        start_time=start_time,
        end_time=end_time,
        is_all_day=is_all_day,
        has_reminder=is_reminder_on,
        reminder_minutes=reminder_minutes,
        is_recurring=is_recurring,
        start_date=start_date,
        end_date=end_date,
    )


def unified_to_graph_body(event: UnifiedEvent) -> dict:
    """Convert a UnifiedEvent to a Microsoft Graph event body for create/update.

    Args:
        event: The UnifiedEvent to convert.

    Returns:
        A dict suitable for POST/PATCH to the Graph events endpoint.
    """
    body: dict = {
        "subject": event.title,
        "body": {
            "contentType": "text",
            "content": event.description,
        },
        "location": {
            "displayName": event.location,
        },
        "isAllDay": event.is_all_day,
    }

    if event.is_all_day and event.start_date:
        body["start"] = {
            "dateTime": datetime(event.start_date.year, event.start_date.month, event.start_date.day).isoformat(),
            "timeZone": "UTC",
        }
        end_date = event.end_date or (event.start_date + timedelta(days=1))
        if end_date == event.start_date:
            end_date = event.start_date + timedelta(days=1)
        body["end"] = {
            "dateTime": datetime(end_date.year, end_date.month, end_date.day).isoformat(),
            "timeZone": "UTC",
        }
    else:
        tz = get_local_timezone()
        # Graph expects Windows timezone names. Use IANA name and let Graph resolve.
        tz_name = str(tz)
        body["start"] = {
            "dateTime": event.start_time.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": tz_name,
        }
        body["end"] = {
            "dateTime": event.end_time.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": tz_name,
        }

    # Reminders
    if event.has_reminder and event.reminder_minutes > 0:
        body["isReminderOn"] = True
        body["reminderMinutesBeforeStart"] = event.reminder_minutes
    else:
        body["isReminderOn"] = False

    return body


# ── Graph datetime parsing ───────────────────────────────────────────────────


# Map of common Windows timezone names to IANA names
_WINDOWS_TO_IANA = {
    "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago",
    "Mountain Standard Time": "America/Denver",
    "Pacific Standard Time": "America/Los_Angeles",
    "UTC": "UTC",
    "GMT Standard Time": "Europe/London",
    "Central European Standard Time": "Europe/Berlin",
    "India Standard Time": "Asia/Kolkata",
    "Tokyo Standard Time": "Asia/Tokyo",
    "China Standard Time": "Asia/Shanghai",
    "AUS Eastern Standard Time": "Australia/Sydney",
}


def _resolve_timezone(tz_name: str | None) -> ZoneInfo:
    """Resolve a timezone name (Windows or IANA) to a ZoneInfo object."""
    from zoneinfo import ZoneInfo

    if not tz_name:
        return get_local_timezone()

    # Try IANA name first
    try:
        return ZoneInfo(tz_name)
    except (KeyError, ValueError):
        pass

    # Try Windows name mapping
    iana = _WINDOWS_TO_IANA.get(tz_name)
    if iana:
        return ZoneInfo(iana)

    logger.warning("Unknown timezone '%s', falling back to local timezone.", tz_name)
    return get_local_timezone()


def _parse_graph_datetime(dt_dict: dict) -> datetime:
    """Parse a Graph API datetime dict into a timezone-aware Python datetime.

    Graph format: {"dateTime": "2026-02-14T09:00:00.0000000", "timeZone": "Eastern Standard Time"}
    """
    from zoneinfo import ZoneInfo

    dt_str = dt_dict.get("dateTime", "")
    tz_name = dt_dict.get("timeZone")

    # Strip fractional seconds beyond 6 digits (Python handles up to microseconds)
    dt_str = re.sub(r"(\.\d{6})\d+", r"\1", dt_str)
    # Remove trailing fractional zeros if present
    dt_str = re.sub(r"\.0+$", "", dt_str)

    try:
        dt = datetime.fromisoformat(dt_str)
    except ValueError:
        # Fallback: try dateutil
        from dateutil.parser import isoparse
        dt = isoparse(dt_str)

    if dt.tzinfo is None:
        tz = _resolve_timezone(tz_name)
        dt = dt.replace(tzinfo=tz)

    return dt
