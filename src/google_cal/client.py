"""Google Calendar API wrapper with incremental sync support."""

import logging
from datetime import date, datetime, timedelta

from googleapiclient.errors import HttpError

from src.sync.models import UnifiedEvent
from src.utils.datetime_utils import (
    from_google_date,
    from_google_datetime,
    to_google_date,
    to_google_datetime,
)

logger = logging.getLogger("outlook_gcal_sync.google.client")


class GoogleCalendarClient:
    """Wrapper around Google Calendar API v3 events resource."""

    def __init__(self, service, calendar_id: str = "primary"):
        self.service = service
        self.calendar_id = calendar_id

    def list_events(
        self,
        time_min: datetime,
        time_max: datetime,
        sync_token: str | None = None,
    ) -> tuple[list[dict], str | None]:
        """List events in the date range, with optional incremental sync.

        Args:
            time_min: Start of date range (inclusive).
            time_max: End of date range (inclusive).
            sync_token: If provided, performs incremental sync.

        Returns:
            Tuple of (events_list, next_sync_token).
            If sync_token was invalid (410), performs full sync.
        """
        all_events = []
        page_token = None
        next_sync_token = None

        try:
            while True:
                kwargs = {
                    "calendarId": self.calendar_id,
                    "singleEvents": True,
                    "orderBy": "startTime",
                    "maxResults": 250,
                }

                if sync_token and not page_token:
                    # Incremental sync — don't pass time bounds
                    kwargs["syncToken"] = sync_token
                else:
                    kwargs["timeMin"] = to_google_datetime(time_min)
                    kwargs["timeMax"] = to_google_datetime(time_max)

                if page_token:
                    kwargs["pageToken"] = page_token

                result = self.service.events().list(**kwargs).execute()
                all_events.extend(result.get("items", []))
                page_token = result.get("nextPageToken")
                if not page_token:
                    next_sync_token = result.get("nextSyncToken")
                    break

        except HttpError as e:
            if e.resp.status == 410:
                # Sync token expired — fall back to full sync
                logger.warning("Google sync token expired (410). Performing full sync.")
                return self.list_events(time_min, time_max, sync_token=None)
            raise

        logger.info(
            "Listed %d events from Google Calendar (incremental=%s).",
            len(all_events), sync_token is not None,
        )
        return all_events, next_sync_token

    def get_event(self, event_id: str) -> dict:
        """Get a single event by ID."""
        return (
            self.service.events()
            .get(calendarId=self.calendar_id, eventId=event_id)
            .execute()
        )

    def insert_event(self, event_body: dict) -> dict:
        """Create a new event. Returns the created event resource."""
        result = (
            self.service.events()
            .insert(calendarId=self.calendar_id, body=event_body)
            .execute()
        )
        logger.info(
            "Created Google event '%s' (ID: %s)",
            result.get("summary", ""), result.get("id", ""),
        )
        return result

    def update_event(self, event_id: str, event_body: dict) -> dict:
        """Update an existing event (full replacement)."""
        result = (
            self.service.events()
            .update(calendarId=self.calendar_id, eventId=event_id, body=event_body)
            .execute()
        )
        logger.info(
            "Updated Google event '%s' (ID: %s)",
            result.get("summary", ""), event_id,
        )
        return result

    def patch_event(self, event_id: str, patch_body: dict) -> dict:
        """Partial update of specific fields."""
        result = (
            self.service.events()
            .patch(calendarId=self.calendar_id, eventId=event_id, body=patch_body)
            .execute()
        )
        logger.info("Patched Google event (ID: %s)", event_id)
        return result

    def delete_event(self, event_id: str) -> None:
        """Delete an event (moves to trash, recoverable for 30 days)."""
        self.service.events().delete(
            calendarId=self.calendar_id, eventId=event_id,
        ).execute()
        logger.info("Deleted Google event (ID: %s)", event_id)


def unified_to_google_body(
    event: UnifiedEvent,
    outlook_id: str | None = None,
) -> dict:
    """Convert a UnifiedEvent to a Google Calendar API event body.

    Args:
        event: The event to convert.
        outlook_id: If set, stored in extendedProperties for cross-system matching.
    """
    body: dict = {
        "summary": event.title,
        "description": event.description,
        "location": event.location,
    }

    if event.is_all_day and event.start_date:
        body["start"] = {"date": to_google_date(event.start_date)}
        # Google uses exclusive end date for all-day events
        end_date = event.end_date or (event.start_date + timedelta(days=1))
        if end_date == event.start_date:
            end_date = event.start_date + timedelta(days=1)
        body["end"] = {"date": to_google_date(end_date)}
    else:
        tz_name = str(event.start_time.tzinfo) if event.start_time.tzinfo else None
        body["start"] = {"dateTime": to_google_datetime(event.start_time)}
        body["end"] = {"dateTime": to_google_datetime(event.end_time)}
        if tz_name:
            body["start"]["timeZone"] = tz_name
            body["end"]["timeZone"] = tz_name

    # Reminders
    if event.has_reminder and event.reminder_minutes > 0:
        body["reminders"] = {
            "useDefault": False,
            "overrides": [{"method": "popup", "minutes": event.reminder_minutes}],
        }
    else:
        body["reminders"] = {"useDefault": True}

    # Extended properties for sync tracking
    if outlook_id:
        body["extendedProperties"] = {
            "private": {"outlookSyncId": str(outlook_id)},
        }

    return body


def google_event_to_unified(event: dict) -> UnifiedEvent:
    """Convert a Google Calendar API event dict to a UnifiedEvent."""
    start = event.get("start", {})
    end = event.get("end", {})

    is_all_day = "date" in start

    if is_all_day:
        start_date = from_google_date(start["date"])
        end_date = from_google_date(end["date"])
        # Build datetime at midnight for the unified model
        from src.utils.datetime_utils import get_local_timezone
        tz = get_local_timezone()
        start_time = datetime(start_date.year, start_date.month, start_date.day, tzinfo=tz)
        end_time = datetime(end_date.year, end_date.month, end_date.day, tzinfo=tz)
    else:
        tz_name = start.get("timeZone")
        start_time = from_google_datetime(start["dateTime"], tz_name)
        end_time = from_google_datetime(end["dateTime"], end.get("timeZone", tz_name))
        start_date = None
        end_date = None

    # Reminders
    reminders = event.get("reminders", {})
    has_reminder = False
    reminder_minutes = 0
    if not reminders.get("useDefault", True):
        overrides = reminders.get("overrides", [])
        if overrides:
            has_reminder = True
            reminder_minutes = overrides[0].get("minutes", 0)

    return UnifiedEvent(
        source="google",
        source_id=event["id"],
        title=event.get("summary", ""),
        description=event.get("description", ""),
        location=event.get("location", ""),
        start_time=start_time,
        end_time=end_time,
        is_all_day=is_all_day,
        has_reminder=has_reminder,
        reminder_minutes=reminder_minutes,
        is_recurring="recurringEventId" in event,
        start_date=start_date,
        end_date=end_date,
    )


def get_outlook_sync_id(event: dict) -> str | None:
    """Extract the outlookSyncId from a Google event's extended properties."""
    props = event.get("extendedProperties", {}).get("private", {})
    return props.get("outlookSyncId")
