"""Read calendar events from Outlook via Microsoft Graph API."""

from __future__ import annotations

import logging
from datetime import date, datetime, timedelta, timezone

from src.outlook.graph_client import GraphCalendarClient, graph_event_to_unified
from src.sync.models import UnifiedEvent

logger = logging.getLogger("outlook_gcal_sync.outlook.reader")


def read_outlook_events(
    calendar_name: str,
    days_back: int = 30,
    days_forward: int = 90,
    *,
    graph_client: GraphCalendarClient | None = None,
) -> list[UnifiedEvent]:
    """Read events from an Outlook calendar via Microsoft Graph API.

    Args:
        calendar_name: Name of the Outlook calendar (used to resolve Graph calendar ID).
        days_back: Number of days in the past to include.
        days_forward: Number of days in the future to include.
        graph_client: An authenticated GraphCalendarClient instance.

    Returns:
        List of UnifiedEvent objects.

    Raises:
        RuntimeError: If graph_client is not provided.
    """
    if graph_client is None:
        raise RuntimeError(
            "Microsoft Graph client not initialized. Run 'outlook-gcal-sync setup-microsoft' first."
        )

    today = date.today()
    start = datetime(today.year, today.month, today.day, tzinfo=timezone.utc) - timedelta(days=days_back)
    end = datetime(today.year, today.month, today.day, 23, 59, 59, tzinfo=timezone.utc) + timedelta(days=days_forward)

    logger.info("Reading Outlook events from '%s' (%s to %s)...", calendar_name, start.date(), end.date())

    # Resolve calendar name to Graph ID
    calendar_id = graph_client.resolve_calendar_id(calendar_name)

    # Fetch events
    raw_events = graph_client.list_events(calendar_id, start, end)

    # Convert to UnifiedEvent
    events = []
    for raw in raw_events:
        try:
            event = graph_event_to_unified(raw)
            events.append(event)
        except Exception as e:
            subject = raw.get("subject", "(unknown)")
            logger.warning("Failed to parse Outlook event '%s': %s", subject, e)

    logger.info("Read %d events from Outlook.", len(events))
    return events
