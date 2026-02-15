"""Create, update, and delete calendar events in Outlook via Microsoft Graph API."""

import logging

from src.outlook.graph_client import GraphCalendarClient, unified_to_graph_body
from src.sync.models import UnifiedEvent

logger = logging.getLogger("outlook_gcal_sync.outlook.writer")

# Module-level client, set by the sync engine during initialization.
_graph_client: GraphCalendarClient | None = None
_calendar_id: str | None = None


def set_graph_client(client: GraphCalendarClient, calendar_name: str) -> None:
    """Wire up the Graph client and resolve the calendar ID.

    Called by SyncEngine.__init__ before any sync operations.
    """
    global _graph_client, _calendar_id
    _graph_client = client
    _calendar_id = client.resolve_calendar_id(calendar_name)
    logger.debug("Event writer initialized (calendar_id=%s)", _calendar_id)


def _ensure_client() -> tuple[GraphCalendarClient, str]:
    """Return the client and calendar_id, or raise if not initialized."""
    if _graph_client is None or _calendar_id is None:
        raise RuntimeError(
            "Microsoft Graph client not initialized. Run 'outlook-gcal-sync setup-microsoft' first."
        )
    return _graph_client, _calendar_id


def create_event(calendar_name: str, event: UnifiedEvent) -> str:
    """Create a new event in Outlook and return its ID.

    Args:
        calendar_name: Name of the Outlook calendar (unused — calendar_id is pre-resolved).
        event: The event to create.

    Returns:
        The Outlook (Graph) event ID as a string.
    """
    client, cal_id = _ensure_client()
    body = unified_to_graph_body(event)
    result = client.create_event(cal_id, body)
    event_id = result["id"]
    logger.info("Created Outlook event '%s' (ID: %s)", event.title, event_id)
    return event_id


def update_event(calendar_name: str, outlook_id: str, event: UnifiedEvent) -> None:
    """Update an existing Outlook event.

    Args:
        calendar_name: Name of the Outlook calendar (unused — uses event ID directly).
        outlook_id: The Graph event ID to update.
        event: The event data to apply.
    """
    client, _ = _ensure_client()
    body = unified_to_graph_body(event)
    client.update_event(outlook_id, body)
    logger.info("Updated Outlook event '%s' (ID: %s)", event.title, outlook_id)


def delete_event(calendar_name: str, outlook_id: str) -> None:
    """Delete an event from Outlook.

    Args:
        calendar_name: Name of the Outlook calendar (unused — uses event ID directly).
        outlook_id: The Graph event ID to delete.
    """
    client, _ = _ensure_client()
    client.delete_event(outlook_id)
    logger.info("Deleted Outlook event (ID: %s)", outlook_id)
