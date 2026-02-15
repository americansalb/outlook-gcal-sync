"""Conflict resolution strategies for bidirectional sync."""

import logging

from src.sync.models import UnifiedEvent

logger = logging.getLogger("outlook_gcal_sync.sync.conflict")


def resolve_conflict(
    outlook_event: UnifiedEvent,
    google_event: UnifiedEvent,
    strategy: str,
) -> str:
    """Determine which side wins when both have changed.

    Args:
        outlook_event: The current Outlook version.
        google_event: The current Google version.
        strategy: One of "outlook-wins", "google-wins", "newest".

    Returns:
        "outlook" or "google" — the winner.
    """
    if strategy == "outlook-wins":
        logger.debug("Conflict on '%s': Outlook wins (by policy).", outlook_event.title)
        return "outlook"

    if strategy == "google-wins":
        logger.debug("Conflict on '%s': Google wins (by policy).", google_event.title)
        return "google"

    if strategy == "newest":
        # AppleScript doesn't reliably expose modification timestamps,
        # so we fall back to outlook-wins if we can't compare.
        logger.debug(
            "Conflict on '%s': 'newest' strategy requested but modification "
            "timestamps unavailable. Falling back to outlook-wins.",
            outlook_event.title,
        )
        return "outlook"

    logger.warning("Unknown conflict strategy '%s'. Defaulting to outlook-wins.", strategy)
    return "outlook"
