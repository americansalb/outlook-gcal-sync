"""Event matching logic for correlating Outlook and Google events."""

import logging
from datetime import timedelta

from src.sync.models import SyncPair, UnifiedEvent

logger = logging.getLogger("outlook_gcal_sync.sync.matcher")


def match_by_sync_pairs(
    outlook_events: list[UnifiedEvent],
    google_events: list[UnifiedEvent],
    pairs: list[SyncPair],
    google_outlook_id_map: dict[str, str],
) -> tuple[
    list[tuple[UnifiedEvent, UnifiedEvent, SyncPair]],  # matched
    list[UnifiedEvent],  # unmatched outlook
    list[UnifiedEvent],  # unmatched google
    list[SyncPair],      # orphaned pairs (both sides deleted)
]:
    """Match events using sync state and extended properties.

    Args:
        outlook_events: Events from Outlook.
        google_events: Events from Google.
        pairs: Existing sync pairs from the state DB.
        google_outlook_id_map: Map of google_event_id -> outlookSyncId
            from Google extendedProperties.

    Returns:
        Tuple of (matched_triples, unmatched_outlook, unmatched_google, orphaned_pairs).
    """
    # Build lookup indices
    outlook_by_id = {e.source_id: e for e in outlook_events}
    google_by_id = {e.source_id: e for e in google_events}
    pairs_by_outlook = {p.outlook_id: p for p in pairs if p.outlook_id}
    pairs_by_google = {p.google_id: p for p in pairs if p.google_id}

    matched = []
    matched_outlook_ids = set()
    matched_google_ids = set()
    orphaned = []

    # Phase 1: Match via existing sync pairs
    for pair in pairs:
        o_event = outlook_by_id.get(pair.outlook_id) if pair.outlook_id else None
        g_event = google_by_id.get(pair.google_id) if pair.google_id else None

        if o_event and g_event:
            matched.append((o_event, g_event, pair))
            matched_outlook_ids.add(o_event.source_id)
            matched_google_ids.add(g_event.source_id)
        elif o_event and not g_event:
            # Google event deleted — will be handled by engine as deletion
            matched.append((o_event, None, pair))  # type: ignore[arg-type]
            matched_outlook_ids.add(o_event.source_id)
        elif g_event and not o_event:
            # Outlook event deleted — will be handled by engine as deletion
            matched.append((None, g_event, pair))  # type: ignore[arg-type]
            matched_google_ids.add(g_event.source_id)
        else:
            # Both deleted
            orphaned.append(pair)

    # Phase 2: Match unmatched Google events by extendedProperties.outlookSyncId
    for g_event in google_events:
        if g_event.source_id in matched_google_ids:
            continue
        outlook_id = google_outlook_id_map.get(g_event.source_id)
        if outlook_id and outlook_id in outlook_by_id:
            o_event = outlook_by_id[outlook_id]
            if o_event.source_id not in matched_outlook_ids:
                # Found a match via extended properties but no pair exists yet
                matched.append((o_event, g_event, None))  # type: ignore[arg-type]
                matched_outlook_ids.add(o_event.source_id)
                matched_google_ids.add(g_event.source_id)

    # Phase 3: Fuzzy match remaining unmatched events
    unmatched_outlook = [e for e in outlook_events if e.source_id not in matched_outlook_ids]
    unmatched_google = [e for e in google_events if e.source_id not in matched_google_ids]

    fuzzy_matched_o = set()
    fuzzy_matched_g = set()

    for o_event in unmatched_outlook:
        for g_event in unmatched_google:
            if g_event.source_id in fuzzy_matched_g:
                continue
            if _fuzzy_match(o_event, g_event):
                matched.append((o_event, g_event, None))  # type: ignore[arg-type]
                fuzzy_matched_o.add(o_event.source_id)
                fuzzy_matched_g.add(g_event.source_id)
                logger.info(
                    "Fuzzy matched: Outlook '%s' <-> Google '%s'",
                    o_event.title, g_event.title,
                )
                break

    final_unmatched_outlook = [e for e in unmatched_outlook if e.source_id not in fuzzy_matched_o]
    final_unmatched_google = [e for e in unmatched_google if e.source_id not in fuzzy_matched_g]

    logger.info(
        "Matching: %d matched, %d new Outlook, %d new Google, %d orphaned pairs",
        len(matched), len(final_unmatched_outlook),
        len(final_unmatched_google), len(orphaned),
    )

    return matched, final_unmatched_outlook, final_unmatched_google, orphaned


def _fuzzy_match(o: UnifiedEvent, g: UnifiedEvent) -> bool:
    """Check if two events are likely the same using heuristics.

    Matches if title is identical (case-insensitive) and start times
    are within 2 minutes of each other.
    """
    if not o.title or not g.title:
        return False

    if o.title.strip().lower() != g.title.strip().lower():
        return False

    # Compare start times
    time_diff = abs((o.start_time - g.start_time).total_seconds())
    return time_diff <= 120  # Within 2 minutes
