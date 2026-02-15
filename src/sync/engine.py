"""Core sync engine: match, diff, resolve, and apply changes."""

import fnmatch
import logging
from datetime import datetime, timedelta, timezone

from src.google_cal.client import (
    GoogleCalendarClient,
    get_outlook_sync_id,
    google_event_to_unified,
    unified_to_google_body,
)
from src.outlook.event_reader import read_outlook_events
from src.outlook.event_writer import create_event, delete_event, set_graph_client, update_event
from src.outlook.graph_client import GraphCalendarClient
from src.sync.conflict import resolve_conflict
from src.sync.matcher import match_by_sync_pairs
from src.sync.models import SyncAction, SyncPair, UnifiedEvent
from src.sync.state import SyncStateStore

logger = logging.getLogger("outlook_gcal_sync.sync.engine")


class SyncEngine:
    """Orchestrates bidirectional calendar sync."""

    def __init__(
        self,
        config: dict,
        google_client: GoogleCalendarClient,
        state_store: SyncStateStore,
        graph_client: GraphCalendarClient | None = None,
    ):
        self.config = config
        self.google = google_client
        self.state = state_store
        self.graph_client = graph_client
        self.sync_config = config["sync"]
        self.outlook_calendar = config["outlook"]["calendar_name"]
        self.dry_run = self.sync_config.get("dry_run", False)
        self.direction = self.sync_config.get("direction", "both")
        self.conflict_strategy = self.sync_config.get("conflict_resolution", "outlook-wins")
        self.exclude_patterns = self.sync_config.get("exclude_patterns", [])

        # Wire up the Graph client for the event writer module
        if graph_client:
            set_graph_client(graph_client, self.outlook_calendar)

    def run(self) -> dict:
        """Execute a full sync cycle. Returns a summary dict."""
        logger.info("=== Starting sync (direction=%s, dry_run=%s) ===", self.direction, self.dry_run)

        # Phase 1: Fetch
        outlook_events = self._fetch_outlook()
        google_events_raw, google_sync_token = self._fetch_google()

        # Convert Google events
        google_events = []
        google_outlook_id_map = {}  # google_id -> outlook_id from extendedProperties
        for raw in google_events_raw:
            if raw.get("status") == "cancelled":
                continue
            g_event = google_event_to_unified(raw)
            google_events.append(g_event)
            outlook_id = get_outlook_sync_id(raw)
            if outlook_id:
                google_outlook_id_map[g_event.source_id] = outlook_id

        # Apply exclusion filters
        outlook_events = self._filter_events(outlook_events)
        google_events = self._filter_events(google_events)

        logger.info("Fetched %d Outlook events, %d Google events.", len(outlook_events), len(google_events))

        # Phase 2: Match
        pairs = self.state.get_all_pairs()
        matched, new_outlook, new_google, orphaned = match_by_sync_pairs(
            outlook_events, google_events, pairs, google_outlook_id_map,
        )

        # Phase 3 & 4: Diff and resolve → build action list
        actions = self._build_actions(matched, new_outlook, new_google, orphaned)

        # Phase 5: Apply
        summary = self._apply_actions(actions)

        # Phase 6: Persist
        if not self.dry_run and google_sync_token:
            self.state.set_metadata("google_sync_token", google_sync_token)
        now = datetime.now(timezone.utc).isoformat()
        if not self.dry_run:
            self.state.set_metadata("last_sync_time", now)

        logger.info(
            "=== Sync complete: %d created, %d updated, %d deleted, %d skipped, %d conflicts ===",
            summary["created"], summary["updated"], summary["deleted"],
            summary["skipped"], summary["conflicts"],
        )
        return summary

    def _fetch_outlook(self) -> list[UnifiedEvent]:
        """Fetch events from Outlook via Microsoft Graph API."""
        return read_outlook_events(
            calendar_name=self.outlook_calendar,
            days_back=self.sync_config["days_back"],
            days_forward=self.sync_config["days_forward"],
            graph_client=self.graph_client,
        )

    def _fetch_google(self) -> tuple[list[dict], str | None]:
        """Fetch events from Google Calendar API."""
        now = datetime.now(timezone.utc)
        time_min = now - timedelta(days=self.sync_config["days_back"])
        time_max = now + timedelta(days=self.sync_config["days_forward"])

        sync_token = self.state.get_metadata("google_sync_token")
        return self.google.list_events(time_min, time_max, sync_token=sync_token)

    def _filter_events(self, events: list[UnifiedEvent]) -> list[UnifiedEvent]:
        """Apply exclusion patterns to filter out unwanted events."""
        if not self.exclude_patterns:
            return events

        filtered = []
        for event in events:
            excluded = False
            for pattern in self.exclude_patterns:
                if fnmatch.fnmatch(event.title, pattern):
                    logger.debug("Excluding event '%s' (matches pattern '%s')", event.title, pattern)
                    excluded = True
                    break
            if not excluded:
                filtered.append(event)
        return filtered

    def _build_actions(
        self,
        matched: list[tuple],
        new_outlook: list[UnifiedEvent],
        new_google: list[UnifiedEvent],
        orphaned: list[SyncPair],
    ) -> list[SyncAction]:
        """Analyze diffs and build list of sync actions."""
        actions = []

        # Handle matched pairs (existing synced events)
        for item in matched:
            o_event, g_event, pair = item

            if o_event and g_event and pair:
                # Both exist — check for changes
                o_hash = o_event.content_hash()
                g_hash = g_event.content_hash()
                last_hash = pair.last_synced_hash

                o_changed = o_hash != last_hash
                g_changed = g_hash != last_hash

                if not o_changed and not g_changed:
                    continue  # No changes

                if o_changed and not g_changed:
                    # Outlook changed → push to Google
                    if self.direction in ("both", "outlook-to-google"):
                        actions.append(SyncAction("update", "o2g", o_event, pair.google_id, pair))

                elif g_changed and not o_changed:
                    # Google changed → push to Outlook
                    if self.direction in ("both", "google-to-outlook"):
                        actions.append(SyncAction("update", "g2o", g_event, pair.outlook_id, pair))

                else:
                    # Both changed — conflict
                    winner = resolve_conflict(o_event, g_event, self.conflict_strategy)
                    if winner == "outlook" and self.direction in ("both", "outlook-to-google"):
                        actions.append(SyncAction("update", "o2g", o_event, pair.google_id, pair))
                    elif winner == "google" and self.direction in ("both", "google-to-outlook"):
                        actions.append(SyncAction("update", "g2o", g_event, pair.outlook_id, pair))

            elif o_event and not g_event and pair:
                # Google event deleted
                if self.direction in ("both", "google-to-outlook"):
                    actions.append(SyncAction("delete", "g2o", None, pair.outlook_id, pair))
                elif self.direction == "outlook-to-google":
                    # Re-create on Google (Outlook is source of truth)
                    actions.append(SyncAction("create", "o2g", o_event, None, pair))

            elif g_event and not o_event and pair:
                # Outlook event deleted
                if self.direction in ("both", "outlook-to-google"):
                    actions.append(SyncAction("delete", "o2g", None, pair.google_id, pair))
                elif self.direction == "google-to-outlook":
                    # Re-create on Outlook (Google is source of truth)
                    actions.append(SyncAction("create", "g2o", g_event, None, pair))

            elif o_event and g_event and not pair:
                # Matched (fuzzy or extended props) but no pair — create pair
                actions.append(SyncAction("update", "o2g", o_event, g_event.source_id, None))

        # Handle new unmatched Outlook events → create on Google
        if self.direction in ("both", "outlook-to-google"):
            for event in new_outlook:
                actions.append(SyncAction("create", "o2g", event, None, None))

        # Handle new unmatched Google events → create on Outlook
        if self.direction in ("both", "google-to-outlook"):
            for event in new_google:
                actions.append(SyncAction("create", "g2o", event, None, None))

        # Clean up orphaned pairs
        for pair in orphaned:
            logger.debug("Cleaning up orphaned pair (outlook=%s, google=%s)", pair.outlook_id, pair.google_id)
            if not self.dry_run:
                self.state.delete_pair(pair.id)

        logger.info("Built %d sync actions.", len(actions))
        return actions

    def _apply_actions(self, actions: list[SyncAction]) -> dict:
        """Execute sync actions. Returns summary counts."""
        summary = {"created": 0, "updated": 0, "deleted": 0, "skipped": 0, "conflicts": 0, "errors": 0}

        for action in actions:
            try:
                if self.dry_run:
                    title = action.source_event.title if action.source_event else "(deleted)"
                    logger.info("[DRY RUN] %s %s: '%s'", action.action, action.direction, title)
                    summary[action.action + "d" if action.action != "delete" else "deleted"] += 1
                    continue

                if action.action == "create" and action.direction == "o2g":
                    self._create_on_google(action)
                    summary["created"] += 1

                elif action.action == "create" and action.direction == "g2o":
                    self._create_on_outlook(action)
                    summary["created"] += 1

                elif action.action == "update" and action.direction == "o2g":
                    self._update_on_google(action)
                    summary["updated"] += 1

                elif action.action == "update" and action.direction == "g2o":
                    self._update_on_outlook(action)
                    summary["updated"] += 1

                elif action.action == "delete" and action.direction == "o2g":
                    self._delete_on_google(action)
                    summary["deleted"] += 1

                elif action.action == "delete" and action.direction == "g2o":
                    self._delete_on_outlook(action)
                    summary["deleted"] += 1

            except Exception as e:
                title = action.source_event.title if action.source_event else "(unknown)"
                logger.error(
                    "Failed to %s %s event '%s': %s",
                    action.action, action.direction, title, e,
                )
                summary["errors"] += 1

        return summary

    def _create_on_google(self, action: SyncAction) -> None:
        """Create a new event on Google Calendar."""
        event = action.source_event
        body = unified_to_google_body(event, outlook_id=event.source_id)
        result = self.google.insert_event(body)
        now = datetime.now(timezone.utc).isoformat()

        pair = SyncPair(
            id=action.pair.id if action.pair else None,
            outlook_id=event.source_id,
            google_id=result["id"],
            last_synced_hash=event.content_hash(),
            last_sync_time=now,
            sync_direction="o2g",
            status="synced",
            outlook_title=event.title,
        )
        self.state.upsert_pair(pair)
        self.state.log_action("create", "o2g", event.source_id, result["id"], event.title)

    def _create_on_outlook(self, action: SyncAction) -> None:
        """Create a new event on Outlook."""
        event = action.source_event
        outlook_id = create_event(self.outlook_calendar, event)
        now = datetime.now(timezone.utc).isoformat()

        # Update Google event with outlookSyncId
        try:
            self.google.patch_event(event.source_id, {
                "extendedProperties": {"private": {"outlookSyncId": outlook_id}},
            })
        except Exception as e:
            logger.warning("Failed to set outlookSyncId on Google event: %s", e)

        pair = SyncPair(
            id=action.pair.id if action.pair else None,
            outlook_id=outlook_id,
            google_id=event.source_id,
            last_synced_hash=event.content_hash(),
            last_sync_time=now,
            sync_direction="g2o",
            status="synced",
            outlook_title=event.title,
        )
        self.state.upsert_pair(pair)
        self.state.log_action("create", "g2o", outlook_id, event.source_id, event.title)

    def _update_on_google(self, action: SyncAction) -> None:
        """Update an existing Google event from Outlook data."""
        event = action.source_event
        body = unified_to_google_body(event, outlook_id=event.source_id)
        self.google.update_event(action.target_id, body)
        now = datetime.now(timezone.utc).isoformat()

        if action.pair:
            action.pair.last_synced_hash = event.content_hash()
            action.pair.last_sync_time = now
            action.pair.status = "synced"
            action.pair.outlook_title = event.title
            self.state.upsert_pair(action.pair)
        else:
            pair = SyncPair(
                id=None,
                outlook_id=event.source_id,
                google_id=action.target_id,
                last_synced_hash=event.content_hash(),
                last_sync_time=now,
                sync_direction="o2g",
                status="synced",
                outlook_title=event.title,
            )
            self.state.upsert_pair(pair)

        self.state.log_action("update", "o2g", event.source_id, action.target_id, event.title)

    def _update_on_outlook(self, action: SyncAction) -> None:
        """Update an existing Outlook event from Google data."""
        event = action.source_event
        update_event(self.outlook_calendar, action.target_id, event)
        now = datetime.now(timezone.utc).isoformat()

        if action.pair:
            action.pair.last_synced_hash = event.content_hash()
            action.pair.last_sync_time = now
            action.pair.status = "synced"
            action.pair.outlook_title = event.title
            self.state.upsert_pair(action.pair)

        self.state.log_action("update", "g2o", action.target_id, event.source_id, event.title)

    def _delete_on_google(self, action: SyncAction) -> None:
        """Delete an event from Google Calendar."""
        self.google.delete_event(action.target_id)
        if action.pair:
            self.state.delete_pair(action.pair.id)
            self.state.log_action("delete", "o2g", action.pair.outlook_id, action.target_id, action.pair.outlook_title)

    def _delete_on_outlook(self, action: SyncAction) -> None:
        """Delete an event from Outlook."""
        delete_event(self.outlook_calendar, action.target_id)
        if action.pair:
            self.state.delete_pair(action.pair.id)
            self.state.log_action("delete", "g2o", action.target_id, action.pair.google_id, action.pair.outlook_title)
