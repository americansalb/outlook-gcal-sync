"""Data models for sync operations."""

from dataclasses import dataclass, field
from datetime import date, datetime


@dataclass
class UnifiedEvent:
    """A calendar event normalized from either Outlook or Google."""

    source: str  # "outlook" or "google"
    source_id: str  # Original ID in source system
    title: str
    description: str
    location: str
    start_time: datetime  # Timezone-aware for timed events
    end_time: datetime  # Timezone-aware for timed events
    is_all_day: bool
    has_reminder: bool
    reminder_minutes: int  # 0 if no reminder
    is_recurring: bool
    start_date: date | None = None  # For all-day events
    end_date: date | None = None  # For all-day events (exclusive end)

    def content_hash(self) -> str:
        """Compute a content hash of sync-relevant fields."""
        from src.utils.hashing import content_hash

        if self.is_all_day and self.start_date:
            start_iso = self.start_date.isoformat()
            end_iso = self.end_date.isoformat() if self.end_date else start_iso
        else:
            start_iso = self.start_time.isoformat()
            end_iso = self.end_time.isoformat()

        return content_hash(
            title=self.title,
            description=self.description,
            location=self.location,
            start_iso=start_iso,
            end_iso=end_iso,
            is_all_day=self.is_all_day,
            has_reminder=self.has_reminder,
            reminder_minutes=self.reminder_minutes,
        )


@dataclass
class SyncPair:
    """Tracks the mapping between an Outlook event and a Google event."""

    id: int | None  # SQLite rowid, None for new pairs
    outlook_id: str | None
    google_id: str | None
    last_synced_hash: str
    last_sync_time: str  # ISO 8601
    sync_direction: str  # "o2g" | "g2o" | "both"
    status: str = "synced"  # synced | pending | conflict | error
    outlook_title: str = ""  # For human-readable debugging


@dataclass
class SyncAction:
    """An action to take during sync."""

    action: str  # "create" | "update" | "delete"
    direction: str  # "o2g" (outlook-to-google) | "g2o" (google-to-outlook)
    source_event: UnifiedEvent | None  # The event to sync from
    target_id: str | None  # The target event ID to update/delete (None for create)
    pair: SyncPair | None  # Existing sync pair, if any
