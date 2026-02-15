"""Content hashing for change detection."""

import hashlib
import json


def content_hash(
    title: str,
    description: str,
    location: str,
    start_iso: str,
    end_iso: str,
    is_all_day: bool,
    has_reminder: bool,
    reminder_minutes: int,
) -> str:
    """Compute a SHA-256 hash of normalized event fields.

    Used to detect whether an event has changed since last sync.
    Normalization ensures semantically identical events produce the same hash.
    """
    normalized = {
        "title": (title or "").strip(),
        "description": (description or "").strip(),
        "location": (location or "").strip(),
        "start": start_iso,
        "end": end_iso,
        "is_all_day": is_all_day,
        "has_reminder": has_reminder,
        "reminder_minutes": reminder_minutes if has_reminder else 0,
    }
    canonical = json.dumps(normalized, sort_keys=True)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()
