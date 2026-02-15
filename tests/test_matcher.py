"""Tests for event matching logic."""

from datetime import datetime
from zoneinfo import ZoneInfo

from src.sync.matcher import match_by_sync_pairs, _fuzzy_match
from src.sync.models import SyncPair, UnifiedEvent

ET = ZoneInfo("America/New_York")


def _make_event(source, source_id, title, start_hour=9, start_min=0):
    return UnifiedEvent(
        source=source,
        source_id=source_id,
        title=title,
        description="",
        location="",
        start_time=datetime(2026, 3, 15, start_hour, start_min, 0, tzinfo=ET),
        end_time=datetime(2026, 3, 15, start_hour + 1, start_min, 0, tzinfo=ET),
        is_all_day=False,
        has_reminder=False,
        reminder_minutes=0,
        is_recurring=False,
    )


class TestFuzzyMatch:
    def test_exact_match(self):
        o = _make_event("outlook", "1", "Standup")
        g = _make_event("google", "g1", "Standup")
        assert _fuzzy_match(o, g) is True

    def test_case_insensitive(self):
        o = _make_event("outlook", "1", "Team STANDUP")
        g = _make_event("google", "g1", "team standup")
        assert _fuzzy_match(o, g) is True

    def test_different_title(self):
        o = _make_event("outlook", "1", "Standup")
        g = _make_event("google", "g1", "Lunch")
        assert _fuzzy_match(o, g) is False

    def test_time_within_threshold(self):
        o = _make_event("outlook", "1", "Standup", start_min=0)
        g = _make_event("google", "g1", "Standup", start_min=1)
        assert _fuzzy_match(o, g) is True

    def test_time_outside_threshold(self):
        o = _make_event("outlook", "1", "Standup", start_hour=9)
        g = _make_event("google", "g1", "Standup", start_hour=14)
        assert _fuzzy_match(o, g) is False

    def test_empty_title(self):
        o = _make_event("outlook", "1", "")
        g = _make_event("google", "g1", "")
        assert _fuzzy_match(o, g) is False


class TestMatchBySyncPairs:
    def test_match_via_existing_pair(self):
        o = _make_event("outlook", "1", "Standup")
        g = _make_event("google", "g1", "Standup")
        pair = SyncPair(
            id=1, outlook_id="1", google_id="g1",
            last_synced_hash="h", last_sync_time="t",
            sync_direction="both", status="synced",
        )

        matched, new_o, new_g, orphaned = match_by_sync_pairs(
            [o], [g], [pair], {}
        )
        assert len(matched) == 1
        assert len(new_o) == 0
        assert len(new_g) == 0

    def test_new_outlook_event(self):
        o = _make_event("outlook", "1", "New Meeting")
        matched, new_o, new_g, orphaned = match_by_sync_pairs(
            [o], [], [], {}
        )
        assert len(matched) == 0
        assert len(new_o) == 1

    def test_new_google_event(self):
        g = _make_event("google", "g1", "New Meeting")
        matched, new_o, new_g, orphaned = match_by_sync_pairs(
            [], [g], [], {}
        )
        assert len(matched) == 0
        assert len(new_g) == 1

    def test_orphaned_pair(self):
        pair = SyncPair(
            id=1, outlook_id="1", google_id="g1",
            last_synced_hash="h", last_sync_time="t",
            sync_direction="both", status="synced",
        )
        matched, new_o, new_g, orphaned = match_by_sync_pairs(
            [], [], [pair], {}
        )
        assert len(orphaned) == 1

    def test_match_via_extended_properties(self):
        o = _make_event("outlook", "1", "Standup")
        g = _make_event("google", "g1", "Standup")
        # No existing pair, but Google event has outlookSyncId
        matched, new_o, new_g, orphaned = match_by_sync_pairs(
            [o], [g], [], {"g1": "1"}
        )
        assert len(matched) == 1
        assert len(new_o) == 0
        assert len(new_g) == 0

    def test_fuzzy_match_fallback(self):
        o = _make_event("outlook", "1", "Standup")
        g = _make_event("google", "g1", "Standup")
        # No pair, no extended properties — should fuzzy match
        matched, new_o, new_g, orphaned = match_by_sync_pairs(
            [o], [g], [], {}
        )
        assert len(matched) == 1
        assert len(new_o) == 0
        assert len(new_g) == 0

    def test_outlook_deleted(self):
        g = _make_event("google", "g1", "Standup")
        pair = SyncPair(
            id=1, outlook_id="1", google_id="g1",
            last_synced_hash="h", last_sync_time="t",
            sync_direction="both", status="synced",
        )
        matched, new_o, new_g, orphaned = match_by_sync_pairs(
            [], [g], [pair], {}
        )
        # Should appear as matched with outlook=None
        assert len(matched) == 1
        assert matched[0][0] is None  # outlook event
        assert matched[0][1] is not None  # google event
