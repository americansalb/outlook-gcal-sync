"""Tests for sync state persistence."""

import pytest
from src.sync.state import SyncStateStore
from src.sync.models import SyncPair


@pytest.fixture
def state_store(tmp_path):
    db_path = str(tmp_path / "test_state.db")
    store = SyncStateStore(db_path)
    yield store
    store.close()


class TestSyncStateStore:
    def test_upsert_and_get_by_outlook_id(self, state_store):
        pair = SyncPair(
            id=None, outlook_id="o123", google_id="g456",
            last_synced_hash="hash1", last_sync_time="2026-03-15T00:00:00Z",
            sync_direction="both", status="synced", outlook_title="Test",
        )
        row_id = state_store.upsert_pair(pair)
        assert row_id is not None

        result = state_store.get_pair_by_outlook_id("o123")
        assert result is not None
        assert result.google_id == "g456"
        assert result.outlook_title == "Test"

    def test_get_by_google_id(self, state_store):
        pair = SyncPair(
            id=None, outlook_id="o123", google_id="g456",
            last_synced_hash="hash1", last_sync_time="2026-03-15T00:00:00Z",
            sync_direction="both", status="synced",
        )
        state_store.upsert_pair(pair)

        result = state_store.get_pair_by_google_id("g456")
        assert result is not None
        assert result.outlook_id == "o123"

    def test_get_nonexistent(self, state_store):
        assert state_store.get_pair_by_outlook_id("nonexistent") is None
        assert state_store.get_pair_by_google_id("nonexistent") is None

    def test_update_existing(self, state_store):
        pair = SyncPair(
            id=None, outlook_id="o123", google_id="g456",
            last_synced_hash="hash1", last_sync_time="2026-03-15T00:00:00Z",
            sync_direction="both", status="synced", outlook_title="Original",
        )
        row_id = state_store.upsert_pair(pair)

        pair.id = row_id
        pair.last_synced_hash = "hash2"
        pair.outlook_title = "Updated"
        state_store.upsert_pair(pair)

        result = state_store.get_pair_by_outlook_id("o123")
        assert result.last_synced_hash == "hash2"
        assert result.outlook_title == "Updated"

    def test_delete_pair(self, state_store):
        pair = SyncPair(
            id=None, outlook_id="o123", google_id="g456",
            last_synced_hash="hash1", last_sync_time="2026-03-15T00:00:00Z",
            sync_direction="both", status="synced",
        )
        row_id = state_store.upsert_pair(pair)
        state_store.delete_pair(row_id)
        assert state_store.get_pair_by_outlook_id("o123") is None

    def test_get_all_pairs(self, state_store):
        for i in range(3):
            pair = SyncPair(
                id=None, outlook_id=f"o{i}", google_id=f"g{i}",
                last_synced_hash=f"hash{i}", last_sync_time="2026-03-15T00:00:00Z",
                sync_direction="both", status="synced",
            )
            state_store.upsert_pair(pair)
        assert len(state_store.get_all_pairs()) == 3

    def test_metadata(self, state_store):
        state_store.set_metadata("test_key", "test_value")
        assert state_store.get_metadata("test_key") == "test_value"
        assert state_store.get_metadata("nonexistent") is None

    def test_log_action(self, state_store):
        state_store.log_action("create", "o2g", "o123", "g456", "Test Event")
        stats = state_store.get_stats()
        assert stats["recent_actions"].get("create", 0) >= 1

    def test_reset(self, state_store):
        pair = SyncPair(
            id=None, outlook_id="o123", google_id="g456",
            last_synced_hash="hash1", last_sync_time="2026-03-15T00:00:00Z",
            sync_direction="both", status="synced",
        )
        state_store.upsert_pair(pair)
        state_store.set_metadata("google_sync_token", "token123")
        state_store.reset()

        assert len(state_store.get_all_pairs()) == 0
        assert state_store.get_metadata("google_sync_token") is None
