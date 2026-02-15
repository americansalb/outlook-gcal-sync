"""Shared test fixtures."""

import pytest
from datetime import datetime
from zoneinfo import ZoneInfo

from src.sync.models import UnifiedEvent, SyncPair


ET = ZoneInfo("America/New_York")


@pytest.fixture
def sample_outlook_event():
    return UnifiedEvent(
        source="outlook",
        source_id="1001",
        title="Team Standup",
        description="Daily sync meeting",
        location="Conference Room A",
        start_time=datetime(2026, 3, 15, 9, 0, 0, tzinfo=ET),
        end_time=datetime(2026, 3, 15, 9, 30, 0, tzinfo=ET),
        is_all_day=False,
        has_reminder=True,
        reminder_minutes=15,
        is_recurring=False,
    )


@pytest.fixture
def sample_google_event():
    return UnifiedEvent(
        source="google",
        source_id="g_abc123",
        title="Team Standup",
        description="Daily sync meeting",
        location="Conference Room A",
        start_time=datetime(2026, 3, 15, 9, 0, 0, tzinfo=ET),
        end_time=datetime(2026, 3, 15, 9, 30, 0, tzinfo=ET),
        is_all_day=False,
        has_reminder=True,
        reminder_minutes=15,
        is_recurring=False,
    )


@pytest.fixture
def sample_all_day_event():
    from datetime import date
    return UnifiedEvent(
        source="outlook",
        source_id="1002",
        title="Company Holiday",
        description="Office closed",
        location="",
        start_time=datetime(2026, 3, 20, 0, 0, 0, tzinfo=ET),
        end_time=datetime(2026, 3, 21, 0, 0, 0, tzinfo=ET),
        is_all_day=True,
        has_reminder=False,
        reminder_minutes=0,
        is_recurring=False,
        start_date=date(2026, 3, 20),
        end_date=date(2026, 3, 21),
    )


@pytest.fixture
def sample_sync_pair():
    return SyncPair(
        id=1,
        outlook_id="1001",
        google_id="g_abc123",
        last_synced_hash="abc123hash",
        last_sync_time="2026-03-14T12:00:00+00:00",
        sync_direction="both",
        status="synced",
        outlook_title="Team Standup",
    )
