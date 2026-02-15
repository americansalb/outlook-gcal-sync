"""Tests for content hashing."""

from src.utils.hashing import content_hash


def test_same_input_same_hash():
    h1 = content_hash("Meeting", "Notes", "Room A", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, True, 15)
    h2 = content_hash("Meeting", "Notes", "Room A", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, True, 15)
    assert h1 == h2


def test_different_title_different_hash():
    h1 = content_hash("Meeting", "", "", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, False, 0)
    h2 = content_hash("Lunch", "", "", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, False, 0)
    assert h1 != h2


def test_whitespace_normalized():
    h1 = content_hash("Meeting", "Notes", "Room A", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, False, 0)
    h2 = content_hash("Meeting ", " Notes ", " Room A", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, False, 0)
    assert h1 == h2


def test_reminder_ignored_when_disabled():
    h1 = content_hash("Meeting", "", "", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, False, 0)
    h2 = content_hash("Meeting", "", "", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, False, 15)
    assert h1 == h2  # has_reminder=False, so reminder_minutes doesn't matter


def test_reminder_matters_when_enabled():
    h1 = content_hash("Meeting", "", "", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, True, 10)
    h2 = content_hash("Meeting", "", "", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, True, 15)
    assert h1 != h2


def test_hash_is_sha256():
    h = content_hash("Test", "", "", "2026-03-15T09:00:00", "2026-03-15T10:00:00", False, False, 0)
    assert len(h) == 64  # SHA-256 hex digest length
    assert all(c in "0123456789abcdef" for c in h)
