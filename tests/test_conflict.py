"""Tests for conflict resolution."""

from src.sync.conflict import resolve_conflict


def test_outlook_wins(sample_outlook_event, sample_google_event):
    assert resolve_conflict(sample_outlook_event, sample_google_event, "outlook-wins") == "outlook"


def test_google_wins(sample_outlook_event, sample_google_event):
    assert resolve_conflict(sample_outlook_event, sample_google_event, "google-wins") == "google"


def test_newest_falls_back_to_outlook(sample_outlook_event, sample_google_event):
    # Since AppleScript doesn't expose modification timestamps,
    # "newest" should fall back to "outlook"
    assert resolve_conflict(sample_outlook_event, sample_google_event, "newest") == "outlook"


def test_unknown_strategy(sample_outlook_event, sample_google_event):
    assert resolve_conflict(sample_outlook_event, sample_google_event, "invalid") == "outlook"
