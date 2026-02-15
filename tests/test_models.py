"""Tests for data models."""


def test_content_hash_consistency(sample_outlook_event):
    h1 = sample_outlook_event.content_hash()
    h2 = sample_outlook_event.content_hash()
    assert h1 == h2


def test_content_hash_differs_on_change(sample_outlook_event):
    h1 = sample_outlook_event.content_hash()
    sample_outlook_event.title = "Changed Title"
    h2 = sample_outlook_event.content_hash()
    assert h1 != h2


def test_matching_events_same_hash(sample_outlook_event, sample_google_event):
    # Two events with identical content should have the same hash
    assert sample_outlook_event.content_hash() == sample_google_event.content_hash()


def test_all_day_event_hash(sample_all_day_event):
    h = sample_all_day_event.content_hash()
    assert len(h) == 64  # SHA-256
