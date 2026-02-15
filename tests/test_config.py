"""Tests for configuration loading."""

import os
import pytest
from src.config import load_config, _deep_merge


class TestDeepMerge:
    def test_simple_merge(self):
        base = {"a": 1, "b": 2}
        override = {"b": 3, "c": 4}
        result = _deep_merge(base, override)
        assert result == {"a": 1, "b": 3, "c": 4}

    def test_nested_merge(self):
        base = {"sync": {"direction": "both", "days_back": 30}}
        override = {"sync": {"direction": "outlook-to-google"}}
        result = _deep_merge(base, override)
        assert result["sync"]["direction"] == "outlook-to-google"
        assert result["sync"]["days_back"] == 30


class TestLoadConfig:
    def test_defaults_when_no_file(self, tmp_path):
        config = load_config(str(tmp_path / "nonexistent.yaml"))
        assert config["outlook"]["calendar_name"] == "Calendar"
        assert config["sync"]["direction"] == "both"
        assert config["sync"]["days_back"] == 30

    def test_load_from_file(self, tmp_path):
        config_file = tmp_path / "config.yaml"
        config_file.write_text(
            "outlook:\n  calendar_name: Work\nsync:\n  days_back: 60\n"
        )
        config = load_config(str(config_file))
        assert config["outlook"]["calendar_name"] == "Work"
        assert config["sync"]["days_back"] == 60
        assert config["sync"]["days_forward"] == 90  # default preserved

    def test_paths_expanded(self, tmp_path):
        config = load_config(str(tmp_path / "nonexistent.yaml"))
        assert "~" not in config["state"]["db_path"]
