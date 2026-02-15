"""Configuration loading and validation."""

import os
from pathlib import Path

import yaml


DEFAULT_CONFIG_PATH = "~/.config/outlook-gcal-sync/config.yaml"

DEFAULT_CONFIG = {
    "outlook": {
        "calendar_name": "Calendar",
    },
    "microsoft": {
        "client_id": "",
        "tenant_id": "common",
        "calendar_id": "",
        "token_cache_path": "~/.config/outlook-gcal-sync/ms_token_cache.json",
    },
    "google": {
        "calendar_id": "primary",
        "credentials_path": "~/.config/outlook-gcal-sync/credentials.json",
        "token_path": "~/.config/outlook-gcal-sync/token.json",
    },
    "sync": {
        "direction": "both",
        "days_back": 30,
        "days_forward": 90,
        "conflict_resolution": "outlook-wins",
        "dry_run": False,
        "exclude_patterns": [],
    },
    "state": {
        "db_path": "~/.config/outlook-gcal-sync/sync_state.db",
    },
    "logging": {
        "level": "INFO",
        "file": "~/.config/outlook-gcal-sync/sync.log",
        "max_bytes": 5_242_880,
        "backup_count": 3,
    },
}


def _expand_paths(config: dict) -> dict:
    """Expand ~ in all path-like config values."""
    for section in config.values():
        if not isinstance(section, dict):
            continue
        for key, value in section.items():
            if isinstance(value, str) and ("~" in value or key.endswith("_path") or key == "file"):
                section[key] = os.path.expanduser(value)
    return config


def _deep_merge(base: dict, override: dict) -> dict:
    """Merge override into base, preferring override values."""
    result = base.copy()
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_config(config_path: str | None = None) -> dict:
    """Load configuration from YAML file, merged with defaults.

    Args:
        config_path: Path to config YAML file. If None, uses DEFAULT_CONFIG_PATH.

    Returns:
        Merged configuration dictionary with paths expanded.
    """
    path = Path(os.path.expanduser(config_path or DEFAULT_CONFIG_PATH))

    if path.exists():
        with open(path) as f:
            user_config = yaml.safe_load(f) or {}
        config = _deep_merge(DEFAULT_CONFIG, user_config)
    else:
        config = DEFAULT_CONFIG.copy()

    return _expand_paths(config)


def ensure_config_dir(config: dict) -> None:
    """Create the config directory and any parent directories if needed."""
    for section in config.values():
        if not isinstance(section, dict):
            continue
        for key, value in section.items():
            if isinstance(value, str) and (key.endswith("_path") or key == "file"):
                Path(value).parent.mkdir(parents=True, exist_ok=True)
