#!/usr/bin/env python3
"""Utility to verify Google Calendar API authentication."""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from src.config import load_config
from src.google_cal.auth import build_calendar_service


def main():
    config = load_config()

    print("Verifying Google Calendar API authentication...")
    try:
        service = build_calendar_service(
            config["google"]["credentials_path"],
            config["google"]["token_path"],
        )
        result = service.calendarList().list().execute()
        calendars = result.get("items", [])
        print(f"\nAuthentication successful!")
        print(f"Found {len(calendars)} calendar(s):")
        for cal in calendars:
            primary = " (primary)" if cal.get("primary") else ""
            print(f"  - {cal['summary']}{primary}")
    except FileNotFoundError as e:
        print(f"\nError: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\nAuthentication failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
