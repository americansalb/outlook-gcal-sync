#!/usr/bin/env python3
"""Utility to list available Outlook calendars via AppleScript."""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from src.outlook.applescript_bridge import check_outlook_running, list_calendars


def main():
    if not check_outlook_running():
        print("Error: Microsoft Outlook is not running. Please launch it first.")
        sys.exit(1)

    calendars = list_calendars()
    if not calendars:
        print("No calendars found in Outlook.")
        sys.exit(1)

    print(f"Found {len(calendars)} calendar(s) in Outlook:\n")
    for i, name in enumerate(calendars, 1):
        print(f"  {i}. {name}")
    print("\nUse the calendar name in your config.yaml under outlook.calendar_name")


if __name__ == "__main__":
    main()
