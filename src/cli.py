"""CLI entry point for outlook-gcal-sync."""

import argparse
import sys
import logging

from src.config import load_config, ensure_config_dir
from src.utils.logging_config import setup_logging


def cmd_setup(config: dict) -> None:
    """Interactive Google OAuth setup."""
    from src.google_cal.auth import build_calendar_service

    print("Setting up Google Calendar authentication...")
    print("A browser window will open for you to authorize access.\n")

    try:
        service = build_calendar_service(
            config["google"]["credentials_path"],
            config["google"]["token_path"],
        )
        # Verify access by listing calendars
        result = service.calendarList().list().execute()
        calendars = result.get("items", [])
        print(f"\nSuccess! Connected to Google Calendar.")
        print(f"Found {len(calendars)} calendar(s):")
        for cal in calendars:
            primary = " (primary)" if cal.get("primary") else ""
            print(f"  - {cal['summary']}{primary} [{cal['id']}]")
        print(f"\nToken saved to: {config['google']['token_path']}")
    except FileNotFoundError as e:
        print(f"\nError: {e}")
        print("\nTo set up Google Calendar API credentials:")
        print("1. Go to https://console.cloud.google.com/")
        print("2. Create a new project (e.g., 'Calendar Sync')")
        print("3. Enable the Google Calendar API")
        print("4. Create OAuth 2.0 Client ID (Desktop app)")
        print("5. Download the JSON and save it to:")
        print(f"   {config['google']['credentials_path']}")
        sys.exit(1)
    except Exception as e:
        print(f"\nError during setup: {e}")
        sys.exit(1)


def cmd_list_calendars(config: dict) -> None:
    """List available Outlook calendars."""
    from src.outlook.applescript_bridge import check_outlook_running, list_calendars

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
    print(f"\nCurrent config uses: '{config['outlook']['calendar_name']}'")


def cmd_sync(config: dict, dry_run: bool = False) -> None:
    """Run the sync engine."""
    from src.google_cal.auth import build_calendar_service
    from src.google_cal.client import GoogleCalendarClient
    from src.outlook.applescript_bridge import check_outlook_running
    from src.sync.engine import SyncEngine
    from src.sync.state import SyncStateStore

    logger = logging.getLogger("outlook_gcal_sync")

    # Pre-checks
    if not check_outlook_running():
        logger.error("Microsoft Outlook is not running. Please launch it first.")
        sys.exit(1)

    if dry_run:
        config["sync"]["dry_run"] = True

    # Initialize components
    try:
        service = build_calendar_service(
            config["google"]["credentials_path"],
            config["google"]["token_path"],
        )
    except FileNotFoundError:
        logger.error(
            "Google credentials not found. Run 'outlook-gcal-sync setup' first."
        )
        sys.exit(1)

    google_client = GoogleCalendarClient(service, config["google"]["calendar_id"])
    state_store = SyncStateStore(config["state"]["db_path"])

    try:
        engine = SyncEngine(config, google_client, state_store)
        summary = engine.run()

        if dry_run:
            print("\n--- DRY RUN SUMMARY ---")
        else:
            print("\n--- SYNC SUMMARY ---")
        print(f"  Created: {summary['created']}")
        print(f"  Updated: {summary['updated']}")
        print(f"  Deleted: {summary['deleted']}")
        print(f"  Skipped: {summary['skipped']}")
        print(f"  Conflicts: {summary['conflicts']}")
        if summary.get("errors", 0) > 0:
            print(f"  Errors: {summary['errors']}")
    finally:
        state_store.close()


def cmd_status(config: dict) -> None:
    """Show sync status and statistics."""
    from src.sync.state import SyncStateStore
    import os

    db_path = config["state"]["db_path"]
    if not os.path.exists(db_path):
        print("No sync state found. Run 'outlook-gcal-sync sync' first.")
        return

    state_store = SyncStateStore(db_path)
    try:
        stats = state_store.get_stats()
        print("--- SYNC STATUS ---")
        print(f"  Total synced event pairs: {stats['total_pairs']}")
        print(f"  Last sync: {stats['last_sync_time'] or 'Never'}")
        if stats["recent_actions"]:
            print("  Actions in last 24h:")
            for action, count in stats["recent_actions"].items():
                print(f"    {action}: {count}")
        else:
            print("  No actions in last 24h.")
    finally:
        state_store.close()


def cmd_reset(config: dict) -> None:
    """Reset sync state."""
    from src.sync.state import SyncStateStore
    import os

    db_path = config["state"]["db_path"]
    if not os.path.exists(db_path):
        print("No sync state to reset.")
        return

    response = input("This will clear all sync state. Events will NOT be deleted. Continue? [y/N] ")
    if response.strip().lower() != "y":
        print("Aborted.")
        return

    state_store = SyncStateStore(db_path)
    try:
        state_store.reset()
        print("Sync state reset. Next sync will be a full sync.")
    finally:
        state_store.close()


def main():
    parser = argparse.ArgumentParser(
        prog="outlook-gcal-sync",
        description="Bidirectional sync between Outlook for Mac and Google Calendar",
    )
    parser.add_argument(
        "-c", "--config",
        help="Path to config YAML file",
        default=None,
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # setup
    subparsers.add_parser("setup", help="Set up Google Calendar authentication")

    # sync
    sync_parser = subparsers.add_parser("sync", help="Run calendar sync")
    sync_parser.add_argument("--dry-run", action="store_true", help="Preview changes without applying")
    sync_parser.add_argument(
        "--direction",
        choices=["both", "outlook-to-google", "google-to-outlook"],
        help="Override sync direction",
    )

    # status
    subparsers.add_parser("status", help="Show sync status and statistics")

    # list-calendars
    subparsers.add_parser("list-calendars", help="List available Outlook calendars")

    # reset
    subparsers.add_parser("reset", help="Reset sync state (events are not deleted)")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    # Load config
    config = load_config(args.command != "setup" and args.config or args.config)
    ensure_config_dir(config)
    setup_logging(config)

    # Override direction if specified
    if args.command == "sync" and args.direction:
        config["sync"]["direction"] = args.direction

    # Dispatch
    if args.command == "setup":
        cmd_setup(config)
    elif args.command == "sync":
        cmd_sync(config, dry_run=args.dry_run)
    elif args.command == "status":
        cmd_status(config)
    elif args.command == "list-calendars":
        cmd_list_calendars(config)
    elif args.command == "reset":
        cmd_reset(config)


if __name__ == "__main__":
    main()
