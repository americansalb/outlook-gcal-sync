"""Create, update, and delete calendar events in Outlook for Mac via AppleScript."""

import logging

from src.outlook.applescript_bridge import (
    escape_applescript_string,
    run_applescript_with_retry,
)
from src.sync.models import UnifiedEvent
from src.utils.datetime_utils import to_applescript_date_components

logger = logging.getLogger("outlook_gcal_sync.outlook.writer")


def _build_date_setter(var_name: str, components: dict) -> str:
    """Build AppleScript lines to set a date variable from components."""
    return f"""
    set {var_name} to current date
    set year of {var_name} to {components['year']}
    set month of {var_name} to {components['month']}
    set day of {var_name} to {components['day']}
    set hours of {var_name} to {components['hour']}
    set minutes of {var_name} to {components['minute']}
    set seconds of {var_name} to {components['second']}
"""


def create_event(calendar_name: str, event: UnifiedEvent) -> str:
    """Create a new event in Outlook and return its ID.

    Args:
        calendar_name: Name of the Outlook calendar.
        event: The event to create.

    Returns:
        The Outlook event ID as a string.
    """
    cal_name = escape_applescript_string(calendar_name)
    title = escape_applescript_string(event.title)
    location = escape_applescript_string(event.location)
    description = escape_applescript_string(event.description)

    start_comp = to_applescript_date_components(event.start_time)
    end_comp = to_applescript_date_components(event.end_time)

    start_setter = _build_date_setter("startDate", start_comp)
    end_setter = _build_date_setter("endDate", end_comp)

    reminder_lines = ""
    if event.has_reminder and event.reminder_minutes > 0:
        reminder_lines = f"""
        set has reminder of newEvent to true
        set reminder time of newEvent to {event.reminder_minutes}
"""

    script = f'''
tell application "Microsoft Outlook"
    set theCalendar to calendar "{cal_name}"
    {start_setter}
    {end_setter}
    set newEvent to make new calendar event with properties {{¬
        subject:"{title}", ¬
        start time:startDate, ¬
        end time:endDate, ¬
        all day flag:{str(event.is_all_day).lower()}, ¬
        location:"{location}", ¬
        content:"{description}" ¬
    }}
    {reminder_lines}
    return id of newEvent
end tell
'''

    result = run_applescript_with_retry(script)
    event_id = result.strip()
    logger.info("Created Outlook event '%s' (ID: %s)", event.title, event_id)
    return event_id


def update_event(calendar_name: str, outlook_id: str, event: UnifiedEvent) -> None:
    """Update an existing Outlook event.

    Args:
        calendar_name: Name of the Outlook calendar.
        outlook_id: The Outlook event ID to update.
        event: The event data to apply.
    """
    cal_name = escape_applescript_string(calendar_name)
    title = escape_applescript_string(event.title)
    location = escape_applescript_string(event.location)
    description = escape_applescript_string(event.description)

    start_comp = to_applescript_date_components(event.start_time)
    end_comp = to_applescript_date_components(event.end_time)

    start_setter = _build_date_setter("startDate", start_comp)
    end_setter = _build_date_setter("endDate", end_comp)

    reminder_lines = ""
    if event.has_reminder and event.reminder_minutes > 0:
        reminder_lines = f"""
        set has reminder of theEvent to true
        set reminder time of theEvent to {event.reminder_minutes}
"""
    else:
        reminder_lines = """
        set has reminder of theEvent to false
"""

    script = f'''
tell application "Microsoft Outlook"
    set theCalendar to calendar "{cal_name}"
    set theEvent to first calendar event of theCalendar whose id is {outlook_id}
    {start_setter}
    {end_setter}
    set subject of theEvent to "{title}"
    set location of theEvent to "{location}"
    set content of theEvent to "{description}"
    set start time of theEvent to startDate
    set end time of theEvent to endDate
    set all day flag of theEvent to {str(event.is_all_day).lower()}
    {reminder_lines}
end tell
'''

    run_applescript_with_retry(script)
    logger.info("Updated Outlook event '%s' (ID: %s)", event.title, outlook_id)


def delete_event(calendar_name: str, outlook_id: str) -> None:
    """Delete an event from Outlook.

    Args:
        calendar_name: Name of the Outlook calendar.
        outlook_id: The Outlook event ID to delete.
    """
    cal_name = escape_applescript_string(calendar_name)

    script = f'''
tell application "Microsoft Outlook"
    set theCalendar to calendar "{cal_name}"
    set theEvent to first calendar event of theCalendar whose id is {outlook_id}
    delete theEvent
end tell
'''

    run_applescript_with_retry(script)
    logger.info("Deleted Outlook event (ID: %s)", outlook_id)
