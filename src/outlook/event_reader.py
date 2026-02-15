"""Read calendar events from Outlook for Mac via AppleScript."""

import logging
from datetime import date, datetime, timedelta

from src.outlook.applescript_bridge import (
    escape_applescript_string,
    run_applescript_with_retry,
)
from src.sync.models import UnifiedEvent
from src.utils.datetime_utils import parse_applescript_components

logger = logging.getLogger("outlook_gcal_sync.outlook.reader")

# Field delimiter for AppleScript output. Using a control character to avoid
# conflicts with event content.
FIELD_DELIM = "|||"
RECORD_DELIM = "<<<>>>"


def _build_read_script(calendar_name: str, start_date: date, end_date: date) -> str:
    """Build AppleScript to read events from an Outlook calendar.

    Uses numeric date components to avoid locale-dependent parsing.
    """
    cal_name = escape_applescript_string(calendar_name)
    delim = FIELD_DELIM
    rec_delim = RECORD_DELIM

    # AppleScript date construction uses "current date" as a base then sets components.
    # We filter by iterating and comparing, since `whose` can be slow/unreliable.
    script = f'''
tell application "Microsoft Outlook"
    set theCalendar to calendar "{cal_name}"
    set allEvents to every calendar event of theCalendar

    set startYear to {start_date.year}
    set startMonth to {start_date.month}
    set startDay to {start_date.day}
    set endYear to {end_date.year}
    set endMonth to {end_date.month}
    set endDay to {end_date.day}

    -- Build boundary dates
    set filterStart to current date
    set year of filterStart to startYear
    set month of filterStart to startMonth
    set day of filterStart to startDay
    set hours of filterStart to 0
    set minutes of filterStart to 0
    set seconds of filterStart to 0

    set filterEnd to current date
    set year of filterEnd to endYear
    set month of filterEnd to endMonth
    set day of filterEnd to endDay
    set hours of filterEnd to 23
    set minutes of filterEnd to 59
    set seconds of filterEnd to 59

    set output to ""
    repeat with anEvent in allEvents
        set evStart to start time of anEvent
        if evStart >= filterStart and evStart <= filterEnd then
            set evId to id of anEvent
            set evSubject to subject of anEvent

            -- Extract start time components
            set sy to year of evStart
            set smo to (month of evStart) as integer
            set sd to day of evStart
            set sh to hours of evStart
            set smi to minutes of evStart
            set ss to seconds of evStart

            set evEnd to end time of anEvent
            set ey to year of evEnd
            set emo to (month of evEnd) as integer
            set ed to day of evEnd
            set eh to hours of evEnd
            set emi to minutes of evEnd
            set es to seconds of evEnd

            set evAllDay to all day flag of anEvent
            set evIsRecurring to is recurring of anEvent

            -- Safe property access (location/content may be missing)
            set evLocation to ""
            try
                set evLocation to location of anEvent
            end try

            set evContent to ""
            try
                set evContent to plain text content of anEvent
            end try
            if evContent is missing value then set evContent to ""

            set evHasReminder to has reminder of anEvent
            set evReminderTime to 0
            if evHasReminder then
                try
                    set evReminderTime to reminder time of anEvent
                end try
            end if

            -- Build delimited record
            set output to output & evId & "{delim}" & ¬
                evSubject & "{delim}" & ¬
                sy & "{delim}" & smo & "{delim}" & sd & "{delim}" & sh & "{delim}" & smi & "{delim}" & ss & "{delim}" & ¬
                ey & "{delim}" & emo & "{delim}" & ed & "{delim}" & eh & "{delim}" & emi & "{delim}" & es & "{delim}" & ¬
                evAllDay & "{delim}" & ¬
                evLocation & "{delim}" & ¬
                evContent & "{delim}" & ¬
                evHasReminder & "{delim}" & ¬
                evReminderTime & "{delim}" & ¬
                evIsRecurring & "{rec_delim}"
        end if
    end repeat
    return output
end tell
'''
    return script


def _parse_bool(value: str) -> bool:
    """Parse AppleScript boolean string."""
    return value.strip().lower() == "true"


def _parse_events(raw_output: str) -> list[UnifiedEvent]:
    """Parse the delimited AppleScript output into UnifiedEvent objects."""
    events = []
    if not raw_output.strip():
        return events

    records = raw_output.split(RECORD_DELIM)
    for record in records:
        record = record.strip()
        if not record:
            continue

        fields = record.split(FIELD_DELIM)
        if len(fields) < 20:
            logger.warning("Skipping malformed record with %d fields: %s", len(fields), record[:100])
            continue

        try:
            outlook_id = fields[0].strip()
            title = fields[1].strip()

            # Start time components
            sy, smo, sd = int(fields[2]), int(fields[3]), int(fields[4])
            sh, smi, ss = int(fields[5]), int(fields[6]), int(fields[7])

            # End time components
            ey, emo, ed = int(fields[8]), int(fields[9]), int(fields[10])
            eh, emi, es = int(fields[11]), int(fields[12]), int(fields[13])

            is_all_day = _parse_bool(fields[14])
            location = fields[15].strip()
            description = fields[16].strip()
            has_reminder = _parse_bool(fields[17])
            reminder_minutes = int(fields[18].strip()) if fields[18].strip() else 0
            is_recurring = _parse_bool(fields[19])

            start_time = parse_applescript_components(sy, smo, sd, sh, smi, ss)
            end_time = parse_applescript_components(ey, emo, ed, eh, emi, es)

            event = UnifiedEvent(
                source="outlook",
                source_id=str(outlook_id),
                title=title,
                description=description,
                location=location,
                start_time=start_time,
                end_time=end_time,
                is_all_day=is_all_day,
                has_reminder=has_reminder,
                reminder_minutes=reminder_minutes,
                is_recurring=is_recurring,
                start_date=start_time.date() if is_all_day else None,
                end_date=end_time.date() if is_all_day else None,
            )
            events.append(event)

        except (ValueError, IndexError) as e:
            logger.warning("Failed to parse event record: %s. Error: %s", record[:100], e)
            continue

    return events


def read_outlook_events(
    calendar_name: str,
    days_back: int = 30,
    days_forward: int = 90,
) -> list[UnifiedEvent]:
    """Read events from an Outlook calendar within the specified date range.

    Args:
        calendar_name: Name of the Outlook calendar.
        days_back: Number of days in the past to include.
        days_forward: Number of days in the future to include.

    Returns:
        List of UnifiedEvent objects.
    """
    today = date.today()
    start_date = today - timedelta(days=days_back)
    end_date = today + timedelta(days=days_forward)

    logger.info(
        "Reading Outlook events from '%s' (%s to %s)...",
        calendar_name, start_date, end_date,
    )

    script = _build_read_script(calendar_name, start_date, end_date)
    raw_output = run_applescript_with_retry(script, timeout=120)

    events = _parse_events(raw_output)
    logger.info("Read %d events from Outlook.", len(events))
    return events
