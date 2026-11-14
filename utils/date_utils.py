"""Date and timezone utility functions."""
import datetime
from zoneinfo import ZoneInfo
from typing import Union


def get_timezone(timezone_name: str = "Asia/Kolkata") -> ZoneInfo:
    """Return a ZoneInfo instance for the specified timezone name, fallback to Asia/Kolkata."""
    try:
        return ZoneInfo(timezone_name)
    except Exception:
        return ZoneInfo("Asia/Kolkata")


def get_current_date(timezone_name: str = "Asia/Kolkata") -> datetime.date:
    """Return the current calendar date in the specified application timezone."""
    tz = get_timezone(timezone_name)
    now_in_tz = datetime.datetime.now(tz)
    return now_in_tz.date()


def convert_to_timezone(
    dt: datetime.datetime, timezone_name: str = "Asia/Kolkata"
) -> datetime.datetime:
    """Convert a datetime object to the target timezone.
    
    If dt is timezone-naive, it is assumed to be in the local target timezone.
    If dt is timezone-aware, it is converted to the target timezone.
    """
    tz = get_timezone(timezone_name)
    if dt.tzinfo is None:
        # Naive datetime: treat as already in application timezone
        return dt.replace(tzinfo=tz)
    else:
        # Aware datetime: convert to application timezone
        return dt.astimezone(tz)


def to_calendar_date(
    dt_or_date: Union[datetime.date, datetime.datetime],
    timezone_name: str = "Asia/Kolkata",
) -> datetime.date:
    """Normalize a datetime or date object into a pure calendar date in target timezone.
    
    If it's already a datetime.date (and not datetime.datetime), returns it directly.
    If it's a datetime.datetime, converts timezone (if aware) and extracts .date().
    """
    if isinstance(dt_or_date, datetime.datetime):
        aware_dt = convert_to_timezone(dt_or_date, timezone_name)
        return aware_dt.date()
    elif isinstance(dt_or_date, datetime.date):
        return dt_or_date
    else:
        raise TypeError(f"Expected datetime.date or datetime.datetime, got {type(dt_or_date)}")


def calculate_calendar_day_difference(
    last_updated_date: datetime.date,
    current_date: datetime.date,
) -> int:
    """Calculate the number of calendar days between last_updated_date and current_date.
    
    Formula:
        Difference in Days = Current Date - Last Updated Date
    
    Uses pure calendar dates so time of day or elapsed hours do not affect the result.
    Example:
        current_date = 2026-11-13
        last_updated_date = 2026-11-08 -> 5 days
        last_updated_date = 2026-11-07 -> 6 days
        last_updated_date = 2026-11-14 -> -1 days (future date)
    """
    if not isinstance(last_updated_date, datetime.date) or not isinstance(current_date, datetime.date):
        raise TypeError("Both dates must be datetime.date instances")
    return (current_date - last_updated_date).days
