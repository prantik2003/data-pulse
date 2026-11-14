"""Date parsing and normalization module supporting diverse formats, serial dates, and timezones."""
import datetime
import math
import re
from dataclasses import dataclass
from typing import Any, List, Optional, Tuple, Union
from zoneinfo import ZoneInfo
from dateutil import parser as dateutil_parser

from utils.date_utils import convert_to_timezone, to_calendar_date


@dataclass
class DateParseResult:
    """Result of parsing an individual date value."""
    
    success: bool
    date_obj: Optional[datetime.date] = None
    is_ambiguous: bool = False
    error_message: Optional[str] = None
    original_value: Any = None


class DateParser:
    """Robust parser for diverse date representations with ambiguity resolution."""

    # Reasonable range for Excel serial dates:
    # 25569 = 1970-01-01
    # 54789 = 2050-01-01
    EXCEL_SERIAL_MIN = 25569
    EXCEL_SERIAL_MAX = 54789
    EXCEL_EPOCH = datetime.date(1899, 12, 30)

    # Standard explicit date format patterns
    EXPLICIT_STRPTIME_FORMATS = [
        # Year first (unambiguous)
        ("%Y-%m-%d", False),
        ("%Y/%m/%d", False),
        ("%Y.%m.%d", False),
        # Day month text year (unambiguous)
        ("%d-%b-%Y", False),
        ("%d-%B-%Y", False),
        ("%d %b %Y", False),
        ("%d %B %Y", False),
        ("%b %d, %Y", False),
        ("%B %d, %Y", False),
        ("%b %d %Y", False),
        ("%B %d %Y", False),
    ]

    @classmethod
    def parse_date(
        cls,
        val: Any,
        preference: str = "Auto-detect",
        timezone_name: str = "Asia/Kolkata",
        column_context_day_first: Optional[bool] = None,
    ) -> DateParseResult:
        """Parse any input value into a normalized Python datetime.date object.
        
        Args:
            val: The input value (string, int, float, datetime, date, Timestamp, etc.)
            preference: User format preference: 'Auto-detect', 'DD/MM/YYYY', 'MM/DD/YYYY', 'YYYY-MM-DD'
            timezone_name: Target timezone for timestamp normalization (default: Asia/Kolkata)
            column_context_day_first: Hint inferred from other unambiguous rows in the same column
            
        Returns:
            DateParseResult with success status and parsed calendar date.
        """
        if val is None:
            return DateParseResult(success=False, error_message="Empty date value (None).", original_value=val)

        # Handle float NaN or string 'nan'/'none'/'null'
        if isinstance(val, float) and math.isnan(val):
            return DateParseResult(success=False, error_message="Empty date value (NaN).", original_value=val)

        # 1. Native Python date (and not datetime)
        if isinstance(val, datetime.date) and not isinstance(val, datetime.datetime):
            return DateParseResult(success=True, date_obj=val, original_value=val)

        # 2. Native Python datetime or Pandas Timestamp
        if isinstance(val, datetime.datetime) or hasattr(val, "to_pydatetime"):
            try:
                py_dt = val.to_pydatetime() if hasattr(val, "to_pydatetime") else val
                cal_date = to_calendar_date(py_dt, timezone_name)
                return DateParseResult(success=True, date_obj=cal_date, original_value=val)
            except Exception as e:
                return DateParseResult(success=False, error_message=f"Datetime conversion failed: {str(e)}", original_value=val)

        # 3. Numeric: Excel serial date or invalid numeric ID
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            return cls._parse_numeric_serial(val)

        # 4. String representations
        str_val = str(val).strip()
        if not str_val or str_val.lower() in ("nan", "none", "null", "nat", "n/a", "-"):
            return DateParseResult(success=False, error_message="Empty date string.", original_value=val)

        # Check if string is a numeric string that could be an Excel serial date
        if re.match(r'^\d+(\.\d+)?$', str_val):
            try:
                numeric_val = float(str_val)
                # If in Excel serial date range, try serial date parsing
                if cls.EXCEL_SERIAL_MIN <= numeric_val <= cls.EXCEL_SERIAL_MAX:
                    return cls._parse_numeric_serial(numeric_val)
                else:
                    return DateParseResult(
                        success=False,
                        error_message=f"Numeric value '{str_val}' is outside plausible date ranges and treated as an ID, not a date.",
                        original_value=val,
                    )
            except ValueError:
                pass

        # Parse string using format rules and preferences
        return cls._parse_string_date(str_val, preference, timezone_name, column_context_day_first)

    @classmethod
    def _parse_numeric_serial(cls, num: Union[int, float]) -> DateParseResult:
        """Convert Excel serial date number to Python date."""
        if not (cls.EXCEL_SERIAL_MIN <= num <= cls.EXCEL_SERIAL_MAX):
            return DateParseResult(
                success=False,
                error_message=f"Numeric value {num} is outside valid Excel date range (1970 - 2050).",
                original_value=num,
            )
        try:
            # Excel serial days from 1899-12-30
            days = int(math.floor(num))
            result_date = cls.EXCEL_EPOCH + datetime.timedelta(days=days)
            return DateParseResult(success=True, date_obj=result_date, original_value=num)
        except Exception as e:
            return DateParseResult(
                success=False,
                error_message=f"Failed to convert Excel serial date {num}: {str(e)}",
                original_value=num,
            )

    @classmethod
    def _parse_string_date(
        cls,
        text: str,
        preference: str,
        timezone_name: str,
        column_context_day_first: Optional[bool],
    ) -> DateParseResult:
        """Parse structured string representations of dates."""
        clean_text = text.strip()

        # Handle ISO strings with T or Z
        if "T" in clean_text or clean_text.endswith("Z"):
            try:
                # Replace trailing Z with UTC offset
                iso_str = clean_text.replace("Z", "+00:00")
                dt = datetime.datetime.fromisoformat(iso_str)
                cal_date = to_calendar_date(dt, timezone_name)
                return DateParseResult(success=True, date_obj=cal_date, original_value=text)
            except Exception:
                pass  # Fall through to dateutil parser

        # Try unambiguous explicit formats
        for fmt, _ in cls.EXPLICIT_STRPTIME_FORMATS:
            try:
                dt = datetime.datetime.strptime(clean_text, fmt)
                return DateParseResult(success=True, date_obj=dt.date(), original_value=text)
            except ValueError:
                continue

        # Check for slash / dash / dot formats: DD/MM/YYYY or MM/DD/YYYY or DD-MM-YYYY or DD.MM.YYYY
        pattern_match = re.match(r'^(\d{1,2})([\/\-\.])(\d{1,2})([\/\-\.])(\d{4})(?:\s+.*)?$', clean_text)
        if pattern_match:
            p1 = int(pattern_match.group(1))
            sep = pattern_match.group(2)
            p2 = int(pattern_match.group(3))
            year = int(pattern_match.group(5))

            # Determine day and month
            is_ambiguous = (p1 <= 12 and p2 <= 12 and p1 != p2)

            day_first: bool
            if preference == "DD/MM/YYYY":
                day_first = True
            elif preference == "MM/DD/YYYY":
                day_first = False
            elif column_context_day_first is not None:
                day_first = column_context_day_first
            elif p1 > 12:
                # Unambiguously DD/MM/YYYY
                day_first = True
            elif p2 > 12:
                # Unambiguously MM/DD/YYYY
                day_first = False
            elif sep in ("-", "."):
                # By international standard, hyphenated/dotted (e.g. 05-11-2026 or 05.11.2026) is usually DD-MM-YYYY
                day_first = True
            else:
                # Default for Indian/British/Standard formats is DD/MM/YYYY
                day_first = True

            day, month = (p1, p2) if day_first else (p2, p1)

            try:
                parsed_date = datetime.date(year, month, day)
                return DateParseResult(
                    success=True,
                    date_obj=parsed_date,
                    is_ambiguous=is_ambiguous,
                    original_value=text,
                )
            except ValueError as e:
                # Try the reverse if ValueError occurred (e.g. day was out of range for month)
                try:
                    alt_day, alt_month = (p2, p1) if day_first else (p1, p2)
                    alt_date = datetime.date(year, alt_month, alt_day)
                    return DateParseResult(
                        success=True,
                        date_obj=alt_date,
                        is_ambiguous=is_ambiguous,
                        original_value=text,
                    )
                except ValueError:
                    return DateParseResult(
                        success=False,
                        error_message=f"Invalid calendar date values ({str(e)}).",
                        original_value=text,
                    )

        # Fallback to dateutil parser with dayfirst setting
        day_first_hint = True
        if preference == "MM/DD/YYYY":
            day_first_hint = False
        elif column_context_day_first is not None:
            day_first_hint = column_context_day_first

        try:
            dt = dateutil_parser.parse(clean_text, dayfirst=day_first_hint)
            cal_date = to_calendar_date(dt, timezone_name)
            return DateParseResult(success=True, date_obj=cal_date, original_value=text)
        except Exception as e:
            return DateParseResult(
                success=False,
                error_message=f"Unrecognized date format: '{clean_text}' ({str(e)})",
                original_value=text,
            )

    @classmethod
    def infer_column_day_first(cls, sample_values: List[Any]) -> Optional[bool]:
        """Analyze sample values in a column to establish whether DD/MM or MM/DD is used.
        
        If we find numbers > 12 in the first position, DD/MM is confirmed.
        If we find numbers > 12 in the second position, MM/DD is confirmed.
        """
        for val in sample_values:
            if val is None:
                continue
            s = str(val).strip()
            match = re.match(r'^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.]\d{4}', s)
            if match:
                p1, p2 = int(match.group(1)), int(match.group(2))
                if p1 > 12 and p2 <= 12:
                    return True  # Must be DD/MM
                elif p2 > 12 and p1 <= 12:
                    return False  # Must be MM/DD
        return None

    @classmethod
    def is_likely_date_column(
        cls,
        sample_values: List[Any],
        allow_numeric_serial: bool = False,
    ) -> Tuple[float, int, int]:
        """Check the proportion of sample values that can be parsed as dates.
        
        Args:
            sample_values: Sample of values from the column.
            allow_numeric_serial: Whether to treat raw numbers as potential Excel serial dates.
                                  Only enabled if the header suggests a date to prevent numeric
                                  IDs/salaries from being misinterpreted as dates.
        
        Returns:
            Tuple of (valid_ratio: float, valid_count: int, total_checked: int)
        """
        valid_count = 0
        total_checked = 0

        # First infer if column has day-first pattern
        day_first_hint = cls.infer_column_day_first(sample_values)

        for val in sample_values:
            if val is None or (isinstance(val, float) and math.isnan(val)):
                continue
            str_val = str(val).strip()
            if not str_val or str_val.lower() in ("nan", "none", "null", "nat", "n/a", "-"):
                continue

            # If value is numeric and numeric serial dates are not allowed for this column
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                total_checked += 1
                if not allow_numeric_serial:
                    continue  # Do not treat ordinary numeric ID / amount as date

            total_checked += 1
            res = cls.parse_date(val, column_context_day_first=day_first_hint)
            if res.success:
                valid_count += 1

        if total_checked == 0:
            return 0.0, 0, 0

        return valid_count / total_checked, valid_count, total_checked
