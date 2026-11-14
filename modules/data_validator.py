"""Data validation module for individual record fields."""
import datetime
import math
import re
from dataclasses import dataclass
from typing import Any, Optional

from modules.date_parser import DateParser, DateParseResult
from utils.error_handler import ErrorCategory, ProcessingError


@dataclass
class RecordValidationResult:
    """Outcome of validating a single row/record."""
    
    is_valid: bool
    clean_name: Optional[str] = None
    parsed_date: Optional[datetime.date] = None
    days_difference: Optional[int] = None
    error: Optional[ProcessingError] = None


class DataValidator:
    """Validates individual record fields (Name and Date)."""

    @classmethod
    def validate_record(
        cls,
        raw_name: Any,
        raw_date: Any,
        row_number: int,
        source_file: str,
        source_sheet: str,
        current_date: datetime.date,
        date_preference: str = "Auto-detect",
        timezone_name: str = "Asia/Kolkata",
        column_context_day_first: Optional[bool] = None,
        name_col_name: str = "Name",
        date_col_name: str = "Last Updated",
    ) -> RecordValidationResult:
        """Validate and clean a single record from a worksheet.
        
        Args:
            raw_name: Uncleaned name value
            raw_date: Unparsed date value
            row_number: Excel/CSV row number (1-based index)
            source_file: Name of source file
            source_sheet: Name of source worksheet
            current_date: Reference current date for filtering
            date_preference: Format preference for date parsing
            timezone_name: Application timezone
            column_context_day_first: Format hint from column context
            name_col_name: Header name of the Name column
            date_col_name: Header name of the Date column
            
        Returns:
            RecordValidationResult containing cleaned values or structured error.
        """
        # 1. Validate Name
        clean_name, name_error = cls._validate_name(
            raw_name, row_number, source_file, source_sheet, name_col_name
        )
        if name_error:
            return RecordValidationResult(is_valid=False, error=name_error)

        # 2. Validate and Parse Date
        parsed_date, date_error = cls._validate_date(
            raw_date,
            row_number,
            source_file,
            source_sheet,
            current_date,
            date_preference,
            timezone_name,
            column_context_day_first,
            date_col_name,
        )
        if date_error:
            return RecordValidationResult(
                is_valid=False,
                clean_name=clean_name,
                error=date_error,
            )

        # Calculate calendar day difference
        days_difference = (current_date - parsed_date).days

        return RecordValidationResult(
            is_valid=True,
            clean_name=clean_name,
            parsed_date=parsed_date,
            days_difference=days_difference,
        )

    @classmethod
    def _validate_name(
        cls,
        val: Any,
        row_number: int,
        source_file: str,
        source_sheet: str,
        col_name: str,
    ) -> tuple[Optional[str], Optional[ProcessingError]]:
        """Validate that the name field is not empty, NaN, or pure punctuation."""
        if val is None:
            return None, ProcessingError(
                source_file=source_file,
                source_sheet=source_sheet,
                row_number=row_number,
                column_name=col_name,
                error_category=ErrorCategory.MISSING_NAME,
                explanation="Name field is missing (null/None).",
                original_value=val,
            )

        if isinstance(val, float) and math.isnan(val):
            return None, ProcessingError(
                source_file=source_file,
                source_sheet=source_sheet,
                row_number=row_number,
                column_name=col_name,
                error_category=ErrorCategory.MISSING_NAME,
                explanation="Name field is missing (NaN).",
                original_value=val,
            )

        text = str(val).strip()
        if not text or text.lower() in ("nan", "none", "null", "n/a", "-"):
            return None, ProcessingError(
                source_file=source_file,
                source_sheet=source_sheet,
                row_number=row_number,
                column_name=col_name,
                error_category=ErrorCategory.MISSING_NAME,
                explanation="Name field is empty or contains only whitespace.",
                original_value=val,
            )

        # Check if text is only punctuation / symbols
        if not re.search(r'[\w]', text):
            return None, ProcessingError(
                source_file=source_file,
                source_sheet=source_sheet,
                row_number=row_number,
                column_name=col_name,
                error_category=ErrorCategory.MISSING_NAME,
                explanation="Name contains no alphanumeric characters.",
                original_value=val,
            )

        return text, None

    @classmethod
    def _validate_date(
        cls,
        val: Any,
        row_number: int,
        source_file: str,
        source_sheet: str,
        current_date: datetime.date,
        date_preference: str,
        timezone_name: str,
        column_context_day_first: Optional[bool],
        col_name: str,
    ) -> tuple[Optional[datetime.date], Optional[ProcessingError]]:
        """Parse and validate date, checking for missing, invalid, or future dates."""
        if val is None or (isinstance(val, float) and math.isnan(val)):
            return None, ProcessingError(
                source_file=source_file,
                source_sheet=source_sheet,
                row_number=row_number,
                column_name=col_name,
                error_category=ErrorCategory.MISSING_DATE,
                explanation="Last Updated Date is empty / missing.",
                original_value=val,
            )

        str_val = str(val).strip()
        if not str_val or str_val.lower() in ("nan", "none", "null", "nat", "n/a", "-"):
            return None, ProcessingError(
                source_file=source_file,
                source_sheet=source_sheet,
                row_number=row_number,
                column_name=col_name,
                error_category=ErrorCategory.MISSING_DATE,
                explanation="Last Updated Date is empty string.",
                original_value=val,
            )

        parse_res: DateParseResult = DateParser.parse_date(
            val,
            preference=date_preference,
            timezone_name=timezone_name,
            column_context_day_first=column_context_day_first,
        )

        if not parse_res.success or parse_res.date_obj is None:
            return None, ProcessingError(
                source_file=source_file,
                source_sheet=source_sheet,
                row_number=row_number,
                column_name=col_name,
                error_category=ErrorCategory.INVALID_DATE,
                explanation=parse_res.error_message or "Failed to parse date value.",
                original_value=val,
            )

        parsed_date = parse_res.date_obj

        # Check for future dates: Last updated cannot be after the current reference date
        if parsed_date > current_date:
            days_ahead = (parsed_date - current_date).days
            return None, ProcessingError(
                source_file=source_file,
                source_sheet=source_sheet,
                row_number=row_number,
                column_name=col_name,
                error_category=ErrorCategory.FUTURE_DATE,
                explanation=f"Date '{parsed_date.isoformat()}' is in the future ({days_ahead} day(s) after current date {current_date.isoformat()}).",
                original_value=val,
            )

        return parsed_date, None
