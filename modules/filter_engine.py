"""Filtering engine for record evaluation and calendar-day difference calculation."""
import datetime
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from modules.data_validator import DataValidator, RecordValidationResult
from utils.error_handler import ProcessingError


@dataclass
class Record:
    """Represents a validated record from an input worksheet."""
    
    name: str
    last_updated: datetime.date
    days_difference: int
    source_file: str
    source_sheet: str
    row_number: int
    incident: str = ""
    sctask: str = ""
    short_description: str = ""

    def to_dict(self, date_format: str = "%d-%b-%Y") -> Dict[str, Any]:
        """Convert record to dictionary for display and report export."""
        return {
            "Incident": self.incident,
            "SCTask": self.sctask,
            "Short Description": self.short_description,
            "Name": self.name,
            "Last Updated": self.last_updated.strftime(date_format),
            "Source File": self.source_file,
        }


@dataclass
class FilterResult:
    """Consolidated outcome of applying filter rules across one or more datasets."""
    
    matching_records: List[Record] = field(default_factory=list)
    skipped_records: List[Record] = field(default_factory=list)
    errors: List[ProcessingError] = field(default_factory=list)
    threshold_days: int = 5
    current_date: Optional[datetime.date] = None
    total_examined: int = 0

    @property
    def total_matching(self) -> int:
        return len(self.matching_records)

    @property
    def total_skipped(self) -> int:
        return len(self.skipped_records)

    @property
    def total_errors(self) -> int:
        return len(self.errors)


class FilterEngine:
    """Evaluates tabular records against the calendar-day aging threshold."""

    @classmethod
    def filter_records(
        cls,
        rows_data: List[Dict[str, Any]],
        name_col: str,
        date_col: str,
        source_file: str,
        source_sheet: str,
        current_date: datetime.date,
        threshold_days: int = 5,
        date_preference: str = "Auto-detect",
        timezone_name: str = "Asia/Kolkata",
        column_context_day_first: Optional[bool] = None,
        incident_col: Optional[str] = None,
        sctask_col: Optional[str] = None,
        short_desc_col: Optional[str] = None,
    ) -> FilterResult:
        """Filter a list of rows for a single worksheet.
        
        Args:
            rows_data: List of row dictionaries (e.g. df.to_dict(orient='records'))
            name_col: Column name containing Name
            date_col: Column name containing Last Updated Date
            source_file: File name for traceability
            source_sheet: Sheet name for traceability
            current_date: Dynamic reference date
            threshold_days: Days threshold (strict > condition, default: 5)
            date_preference: Format preference for ambiguous dates
            timezone_name: Application timezone
            column_context_day_first: Hint inferred from column sample
            incident_col: Optional column name for Incident ticket number
            sctask_col: Optional column name for SCTask ticket number
            short_desc_col: Optional column name for Short Description
            
        Returns:
            FilterResult containing matching records, skipped records, and errors.
        """
        result = FilterResult(
            threshold_days=threshold_days,
            current_date=current_date,
            total_examined=len(rows_data),
        )

        for idx, row in enumerate(rows_data):
            row_num = idx + 2  # 1-based index with header as row 1
            raw_name = row.get(name_col)
            raw_date = row.get(date_col)

            # Validate record
            val_res: RecordValidationResult = DataValidator.validate_record(
                raw_name=raw_name,
                raw_date=raw_date,
                row_number=row_num,
                source_file=source_file,
                source_sheet=source_sheet,
                current_date=current_date,
                date_preference=date_preference,
                timezone_name=timezone_name,
                column_context_day_first=column_context_day_first,
                name_col_name=name_col,
                date_col_name=date_col,
            )

            if not val_res.is_valid:
                if val_res.error:
                    result.errors.append(val_res.error)
                continue

            # Extract incident, sctask, and short_description values
            inc_val = ""
            if incident_col and row.get(incident_col) is not None:
                v = str(row.get(incident_col)).strip()
                if v.lower() not in ("nan", "none", "null"):
                    inc_val = v

            sctask_val = ""
            if sctask_col and row.get(sctask_col) is not None:
                v = str(row.get(sctask_col)).strip()
                if v.lower() not in ("nan", "none", "null"):
                    sctask_val = v

            desc_val = ""
            if short_desc_col and row.get(short_desc_col) is not None:
                v = str(row.get(short_desc_col)).strip()
                if v.lower() not in ("nan", "none", "null"):
                    desc_val = v

            record = Record(
                name=val_res.clean_name,
                last_updated=val_res.parsed_date,
                days_difference=val_res.days_difference,
                source_file=source_file,
                source_sheet=source_sheet,
                row_number=row_num,
                incident=inc_val,
                sctask=sctask_val,
                short_description=desc_val,
            )

            # Filtering Condition: STRICTLY GREATER THAN threshold
            # Days Difference > Threshold
            # Example: threshold = 5:
            # diff = 5 -> EXCLUDED (skipped)
            # diff = 6 -> INCLUDED (matching)
            if record.days_difference > threshold_days:
                result.matching_records.append(record)
            else:
                result.skipped_records.append(record)

        return result
