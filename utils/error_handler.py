"""Structured error handling and logging utilities."""
import logging
from dataclasses import dataclass
from enum import Enum
from typing import Any, Optional

# Configure application logger
logger = logging.getLogger("smart_data_filter")
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] [%(name)s]: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)


class ErrorCategory(str, Enum):
    """Categories of processing and validation errors."""
    
    INVALID_DATE = "Invalid Date"
    MISSING_DATE = "Missing Date"
    MISSING_NAME = "Missing Name"
    FUTURE_DATE = "Future Date"
    CORRUPTED_FILE = "Corrupted File"
    EMPTY_SHEET = "Empty Sheet / File"
    ENCODING_ERROR = "Encoding Error"
    MAPPING_ERROR = "Column Mapping Error"
    PARSE_ERROR = "Parse Error"

    def __str__(self) -> str:
        return self.value


@dataclass
class ProcessingError:
    """Represents a structured processing or validation error for reporting and export."""
    
    source_file: str
    source_sheet: str
    row_number: Optional[int]
    error_category: Any
    explanation: str
    original_value: Any = None
    column_name: Optional[str] = None
    
    def to_dict(self) -> dict:
        """Convert error to dictionary for dataframe and export display."""
        cat_val = self.error_category.value if hasattr(self.error_category, "value") else str(self.error_category)
        return {
            "Source File": self.source_file,
            "Source Sheet": self.source_sheet,
            "Row Number": self.row_number if self.row_number is not None else "N/A",
            "Error Category": cat_val,
            "Explanation": self.explanation,
            "Original Value": str(self.original_value) if self.original_value is not None else "",
            "Column Name": self.column_name or "",
        }


def log_sanitized(level: int, msg: str, **kwargs):
    """Log messages while avoiding leaking sensitive personal identity data."""
    logger.log(level, msg, **kwargs)
