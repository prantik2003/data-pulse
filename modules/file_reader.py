"""File reading module for XLSX and CSV files with delimiter and encoding auto-detection."""
import csv
import io
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

import openpyxl
import pandas as pd

from utils.error_handler import ErrorCategory, ProcessingError, log_sanitized
from utils.file_utils import get_safe_filename, is_valid_extension


@dataclass
class SheetData:
    """Represents a single tabular sheet extracted from an Excel or CSV file."""
    
    sheet_name: str
    df: pd.DataFrame
    columns: List[str]
    row_count: int
    is_empty: bool = False
    warning: Optional[str] = None


@dataclass
class FileData:
    """Represents an uploaded file and all its worksheets."""
    
    file_name: str
    file_type: str  # 'xlsx' or 'csv'
    file_size: int
    sheets: Dict[str, SheetData] = field(default_factory=dict)
    error: Optional[str] = None

    @property
    def sheet_names(self) -> List[str]:
        return list(self.sheets.keys())


class FileReader:
    """Safely reads XLSX and CSV files with encoding and delimiter auto-detection."""

    SUPPORTED_ENCODINGS = ["utf-8-sig", "utf-8", "latin-1", "cp1252", "iso-8859-1"]
    COMMON_DELIMITERS = [",", ";", "\t", "|"]

    @classmethod
    def read_file(
        cls,
        file_source: Union[bytes, io.BytesIO, str, Path],
        file_name: str,
    ) -> FileData:
        """Read an uploaded file (XLSX or CSV) and return a structured FileData object."""
        safe_name = get_safe_filename(file_name)
        ext = Path(safe_name).suffix.lower()

        # Get raw bytes
        if isinstance(file_source, bytes):
            raw_bytes = file_source
        elif isinstance(file_source, (io.BytesIO, io.BufferedReader)):
            file_source.seek(0)
            raw_bytes = file_source.read()
        elif isinstance(file_source, (str, Path)):
            with open(file_source, "rb") as f:
                raw_bytes = f.read()
        elif hasattr(file_source, "getvalue"):
            raw_bytes = file_source.getvalue()
        elif hasattr(file_source, "read"):
            if hasattr(file_source, "seek"):
                file_source.seek(0)
            raw_bytes = file_source.read()
        else:
            return FileData(
                file_name=safe_name,
                file_type="unknown",
                file_size=0,
                error="Unsupported file input stream.",
            )

        file_size = len(raw_bytes)
        if file_size == 0:
            return FileData(
                file_name=safe_name,
                file_type=ext.replace(".", ""),
                file_size=0,
                error="Uploaded file is completely empty (0 bytes).",
            )

        if ext == ".xlsx":
            return cls._read_xlsx(raw_bytes, safe_name, file_size)
        elif ext == ".csv":
            return cls._read_csv(raw_bytes, safe_name, file_size)
        else:
            return FileData(
                file_name=safe_name,
                file_type=ext.replace(".", ""),
                file_size=file_size,
                error=f"Unsupported file format '{ext}'. Only .xlsx and .csv files are supported.",
            )

    @classmethod
    def _read_xlsx(cls, raw_bytes: bytes, file_name: str, file_size: int) -> FileData:
        """Read all sheets from an Excel workbook."""
        file_data = FileData(file_name=file_name, file_type="xlsx", file_size=file_size)
        
        try:
            # First inspect sheet names using openpyxl in read-only mode to prevent high memory usage
            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
        except Exception as e:
            file_data.error = f"Corrupted or invalid Excel file: {str(e)}"
            return file_data

        if not sheet_names:
            file_data.error = "Excel workbook contains no worksheets."
            return file_data

        for sheet_name in sheet_names:
            try:
                # Read worksheet with pandas
                # Preserve date parsing by openpyxl/pandas
                df = pd.read_excel(
                    io.BytesIO(raw_bytes),
                    sheet_name=sheet_name,
                    engine="openpyxl",
                )
                
                # Check if worksheet is empty
                if df.empty or len(df.columns) == 0:
                    file_data.sheets[sheet_name] = SheetData(
                        sheet_name=sheet_name,
                        df=pd.DataFrame(),
                        columns=[],
                        row_count=0,
                        is_empty=True,
                        warning=f"Worksheet '{sheet_name}' is empty.",
                    )
                    continue

                # Ensure all column headers are strings and strip surrounding whitespace
                df.columns = [str(col).strip() if col is not None else f"Unnamed_{i}" for i, col in enumerate(df.columns)]
                
                file_data.sheets[sheet_name] = SheetData(
                    sheet_name=sheet_name,
                    df=df,
                    columns=list(df.columns),
                    row_count=len(df),
                    is_empty=False,
                )
            except Exception as e:
                file_data.sheets[sheet_name] = SheetData(
                    sheet_name=sheet_name,
                    df=pd.DataFrame(),
                    columns=[],
                    row_count=0,
                    is_empty=True,
                    warning=f"Error reading worksheet '{sheet_name}': {str(e)}",
                )

        return file_data

    @classmethod
    def _read_csv(cls, raw_bytes: bytes, file_name: str, file_size: int) -> FileData:
        """Read a CSV file with automatic encoding and delimiter detection."""
        file_data = FileData(file_name=file_name, file_type="csv", file_size=file_size)
        
        # Try decoding with supported encodings
        decoded_text: Optional[str] = None
        successful_encoding: Optional[str] = None

        for enc in cls.SUPPORTED_ENCODINGS:
            try:
                decoded_text = raw_bytes.decode(enc)
                successful_encoding = enc
                break
            except UnicodeDecodeError:
                continue

        if decoded_text is None:
            file_data.error = (
                "Unable to decode CSV file. The file character encoding could not be determined. "
                "Please save the file using UTF-8 encoding and try again."
            )
            return file_data

        if not decoded_text.strip():
            file_data.error = "CSV file is empty (contains no text or data)."
            return file_data

        # Detect delimiter using csv.Sniffer and heuristic verification
        delimiter = cls._detect_delimiter(decoded_text)

        try:
            df = pd.read_csv(
                io.StringIO(decoded_text),
                sep=delimiter,
                encoding=successful_encoding,
                on_bad_lines="skip",
            )
            
            if df.empty or len(df.columns) == 0:
                file_data.sheets["Sheet1"] = SheetData(
                    sheet_name="Sheet1",
                    df=pd.DataFrame(),
                    columns=[],
                    row_count=0,
                    is_empty=True,
                    warning="CSV file contains no data rows.",
                )
                return file_data

            # Normalize column names to strings
            df.columns = [str(col).strip() if col is not None else f"Unnamed_{i}" for i, col in enumerate(df.columns)]

            file_data.sheets["Sheet1"] = SheetData(
                sheet_name="Sheet1",
                df=df,
                columns=list(df.columns),
                row_count=len(df),
                is_empty=False,
            )
        except Exception as e:
            file_data.error = f"Failed to parse CSV file: {str(e)}"

        return file_data

    @classmethod
    def _detect_delimiter(cls, text: str) -> str:
        """Heuristically and dynamically detect the CSV delimiter."""
        sample = text[:16384]  # First 16KB
        
        # First try csv.Sniffer
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
            if dialect.delimiter in cls.COMMON_DELIMITERS:
                return dialect.delimiter
        except Exception:
            pass

        # Fallback heuristic: count occurrences on the first non-empty line
        first_line = ""
        for line in sample.splitlines():
            if line.strip():
                first_line = line
                break

        if not first_line:
            return ","

        delimiter_counts = {sep: first_line.count(sep) for sep in cls.COMMON_DELIMITERS}
        best_delimiter = max(delimiter_counts, key=delimiter_counts.get)
        
        return best_delimiter if delimiter_counts[best_delimiter] > 0 else ","
