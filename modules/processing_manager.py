"""Processing manager orchestrating multi-file, multi-sheet workflows."""
import datetime
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd

from config.settings import Settings
from modules.column_detector import ColumnDetectionResult, ColumnDetector
from modules.date_parser import DateParser
from modules.export_generator import ExportGenerator
from modules.file_reader import FileData, FileReader, SheetData
from modules.filter_engine import FilterEngine, FilterResult, Record
from utils.date_utils import get_current_date
from utils.error_handler import ErrorCategory, ProcessingError


@dataclass
class SheetConfig:
    """Configuration mapping for an individual sheet to be processed."""
    
    file_name: str
    sheet_name: str
    name_col: str
    date_col: str
    is_active: bool = True
    detection_result: Optional[ColumnDetectionResult] = None


@dataclass
class ProcessingResult:
    """Full aggregated outcome across all uploaded files and worksheets."""
    
    matching_records: List[Record] = field(default_factory=list)
    skipped_records: List[Record] = field(default_factory=list)
    errors: List[ProcessingError] = field(default_factory=list)
    sheet_breakdown: List[Dict[str, Any]] = field(default_factory=list)
    stats: Dict[str, Any] = field(default_factory=dict)
    
    excel_bytes: Optional[bytes] = None
    txt_report: Optional[str] = None


class ProcessingManager:
    """Coordinates file ingestion, column detection, filtering, and reporting."""

    @classmethod
    def analyze_uploaded_files(
        cls,
        files_data: List[FileData],
        sample_size: int = 50,
    ) -> Dict[str, Dict[str, SheetConfig]]:
        """Run intelligent column detection across all files and worksheets.
        
        Returns:
            Nested dict: {file_name: {sheet_name: SheetConfig}}
        """
        file_configs: Dict[str, Dict[str, SheetConfig]] = {}

        for f_data in files_data:
            if f_data.error:
                continue
            
            file_configs[f_data.file_name] = {}
            for s_name, s_data in f_data.sheets.items():
                if s_data.is_empty or len(s_data.columns) == 0:
                    continue

                detection = ColumnDetector.detect_columns(s_data.df, sample_size=sample_size)
                
                # Pick top column or first available column as fallback
                fallback_name = detection.detected_name_col or (s_data.columns[0] if s_data.columns else "")
                fallback_date = detection.detected_date_col or (
                    s_data.columns[1] if len(s_data.columns) > 1 else (s_data.columns[0] if s_data.columns else "")
                )

                file_configs[f_data.file_name][s_name] = SheetConfig(
                    file_name=f_data.file_name,
                    sheet_name=s_name,
                    name_col=fallback_name,
                    date_col=fallback_date,
                    is_active=True,
                    detection_result=detection,
                )

        return file_configs

    @classmethod
    def execute_processing(
        cls,
        files_data: List[FileData],
        sheet_configs: Dict[str, Dict[str, SheetConfig]],
        threshold_days: int = 5,
        current_date: Optional[datetime.date] = None,
        date_preference: str = "Auto-detect",
        timezone_name: str = "Asia/Kolkata",
    ) -> ProcessingResult:
        """Run complete extraction, filtering, and export generation across configured sheets."""
        if current_date is None:
            current_date = get_current_date(timezone_name)

        all_matching: List[Record] = []
        all_skipped: List[Record] = []
        all_errors: List[ProcessingError] = []
        sheet_breakdown: List[Dict[str, Any]] = []

        total_files_processed = 0
        total_sheets_processed = 0
        total_examined = 0

        # Create quick lookup for files
        file_map = {f.file_name: f for f in files_data}

        for f_name, f_data in file_map.items():
            if f_data.error:
                all_errors.append(
                    ProcessingError(
                        source_file=f_name,
                        source_sheet="N/A",
                        row_number=None,
                        error_category=ErrorCategory.CORRUPTED_FILE,
                        explanation=f_data.error,
                    )
                )
                continue

            sheets_for_file = sheet_configs.get(f_name, {})
            file_had_active_sheet = False

            for s_name, s_config in sheets_for_file.items():
                if not s_config.is_active:
                    continue

                s_data = f_data.sheets.get(s_name)
                if not s_data or s_data.is_empty:
                    if s_data and s_data.warning:
                        all_errors.append(
                            ProcessingError(
                                source_file=f_name,
                                source_sheet=s_name,
                                row_number=None,
                                error_category=ErrorCategory.EMPTY_SHEET,
                                explanation=s_data.warning,
                            )
                        )
                    continue

                file_had_active_sheet = True
                total_sheets_processed += 1

                # Validate column selections exist in sheet
                if s_config.name_col not in s_data.columns or s_config.date_col not in s_data.columns:
                    all_errors.append(
                        ProcessingError(
                            source_file=f_name,
                            source_sheet=s_name,
                            row_number=None,
                            error_category=ErrorCategory.MAPPING_ERROR,
                            explanation=f"Selected columns '{s_config.name_col}' or '{s_config.date_col}' not found in sheet.",
                        )
                    )
                    continue

                # Infer day first pattern from sample values of the date column
                date_sample = list(s_data.df[s_config.date_col].dropna().head(40))
                day_first_hint = DateParser.infer_column_day_first(date_sample)

                # Convert rows to dictionaries for processing
                rows_data = s_data.df.to_dict(orient="records")
                total_examined += len(rows_data)

                # Apply filtering
                sheet_res = FilterEngine.filter_records(
                    rows_data=rows_data,
                    name_col=s_config.name_col,
                    date_col=s_config.date_col,
                    source_file=f_name,
                    source_sheet=s_name,
                    current_date=current_date,
                    threshold_days=threshold_days,
                    date_preference=date_preference,
                    timezone_name=timezone_name,
                    column_context_day_first=day_first_hint,
                )

                all_matching.extend(sheet_res.matching_records)
                all_skipped.extend(sheet_res.skipped_records)
                all_errors.extend(sheet_res.errors)

                sheet_breakdown.append({
                    "File": f_name,
                    "Sheet": s_name,
                    "Name Column": s_config.name_col,
                    "Date Column": s_config.date_col,
                    "Examined": sheet_res.total_examined,
                    "Matching (> Threshold)": sheet_res.total_matching,
                    "Skipped (<= Threshold)": sheet_res.total_skipped,
                    "Errors": sheet_res.total_errors,
                })

            if file_had_active_sheet:
                total_files_processed += 1

        # Calculate error categories
        invalid_date_count = sum(1 for e in all_errors if e.error_category == ErrorCategory.INVALID_DATE)
        missing_date_count = sum(1 for e in all_errors if e.error_category == ErrorCategory.MISSING_DATE)
        missing_name_count = sum(1 for e in all_errors if e.error_category == ErrorCategory.MISSING_NAME)
        future_date_count = sum(1 for e in all_errors if e.error_category == ErrorCategory.FUTURE_DATE)

        stats = {
            "processing_date": current_date,
            "threshold_days": threshold_days,
            "timezone": timezone_name,
            "total_files": total_files_processed,
            "total_sheets": total_sheets_processed,
            "total_examined": total_examined,
            "total_matching": len(all_matching),
            "total_skipped": len(all_skipped),
            "total_errors": len(all_errors),
            "invalid_date_count": invalid_date_count,
            "missing_date_count": missing_date_count,
            "missing_name_count": missing_name_count,
            "future_date_count": future_date_count,
        }

        # Generate report bytes
        excel_bytes = ExportGenerator.generate_excel_bytes(
            matching_records=all_matching,
            errors=all_errors,
            summary_stats=stats,
        )

        txt_report = ExportGenerator.generate_txt_report(
            matching_records=all_matching,
            processing_date=current_date,
            threshold_days=threshold_days,
        )

        return ProcessingResult(
            matching_records=all_matching,
            skipped_records=all_skipped,
            errors=all_errors,
            sheet_breakdown=sheet_breakdown,
            stats=stats,
            excel_bytes=excel_bytes,
            txt_report=txt_report,
        )
