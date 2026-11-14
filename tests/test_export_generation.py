"""Unit tests for Excel and TXT report generation."""
import datetime
import io
import openpyxl
import pytest

from modules.export_generator import ExportGenerator
from modules.filter_engine import Record
from utils.error_handler import ErrorCategory, ProcessingError
from utils.file_utils import sanitize_cell_for_export


class TestExportGeneration:
    """Test suite for Excel workbook styling, sheets, formula sanitization, and TXT report formats."""

    REF_DATE = datetime.date(2026, 11, 13)

    def test_excel_export_structure_and_sheets(self):
        records = [
            Record(
                name="Michael Smith",
                last_updated=datetime.date(2026, 11, 7),
                days_difference=6,
                source_file="file1.xlsx",
                source_sheet="Sheet1",
                row_number=2,
            ),
            Record(
                name="Sarah Johnson",
                last_updated=datetime.date(2026, 11, 1),
                days_difference=12,
                source_file="file2.csv",
                source_sheet="Sheet1",
                row_number=5,
            ),
        ]

        errors = [
            ProcessingError(
                source_file="file1.xlsx",
                source_sheet="Sheet1",
                row_number=3,
                error_category=ErrorCategory.INVALID_DATE,
                explanation="Failed to parse date value",
                original_value="invalid-date",
            )
        ]

        stats = {
            "processing_date": self.REF_DATE,
            "threshold_days": 5,
            "timezone": "Asia/Kolkata",
            "total_files": 2,
            "total_sheets": 2,
            "total_examined": 5,
            "total_matching": 2,
            "total_skipped": 2,
            "total_errors": 1,
            "invalid_date_count": 1,
            "missing_date_count": 0,
            "missing_name_count": 0,
            "future_date_count": 0,
        }

        excel_bytes = ExportGenerator.generate_excel_bytes(records, errors, stats)
        assert len(excel_bytes) > 0

        # Reload workbook with openpyxl to verify valid Excel format
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        assert "Filtered Records" in wb.sheetnames

        # Check Filtered Records rows
        ws_match = wb["Filtered Records"]
        assert ws_match.max_row == 3  # Header + 2 data rows
        assert ws_match["A2"].value == "Michael Smith"
        assert ws_match["B2"].value == "07-Nov-2026"
        assert ws_match["C2"].value == 6

    def test_excel_formula_injection_sanitization(self):
        # Untrusted strings starting with formula prefixes
        assert sanitize_cell_for_export("=cmd|' /C calc'!A0") == "'=cmd|' /C calc'!A0"
        assert sanitize_cell_for_export("+12345text") == "'+12345text"
        assert sanitize_cell_for_export("@SUM(1,2)") == "'@SUM(1,2)"
        # Normal negative number should NOT be quoted
        assert sanitize_cell_for_export("-5") == "-5"
        # Normal strings shouldn't be touched
        assert sanitize_cell_for_export("Normal Name") == "Normal Name"

    def test_txt_report_format(self):
        records = [
            Record(
                name="Michael Smith",
                last_updated=datetime.date(2026, 11, 7),
                days_difference=6,
                source_file="file1.xlsx",
                source_sheet="Sheet1",
                row_number=2,
            ),
        ]

        txt = ExportGenerator.generate_txt_report(
            matching_records=records,
            processing_date=self.REF_DATE,
            threshold_days=5,
        )

        assert "DataPulse - RESULTS" in txt
        assert "Filtering Threshold: More than 5 days" in txt
        assert "Name: Michael Smith" in txt
        assert "Days Difference: 6" in txt
        assert "Total Matching Records: 1" in txt

    def test_txt_report_empty_state(self):
        txt = ExportGenerator.generate_txt_report(
            matching_records=[],
            processing_date=self.REF_DATE,
            threshold_days=5,
        )
        assert "No records satisfied the filtering condition." in txt
        assert "Total Matching Records: 0" in txt
