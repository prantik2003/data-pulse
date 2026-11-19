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

        # Check Filtered Records headers and rows
        ws_match = wb["Filtered Records"]
        assert ws_match.max_row == 4  # Banner + Header + 2 data rows
        banner = ws_match.cell(row=1, column=1).value
        assert "OTHER MATCHING RECORDS (2 Records)" in banner
        headers = [ws_match.cell(row=2, column=c).value for c in range(1, 4)]
        assert headers == ["Name", "Last Updated", "Source File"]
        assert ws_match["A3"].value == "Michael Smith"
        assert ws_match["B3"].value == "07-Nov-2026"
        assert ws_match["C3"].value == "file1.xlsx"

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

        assert "DataPulse - CONSOLIDATED REPORT" in txt
        assert "Threshold: More than 5 calendar days" in txt
        assert "OTHER MATCHING RECORDS (Total: 1)" in txt
        assert "Michael Smith" in txt
        assert "file1.xlsx" in txt
        assert "Total Matching Records: 1" in txt

    def test_txt_report_empty_state(self):
        txt = ExportGenerator.generate_txt_report(
            matching_records=[],
            processing_date=self.REF_DATE,
            threshold_days=5,
        )
        assert "No records satisfied the filtering condition." in txt
        assert "Total Matching Records: 0" in txt

    def test_ticket_columns_export_and_txt(self):
        records = [
            Record(
                name="Sambit Chakraborty",
                last_updated=datetime.date(2026, 11, 3),
                days_difference=11,
                source_file="incident _2_.xlsx",
                source_sheet="Page 1",
                row_number=2,
                incident="INC27334682",
                sctask="",
                short_description="Environmental reports do not work",
            ),
            Record(
                name="Tannu Kumari",
                last_updated=datetime.date(2026, 10, 12),
                days_difference=33,
                source_file="sc_task _1_.xlsx",
                source_sheet="Page 1",
                row_number=3,
                incident="",
                sctask="SCTASK19669063",
                short_description="Create user account",
            ),
        ]
        excel_bytes = ExportGenerator.generate_excel_bytes(records)
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        ws = wb["Filtered Records"]

        # Section 1: INCIDENTS
        assert ws["A1"].value == "INCIDENTS (1 Records)"
        inc_headers = [ws.cell(row=2, column=c).value for c in range(1, 6)]
        assert inc_headers == ["Incident", "Short Description", "Name", "Last Updated", "Source File"]
        assert ws["A3"].value == "INC27334682"
        assert ws["B3"].value == "Environmental reports do not work"
        assert ws["C3"].value == "Sambit Chakraborty"
        assert ws["D3"].value == "03-Nov-2026"
        assert ws["E3"].value == "incident _2_.xlsx"

        # Section 2: SCTASKS (separated by blank row 4)
        assert ws["A5"].value == "SERVICE CATALOG TASKS (SCTASKS) (1 Records)"
        sctask_headers = [ws.cell(row=6, column=c).value for c in range(1, 6)]
        assert sctask_headers == ["SCTask", "Short Description", "Name", "Last Updated", "Source File"]
        assert ws["A7"].value == "SCTASK19669063"
        assert ws["B7"].value == "Create user account"
        assert ws["C7"].value == "Tannu Kumari"
        assert ws["D7"].value == "12-Oct-2026"
        assert ws["E7"].value == "sc_task _1_.xlsx"

        txt = ExportGenerator.generate_txt_report(records, self.REF_DATE, 5)
        assert "INCIDENTS (Total: 1)" in txt
        assert "INC27334682" in txt
        assert "Environmental reports do not work" in txt
        assert "SERVICE CATALOG TASKS (SCTASKS) (Total: 1)" in txt
        assert "SCTASK19669063" in txt
        assert "Create user account" in txt
        assert "+-------------+" in txt or "+-" in txt  # ASCII table separator

    def test_txt_report_alignment_and_wrapping(self):
        records = [
            Record(
                name="Sambit Chakraborty",
                last_updated=datetime.date(2026, 11, 3),
                days_difference=11,
                source_file="incident _2_.xlsx",
                source_sheet="Page 1",
                row_number=2,
                incident="INC27334682",
                short_description="Environmental reports \tdo not work with \nthe monitoring analyst role.",
            ),
            Record(
                name="Tannu Kumari",
                last_updated=datetime.date(2026, 10, 12),
                days_difference=33,
                source_file="sc_task _1_.xlsx",
                source_sheet="Page 1",
                row_number=3,
                sctask="SCTASK19481771",
                short_description="GLIMS Labvantage 8.9 Upgrade activities",
            ),
        ]
        txt = ExportGenerator.generate_txt_report(records, self.REF_DATE, 5)
        assert "\r\n" in txt
        lines = txt.split("\r\n")

        # Find tables and check that every line in a table has the exact same character length
        inc_lines = [l for l in lines if l.startswith("|") or l.startswith("+")]
        assert len(inc_lines) > 0
        table_width = len(inc_lines[0])
        for l in inc_lines:
            assert len(l) == table_width, f"Mismatched line length in table: expected {table_width}, got {len(l)} for '{l}'"

        # Tabs and embedded newlines should be cleaned up
        assert "\t" not in txt

    def test_rtf_report_generation(self):
        records = [
            Record(
                name="Sambit Chakraborty",
                last_updated=datetime.date(2026, 11, 3),
                days_difference=11,
                source_file="incident _2_.xlsx",
                source_sheet="Page 1",
                row_number=2,
                incident="INC27334682",
                short_description="Environmental reports do not work with analyst role.",
            ),
            Record(
                name="Tannu Kumari",
                last_updated=datetime.date(2026, 10, 12),
                days_difference=33,
                source_file="sc_task _1_.xlsx",
                source_sheet="Page 1",
                row_number=3,
                sctask="SCTASK19481771",
                short_description="GLIMS Labvantage 8.9 Upgrade activities",
            ),
        ]
        rtf = ExportGenerator.generate_rtf_report(records, self.REF_DATE, 5)
        assert rtf.startswith(r"{\rtf1")
        assert rtf.endswith("}")
        assert "INCIDENTS" in rtf
        assert "INC27334682" in rtf
        assert "SERVICE CATALOG TASKS (SCTASKS)" in rtf
        assert "SCTASK19481771" in rtf
        assert r"\trowd" in rtf
        assert r"\cell" in rtf
        assert r"\row" in rtf

    def test_rtf_report_empty_state(self):
        rtf = ExportGenerator.generate_rtf_report([], self.REF_DATE, 5)
        assert rtf.startswith(r"{\rtf1")
        assert "No records satisfied the filtering condition." in rtf
        assert "Total Matching Records: 0" in rtf


