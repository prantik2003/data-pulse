"""Integration tests for file ingestion, multi-sheet handling, and processing workflows."""
import datetime
import io
import pandas as pd
import pytest

from modules.file_reader import FileReader
from modules.processing_manager import ProcessingManager


class TestFileProcessingIntegration:
    """End-to-end integration test suite across XLSX, CSV, and multi-sheet structures."""

    REF_DATE = datetime.date(2026, 11, 13)

    def _create_xlsx_bytes(self, sheets_dict: dict[str, pd.DataFrame]) -> bytes:
        """Helper to create multi-sheet Excel bytes in memory."""
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            for s_name, df in sheets_dict.items():
                df.to_excel(writer, sheet_name=s_name, index=False)
        return out.getvalue()

    def test_1_upload_one_xlsx_file(self):
        df = pd.DataFrame({
            "Employee Name": ["Alice", "Bob"],
            "Last Updated": ["2026-11-01", "2026-11-12"],
        })
        raw_xlsx = self._create_xlsx_bytes({"Staff": df})
        file_data = FileReader.read_file(raw_xlsx, "staff.xlsx")

        assert file_data.error is None
        assert "Staff" in file_data.sheets
        assert file_data.sheets["Staff"].row_count == 2

    def test_2_upload_one_csv_file_with_semicolon(self):
        csv_content = "Resource;Modified On;Department\nCharlie;2026-10-25;Engineering\nDavid;2026-11-11;Design\n"
        file_data = FileReader.read_file(csv_content.encode("utf-8"), "team.csv")

        assert file_data.error is None
        assert "Sheet1" in file_data.sheets
        assert file_data.sheets["Sheet1"].row_count == 2
        assert "Resource" in file_data.sheets["Sheet1"].columns

    def test_3_multiple_worksheets_processed_independently(self):
        sheet1_df = pd.DataFrame({
            "Name": ["John", "Alex"],
            "Last Updated": ["2026-11-01", "2026-11-12"],  # John matches (12d), Alex skipped (1d)
        })
        sheet2_df = pd.DataFrame({
            "Assigned To": ["Sarah", "David"],
            "Modification Timestamp": ["2026-10-20", "2026-11-08"],  # Sarah matches (24d), David skipped (5d)
        })

        raw_xlsx = self._create_xlsx_bytes({"Engineering": sheet1_df, "Design": sheet2_df})
        file_data = FileReader.read_file(raw_xlsx, "company.xlsx")
        configs = ProcessingManager.analyze_uploaded_files([file_data])

        # Verify independent column detection
        assert configs["company.xlsx"]["Engineering"].name_col == "Name"
        assert configs["company.xlsx"]["Engineering"].date_col == "Last Updated"
        assert configs["company.xlsx"]["Design"].name_col == "Assigned To"
        assert configs["company.xlsx"]["Design"].date_col == "Modification Timestamp"

        res = ProcessingManager.execute_processing(
            files_data=[file_data],
            sheet_configs=configs,
            threshold_days=5,
            current_date=self.REF_DATE,
        )

        assert res.stats["total_sheets"] == 2
        assert res.stats["total_examined"] == 4
        assert res.stats["total_matching"] == 2
        assert res.stats["total_skipped"] == 2

        matching_names = [r.name for r in res.matching_records]
        assert "John" in matching_names
        assert "Sarah" in matching_names

    def test_4_mixed_xlsx_and_csv_files_with_different_headers_and_date_formats(self):
        # File A: XLSX with DD/MM/YYYY
        df_a = pd.DataFrame({
            "Employee": ["Michael", "Dwight"],
            "Updated At": ["07/11/2026", "12/11/2026"],  # 6d (match), 1d (skip)
        })
        raw_a = self._create_xlsx_bytes({"Roster": df_a})
        file_a = FileReader.read_file(raw_a, "file_a.xlsx")

        # File B: CSV with DD-MMM-YYYY and tab delimiter
        csv_b = "User Name\tLast Modified\nJim\t20-Oct-2026\nPam\t10-Nov-2026\n"
        file_b = FileReader.read_file(csv_b.encode("utf-8"), "file_b.csv")

        files = [file_a, file_b]
        configs = ProcessingManager.analyze_uploaded_files(files)

        res = ProcessingManager.execute_processing(
            files_data=files,
            sheet_configs=configs,
            threshold_days=5,
            current_date=self.REF_DATE,
        )

        assert res.stats["total_files"] == 2
        assert res.stats["total_sheets"] == 2
        assert res.stats["total_examined"] == 4
        assert res.stats["total_matching"] == 2  # Michael (6d), Jim (24d)
        assert res.stats["total_skipped"] == 2   # Dwight (1d), Pam (3d)

    def test_5_file_with_zero_matching_records(self):
        df = pd.DataFrame({
            "Name": ["Recent User 1", "Recent User 2"],
            "Last Updated": ["2026-11-12", "2026-11-11"],  # Both <= 5 days
        })
        raw = self._create_xlsx_bytes({"Recent": df})
        file_data = FileReader.read_file(raw, "recent.xlsx")
        configs = ProcessingManager.analyze_uploaded_files([file_data])

        res = ProcessingManager.execute_processing(
            files_data=[file_data],
            sheet_configs=configs,
            threshold_days=5,
            current_date=self.REF_DATE,
        )

        assert res.stats["total_matching"] == 0
        assert res.stats["total_skipped"] == 2
        assert res.excel_bytes is not None
        assert "No records satisfied the filtering condition" in res.txt_report

    def test_6_corrupted_or_unsupported_file(self):
        # Corrupted xlsx (random garbage bytes)
        file_corrupt = FileReader.read_file(b"PK\x03\x04not_a_valid_zip", "corrupt.xlsx")
        assert file_corrupt.error is not None

        # Unsupported format (.pdf)
        file_pdf = FileReader.read_file(b"%PDF-1.4...", "doc.pdf")
        assert file_pdf.error is not None
        assert "Unsupported file format" in file_pdf.error
