"""Unit tests for intelligent column detection and header normalization."""
import pandas as pd
import pytest

from modules.column_detector import ColumnDetector
from utils.header_normalizer import normalize_header


class TestHeaderNormalization:
    """Test suite for header normalization functionality."""

    def test_exact_and_lowercase(self):
        assert normalize_header("name") == "name"
        assert normalize_header("Name") == "name"
        assert normalize_header("NAME") == "name"

    def test_underscore_and_hyphen(self):
        assert normalize_header("Employee_Name") == "employee name"
        assert normalize_header("EMPLOYEE-NAME") == "employee name"
        assert normalize_header("last_updated_date") == "last updated date"

    def test_camel_case(self):
        assert normalize_header("employeeName") == "employee name"
        assert normalize_header("lastUpdatedDate") == "last updated date"
        assert normalize_header("UserFullName") == "user full name"

    def test_special_characters_and_spaces(self):
        assert normalize_header("  Last Updated (UTC)  ") == "last updated utc"
        assert normalize_header("Employee/Resource #") == "employee resource"
        assert normalize_header("Date.Modified: [Archive]") == "date modified archive"


class TestColumnDetection:
    """Test suite for header matching, fuzzy ranking, and content validation."""

    def test_1_exact_header_match(self):
        df = pd.DataFrame({
            "Name": ["John Doe", "Jane Smith"],
            "Last Updated": ["2026-11-01", "2026-11-02"],
            "Age": [30, 25],
        })
        res = ColumnDetector.detect_columns(df)
        assert res.detected_name_col == "Name"
        assert res.detected_date_col == "Last Updated"
        assert res.name_status == "high"
        assert res.date_status == "high"

    def test_2_different_capitalization(self):
        df = pd.DataFrame({
            "EMPLOYEE NAME": ["Alice", "Bob"],
            "LAST MODIFIED": ["2026-10-15", "2026-10-20"],
        })
        res = ColumnDetector.detect_columns(df)
        assert res.detected_name_col == "EMPLOYEE NAME"
        assert res.detected_date_col == "LAST MODIFIED"
        assert res.name_status == "high"
        assert res.date_status == "high"

    def test_3_underscore_and_hyphen_variations(self):
        df = pd.DataFrame({
            "User_Name": ["Charlie", "Diana"],
            "modified-date": ["2026-09-01", "2026-09-05"],
        })
        res = ColumnDetector.detect_columns(df)
        assert res.detected_name_col == "User_Name"
        assert res.detected_date_col == "modified-date"

    def test_4_camel_case_headings(self):
        df = pd.DataFrame({
            "employeeName": ["Edward", "Fiona"],
            "lastActivityDate": ["2026-11-05", "2026-11-08"],
        })
        res = ColumnDetector.detect_columns(df)
        assert res.detected_name_col == "employeeName"
        assert res.detected_date_col == "lastActivityDate"

    def test_5_fuzzy_matching(self):
        df = pd.DataFrame({
            "Staff Member Full Name": ["George", "Hannah"],
            "Date of Latest Update": ["2026-11-01", "2026-11-02"],
        })
        res = ColumnDetector.detect_columns(df)
        assert res.detected_name_col == "Staff Member Full Name"
        assert res.detected_date_col == "Date of Latest Update"

    def test_6_different_column_headings_across_separate_datasets(self):
        # Dataset 1: Resource & Modified On
        df1 = pd.DataFrame({
            "Resource": ["Ian", "Jack"],
            "Modified On": ["2026-10-01", "2026-10-02"],
        })
        res1 = ColumnDetector.detect_columns(df1)
        assert res1.detected_name_col == "Resource"
        assert res1.detected_date_col == "Modified On"

        # Dataset 2: Assigned To & Modification Timestamp
        df2 = pd.DataFrame({
            "Assigned To": ["Kevin", "Laura"],
            "Modification Timestamp": ["2026-10-05T10:00:00", "2026-10-06T12:00:00"],
        })
        res2 = ColumnDetector.detect_columns(df2)
        assert res2.detected_name_col == "Assigned To"
        assert res2.detected_date_col == "Modification Timestamp"

    def test_7_multiple_possible_name_columns_disfavors_company_or_file(self):
        df = pd.DataFrame({
            "Company Name": ["Acme Corp", "Globex Inc"],
            "Employee Name": ["Michael Scott", "Jim Halpert"],
            "Updated At": ["2026-11-01", "2026-11-02"],
        })
        res = ColumnDetector.detect_columns(df)
        # Should pick Employee Name over Company Name
        assert res.detected_name_col == "Employee Name"

    def test_8_multiple_possible_date_columns_prioritizes_updated_over_created(self):
        df = pd.DataFrame({
            "Name": ["Pam Beesly", "Dwight Schrute"],
            "Created Date": ["2025-01-01", "2025-01-02"],
            "Last Updated": ["2026-11-01", "2026-11-02"],
        })
        res = ColumnDetector.detect_columns(df)
        # Should pick Last Updated over Created Date
        assert res.detected_date_col == "Last Updated"

    def test_9_missing_name_column(self):
        df = pd.DataFrame({
            "Department": ["Sales", "Accounting"],
            "Amount": [1000, 2500],
            "Last Updated": ["2026-11-01", "2026-11-02"],
        })
        res = ColumnDetector.detect_columns(df)
        assert res.name_status in ("low", "unmatched")
        assert res.needs_confirmation is True

    def test_10_missing_date_column(self):
        df = pd.DataFrame({
            "Employee Name": ["Angela", "Oscar"],
            "Role": ["Accountant", "Accountant"],
            "Salary": [50000, 52000],
        })
        res = ColumnDetector.detect_columns(df)
        assert res.date_status in ("low", "unmatched")
        assert res.needs_confirmation is True

    def test_11_incorrect_candidate_rejected_by_content_validation(self):
        # Column named 'Date' but containing numeric customer IDs
        # Another column has actual dates
        df = pd.DataFrame({
            "User Name": ["Stanley", "Phyllis"],
            "Date": [1001, 1002],  # Fake date column with numeric IDs
            "Record Updated At": ["2026-11-01", "2026-11-02"],
        })
        res = ColumnDetector.detect_columns(df)
        assert res.detected_date_col == "Record Updated At"
