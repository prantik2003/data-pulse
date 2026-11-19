"""Unit tests for the filtering engine and date difference calculations."""
import datetime
import pytest

from modules.filter_engine import FilterEngine
from utils.date_utils import calculate_calendar_day_difference, to_calendar_date
from utils.error_handler import ErrorCategory


class TestFilteringLogic:
    """Test suite validating calendar-day aging logic, boundary conditions, and edge cases."""

    # Reference date from prompt example: 13-Nov-2026
    REF_DATE = datetime.date(2026, 11, 13)

    def test_calendar_day_difference_calculation(self):
        # 13 Nov 2026 - 12 Nov 2026 = 1
        assert calculate_calendar_day_difference(datetime.date(2026, 11, 12), self.REF_DATE) == 1
        # 13 Nov 2026 - 08 Nov 2026 = 5
        assert calculate_calendar_day_difference(datetime.date(2026, 11, 8), self.REF_DATE) == 5
        # 13 Nov 2026 - 07 Nov 2026 = 6
        assert calculate_calendar_day_difference(datetime.date(2026, 11, 7), self.REF_DATE) == 6
        # Future date: 14 Nov 2026
        assert calculate_calendar_day_difference(datetime.date(2026, 11, 14), self.REF_DATE) == -1

    def test_example_scenario_from_spec(self):
        """Test the exact specification example from Section 6.
        Assuming current date is 13 November 2026:
        - John: 12-Nov-2026, Diff: 1 -> Exclude
        - Alex: 10-Nov-2026, Diff: 3 -> Exclude
        - David: 08-Nov-2026, Diff: 5 -> Exclude (STRICTLY > 5)
        - Michael: 07-Nov-2026, Diff: 6 -> Include
        - Sarah: 01-Nov-2026, Diff: 12 -> Include
        - Robert: 20-Oct-2026, Diff: 24 -> Include
        """
        rows = [
            {"Name": "John", "Last Updated": "12-Nov-2026"},
            {"Name": "Alex", "Last Updated": "10-Nov-2026"},
            {"Name": "David", "Last Updated": "08-Nov-2026"},
            {"Name": "Michael", "Last Updated": "07-Nov-2026"},
            {"Name": "Sarah", "Last Updated": "01-Nov-2026"},
            {"Name": "Robert", "Last Updated": "20-Oct-2026"},
        ]

        res = FilterEngine.filter_records(
            rows_data=rows,
            name_col="Name",
            date_col="Last Updated",
            source_file="spec_example.xlsx",
            source_sheet="Sheet1",
            current_date=self.REF_DATE,
            threshold_days=5,
        )

        assert res.total_examined == 6
        assert res.total_matching == 3
        assert res.total_skipped == 3
        assert res.total_errors == 0

        matching_names = [r.name for r in res.matching_records]
        assert matching_names == ["Michael", "Sarah", "Robert"]

        skipped_names = [r.name for r in res.skipped_records]
        assert skipped_names == ["John", "Alex", "David"]

    def test_exact_5_days_difference_is_excluded(self):
        # 08-Nov-2026 to 13-Nov-2026 is exactly 5 days
        rows = [{"Name": "David", "Last Updated": "2026-11-08"}]
        res = FilterEngine.filter_records(
            rows_data=rows,
            name_col="Name",
            date_col="Last Updated",
            source_file="test.csv",
            source_sheet="Sheet1",
            current_date=self.REF_DATE,
            threshold_days=5,
        )
        assert res.total_matching == 0
        assert res.total_skipped == 1
        assert res.skipped_records[0].days_difference == 5

    def test_6_days_difference_is_included(self):
        # 07-Nov-2026 to 13-Nov-2026 is 6 days
        rows = [{"Name": "Michael", "Last Updated": "2026-11-07"}]
        res = FilterEngine.filter_records(
            rows_data=rows,
            name_col="Name",
            date_col="Last Updated",
            source_file="test.csv",
            source_sheet="Sheet1",
            current_date=self.REF_DATE,
            threshold_days=5,
        )
        assert res.total_matching == 1
        assert res.matching_records[0].days_difference == 6

    def test_future_dates_are_excluded_and_flagged(self):
        rows = [
            {"Name": "Future User", "Last Updated": "2026-11-15"},
        ]
        res = FilterEngine.filter_records(
            rows_data=rows,
            name_col="Name",
            date_col="Last Updated",
            source_file="test.csv",
            source_sheet="Sheet1",
            current_date=self.REF_DATE,
            threshold_days=5,
        )
        assert res.total_matching == 0
        assert res.total_errors == 1
        assert res.errors[0].error_category == ErrorCategory.FUTURE_DATE

    def test_missing_dates_and_names(self):
        rows = [
            {"Name": "No Date User", "Last Updated": None},
            {"Name": None, "Last Updated": "2026-10-01"},
            {"Name": "   ", "Last Updated": "2026-10-01"},
        ]
        res = FilterEngine.filter_records(
            rows_data=rows,
            name_col="Name",
            date_col="Last Updated",
            source_file="test.csv",
            source_sheet="Sheet1",
            current_date=self.REF_DATE,
            threshold_days=5,
        )
        assert res.total_matching == 0
        assert res.total_errors == 3

    def test_configurable_threshold(self):
        rows = [
            {"Name": "Alice", "Last Updated": "2026-11-10"},  # 3 days diff
            {"Name": "Bob", "Last Updated": "2026-11-08"},    # 5 days diff
        ]
        # With threshold = 2: both included (3 > 2 and 5 > 2)
        res_t2 = FilterEngine.filter_records(
            rows_data=rows,
            name_col="Name",
            date_col="Last Updated",
            source_file="test.csv",
            source_sheet="Sheet1",
            current_date=self.REF_DATE,
            threshold_days=2,
        )
        assert res_t2.total_matching == 2

        # With threshold = 4: only Bob included (5 > 4)
        res_t4 = FilterEngine.filter_records(
            rows_data=rows,
            name_col="Name",
            date_col="Last Updated",
            source_file="test.csv",
            source_sheet="Sheet1",
            current_date=self.REF_DATE,
            threshold_days=4,
        )
        assert res_t4.total_matching == 1
        assert res_t4.matching_records[0].name == "Bob"

    def test_duplicate_names_are_preserved(self):
        rows = [
            {"Name": "John Smith", "Last Updated": "2026-10-01"},
            {"Name": "John Smith", "Last Updated": "2026-09-15"},
        ]
        res = FilterEngine.filter_records(
            rows_data=rows,
            name_col="Name",
            date_col="Last Updated",
            source_file="test.csv",
            source_sheet="Sheet1",
            current_date=self.REF_DATE,
            threshold_days=5,
        )
        assert res.total_matching == 2
        assert len(res.matching_records) == 2

    def test_calendar_day_calculation_around_midnight(self):
        # Even if timestamp is 23:59:59 on Nov 7, calendar date is Nov 7, so diff to Nov 13 is 6 days
        dt = datetime.datetime(2026, 11, 7, 23, 59, 59)
        cal_date = to_calendar_date(dt, "Asia/Kolkata")
        assert cal_date == datetime.date(2026, 11, 7)
        assert (self.REF_DATE - cal_date).days == 6

    def test_deduplicate_records_merges_source_files(self):
        from modules.filter_engine import Record

        r1 = Record(
            name="Sambit Chakraborty",
            last_updated=datetime.date(2026, 11, 3),
            days_difference=11,
            source_file="incident (2).xlsx",
            source_sheet="Page 1",
            row_number=2,
            incident="INC27334682",
            short_description="Environmental reports do not work",
        )
        r2 = Record(
            name="Sambit Chakraborty",
            last_updated=datetime.date(2026, 11, 3),
            days_difference=11,
            source_file="incident.csv",
            source_sheet="Sheet1",
            row_number=5,
            incident="INC27334682",
            short_description="Environmental reports do not work",
        )
        r3 = Record(
            name="Amol Kusalkar",
            last_updated=datetime.date(2026, 10, 28),
            days_difference=17,
            source_file="incident.csv",
            source_sheet="Sheet1",
            row_number=6,
            incident="INC27712210",
            short_description="Issue migrating LIMS CR",
        )

        deduped = FilterEngine.deduplicate_records([r1, r2, r3])
        assert len(deduped) == 2
        assert deduped[0].incident == "INC27334682"
        assert deduped[0].source_file == "incident (2).xlsx, incident.csv"
        assert deduped[1].incident == "INC27712210"
        assert deduped[1].source_file == "incident.csv"
