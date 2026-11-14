"""Unit tests for date parsing, format normalization, and ambiguity resolution."""
import datetime
from zoneinfo import ZoneInfo
import pandas as pd
import pytest

from modules.date_parser import DateParser


class TestDateParsing:
    """Test suite covering all required date formats and edge cases."""

    def test_1_dd_mm_yyyy(self):
        res = DateParser.parse_date("25/11/2026", preference="Auto-detect")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 25)

    def test_2_mm_dd_yyyy(self):
        res = DateParser.parse_date("11/25/2026", preference="Auto-detect")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 25)

    def test_3_yyyy_mm_dd(self):
        res = DateParser.parse_date("2026-11-05")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_4_dd_mm_yyyy_hyphen(self):
        res = DateParser.parse_date("05-11-2026", preference="DD/MM/YYYY")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_5_dd_mm_yyyy_dot(self):
        res = DateParser.parse_date("05.11.2026", preference="DD/MM/YYYY")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_6_dd_mmm_yyyy(self):
        res = DateParser.parse_date("05-Nov-2026")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_7_mmm_dd_yyyy(self):
        res = DateParser.parse_date("Nov 05, 2026")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_8_yyyy_mm_dd_slash(self):
        res = DateParser.parse_date("2026/11/05")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_9_iso_timestamp(self):
        res = DateParser.parse_date("2026-11-05T14:30:00")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_10_timestamp_with_timezone(self):
        # 2026-11-05T23:00:00Z in UTC will be 2026-11-06 in Asia/Kolkata (+05:30)
        res = DateParser.parse_date("2026-11-05T23:00:00Z", timezone_name="Asia/Kolkata")
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 6)

    def test_11_excel_serial_date(self):
        # Excel serial date 46331 is 2026-11-05
        res = DateParser.parse_date(46331)
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_12_native_python_datetime(self):
        dt = datetime.datetime(2026, 11, 5, 10, 15, 0)
        res = DateParser.parse_date(dt)
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_13_native_python_date(self):
        d = datetime.date(2026, 11, 5)
        res = DateParser.parse_date(d)
        assert res.success is True
        assert res.date_obj == datetime.date(2026, 11, 5)

    def test_14_empty_values(self):
        assert DateParser.parse_date(None).success is False
        assert DateParser.parse_date("").success is False
        assert DateParser.parse_date("   ").success is False
        assert DateParser.parse_date("nan").success is False
        assert DateParser.parse_date(float("nan")).success is False

    def test_15_invalid_dates(self):
        assert DateParser.parse_date("not a date").success is False
        assert DateParser.parse_date("32/13/2026").success is False
        assert DateParser.parse_date(99999999).success is False  # Outside serial range

    def test_16_ambiguous_date_preference_resolution(self):
        # 05/11/2026
        # With DD/MM/YYYY preference -> Nov 5, 2026
        res_dd = DateParser.parse_date("05/11/2026", preference="DD/MM/YYYY")
        assert res_dd.success is True
        assert res_dd.date_obj == datetime.date(2026, 11, 5)

        # With MM/DD/YYYY preference -> May 11, 2026
        res_mm = DateParser.parse_date("05/11/2026", preference="MM/DD/YYYY")
        assert res_mm.success is True
        assert res_mm.date_obj == datetime.date(2026, 5, 11)

    def test_17_leap_year_date(self):
        res = DateParser.parse_date("29/02/2024")
        assert res.success is True
        assert res.date_obj == datetime.date(2024, 2, 29)

        # Invalid leap year (2023 is not a leap year)
        res_invalid = DateParser.parse_date("29/02/2023")
        assert res_invalid.success is False
