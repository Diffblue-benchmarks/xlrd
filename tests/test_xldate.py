# -*- coding: utf-8 -*-
"""
Tests for xlrd.xldate — date/time conversion utilities.
"""
import datetime
import pytest

from xlrd.xldate import (
    XLDateAmbiguous,
    XLDateBadDatemode,
    XLDateBadTuple,
    XLDateNegative,
    XLDateTooLarge,
    xldate_as_datetime,
    xldate_as_tuple,
    xldate_from_date_tuple,
    xldate_from_datetime_tuple,
    xldate_from_time_tuple,
)


# ---------------------------------------------------------------------------
# xldate_as_tuple
# ---------------------------------------------------------------------------

class TestXldateAsTuple:

    def test_zero_value_returns_time_zero(self):
        assert xldate_as_tuple(0.0, 0) == (0, 0, 0, 0, 0, 0)

    def test_zero_value_datemode_1(self):
        assert xldate_as_tuple(0.0, 1) == (0, 0, 0, 0, 0, 0)

    def test_negative_xldate_raises(self):
        with pytest.raises(XLDateNegative):
            xldate_as_tuple(-1.0, 0)

    def test_bad_datemode_raises(self):
        with pytest.raises(XLDateBadDatemode):
            xldate_as_tuple(40000.0, 2)

    def test_bad_datemode_negative_raises(self):
        with pytest.raises(XLDateBadDatemode):
            xldate_as_tuple(40000.0, -1)

    def test_ambiguous_1900_raises(self):
        # datemode 0 and 1.0 <= xldate < 61.0 is ambiguous (Excel 1900 leap year bug)
        with pytest.raises(XLDateAmbiguous):
            xldate_as_tuple(30.0, 0)

    def test_ambiguous_boundary_1(self):
        with pytest.raises(XLDateAmbiguous):
            xldate_as_tuple(1.0, 0)

    def test_ambiguous_boundary_60(self):
        with pytest.raises(XLDateAmbiguous):
            xldate_as_tuple(60.0, 0)

    def test_too_large_raises(self):
        # 2958466 is the too-large threshold for datemode 0
        with pytest.raises(XLDateTooLarge):
            xldate_as_tuple(2958466.0, 0)

    def test_too_large_datemode_1_raises(self):
        with pytest.raises(XLDateTooLarge):
            xldate_as_tuple(2958465.0, 1)  # threshold is 2958466 - 1462

    def test_time_only_no_date(self):
        # 0 < xldate < 1 -> time only
        # 0.5 = noon
        result = xldate_as_tuple(0.5, 0)
        assert result == (0, 0, 0, 12, 0, 0)

    def test_time_fraction_midnight_next_day(self):
        # A fractional part that rounds to 86400 (midnight of next day).
        # 0.9999999 * 86400 rounds to 86400, so xldays increments to 1.
        # In datemode 0, xldays=1 is ambiguous (1900 bug range 1..60).
        with pytest.raises(XLDateAmbiguous):
            xldate_as_tuple(0.9999999, 0)

    def test_known_date_datemode_0(self):
        # Excel serial 61 = 1900-03-01 in datemode 0 (first non-ambiguous date)
        result = xldate_as_tuple(61.0, 0)
        assert result == (1900, 3, 1, 0, 0, 0)

    def test_known_date_2000_01_01_datemode_0(self):
        # 2000-01-01 in 1900-based system = serial 36526
        result = xldate_as_tuple(36526.0, 0)
        assert result == (2000, 1, 1, 0, 0, 0)

    def test_known_date_datemode_1(self):
        # In 1904-based system, serial 1 = 1904-01-02
        result = xldate_as_tuple(1.0, 1)
        assert result == (1904, 1, 2, 0, 0, 0)

    def test_datetime_with_time_component(self):
        # 36526.5 = 2000-01-01 12:00:00
        result = xldate_as_tuple(36526.5, 0)
        assert result == (2000, 1, 1, 12, 0, 0)

    def test_datetime_with_minutes_seconds(self):
        # 36526 + (1*3600 + 30*60 + 45) / 86400.0
        fraction = (1 * 3600 + 30 * 60 + 45) / 86400.0
        result = xldate_as_tuple(36526.0 + fraction, 0)
        assert result == (2000, 1, 1, 1, 30, 45)


# ---------------------------------------------------------------------------
# xldate_as_datetime
# ---------------------------------------------------------------------------

class TestXldateAsDatetime:

    def test_zero_gives_1899_12_31_for_datemode_0(self):
        # For datemode 0, xldate < 60, epoch is 1899-12-31; days=0, so 1899-12-31
        result = xldate_as_datetime(0.0, 0)
        assert result == datetime.datetime(1899, 12, 31)

    def test_serial_1_datemode_1(self):
        # Serial 1 in 1904-mode = 1904-01-02
        result = xldate_as_datetime(1.0, 1)
        assert result == datetime.datetime(1904, 1, 2)

    def test_known_date_2000_01_01_datemode_0(self):
        result = xldate_as_datetime(36526.0, 0)
        assert result == datetime.datetime(2000, 1, 1, 0, 0, 0)

    def test_datetime_with_time_component(self):
        # 36526.5 = 2000-01-01 12:00:00
        result = xldate_as_datetime(36526.5, 0)
        assert result == datetime.datetime(2000, 1, 1, 12, 0, 0)

    def test_datemode_1_epoch_1904(self):
        # Serial 0 in 1904-mode = 1904-01-01
        result = xldate_as_datetime(0.0, 1)
        assert result == datetime.datetime(1904, 1, 1)

    def test_pre_60_uses_1899_12_31_epoch(self):
        # xldate < 60 uses epoch_1900 = 1899-12-31
        result = xldate_as_datetime(1.0, 0)
        assert result == datetime.datetime(1899, 12, 31) + datetime.timedelta(days=1)

    def test_post_60_uses_1899_12_30_epoch(self):
        # xldate >= 60 uses epoch_1900_minus_1 = 1899-12-30
        result = xldate_as_datetime(61.0, 0)
        assert result == datetime.datetime(1900, 3, 1)

    def test_millisecond_precision(self):
        # fraction with milliseconds
        # 0.5 day = 43200 seconds = 12:00:00.000
        result = xldate_as_datetime(36526.5, 0)
        assert result.hour == 12
        assert result.minute == 0
        assert result.second == 0
        assert result.microsecond == 0


# ---------------------------------------------------------------------------
# xldate_from_date_tuple
# ---------------------------------------------------------------------------

class TestXldateFromDateTuple:

    def test_zero_tuple_returns_zero(self):
        assert xldate_from_date_tuple((0, 0, 0), 0) == 0.0

    def test_bad_datemode_raises(self):
        with pytest.raises(XLDateBadDatemode):
            xldate_from_date_tuple((2000, 1, 1), 2)

    def test_invalid_year_too_low_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_date_tuple((1800, 1, 1), 0)

    def test_invalid_year_too_high_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_date_tuple((10000, 1, 1), 0)

    def test_invalid_month_zero_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_date_tuple((2000, 0, 1), 0)

    def test_invalid_month_13_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_date_tuple((2000, 13, 1), 0)

    def test_invalid_day_zero_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_date_tuple((2000, 1, 0), 0)

    def test_invalid_day_32_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_date_tuple((2000, 1, 32), 0)

    def test_invalid_day_feb_29_non_leap_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_date_tuple((1900, 2, 29), 0)

    def test_valid_day_feb_29_leap_year(self):
        # 2000 is a leap year
        result = xldate_from_date_tuple((2000, 2, 29), 0)
        assert result > 0

    def test_ambiguous_before_1900_03_01_raises(self):
        with pytest.raises(XLDateAmbiguous):
            xldate_from_date_tuple((1900, 2, 28), 0)

    def test_known_date_2000_01_01_datemode_0(self):
        result = xldate_from_date_tuple((2000, 1, 1), 0)
        assert result == 36526.0

    def test_roundtrip_datemode_0(self):
        original = (2023, 6, 15)
        serial = xldate_from_date_tuple(original, 0)
        result = xldate_as_tuple(serial, 0)
        assert result[:3] == original

    def test_roundtrip_datemode_1(self):
        original = (2023, 6, 15)
        serial = xldate_from_date_tuple(original, 1)
        result = xldate_as_tuple(serial, 1)
        assert result[:3] == original


# ---------------------------------------------------------------------------
# xldate_from_time_tuple
# ---------------------------------------------------------------------------

class TestXldateFromTimeTuple:

    def test_midnight(self):
        assert xldate_from_time_tuple((0, 0, 0)) == 0.0

    def test_noon(self):
        result = xldate_from_time_tuple((12, 0, 0))
        assert abs(result - 0.5) < 1e-10

    def test_end_of_day(self):
        # 23:59:59
        result = xldate_from_time_tuple((23, 59, 59))
        assert 0.0 < result < 1.0

    def test_invalid_hour_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_time_tuple((24, 0, 0))

    def test_invalid_negative_hour_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_time_tuple((-1, 0, 0))

    def test_invalid_minute_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_time_tuple((0, 60, 0))

    def test_invalid_second_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_time_tuple((0, 0, 60))

    def test_roundtrip_via_tuple(self):
        # 6:30:15 as fraction
        fraction = xldate_from_time_tuple((6, 30, 15))
        result = xldate_as_tuple(fraction, 0)
        assert result == (0, 0, 0, 6, 30, 15)


# ---------------------------------------------------------------------------
# xldate_from_datetime_tuple
# ---------------------------------------------------------------------------

class TestXldateFromDatetimeTuple:

    def test_known_datetime(self):
        # 2000-01-01 12:00:00 = 36526.5
        result = xldate_from_datetime_tuple((2000, 1, 1, 12, 0, 0), 0)
        assert abs(result - 36526.5) < 1e-6

    def test_roundtrip(self):
        original = (2023, 6, 15, 9, 45, 30)
        serial = xldate_from_datetime_tuple(original, 0)
        result = xldate_as_tuple(serial, 0)
        assert result == original

    def test_invalid_date_part_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_datetime_tuple((10000, 1, 1, 0, 0, 0), 0)

    def test_invalid_time_part_raises(self):
        with pytest.raises(XLDateBadTuple):
            xldate_from_datetime_tuple((2000, 1, 1, 25, 0, 0), 0)


# ---------------------------------------------------------------------------
# xldate_from_date_tuple — xldays <= 0 raises XLDateBadTuple (line 214)
# ---------------------------------------------------------------------------

class TestXldateFromDateTupleExtraEdgeCases:

    def test_date_before_mac_epoch_raises(self):
        """xldate_from_date_tuple raises when xldays <= 0 in Mac 1904 datemode (line 214)."""
        from xlrd.xldate import xldate_from_date_tuple, XLDateBadTuple
        # 1900-01-01 is before the Mac 1904 epoch, so xldays <= 0
        with pytest.raises(XLDateBadTuple, match="Invalid"):
            xldate_from_date_tuple((1900, 1, 1), datemode=1)
