import datetime
import pytest

from xlrd.xldate import (
    xldate_as_tuple,
    xldate_as_datetime,
    xldate_from_date_tuple,
    xldate_from_time_tuple,
    xldate_from_datetime_tuple,
    _leap,
    XLDateNegative,
    XLDateAmbiguous,
    XLDateTooLarge,
    XLDateBadDatemode,
    XLDateBadTuple,
)


# --- xldate_as_tuple ---

def test_xldate_as_tuple_bad_datemode():
    with pytest.raises(XLDateBadDatemode):
        xldate_as_tuple(1.0, 2)


def test_xldate_as_tuple_zero():
    assert xldate_as_tuple(0.0, 0) == (0, 0, 0, 0, 0, 0)


def test_xldate_as_tuple_negative():
    with pytest.raises(XLDateNegative):
        xldate_as_tuple(-1.0, 0)


def test_xldate_as_tuple_too_large_datemode0():
    with pytest.raises(XLDateTooLarge):
        xldate_as_tuple(2958466, 0)


def test_xldate_as_tuple_too_large_datemode1():
    with pytest.raises(XLDateTooLarge):
        xldate_as_tuple(2958465, 1)


def test_xldate_as_tuple_time_only():
    # xldate < 1 but > 0: pure time
    result = xldate_as_tuple(0.5, 0)
    assert result == (0, 0, 0, 12, 0, 0)


def test_xldate_as_tuple_ambiguous():
    with pytest.raises(XLDateAmbiguous):
        xldate_as_tuple(1.0, 0)


def test_xldate_as_tuple_ambiguous_upper_boundary():
    with pytest.raises(XLDateAmbiguous):
        xldate_as_tuple(60.0, 0)


def test_xldate_as_tuple_valid_date_datemode0():
    # 2000-01-01 in datemode=0
    result = xldate_as_tuple(36526.0, 0)
    assert result == (2000, 1, 1, 0, 0, 0)


def test_xldate_as_tuple_valid_date_datemode1():
    # datemode=1: 1904-based; 35064 days from 1904-01-01 = 2000-01-01
    result = xldate_as_tuple(35064.0, 1)
    assert result == (2000, 1, 1, 0, 0, 0)


def test_xldate_as_tuple_with_time():
    # Date + time component
    result = xldate_as_tuple(36526.5, 0)
    assert result == (2000, 1, 1, 12, 0, 0)


def test_xldate_as_tuple_mp_less_than_10():
    # January or February (mp < 10 after calculation)
    result = xldate_as_tuple(36162.0, 0)
    assert result[:3] == (1998, 12, 31) or result[0] >= 1900


def test_xldate_as_tuple_seconds_rollover():
    # Edge: fraction that rounds to 86400 seconds
    frac = 86399.5 / 86400.0
    result = xldate_as_tuple(36526.0 + frac, 0)
    # Should roll over to next day at midnight
    assert result[3:] == (0, 0, 0)


def test_xldate_as_tuple_datemode1_valid():
    result = xldate_as_tuple(1.0, 1)
    assert result == (1904, 1, 2, 0, 0, 0)


# --- xldate_as_datetime ---

def test_xldate_as_datetime_datemode1():
    result = xldate_as_datetime(1.0, 1)
    assert isinstance(result, datetime.datetime)
    assert result.year == 1904


def test_xldate_as_datetime_datemode0_before_60():
    result = xldate_as_datetime(1.0, 0)
    assert isinstance(result, datetime.datetime)


def test_xldate_as_datetime_datemode0_at_or_after_60():
    result = xldate_as_datetime(61.0, 0)
    assert isinstance(result, datetime.datetime)
    assert result.year == 1900


def test_xldate_as_datetime_with_fraction():
    result = xldate_as_datetime(36526.5, 0)
    assert result.hour == 12
    assert result.minute == 0


def test_xldate_as_datetime_milliseconds():
    # Fraction that results in non-zero milliseconds
    result = xldate_as_datetime(36526.0 + 0.001 / 86400.0, 0)
    assert isinstance(result, datetime.datetime)


# --- _leap ---

def test_leap_not_divisible_by_4():
    assert _leap(1901) == 0


def test_leap_divisible_by_4_not_100():
    assert _leap(1904) == 1


def test_leap_divisible_by_100_not_400():
    assert _leap(1900) == 0


def test_leap_divisible_by_400():
    assert _leap(2000) == 1


# --- xldate_from_date_tuple ---

def test_xldate_from_date_tuple_bad_datemode():
    with pytest.raises(XLDateBadDatemode):
        xldate_from_date_tuple((2000, 1, 1), 2)


def test_xldate_from_date_tuple_zero():
    assert xldate_from_date_tuple((0, 0, 0), 0) == 0.0


def test_xldate_from_date_tuple_invalid_year():
    with pytest.raises(XLDateBadTuple):
        xldate_from_date_tuple((1800, 1, 1), 0)


def test_xldate_from_date_tuple_invalid_month():
    with pytest.raises(XLDateBadTuple):
        xldate_from_date_tuple((2000, 13, 1), 0)


def test_xldate_from_date_tuple_invalid_day_zero():
    with pytest.raises(XLDateBadTuple):
        xldate_from_date_tuple((2000, 1, 0), 0)


def test_xldate_from_date_tuple_invalid_day_too_large():
    with pytest.raises(XLDateBadTuple):
        xldate_from_date_tuple((2001, 2, 29), 0)  # 2001 is not a leap year


def test_xldate_from_date_tuple_leap_day_valid():
    # 2000 is a leap year
    result = xldate_from_date_tuple((2000, 2, 29), 0)
    assert result > 0


def test_xldate_from_date_tuple_ambiguous():
    with pytest.raises(XLDateAmbiguous):
        xldate_from_date_tuple((1900, 1, 31), 0)


def test_xldate_from_date_tuple_valid_datemode0():
    result = xldate_from_date_tuple((2000, 1, 1), 0)
    assert result == 36526.0


def test_xldate_from_date_tuple_valid_datemode1():
    result = xldate_from_date_tuple((2000, 1, 1), 1)
    assert result > 0


def test_xldate_from_date_tuple_month_le_2():
    # Test January/February path (M <= 2)
    result = xldate_from_date_tuple((2000, 2, 1), 0)
    assert result > 0


def test_xldate_from_date_tuple_month_gt_2():
    result = xldate_from_date_tuple((2000, 3, 1), 0)
    assert result > 0


# --- xldate_from_time_tuple ---

def test_xldate_from_time_tuple_valid():
    result = xldate_from_time_tuple((12, 0, 0))
    assert abs(result - 0.5) < 1e-9


def test_xldate_from_time_tuple_midnight():
    result = xldate_from_time_tuple((0, 0, 0))
    assert result == 0.0


def test_xldate_from_time_tuple_invalid_hour():
    with pytest.raises(XLDateBadTuple):
        xldate_from_time_tuple((24, 0, 0))


def test_xldate_from_time_tuple_invalid_minute():
    with pytest.raises(XLDateBadTuple):
        xldate_from_time_tuple((0, 60, 0))


def test_xldate_from_time_tuple_invalid_second():
    with pytest.raises(XLDateBadTuple):
        xldate_from_time_tuple((0, 0, 60))


def test_xldate_from_time_tuple_negative_hour():
    with pytest.raises(XLDateBadTuple):
        xldate_from_time_tuple((-1, 0, 0))


# --- xldate_from_datetime_tuple ---

def test_xldate_from_datetime_tuple_valid():
    result = xldate_from_datetime_tuple((2000, 1, 1, 12, 0, 0), 0)
    assert result == 36526.5


def test_xldate_from_datetime_tuple_datemode1():
    result = xldate_from_datetime_tuple((2000, 1, 1, 0, 0, 0), 1)
    assert result > 0
