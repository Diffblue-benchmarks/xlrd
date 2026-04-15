"""
Tests targeting uncovered lines in xlrd/book.py Book.handle_datemode.
"""
import struct
import sys
import pytest

from xlrd.book import Book


def _make_book(verbosity=0):
    """Create a minimal Book instance with attributes needed by handle_datemode."""
    bk = Book()
    bk.logfile = sys.stdout
    bk.verbosity = verbosity
    bk.datemode = 0
    return bk


def _make_datemode_data(datemode):
    """Build DATEMODE record data."""
    return struct.pack('<H', datemode)


class TestHandleDatemode:

    def test_datemode_zero_sets_attribute(self):
        """handle_datemode sets datemode to 0."""
        bk = _make_book()
        data = _make_datemode_data(0)
        bk.handle_datemode(data)
        assert bk.datemode == 0

    def test_datemode_one_sets_attribute(self):
        """handle_datemode sets datemode to 1."""
        bk = _make_book()
        data = _make_datemode_data(1)
        bk.handle_datemode(data)
        assert bk.datemode == 1

    def test_datemode_with_verbosity_logs(self):
        """handle_datemode with verbosity logs the datemode value."""
        bk = _make_book(verbosity=1)
        data = _make_datemode_data(1)
        bk.handle_datemode(data)
        assert bk.datemode == 1

    def test_datemode_invalid_raises_assertion(self):
        """handle_datemode raises AssertionError for invalid datemode values."""
        bk = _make_book()
        data = _make_datemode_data(2)
        with pytest.raises(AssertionError):
            bk.handle_datemode(data)
