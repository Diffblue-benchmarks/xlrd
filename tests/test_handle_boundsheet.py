"""
Tests targeting uncovered lines in xlrd/book.py Book.handle_boundsheet.
"""
import struct
import sys
import pytest

from xlrd.book import Book
from xlrd.biffh import XL_BOUNDSHEET_WORKSHEET


def _make_book(biff_version=80):
    """Create a minimal Book instance with attributes needed by handle_boundsheet."""
    bk = Book()
    bk.logfile = sys.stdout
    bk.verbosity = 0
    bk.encoding_override = None
    bk.encoding = 'ascii'
    bk.codepage = 1252
    bk.biff_version = biff_version
    bk.base = 0
    bk._sheetsoffset = 0
    return bk


def _make_biff4w_data(sheet_name, encoding='ascii'):
    """Build BOUNDSHEET data for BIFF4W: just length-prefixed sheet name."""
    name_bytes = sheet_name.encode(encoding)
    return struct.pack('<B', len(name_bytes)) + name_bytes


def _make_biff_pre_unicode_data(sheet_name, offset=0, visibility=0,
                                 sheet_type=0, encoding='ascii'):
    """Build BOUNDSHEET data for non-BIFF4W pre-Unicode (bv < 80)."""
    name_bytes = sheet_name.encode(encoding)
    header = struct.pack('<iBB', offset, visibility, sheet_type)
    name_data = struct.pack('<B', len(name_bytes)) + name_bytes
    return header + name_data


def _make_biff8_data(sheet_name, offset=0, visibility=0, sheet_type=0):
    """Build BOUNDSHEET data for BIFF8 (bv >= 80) with Unicode name."""
    name_utf16 = sheet_name.encode('utf-16-le')
    nchars = len(sheet_name)
    header = struct.pack('<iBB', offset, visibility, sheet_type)
    # unpack_unicode: lenlen=1 so 1 byte length, 1 byte options, then chars
    name_data = struct.pack('<BB', nchars, 0x00) + name_utf16
    return header + name_data


class TestHandleBoundsheetBiff4W:

    def test_biff4w_first_sheet_abs_posn_from_sheetsoffset(self):
        """BIFF4W with empty _sh_abs_posn: abs_posn = _sheetsoffset + base."""
        bk = _make_book(biff_version=45)
        bk._sheetsoffset = 100
        bk.base = 50
        data = _make_biff4w_data('Sheet1')
        bk.handle_boundsheet(data)
        assert bk._sh_abs_posn[0] == 150  # sheetsoffset + base

    def test_biff4w_first_sheet_name_recorded(self):
        """BIFF4W first sheet: sheet_name and visibility are set correctly."""
        bk = _make_book(biff_version=45)
        bk._sheetsoffset = 0
        bk.base = 0
        data = _make_biff4w_data('MySheet')
        bk.handle_boundsheet(data)
        assert bk._sheet_names[0] == 'MySheet'
        assert bk._sheet_visibility[0] == 0

    def test_biff4w_second_sheet_abs_posn_minus_one(self):
        """BIFF4W with non-empty _sh_abs_posn: abs_posn = -1."""
        bk = _make_book(biff_version=45)
        bk._sheetsoffset = 100
        bk.base = 0
        # First call fills _sh_abs_posn
        data1 = _make_biff4w_data('Sheet1')
        bk.handle_boundsheet(data1)
        assert len(bk._sh_abs_posn) == 1
        # Second call: _sh_abs_posn is non-empty => abs_posn = -1
        data2 = _make_biff4w_data('Sheet2')
        bk.handle_boundsheet(data2)
        assert bk._sh_abs_posn[1] == -1

    def test_biff4w_all_sheets_count_increments(self):
        """BIFF4W: _all_sheets_count increments on each call."""
        bk = _make_book(biff_version=45)
        bk._sheetsoffset = 0
        bk.base = 0
        bk.handle_boundsheet(_make_biff4w_data('Sheet1'))
        bk.handle_boundsheet(_make_biff4w_data('Sheet2'))
        assert bk._all_sheets_count == 2


class TestHandleBoundsheetPreUnicode:

    def test_pre_unicode_sheet_name_decoded(self):
        """BIFF5 (pre-Unicode): sheet name read via unpack_string."""
        bk = _make_book(biff_version=50)
        data = _make_biff_pre_unicode_data('Sheet1', offset=100)
        bk.handle_boundsheet(data)
        assert bk._sheet_names[0] == 'Sheet1'

    def test_pre_unicode_abs_posn_from_offset(self):
        """BIFF5: abs_posn = offset + base."""
        bk = _make_book(biff_version=50)
        bk.base = 10
        data = _make_biff_pre_unicode_data('Sheet1', offset=200)
        bk.handle_boundsheet(data)
        assert bk._sh_abs_posn[0] == 210


class TestHandleBoundsheetVerbosity:

    def test_verbosity_2_logs_boundsheet_info(self, capsys):
        """verbosity >= 2: fprintf is called with boundsheet info."""
        bk = _make_book(biff_version=80)
        bk.verbosity = 2
        data = _make_biff8_data('Sheet1', offset=100)
        bk.handle_boundsheet(data)
        captured = capsys.readouterr()
        assert 'BOUNDSHEET' in captured.out


class TestHandleBoundsheetNonWorksheet:

    def test_chart_sheet_appends_minus_one_to_all_sheets_map(self):
        """Non-worksheet (Chart type=2): _all_sheets_map gets -1."""
        bk = _make_book(biff_version=80)
        data = _make_biff8_data('MyChart', offset=0, sheet_type=2)
        bk.handle_boundsheet(data)
        assert bk._all_sheets_map[0] == -1

    def test_chart_sheet_not_added_to_sheet_names(self):
        """Non-worksheet type: sheet name is not added to _sheet_names."""
        bk = _make_book(biff_version=80)
        data = _make_biff8_data('MyChart', offset=0, sheet_type=2)
        bk.handle_boundsheet(data)
        assert len(bk._sheet_names) == 0

    def test_macro_sheet_appends_minus_one_to_all_sheets_map(self):
        """Macro sheet (type=1): _all_sheets_map gets -1."""
        bk = _make_book(biff_version=80)
        data = _make_biff8_data('MyMacro', offset=0, sheet_type=1)
        bk.handle_boundsheet(data)
        assert bk._all_sheets_map[0] == -1

    def test_unknown_sheet_type_logged_with_verbosity(self, capsys):
        """Non-worksheet type with verbosity >= 1: NOTE log message emitted."""
        bk = _make_book(biff_version=80)
        bk.verbosity = 1
        data = _make_biff8_data('UnknownSheet', offset=0, sheet_type=0x99)
        bk.handle_boundsheet(data)
        captured = capsys.readouterr()
        assert 'Ignoring non-worksheet' in captured.out

    def test_chart_sheet_type_logged_with_verbosity(self, capsys):
        """Chart type (2) with verbosity >= 1: NOTE log includes 'Chart'."""
        bk = _make_book(biff_version=80)
        bk.verbosity = 1
        data = _make_biff8_data('Chart1', offset=0, sheet_type=2)
        bk.handle_boundsheet(data)
        captured = capsys.readouterr()
        assert 'Chart' in captured.out

    def test_vba_sheet_appends_minus_one_to_all_sheets_map(self):
        """VBA module (type=6): _all_sheets_map gets -1."""
        bk = _make_book(biff_version=80)
        data = _make_biff8_data('VBAModule', offset=0, sheet_type=6)
        bk.handle_boundsheet(data)
        assert bk._all_sheets_map[0] == -1
