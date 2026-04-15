"""Tests for Book.getbof targeting uncovered lines."""
import io
import struct
import pytest

from xlrd.book import Book
from xlrd.biffh import (
    XLRDError,
    XL_WORKBOOK_GLOBALS,
    XL_WORKBOOK_GLOBALS_4W,
    XL_WORKSHEET,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_bof_book(mem_bytes, verbosity=0):
    """Create a minimally-initialized Book with the given raw bytes."""
    bk = Book()
    bk.mem = mem_bytes
    bk._position = 0
    bk.logfile = io.StringIO()
    bk.verbosity = verbosity
    return bk


def _bof_bytes(opcode, data):
    """Pack opcode(2) + length(2) + data bytes for a BOF-like record."""
    return struct.pack('<HH', opcode, len(data)) + data


# ---------------------------------------------------------------------------
# Error cases
# ---------------------------------------------------------------------------

class TestGetbofErrorCases:

    def test_eof_before_opcode(self):
        """Line 1282: opcode == MY_EOF when mem is empty."""
        bk = _make_bof_book(b'')
        with pytest.raises(XLRDError, match='Expected BOF record; met end of file'):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_eof_before_length(self):
        """Line 1287: length == MY_EOF when mem has opcode only."""
        bk = _make_bof_book(struct.pack('<H', 0x0809))
        with pytest.raises(XLRDError, match=r'Incomplete BOF record\[1\]; met end of file'):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_invalid_length_too_short(self):
        """Line 1289: length < 4 triggers invalid length error."""
        bk = _make_bof_book(struct.pack('<HH', 0x0809, 3))
        with pytest.raises(XLRDError, match='Invalid length'):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_invalid_length_too_long(self):
        """Line 1289: length > 20 triggers invalid length error."""
        bk = _make_bof_book(struct.pack('<HH', 0x0809, 21))
        with pytest.raises(XLRDError, match='Invalid length'):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_incomplete_data(self):
        """Line 1296: data shorter than declared length triggers error."""
        # length says 8 but only 4 bytes of data follow
        bk = _make_bof_book(struct.pack('<HH', 0x0809, 8) + b'\x00' * 4)
        with pytest.raises(XLRDError, match=r'Incomplete BOF record\[2\]; met end of file'):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_workspace_file_error(self):
        """Lines 1341-1342: streamtype == 0x0100 with version >= 50."""
        data = struct.pack('<HHHH', 0x0600, 0x0100, 0x0DBB, 0x07CC)
        bk = _make_bof_book(_bof_bytes(0x0809, data))
        with pytest.raises(XLRDError, match='Workspace file'):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_bof_not_workbook_worksheet_error(self):
        """Lines 1343-1346: unrecognized streamtype with version >= 50."""
        data = struct.pack('<HHHH', 0x0600, 0x9999, 0x0DBB, 0x07CC)
        bk = _make_bof_book(_bof_bytes(0x0809, data))
        with pytest.raises(XLRDError, match='BOF not workbook/worksheet'):
            bk.getbof(XL_WORKBOOK_GLOBALS)


# ---------------------------------------------------------------------------
# Version detection
# ---------------------------------------------------------------------------

class TestGetbofVersionDetection:

    def test_version_50_year_too_old(self):
        """Lines 1312-1314: version2=0x0500 with year<1994 → BIFF5 (version=50)."""
        data = struct.pack('<HHHH', 0x0500, XL_WORKBOOK_GLOBALS, 100, 1993)
        bk = _make_bof_book(_bof_bytes(0x0809, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 50

    def test_version_50_special_build(self):
        """Lines 1312-1314: version2=0x0500 with build in (2412, 3218, 3321) → BIFF5."""
        data = struct.pack('<HHHH', 0x0500, XL_WORKBOOK_GLOBALS, 2412, 1995)
        bk = _make_bof_book(_bof_bytes(0x0809, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 50

    def test_version_70_normal_build(self):
        """Line 1316: version2=0x0500, year>=1994, non-special build → BIFF7 (version=70)."""
        data = struct.pack('<HHHH', 0x0500, XL_WORKBOOK_GLOBALS, 100, 1995)
        bk = _make_bof_book(_bof_bytes(0x0809, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 70

    def test_dodgy_version_known_mapping(self):
        """Lines 1319-1325: version2=0x0000 maps to version=21 via dodgy dict."""
        data = struct.pack('<HHHH', 0x0000, XL_WORKBOOK_GLOBALS, 0, 0)
        bk = _make_bof_book(_bof_bytes(0x0809, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 21

    def test_dodgy_version_unknown_mapping(self):
        """Lines 1319-1325: version2 not in dodgy dict → version=0."""
        data = struct.pack('<HHHH', 0xAAAA, XL_WORKBOOK_GLOBALS, 0, 0)
        bk = _make_bof_book(_bof_bytes(0x0809, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 0

    def test_version1_0x04_returns_40(self):
        """Lines 1326-1327: version1=0x04 (opcode 0x0409) → BIFF4 (version=40)."""
        # data[0:2]=version2, data[2:4]=streamtype
        data = struct.pack('<HH', 0x0000, XL_WORKBOOK_GLOBALS) + b'\x00\x00'
        bk = _make_bof_book(_bof_bytes(0x0409, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 40

    def test_version1_0x02_returns_30(self):
        """Lines 1326-1327: version1=0x02 (opcode 0x0209) → BIFF3 (version=30)."""
        data = struct.pack('<HH', 0x0000, XL_WORKBOOK_GLOBALS) + b'\x00\x00'
        bk = _make_bof_book(_bof_bytes(0x0209, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 30

    def test_version1_0x00_returns_21(self):
        """Lines 1326-1327: version1=0x00 (opcode 0x0009) → BIFF2 (version=21)."""
        data = struct.pack('<HH', 0x0000, XL_WORKBOOK_GLOBALS)
        bk = _make_bof_book(_bof_bytes(0x0009, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 21

    def test_version_45_workbook_globals_4w(self):
        """Lines 1329-1330: version=40 with XL_WORKBOOK_GLOBALS_4W streamtype → version=45."""
        data = struct.pack('<HH', 0x0000, XL_WORKBOOK_GLOBALS_4W) + b'\x00\x00'
        bk = _make_bof_book(_bof_bytes(0x0409, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 45

    def test_version_lt50_worksheet_streamtype(self):
        """Lines 1339-1340: version<50 with XL_WORKSHEET streamtype returns version."""
        data = struct.pack('<HH', 0x0000, XL_WORKSHEET) + b'\x00\x00'
        bk = _make_bof_book(_bof_bytes(0x0409, data))
        assert bk.getbof(XL_WORKBOOK_GLOBALS) == 40


# ---------------------------------------------------------------------------
# Verbosity
# ---------------------------------------------------------------------------

class TestGetbofVerbosity:

    def test_verbosity_2_prints_bof_info(self):
        """Line 1333: verbosity >= 2 triggers BOF info print to logfile."""
        logfile = io.StringIO()
        data = struct.pack('<HHHH', 0x0600, XL_WORKBOOK_GLOBALS, 0x0DBB, 0x07CC)
        bk = _make_bof_book(_bof_bytes(0x0809, data), verbosity=2)
        bk.logfile = logfile
        version = bk.getbof(XL_WORKBOOK_GLOBALS)
        assert version == 80
        assert 'BOF' in logfile.getvalue()
