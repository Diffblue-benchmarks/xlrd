"""
Tests targeting uncovered lines in xlrd/book.py Book.handle_supbook.
"""
import struct
import sys
import pytest

from xlrd.book import Book, SUPBOOK_INTERNAL, SUPBOOK_ADDIN, SUPBOOK_DDEOLE, SUPBOOK_EXTERNAL


def _make_book(verbosity=0):
    """Create a minimal Book instance with attributes needed by handle_supbook."""
    bk = Book()
    bk.logfile = sys.stdout
    bk.verbosity = verbosity
    bk._supbook_count = 0
    bk._supbook_types = []
    bk._supbook_locals_inx = None
    bk._supbook_addins_inx = None
    bk._all_sheets_map = []
    return bk


def _pack_unicode_str(s):
    """Pack a string as a BIFF unicode string (lenlen=2, ascii options=0)."""
    encoded = s.encode('ascii')
    return struct.pack('<H', len(s)) + b'\x00' + encoded


def _make_internal_data(num_sheets):
    """Build SUPBOOK data for an internal reference."""
    return struct.pack('<H', num_sheets) + b'\x01\x04'


def _make_addin_data():
    """Build SUPBOOK data for add-in functions."""
    return b'\x01\x00\x01\x3A'


def _make_ddeole_data(url):
    """Build SUPBOOK data for DDE/OLE with zero sheets."""
    return struct.pack('<H', 0) + _pack_unicode_str(url)


def _make_external_data(url, sheet_names):
    """Build SUPBOOK data for an external workbook with given sheet names."""
    data = struct.pack('<H', len(sheet_names)) + _pack_unicode_str(url)
    for name in sheet_names:
        data += _pack_unicode_str(name)
    return data


class TestHandleSupbook:

    def test_internal_sets_supbook_type(self):
        """handle_supbook sets SUPBOOK_INTERNAL when data[2:4] == b'\\x01\\x04'."""
        bk = _make_book()
        data = _make_internal_data(num_sheets=3)
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_INTERNAL

    def test_internal_sets_locals_inx(self):
        """handle_supbook sets _supbook_locals_inx for internal supbook."""
        bk = _make_book()
        data = _make_internal_data(num_sheets=2)
        bk.handle_supbook(data)
        assert bk._supbook_locals_inx == 0

    def test_internal_increments_supbook_count(self):
        """handle_supbook increments _supbook_count for internal supbook."""
        bk = _make_book()
        data = _make_internal_data(num_sheets=1)
        bk.handle_supbook(data)
        assert bk._supbook_count == 1

    def test_internal_appends_to_supbook_types(self):
        """handle_supbook appends to _supbook_types."""
        bk = _make_book()
        data = _make_internal_data(num_sheets=1)
        bk.handle_supbook(data)
        assert len(bk._supbook_types) == 1

    def test_internal_with_verbosity(self):
        """handle_supbook with verbosity >= 2 logs internal supbook info."""
        bk = _make_book(verbosity=2)
        data = _make_internal_data(num_sheets=1)
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_INTERNAL

    def test_addin_sets_supbook_type(self):
        """handle_supbook sets SUPBOOK_ADDIN when data[0:4] == b'\\x01\\x00\\x01\\x3A'."""
        bk = _make_book()
        data = _make_addin_data()
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_ADDIN

    def test_addin_sets_addins_inx(self):
        """handle_supbook sets _supbook_addins_inx for add-in supbook."""
        bk = _make_book()
        data = _make_addin_data()
        bk.handle_supbook(data)
        assert bk._supbook_addins_inx == 0

    def test_addin_increments_supbook_count(self):
        """handle_supbook increments _supbook_count for add-in supbook."""
        bk = _make_book()
        data = _make_addin_data()
        bk.handle_supbook(data)
        assert bk._supbook_count == 1

    def test_addin_with_verbosity(self):
        """handle_supbook with verbosity >= 2 logs add-in supbook info."""
        bk = _make_book(verbosity=2)
        data = _make_addin_data()
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_ADDIN

    def test_ddeole_sets_supbook_type(self):
        """handle_supbook sets SUPBOOK_DDEOLE when num_sheets == 0."""
        bk = _make_book()
        data = _make_ddeole_data('dde_document')
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_DDEOLE

    def test_ddeole_increments_supbook_count(self):
        """handle_supbook increments _supbook_count for DDE/OLE supbook."""
        bk = _make_book()
        data = _make_ddeole_data('ole_document')
        bk.handle_supbook(data)
        assert bk._supbook_count == 1

    def test_ddeole_with_verbosity(self):
        """handle_supbook with verbosity >= 2 logs DDE/OLE supbook info."""
        bk = _make_book(verbosity=2)
        data = _make_ddeole_data('doc')
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_DDEOLE

    def test_external_sets_supbook_type(self):
        """handle_supbook sets SUPBOOK_EXTERNAL for external workbook."""
        bk = _make_book()
        data = _make_external_data('external.xls', ['Sheet1'])
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_EXTERNAL

    def test_external_increments_supbook_count(self):
        """handle_supbook increments _supbook_count for external supbook."""
        bk = _make_book()
        data = _make_external_data('external.xls', ['Sheet1'])
        bk.handle_supbook(data)
        assert bk._supbook_count == 1

    def test_external_multiple_sheets(self):
        """handle_supbook handles external supbook with multiple sheet names."""
        bk = _make_book()
        data = _make_external_data('workbook.xls', ['Alpha', 'Beta', 'Gamma'])
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_EXTERNAL

    def test_external_with_verbosity(self):
        """handle_supbook with verbosity >= 2 logs external supbook info."""
        bk = _make_book(verbosity=2)
        data = _make_external_data('ext.xls', ['Sheet1', 'Sheet2'])
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_EXTERNAL

    def test_external_struct_error_with_verbosity(self):
        """handle_supbook handles struct.error when unpacking sheet names and verbosity > 0."""
        bk = _make_book(verbosity=1)
        # Build data claiming 2 sheets but only provide data for the url
        # Append only 1 byte for the next string length (needs 2 for lenlen=2 -> struct.error)
        url = 'trunc.xls'
        data = struct.pack('<H', 2) + _pack_unicode_str(url) + b'\x03'
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_EXTERNAL

    def test_external_struct_error_silent(self):
        """handle_supbook handles struct.error silently when verbosity == 0."""
        bk = _make_book(verbosity=0)
        url = 'trunc.xls'
        data = struct.pack('<H', 2) + _pack_unicode_str(url) + b'\x03'
        bk.handle_supbook(data)
        assert bk._supbook_types[-1] == SUPBOOK_EXTERNAL

    def test_multiple_supbooks_accumulate(self):
        """handle_supbook accumulates multiple entries in _supbook_types."""
        bk = _make_book()
        bk.handle_supbook(_make_internal_data(1))
        bk.handle_supbook(_make_addin_data())
        bk.handle_supbook(_make_ddeole_data('doc'))
        assert len(bk._supbook_types) == 3
        assert bk._supbook_count == 3
        assert bk._supbook_types[0] == SUPBOOK_INTERNAL
        assert bk._supbook_types[1] == SUPBOOK_ADDIN
        assert bk._supbook_types[2] == SUPBOOK_DDEOLE
