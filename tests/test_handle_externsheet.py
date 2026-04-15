"""
Tests targeting uncovered lines in xlrd/book.py Book.handle_externsheet.
"""
import struct
import sys
import pytest

from xlrd.book import Book
from xlrd.biffh import XLRDError, XL_CONTINUE


def _make_book(biff_version=80, verbosity=0):
    """Create a minimal Book instance with attributes needed by handle_externsheet."""
    bk = Book()
    bk.logfile = sys.stdout
    bk.verbosity = verbosity
    bk.biff_version = biff_version
    bk.encoding = 'ascii'
    bk.encoding_override = 'ascii'
    bk.codepage = 1252
    bk._extnsht_count = 0
    bk._externsheet_info = []
    bk._externsheet_type_b57 = []
    bk._extnsht_name_from_num = {}
    return bk


def _make_biff8_data(refs):
    """Build EXTERNSHEET data for BIFF8 with given list of (record, first, last) tuples."""
    data = struct.pack('<H', len(refs))
    for r in refs:
        data += struct.pack('<HHH', *r)
    return data


class TestHandleExternsheetBiff8:

    def test_increments_extnsht_count(self):
        """handle_externsheet increments _extnsht_count."""
        bk = _make_book(biff_version=80)
        data = _make_biff8_data([])
        bk.handle_externsheet(data)
        assert bk._extnsht_count == 1

    def test_zero_refs_no_externsheet_info(self):
        """handle_externsheet with zero refs leaves _externsheet_info empty."""
        bk = _make_book(biff_version=80)
        data = _make_biff8_data([])
        bk.handle_externsheet(data)
        assert bk._externsheet_info == []

    def test_single_ref_appended(self):
        """handle_externsheet appends a single externsheet info tuple."""
        bk = _make_book(biff_version=80)
        data = _make_biff8_data([(0, 0, 0)])
        bk.handle_externsheet(data)
        assert len(bk._externsheet_info) == 1
        assert bk._externsheet_info[0] == (0, 0, 0)

    def test_multiple_refs_appended(self):
        """handle_externsheet appends multiple externsheet info tuples."""
        bk = _make_book(biff_version=80)
        refs = [(0, 0, 0), (1, 2, 3), (2, 5, 7)]
        data = _make_biff8_data(refs)
        bk.handle_externsheet(data)
        assert len(bk._externsheet_info) == 3
        assert bk._externsheet_info[1] == (1, 2, 3)

    def test_verbosity2_logs_refs(self):
        """handle_externsheet with verbosity >= 2 logs each ref."""
        bk = _make_book(biff_version=80, verbosity=2)
        data = _make_biff8_data([(0, 1, 2)])
        bk.handle_externsheet(data)
        assert bk._externsheet_info[0] == (0, 1, 2)

    def test_continue_record_extends_data(self, mocker):
        """handle_externsheet reads CONTINUE record when data is too short."""
        bk = _make_book(biff_version=80)
        # Claim 2 refs but only provide partial data (header + 1 ref instead of 2)
        num_refs = 2
        partial_data = struct.pack('<H', num_refs) + struct.pack('<HHH', 0, 0, 0)
        continuation = struct.pack('<HHH', 1, 2, 3)
        mocker.patch.object(bk, 'get_record_parts', return_value=(XL_CONTINUE, len(continuation), continuation))
        bk.handle_externsheet(partial_data)
        assert len(bk._externsheet_info) == 2

    def test_continue_record_with_verbosity1(self, mocker):
        """handle_externsheet logs when data is too short and verbosity >= 1."""
        bk = _make_book(biff_version=80, verbosity=1)
        num_refs = 2
        partial_data = struct.pack('<H', num_refs) + struct.pack('<HHH', 0, 0, 0)
        continuation = struct.pack('<HHH', 1, 2, 3)
        mocker.patch.object(bk, 'get_record_parts', return_value=(XL_CONTINUE, len(continuation), continuation))
        bk.handle_externsheet(partial_data)
        assert len(bk._externsheet_info) == 2

    def test_missing_continue_raises(self, mocker):
        """handle_externsheet raises XLRDError when CONTINUE is missing."""
        bk = _make_book(biff_version=80)
        num_refs = 2
        partial_data = struct.pack('<H', num_refs) + struct.pack('<HHH', 0, 0, 0)
        # Return a non-CONTINUE record code
        mocker.patch.object(bk, 'get_record_parts', return_value=(0x00, 0, b''))
        with pytest.raises(XLRDError):
            bk.handle_externsheet(partial_data)


class TestHandleExternsheetBiff7:

    def test_type1_appended(self):
        """handle_externsheet (biff < 8) appends ty=1 (Encoded URL)."""
        bk = _make_book(biff_version=70)
        data = struct.pack('<BB', 0, 1)
        bk.handle_externsheet(data)
        assert bk._externsheet_type_b57 == [1]

    def test_type2_appended(self):
        """handle_externsheet (biff < 8) appends ty=2 (Current sheet)."""
        bk = _make_book(biff_version=70)
        data = struct.pack('<BB', 0, 2)
        bk.handle_externsheet(data)
        assert bk._externsheet_type_b57 == [2]

    def test_type3_sets_sheet_name(self):
        """handle_externsheet (biff < 8) ty=3 stores sheet name."""
        bk = _make_book(biff_version=70)
        name = b'Sheet1'
        nc = len(name)
        data = struct.pack('<BB', nc, 3) + name
        bk.handle_externsheet(data)
        assert bk._extnsht_name_from_num[1] == 'Sheet1'
        assert bk._externsheet_type_b57 == [3]

    def test_type4_appended(self):
        """handle_externsheet (biff < 8) appends ty=4 (Nonspecific sheet)."""
        bk = _make_book(biff_version=70)
        data = struct.pack('<BB', 0, 4)
        bk.handle_externsheet(data)
        assert bk._externsheet_type_b57 == [4]

    def test_unknown_type_becomes_zero(self):
        """handle_externsheet (biff < 8) replaces unknown ty with 0."""
        bk = _make_book(biff_version=70)
        data = struct.pack('<BB', 0, 9)
        bk.handle_externsheet(data)
        assert bk._externsheet_type_b57 == [0]

    def test_type0_becomes_zero(self):
        """handle_externsheet (biff < 8) replaces ty=0 with 0 (out of 1-4 range)."""
        bk = _make_book(biff_version=70)
        data = struct.pack('<BB', 0, 0)
        bk.handle_externsheet(data)
        assert bk._externsheet_type_b57 == [0]

    def test_verbosity2_logs_info(self):
        """handle_externsheet (biff < 8) with verbosity >= 2 logs info."""
        bk = _make_book(biff_version=70, verbosity=2)
        name = b'MySheet'
        nc = len(name)
        data = struct.pack('<BB', nc, 3) + name
        bk.handle_externsheet(data)
        assert bk._extnsht_name_from_num[1] == 'MySheet'
        assert bk._externsheet_type_b57 == [3]

    def test_verbosity2_type1_logs(self):
        """handle_externsheet (biff < 8) verbosity >= 2 logs for ty=1."""
        bk = _make_book(biff_version=70, verbosity=2)
        data = struct.pack('<BB', 0, 1)
        bk.handle_externsheet(data)
        assert bk._externsheet_type_b57 == [1]

    def test_verbosity2_unknown_type_logs(self):
        """handle_externsheet (biff < 8) verbosity >= 2 logs unknown type."""
        bk = _make_book(biff_version=70, verbosity=2)
        data = struct.pack('<BB', 0, 99)
        bk.handle_externsheet(data)
        assert bk._externsheet_type_b57 == [0]

    def test_increments_extnsht_count(self):
        """handle_externsheet (biff < 8) increments _extnsht_count."""
        bk = _make_book(biff_version=70)
        data = struct.pack('<BB', 0, 1)
        bk.handle_externsheet(data)
        assert bk._extnsht_count == 1

    def test_multiple_calls_accumulate(self):
        """handle_externsheet accumulates entries across multiple calls."""
        bk = _make_book(biff_version=70)
        bk.handle_externsheet(struct.pack('<BB', 0, 1))
        bk.handle_externsheet(struct.pack('<BB', 0, 2))
        assert bk._extnsht_count == 2
        assert bk._externsheet_type_b57 == [1, 2]
