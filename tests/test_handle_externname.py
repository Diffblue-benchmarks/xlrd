"""
Tests targeting uncovered lines in xlrd/book.py Book.handle_externname.
"""
import struct
import sys
import pytest

from xlrd.book import Book, SUPBOOK_ADDIN, SUPBOOK_INTERNAL


def _make_book(biff_version=80, verbosity=0):
    """Create a minimal Book instance with attributes needed by handle_externname."""
    bk = Book()
    bk.logfile = sys.stdout
    bk.verbosity = verbosity
    bk.biff_version = biff_version
    bk._supbook_types = []
    bk.addin_func_names = []
    return bk


def _make_externname_data(name, option_flags=0, other_info=0):
    """Build EXTERNNAME record data for biff8+."""
    header = struct.pack('<HI', option_flags, other_info)
    encoded = name.encode('latin_1')
    name_part = struct.pack('<B', len(encoded)) + b'\x00' + encoded
    return header + name_part


class TestHandleExternname:

    def test_biff8_non_addin_does_not_append_name(self):
        """handle_externname does not append name for non-SUPBOOK_ADDIN type."""
        bk = _make_book(biff_version=80)
        bk._supbook_types = [SUPBOOK_INTERNAL]
        data = _make_externname_data('MyFunc')
        bk.handle_externname(data)
        assert bk.addin_func_names == []

    def test_biff8_addin_appends_name(self):
        """handle_externname appends name to addin_func_names when supbook type is SUPBOOK_ADDIN."""
        bk = _make_book(biff_version=80)
        bk._supbook_types = [SUPBOOK_ADDIN]
        data = _make_externname_data('MyAddinFunc')
        bk.handle_externname(data)
        assert bk.addin_func_names == ['MyAddinFunc']

    def test_biff8_addin_appends_multiple_names(self):
        """handle_externname appends multiple names on successive calls."""
        bk = _make_book(biff_version=80)
        bk._supbook_types = [SUPBOOK_ADDIN]
        bk.handle_externname(_make_externname_data('FuncA'))
        bk.handle_externname(_make_externname_data('FuncB'))
        assert bk.addin_func_names == ['FuncA', 'FuncB']

    def test_biff8_with_verbosity_logs(self):
        """handle_externname with verbosity >= 2 logs the EXTERNNAME info."""
        bk = _make_book(biff_version=80, verbosity=2)
        bk._supbook_types = [SUPBOOK_INTERNAL]
        data = _make_externname_data('LoggedFunc')
        bk.handle_externname(data)
        assert bk.addin_func_names == []

    def test_biff8_addin_with_verbosity_logs_and_appends(self):
        """handle_externname with verbosity >= 2 and SUPBOOK_ADDIN both appends and logs."""
        bk = _make_book(biff_version=80, verbosity=2)
        bk._supbook_types = [SUPBOOK_ADDIN]
        data = _make_externname_data('AddinVerbose')
        bk.handle_externname(data)
        assert bk.addin_func_names == ['AddinVerbose']

    def test_biff_version_below_80_skips_processing(self):
        """handle_externname skips processing for biff_version < 80."""
        bk = _make_book(biff_version=70)
        bk._supbook_types = [SUPBOOK_ADDIN]
        data = _make_externname_data('ShouldNotAdd')
        bk.handle_externname(data)
        assert bk.addin_func_names == []

    def test_biff8_extra_data_parsed(self):
        """handle_externname correctly handles data with extra bytes after name."""
        bk = _make_book(biff_version=80)
        bk._supbook_types = [SUPBOOK_INTERNAL]
        data = _make_externname_data('WithExtra') + b'\xDE\xAD\xBE\xEF'
        bk.handle_externname(data)
        assert bk.addin_func_names == []
