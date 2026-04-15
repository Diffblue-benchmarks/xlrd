"""Tests for Book.handle_name targeting uncovered lines."""
import io
from struct import pack

import pytest

from xlrd.book import Book


def _make_book(biff_version=80, verbosity=0, encoding_override=None, codepage=None):
    """Create a Book with minimum state needed for handle_name."""
    bk = Book()
    bk.logfile = io.StringIO()
    bk.verbosity = verbosity
    bk.biff_version = biff_version
    bk.encoding_override = encoding_override
    bk.codepage = codepage
    bk.raw_user_name = False
    bk.name_obj_list = []
    return bk


def _make_name_data(option_flags=0, kb_shortcut=0, name_len=3,
                    fmla_len=0, extsht_index=0, sheet_index=0,
                    name_bytes=b'\x00foo'):
    """Build a NAME record data buffer."""
    header = pack("<HBBHHH4B",
                  option_flags, kb_shortcut, name_len,
                  fmla_len, extsht_index, sheet_index,
                  0, 0, 0, 0)
    return header + name_bytes


# ---------------------------------------------------------------------------
# Early return for BIFF version < 50 (lines 948-949)
# ---------------------------------------------------------------------------

class TestHandleNameEarlyReturn:

    def test_biff_version_below_50_returns_immediately(self):
        bk = _make_book(biff_version=45)
        bk.handle_name(b'')
        assert bk.name_obj_list == []

    def test_biff_version_20_returns_immediately(self):
        bk = _make_book(biff_version=20)
        bk.handle_name(b'')
        assert bk.name_obj_list == []


# ---------------------------------------------------------------------------
# BIFF8 path: unpack_unicode_update_pos (lines 979-980)
# ---------------------------------------------------------------------------

class TestHandleNameBiff8:

    def test_biff8_regular_name_appended_to_name_obj_list(self):
        bk = _make_book(biff_version=80)
        # option_flags=0, name_len=3, data after header: options=0 (compressed) + b'foo'
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'\x00foo')
        bk.handle_name(data)
        assert len(bk.name_obj_list) == 1
        nobj = bk.name_obj_list[0]
        assert nobj.name == 'foo'

    def test_biff8_name_index_assigned_correctly(self):
        bk = _make_book(biff_version=80)
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'\x00foo')
        bk.handle_name(data)
        assert bk.name_obj_list[0].name_index == 0

    def test_biff8_second_name_has_incremented_index(self):
        bk = _make_book(biff_version=80)
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'\x00foo')
        bk.handle_name(data)
        data2 = _make_name_data(option_flags=0, name_len=3, name_bytes=b'\x00bar')
        bk.handle_name(data2)
        assert bk.name_obj_list[1].name_index == 1

    def test_biff8_option_flags_set_on_name_obj(self):
        bk = _make_book(biff_version=80)
        # macro flag (bit 3) set
        data = _make_name_data(option_flags=0x08, name_len=3, name_bytes=b'\x00foo')
        bk.handle_name(data)
        nobj = bk.name_obj_list[0]
        assert nobj.macro == 1

    def test_biff8_sheet_index_and_extsht_stored(self):
        bk = _make_book(biff_version=80)
        data = _make_name_data(option_flags=0, name_len=3, extsht_index=2,
                               sheet_index=1, name_bytes=b'\x00foo')
        bk.handle_name(data)
        nobj = bk.name_obj_list[0]
        assert nobj.extn_sheet_num == 2
        assert nobj.excel_sheet_index == 1

    def test_biff8_raw_formula_and_basic_formula_len_stored(self):
        bk = _make_book(biff_version=80)
        data = _make_name_data(option_flags=0, name_len=3, fmla_len=4,
                               name_bytes=b'\x00foo')
        bk.handle_name(data)
        nobj = bk.name_obj_list[0]
        assert nobj.basic_formula_len == 4
        assert nobj.evaluated == 0

    def test_biff8_scope_initially_none(self):
        bk = _make_book(biff_version=80)
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'\x00foo')
        bk.handle_name(data)
        assert bk.name_obj_list[0].scope is None


# ---------------------------------------------------------------------------
# BIFF < 80 but >= 50: unpack_string_update_pos (lines 977-978)
# ---------------------------------------------------------------------------

class TestHandleNameBiff7:

    def test_biff7_name_parsed_using_encoding(self):
        bk = _make_book(biff_version=70, encoding_override='latin-1')
        # For BIFF < 80 the name bytes are raw (no options byte)
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'bar')
        bk.handle_name(data)
        assert len(bk.name_obj_list) == 1
        assert bk.name_obj_list[0].name == 'bar'

    def test_biff5_name_parsed_correctly(self):
        bk = _make_book(biff_version=50, encoding_override='latin-1')
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'baz')
        bk.handle_name(data)
        assert bk.name_obj_list[0].name == 'baz'


# ---------------------------------------------------------------------------
# Builtin name lookup (lines 991-993)
# ---------------------------------------------------------------------------

class TestHandleNameBuiltin:

    def test_builtin_flag_triggers_name_lookup(self):
        bk = _make_book(biff_version=80)
        # option_flags = 0x20 → builtin=1; name_len=1; builtin code \x06 = Print_Area
        data = _make_name_data(option_flags=0x20, name_len=1, name_bytes=b'\x00\x06')
        bk.handle_name(data)
        nobj = bk.name_obj_list[0]
        assert nobj.builtin == 1
        assert nobj.name == 'Print_Area'

    def test_unknown_builtin_code_returns_question_marks(self):
        bk = _make_book(biff_version=80)
        # option_flags = 0x20 → builtin=1; unknown code \xff
        data = _make_name_data(option_flags=0x20, name_len=1, name_bytes=b'\x00\xff')
        bk.handle_name(data)
        nobj = bk.name_obj_list[0]
        assert nobj.builtin == 1
        assert nobj.name == '??Unknown??'

    def test_non_builtin_name_not_looked_up(self):
        bk = _make_book(biff_version=80)
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'\x00foo')
        bk.handle_name(data)
        assert bk.name_obj_list[0].name == 'foo'


# ---------------------------------------------------------------------------
# Verbose / debug output (lines 984-989, 993, 998-1003)
# ---------------------------------------------------------------------------

class TestHandleNameVerbose:

    def test_verbose_logs_name_record_info(self):
        bk = _make_book(biff_version=80, verbosity=2)
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'\x00foo')
        bk.handle_name(data)
        output = bk.logfile.getvalue()
        assert 'NAME' in output

    def test_verbose_builtin_logs_builtin_name(self):
        bk = _make_book(biff_version=80, verbosity=2)
        data = _make_name_data(option_flags=0x20, name_len=1, name_bytes=b'\x00\x06')
        bk.handle_name(data)
        output = bk.logfile.getvalue()
        assert 'builtin' in output

    def test_verbose_dumps_name_object(self):
        bk = _make_book(biff_version=80, verbosity=2)
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'\x00foo')
        bk.handle_name(data)
        output = bk.logfile.getvalue()
        assert 'handle_name' in output

    def test_non_verbose_produces_no_output(self):
        bk = _make_book(biff_version=80, verbosity=0)
        data = _make_name_data(option_flags=0, name_len=3, name_bytes=b'\x00foo')
        bk.handle_name(data)
        assert bk.logfile.getvalue() == ''
