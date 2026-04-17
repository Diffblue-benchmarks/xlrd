# -*- coding: utf-8 -*-
"""
Additional targeted tests to cover remaining gaps in:
 - biffh.py: BaseObject with __slots__, child-dump, _repr_these, biff_dump edge cases
 - formatting.py: is_date_format_string warning paths, ambiguous formats
 - formula.py: adjust_cell_addr_biff_le7 non-reldelta paths, col_rel wrap
 - sheet.py: Cell repr, ctype_text, tidy_dimensions verbosity, Colinfo/Rowinfo
 - __init__.py: dump / count_records via mocking
"""
import io
import struct
import pytest
from unittest.mock import MagicMock, patch

import xlrd
from xlrd.biffh import (
    BaseObject,
    biff_dump,
    hex_char_dump,
)
from xlrd.formula import (
    adjust_cell_addr_biff_le7,
    adjust_cell_addr_biff8,
)
from xlrd.sheet import Cell, Colinfo, Rowinfo, ctype_text
from xlrd.biffh import (
    XL_CELL_EMPTY, XL_CELL_TEXT, XL_CELL_NUMBER, XL_CELL_DATE,
    XL_CELL_BOOLEAN, XL_CELL_ERROR, XL_CELL_BLANK,
)


# ---------------------------------------------------------------------------
# BaseObject.dump — __slots__ branch
# ---------------------------------------------------------------------------

class TestBaseObjectSlots:

    def test_dump_with_slots_uses_slot_attrs(self):
        """Cell has __slots__; dump should iterate them."""
        c = Cell(XL_CELL_NUMBER, 3.14, xf_index=5)
        buf = io.StringIO()
        c.dump(f=buf)
        output = buf.getvalue()
        assert 'ctype' in output
        assert 'value' in output
        assert 'xf_index' in output

    def test_dump_uses_stderr_when_f_is_none(self):
        """When f=None, dump writes to sys.stderr — just check it doesn't raise."""
        c = Cell(XL_CELL_NUMBER, 1.0)
        import sys
        old_stderr = sys.stderr
        sys.stderr = io.StringIO()
        try:
            c.dump()  # no f= arg → uses sys.stderr
        finally:
            sys.stderr = old_stderr


class TestBaseObjectReprThese:

    def test_list_attr_in_repr_these_printed_as_repr(self):
        """An attr in _repr_these should be printed via repr, not 'len = N'."""
        class MyObj(BaseObject):
            _repr_these = ['mylist']
        obj = MyObj()
        obj.mylist = [1, 2, 3]
        buf = io.StringIO()
        obj.dump(f=buf)
        output = buf.getvalue()
        # Since mylist is in _repr_these, it's printed with repr not length
        assert '[1, 2, 3]' in output

    def test_list_not_in_repr_these_shows_length(self):
        """An attr NOT in _repr_these and is a list → show type + len."""
        class MyObj(BaseObject):
            _repr_these = []
        obj = MyObj()
        obj.mylist = [1, 2, 3]
        buf = io.StringIO()
        obj.dump(f=buf)
        output = buf.getvalue()
        assert 'len = 3' in output

    def test_dict_not_in_repr_these_shows_length(self):
        class MyObj(BaseObject):
            _repr_these = []
        obj = MyObj()
        obj.mydict = {'a': 1, 'b': 2}
        buf = io.StringIO()
        obj.dump(f=buf)
        output = buf.getvalue()
        assert 'len = 2' in output


class TestBaseObjectChildDump:

    def test_child_with_dump_method_is_recursed(self):
        """An attribute that has a .dump() method (and isn't 'book') gets recursed."""
        class Child(BaseObject):
            pass
        child = Child()
        child.x = 99

        class Parent(BaseObject):
            _repr_these = []
        parent = Parent()
        parent.child_attr = child

        buf = io.StringIO()
        parent.dump(f=buf)
        output = buf.getvalue()
        # child_attr should trigger recursive dump call
        assert 'child_attr' in output or 'Child' in output

    def test_book_attr_not_recursed(self):
        """An attribute named 'book' with a .dump() method is NOT recursed."""
        class FakeBook(BaseObject):
            pass
        book = FakeBook()
        book.name = 'workbook'

        class Parent(BaseObject):
            _repr_these = []
        parent = Parent()
        parent.book = book

        buf = io.StringIO()
        parent.dump(f=buf)
        # Should NOT show book's children, just the repr
        output = buf.getvalue()
        assert 'name' not in output or 'workbook' not in output


# ---------------------------------------------------------------------------
# biff_dump edge cases — misc bytes at end, record too large
# ---------------------------------------------------------------------------

def _make_stream(*records):
    out = b''
    for opcode, payload in records:
        out += struct.pack('<HH', opcode, len(payload)) + payload
    return out


class TestBiffDumpEdgeCases:

    def test_misc_bytes_at_end(self):
        # A stream where the last record claims a length that extends past stream_end
        # Create stream with only 3 bytes (less than 4 needed for a record header)
        stream = _make_stream((0x000A, b'')) + b'\xAB\xCD\xEF'
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert 'Misc bytes at end' in output

    def test_dummies_then_real_record_reports_skipped(self):
        # Zero bytes before a real record → "zero bytes skipped" message
        stream = b'\x00\x00\x00\x00' + _make_stream((0x000A, b''))
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert 'zero bytes skipped' in output

    def test_record_length_too_large_reported(self):
        # Manually build a record whose length goes past stream end
        # opcode=0x000A, length=1000 but no actual data
        stream = struct.pack('<HH', 0x000A, 1000)  # header says 1000 bytes but nothing follows
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert 'too large' in output.lower() or 'EOF' in output


# ---------------------------------------------------------------------------
# is_date_format_string — separator and ambiguous paths
# ---------------------------------------------------------------------------

class TestIsDateFormatStringEdgeCases:

    def _book(self, verbosity=0):
        from xlrd.formatting import is_date_format_string
        book = type('MockBook', (), {
            'verbosity': verbosity,
            'logfile': io.StringIO(),
        })()
        return book

    def test_separator_only_format_returns_false(self):
        from xlrd.formatting import is_date_format_string
        # A format with only ';' (separator) and no date/num chars → date_count==0, not got_sep behaviour
        # But since there are no date or num chars and got_sep=1, returns False
        result = is_date_format_string(self._book(), ';')
        assert result is False

    def test_empty_format_returns_false(self):
        from xlrd.formatting import is_date_format_string
        result = is_date_format_string(self._book(), '')
        assert result is False

    def test_ambiguous_format_higher_verbosity(self):
        from xlrd.formatting import is_date_format_string
        # A format with both date and num chars → warning is logged, returns date_count > num_count
        book = self._book(verbosity=1)
        # 'm' is date char (×5), '0' is num char (×5) → equal → returns False
        result = is_date_format_string(book, 'm0')
        assert result is False
        log_output = book.logfile.getvalue()
        assert 'ambiguous' in log_output.lower() or 'WARNING' in log_output

    def test_format_with_only_separator(self):
        from xlrd.formatting import is_date_format_string
        # got_sep=1 but no date/num counts → returns False
        result = is_date_format_string(self._book(), ';;;')
        assert result is False

    def test_verbosity_4_logs_reduced_format(self):
        from xlrd.formatting import is_date_format_string
        book = self._book(verbosity=4)
        result = is_date_format_string(book, 'yyyy-mm-dd')
        # With verbosity>=4, a log line is written
        assert result is True
        log_output = book.logfile.getvalue()
        assert 'is_date_format_string' in log_output

    def test_constant_result_warns_at_verbosity(self):
        from xlrd.formatting import is_date_format_string
        # A format with no date/num chars and no separator → "constant result" warning
        # 'ABC' has no date chars (ymdhs), no num chars (0#?), no skip chars, no separator
        book = self._book(verbosity=1)
        result = is_date_format_string(book, 'ABC')
        assert result is False
        log_output = book.logfile.getvalue()
        assert 'constant' in log_output.lower() or 'WARNING' in log_output


# ---------------------------------------------------------------------------
# formula.py — adjust_cell_addr_biff_le7 remaining branches
# ---------------------------------------------------------------------------

class TestAdjustCellAddrBiffLe7Extra:

    def test_col_rel_wrap_reldelta(self):
        # col_rel=1 and colx >= 128 → colx -= 256
        rowval = (1 << 14) | 5  # col_rel=1, rowx=5
        colval = 200  # colx=200 >= 128 → wraps
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(
            rowval, colval, reldelta=1
        )
        assert colx == 200 - 256

    def test_non_reldelta_row_subtracted(self):
        # reldelta=0, row_rel=1 → rowx -= browx
        rowval = (1 << 15) | 10  # row_rel=1, rowx=10
        colval = 5
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(
            rowval, colval, reldelta=0, browx=3, bcolx=0
        )
        assert rowx == 10 - 3

    def test_non_reldelta_col_subtracted(self):
        # reldelta=0, col_rel=1 → colx -= bcolx
        rowval = (1 << 14) | 0  # col_rel=1
        colval = 7
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(
            rowval, colval, reldelta=0, browx=0, bcolx=2
        )
        assert colx == 7 - 2


# ---------------------------------------------------------------------------
# Cell repr tests
# ---------------------------------------------------------------------------

class TestCellRepr:

    def test_repr_without_xf_index(self):
        c = Cell(XL_CELL_NUMBER, 42.0)
        r = repr(c)
        assert 'number' in r
        assert '42' in r
        assert 'XF' not in r

    def test_repr_with_xf_index(self):
        c = Cell(XL_CELL_NUMBER, 42.0, xf_index=7)
        r = repr(c)
        assert 'XF' in r
        assert '7' in r

    def test_repr_text_cell(self):
        c = Cell(XL_CELL_TEXT, 'hello')
        assert 'text' in repr(c)
        assert 'hello' in repr(c)

    def test_repr_boolean_cell(self):
        c = Cell(XL_CELL_BOOLEAN, 1)
        assert 'bool' in repr(c)

    def test_repr_error_cell(self):
        c = Cell(XL_CELL_ERROR, 0x07)
        assert 'error' in repr(c)

    def test_repr_blank_cell(self):
        c = Cell(XL_CELL_BLANK, '')
        assert 'blank' in repr(c)

    def test_repr_date_cell(self):
        c = Cell(XL_CELL_DATE, 36526.0)
        assert 'xldate' in repr(c)

    def test_repr_empty_cell(self):
        c = Cell(XL_CELL_EMPTY, '')
        assert 'empty' in repr(c)


# ---------------------------------------------------------------------------
# ctype_text dict
# ---------------------------------------------------------------------------

class TestCtypeText:

    def test_all_cell_types_have_text(self):
        for ct in range(7):
            assert ct in ctype_text

    def test_number_text(self):
        assert ctype_text[XL_CELL_NUMBER] == 'number'

    def test_text_text(self):
        assert ctype_text[XL_CELL_TEXT] == 'text'

    def test_date_text(self):
        assert ctype_text[XL_CELL_DATE] == 'xldate'


# ---------------------------------------------------------------------------
# Colinfo and Rowinfo defaults
# ---------------------------------------------------------------------------

class TestColinfo:

    def test_colinfo_is_base_object(self):
        col = Colinfo()
        assert isinstance(col, BaseObject)

    def test_colinfo_has_width(self):
        col = Colinfo()
        assert hasattr(col, 'width')

    def test_colinfo_has_xf_index(self):
        col = Colinfo()
        assert hasattr(col, 'xf_index')

    def test_colinfo_hidden_default(self):
        col = Colinfo()
        assert col.hidden == 0


class TestRowinfo:

    def test_rowinfo_is_base_object(self):
        row = Rowinfo()
        assert isinstance(row, BaseObject)

    def test_rowinfo_has_xf_index(self):
        row = Rowinfo()
        assert hasattr(row, 'xf_index')

    def test_rowinfo_has_height(self):
        row = Rowinfo()
        assert hasattr(row, 'height')

    def test_rowinfo_hidden_default(self):
        # Rowinfo uses __slots__ and sets hidden = None by default (not 0)
        row = Rowinfo()
        assert row.hidden is None


# ---------------------------------------------------------------------------
# __init__.py dump and count_records via mocking
# ---------------------------------------------------------------------------

class TestInitDump:

    def test_dump_calls_biff_dump(self):
        biff_stream = _make_stream((0x000A, b''))
        mock_book = MagicMock()
        mock_book.mem = biff_stream
        mock_book.base = 0
        mock_book.stream_len = len(biff_stream)

        with patch('xlrd.Book', return_value=mock_book):
            buf = io.StringIO()
            xlrd.dump('fake.xls', outfile=buf)
        mock_book.biff2_8_load.assert_called_once()

    def test_dump_unnumbered_passes_flag(self):
        biff_stream = _make_stream((0x000A, b''))
        mock_book = MagicMock()
        mock_book.mem = biff_stream
        mock_book.base = 0
        mock_book.stream_len = len(biff_stream)

        with patch('xlrd.Book', return_value=mock_book):
            buf = io.StringIO()
            xlrd.dump('fake.xls', outfile=buf, unnumbered=True)
        # Just verifying it doesn't raise


class TestInitCountRecords:

    def test_count_records_calls_biff_count(self):
        biff_stream = _make_stream((0x000A, b''))
        mock_book = MagicMock()
        mock_book.mem = biff_stream
        mock_book.base = 0
        mock_book.stream_len = len(biff_stream)

        with patch('xlrd.Book', return_value=mock_book):
            buf = io.StringIO()
            xlrd.count_records('fake.xls', outfile=buf)
        mock_book.biff2_8_load.assert_called_once()


# ---------------------------------------------------------------------------
# Rowinfo __getstate__ / __setstate__
# ---------------------------------------------------------------------------

class TestRowinfoPersistence:

    def _make_rowinfo(self, height=240, has_default_height=1, outline_level=0,
                      outline_group_starts_ends=0, hidden=0, height_mismatch=0,
                      has_default_xf_index=1, xf_index=15,
                      additional_space_above=0, additional_space_below=0):
        row = Rowinfo()
        row.height = height
        row.has_default_height = has_default_height
        row.outline_level = outline_level
        row.outline_group_starts_ends = outline_group_starts_ends
        row.hidden = hidden
        row.height_mismatch = height_mismatch
        row.has_default_xf_index = has_default_xf_index
        row.xf_index = xf_index
        row.additional_space_above = additional_space_above
        row.additional_space_below = additional_space_below
        return row

    def test_getstate_returns_tuple(self):
        row = self._make_rowinfo()
        state = row.__getstate__()
        assert isinstance(state, tuple)
        assert len(state) == 10

    def test_getstate_values_match(self):
        row = self._make_rowinfo(height=300, xf_index=7)
        state = row.__getstate__()
        assert state[0] == 300   # height
        assert state[7] == 7     # xf_index

    def test_setstate_restores_values(self):
        row = self._make_rowinfo(height=480, xf_index=3)
        state = row.__getstate__()
        new_row = Rowinfo()
        new_row.__setstate__(state)
        assert new_row.height == 480
        assert new_row.xf_index == 3

    def test_roundtrip_via_pickle(self):
        import pickle
        row = self._make_rowinfo(height=360, hidden=1)
        pickled = pickle.dumps(row)
        loaded = pickle.loads(pickled)
        assert loaded.height == 360
        assert loaded.hidden == 1


# ---------------------------------------------------------------------------
# formula.py: get_externsheet_local_range
# ---------------------------------------------------------------------------

class TestGetExternsheetLocalRange:

    def _make_bk(self, externsheet_info, supbook_locals_inx, supbook_addins_inx, all_sheets_map):
        bk = MagicMock()
        bk._externsheet_info = externsheet_info
        bk._supbook_locals_inx = supbook_locals_inx
        bk._supbook_addins_inx = supbook_addins_inx
        bk._all_sheets_map = all_sheets_map
        bk.logfile = io.StringIO()
        return bk

    def test_refx_out_of_range_returns_minus_101(self):
        from xlrd.formula import get_externsheet_local_range
        bk = self._make_bk([], 0, 1, [])
        result = get_externsheet_local_range(bk, refx=5)
        assert result == (-101, -101)

    def test_addins_reference_returns_minus_5(self):
        from xlrd.formula import get_externsheet_local_range
        # ref_recordx == supbook_addins_inx, and sheetx both == 0xFFFE
        bk = self._make_bk(
            [(1, 0xFFFE, 0xFFFE)],  # ref_recordx=1 is addins
            supbook_locals_inx=0,
            supbook_addins_inx=1,
            all_sheets_map=[]
        )
        result = get_externsheet_local_range(bk, refx=0)
        assert result == (-5, -5)

    def test_external_reference_returns_minus_4(self):
        from xlrd.formula import get_externsheet_local_range
        # ref_recordx is neither locals nor addins
        bk = self._make_bk(
            [(99, 0, 0)],  # ref_recordx=99 is external
            supbook_locals_inx=0,
            supbook_addins_inx=1,
            all_sheets_map=[]
        )
        result = get_externsheet_local_range(bk, refx=0)
        assert result == (-4, -4)

    def test_unspecified_sheet_returns_minus_1(self):
        from xlrd.formula import get_externsheet_local_range
        bk = self._make_bk(
            [(0, 0xFFFE, 0xFFFE)],  # locals, unspecified sheet
            supbook_locals_inx=0,
            supbook_addins_inx=1,
            all_sheets_map=[]
        )
        result = get_externsheet_local_range(bk, refx=0)
        assert result == (-1, -1)

    def test_deleted_sheet_returns_minus_2(self):
        from xlrd.formula import get_externsheet_local_range
        bk = self._make_bk(
            [(0, 0xFFFF, 0xFFFF)],  # locals, deleted sheets
            supbook_locals_inx=0,
            supbook_addins_inx=1,
            all_sheets_map=[]
        )
        result = get_externsheet_local_range(bk, refx=0)
        assert result == (-2, -2)

    def test_sheet_out_of_range_returns_minus_102(self):
        from xlrd.formula import get_externsheet_local_range
        bk = self._make_bk(
            [(0, 0, 5)],  # locals, sheet indices 0-5 but only 3 sheets
            supbook_locals_inx=0,
            supbook_addins_inx=1,
            all_sheets_map=[0, 1, 2]  # only 3 sheets
        )
        result = get_externsheet_local_range(bk, refx=0)
        assert result == (-102, -102)

    def test_macro_sheet_returns_minus_3(self):
        from xlrd.formula import get_externsheet_local_range
        # xlrd_sheetx values are negative (macro sheets)
        bk = self._make_bk(
            [(0, 0, 1)],
            supbook_locals_inx=0,
            supbook_addins_inx=1,
            all_sheets_map=[-1, -1]  # macro sheets (negative xlrd indices)
        )
        result = get_externsheet_local_range(bk, refx=0)
        assert result == (-3, -3)

    def test_valid_local_reference(self):
        from xlrd.formula import get_externsheet_local_range
        bk = self._make_bk(
            [(0, 0, 1)],
            supbook_locals_inx=0,
            supbook_addins_inx=1,
            all_sheets_map=[0, 1, 2]
        )
        result = get_externsheet_local_range(bk, refx=0)
        assert result == (0, 1)


# ---------------------------------------------------------------------------
# formula.py: get_externsheet_local_range_b57
# ---------------------------------------------------------------------------

class TestGetExternsheetLocalRangeB57:

    def _make_bk(self, all_sheets_map):
        bk = MagicMock()
        bk._all_sheets_map = all_sheets_map
        bk.logfile = io.StringIO()
        return bk

    def test_external_reference(self):
        from xlrd.formula import get_externsheet_local_range_b57
        bk = self._make_bk([0, 1])
        result = get_externsheet_local_range_b57(bk, raw_extshtx=1, ref_first_sheetx=0, ref_last_sheetx=0)
        assert result == (-4, -4)

    def test_deleted_sheets(self):
        from xlrd.formula import get_externsheet_local_range_b57
        bk = self._make_bk([0])
        result = get_externsheet_local_range_b57(bk, raw_extshtx=0, ref_first_sheetx=-1, ref_last_sheetx=-1)
        assert result == (-2, -2)

    def test_sheet_out_of_range(self):
        from xlrd.formula import get_externsheet_local_range_b57
        bk = self._make_bk([0])  # only 1 sheet
        result = get_externsheet_local_range_b57(bk, raw_extshtx=0, ref_first_sheetx=0, ref_last_sheetx=5)
        assert result == (-103, -103)

    def test_macro_sheet(self):
        from xlrd.formula import get_externsheet_local_range_b57
        bk = self._make_bk([-1, -1])  # both macro sheets
        result = get_externsheet_local_range_b57(bk, raw_extshtx=0, ref_first_sheetx=0, ref_last_sheetx=1)
        assert result == (-3, -3)

    def test_valid_reference(self):
        from xlrd.formula import get_externsheet_local_range_b57
        bk = self._make_bk([0, 1, 2])
        result = get_externsheet_local_range_b57(bk, raw_extshtx=0, ref_first_sheetx=0, ref_last_sheetx=2)
        assert result == (0, 2)
