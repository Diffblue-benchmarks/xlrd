# -*- coding: utf-8 -*-
"""Unit tests for xlrd.sheet module."""

import sys
import pytest
from struct import pack

from xlrd.sheet import (
    Cell, Colinfo, Rowinfo, Sheet,
    XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_EMPTY,
    XL_CELL_NUMBER, XL_CELL_TEXT,
    unpack_RK,
)
from xlrd.biffh import XLRDError


# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------

class MockBook:
    biff_version = 80
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    _xf_index_to_xl_type_map = {}
    _sheet_visibility = [0, 0, 0, 0, 0]


class FmtBook(MockBook):
    formatting_info = True


class RaggedBook(MockBook):
    ragged_rows = True


@pytest.fixture
def book():
    return MockBook()


@pytest.fixture
def sheet(book):
    return Sheet(book, 0, 'TestSheet', 0)


@pytest.fixture
def fmt_book():
    return FmtBook()


@pytest.fixture
def fmt_sheet(fmt_book):
    return Sheet(fmt_book, 0, 'FmtSheet', 0)


@pytest.fixture
def ragged_book():
    return RaggedBook()


@pytest.fixture
def ragged_sheet(ragged_book):
    return Sheet(ragged_book, 0, 'RaggedSheet', 0)


def _make_sheet_with_data(book=None):
    """Create a 3x3 sheet with test data."""
    if book is None:
        book = MockBook()
    sh = Sheet(book, 0, 'DataSheet', 0)
    sh.put_cell(0, 0, XL_CELL_TEXT, 'hello', -1)
    sh.put_cell(0, 1, XL_CELL_NUMBER, 42.0, -1)
    sh.put_cell(0, 2, XL_CELL_BOOLEAN, 1, -1)
    sh.put_cell(1, 0, XL_CELL_NUMBER, 3.14, -1)
    sh.put_cell(1, 1, XL_CELL_TEXT, 'world', -1)
    sh.put_cell(1, 2, XL_CELL_EMPTY, '', -1)
    sh.put_cell(2, 0, XL_CELL_DATE, 42000.0, -1)
    sh.put_cell(2, 1, XL_CELL_NUMBER, 0.0, -1)
    sh.put_cell(2, 2, XL_CELL_TEXT, 'end', -1)
    return sh


# ---------------------------------------------------------------------------
# Sheet.__init__ tests
# ---------------------------------------------------------------------------

class TestSheetInit:
    def test_basic_attributes(self, sheet):
        assert sheet.name == 'TestSheet'
        assert sheet.number == 0
        assert sheet.nrows == 0
        assert sheet.ncols == 0
        assert sheet.biff_version == 80

    def test_utter_max_rows_biff8(self, sheet):
        assert sheet.utter_max_rows == 65536

    def test_utter_max_rows_biff7(self, book):
        book.biff_version = 70
        sh = Sheet(book, 0, 'S', 0)
        assert sh.utter_max_rows == 16384

    def test_ragged_rows_puts_ragged(self, ragged_book):
        sh = Sheet(ragged_book, 0, 'S', 0)
        assert sh.put_cell.__func__ is Sheet.put_cell_ragged

    def test_unragged_rows_puts_unragged(self, sheet):
        assert sheet.put_cell.__func__ is Sheet.put_cell_unragged

    def test_initial_cell_storage_empty(self, sheet):
        assert sheet._cell_values == []
        assert sheet._cell_types == []
        assert sheet._cell_xf_indexes == []

    def test_misc_defaults(self, sheet):
        assert sheet.defcolwidth is None
        assert sheet.standardwidth is None
        assert sheet.default_row_height is None
        assert sheet.hyperlink_list == []
        assert sheet.merged_cells == []
        assert sheet.cooked_normal_view_mag_factor == 100
        assert sheet.cooked_page_break_preview_mag_factor == 60


# ---------------------------------------------------------------------------
# Cell tests
# ---------------------------------------------------------------------------

class TestCell:
    def test_cell_init_no_xf(self):
        c = Cell(XL_CELL_TEXT, 'hello')
        assert c.ctype == XL_CELL_TEXT
        assert c.value == 'hello'
        assert c.xf_index is None

    def test_cell_init_with_xf(self):
        c = Cell(XL_CELL_NUMBER, 3.14, 5)
        assert c.ctype == XL_CELL_NUMBER
        assert c.value == 3.14
        assert c.xf_index == 5

    def test_cell_repr_no_xf(self):
        c = Cell(XL_CELL_TEXT, 'hi')
        r = repr(c)
        assert 'text' in r
        assert 'hi' in r
        assert 'XF' not in r

    def test_cell_repr_with_xf(self):
        c = Cell(XL_CELL_NUMBER, 42.0, 3)
        r = repr(c)
        assert 'number' in r
        assert 'XF' in r
        assert '3' in r


# ---------------------------------------------------------------------------
# Sheet.cell / cell_value / cell_type
# ---------------------------------------------------------------------------

class TestCellAccess:
    def test_cell_returns_cell_object(self, sheet):
        sheet.put_cell(0, 0, XL_CELL_TEXT, 'hi', -1)
        c = sheet.cell(0, 0)
        assert isinstance(c, Cell)
        assert c.ctype == XL_CELL_TEXT
        assert c.value == 'hi'
        assert c.xf_index is None

    def test_cell_with_formatting_info(self, fmt_sheet):
        fmt_sheet.put_cell(0, 0, XL_CELL_TEXT, 'hi', 7)
        c = fmt_sheet.cell(0, 0)
        assert c.xf_index == 7

    def test_cell_value(self, sheet):
        sheet.put_cell(0, 0, XL_CELL_NUMBER, 99.9, -1)
        assert sheet.cell_value(0, 0) == 99.9

    def test_cell_type(self, sheet):
        sheet.put_cell(0, 0, XL_CELL_BOOLEAN, 1, -1)
        assert sheet.cell_type(0, 0) == XL_CELL_BOOLEAN


# ---------------------------------------------------------------------------
# Sheet.cell_xf_index
# ---------------------------------------------------------------------------

class TestCellXfIndex:
    def test_cell_without_fmt_info_raises(self, sheet):
        sheet.put_cell(0, 0, XL_CELL_TEXT, 'hi', -1)
        with pytest.raises(XLRDError):
            sheet.cell_xf_index(0, 0)

    def test_direct_xf_index(self, fmt_sheet):
        fmt_sheet.put_cell(0, 0, XL_CELL_TEXT, 'hi', 5)
        assert fmt_sheet.cell_xf_index(0, 0) == 5

    def test_rowinfo_fallback(self, fmt_sheet):
        fmt_sheet.put_cell(0, 0, XL_CELL_TEXT, 'hi', -1)
        ri = Rowinfo()
        ri.xf_index = 3
        fmt_sheet.rowinfo_map[0] = ri
        assert fmt_sheet.cell_xf_index(0, 0) == 3

    def test_colinfo_fallback(self, fmt_sheet):
        fmt_sheet.put_cell(0, 0, XL_CELL_TEXT, 'hi', -1)
        ci = Colinfo()
        ci.xf_index = 7
        fmt_sheet.colinfo_map[0] = ci
        assert fmt_sheet.cell_xf_index(0, 0) == 7

    def test_colinfo_xf_minus1_returns_15(self, fmt_sheet):
        fmt_sheet.put_cell(0, 0, XL_CELL_TEXT, 'hi', -1)
        ci = Colinfo()
        ci.xf_index = -1
        fmt_sheet.colinfo_map[0] = ci
        assert fmt_sheet.cell_xf_index(0, 0) == 15

    def test_default_fallback_15(self, fmt_sheet):
        fmt_sheet.put_cell(0, 0, XL_CELL_TEXT, 'hi', -1)
        assert fmt_sheet.cell_xf_index(0, 0) == 15


# ---------------------------------------------------------------------------
# Sheet.row_len / row / __getitem__ / get_rows
# ---------------------------------------------------------------------------

class TestRowAccess:
    def test_row_len(self):
        sh = _make_sheet_with_data()
        assert sh.row_len(0) == 3

    def test_row_returns_cells(self):
        sh = _make_sheet_with_data()
        row = sh.row(0)
        assert len(row) == 3
        assert isinstance(row[0], Cell)
        assert row[0].value == 'hello'

    def test_getitem_row_index(self):
        sh = _make_sheet_with_data()
        row = sh[0]
        assert len(row) == 3

    def test_getitem_row_col_tuple(self):
        sh = _make_sheet_with_data()
        c = sh[0, 1]
        assert c.value == 42.0

    def test_get_rows_generator(self):
        sh = _make_sheet_with_data()
        rows = list(sh.get_rows())
        assert len(rows) == 3
        assert rows[0][0].value == 'hello'

    def test_iter_sheet(self):
        sh = _make_sheet_with_data()
        rows = list(sh)
        assert len(rows) == 3


# ---------------------------------------------------------------------------
# Sheet.row_types / row_values / row_slice
# ---------------------------------------------------------------------------

class TestRowSlicing:
    def test_row_types_full(self):
        sh = _make_sheet_with_data()
        types = sh.row_types(0)
        assert len(types) == 3
        assert types[0] == XL_CELL_TEXT
        assert types[1] == XL_CELL_NUMBER

    def test_row_types_with_end(self):
        sh = _make_sheet_with_data()
        types = sh.row_types(0, 0, 2)
        assert len(types) == 2

    def test_row_types_with_start(self):
        sh = _make_sheet_with_data()
        types = sh.row_types(0, 1)
        assert len(types) == 2

    def test_row_values_full(self):
        sh = _make_sheet_with_data()
        vals = sh.row_values(0)
        assert vals[0] == 'hello'
        assert vals[1] == 42.0

    def test_row_values_slice(self):
        sh = _make_sheet_with_data()
        vals = sh.row_values(0, 1, 3)
        assert len(vals) == 2
        assert vals[0] == 42.0

    def test_row_slice_full(self):
        sh = _make_sheet_with_data()
        slc = sh.row_slice(0)
        assert len(slc) == 3
        assert all(isinstance(c, Cell) for c in slc)

    def test_row_slice_negative_start(self):
        sh = _make_sheet_with_data()
        slc = sh.row_slice(0, -2)
        assert len(slc) == 2

    def test_row_slice_negative_start_clamp(self):
        sh = _make_sheet_with_data()
        slc = sh.row_slice(0, -100)
        assert len(slc) == 3

    def test_row_slice_negative_end(self):
        sh = _make_sheet_with_data()
        slc = sh.row_slice(0, 0, -1)
        assert len(slc) == 2

    def test_row_slice_end_none(self):
        sh = _make_sheet_with_data()
        slc = sh.row_slice(0, 0, None)
        assert len(slc) == 3

    def test_row_slice_end_exceeds_nc(self):
        sh = _make_sheet_with_data()
        slc = sh.row_slice(0, 0, 100)
        assert len(slc) == 3


# ---------------------------------------------------------------------------
# Sheet.col_slice / col_values / col_types
# ---------------------------------------------------------------------------

class TestColSlicing:
    def test_col_slice_full(self):
        sh = _make_sheet_with_data()
        slc = sh.col_slice(0)
        assert len(slc) == 3
        assert slc[0].value == 'hello'

    def test_col_slice_partial(self):
        sh = _make_sheet_with_data()
        slc = sh.col_slice(0, 1, 3)
        assert len(slc) == 2

    def test_col_slice_negative_start(self):
        sh = _make_sheet_with_data()
        slc = sh.col_slice(0, -2)
        assert len(slc) == 2

    def test_col_slice_negative_start_clamp(self):
        sh = _make_sheet_with_data()
        slc = sh.col_slice(0, -100)
        assert len(slc) == 3

    def test_col_slice_negative_end(self):
        sh = _make_sheet_with_data()
        slc = sh.col_slice(0, 0, -1)
        assert len(slc) == 2

    def test_col_slice_end_exceeds_nrows(self):
        sh = _make_sheet_with_data()
        slc = sh.col_slice(0, 0, 100)
        assert len(slc) == 3

    def test_col_values_full(self):
        sh = _make_sheet_with_data()
        vals = sh.col_values(0)
        assert vals == ['hello', 3.14, 42000.0]

    def test_col_values_partial(self):
        sh = _make_sheet_with_data()
        vals = sh.col_values(0, 0, 2)
        assert len(vals) == 2

    def test_col_values_negative_start(self):
        sh = _make_sheet_with_data()
        vals = sh.col_values(0, -2)
        assert len(vals) == 2

    def test_col_values_negative_end(self):
        sh = _make_sheet_with_data()
        vals = sh.col_values(0, 0, -1)
        assert len(vals) == 2

    def test_col_types_full(self):
        sh = _make_sheet_with_data()
        types = sh.col_types(0)
        assert types[0] == XL_CELL_TEXT
        assert types[1] == XL_CELL_NUMBER

    def test_col_types_partial(self):
        sh = _make_sheet_with_data()
        types = sh.col_types(0, 1, 3)
        assert len(types) == 2

    def test_col_types_negative_start(self):
        sh = _make_sheet_with_data()
        types = sh.col_types(0, -1)
        assert len(types) == 1

    def test_col_types_negative_end(self):
        sh = _make_sheet_with_data()
        types = sh.col_types(0, 0, -1)
        assert len(types) == 2

    def test_col_is_col_slice(self):
        sh = _make_sheet_with_data()
        col = sh.col(0)
        col_slc = sh.col_slice(0)
        assert len(col) == len(col_slc)
        for a, b in zip(col, col_slc):
            assert a.ctype == b.ctype
            assert a.value == b.value


# ---------------------------------------------------------------------------
# Sheet.__repr__
# ---------------------------------------------------------------------------

class TestSheetRepr:
    def test_repr_contains_name_and_number(self, sheet):
        r = repr(sheet)
        assert 'TestSheet' in r
        assert '0' in r


# ---------------------------------------------------------------------------
# put_cell_ragged / put_cell_unragged
# ---------------------------------------------------------------------------

class TestPutCell:
    def test_put_cell_unragged_creates_rows(self, sheet):
        sheet.put_cell(0, 0, XL_CELL_TEXT, 'a', -1)
        sheet.put_cell(0, 1, XL_CELL_TEXT, 'b', -1)
        assert sheet.nrows == 1
        assert sheet.ncols == 2

    def test_put_cell_unragged_extend_column(self, sheet):
        sheet.put_cell(0, 0, XL_CELL_TEXT, 'a', -1)
        sheet.put_cell(0, 3, XL_CELL_TEXT, 'd', -1)
        assert sheet.ncols == 4
        assert sheet.cell_value(0, 3) == 'd'

    def test_put_cell_ragged_no_fill(self, ragged_sheet):
        ragged_sheet.put_cell(0, 0, XL_CELL_TEXT, 'a', -1)
        ragged_sheet.put_cell(1, 2, XL_CELL_NUMBER, 1.0, -1)
        assert ragged_sheet.row_len(1) == 3
        assert ragged_sheet.cell_type(1, 0) == XL_CELL_EMPTY

    def test_put_cell_unragged_with_fmt_info(self, fmt_sheet):
        fmt_sheet.put_cell(0, 0, XL_CELL_TEXT, 'a', 3)
        assert fmt_sheet._cell_xf_indexes[0][0] == 3

    def test_put_cell_ragged_with_fmt_info(self):
        book = FmtBook()
        book.ragged_rows = True
        sh = Sheet(book, 0, 'S', 0)
        sh.put_cell(0, 0, XL_CELL_TEXT, 'a', 5)
        assert sh._cell_xf_indexes[0][0] == 5


# ---------------------------------------------------------------------------
# req_fmt_info / computed_column_width
# ---------------------------------------------------------------------------

class TestReqFmtInfo:
    def test_raises_without_formatting_info(self, sheet):
        with pytest.raises(XLRDError):
            sheet.req_fmt_info()

    def test_no_raise_with_formatting_info(self, fmt_sheet):
        fmt_sheet.req_fmt_info()  # should not raise


class TestComputedColumnWidth:
    def test_raises_without_fmt_info(self, sheet):
        with pytest.raises(XLRDError):
            sheet.computed_column_width(0)

    def test_default_width_no_colinfo(self, fmt_sheet):
        assert fmt_sheet.computed_column_width(0) == 8 * 256

    def test_defcolwidth_used(self, fmt_sheet):
        fmt_sheet.defcolwidth = 10
        assert fmt_sheet.computed_column_width(0) == 10 * 256

    def test_standardwidth_used_biff8(self, fmt_sheet):
        fmt_sheet.standardwidth = 2048
        assert fmt_sheet.computed_column_width(0) == 2048

    def test_colinfo_width_biff8(self, fmt_sheet):
        ci = Colinfo()
        ci.width = 3000
        fmt_sheet.colinfo_map[0] = ci
        assert fmt_sheet.computed_column_width(0) == 3000

    def test_biff3_colinfo_width(self, fmt_book):
        fmt_book.biff_version = 30
        sh = Sheet(fmt_book, 0, 'S', 0)
        ci = Colinfo()
        ci.width = 1500
        sh.colinfo_map[0] = ci
        assert sh.computed_column_width(0) == 1500

    def test_biff3_defcolwidth_fallback(self, fmt_book):
        fmt_book.biff_version = 30
        sh = Sheet(fmt_book, 0, 'S', 0)
        sh.defcolwidth = 5
        assert sh.computed_column_width(0) == 5 * 256


# ---------------------------------------------------------------------------
# Rowinfo tests
# ---------------------------------------------------------------------------

class TestRowinfo:
    def test_init_defaults_none(self):
        ri = Rowinfo()
        assert ri.height is None
        assert ri.has_default_height is None
        assert ri.outline_level is None
        assert ri.hidden is None
        assert ri.xf_index is None

    def test_getstate_returns_tuple(self):
        ri = Rowinfo()
        ri.height = 300
        ri.has_default_height = 0
        ri.outline_level = 0
        ri.outline_group_starts_ends = 0
        ri.hidden = 0
        ri.height_mismatch = 0
        ri.has_default_xf_index = 1
        ri.xf_index = 5
        ri.additional_space_above = 0
        ri.additional_space_below = 0
        state = ri.__getstate__()
        assert isinstance(state, tuple)
        assert len(state) == 10
        assert state[0] == 300
        assert state[7] == 5

    def test_setstate_restores_values(self):
        ri = Rowinfo()
        ri.height = 400
        ri.has_default_height = 1
        ri.outline_level = 2
        ri.outline_group_starts_ends = 0
        ri.hidden = 0
        ri.height_mismatch = 0
        ri.has_default_xf_index = 0
        ri.xf_index = 0
        ri.additional_space_above = 0
        ri.additional_space_below = 0
        state = ri.__getstate__()
        ri2 = Rowinfo()
        ri2.__setstate__(state)
        assert ri2.height == 400
        assert ri2.has_default_height == 1
        assert ri2.outline_level == 2

    def test_roundtrip_getstate_setstate(self):
        ri = Rowinfo()
        ri.height = 255
        ri.has_default_height = 0
        ri.outline_level = 1
        ri.outline_group_starts_ends = 1
        ri.hidden = 0
        ri.height_mismatch = 1
        ri.has_default_xf_index = 1
        ri.xf_index = 10
        ri.additional_space_above = 1
        ri.additional_space_below = 0
        state = ri.__getstate__()
        ri_copy = Rowinfo()
        ri_copy.__setstate__(state)
        assert ri_copy.__getstate__() == state


# ---------------------------------------------------------------------------
# unpack_RK tests
# ---------------------------------------------------------------------------

class TestUnpackRK:
    def test_integer_value(self):
        # flags=2 means signed integer in bits 31..2
        i = 42
        rk = pack('<i', (i << 2) | 2)
        assert unpack_RK(rk) == 42.0

    def test_integer_divided_by_100(self):
        i = 42
        rk = pack('<i', (i << 2) | 3)
        assert abs(unpack_RK(rk) - 0.42) < 1e-10

    def test_float_value(self):
        import struct
        d = 100.0
        packed = struct.pack('<d', d)
        rk = bytes([packed[4] & 252]) + packed[5:8]
        result = unpack_RK(rk)
        assert result == 100.0

    def test_float_divided_by_100(self):
        import struct
        d = 100.0
        packed = struct.pack('<d', d)
        # Set bit 0 (divide by 100), clear bit 1 (float path)
        rk = bytes([packed[4] | 0x01]) + packed[5:8]
        result = unpack_RK(rk)
        assert abs(result - 1.0) < 1e-10

    def test_negative_integer(self):
        i = -10
        rk = pack('<i', (i << 2) | 2)
        result = unpack_RK(rk)
        assert result == -10.0

    def test_zero_integer(self):
        rk = pack('<i', 0 | 2)
        result = unpack_RK(rk)
        assert result == 0.0


# ---------------------------------------------------------------------------
# Sheet.handle_hlink / get_nul_terminated_unicode
# ---------------------------------------------------------------------------

def _make_hlink_header(options):
    """Build the 32-byte header for a hyperlink record."""
    import struct
    guid0 = b"\xD0\xC9\xEA\x79\xF9\xBA\xCE\x11\x8C\x82\x00\xAA\x00\x4B\xA9\x0B"
    dummy = b"\x02\x00\x00\x00"
    return struct.pack('<HHHH16s4si', 0, 0, 0, 0, guid0, dummy, options)


def _nul_terminated_unicode_field(text):
    """Encode text (without NUL) as a NUL-terminated UTF-16le field with 4-byte length prefix."""
    import struct
    text_with_nul = text + '\x00'
    encoded = text_with_nul.encode('UTF-16le')
    count = len(text_with_nul)  # number of characters including NUL
    return struct.pack('<L', count) + encoded


class TestHandleHlinkGetNulTerminatedUnicode:
    def test_description_parsed_from_hlink(self, sheet):
        """handle_hlink with description option invokes get_nul_terminated_unicode."""
        options = 0x04  # has description; workbook link type; no textmark
        header = _make_hlink_header(options)
        desc_field = _nul_terminated_unicode_field('MyDescription')
        data = header + desc_field

        sheet.handle_hlink(data)

        assert len(sheet.hyperlink_list) == 1
        assert sheet.hyperlink_list[0].desc == 'MyDescription'

    def test_target_parsed_from_hlink(self, sheet):
        """handle_hlink with target option invokes get_nul_terminated_unicode for target."""
        options = 0x80  # has target; no description; no moniker; workbook type
        header = _make_hlink_header(options)
        target_field = _nul_terminated_unicode_field('TargetFrame')
        data = header + target_field

        sheet.handle_hlink(data)

        assert len(sheet.hyperlink_list) == 1
        assert sheet.hyperlink_list[0].target == 'TargetFrame'

    def test_description_and_target_both_parsed(self, sheet):
        """handle_hlink with both description and target bits reads both strings."""
        options = 0x84  # has description (0x04) + has target (0x80)
        header = _make_hlink_header(options)
        desc_field = _nul_terminated_unicode_field('Desc')
        target_field = _nul_terminated_unicode_field('Frame')
        data = header + desc_field + target_field

        sheet.handle_hlink(data)

        assert sheet.hyperlink_list[0].desc == 'Desc'
        assert sheet.hyperlink_list[0].target == 'Frame'
