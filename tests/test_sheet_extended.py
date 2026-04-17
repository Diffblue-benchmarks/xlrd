# -*- coding: utf-8 -*-
"""
Extended tests for xlrd.sheet.Sheet — put_cell methods, cell_xf_index,
req_fmt_info, Cell class, and tidy_dimensions.
"""
import io
import sys
from array import array

import pytest

from xlrd.sheet import Sheet, Cell, empty_cell
from xlrd.biffh import (
    XL_CELL_EMPTY, XL_CELL_TEXT, XL_CELL_NUMBER, XL_CELL_DATE,
    XL_CELL_BOOLEAN, XL_CELL_ERROR,
)


def _make_book(biff_version=80, formatting_info=False, ragged_rows=False):
    return type('MockBook', (), {
        'biff_version': biff_version,
        'logfile': io.StringIO(),
        'verbosity': 0,
        'formatting_info': formatting_info,
        'ragged_rows': ragged_rows,
        '_xf_index_to_xl_type_map': {15: XL_CELL_NUMBER},
        '_sheet_visibility': [0],
    })()


def _make_empty_sheet(biff_version=80, formatting_info=False, ragged_rows=False):
    book = _make_book(biff_version=biff_version,
                      formatting_info=formatting_info,
                      ragged_rows=ragged_rows)
    return Sheet(book, position=0, name='TestSheet', number=0)


# ---------------------------------------------------------------------------
# Cell class
# ---------------------------------------------------------------------------

class TestCell:

    def test_cell_stores_type(self):
        c = Cell(XL_CELL_NUMBER, 3.14)
        assert c.ctype == XL_CELL_NUMBER

    def test_cell_stores_value(self):
        c = Cell(XL_CELL_TEXT, 'hello')
        assert c.value == 'hello'

    def test_cell_xf_index_default_none(self):
        c = Cell(XL_CELL_NUMBER, 1.0)
        assert c.xf_index is None

    def test_cell_with_xf_index(self):
        c = Cell(XL_CELL_NUMBER, 1.0, xf_index=5)
        assert c.xf_index == 5

    def test_cell_repr_contains_type_and_value(self):
        c = Cell(XL_CELL_NUMBER, 42.0)
        r = repr(c)
        assert 'number' in r.lower() or str(XL_CELL_NUMBER) in r or '42' in r

    def test_empty_cell_type(self):
        assert empty_cell.ctype == XL_CELL_EMPTY

    def test_empty_cell_value(self):
        assert empty_cell.value == ''


# ---------------------------------------------------------------------------
# put_cell_ragged / put_cell_unragged — building a sheet cell by cell
# ---------------------------------------------------------------------------

class TestPutCellUnragged:

    def test_put_single_cell(self):
        sh = _make_empty_sheet(ragged_rows=False)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 42.0, -1)
        assert sh.nrows == 1
        assert sh.ncols == 1
        assert sh.cell_value(0, 0) == 42.0

    def test_put_cells_expands_ncols(self):
        sh = _make_empty_sheet(ragged_rows=False)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, -1)
        sh.put_cell(0, 4, XL_CELL_NUMBER, 5.0, -1)
        assert sh.ncols == 5

    def test_put_cells_expands_nrows(self):
        sh = _make_empty_sheet(ragged_rows=False)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, -1)
        sh.put_cell(2, 0, XL_CELL_NUMBER, 3.0, -1)
        assert sh.nrows == 3

    def test_put_text_cell(self):
        sh = _make_empty_sheet(ragged_rows=False)
        sh.put_cell(0, 0, XL_CELL_TEXT, 'hello', -1)
        assert sh.cell_type(0, 0) == XL_CELL_TEXT
        assert sh.cell_value(0, 0) == 'hello'

    def test_put_multiple_cells_same_row(self):
        sh = _make_empty_sheet(ragged_rows=False)
        for i in range(5):
            sh.put_cell(0, i, XL_CELL_NUMBER, float(i), -1)
        assert sh.row_values(0) == [0.0, 1.0, 2.0, 3.0, 4.0]

    def test_put_cells_different_rows(self):
        sh = _make_empty_sheet(ragged_rows=False)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 10.0, -1)
        sh.put_cell(1, 0, XL_CELL_NUMBER, 20.0, -1)
        sh.put_cell(2, 0, XL_CELL_NUMBER, 30.0, -1)
        assert sh.col_values(0) == [10.0, 20.0, 30.0]


class TestPutCellRagged:

    def test_put_single_cell(self):
        sh = _make_empty_sheet(ragged_rows=True)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 99.0, -1)
        assert sh.nrows == 1
        assert sh.cell_value(0, 0) == 99.0

    def test_ragged_rows_different_lengths(self):
        sh = _make_empty_sheet(ragged_rows=True)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, -1)
        sh.put_cell(0, 2, XL_CELL_NUMBER, 3.0, -1)  # skip col 1
        sh.put_cell(1, 0, XL_CELL_NUMBER, 4.0, -1)
        assert sh.row_len(0) == 3
        assert sh.row_len(1) == 1

    def test_ragged_ncols_tracked(self):
        sh = _make_empty_sheet(ragged_rows=True)
        sh.put_cell(0, 5, XL_CELL_NUMBER, 1.0, -1)
        assert sh.ncols == 6

    def test_text_cell_ragged(self):
        sh = _make_empty_sheet(ragged_rows=True)
        sh.put_cell(0, 0, XL_CELL_TEXT, 'world', -1)
        assert sh.cell_value(0, 0) == 'world'
        assert sh.cell_type(0, 0) == XL_CELL_TEXT


# ---------------------------------------------------------------------------
# req_fmt_info
# ---------------------------------------------------------------------------

class TestReqFmtInfo:

    def test_raises_when_formatting_info_false(self):
        sh = _make_empty_sheet(formatting_info=False)
        with pytest.raises(Exception):
            sh.req_fmt_info()

    def test_does_not_raise_when_formatting_info_true(self):
        sh = _make_empty_sheet(formatting_info=True)
        sh.req_fmt_info()  # should not raise


# ---------------------------------------------------------------------------
# cell_xf_index
# ---------------------------------------------------------------------------

class TestCellXfIndex:

    def _make_fmt_sheet(self):
        """Sheet with formatting_info=True and one cell at (0,0) with xf_index=5."""
        sh = _make_empty_sheet(formatting_info=True)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, 5)
        return sh

    def test_explicit_xf_index_returned(self):
        sh = self._make_fmt_sheet()
        assert sh.cell_xf_index(0, 0) == 5

    def test_negative_xf_falls_back_to_rowinfo(self):
        sh = _make_empty_sheet(formatting_info=True)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, -1)
        # Provide a rowinfo with xf_index=7
        rowinfo = type('Rowinfo', (), {'xf_index': 7})()
        sh.rowinfo_map[0] = rowinfo
        assert sh.cell_xf_index(0, 0) == 7

    def test_negative_xf_falls_back_to_colinfo(self):
        sh = _make_empty_sheet(formatting_info=True)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, -1)
        # Provide a colinfo with xf_index=9
        colinfo = type('Colinfo', (), {'xf_index': 9})()
        sh.colinfo_map[0] = colinfo
        assert sh.cell_xf_index(0, 0) == 9

    def test_negative_xf_falls_back_to_default_15(self):
        sh = _make_empty_sheet(formatting_info=True)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, -1)
        # No rowinfo or colinfo → hardwired default 15
        assert sh.cell_xf_index(0, 0) == 15

    def test_colinfo_minus_one_maps_to_15(self):
        sh = _make_empty_sheet(formatting_info=True)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, -1)
        colinfo = type('Colinfo', (), {'xf_index': -1})()
        sh.colinfo_map[0] = colinfo
        # colinfo xf_index == -1 → return 15
        assert sh.cell_xf_index(0, 0) == 15


# ---------------------------------------------------------------------------
# tidy_dimensions
# ---------------------------------------------------------------------------

class TestTidyDimensions:

    def test_tidy_pads_short_rows(self):
        # Build a 3-column sheet but only write 1 cell in row 0
        sh = _make_empty_sheet(ragged_rows=False)
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, -1)
        sh.put_cell(0, 1, XL_CELL_NUMBER, 2.0, -1)
        sh.put_cell(1, 0, XL_CELL_NUMBER, 3.0, -1)
        # After tidy_dimensions, row 1 should be padded to ncols=2
        sh.tidy_dimensions()
        assert len(sh._cell_values[1]) == 2

    def test_tidy_with_zero_verbosity_no_crash(self):
        sh = _make_empty_sheet()
        sh.put_cell(0, 0, XL_CELL_NUMBER, 1.0, -1)
        sh.tidy_dimensions()  # should not raise


# ---------------------------------------------------------------------------
# Sheet.__getitem__ edge cases
# ---------------------------------------------------------------------------

class TestSheetGetitemEdgeCases:

    def _make_sheet_3x3(self):
        sh = _make_empty_sheet()
        for r in range(3):
            for c in range(3):
                sh.put_cell(r, c, XL_CELL_NUMBER, float(r * 3 + c), -1)
        return sh

    def test_getitem_single_int_returns_row(self):
        sh = self._make_sheet_3x3()
        row = sh[0]
        assert len(row) == 3

    def test_getitem_tuple_returns_cell(self):
        sh = self._make_sheet_3x3()
        cell = sh[2, 2]
        assert cell.value == 8.0

    def test_getitem_negative_row(self):
        sh = self._make_sheet_3x3()
        row = sh[-1]
        assert len(row) == 3

    def test_iter_returns_all_rows(self):
        sh = self._make_sheet_3x3()
        rows = list(sh)
        assert len(rows) == 3

    def test_get_rows_generator(self):
        sh = self._make_sheet_3x3()
        rows = list(sh.get_rows())
        assert len(rows) == 3
        assert all(len(r) == 3 for r in rows)


# ---------------------------------------------------------------------------
# Sheet biff_version affects utter_max_rows
# ---------------------------------------------------------------------------

class TestSheetBiffVersionEffect:

    def test_biff8_utter_max_rows(self):
        sh = _make_empty_sheet(biff_version=80)
        assert sh.utter_max_rows == 65536

    def test_biff5_utter_max_rows(self):
        sh = _make_empty_sheet(biff_version=50)
        assert sh.utter_max_rows == 16384

    def test_utter_max_cols_always_256(self):
        for bv in (20, 30, 50, 80):
            sh = _make_empty_sheet(biff_version=bv)
            assert sh.utter_max_cols == 256


# ---------------------------------------------------------------------------
# Sheet.tidy_dimensions — MERGEDCELLS edge cases
# ---------------------------------------------------------------------------

class TestTidyDimensionsMergedCells:
    """Tests for tidy_dimensions() MERGEDCELLS processing paths."""

    def _make_sheet_with_data(self, nrows=2, ncols=2, ragged_rows=True):
        sh = _make_empty_sheet(ragged_rows=ragged_rows)
        sh.logfile = io.StringIO()
        # put some cells to build up _cell_types / _cell_values structure
        for r in range(nrows):
            sh.put_cell(r, 0, XL_CELL_TEXT, 'x', -1)
        sh.ncols = ncols
        return sh

    def test_bad_merged_cell_range_logs_warning(self):
        """Lines 608-610: MERGEDCELLS range with chi < clo logs a warning."""
        sh = self._make_sheet_with_data()
        sh.merged_cells = [(0, 3, 0, -1)]  # chi=-1 < clo=0 → invalid
        sh.tidy_dimensions()
        assert 'WARNING' in sh.logfile.getvalue()
        assert 'bad range' in sh.logfile.getvalue()

    def test_merged_cells_extend_ncols(self):
        """Lines 613-615: merged cells extend ncols and set _first_full_rowx=-2."""
        sh = _make_empty_sheet(ragged_rows=True)
        sh.logfile = io.StringIO()
        sh.ncols = 2
        sh.nrows = 0  # no rows → fix-ragged loop is no-op
        sh.merged_cells = [(0, 1, 0, 5)]  # chi=5 > ncols=2
        sh.tidy_dimensions()
        assert sh.ncols == 5
        assert sh._first_full_rowx == -2

    def test_merged_cells_extend_nrows(self):
        """Lines 616-620: merged cells extend nrows by inserting an empty cell."""
        sh = self._make_sheet_with_data(nrows=2, ncols=2)
        sh.merged_cells = [(0, 8, 0, 2)]  # rhi=8 > nrows=2
        sh.tidy_dimensions()
        assert sh.nrows == 8

    def test_dimensions_mismatch_logs_note_with_verbosity(self):
        """Lines 621-632: verbosity >= 1 and nrows/ncols != _dimnrows/_dimncols logs NOTE."""
        sh = _make_empty_sheet(ragged_rows=False)
        sh.logfile = io.StringIO()
        sh.verbosity = 1
        sh.put_cell(0, 0, XL_CELL_TEXT, 'x', -1)
        sh.put_cell(1, 0, XL_CELL_TEXT, 'y', -1)
        sh._dimnrows = 5  # differs from nrows=2 → triggers NOTE
        sh._dimncols = sh.ncols
        sh.tidy_dimensions()
        assert 'NOTE' in sh.logfile.getvalue()
        assert 'DIMENSIONS' in sh.logfile.getvalue()

    def test_first_full_rowx_minus2_pads_all_rows(self):
        """Lines 641-642: _first_full_rowx == -2 sets ubound = nrows for full ragged-fix loop."""
        sh = _make_empty_sheet(ragged_rows=False)
        sh.logfile = io.StringIO()
        # Add a cell at col 0 only → row 0 has length 1
        sh.put_cell(0, 0, XL_CELL_TEXT, 'x', -1)
        assert len(sh._cell_types[0]) == 1
        # Merged cell extends ncols to 3, setting _first_full_rowx = -2
        sh.merged_cells = [(0, 1, 0, 3)]  # chi=3 > ncols=1
        sh.tidy_dimensions()
        # _first_full_rowx should be -2 (set by merged-cell ncols extension)
        assert sh._first_full_rowx == -2
        # Row 0 should be padded to ncols=3
        assert len(sh._cell_types[0]) == 3
