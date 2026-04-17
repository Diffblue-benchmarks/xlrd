# -*- coding: utf-8 -*-
"""
Tests for xlrd.sheet.Sheet — cell access methods using a minimal mock book.
"""
import sys
import io
from array import array

import pytest

from xlrd.sheet import Sheet, Cell, empty_cell
from xlrd.biffh import (
    XL_CELL_EMPTY, XL_CELL_TEXT, XL_CELL_NUMBER, XL_CELL_DATE,
    XL_CELL_BOOLEAN, XL_CELL_ERROR,
)


def _make_book(biff_version=80, formatting_info=False, ragged_rows=False):
    """Return a minimal mock book object sufficient to construct a Sheet."""
    book = type('MockBook', (), {
        'biff_version': biff_version,
        'logfile': io.StringIO(),
        'verbosity': 0,
        'formatting_info': formatting_info,
        'ragged_rows': ragged_rows,
        '_xf_index_to_xl_type_map': {},
        '_sheet_visibility': [0],
    })()
    return book


def _make_sheet(nrows=3, ncols=3, biff_version=80, formatting_info=False,
                ragged_rows=False):
    """
    Construct a Sheet pre-populated with a nrows×ncols grid of cells.
    Cell (r, c) has type XL_CELL_NUMBER and value float(r * ncols + c).
    """
    book = _make_book(biff_version=biff_version,
                      formatting_info=formatting_info,
                      ragged_rows=ragged_rows)
    sh = Sheet(book, position=0, name="Sheet1", number=0)
    sh.nrows = nrows
    sh.ncols = ncols

    for r in range(nrows):
        row_types = array('B', [XL_CELL_NUMBER] * ncols)
        row_values = [float(r * ncols + c) for c in range(ncols)]
        sh._cell_types.append(row_types)
        sh._cell_values.append(row_values)
        if formatting_info:
            sh._cell_xf_indexes.append(array('h', [-1] * ncols))

    return sh


class TestSheetBasicAttributes:

    def test_sheet_name(self):
        sh = _make_sheet()
        assert sh.name == "Sheet1"

    def test_nrows(self):
        sh = _make_sheet(nrows=4, ncols=2)
        assert sh.nrows == 4

    def test_ncols(self):
        sh = _make_sheet(nrows=4, ncols=2)
        assert sh.ncols == 2

    def test_visibility_defaults_to_zero(self):
        sh = _make_sheet()
        assert sh.visibility == 0


class TestSheetCellAccess:

    def test_cell_value_first_cell(self):
        sh = _make_sheet()
        assert sh.cell_value(0, 0) == 0.0

    def test_cell_value_middle(self):
        sh = _make_sheet(nrows=3, ncols=3)
        # Row 1, Col 1 = 1*3+1 = 4
        assert sh.cell_value(1, 1) == 4.0

    def test_cell_type_is_number(self):
        sh = _make_sheet()
        assert sh.cell_type(0, 0) == XL_CELL_NUMBER

    def test_cell_returns_cell_object(self):
        sh = _make_sheet()
        c = sh.cell(0, 0)
        assert isinstance(c, Cell)
        assert c.ctype == XL_CELL_NUMBER
        assert c.value == 0.0

    def test_cell_negative_index_wraps(self):
        sh = _make_sheet(nrows=3, ncols=3)
        # Row -1 => row 2, col -1 => col 2 = value 2*3+2 = 8
        assert sh.cell_value(-1, -1) == 8.0


class TestSheetRowAccess:

    def test_row_returns_list_of_cells(self):
        sh = _make_sheet(nrows=2, ncols=3)
        row = sh.row(0)
        assert len(row) == 3
        assert all(isinstance(c, Cell) for c in row)

    def test_row_values(self):
        sh = _make_sheet(nrows=2, ncols=3)
        vals = sh.row_values(0)
        assert vals == [0.0, 1.0, 2.0]

    def test_row_values_with_start(self):
        sh = _make_sheet(nrows=2, ncols=3)
        vals = sh.row_values(0, start_colx=1)
        assert vals == [1.0, 2.0]

    def test_row_values_with_end(self):
        sh = _make_sheet(nrows=2, ncols=3)
        vals = sh.row_values(0, start_colx=0, end_colx=2)
        assert vals == [0.0, 1.0]

    def test_row_types(self):
        sh = _make_sheet(nrows=2, ncols=3)
        types = sh.row_types(0)
        assert list(types) == [XL_CELL_NUMBER] * 3

    def test_row_len(self):
        sh = _make_sheet(nrows=2, ncols=3)
        assert sh.row_len(0) == 3


class TestSheetColumnAccess:

    def test_col_values(self):
        sh = _make_sheet(nrows=3, ncols=3)
        # Column 0: values are 0.0, 3.0, 6.0
        vals = sh.col_values(0)
        assert vals == [0.0, 3.0, 6.0]

    def test_col_values_with_range(self):
        sh = _make_sheet(nrows=3, ncols=3)
        vals = sh.col_values(0, start_rowx=1, end_rowx=3)
        assert vals == [3.0, 6.0]

    def test_col_types(self):
        sh = _make_sheet(nrows=3, ncols=3)
        types = sh.col_types(0)
        assert types == [XL_CELL_NUMBER] * 3

    def test_col_slice_returns_cells(self):
        sh = _make_sheet(nrows=3, ncols=3)
        cells = sh.col_slice(0)
        assert len(cells) == 3
        assert all(isinstance(c, Cell) for c in cells)


class TestSheetSlicing:

    def test_row_slice(self):
        sh = _make_sheet(nrows=2, ncols=3)
        cells = sh.row_slice(0, start_colx=1, end_colx=3)
        assert len(cells) == 2
        assert cells[0].value == 1.0
        assert cells[1].value == 2.0

    def test_row_slice_negative_start(self):
        sh = _make_sheet(nrows=2, ncols=3)
        cells = sh.row_slice(0, start_colx=-2)
        assert len(cells) == 2

    def test_col_slice_with_negative_end(self):
        sh = _make_sheet(nrows=3, ncols=3)
        cells = sh.col_slice(0, start_rowx=0, end_rowx=-1)
        assert len(cells) == 2


class TestSheetGetitem:

    def test_getitem_row_index(self):
        sh = _make_sheet(nrows=2, ncols=3)
        row = sh[0]
        assert len(row) == 3
        assert all(isinstance(c, Cell) for c in row)

    def test_getitem_row_col_tuple(self):
        sh = _make_sheet(nrows=3, ncols=3)
        cell = sh[1, 2]
        assert cell.value == 5.0  # 1*3+2

    def test_iter_yields_rows(self):
        sh = _make_sheet(nrows=3, ncols=3)
        rows = list(sh)
        assert len(rows) == 3
        assert all(len(r) == 3 for r in rows)


class TestSheetGetRows:

    def test_get_rows_is_generator(self):
        sh = _make_sheet(nrows=2, ncols=2)
        rows = sh.get_rows()
        import types
        assert isinstance(rows, types.GeneratorType)

    def test_get_rows_count(self):
        sh = _make_sheet(nrows=5, ncols=2)
        assert sum(1 for _ in sh.get_rows()) == 5


class TestEmptyCell:

    def test_empty_cell_type_is_empty(self):
        assert empty_cell.ctype == XL_CELL_EMPTY

    def test_empty_cell_value_is_empty_string(self):
        assert empty_cell.value == ''
