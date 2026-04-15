# -*- coding: utf-8 -*-
"""Unit tests for Sheet.insert_new_BIFF20_xf."""

import sys
import pytest
from struct import pack

from xlrd.sheet import Sheet, XL_CELL_NUMBER, XL_CELL_DATE
from xlrd.biffh import FUN, FDT, FNU, FGE, FTX
from xlrd.formatting import Format


class MockBook:
    biff_version = 20
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    _sheet_visibility = [0, 0, 0, 0, 0]

    def __init__(self):
        self._xf_index_to_xl_type_map = {}
        self.xf_list = []
        self.format_map = {}
        self.format_list = []


@pytest.fixture
def book():
    return MockBook()


@pytest.fixture
def sheet(book):
    return Sheet(book, 0, 'TestSheet', 0)


class TestInsertNewBIFF20XF:

    def test_returns_xfx_index(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xfx = sheet.insert_new_BIFF20_xf(cell_attr)
        assert xfx == 0

    def test_returns_incremented_index_on_second_call(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xfx0 = sheet.insert_new_BIFF20_xf(cell_attr)
        cell_attr2 = pack('<BBB', 0x00, 0x01, 0x00)
        xfx1 = sheet.insert_new_BIFF20_xf(cell_attr2)
        assert xfx0 == 0
        assert xfx1 == 1

    def test_xf_appended_to_book_xf_list(self, sheet, book):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        sheet.insert_new_BIFF20_xf(cell_attr)
        assert len(book.xf_list) == 1

    def test_xf_index_set_on_xf_object(self, sheet, book):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xfx = sheet.insert_new_BIFF20_xf(cell_attr)
        assert book.xf_list[0].xf_index == xfx

    def test_format_added_to_format_map_when_missing(self, sheet, book):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        sheet.insert_new_BIFF20_xf(cell_attr)
        assert 0 in book.format_map

    def test_format_added_to_format_list_when_missing(self, sheet, book):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        sheet.insert_new_BIFF20_xf(cell_attr)
        assert len(book.format_list) == 1

    def test_format_not_duplicated_when_already_in_map(self, sheet, book):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        existing_fmt = Format(0, FUN, u"General")
        book.format_map[0] = existing_fmt
        book.format_list.append(existing_fmt)
        sheet.insert_new_BIFF20_xf(cell_attr)
        assert len(book.format_list) == 1

    def test_xf_index_to_xl_type_map_updated(self, sheet, book):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xfx = sheet.insert_new_BIFF20_xf(cell_attr)
        assert sheet._xf_index_to_xl_type_map[xfx] == XL_CELL_NUMBER

    def test_cell_attr_to_xfx_updated(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xfx = sheet.insert_new_BIFF20_xf(cell_attr)
        assert sheet._cell_attr_to_xfx[cell_attr] == xfx

    def test_cell_type_number_for_fun_format(self, sheet, book):
        # format_key=0 -> Format with type FUN -> XL_CELL_NUMBER
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xfx = sheet.insert_new_BIFF20_xf(cell_attr)
        assert sheet._xf_index_to_xl_type_map[xfx] == XL_CELL_NUMBER

    def test_cell_type_date_for_fdt_format(self, sheet, book):
        # Use a format_key=1 (FDT) pre-populated in format_map
        fmt = Format(1, FDT, u"DD/MM/YYYY")
        book.format_map[1] = fmt
        book.format_list.append(fmt)
        # font_and_format=0x01 => format_key = 0x01 & 0x3F = 1
        cell_attr = pack('<BBB', 0x00, 0x01, 0x00)
        xfx = sheet.insert_new_BIFF20_xf(cell_attr)
        assert sheet._xf_index_to_xl_type_map[xfx] == XL_CELL_DATE

    def test_style_zero_sets_parent_style_index(self, sheet, book):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        sheet.insert_new_BIFF20_xf(cell_attr, style=0)
        assert book.xf_list[0].parent_style_index == 0x0FFF

    def test_style_nonzero_sets_parent_style_index_zero(self, sheet, book):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        sheet.insert_new_BIFF20_xf(cell_attr, style=1)
        assert book.xf_list[0].parent_style_index == 0

    def test_high_verbosity_does_not_raise(self, book):
        book.verbosity = 3
        s = Sheet(book, 0, 'V3Sheet', 0)
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xfx = s.insert_new_BIFF20_xf(cell_attr)
        assert xfx == 0

    def test_format_key_nonzero_missing_logs_and_creates_fmt(self, sheet, book):
        # format_key=5, not in format_map -> creates Format(5, FUN, "General")
        # font_and_format=0x05 => format_key=5
        cell_attr = pack('<BBB', 0x00, 0x05, 0x00)
        xfx = sheet.insert_new_BIFF20_xf(cell_attr)
        assert 5 in book.format_map
        assert book.format_map[5].format_key == 5

    def test_multiple_xfs_correct_indices(self, sheet, book):
        attrs = [pack('<BBB', 0x00, i, 0x00) for i in range(3)]
        results = [sheet.insert_new_BIFF20_xf(attr) for attr in attrs]
        assert results == [0, 1, 2]
        assert len(book.xf_list) == 3
