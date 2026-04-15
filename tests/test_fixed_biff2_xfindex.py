# -*- coding: utf-8 -*-
"""Unit tests for Sheet.fixed_BIFF2_xfindex."""

import sys
import pytest
from struct import pack

from xlrd.sheet import Sheet
from xlrd.biffh import XLRDError


class MockBook:
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    _sheet_visibility = [0, 0, 0, 0, 0]

    def __init__(self, biff_version=20):
        self.biff_version = biff_version
        self._xf_index_to_xl_type_map = {}
        self.xf_list = []
        self.format_map = {}
        self.format_list = []


@pytest.fixture
def book_biff21():
    return MockBook(biff_version=21)


@pytest.fixture
def sheet_biff21(book_biff21):
    return Sheet(book_biff21, 0, 'TestSheet', 0)


@pytest.fixture
def book_biff20():
    return MockBook(biff_version=20)


@pytest.fixture
def sheet_biff20(book_biff20):
    return Sheet(book_biff20, 0, 'TestSheet', 0)


class TestFixedBIFF2XFIndex:

    def test_biff21_with_xf_list_true_xfx_not_none(self, sheet_biff21, book_biff21):
        book_biff21.xf_list.append(object())
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        result = sheet_biff21.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0, true_xfx=5)
        assert result == 5

    def test_biff21_with_xf_list_true_xfx_none_uses_cell_attr(self, sheet_biff21, book_biff21):
        book_biff21.xf_list.append(object())
        # cell_attr[0] & 0x3F = 3
        cell_attr = pack('<BBB', 0x03, 0x00, 0x00)
        result = sheet_biff21.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
        assert result == 3

    def test_biff21_xf_index_63_no_ixfe_raises(self, sheet_biff21, book_biff21):
        book_biff21.xf_list.append(object())
        # cell_attr[0] & 0x3F = 0x3F = 63
        cell_attr = pack('<BBB', 0x3F, 0x00, 0x00)
        with pytest.raises(XLRDError):
            sheet_biff21.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)

    def test_biff21_xf_index_63_with_ixfe_returns_ixfe(self, sheet_biff21, book_biff21):
        book_biff21.xf_list.append(object())
        sheet_biff21._ixfe = 7
        cell_attr = pack('<BBB', 0x3F, 0x00, 0x00)
        result = sheet_biff21.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
        assert result == 7

    def test_biff21_empty_xf_list_falls_through_to_biff20(self, sheet_biff21, book_biff21):
        # biff_version == 21 but no xf_list -> downgrade to 20, insert xfs
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        result = sheet_biff21.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
        assert book_biff21.biff_version == 20
        assert sheet_biff21.biff_version == 20
        assert result is not None

    def test_biff20_cache_hit_returns_cached_xfx(self, sheet_biff20):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        sheet_biff20._cell_attr_to_xfx[cell_attr] = 42
        result = sheet_biff20.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
        assert result == 42

    def test_biff20_no_xf_list_inserts_16_styles_then_new_xf(self, sheet_biff20, book_biff20):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        result = sheet_biff20.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
        # 16 default styles + 1 for the cell_attr
        assert len(book_biff20.xf_list) == 17
        assert result == 16

    def test_biff20_second_call_uses_cache(self, sheet_biff20, book_biff20):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        result1 = sheet_biff20.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
        xf_count_after_first = len(book_biff20.xf_list)
        result2 = sheet_biff20.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
        assert result1 == result2
        assert len(book_biff20.xf_list) == xf_count_after_first

    def test_biff20_with_existing_xf_list_no_bulk_insert(self, sheet_biff20, book_biff20):
        # Pre-populate xf_list so bulk insert is skipped
        sheet_biff20.insert_new_BIFF20_xf(cell_attr=b"\x40\x00\x00", style=1)
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        result = sheet_biff20.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
        assert result == 1

    def test_biff20_high_verbosity_logs(self, book_biff20):
        book_biff20.verbosity = 2
        sheet = Sheet(book_biff20, 0, 'VerbSheet', 0)
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        result = sheet.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
        assert result is not None

    def test_biff20_nonzero_xfx_slot_raises_assertion(self, sheet_biff20):
        # cell_attr[0] & 0x3F != 0 should trigger assert
        cell_attr = pack('<BBB', 0x01, 0x00, 0x00)
        with pytest.raises(AssertionError):
            sheet_biff20.fixed_BIFF2_xfindex(cell_attr, rowx=0, colx=0)
