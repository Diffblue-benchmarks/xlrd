# -*- coding: utf-8 -*-
"""Unit tests for Sheet.fake_XF_from_BIFF20_cell_attr."""

import sys
import pytest
from struct import pack

from xlrd.sheet import Sheet


class MockBook:
    biff_version = 20
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    _xf_index_to_xl_type_map = {}
    _sheet_visibility = [0, 0, 0, 0, 0]


@pytest.fixture
def sheet():
    return Sheet(MockBook(), 0, 'TestSheet', 0)


# ---------------------------------------------------------------------------
# fake_XF_from_BIFF20_cell_attr tests
# ---------------------------------------------------------------------------

class TestFakeXFFromBIFF20CellAttr:

    def test_returns_xf_object(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        from xlrd.formatting import XF
        assert isinstance(xf, XF)

    def test_alignment_defaults(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.alignment.indent_level == 0
        assert xf.alignment.shrink_to_fit == 0
        assert xf.alignment.text_direction == 0
        assert xf.alignment.vert_align == 2
        assert xf.alignment.rotation == 0

    def test_border_diag_defaults(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.border.diag_up == 0
        assert xf.border.diag_down == 0
        assert xf.border.diag_colour_index == 0
        assert xf.border.diag_line_style == 0

    def test_format_key_extracted(self, sheet):
        # font_and_format = 0x15 => format_key = 0x15 & 0x3F = 0x15 = 21
        cell_attr = pack('<BBB', 0x00, 0x15, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.format_key == 0x15

    def test_font_index_extracted(self, sheet):
        # font_and_format = 0x80 => font_index = (0x80 & 0xC0) >> 6 = 2
        cell_attr = pack('<BBB', 0x00, 0x80, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.font_index == 2

    def test_protection_cell_locked(self, sheet):
        # prot_bits bit 6 => cell_locked
        cell_attr = pack('<BBB', 0x40, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.protection.cell_locked == 1

    def test_protection_formula_hidden(self, sheet):
        # prot_bits bit 7 => formula_hidden
        cell_attr = pack('<BBB', 0x80, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.protection.formula_hidden == 1

    def test_hor_align_extracted(self, sheet):
        # halign_etc & 0x07 => hor_align
        cell_attr = pack('<BBB', 0x00, 0x00, 0x03)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.alignment.hor_align == 3

    def test_border_sides_set_when_mask_bits_set(self, sheet):
        # halign_etc = 0x78 => left(0x08), right(0x10), top(0x20), bottom(0x40)
        cell_attr = pack('<BBB', 0x00, 0x00, 0x78)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        for side in ('left', 'right', 'top', 'bottom'):
            assert getattr(xf.border, side + '_colour_index') == 8
            assert getattr(xf.border, side + '_line_style') == 1

    def test_border_sides_clear_when_mask_bits_unset(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        for side in ('left', 'right', 'top', 'bottom'):
            assert getattr(xf.border, side + '_colour_index') == 0
            assert getattr(xf.border, side + '_line_style') == 0

    def test_background_fill_pattern_set_when_bit80(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x80)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.background.fill_pattern == 17

    def test_background_fill_pattern_zero_when_bit80_clear(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.background.fill_pattern == 0

    def test_background_colour_indices(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.background.background_colour_index == 9
        assert xf.background.pattern_colour_index == 8

    def test_parent_style_index_non_style(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr, style=0)
        assert xf.parent_style_index == 0x0FFF

    def test_parent_style_index_style(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr, style=1)
        assert xf.parent_style_index == 0

    def test_all_flags_set_to_one(self, sheet):
        cell_attr = pack('<BBB', 0x00, 0x00, 0x00)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        for attr_stem in ('format', 'font', 'alignment', 'border', 'background', 'protection'):
            assert getattr(xf, '_' + attr_stem + '_flag') == 1

    def test_combined_bits(self, sheet):
        # prot_bits=0xC0 (locked+hidden), font_and_format=0xC3 (font=3, format=3), halign_etc=0xFF
        cell_attr = pack('<BBB', 0xC0, 0xC3, 0xFF)
        xf = sheet.fake_XF_from_BIFF20_cell_attr(cell_attr)
        assert xf.protection.cell_locked == 1
        assert xf.protection.formula_hidden == 1
        assert xf.format_key == 3
        assert xf.font_index == 3
        assert xf.alignment.hor_align == 7
        assert xf.background.fill_pattern == 17
