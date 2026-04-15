# -*- coding: utf-8 -*-
"""Unit tests for xlrd.formatting module."""

import sys
import io
import struct
import pytest

from xlrd.formatting import (
    initialise_colour_map,
    nearest_colour_index,
    EqNeAttrs,
    Format,
    is_date_format_string,
    fill_in_standard_formats,
    check_colour_indexes_in_obj,
    handle_efont,
    handle_font,
    handle_format,
    handle_palette,
    palette_epilogue,
    handle_style,
    handle_xf,
    xf_epilogue,
    initialise_book,
    FUN, FDT, FNU, FGE, FTX,
    excel_default_palette_b8,
)
from xlrd.biffh import BaseObject, XLRDError


class MockBook(object):
    """Minimal book mock for testing formatting functions."""

    def __init__(self, biff_version=80, formatting_info=True, verbosity=0):
        self.biff_version = biff_version
        self.formatting_info = formatting_info
        self.verbosity = verbosity
        self.logfile = io.StringIO()
        self.colour_map = {}
        self.colour_indexes_used = {}
        self.palette_record = []
        self.font_list = []
        self.xf_list = []
        self.format_map = {}
        self.format_list = []
        self.style_name_map = {}
        self.xfcount = 0
        self.actualfmtcount = 0
        self.encoding = 'ascii'
        self._xf_index_to_xl_type_map = {}
        self._xf_epilogue_done = 0

    def derive_encoding(self):
        if not self.encoding:
            self.encoding = 'ascii'

    def is_date_format_string(self, fmt):
        return is_date_format_string(self, fmt)


# ===== initialise_colour_map =====

def test_initialise_colour_map_no_formatting_info():
    book = MockBook(formatting_info=False)
    initialise_colour_map(book)
    assert book.colour_map == {}
    assert book.colour_indexes_used == {}


def test_initialise_colour_map_with_formatting_info_biff8():
    book = MockBook(biff_version=80, formatting_info=True)
    initialise_colour_map(book)
    # 8 invariant colours
    for i in range(8):
        assert book.colour_map[i] == excel_default_palette_b8[i]
    # palette colours start at index 8
    ndpal = len(excel_default_palette_b8)
    for i in range(ndpal):
        assert book.colour_map[i + 8] == excel_default_palette_b8[i]
    # specials
    assert book.colour_map[ndpal + 8] is None
    assert book.colour_map[ndpal + 8 + 1] is None
    assert book.colour_map[0x51] is None
    assert book.colour_map[0x7FFF] is None


def test_initialise_colour_map_sets_colour_indexes_used_empty():
    book = MockBook(formatting_info=True)
    initialise_colour_map(book)
    assert book.colour_indexes_used == {}


# ===== nearest_colour_index =====

def test_nearest_colour_index_exact_match():
    colour_map = {0: (255, 0, 0), 1: (0, 255, 0), 2: (0, 0, 255)}
    result = nearest_colour_index(colour_map, (0, 255, 0))
    assert result == 1


def test_nearest_colour_index_closest_colour():
    colour_map = {0: (0, 0, 0), 1: (255, 255, 255)}
    result = nearest_colour_index(colour_map, (200, 200, 200))
    assert result == 1


def test_nearest_colour_index_skips_none():
    colour_map = {0: None, 1: (100, 100, 100)}
    result = nearest_colour_index(colour_map, (100, 100, 100))
    assert result == 1


def test_nearest_colour_index_returns_int():
    colour_map = {0: (0, 0, 0)}
    result = nearest_colour_index(colour_map, (0, 0, 0))
    assert isinstance(result, int)


def test_nearest_colour_index_all_none():
    colour_map = {0: None, 1: None}
    result = nearest_colour_index(colour_map, (128, 128, 128))
    assert result == 0  # default best_colourx


# ===== EqNeAttrs =====

class SampleEqNe(EqNeAttrs):
    def __init__(self, x, y):
        self.x = x
        self.y = y


def test_eqneattrs_eq_equal():
    a = SampleEqNe(1, 2)
    b = SampleEqNe(1, 2)
    assert a == b


def test_eqneattrs_eq_not_equal():
    a = SampleEqNe(1, 2)
    b = SampleEqNe(1, 3)
    assert not (a == b)


def test_eqneattrs_ne_different():
    a = SampleEqNe(1, 2)
    b = SampleEqNe(1, 3)
    assert a != b


def test_eqneattrs_ne_same():
    a = SampleEqNe(1, 2)
    b = SampleEqNe(1, 2)
    assert not (a != b)


# ===== Format.__init__ =====

def test_format_init():
    fmt = Format(42, FDT, 'dd/mm/yyyy')
    assert fmt.format_key == 42
    assert fmt.type == FDT
    assert fmt.format_str == 'dd/mm/yyyy'


def test_format_init_fun():
    fmt = Format(0, FUN, None)
    assert fmt.format_key == 0
    assert fmt.type == FUN
    assert fmt.format_str is None


# ===== is_date_format_string =====

def test_is_date_format_string_date_format():
    book = MockBook()
    assert is_date_format_string(book, 'dd/mm/yyyy') is True


def test_is_date_format_string_numeric_format():
    book = MockBook()
    assert is_date_format_string(book, '0.00') is False


def test_is_date_format_string_general():
    book = MockBook()
    assert is_date_format_string(book, 'General') is False


def test_is_date_format_string_text():
    book = MockBook()
    assert is_date_format_string(book, '@') is False


def test_is_date_format_string_with_text_escape():
    book = MockBook()
    # Text in quotes should be stripped
    assert is_date_format_string(book, '"Year: "yyyy') is True


def test_is_date_format_string_time_format():
    book = MockBook()
    assert is_date_format_string(book, 'hh:mm:ss') is True


def test_is_date_format_string_backslash_escape():
    book = MockBook()
    # escaped chars like \h should be skipped
    result = is_date_format_string(book, r'hh\hmm\mss\s')
    assert result is True


def test_is_date_format_string_numeric_exp_format():
    book = MockBook()
    assert is_date_format_string(book, '0.00E+00') is False


def test_is_date_format_string_bracketed():
    book = MockBook()
    # [h] style format
    result = is_date_format_string(book, '[h]:mm')
    assert result is True


def test_is_date_format_string_ambiguous_returns_bool():
    book = MockBook(verbosity=1)
    result = is_date_format_string(book, 'dd/mm 0.00')
    assert isinstance(result, bool)


# ===== handle_efont =====

def test_handle_efont_no_formatting_info():
    book = MockBook(biff_version=20, formatting_info=False)
    from xlrd.formatting import Font
    f = Font()
    f.colour_index = 0
    f.font_index = 0
    book.font_list = [f]
    data = struct.pack('<H', 0x0010)
    handle_efont(book, data)
    assert f.colour_index == 0  # unchanged because formatting_info is False


def test_handle_efont_sets_colour_index():
    book = MockBook(biff_version=20, formatting_info=True)
    from xlrd.formatting import Font
    f = Font()
    f.colour_index = 0
    f.font_index = 0
    book.font_list = [f]
    data = struct.pack('<H', 0x0022)
    handle_efont(book, data)
    assert f.colour_index == 0x0022


# ===== handle_font =====

def _make_font_data_biff8(height=200, options=0, colour_index=8, weight=400,
                           escapement=0, underline=0, family=0, charset=0,
                           name='Arial'):
    """Build a BIFF8 FONT record data bytes."""
    name_bytes = name.encode('utf-8')
    data = struct.pack('<HHHHHBBB',
                       height, options, colour_index, weight,
                       escapement, underline, family, charset)
    # add padding to position 14 (2 bytes for lenlen=1 unicode string)
    data += b'\x00'  # pad byte
    data += struct.pack('B', len(name_bytes))  # length
    data += b'\x00'  # grbit (not compressed)
    data += name_bytes
    return data


def test_handle_font_no_formatting_info():
    book = MockBook(biff_version=80, formatting_info=False)
    data = _make_font_data_biff8()
    handle_font(book, data)
    assert len(book.font_list) == 0


def test_handle_font_biff8_creates_font():
    book = MockBook(biff_version=80, formatting_info=True)
    data = _make_font_data_biff8(height=200, name='Arial')
    handle_font(book, data)
    assert len(book.font_list) == 1
    assert book.font_list[0].height == 200


def test_handle_font_inserts_dummy_at_index_4():
    book = MockBook(biff_version=80, formatting_info=True)
    # Add 4 fonts first so next one triggers dummy insertion
    from xlrd.formatting import Font
    for i in range(4):
        f = Font()
        f.font_index = i
        book.font_list.append(f)
    data = _make_font_data_biff8(name='Times')
    handle_font(book, data)
    assert len(book.font_list) == 6
    assert book.font_list[4].name == 'Dummy Font'


def test_handle_font_biff5():
    book = MockBook(biff_version=50, formatting_info=True)
    # BIFF5 has same structure as BIFF8 but uses unpack_string for name
    data = _make_font_data_biff8(height=180, name='Helvetica')
    from xlrd.formatting import Font
    # BIFF 5 uses unpack_string: build appropriate data
    name = 'Helvetica'
    name_bytes = name.encode('ascii')
    raw = struct.pack('<HHHHHBBB', 180, 0, 8, 400, 0, 0, 0, 0)
    raw += b'\x00'  # padding
    raw += struct.pack('B', len(name_bytes))
    raw += name_bytes
    handle_font(book, raw)
    assert len(book.font_list) == 1
    assert book.font_list[0].height == 180


# ===== handle_format =====

def test_handle_format_biff8():
    book = MockBook(biff_version=80, formatting_info=True)
    from xlrd.biffh import XL_FORMAT
    # Build a BIFF8 FORMAT record: 2 bytes key + unicode string
    fmtkey = 200
    fmt_str = 'dd/mm/yyyy'
    key_bytes = struct.pack('<H', fmtkey)
    # unicode string: 2 bytes len + grbit + chars
    str_bytes = struct.pack('<H', len(fmt_str)) + b'\x00' + fmt_str.encode('ascii')
    data = key_bytes + str_bytes
    handle_format(book, data, XL_FORMAT)
    assert fmtkey in book.format_map
    assert book.format_map[fmtkey].format_str == fmt_str


def test_handle_format_increments_count():
    book = MockBook(biff_version=80, formatting_info=True)
    from xlrd.biffh import XL_FORMAT
    fmtkey = 201
    fmt_str = '0.00'
    key_bytes = struct.pack('<H', fmtkey)
    str_bytes = struct.pack('<H', len(fmt_str)) + b'\x00' + fmt_str.encode('ascii')
    data = key_bytes + str_bytes
    handle_format(book, data, XL_FORMAT)
    assert book.actualfmtcount == 1


# ===== handle_palette =====

def _make_palette_data(colours):
    """Build a PALETTE record for given list of (r, g, b) colours."""
    n = len(colours)
    data = struct.pack('<H', n)
    for r, g, b in colours:
        # colour stored as 0x00bbggrr (little-endian 4 bytes, signed int)
        val = r | (g << 8) | (b << 16)
        data += struct.pack('<i', val)
    return data


def test_handle_palette_no_formatting_info():
    book = MockBook(biff_version=80, formatting_info=False)
    initialise_colour_map(book)
    colours = [(255, 0, 0)] * 56
    data = _make_palette_data(colours)
    handle_palette(book, data)
    assert book.palette_record == []


def test_handle_palette_biff8_updates_colour_map():
    book = MockBook(biff_version=80, formatting_info=True)
    initialise_colour_map(book)
    n_colours = 56
    colours = [(i * 4, i * 4, i * 4) for i in range(n_colours)]
    data = _make_palette_data(colours)
    handle_palette(book, data)
    assert len(book.palette_record) == n_colours
    assert book.colour_map[8] == colours[0]


def test_handle_palette_wrong_size_raises():
    book = MockBook(biff_version=80, formatting_info=True)
    initialise_colour_map(book)
    # Too short data
    data = struct.pack('<H', 56) + b'\x00' * 4  # only 1 colour, not 56
    with pytest.raises(XLRDError):
        handle_palette(book, data)


# ===== palette_epilogue =====

def test_palette_epilogue_marks_used_colour_indexes():
    book = MockBook(formatting_info=True)
    initialise_colour_map(book)
    from xlrd.formatting import Font
    f = Font()
    f.font_index = 0
    f.colour_index = 8  # known colour
    f.name = 'Arial'
    book.font_list = [f]
    palette_epilogue(book)
    assert 8 in book.colour_indexes_used


def test_palette_epilogue_skips_dummy_font():
    book = MockBook(formatting_info=True)
    initialise_colour_map(book)
    from xlrd.formatting import Font
    f = Font()
    f.font_index = 4  # dummy font
    f.colour_index = 8
    book.font_list = [f]
    palette_epilogue(book)
    assert 8 not in book.colour_indexes_used


def test_palette_epilogue_skips_system_colour():
    book = MockBook(formatting_info=True)
    initialise_colour_map(book)
    from xlrd.formatting import Font
    f = Font()
    f.font_index = 0
    f.colour_index = 0x7fff  # system window text colour
    f.name = 'Arial'
    book.font_list = [f]
    palette_epilogue(book)
    assert 0x7fff not in book.colour_indexes_used


# ===== handle_style =====

def test_handle_style_no_formatting_info():
    book = MockBook(biff_version=80, formatting_info=False)
    data = struct.pack('<HBB', 0x8000, 0, 255)
    handle_style(book, data)
    assert book.style_name_map == {}


def test_handle_style_builtin_normal():
    book = MockBook(biff_version=80, formatting_info=True)
    # Built-in bit set (0x8000), built_in_id=0 (Normal), level=255
    data = struct.pack('<HBB', 0x8000, 0, 255)
    handle_style(book, data)
    assert 'Normal' in book.style_name_map
    assert book.style_name_map['Normal'][0] == 1  # built_in


def test_handle_style_builtin_rowlevel():
    book = MockBook(biff_version=80, formatting_info=True)
    # built_in_id=1 (RowLevel_), level=0 => RowLevel_1
    data = struct.pack('<HBB', 0x8001, 1, 0)
    handle_style(book, data)
    assert 'RowLevel_1' in book.style_name_map


def test_handle_style_erroneous_record():
    book = MockBook(biff_version=80, formatting_info=True)
    data = b'\x00\x00\x00\x00'
    handle_style(book, data)
    assert 'Normal' in book.style_name_map


def test_handle_style_user_defined_biff8():
    book = MockBook(biff_version=80, formatting_info=True)
    name = 'MyStyle'
    # flag_and_xfx does not have 0x8000 bit -> user-defined
    # bv >= 80, name stored as unicode at offset 2 with lenlen=2
    name_bytes = name.encode('utf-16-le')
    str_data = struct.pack('<H', len(name)) + b'\x01' + name_bytes
    flag_data = struct.pack('<HBB', 0x0000, 0, 0)
    data = flag_data[:2] + str_data
    handle_style(book, data)
    assert 'MyStyle' in book.style_name_map


# ===== check_colour_indexes_in_obj =====

def test_check_colour_indexes_in_obj_known_index():
    book = MockBook(formatting_info=True)
    book.colour_map = {8: (255, 0, 0)}
    obj = BaseObject()
    obj.colour_index = 8
    check_colour_indexes_in_obj(book, obj, 0)
    assert 8 in book.colour_indexes_used


def test_check_colour_indexes_in_obj_unknown_index():
    book = MockBook(formatting_info=True)
    book.colour_map = {}
    obj = BaseObject()
    obj.colour_index = 99
    check_colour_indexes_in_obj(book, obj, 0)
    assert 99 not in book.colour_indexes_used


# ===== fill_in_standard_formats =====

def test_fill_in_standard_formats_populates_map():
    book = MockBook()
    book.format_map = {}
    fill_in_standard_formats(book)
    # 0x0e is a date format code
    assert 0x0e in book.format_map
    assert book.format_map[0x0e].type == FDT


def test_fill_in_standard_formats_does_not_overwrite():
    book = MockBook()
    existing = Format(0x0e, FNU, 'custom')
    book.format_map = {0x0e: existing}
    fill_in_standard_formats(book)
    assert book.format_map[0x0e] is existing


# ===== handle_xf (BIFF8) =====

def _make_xf_data_biff8(font_index=0, format_key=0, pkd_type_par=0,
                          pkd_align1=0, rotation=0, pkd_align2=0,
                          pkd_used=0xFC, pkd_brdbkg1=0, pkd_brdbkg2=0,
                          pkd_brdbkg3=0):
    return struct.pack('<HHHBBBBIiH',
                       font_index, format_key, pkd_type_par,
                       pkd_align1, rotation, pkd_align2,
                       pkd_used, pkd_brdbkg1, pkd_brdbkg2, pkd_brdbkg3)


def test_handle_xf_biff8_creates_xf():
    book = MockBook(biff_version=80, formatting_info=True)
    fill_in_standard_formats(book)
    data = _make_xf_data_biff8()
    handle_xf(book, data)
    assert len(book.xf_list) == 1
    assert book.xfcount == 1


def test_handle_xf_biff8_increments_xfcount():
    book = MockBook(biff_version=80, formatting_info=True)
    fill_in_standard_formats(book)
    data = _make_xf_data_biff8()
    handle_xf(book, data)
    handle_xf(book, data)
    assert book.xfcount == 2


def test_handle_xf_biff8_fills_standard_formats_once():
    book = MockBook(biff_version=80, formatting_info=True)
    assert not book.format_map  # empty to start
    data = _make_xf_data_biff8()
    handle_xf(book, data)
    # standard formats should now be filled in
    assert 0x0e in book.format_map


# ===== xf_epilogue =====

def _setup_book_with_xfs(biff_version=80, formatting_info=True):
    book = MockBook(biff_version=biff_version, formatting_info=formatting_info)
    fill_in_standard_formats(book)
    data = _make_xf_data_biff8(pkd_type_par=0x0004)  # is_style=1
    handle_xf(book, data)
    # add a cell xf referencing style xf[0]
    data2 = _make_xf_data_biff8(pkd_type_par=0x0000)  # is_style=0
    handle_xf(book, data2)
    return book


def test_xf_epilogue_sets_done_flag():
    book = _setup_book_with_xfs()
    xf_epilogue(book)
    assert book._xf_epilogue_done == 1


def test_xf_epilogue_updates_xl_type_map():
    book = _setup_book_with_xfs()
    xf_epilogue(book)
    for xf in book.xf_list:
        assert xf.xf_index in book._xf_index_to_xl_type_map


# ===== initialise_book =====

def test_initialise_book_sets_colour_map():
    book = MockBook(biff_version=80, formatting_info=True)
    initialise_book(book)
    assert len(book.colour_map) > 0


def test_initialise_book_attaches_methods():
    book = MockBook(biff_version=80, formatting_info=True)
    initialise_book(book)
    assert hasattr(book.__class__, 'handle_font')
    assert hasattr(book.__class__, 'handle_format')
    assert hasattr(book.__class__, 'is_date_format_string')
    assert hasattr(book.__class__, 'xf_epilogue')


def test_initialise_book_sets_xf_epilogue_done():
    book = MockBook(biff_version=80, formatting_info=True)
    initialise_book(book)
    assert book._xf_epilogue_done == 0


# ===== xf_epilogue: additional coverage =====

def test_xf_epilogue_blah_logging():
    """Line 1023: verbose logging when verbosity >= 3."""
    book = MockBook(biff_version=80, formatting_info=True, verbosity=3)
    fill_in_standard_formats(book)
    data = _make_xf_data_biff8(pkd_type_par=0x0004)
    handle_xf(book, data)
    xf_epilogue(book)
    assert "xf_epilogue called" in book.logfile.getvalue()


def test_xf_epilogue_keyerror_format_key():
    """Lines 1038-1039: format_key not in format_map yields XL_CELL_TEXT."""
    from xlrd.biffh import XL_CELL_TEXT
    book = MockBook(biff_version=80, formatting_info=True)
    fill_in_standard_formats(book)
    data = _make_xf_data_biff8()
    handle_xf(book, data)
    book.xf_list[0].format_key = 9999
    xf_epilogue(book)
    assert book._xf_index_to_xl_type_map[0] == XL_CELL_TEXT


def test_xf_epilogue_no_formatting_info_continues():
    """Line 1043: continue when not formatting_info."""
    book = MockBook(biff_version=80, formatting_info=False)
    fill_in_standard_formats(book)
    data = _make_xf_data_biff8(pkd_type_par=0x0000)
    handle_xf(book, data)
    xf_epilogue(book)
    assert book._xf_epilogue_done == 1


def test_xf_epilogue_invalid_parent_style_index_blah1():
    """Lines 1047-1048, 1052: invalid parent_style_index corrected with warning."""
    book = MockBook(biff_version=80, formatting_info=True, verbosity=1)
    fill_in_standard_formats(book)
    # is_style=0, parent_style_index=1; but only 1 XF exists → out of range
    data = _make_xf_data_biff8(pkd_type_par=0x0010)
    handle_xf(book, data)
    xf_epilogue(book)
    assert book.xf_list[0].parent_style_index == 0
    assert "parent_style_index" in book.logfile.getvalue()


def test_xf_epilogue_parent_style_self_reference():
    """Lines 1055-1056: parent_style_index equals xf_index."""
    book = MockBook(biff_version=80, formatting_info=True, verbosity=1)
    fill_in_standard_formats(book)
    # style XF at index 0
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0004))
    # cell XF at index 1, parent_style_index=1 (same as xf_index)
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0010))
    xf_epilogue(book)
    assert "parent_style_index is also" in book.logfile.getvalue()


def test_xf_epilogue_parent_not_style():
    """Lines 1059-1060: parent XF has style flag not set."""
    book = MockBook(biff_version=80, formatting_info=True, verbosity=1)
    fill_in_standard_formats(book)
    # Two cell XFs (is_style=0); xf[1] parent=0; xf[0] is not a style
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0000))
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0000))
    xf_epilogue(book)
    assert "style flag not set" in book.logfile.getvalue()


def test_xf_epilogue_parent_out_of_order():
    """Line 1064: parent_style_index > xf_index."""
    book = MockBook(biff_version=80, formatting_info=True, verbosity=1)
    fill_in_standard_formats(book)
    # style XF at index 0
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0004))
    # cell XF at index 1, parent_style_index=2 (> xf_index=1)
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0020))
    # style XF at index 2
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0004))
    xf_epilogue(book)
    assert "out of order" in book.logfile.getvalue()


def test_xf_epilogue_check_same_different_alignment():
    """Lines 1027-1028, 1069: check_same logs when alignment differs."""
    book = MockBook(biff_version=80, formatting_info=True, verbosity=1)
    fill_in_standard_formats(book)
    # style XF at index 0, hor_align=0, all flags=0
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0004, pkd_align1=0x00, pkd_used=0x00))
    # cell XF at index 1, hor_align=1 (differs from parent), all flags=0
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0000, pkd_align1=0x01, pkd_used=0x00))
    xf_epilogue(book)
    assert "alignment different" in book.logfile.getvalue()


def test_xf_epilogue_check_same_all_attrs():
    """Lines 1069, 1071, 1073, 1075: check_same called for all attribute types."""
    book = MockBook(biff_version=80, formatting_info=True, verbosity=1)
    fill_in_standard_formats(book)
    # style XF at index 0, all flags=0
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0004, pkd_used=0x00))
    # cell XF at index 1, parent=0, all flags=0
    handle_xf(book, _make_xf_data_biff8(pkd_type_par=0x0000, pkd_used=0x00))
    xf_epilogue(book)
    assert book._xf_epilogue_done == 1


def test_xf_epilogue_format_key_mismatch():
    """Lines 1077-1078: format_key differs from parent with _format_flag=0."""
    book = MockBook(biff_version=80, formatting_info=True, verbosity=1)
    fill_in_standard_formats(book)
    # style XF, format_key=0, _format_flag=0
    handle_xf(book, _make_xf_data_biff8(format_key=0, pkd_type_par=0x0004, pkd_used=0x00))
    # cell XF, format_key=0x0e (differs from parent), _format_flag=0
    handle_xf(book, _make_xf_data_biff8(format_key=0x0e, pkd_type_par=0x0000, pkd_used=0x00))
    xf_epilogue(book)
    assert "fmtk=" in book.logfile.getvalue()


def test_xf_epilogue_font_index_mismatch():
    """Lines 1084-1085: font_index differs from parent with _font_flag=0."""
    book = MockBook(biff_version=80, formatting_info=True, verbosity=1)
    fill_in_standard_formats(book)
    # style XF, font_index=0, _font_flag=0
    handle_xf(book, _make_xf_data_biff8(font_index=0, pkd_type_par=0x0004, pkd_used=0x00))
    # cell XF, font_index=1 (differs from parent), _font_flag=0
    handle_xf(book, _make_xf_data_biff8(font_index=1, pkd_type_par=0x0000, pkd_used=0x00))
    xf_epilogue(book)
    assert "fontx=" in book.logfile.getvalue()
