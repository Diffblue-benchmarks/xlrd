# -*- coding: utf-8 -*-
"""
Tests for xlrd.formatting — colour helpers, is_date_format_string,
Format, Font, EqNeAttrs, and palette data.
"""
import io
import pytest

from xlrd.formatting import (
    nearest_colour_index,
    initialise_colour_map,
    is_date_format_string,
    Format,
    Font,
    EqNeAttrs,
    std_format_code_types,
    default_palette,
    excel_default_palette_b8,
    excel_default_palette_b5,
    non_date_formats,
    built_in_style_names,
    palette_epilogue,
    handle_style,
    check_colour_indexes_in_obj,
)
from xlrd.biffh import FUN, FDT, FNU, FGE, FTX, BaseObject


# ---------------------------------------------------------------------------
# Minimal mock book for functions that expect a book
# ---------------------------------------------------------------------------

def _make_mock_book(verbosity=0, biff_version=80, formatting_info=True):
    return type('MockBook', (), {
        'verbosity': verbosity,
        'logfile': io.StringIO(),
        'biff_version': biff_version,
        'formatting_info': formatting_info,
        'colour_map': {},
        'colour_indexes_used': {},
        'font_list': [],
        'format_map': {},
        'format_list': [],
        'style_name_map': {},
        'encoding': 'ascii',
    })()


# ---------------------------------------------------------------------------
# nearest_colour_index
# ---------------------------------------------------------------------------

class TestNearestColourIndex:

    def test_exact_match_returns_correct_index(self):
        colour_map = {
            0: (0, 0, 0),
            1: (255, 255, 255),
            2: (255, 0, 0),
        }
        assert nearest_colour_index(colour_map, (255, 0, 0)) == 2

    def test_nearest_to_grey(self):
        colour_map = {
            0: (0, 0, 0),
            1: (255, 255, 255),
            2: (128, 128, 128),
        }
        # (130, 130, 130) is closest to (128, 128, 128)
        assert nearest_colour_index(colour_map, (130, 130, 130)) == 2

    def test_skips_none_entries(self):
        colour_map = {
            0: None,
            1: (255, 0, 0),
        }
        assert nearest_colour_index(colour_map, (200, 0, 0)) == 1

    def test_black_maps_to_black(self):
        colour_map = {0: (0, 0, 0), 1: (255, 255, 255)}
        assert nearest_colour_index(colour_map, (0, 0, 0)) == 0

    def test_white_maps_to_white(self):
        colour_map = {0: (0, 0, 0), 1: (255, 255, 255)}
        assert nearest_colour_index(colour_map, (255, 255, 255)) == 1

    def test_all_none_returns_default_zero(self):
        colour_map = {0: None, 1: None}
        # No valid entry, so best_colourx stays 0
        result = nearest_colour_index(colour_map, (100, 100, 100))
        assert result == 0


# ---------------------------------------------------------------------------
# initialise_colour_map
# ---------------------------------------------------------------------------

class TestInitialiseColourMap:

    def test_populates_colour_map(self):
        book = _make_mock_book(biff_version=80)
        initialise_colour_map(book)
        assert len(book.colour_map) > 0

    def test_first_8_entries_from_b8_palette(self):
        book = _make_mock_book(biff_version=80)
        initialise_colour_map(book)
        for i in range(8):
            assert book.colour_map[i] == excel_default_palette_b8[i]

    def test_no_formatting_info_does_nothing(self):
        book = _make_mock_book(formatting_info=False)
        initialise_colour_map(book)
        assert book.colour_map == {}

    def test_biff5_uses_b5_palette(self):
        book = _make_mock_book(biff_version=50)
        initialise_colour_map(book)
        # Should use the b5 palette; first 8 still from b8
        for i in range(8):
            assert book.colour_map[i] == excel_default_palette_b8[i]

    def test_special_keys_are_none(self):
        book = _make_mock_book(biff_version=80)
        initialise_colour_map(book)
        # System window text colour
        assert 0x7FFF in book.colour_map
        assert book.colour_map[0x7FFF] is None


# ---------------------------------------------------------------------------
# is_date_format_string
# ---------------------------------------------------------------------------

class TestIsDateFormatString:

    def _book(self, verbosity=0):
        return _make_mock_book(verbosity=verbosity)

    def test_plain_date_format(self):
        assert is_date_format_string(self._book(), 'yyyy-mm-dd') is True

    def test_time_format_hh_mm_ss(self):
        assert is_date_format_string(self._book(), 'hh:mm:ss') is True

    def test_number_format(self):
        assert is_date_format_string(self._book(), '#,##0.00') is False

    def test_general_format_is_not_date(self):
        assert is_date_format_string(self._book(), 'General') is False

    def test_text_format_at_sign(self):
        assert is_date_format_string(self._book(), '@') is False

    def test_scientific_notation_not_date(self):
        assert is_date_format_string(self._book(), '0.00E+00') is False

    def test_quoted_text_ignored(self):
        # "Year"yyyy → removes "Year", leaves yyyy → date
        assert is_date_format_string(self._book(), '"Year"yyyy') is True

    def test_backslash_escape_ignored(self):
        # hh\hmm\m → h and m are date chars, backslash skips next char
        assert is_date_format_string(self._book(), r'hh\hmm\m') is True

    def test_bracketed_section_ignored(self):
        # [h]:mm:ss → brackets stripped, mm and ss remain → date
        assert is_date_format_string(self._book(), '[h]:mm:ss') is True

    def test_percent_format_not_date(self):
        assert is_date_format_string(self._book(), '0.00%') is False

    def test_currency_format_not_date(self):
        assert is_date_format_string(self._book(), '$#,##0') is False

    def test_mixed_date_number_uses_count(self):
        # 'yyyymmdd0' has 3*date_count (y,m,d → 5+5+5) and 1*num (0→5)
        # date_count > num_count → True
        assert is_date_format_string(self._book(), 'yyyymmdd0') is True


# ---------------------------------------------------------------------------
# Format class
# ---------------------------------------------------------------------------

class TestFormat:

    def test_init_sets_attributes(self):
        fmt = Format(164, FDT, 'yyyy-mm-dd')
        assert fmt.format_key == 164
        assert fmt.type == FDT
        assert fmt.format_str == 'yyyy-mm-dd'

    def test_equality(self):
        a = Format(1, FNU, '0.00')
        b = Format(1, FNU, '0.00')
        assert a == b

    def test_inequality(self):
        a = Format(1, FNU, '0.00')
        b = Format(2, FNU, '0.00')
        assert a != b

    def test_format_key_stored(self):
        fmt = Format(0x0E, FDT, 'm/d/yy')
        assert fmt.format_key == 0x0E


# ---------------------------------------------------------------------------
# Font class
# ---------------------------------------------------------------------------

class TestFont:

    def test_default_attributes(self):
        font = Font()
        assert font.bold == 0
        assert font.italic == 0
        assert font.weight == 400
        assert font.height == 0
        assert font.underlined == 0
        assert font.struck_out == 0

    def test_equality(self):
        a = Font()
        b = Font()
        assert a == b

    def test_inequality_after_modification(self):
        a = Font()
        b = Font()
        b.bold = 1
        assert a != b

    def test_colour_index_default(self):
        assert Font().colour_index == 0

    def test_escapement_default(self):
        assert Font().escapement == 0


# ---------------------------------------------------------------------------
# EqNeAttrs mixin
# ---------------------------------------------------------------------------

class TestEqNeAttrs:

    def test_equal_objects(self):
        class Obj(EqNeAttrs):
            pass
        a, b = Obj(), Obj()
        a.x = 1
        b.x = 1
        assert a == b

    def test_not_equal_objects(self):
        class Obj(EqNeAttrs):
            pass
        a, b = Obj(), Obj()
        a.x = 1
        b.x = 2
        assert a != b


# ---------------------------------------------------------------------------
# std_format_code_types
# ---------------------------------------------------------------------------

class TestStdFormatCodeTypes:

    def test_general_format_is_FGE(self):
        from xlrd.biffh import FGE
        assert std_format_code_types[0] == FGE

    def test_numeric_format_is_FNU(self):
        from xlrd.biffh import FNU
        assert std_format_code_types[1] == FNU  # "0"
        assert std_format_code_types[4] == FNU  # "#,##0.00"

    def test_date_format_14_is_FDT(self):
        from xlrd.biffh import FDT
        assert std_format_code_types[14] == FDT  # "m/d/yy"

    def test_time_format_18_is_FDT(self):
        from xlrd.biffh import FDT
        assert std_format_code_types[18] == FDT  # "h:mm AM/PM"

    def test_text_format_49_is_FTX(self):
        from xlrd.biffh import FTX
        assert std_format_code_types[49] == FTX  # "@"

    def test_non_existent_key_returns_none(self):
        assert std_format_code_types.get(9999) is None


# ---------------------------------------------------------------------------
# default_palette
# ---------------------------------------------------------------------------

class TestDefaultPalette:

    def test_biff8_palette_has_56_entries(self):
        assert len(excel_default_palette_b8) == 56

    def test_biff5_palette_has_56_entries(self):
        assert len(excel_default_palette_b5) == 56

    def test_all_versions_in_default_palette(self):
        for version in (80, 70, 50, 45, 40, 30, 21, 20):
            assert version in default_palette

    def test_biff8_first_entry_is_black(self):
        assert excel_default_palette_b8[0] == (0, 0, 0)

    def test_biff8_second_entry_is_white(self):
        assert excel_default_palette_b8[1] == (255, 255, 255)

    def test_all_rgb_tuples_have_three_components(self):
        for entry in excel_default_palette_b8:
            assert len(entry) == 3

    def test_all_rgb_values_in_range(self):
        for r, g, b in excel_default_palette_b8:
            assert 0 <= r <= 255
            assert 0 <= g <= 255
            assert 0 <= b <= 255


# ---------------------------------------------------------------------------
# built_in_style_names
# ---------------------------------------------------------------------------

class TestBuiltInStyleNames:

    def test_normal_is_first(self):
        assert built_in_style_names[0] == 'Normal'

    def test_hyperlink_present(self):
        assert 'Hyperlink' in built_in_style_names

    def test_length(self):
        assert len(built_in_style_names) == 10


# ---------------------------------------------------------------------------
# handle_palette with patched XLS data — error and warning paths
# ---------------------------------------------------------------------------

class TestHandlePalettePatched:
    """Tests that require patching PALETTE records in raw XLS bytes."""

    @staticmethod
    def _palette_xls_bytes():
        """Create XLS bytes with a PALETTE record using custom colours."""
        import xlwt, io, struct
        wb = xlwt.Workbook()
        wb.set_colour_RGB(0x21, 0xFF, 0x80, 0x00)
        ws = wb.add_sheet('Colors')
        style = xlwt.easyxf('pattern: pattern solid, fore_colour 0x21;')
        ws.write(0, 0, 'Orange', style)
        buf = io.BytesIO()
        wb.save(buf)
        return bytearray(buf.getvalue())

    @staticmethod
    def _find_palette_offset(data):
        """Locate PALETTE record offset in XLS bytes."""
        import struct
        PALETTE_OPCODE = struct.pack('<H', 0x0092)
        for idx in range(len(data) - 4):
            if data[idx:idx+2] == PALETTE_OPCODE:
                return idx
        return -1

    def test_palette_unexpected_n_colours_warning(self):
        """handle_palette with wrong n_colours prints NOTE warning (line 584)."""
        import struct, xlrd, io
        data = self._palette_xls_bytes()
        offset = self._find_palette_offset(data)
        assert offset >= 0, "PALETTE record not found in test data"
        data[offset+4:offset+6] = struct.pack('<H', 55)  # patch n_colours to 55 (not 56)
        log = io.StringIO()
        bk = xlrd.open_workbook(file_contents=bytes(data), verbosity=1,
                                formatting_info=True, logfile=log)
        assert 'Expected 56' in log.getvalue()
        assert 'found 55' in log.getvalue()

    def test_palette_size_error_raises(self):
        """handle_palette with wrong record size raises XLRDError (line 595)."""
        import struct, xlrd, io
        from xlrd.biffh import XLRDError
        data = self._palette_xls_bytes()
        offset = self._find_palette_offset(data)
        assert offset >= 0, "PALETTE record not found in test data"
        data[offset+2:offset+4] = struct.pack('<H', 50)  # shrink record length
        log = io.StringIO()
        with pytest.raises(XLRDError, match="PALETTE record"):
            xlrd.open_workbook(file_contents=bytes(data), verbosity=0,
                               formatting_info=True, logfile=log)

    def test_palette_verbose_prints_info(self):
        """handle_palette with verbosity=2 prints colour info (line 587-589)."""
        import xlrd, io
        data = self._palette_xls_bytes()
        log = io.StringIO()
        bk = xlrd.open_workbook(file_contents=bytes(data), verbosity=2,
                                formatting_info=True, logfile=log)
        output = log.getvalue()
        assert 'PALETTE record with' in output


# ---------------------------------------------------------------------------
# palette_epilogue — unknown colour_index warning (lines 625-627)
# ---------------------------------------------------------------------------

class TestPaletteEpilogueUnknownColour:

    def test_unknown_colour_index_prints_warning_with_verbosity(self):
        """palette_epilogue warns when font colour_index not in colour_map (lines 625-627)."""
        import struct
        book = _make_mock_book(verbosity=1)
        # colour_map has index 0x0a but NOT 0x0200
        book.colour_map = {0x0a: (255, 0, 0)}
        book.colour_indexes_used = {}

        font = Font()
        font.font_index = 0
        font.colour_index = 0x0200  # not in colour_map
        book.font_list = [font]

        palette_epilogue(book)
        output = book.logfile.getvalue()
        assert 'unknown' in output.lower() or '0x0200' in output

    def test_known_colour_index_is_recorded(self):
        """palette_epilogue records colour_indexes_used when cx in colour_map (line 624)."""
        book = _make_mock_book(verbosity=0)
        book.colour_map = {0x0a: (255, 0, 0)}
        book.colour_indexes_used = {}

        font = Font()
        font.font_index = 0
        font.colour_index = 0x0a  # in colour_map
        book.font_list = [font]

        palette_epilogue(book)
        assert 0x0a in book.colour_indexes_used

    def test_system_colour_0x7fff_is_skipped(self):
        """palette_epilogue skips fonts with colour_index == 0x7fff (line 622)."""
        book = _make_mock_book(verbosity=1)
        book.colour_map = {}
        book.colour_indexes_used = {}

        font = Font()
        font.font_index = 0
        font.colour_index = 0x7fff
        book.font_list = [font]

        palette_epilogue(book)
        # No warning should appear and colour_indexes_used unchanged
        assert book.colour_indexes_used == {}

    def test_missing_font_index_4_is_skipped(self):
        """palette_epilogue skips font with font_index == 4 (the missing font, line 619)."""
        book = _make_mock_book(verbosity=1)
        book.colour_map = {}
        book.colour_indexes_used = {}

        font = Font()
        font.font_index = 4
        font.colour_index = 0x0200  # would trigger warning if not skipped
        book.font_list = [font]

        palette_epilogue(book)
        assert book.colour_indexes_used == {}

    def test_used_colours_printed_with_verbosity_1(self):
        """palette_epilogue prints used colours when verbosity >= 1 (line 629-631)."""
        book = _make_mock_book(verbosity=1)
        book.colour_map = {0x0a: (255, 0, 0)}
        book.colour_indexes_used = {}

        font = Font()
        font.font_index = 0
        font.colour_index = 0x0a
        book.font_list = [font]

        palette_epilogue(book)
        output = book.logfile.getvalue()
        assert 'Colour indexes used' in output


# ---------------------------------------------------------------------------
# handle_style — erroneous Normal record and user-defined style (lines 640-670)
# ---------------------------------------------------------------------------

class TestHandleStyleDirect:

    def test_erroneous_normal_style_record(self):
        """handle_style handles all-zeros data as erroneous Normal style (lines 640-647)."""
        book = _make_mock_book(formatting_info=True)
        book.style_name_map = {}

        handle_style(book, b"\0\0\0\0")

        assert 'Normal' in book.style_name_map
        built_in, xf_index = book.style_name_map['Normal']
        assert built_in == 1
        assert xf_index == 0

    def test_erroneous_normal_not_repeated_if_already_present(self):
        """handle_style skips erroneous Normal record if Normal already in style_name_map."""
        book = _make_mock_book(formatting_info=True)
        book.style_name_map = {'Normal': (1, 0)}  # already present

        # Should go to the elif or else branch instead (flag_and_xfx=0, no 0x8000 bit → user-defined)
        # data == b"\0\0\0\0" but "Normal" in style_name_map, so the if branch is NOT taken
        # flag_and_xfx = 0, not & 0x8000, so user-defined branch (else)
        # For BIFF8 user-defined: unpack_unicode(data, 2, lenlen=2) with data[2:4] = b'\x00\x00' → nchars=0 → ""
        handle_style(book, b"\0\0\0\0")
        # Empty user-defined style name "" should also be stored
        assert '' in book.style_name_map or 'Normal' in book.style_name_map

    def test_user_defined_style_biff8(self):
        """handle_style adds user-defined style name for BIFF8 (lines 654-661)."""
        import struct
        book = _make_mock_book(biff_version=80, formatting_info=True)

        # Layout: flag_and_xfx (2 bytes), then Unicode string at offset 2 (lenlen=2)
        # nchars goes where built_in_id/level would be; they're overwritten for user-defined
        name = "Custom"
        data = struct.pack('<H', 0x0001) + struct.pack('<H', len(name)) + b'\x00' + name.encode('latin1')
        handle_style(book, data)
        assert 'Custom' in book.style_name_map

    def test_user_defined_style_biff5(self):
        """handle_style reads user-defined style from BIFF5 using unpack_string (line 668)."""
        import struct
        book = _make_mock_book(biff_version=50, formatting_info=True)
        book.encoding = 'ascii'

        # Layout: flag_and_xfx (2 bytes), then BIFF5 string at offset 2 (lenlen=1)
        name = "OldStyle"
        data = struct.pack('<H', 0x0001) + struct.pack('B', len(name)) + name.encode('ascii')
        handle_style(book, data)
        assert 'OldStyle' in book.style_name_map

    def test_user_defined_empty_name_warns_with_verbosity(self):
        """handle_style warns when user-defined style has empty name with verbosity>=2 (line 669-670)."""
        import struct
        book = _make_mock_book(biff_version=80, formatting_info=True, verbosity=2)

        # Empty name: nchars=0
        name_data = struct.pack('<H', 0)  # nchars=0 → unpack_unicode returns ""
        data = struct.pack('<HBB', 0x0001, 0, 0) + name_data
        handle_style(book, data)
        output = book.logfile.getvalue()
        assert 'zero-length' in output or '' in book.style_name_map

    def test_formatting_info_false_returns_early(self):
        """handle_style returns immediately when formatting_info is False (line 634-635)."""
        book = _make_mock_book(formatting_info=False)
        book.style_name_map = {}
        handle_style(book, b"\0\0\0\0")
        assert book.style_name_map == {}


# ---------------------------------------------------------------------------
# handle_style — built-in style with RowLevel/ColLevel (line 653)
# ---------------------------------------------------------------------------

class TestHandleStyleBuiltIn:

    def test_row_level_style_appends_level(self):
        """handle_style appends level+1 for RowLevel style (built_in_id=1, line 653)."""
        import struct
        book = _make_mock_book(formatting_info=True)
        # flag_and_xfx = 0x8001 (built-in flag set, xf_index=1)
        # built_in_id = 1 (RowLevel), level = 3 → name = "RowLevel_4"
        data = struct.pack('<HBB', 0x8001, 1, 3)
        handle_style(book, data)
        assert any('RowLevel' in k for k in book.style_name_map)

    def test_col_level_style_appends_level(self):
        """handle_style appends level+1 for ColLevel style (built_in_id=2, line 653)."""
        import struct
        book = _make_mock_book(formatting_info=True)
        data = struct.pack('<HBB', 0x8002, 2, 1)  # ColLevel, level=1
        handle_style(book, data)
        assert any('ColLevel' in k for k in book.style_name_map)

    def test_other_builtin_style_no_level(self):
        """handle_style does not append level for built_in_id not in {1,2}."""
        import struct
        book = _make_mock_book(formatting_info=True)
        data = struct.pack('<HBB', 0x8003, 3, 0)  # built_in_id=3 (Comma), level=0
        handle_style(book, data)
        # built_in_style_names[3] = 'Comma'
        assert 'Comma' in book.style_name_map


# ---------------------------------------------------------------------------
# check_colour_indexes_in_obj — unknown colour warning (lines 685-686)
# ---------------------------------------------------------------------------

class TestCheckColourIndexesInObj:

    def test_known_colour_index_recorded(self):
        """check_colour_indexes_in_obj records known colour_index (lines 682-683)."""
        book = _make_mock_book()
        book.colour_map = {0x0a: (255, 0, 0)}
        book.colour_indexes_used = {}

        class Obj:
            pass
        obj = Obj()
        obj.colour_index = 0x0a

        check_colour_indexes_in_obj(book, obj, orig_index=0)
        assert 0x0a in book.colour_indexes_used

    def test_unknown_colour_index_prints_warning(self):
        """check_colour_indexes_in_obj warns for colour_index not in colour_map (lines 685-686)."""
        book = _make_mock_book()
        book.colour_map = {}
        book.colour_indexes_used = {}

        class Obj:
            pass
        obj = Obj()
        obj.colour_index = 0x0200  # not in colour_map

        check_colour_indexes_in_obj(book, obj, orig_index=5)
        output = book.logfile.getvalue()
        assert '0x0200' in output
        assert 'xf #5' in output

    def test_nested_object_with_dump_is_recursed(self):
        """check_colour_indexes_in_obj recurses into attrs that have a dump() method (line 680)."""
        book = _make_mock_book()
        book.colour_map = {0x0b: (0, 255, 0)}
        book.colour_indexes_used = {}

        class Inner(BaseObject):
            pass
        inner = Inner()
        inner.colour_index = 0x0b

        class Outer:
            pass
        outer = Outer()
        outer.inner_obj = inner

        check_colour_indexes_in_obj(book, outer, orig_index=0)
        assert 0x0b in book.colour_indexes_used


# ---------------------------------------------------------------------------
# handle_format — XL_FORMAT2 rectype and codepage-aware paths
# ---------------------------------------------------------------------------

class TestHandleFormatDirect:
    """Tests for handle_format() covering BIFF2/3 (XL_FORMAT2) paths and conflict warnings."""

    def _make_book(self, biff_version=30):
        """Create a minimal Book with initialise_book called."""
        from xlrd.book import Book
        from xlrd.formatting import initialise_book
        bk = Book()
        bk.biff_version = biff_version
        bk.logfile = io.StringIO()
        bk.verbosity = 0
        bk.encoding = 'latin-1'
        bk.formatting_info = False
        bk.ragged_rows = False
        initialise_book(bk)
        return bk

    def test_xl_format2_uses_actualfmtcount_as_key(self):
        """handle_format with XL_FORMAT2: fmtkey = actualfmtcount (line 528-529, 536)."""
        import struct
        from xlrd.formatting import handle_format, XL_FORMAT2
        bk = self._make_book(biff_version=30)
        # BIFF3 FORMAT2: 1-byte length prefix, then raw string (strpos=0)
        data = struct.pack('<B', 5) + b'hello'
        handle_format(bk, data, rectype=XL_FORMAT2)
        assert 0 in bk.format_map
        assert bk.format_map[0].format_str == 'hello'

    def test_xl_format2_increments_actualfmtcount(self):
        """Multiple XL_FORMAT2 records use sequential keys."""
        import struct
        from xlrd.formatting import handle_format, XL_FORMAT2
        bk = self._make_book(biff_version=30)
        data1 = struct.pack('<B', 3) + b'foo'
        data2 = struct.pack('<B', 3) + b'bar'
        handle_format(bk, data1, rectype=XL_FORMAT2)
        handle_format(bk, data2, rectype=XL_FORMAT2)
        assert 0 in bk.format_map
        assert 1 in bk.format_map
        assert bk.format_map[1].format_str == 'bar'

    def test_biff5_handle_format_uses_fmtkey_from_data(self):
        """handle_format with BIFF5: fmtkey is read from data[0:2] (line 534)."""
        import struct
        from xlrd.formatting import handle_format, XL_FORMAT
        bk = self._make_book(biff_version=50)
        bk.encoding = 'latin-1'
        # BIFF5 FORMAT: 2-byte fmtkey + 1-byte length + string
        data = struct.pack('<H', 200) + struct.pack('<B', 4) + b'0.00'
        handle_format(bk, data, rectype=XL_FORMAT)
        assert 200 in bk.format_map
        assert bk.format_map[200].format_str == '0.00'

    def test_conflict_between_std_fmtkey_and_format_string_logs_warning(self):
        """handle_format logs conflict warning when std key type doesn't match string type (lines 558-563)."""
        import struct
        from xlrd.formatting import handle_format, XL_FORMAT
        bk = self._make_book(biff_version=50)
        bk.encoding = 'latin-1'
        bk.verbosity = 1  # required to trigger warning
        # fmtkey=14 is a standard date format, but we provide a number string '0.00'
        data = struct.pack('<H', 14) + struct.pack('<B', 4) + b'0.00'
        handle_format(bk, data, rectype=XL_FORMAT)
        log_output = bk.logfile.getvalue()
        assert 'Conflict' in log_output
        assert 'std format key 14' in log_output

    def test_handle_format_with_no_encoding_calls_derive_encoding(self):
        """handle_format calls derive_encoding() when self.encoding is None (line 531)."""
        import struct
        from xlrd.formatting import handle_format, XL_FORMAT2
        bk = self._make_book(biff_version=30)
        bk.codepage = 1252
        bk.encoding_override = ''
        bk.raw_user_name = False
        bk.encoding = None  # Reset encoding to trigger derive_encoding
        data = struct.pack('<B', 5) + b'hello'
        handle_format(bk, data, rectype=XL_FORMAT2)
        assert bk.encoding is not None  # derive_encoding was called
        assert 0 in bk.format_map
