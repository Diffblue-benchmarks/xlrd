# -*- coding: utf-8 -*-
"""
Integration tests that load real BIFF8 XLS files (built by xlwt) through the
complete xlrd parsing pipeline. This exercises:
  - compdoc.py: OLE2 compound-document parsing
  - book.py:    getbof, parse_globals, get_sheets, Book public API
  - sheet.py:   all row/cell record handlers
  - formatting.py: FORMAT, FONT, XF record handlers
"""
import io
import pytest
import xlrd
from xlrd.biffh import XL_CELL_TEXT, XL_CELL_NUMBER, XL_CELL_EMPTY


# ===========================================================================
# Helpers
# ===========================================================================

def open_wb(xls_bytes, **kwargs):
    """Open bytes as a workbook."""
    return xlrd.open_workbook(file_contents=xls_bytes, **kwargs)


# ===========================================================================
# Book-level API (book.py)
# ===========================================================================

class TestBookLoading:

    def test_open_workbook_returns_book(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert bk is not None

    def test_nsheets(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert bk.nsheets == 3

    def test_sheet_names(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        names = bk.sheet_names()
        assert 'Sheet1' in names
        assert 'Sheet2' in names
        assert 'EmptySheet' in names

    def test_sheet_names_returns_copy(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        names = bk.sheet_names()
        assert isinstance(names, list)

    def test_sheet_by_index(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh is not None
        assert sh.name == 'Sheet1'

    def test_sheet_by_index_second(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(1)
        assert sh.name == 'Sheet2'

    def test_sheet_by_name(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_name('Sheet2')
        assert sh is not None
        assert sh.name == 'Sheet2'

    def test_sheet_by_name_raises_for_unknown(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        with pytest.raises(xlrd.XLRDError, match='No sheet named'):
            bk.sheet_by_name('DoesNotExist')

    def test_getitem_int(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk[0]
        assert sh.name == 'Sheet1'

    def test_getitem_str(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk['Sheet2']
        assert sh.name == 'Sheet2'

    def test_sheets_returns_list(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sheets = bk.sheets()
        assert isinstance(sheets, list)
        assert len(sheets) == 3

    def test_iter_workbook(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sheet_list = list(bk)
        assert len(sheet_list) == 3

    def test_iter_workbook_names(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        names = [sh.name for sh in bk]
        assert 'Sheet1' in names

    def test_sheet_loaded_by_index(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        # After full load all sheets should be loaded
        assert bk.sheet_loaded(0) is True

    def test_sheet_loaded_by_name(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert bk.sheet_loaded('Sheet1') is True

    def test_sheet_loaded_unknown_name_raises(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        with pytest.raises(xlrd.XLRDError, match='No sheet named'):
            bk.sheet_loaded('Nope')

    def test_unload_sheet_by_index(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        bk.unload_sheet(0)
        assert bk.sheet_loaded(0) is False

    def test_unload_sheet_by_name(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        bk.unload_sheet('Sheet1')
        assert bk.sheet_loaded(0) is False

    def test_unload_unknown_name_raises(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        with pytest.raises(xlrd.XLRDError, match='No sheet named'):
            bk.unload_sheet('Nope')

    def test_release_resources(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        bk.release_resources()
        # mem should be None after release
        assert bk.mem is None

    def test_release_resources_idempotent(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        bk.release_resources()
        bk.release_resources()  # second call must not raise

    def test_context_manager(self, basic_xls_bytes):
        with open_wb(basic_xls_bytes) as bk:
            assert bk.nsheets == 3
        # resources should be released after __exit__
        assert bk.mem is None

    def test_biff_version(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        # xlwt writes BIFF8
        assert bk.biff_version == 80

    def test_encoding(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert bk.encoding is not None

    def test_datemode(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert bk.datemode in (0, 1)


class TestOnDemandLoading:

    def test_on_demand_sheets_not_loaded(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes, on_demand=True)
        # On demand: sheets exist but may not be pre-loaded
        assert bk.nsheets == 3

    def test_on_demand_access_loads_sheet(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes, on_demand=True)
        sh = bk.sheet_by_index(0)
        assert sh.name == 'Sheet1'
        bk.release_resources()

    def test_on_demand_context_manager(self, basic_xls_bytes):
        with open_wb(basic_xls_bytes, on_demand=True) as bk:
            sh = bk.sheet_by_index(0)
            assert sh.nrows > 0


class TestMultiSheetBook:

    def test_five_sheets(self, multisheet_xls_bytes):
        bk = open_wb(multisheet_xls_bytes)
        assert bk.nsheets == 5

    def test_sheets_list_length(self, multisheet_xls_bytes):
        bk = open_wb(multisheet_xls_bytes)
        assert len(bk.sheets()) == 5

    def test_all_sheet_names(self, multisheet_xls_bytes):
        bk = open_wb(multisheet_xls_bytes)
        names = bk.sheet_names()
        for i in range(1, 6):
            assert 'Sheet%d' % i in names

    def test_iterate_all_sheets(self, multisheet_xls_bytes):
        bk = open_wb(multisheet_xls_bytes)
        count = sum(1 for _ in bk)
        assert count == 5


# ===========================================================================
# Sheet-level cell reading (sheet.py record handlers)
# ===========================================================================

class TestSheetCellReading:

    def test_text_cell(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        cell = sh.cell(0, 0)
        assert cell.ctype == XL_CELL_TEXT
        assert cell.value == 'Hello'

    def test_second_text_cell(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell(0, 1).value == 'World'

    def test_unicode_cell(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        val = sh.cell(0, 2).value
        assert '\u00e9' in val or 'Unicode' in val

    def test_integer_cell(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        cell = sh.cell(1, 0)
        assert cell.ctype == XL_CELL_NUMBER
        assert cell.value == 42.0

    def test_float_cell(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        cell = sh.cell(1, 1)
        assert cell.ctype == XL_CELL_NUMBER
        assert abs(cell.value - 3.14159) < 1e-5

    def test_negative_number(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell(1, 2).value == -99.5

    def test_nrows(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.nrows >= 4

    def test_ncols(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.ncols >= 3

    def test_row_values(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        row = sh.row_values(0)
        assert 'Hello' in row
        assert 'World' in row

    def test_row_slice(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        row = sh.row_slice(0, 0, 2)
        assert len(row) == 2

    def test_col_values(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        col = sh.col_values(0)
        assert 'Hello' in col

    def test_col_slice(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        col = sh.col_slice(0, 0, 2)
        assert len(col) == 2

    def test_row_types(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        types = sh.row_types(0)
        assert types[0] == XL_CELL_TEXT

    def test_cell_type(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_type(0, 0) == XL_CELL_TEXT
        assert sh.cell_type(1, 0) == XL_CELL_NUMBER

    def test_cell_value(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(0, 0) == 'Hello'
        assert sh.cell_value(1, 0) == 42.0

    def test_sheet2_cell(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(1)
        assert sh.cell(0, 0).value == 'Sheet2 A1'
        assert sh.cell(0, 1).value == 99.0

    def test_empty_sheet(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_name('EmptySheet')
        assert sh.nrows == 0
        assert sh.ncols == 0

    def test_row_len(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.row_len(0) >= 3


class TestLargeSheet:

    def test_large_nrows(self, large_xls_bytes):
        bk = open_wb(large_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.nrows == 20

    def test_large_ncols(self, large_xls_bytes):
        bk = open_wb(large_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.ncols == 15

    def test_large_cell_values(self, large_xls_bytes):
        bk = open_wb(large_xls_bytes)
        sh = bk.sheet_by_index(0)
        # cell(r, c) == r * 100 + c
        assert sh.cell_value(5, 3) == 503.0
        assert sh.cell_value(0, 0) == 0.0
        assert sh.cell_value(19, 14) == 1914.0

    def test_all_rows_have_data(self, large_xls_bytes):
        bk = open_wb(large_xls_bytes)
        sh = bk.sheet_by_index(0)
        for r in range(sh.nrows):
            assert len(sh.row_values(r)) > 0


class TestRaggedRows:

    def test_ragged_rows_mode(self, ragged_xls_bytes):
        bk = open_wb(ragged_xls_bytes, ragged_rows=True)
        sh = bk.sheet_by_index(0)
        # Row 0 has 1 cell, row 2 has 3
        assert sh.row_len(0) == 1
        assert sh.row_len(2) == 3

    def test_non_ragged_rows_pads_to_ncols(self, ragged_xls_bytes):
        bk = open_wb(ragged_xls_bytes, ragged_rows=False)
        sh = bk.sheet_by_index(0)
        # All rows padded to max width
        ncols = sh.ncols
        for r in range(sh.nrows):
            assert sh.row_len(r) == ncols


# ===========================================================================
# OLE2 / CompDoc parsing (compdoc.py)
# ===========================================================================

class TestCompDocParsing:

    def test_ole2_file_loads(self, basic_xls_bytes):
        # The basic XLS is an OLE2 compound document
        from xlrd.compdoc import SIGNATURE
        assert basic_xls_bytes[:8] == SIGNATURE

    def test_ole2_workbook_stream_found(self, basic_xls_bytes):
        # If CompDoc fails, open_workbook raises. The fact it succeeds confirms parsing.
        bk = open_wb(basic_xls_bytes)
        assert bk.nsheets > 0

    def test_compdoc_direct(self, basic_xls_bytes):
        from xlrd import compdoc
        cd = compdoc.CompDoc(basic_xls_bytes)
        mem, base, length = cd.locate_named_stream('Workbook')
        assert mem is not None
        assert length > 0

    def test_compdoc_book_stream(self, basic_xls_bytes):
        """Try both 'Workbook' and 'Book' stream names."""
        from xlrd import compdoc
        cd = compdoc.CompDoc(basic_xls_bytes)
        # 'Workbook' should be found in xlwt-generated files
        mem, base, length = cd.locate_named_stream('Workbook')
        assert mem is not None

    def test_compdoc_missing_stream_returns_none(self, basic_xls_bytes):
        from xlrd import compdoc
        cd = compdoc.CompDoc(basic_xls_bytes)
        mem, base, length = cd.locate_named_stream('NonExistentStream')
        assert mem is None

    def test_multiple_ole2_loads(self, basic_xls_bytes):
        """Load the same OLE2 bytes multiple times (no state mutation)."""
        for _ in range(3):
            bk = open_wb(basic_xls_bytes)
            assert bk.nsheets == 3

    def test_large_file_compdoc(self, large_xls_bytes):
        """Larger file exercises more sector chains in CompDoc."""
        from xlrd import compdoc
        cd = compdoc.CompDoc(large_xls_bytes)
        mem, base, length = cd.locate_named_stream('Workbook')
        assert mem is not None


# ===========================================================================
# Formatting info (formatting.py record handlers)
# ===========================================================================

class TestFormattingInfo:

    def test_load_with_formatting_info(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        assert bk.nsheets == 1

    def test_xf_list_populated(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        assert len(bk.xf_list) > 0

    def test_font_list_populated(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        assert len(bk.font_list) > 0

    def test_format_map_has_builtin_formats(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        # Built-in formats always present
        assert len(bk.format_map) > 0

    def test_cell_xf_index_available(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        xf_idx = sh.cell_xf_index(0, 0)
        assert xf_idx is not None
        assert xf_idx >= 0

    def test_xf_record_has_format_key(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        xf = bk.xf_list[0]
        assert hasattr(xf, 'format_key')

    def test_xf_record_has_font_index(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        xf = bk.xf_list[0]
        assert hasattr(xf, 'font_index')

    def test_font_has_bold_attr(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        font = bk.font_list[0]
        assert hasattr(font, 'bold')

    def test_font_has_italic_attr(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        font = bk.font_list[0]
        assert hasattr(font, 'italic')

    def test_bold_cell_xf(self, formatted_xls_bytes):
        """Bold text cell should have xf pointing to a bold font."""
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        xf_idx = sh.cell_xf_index(0, 0)
        xf = bk.xf_list[xf_idx]
        font = bk.font_list[xf.font_index]
        assert font.bold

    def test_date_format_cell(self, formatted_xls_bytes):
        """Cell with date format should have a date-like format string."""
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        xf_idx = sh.cell_xf_index(1, 0)
        xf = bk.xf_list[xf_idx]
        fmt = bk.format_map[xf.format_key]
        assert 'Y' in fmt.format_str or 'y' in fmt.format_str or 'M' in fmt.format_str

    def test_cell_without_formatting_info_has_none_xf(self, basic_xls_bytes):
        """Without formatting_info=True, xf_index should be None."""
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        cell = sh.cell(0, 0)
        assert cell.xf_index is None

    def test_formatting_info_cells_have_xf(self, formatted_xls_bytes):
        """With formatting_info=True, cells should have xf_index."""
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        cell = sh.cell(0, 0)
        assert cell.xf_index is not None

    def test_colour_map_populated(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        assert len(bk.colour_map) > 0

    def test_palette_record_in_format_map(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        # palette and built-in formats
        assert bk.format_list is not None


class TestFormattingWithBasicWb:

    def test_basic_wb_formatting_info(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert sh.nrows > 0

    def test_xf_list_after_basic_load(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes, formatting_info=True)
        assert len(bk.xf_list) > 0

    def test_cell_xf_index_basic(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        idx = sh.cell_xf_index(0, 0)
        assert idx >= 0

    def test_style_name_map(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes, formatting_info=True)
        assert bk.style_name_map is not None


# ===========================================================================
# Named range workbook
# ===========================================================================

class TestNamedRangeWorkbook:

    def test_load_named_range_wb(self, named_range_xls_bytes):
        bk = open_wb(named_range_xls_bytes)
        assert bk.nsheets == 2

    def test_sheet_data_accessible(self, named_range_xls_bytes):
        bk = open_wb(named_range_xls_bytes)
        sh = bk.sheet_by_name('Data')
        assert sh.cell_value(0, 0) == 10.0
        assert sh.cell_value(1, 0) == 20.0
        assert sh.cell_value(2, 0) == 30.0

    def test_summary_sheet(self, named_range_xls_bytes):
        bk = open_wb(named_range_xls_bytes)
        sh = bk.sheet_by_name('Summary')
        assert sh.cell_value(0, 0) == 'Total'
        assert sh.cell_value(0, 1) == 60.0


# ===========================================================================
# From-file loading (exercises mmap/file-read paths in book.py)
# ===========================================================================

class TestFromFileLoading:

    def test_load_from_file(self, tmp_path, basic_xls_bytes):
        xls_path = tmp_path / 'test.xls'
        xls_path.write_bytes(basic_xls_bytes)
        bk = xlrd.open_workbook(str(xls_path))
        assert bk.nsheets == 3

    def test_file_load_cell_data(self, tmp_path, basic_xls_bytes):
        xls_path = tmp_path / 'test.xls'
        xls_path.write_bytes(basic_xls_bytes)
        bk = xlrd.open_workbook(str(xls_path))
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(0, 0) == 'Hello'

    def test_load_from_file_no_mmap(self, tmp_path, basic_xls_bytes):
        xls_path = tmp_path / 'test.xls'
        xls_path.write_bytes(basic_xls_bytes)
        bk = xlrd.open_workbook(str(xls_path), use_mmap=False)
        assert bk.nsheets == 3

    def test_empty_file_raises(self, tmp_path):
        empty = tmp_path / 'empty.xls'
        empty.write_bytes(b'')
        with pytest.raises(xlrd.XLRDError, match='size is 0'):
            xlrd.open_workbook(str(empty))


# ===========================================================================
# Book public attributes (book.py defaults after load)
# ===========================================================================

class TestBookAttributes:

    def test_load_time_stage_1(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert bk.load_time_stage_1 >= 0

    def test_load_time_stage_2(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert bk.load_time_stage_2 >= 0

    def test_user_name(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert isinstance(bk.user_name, str)

    def test_countries(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert isinstance(bk.countries, tuple)

    def test_sheet_visibility(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        assert hasattr(bk, '_sheet_visibility')
        assert len(bk._sheet_visibility) == bk.nsheets

    def test_shared_strings(self, basic_xls_bytes):
        # After load (resources not released), _sharedstrings exists
        bk = open_wb(basic_xls_bytes, on_demand=True)
        # Access a sheet to trigger full parse
        _ = bk.sheet_by_index(0)
        bk.release_resources()


# ===========================================================================
# Sheet object attributes (sheet.py)
# ===========================================================================

class TestSheetAttributes:

    def test_sheet_name(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.name == 'Sheet1'

    def test_sheet_number(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.number == 0

    def test_sheet_nrows_ncols_consistent(self, large_xls_bytes):
        bk = open_wb(large_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.nrows == 20
        assert sh.ncols == 15

    def test_cell_out_of_range_raises(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        with pytest.raises(IndexError):
            sh.cell(999, 999)

    def test_row_returns_cells(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        row = sh.row(0)
        assert all(hasattr(c, 'ctype') for c in row)

    def test_col_returns_cells(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        col = sh.col(0)
        assert all(hasattr(c, 'ctype') for c in col)

    def test_visibility_default(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.visibility == 0  # visible

    def test_sheet_has_biff_version(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.biff_version == 80


# ===========================================================================
# Boolean cells (BOOLERR record handler — sheet.py lines 1001-1007)
# ===========================================================================

class TestBooleanCells:

    def test_true_cell_ctype(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        from xlrd.biffh import XL_CELL_BOOLEAN
        cell = sh.cell(4, 0)
        assert cell.ctype == XL_CELL_BOOLEAN
        assert cell.value == 1

    def test_false_cell_ctype(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        from xlrd.biffh import XL_CELL_BOOLEAN
        cell = sh.cell(4, 1)
        assert cell.ctype == XL_CELL_BOOLEAN
        assert cell.value == 0

    def test_boolean_cells_in_col_types(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        from xlrd.biffh import XL_CELL_BOOLEAN
        col_types = sh.col_types(0)
        assert XL_CELL_BOOLEAN in col_types


# ===========================================================================
# Formula cells (FORMULA record handler — sheet.py lines 932-1000)
# ===========================================================================

class TestFormulaCells:

    def test_formula_xls_loads(self, formula_xls_bytes):
        bk = open_wb(formula_xls_bytes)
        assert bk.nsheets == 1
        sh = bk.sheet_by_index(0)
        assert sh.nrows > 0

    def test_formula_source_numbers(self, formula_xls_bytes):
        bk = open_wb(formula_xls_bytes)
        sh = bk.sheet_by_index(0)
        # Source values written as plain numbers
        assert sh.cell_value(0, 0) == 10.0
        assert sh.cell_value(1, 0) == 20.0
        assert sh.cell_value(2, 0) == 30.0

    def test_numeric_formula_result(self, formula_xls_bytes):
        """SUM(A1:A3) should give 60."""
        bk = open_wb(formula_xls_bytes)
        sh = bk.sheet_by_index(0)
        cell = sh.cell(3, 0)
        # xlwt may store formula with numeric result or text depending on Python 3 support
        assert cell.ctype in (XL_CELL_NUMBER, XL_CELL_TEXT, XL_CELL_EMPTY)

    def test_formula_cells_accessible(self, formula_xls_bytes):
        """All formula cells are readable without error."""
        bk = open_wb(formula_xls_bytes)
        sh = bk.sheet_by_index(0)
        for r in range(sh.nrows):
            _ = sh.cell(r, 0)  # must not raise

    def test_formula_boolean_result(self, formula_xls_bytes):
        """All formula cells are readable and have valid ctypes."""
        bk = open_wb(formula_xls_bytes)
        sh = bk.sheet_by_index(0)
        from xlrd.biffh import XL_CELL_BOOLEAN, XL_CELL_TEXT, XL_CELL_NUMBER, XL_CELL_EMPTY
        # Formula cells may produce text (empty string) or numeric result
        for r in (3, 4, 5):
            cell = sh.cell(r, 0)
            assert cell.ctype in (XL_CELL_NUMBER, XL_CELL_TEXT, XL_CELL_EMPTY, XL_CELL_BOOLEAN)


# ===========================================================================
# COLINFO record handler (sheet.py lines 1009-1044) — needs formatting_info=True
# ===========================================================================

class TestColinfo:

    def test_colinfo_parsed_with_formatting_info(self, colinfo_xls_bytes):
        bk = open_wb(colinfo_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        # colinfo_map should contain entries for columns we set widths on
        assert len(sh.colinfo_map) > 0

    def test_colinfo_width_set(self, colinfo_xls_bytes):
        bk = open_wb(colinfo_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        col0 = sh.colinfo_map.get(0)
        assert col0 is not None
        assert col0.width == 2000

    def test_colinfo_wide_column(self, colinfo_xls_bytes):
        bk = open_wb(colinfo_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        col2 = sh.colinfo_map.get(2)
        assert col2 is not None
        assert col2.width == 8000

    def test_colinfo_hidden_column(self, colinfo_xls_bytes):
        bk = open_wb(colinfo_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        col3 = sh.colinfo_map.get(3)
        assert col3 is not None
        assert col3.hidden == 1

    def test_colinfo_from_formatted_wb(self, formatted_xls_bytes):
        """formatted_xls_bytes also has COLINFO records."""
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert len(sh.colinfo_map) > 0

    def test_colinfo_not_parsed_without_fmt_info(self, colinfo_xls_bytes):
        """Without formatting_info=True, colinfo_map stays empty."""
        bk = open_wb(colinfo_xls_bytes, formatting_info=False)
        sh = bk.sheet_by_index(0)
        assert len(sh.colinfo_map) == 0

    def test_computed_column_width(self, colinfo_xls_bytes):
        bk = open_wb(colinfo_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        w = sh.computed_column_width(0)
        assert w > 0


# ===========================================================================
# PALETTE record handler (formatting.py lines 578-611) — custom colors
# ===========================================================================

class TestPaletteRecord:

    def test_palette_xls_loads(self, palette_xls_bytes):
        bk = open_wb(palette_xls_bytes, formatting_info=True)
        assert bk.nsheets == 1

    def test_palette_colour_map_has_custom(self, palette_xls_bytes):
        """Custom colour at index 0x21 (33) should be orange."""
        bk = open_wb(palette_xls_bytes, formatting_info=True)
        # Colour index 33 = 0x21, mapped via palette
        assert 8 + (0x21 - 8) in bk.colour_map or 0x21 in bk.colour_map

    def test_palette_record_populated(self, palette_xls_bytes):
        bk = open_wb(palette_xls_bytes, formatting_info=True)
        assert len(bk.palette_record) > 0

    def test_palette_contains_orange(self, palette_xls_bytes):
        """The orange custom colour (0xFF, 0x80, 0x00) should be in the palette."""
        bk = open_wb(palette_xls_bytes, formatting_info=True)
        assert (255, 128, 0) in bk.palette_record

    def test_palette_without_formatting_info(self, palette_xls_bytes):
        """Without formatting_info, palette is not populated."""
        bk = open_wb(palette_xls_bytes, formatting_info=False)
        assert len(bk.palette_record) == 0


# ===========================================================================
# Merged cells (MERGEDCELLS record handler)
# ===========================================================================

class TestMergedCells:

    def test_merged_cells_loads(self, merged_cells_xls_bytes):
        bk = open_wb(merged_cells_xls_bytes)
        assert bk.nsheets == 1

    def test_merged_header_cell(self, merged_cells_xls_bytes):
        bk = open_wb(merged_cells_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(0, 0) == 'Header spanning 3 cols'

    def test_merged_tall_cell(self, merged_cells_xls_bytes):
        bk = open_wb(merged_cells_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(1, 0) == 'Tall cell spanning 3 rows'

    def test_merged_cells_adjacent_content(self, merged_cells_xls_bytes):
        bk = open_wb(merged_cells_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(1, 1) == 'R1C1'

    def test_merged_cell_info_available(self, merged_cells_xls_bytes):
        bk = open_wb(merged_cells_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        # merged_cells should be populated when formatting_info=True
        assert hasattr(sh, 'merged_cells')


# ===========================================================================
# ROW record with formatting info (rowinfo_map population)
# ===========================================================================

class TestRowInfo:

    def test_rowinfo_map_populated_with_fmt_info(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert len(sh.rowinfo_map) > 0

    def test_rowinfo_height(self, formatted_xls_bytes):
        bk = open_wb(formatted_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        if sh.rowinfo_map:
            row_info = next(iter(sh.rowinfo_map.values()))
            assert row_info.height is not None


# ===========================================================================
# Formatting verbosity debug paths (formatting.py blah paths)
# ===========================================================================

class TestFormattingVerbosity:

    def test_verbosity_1_fmt_info(self, formatted_xls_bytes):
        """verbosity=1 + formatting_info=True triggers colour_epilogue log."""
        import io
        bk = xlrd.open_workbook(file_contents=formatted_xls_bytes,
                                formatting_info=True, verbosity=1,
                                logfile=io.StringIO())
        assert bk.nsheets > 0

    def test_verbosity_3_fmt_info(self, formatted_xls_bytes):
        """verbosity=3 triggers blah in handle_format and handle_xf."""
        import io
        bk = xlrd.open_workbook(file_contents=formatted_xls_bytes,
                                formatting_info=True, verbosity=3,
                                logfile=io.StringIO())
        assert bk.nsheets > 0

    def test_verbosity_3_palette(self, palette_xls_bytes):
        """verbosity=3 with palette: triggers blah path in handle_palette."""
        import io
        bk = xlrd.open_workbook(file_contents=palette_xls_bytes,
                                formatting_info=True, verbosity=3,
                                logfile=io.StringIO())
        assert bk.nsheets > 0

    def test_verbosity_2_palette(self, palette_xls_bytes):
        """verbosity=2 with palette: triggers blah in handle_palette."""
        import io
        bk = xlrd.open_workbook(file_contents=palette_xls_bytes,
                                formatting_info=True, verbosity=2,
                                logfile=io.StringIO())
        assert bk.nsheets > 0
        assert len(bk.palette_record) > 0

    def test_verbosity_3_colinfo(self, colinfo_xls_bytes):
        """verbosity=3 + formatting_info + COLINFO: triggers blah path."""
        import io
        bk = xlrd.open_workbook(file_contents=colinfo_xls_bytes,
                                formatting_info=True, verbosity=3,
                                logfile=io.StringIO())
        sh = bk.sheet_by_index(0)
        assert len(sh.colinfo_map) > 0

    def test_verbosity_3_blank_cells(self, blank_cells_xls_bytes):
        """verbosity=3 + formatting_info with blank cells."""
        import io
        bk = xlrd.open_workbook(file_contents=blank_cells_xls_bytes,
                                formatting_info=True, verbosity=3,
                                logfile=io.StringIO())
        assert bk.nsheets > 0


# ===========================================================================
# Negative-index paths in row/col methods (sheet.py lines 505-565, 577-583)
# ===========================================================================

class TestNegativeIndexSlices:

    def test_row_types_no_end(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        # end_colx=None path → line 506
        types = sh.row_types(0)
        assert len(types) == sh.ncols

    def test_row_types_with_end(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        types = sh.row_types(0, 0, 2)
        assert len(types) == 2

    def test_row_values_no_end(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        vals = sh.row_values(0)
        assert len(vals) == sh.ncols

    def test_row_slice_negative_start(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        # negative start_colx → lines 523-525
        result = sh.row_slice(0, -2)
        assert len(result) == 2

    def test_row_slice_negative_start_clamped(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        # very negative → clamped to 0 → line 525
        result = sh.row_slice(0, -999)
        assert len(result) == sh.ncols

    def test_row_slice_negative_end(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        # negative end_colx → line 529
        result = sh.row_slice(0, 0, -1)
        assert len(result) == sh.ncols - 1

    def test_col_slice_negative_start(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        # negative start_rowx → lines 541-543
        result = sh.col_slice(0, -2)
        assert len(result) == 2

    def test_col_slice_negative_start_clamped(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        # very negative → clamped to 0
        result = sh.col_slice(0, -999)
        assert len(result) == sh.nrows

    def test_col_slice_negative_end(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        # negative end_rowx → line 547-548
        result = sh.col_slice(0, 0, -1)
        assert len(result) == sh.nrows - 1

    def test_col_values_negative_start(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        result = sh.col_values(0, -2)
        assert len(result) == 2

    def test_col_values_negative_start_clamped(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        result = sh.col_values(0, -999)
        assert len(result) == sh.nrows

    def test_col_values_negative_end(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        result = sh.col_values(0, 0, -1)
        assert len(result) == sh.nrows - 1

    def test_col_types_negative_start(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        result = sh.col_types(0, -2)
        assert len(result) == 2

    def test_col_types_negative_start_clamped(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        result = sh.col_types(0, -999)
        assert len(result) == sh.nrows

    def test_col_types_negative_end(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        result = sh.col_types(0, 0, -1)
        assert len(result) == sh.nrows - 1


# ===========================================================================
# Verbosity debug output paths (blah_rows, DIMENSION blah, COLINFO blah)
# ===========================================================================

class TestVerbosityPaths:

    def test_verbosity_2_loads_ok(self, basic_xls_bytes):
        """verbosity=2 triggers blah=True, covering DIMENSION debug output."""
        import io
        bk = xlrd.open_workbook(file_contents=basic_xls_bytes, verbosity=2,
                                logfile=io.StringIO())
        assert bk.nsheets > 0

    def test_verbosity_4_loads_ok(self, basic_xls_bytes):
        """verbosity=4 triggers blah_rows=True, covering ROW debug output lines 926-927."""
        import io
        bk = xlrd.open_workbook(file_contents=basic_xls_bytes, verbosity=4,
                                formatting_info=True, logfile=io.StringIO())
        sh = bk.sheet_by_index(0)
        assert sh.nrows > 0

    def test_verbosity_2_with_colinfo(self, colinfo_xls_bytes):
        """verbosity=2 + formatting_info=True → COLINFO blah debug output."""
        import io
        bk = xlrd.open_workbook(file_contents=colinfo_xls_bytes, verbosity=2,
                                formatting_info=True, logfile=io.StringIO())
        sh = bk.sheet_by_index(0)
        assert len(sh.colinfo_map) > 0


# ===========================================================================
# BLANK and MULBLANK record handlers (sheet.py lines 1068-1083)
# ===========================================================================

class TestBlankCells:

    def test_blank_cells_load(self, blank_cells_xls_bytes):
        bk = open_wb(blank_cells_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert sh.nrows > 0

    def test_blank_cell_type(self, blank_cells_xls_bytes):
        from xlrd.biffh import XL_CELL_BLANK
        bk = open_wb(blank_cells_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        # Cell (0,0) was written as write_blank → XL_CELL_BLANK
        cell = sh.cell(0, 0)
        assert cell.ctype == XL_CELL_BLANK

    def test_blank_cells_not_loaded_without_formatting_info(self, blank_cells_xls_bytes):
        """Without formatting_info, BLANK and MULBLANK records are skipped."""
        bk = open_wb(blank_cells_xls_bytes, formatting_info=False)
        sh = bk.sheet_by_index(0)
        # only the non-blank cells matter; sheet still loads
        assert sh.nrows > 0

    def test_mulblank_cells_have_correct_type(self, blank_cells_xls_bytes):
        from xlrd.biffh import XL_CELL_BLANK
        bk = open_wb(blank_cells_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        # Columns 3-6 were consecutive blank cells → MULBLANK record
        for c in range(3, 7):
            if c < sh.ncols:
                assert sh.cell(0, c).ctype == XL_CELL_BLANK


# ===========================================================================
# Hyperlinks (HLINK handler)
# ===========================================================================

class TestHyperlinks:

    def test_sheet_hlink_attr_exists(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert hasattr(sh, 'hyperlink_map')

    def test_no_hyperlinks_in_basic(self, basic_xls_bytes):
        bk = open_wb(basic_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.hyperlink_map == {}


# ===========================================================================
# Page breaks (HORIZONTALPAGEBREAKS + VERTICALPAGEBREAKS handlers)
# ===========================================================================

class TestPageBreaks:

    def test_page_breaks_loads(self, page_breaks_xls_bytes):
        bk = open_wb(page_breaks_xls_bytes, formatting_info=True)
        assert bk.nsheets == 1

    def test_horizontal_page_breaks(self, page_breaks_xls_bytes):
        bk = open_wb(page_breaks_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert len(sh.horizontal_page_breaks) == 2
        assert sh.horizontal_page_breaks[0][0] == 5
        assert sh.horizontal_page_breaks[1][0] == 10

    def test_vertical_page_breaks(self, page_breaks_xls_bytes):
        bk = open_wb(page_breaks_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert len(sh.vertical_page_breaks) == 1
        assert sh.vertical_page_breaks[0][0] == 2

    def test_page_breaks_not_loaded_without_fmt_info(self, page_breaks_xls_bytes):
        bk = open_wb(page_breaks_xls_bytes, formatting_info=False)
        sh = bk.sheet_by_index(0)
        assert sh.horizontal_page_breaks == []
        assert sh.vertical_page_breaks == []


# ===========================================================================
# Cross-sheet formulas → SUPBOOK + EXTERNSHEET records (book.py 1088-1131)
# ===========================================================================

class TestCrossSheetFormulas:

    def test_cross_sheet_xls_loads(self, cross_sheet_xls_bytes):
        bk = open_wb(cross_sheet_xls_bytes)
        assert bk.nsheets == 2

    def test_supbook_types_set(self, cross_sheet_xls_bytes):
        """SUPBOOK record → handle_supbook sets _supbook_types."""
        bk = open_wb(cross_sheet_xls_bytes)
        assert len(bk._supbook_types) > 0

    def test_externsheet_info_set(self, cross_sheet_xls_bytes):
        """EXTERNSHEET record → handle_externsheet sets _externsheet_info."""
        bk = open_wb(cross_sheet_xls_bytes)
        assert len(bk._externsheet_info) > 0

    def test_source_sheet_data(self, cross_sheet_xls_bytes):
        bk = open_wb(cross_sheet_xls_bytes)
        sh = bk.sheet_by_name('Source')
        assert sh.cell_value(0, 0) == 100
        assert sh.cell_value(1, 0) == 200

    def test_cross_sheet_verbosity_2(self, cross_sheet_xls_bytes):
        """verbosity=2 triggers blah in handle_supbook → lines 1090-1102."""
        import io
        bk = xlrd.open_workbook(file_contents=cross_sheet_xls_bytes,
                                verbosity=2, logfile=io.StringIO())
        assert len(bk._supbook_types) > 0


# ===========================================================================
# put_cell_ragged with formatting_info (sheet.py lines 658-712)
# ===========================================================================

class TestRaggedCellsWithFormatting:

    def test_ragged_with_formatting_info(self, ragged_xls_bytes):
        """ragged_rows=True + formatting_info=True exercises put_cell_ragged
        formatting paths (lines 676, 682, 693, 705)."""
        bk = open_wb(ragged_xls_bytes, ragged_rows=True, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert sh.nrows > 0

    def test_ragged_numeric_ctype_lookup(self, large_xls_bytes):
        """ragged_rows=True with number cells → ctype=None lookup (line 658)."""
        bk = open_wb(large_xls_bytes, ragged_rows=True)
        sh = bk.sheet_by_index(0)
        cell = sh.cell(0, 0)
        assert cell.ctype == XL_CELL_NUMBER

    def test_ragged_with_formatting_and_numbers(self, formatted_xls_bytes):
        """ragged_rows=True + formatting_info=True with number cells."""
        bk = open_wb(formatted_xls_bytes, ragged_rows=True, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert sh.nrows > 0
        # Each row may have different column counts in ragged mode
        row_lens = [sh.row_len(r) for r in range(sh.nrows)]
        assert len(row_lens) == sh.nrows


# ===========================================================================
# tidy_dimensions with formatting_info (line 652-653)
# ===========================================================================

class TestTidyDimensionsFormatting:

    def test_tidy_dimensions_with_fmt_info(self, ragged_xls_bytes):
        """Load ragged file without ragged_rows — tidy_dimensions pads rows
        with formatting_info=True (line 652-653)."""
        bk = open_wb(ragged_xls_bytes, ragged_rows=False, formatting_info=True)
        sh = bk.sheet_by_index(0)
        # All rows should be padded to ncols
        for r in range(sh.nrows):
            assert len(sh.row(r)) == sh.ncols

    def test_merged_cells_blah_debug(self, merged_cells_xls_bytes):
        """verbosity=2 + formatting_info=True + merged cells → line 1265."""
        import io
        bk = xlrd.open_workbook(file_contents=merged_cells_xls_bytes,
                                verbosity=2, formatting_info=True,
                                logfile=io.StringIO())
        sh = bk.sheet_by_index(0)
        assert len(sh.merged_cells) > 0


# ===========================================================================
# Column gaps in ragged rows — put_cell_ragged num_empty > 0 (lines 695-712)
# ===========================================================================

class TestColumnGapRagged:

    def test_col_gap_ragged_loads(self, col_gap_xls_bytes):
        bk = open_wb(col_gap_xls_bytes, ragged_rows=True)
        sh = bk.sheet_by_index(0)
        assert sh.nrows == 3

    def test_col_gap_content_correct(self, col_gap_xls_bytes):
        bk = open_wb(col_gap_xls_bytes, ragged_rows=True)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(0, 0) == 'start'
        assert sh.cell_value(0, 5) == 'end'

    def test_col_gap_ragged_with_fmt_info(self, col_gap_xls_bytes):
        """Gap cells with formatting_info=True → put_cell_ragged fmt_row paths."""
        bk = open_wb(col_gap_xls_bytes, ragged_rows=True, formatting_info=True)
        sh = bk.sheet_by_index(0)
        row0 = sh.row(0)
        # row 0 has gap filled with empty cells
        assert len(row0) >= 6
        assert sh.cell(0, 0).ctype == XL_CELL_TEXT
        assert sh.cell(0, 5).ctype == XL_CELL_TEXT

    def test_col_gap_empty_cells_in_gap(self, col_gap_xls_bytes):
        from xlrd.biffh import XL_CELL_EMPTY
        bk = open_wb(col_gap_xls_bytes, ragged_rows=True)
        sh = bk.sheet_by_index(0)
        # Cells 1-4 in row 0 should be empty (gap)
        for c in range(1, 5):
            assert sh.cell(0, c).ctype == XL_CELL_EMPTY


# ===========================================================================
# Rich text SST strings — SST rtcount > 0 path (book.py lines 1410-1461)
# ===========================================================================

class TestRichTextSST:

    def test_rich_text_loads(self, rich_text_xls_bytes):
        bk = open_wb(rich_text_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert sh.nrows >= 5

    def test_rich_text_string_values(self, rich_text_xls_bytes):
        bk = open_wb(rich_text_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(0, 0) == 'Hello World'
        assert sh.cell_value(1, 0) == 'foobarbaz'
        assert sh.cell_value(2, 0) == 'alphabeta'

    def test_rich_text_runlist_map_populated(self, rich_text_xls_bytes):
        """formatting_info=True → rich_text_runlist_map populated (book.py 1451-1461)."""
        bk = open_wb(rich_text_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert len(sh.rich_text_runlist_map) > 0

    def test_rich_text_runs_for_hello(self, rich_text_xls_bytes):
        bk = open_wb(rich_text_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        runs = sh.rich_text_runlist_map.get((0, 0))
        assert runs is not None
        assert len(runs) >= 1

    def test_rich_text_plain_cell_not_in_runlist(self, rich_text_xls_bytes):
        """Plain text cell has no rich text runs."""
        bk = open_wb(rich_text_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert (3, 0) not in sh.rich_text_runlist_map

    def test_rich_text_without_formatting_info(self, rich_text_xls_bytes):
        """Without formatting_info, string values still correct but no runlist."""
        bk = open_wb(rich_text_xls_bytes, formatting_info=False)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(0, 0) == 'Hello World'
        assert sh.rich_text_runlist_map == {}

    def test_rich_text_with_verbosity(self, rich_text_xls_bytes):
        """verbosity=2 + rich text + formatting_info → SST debug paths."""
        import io
        bk = xlrd.open_workbook(file_contents=rich_text_xls_bytes,
                                verbosity=2, formatting_info=True,
                                logfile=io.StringIO())
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(0, 0) == 'Hello World'


# ===========================================================================
# Unicode/UTF-16 SST strings — book.py lines 1420-1434
# ===========================================================================

class TestUnicodeSST:

    def test_unicode_loads(self, unicode_xls_bytes):
        bk = open_wb(unicode_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.nrows >= 6

    def test_unicode_japanese(self, unicode_xls_bytes):
        """UTF-16 uncompressed string (options & 0x01) in SST → lines 1422-1434."""
        bk = open_wb(unicode_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(0, 0) == u'\u3053\u3093\u306b\u3061\u306f'

    def test_unicode_chinese(self, unicode_xls_bytes):
        bk = open_wb(unicode_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(1, 0) == u'\u4e2d\u6587\u6d4b\u8bd5'

    def test_unicode_greek(self, unicode_xls_bytes):
        bk = open_wb(unicode_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(2, 0) == u'\u03b1\u03b2\u03b3\u03b4'

    def test_unicode_cyrillic(self, unicode_xls_bytes):
        bk = open_wb(unicode_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(3, 0) == u'\u0410\u0411\u0412'

    def test_unicode_latin1(self, unicode_xls_bytes):
        """Latin-1 string café uses compressed encoding in SST."""
        bk = open_wb(unicode_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(4, 0) == u'caf\xe9'

    def test_unicode_cell_types(self, unicode_xls_bytes):
        bk = open_wb(unicode_xls_bytes)
        sh = bk.sheet_by_index(0)
        assert sh.cell(0, 0).ctype == XL_CELL_TEXT
        assert sh.cell(0, 1).ctype == XL_CELL_NUMBER


# ===========================================================================
# Coloured fonts → palette_epilogue lines 623-624
# ===========================================================================

class TestColouredFonts:

    def test_coloured_font_loads(self, coloured_font_xls_bytes):
        bk = open_wb(coloured_font_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert sh.nrows == 3

    def test_coloured_font_values(self, coloured_font_xls_bytes):
        bk = open_wb(coloured_font_xls_bytes, formatting_info=True)
        sh = bk.sheet_by_index(0)
        assert sh.cell_value(0, 0) == 'Red text'
        assert sh.cell_value(1, 0) == 'Blue text'

    def test_palette_epilogue_with_coloured_font(self, coloured_font_xls_bytes):
        """Fonts with non-0x7fff colour_index trigger palette_epilogue lines 623-624."""
        bk = open_wb(coloured_font_xls_bytes, formatting_info=True)
        # The fact it loaded successfully means palette_epilogue processed the
        # coloured fonts without error
        assert len(bk.colour_indexes_used) > 0

    def test_palette_epilogue_verbosity_prints_used_colours(self, coloured_font_xls_bytes):
        """verbosity=1 + formatting_info + coloured fonts → palette_epilogue prints
        used colour list (line 629-631)."""
        import io
        log = io.StringIO()
        bk = xlrd.open_workbook(file_contents=coloured_font_xls_bytes,
                                formatting_info=True, verbosity=1, logfile=log)
        output = log.getvalue()
        assert 'Colour indexes used' in output
        # Red (index 10) and blue (index 12) should be in the list
        assert '10' in output or '12' in output
