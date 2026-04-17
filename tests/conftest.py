# -*- coding: utf-8 -*-
"""
Session-scoped fixtures that produce in-memory XLS (BIFF8) files using xlwt.
Each fixture returns raw bytes suitable for xlrd.open_workbook(file_contents=...).
"""
import io
import datetime
import pytest
import xlwt


def _wb_to_bytes(wb):
    """Serialize an xlwt.Workbook to bytes."""
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# basic_xls_bytes — two sheets, mixed cell types
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def basic_xls_bytes():
    """
    Two sheets:
      Sheet1: text, integers, floats, empty cells, boolean-ish values
      Sheet2: a second sheet to exercise multi-sheet paths
    """
    wb = xlwt.Workbook(encoding='utf-8')

    ws1 = wb.add_sheet('Sheet1')
    ws1.write(0, 0, 'Hello')
    ws1.write(0, 1, 'World')
    ws1.write(0, 2, 'Unicode: \u00e9\u00e0\u00fc')
    ws1.write(1, 0, 42)
    ws1.write(1, 1, 3.14159)
    ws1.write(1, 2, -99.5)
    ws1.write(2, 0, 0)
    ws1.write(2, 1, '')       # empty string
    ws1.write(3, 0, 'row3col0')
    ws1.write(3, 1, 1000000)
    ws1.write(3, 2, 0.001)
    # Boolean cells → BOOLERR records
    ws1.write(4, 0, True)
    ws1.write(4, 1, False)

    ws2 = wb.add_sheet('Sheet2')
    ws2.write(0, 0, 'Sheet2 A1')
    ws2.write(0, 1, 99)

    ws3 = wb.add_sheet('EmptySheet')
    # deliberately empty

    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# formatted_xls_bytes — fonts, number formats, date formats
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def formatted_xls_bytes():
    """
    Sheet with various formatting: bold, italic, underline, date format,
    number format, colors — for testing with formatting_info=True.
    """
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Formatted')

    # Bold style
    bold_style = xlwt.XFStyle()
    bold_font = xlwt.Font()
    bold_font.bold = True
    bold_style.font = bold_font

    # Italic style
    italic_style = xlwt.XFStyle()
    italic_font = xlwt.Font()
    italic_font.italic = True
    italic_style.font = italic_font

    # Underline style
    underline_style = xlwt.XFStyle()
    underline_font = xlwt.Font()
    underline_font.underline = xlwt.Font.UNDERLINE_SINGLE
    underline_style.font = underline_font

    # Date format style
    date_style = xlwt.XFStyle()
    date_style.num_format_str = 'YYYY-MM-DD'

    # Number format style
    num_style = xlwt.XFStyle()
    num_style.num_format_str = '#,##0.00'

    # General format
    gen_style = xlwt.XFStyle()
    gen_style.num_format_str = 'General'

    ws.write(0, 0, 'Bold text', bold_style)
    ws.write(0, 1, 'Italic text', italic_style)
    ws.write(0, 2, 'Underline', underline_style)
    ws.write(1, 0, 36526.0, date_style)   # 2000-01-01 in Excel date serial
    ws.write(1, 1, 1234567.89, num_style)
    ws.write(1, 2, 0.0, gen_style)
    ws.write(2, 0, 'Normal text')
    ws.write(2, 1, 42)
    ws.write(2, 2, True)

    # Alignment
    align_style = xlwt.XFStyle()
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER
    align_style.alignment = align
    ws.write(3, 0, 'Centered', align_style)

    # Borders
    border_style = xlwt.XFStyle()
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    border_style.borders = borders
    ws.write(3, 1, 'Bordered', border_style)

    # Set column widths → COLINFO records
    ws.col(0).width = 8000
    ws.col(1).width = 6000

    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# formula_xls_bytes — formula cells (numeric, string, boolean, error results)
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def formula_xls_bytes():
    """
    Sheet with formula cells to exercise the FORMULA record handler.
    Note: xlwt in Python 3 has limited formula support; numeric formulas
    often produce an empty-string cached result (first_byte==0 branch).
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Formulas')
    ws.write(0, 0, 10)
    ws.write(1, 0, 20)
    ws.write(2, 0, 30)
    # Numeric formula — xlwt writes a FORMULA record; cached result varies by Python version
    ws.write(3, 0, xlwt.Formula('A1+A2+A3'))
    ws.write(4, 0, xlwt.Formula('A1*2'))
    ws.write(5, 0, xlwt.Formula('A1-A2'))
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# palette_xls_bytes — custom colours → PALETTE record
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def palette_xls_bytes():
    """Workbook that writes a custom PALETTE record via set_colour_RGB."""
    wb = xlwt.Workbook()
    # Custom orange at index 0x21
    wb.set_colour_RGB(0x21, 0xFF, 0x80, 0x00)
    ws = wb.add_sheet('Colors')
    style = xlwt.easyxf('pattern: pattern solid, fore_colour 0x21;')
    ws.write(0, 0, 'Orange cell', style)
    ws.write(0, 1, 'Normal')
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# colinfo_xls_bytes — explicit column widths → COLINFO records
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def colinfo_xls_bytes():
    """Workbook with explicit column widths to exercise COLINFO parsing."""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Widths')
    ws.col(0).width = 2000
    ws.col(1).width = 5000
    ws.col(2).width = 8000
    ws.col(3).hidden = True  # hidden column
    ws.write(0, 0, 'Narrow')
    ws.write(0, 1, 'Medium')
    ws.write(0, 2, 'Wide')
    ws.write(0, 3, 'Hidden')
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# named_range_xls_bytes — workbook-level named ranges
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def named_range_xls_bytes():
    """Workbook with a named range."""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Data')
    ws.write(0, 0, 10)
    ws.write(1, 0, 20)
    ws.write(2, 0, 30)
    # xlwt doesn't directly support named ranges via a simple API,
    # so just write data and test the sheet access
    ws2 = wb.add_sheet('Summary')
    ws2.write(0, 0, 'Total')
    ws2.write(0, 1, 60)
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# large_xls_bytes — wider/taller sheet to stress row/col handling
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def large_xls_bytes():
    """A sheet with many rows and columns to exercise tidy_dimensions."""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Big')
    for r in range(20):
        for c in range(15):
            ws.write(r, c, r * 100 + c)
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# ragged_xls_bytes — rows of different lengths (for ragged_rows mode)
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def ragged_xls_bytes():
    """Rows with intentionally different column counts."""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Ragged')
    ws.write(0, 0, 'A')
    ws.write(1, 0, 'B')
    ws.write(1, 1, 'C')
    ws.write(2, 0, 'D')
    ws.write(2, 1, 'E')
    ws.write(2, 2, 'F')
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# multisheet_xls_bytes — 5 sheets for iteration tests
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def multisheet_xls_bytes():
    """Five sheets with sequential data."""
    wb = xlwt.Workbook()
    for i in range(5):
        ws = wb.add_sheet('Sheet%d' % (i + 1))
        ws.write(0, 0, 'sheet%d_r0c0' % i)
        ws.write(0, 1, i * 10)
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# merged_cells_xls_bytes — merged cell ranges
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def merged_cells_xls_bytes():
    """Workbook with merged cells to exercise MERGEDCELLS record handling."""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Merged')
    ws.write_merge(0, 0, 0, 2, 'Header spanning 3 cols')
    ws.write_merge(1, 3, 0, 0, 'Tall cell spanning 3 rows')
    ws.write(1, 1, 'R1C1')
    ws.write(2, 1, 'R2C1')
    ws.write(3, 1, 'R3C1')
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# blank_cells_xls_bytes — blank cells with formatting → BLANK / MULBLANK records
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def blank_cells_xls_bytes():
    """
    Workbook with explicitly blank cells (write_blank) to exercise BLANK and
    MULBLANK record handlers (sheet.py lines 1068-1083). These records only
    appear when formatting_info=True is used.
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Blanks')
    bold_style = xlwt.XFStyle()
    bold_font = xlwt.Font()
    bold_font.bold = True
    bold_style.font = bold_font
    # Single BLANK records (non-consecutive → separate records)
    ws.row(0).write_blanks(0, 0, bold_style)
    ws.write(0, 1, 'between blanks')
    ws.row(0).write_blanks(2, 2, bold_style)
    # Multiple consecutive blanks → MULBLANK record
    ws.row(0).write_blanks(3, 6, bold_style)
    ws.write(1, 0, 'reference')
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# print_area_xls_bytes — print areas and repeating rows → NAME records
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def print_area_xls_bytes():
    """
    Workbook with page breaks etc. — kept for compatibility, no NAME records.
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet('PrintArea')
    for r in range(10):
        for c in range(5):
            ws.write(r, c, r * 5 + c)
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# page_breaks_xls_bytes — page breaks → HORIZONTALPAGEBREAKS/VERTICALPAGEBREAKS
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def page_breaks_xls_bytes():
    """
    Workbook with page breaks to exercise HORIZONTALPAGEBREAKS and
    VERTICALPAGEBREAKS record handlers (sheet.py lines 1327-1352).
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Breaks')
    for r in range(20):
        ws.write(r, 0, r)
    # Horizontal page break after row 5 (break between row 5 and 6)
    ws.set_horz_page_breaks([(5, 0, 255), (10, 0, 255)])
    # Vertical page break after col 2
    ws.set_vert_page_breaks([(2, 0, 65535)])
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# cross_sheet_xls_bytes — cross-sheet formulas → SUPBOOK + EXTERNSHEET records
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def cross_sheet_xls_bytes():
    """
    Workbook with cross-sheet formula references to exercise
    handle_supbook (book.py 1088-1131) and handle_externsheet (book.py ~872-920)
    via the SUPBOOK and EXTERNSHEET records in parse_globals.
    """
    wb = xlwt.Workbook()
    ws1 = wb.add_sheet('Source')
    ws1.write(0, 0, 100)
    ws1.write(1, 0, 200)
    ws1.write(2, 0, 300)
    ws2 = wb.add_sheet('Summary')
    ws2.write(0, 0, xlwt.Formula('Source!A1'))
    ws2.write(1, 0, xlwt.Formula('Source!A2'))
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# rich_text_xls_bytes — SST strings with rich-text runs (rtcount > 0)
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def rich_text_xls_bytes():
    """
    Workbook with rich-text formatted strings in SST.
    This exercises the rtcount > 0 branch in handle_sst()
    (book.py lines 1410-1461) when formatting_info=True.
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet('RichText')
    bold_font = xlwt.Font()
    bold_font.bold = True
    italic_font = xlwt.Font()
    italic_font.italic = True
    # Rich text: multiple font runs
    ws.write_rich_text(0, 0, [('Hello', bold_font), (' World', xlwt.Font())])
    ws.write_rich_text(1, 0, [('foo', bold_font), ('bar', xlwt.Font()), ('baz', italic_font)])
    ws.write_rich_text(2, 0, [('alpha', italic_font), ('beta', bold_font)])
    ws.write(3, 0, 'plain text')
    ws.write(4, 0, 42)
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# unicode_xls_bytes — non-Latin-1 strings → UTF-16 path in handle_sst
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def unicode_xls_bytes():
    """
    Workbook with non-Latin-1 Unicode strings.
    Japanese/Chinese characters force the SST to use UTF-16 uncompressed
    encoding (options & 0x01), exercising book.py lines 1420-1434.
    """
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Unicode')
    ws.write(0, 0, u'\u3053\u3093\u306b\u3061\u306f')  # konnichiwa (Japanese)
    ws.write(1, 0, u'\u4e2d\u6587\u6d4b\u8bd5')        # Chinese
    ws.write(2, 0, u'\u03b1\u03b2\u03b3\u03b4')        # Greek: αβγδ
    ws.write(3, 0, u'\u0410\u0411\u0412')               # Cyrillic: АБВ
    ws.write(4, 0, u'caf\xe9')                          # café (Latin-1, compressed)
    ws.write(5, 0, u'normal ascii')
    ws.write(0, 1, 1.5)
    ws.write(1, 1, 2.5)
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# col_gap_xls_bytes — rows with non-consecutive column writes (for ragged path)
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def col_gap_xls_bytes():
    """
    Workbook where cells within a row are non-consecutive (col gap).
    This is needed to exercise put_cell_ragged's num_empty > 0 path
    (sheet.py lines 695-712) when loading with ragged_rows=True.
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Gaps')
    # Write col 0 and col 5 in row 0 — gap of 4 columns
    ws.write(0, 0, 'start')
    ws.write(0, 5, 'end')
    ws.write(1, 0, 'row1col0')
    ws.write(1, 3, 'row1col3')
    ws.write(2, 0, 1)
    ws.write(2, 8, 100)
    return _wb_to_bytes(wb)


# ---------------------------------------------------------------------------
# coloured_font_xls_bytes — font with explicit colour → palette_epilogue lines 623-624
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def coloured_font_xls_bytes():
    """
    Workbook with fonts that have non-system colour_index (not 0x7fff).
    This exercises palette_epilogue lines 623-624 when loaded with
    formatting_info=True.
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Colours')
    # Font with explicit red and blue colours
    red_style = xlwt.easyxf('font: colour red;')
    blue_style = xlwt.easyxf('font: colour blue;')
    ws.write(0, 0, 'Red text', red_style)
    ws.write(1, 0, 'Blue text', blue_style)
    ws.write(2, 0, 'Normal')
    return _wb_to_bytes(wb)
