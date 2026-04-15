"""
Tests for xlrd/book.py targeting uncovered lines.
"""
import io
import struct
import sys
import pytest

import xlrd
from xlrd.book import (
    Book,
    Name,
    colname,
    display_cell_address,
    expand_cell_address,
    open_workbook_xls,
)
from xlrd.biffh import XLRDError


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_record(rec_type, data=b''):
    return struct.pack('<HH', rec_type, len(data)) + data


def make_minimal_xls(sheet_names=None):
    """Build a minimal BIFF8 XLS stream with the given sheet names."""
    if sheet_names is None:
        sheet_names = ['Sheet1']

    bof_wbk = _make_record(0x0809, struct.pack('<HHHH', 0x0600, 0x0005, 0x0DBB, 0x07CC) + b'\x00\x00')
    codepage = _make_record(0x0042, struct.pack('<H', 1200))
    eof = _make_record(0x000A, b'')

    sheet_streams = []
    for _name in sheet_names:
        bof_sheet = _make_record(0x0809, struct.pack('<HHHH', 0x0600, 0x0010, 0x0DBB, 0x07CC) + b'\x00\x00')
        sheet_eof = _make_record(0x000A, b'')
        sheet_streams.append(bof_sheet + sheet_eof)

    globals_without_eof = bof_wbk + codepage

    # First pass: compute total globals size to determine sheet offsets
    bs_placeholder = b''
    for name in sheet_names:
        name_bytes = name.encode('utf-16-le')
        bs_data = struct.pack('<IH', 0, 0) + struct.pack('<BB', len(name), 1) + name_bytes
        bs_placeholder += _make_record(0x0085, bs_data)

    globals_size = len(globals_without_eof) + len(bs_placeholder) + len(eof)

    # Second pass: build boundsheet records with correct offsets
    offset = globals_size
    final_bs_records = b''
    for i, name in enumerate(sheet_names):
        name_bytes = name.encode('utf-16-le')
        bs_data = struct.pack('<IH', offset, 0) + struct.pack('<BB', len(name), 1) + name_bytes
        final_bs_records += _make_record(0x0085, bs_data)
        offset += len(sheet_streams[i])

    full_stream = globals_without_eof + final_bs_records + eof
    for ss in sheet_streams:
        full_stream += ss
    return full_stream


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def simple_book():
    """A Book opened from a minimal one-sheet XLS stream."""
    return xlrd.open_workbook(file_contents=make_minimal_xls(['Sheet1']))


@pytest.fixture
def two_sheet_book():
    """A Book opened from a minimal two-sheet XLS stream."""
    return xlrd.open_workbook(file_contents=make_minimal_xls(['First', 'Second']))


@pytest.fixture
def on_demand_book():
    """A Book opened in on_demand mode."""
    return xlrd.open_workbook(
        file_contents=make_minimal_xls(['SheetA', 'SheetB']),
        on_demand=True,
    )


# ---------------------------------------------------------------------------
# open_workbook_xls
# ---------------------------------------------------------------------------

class TestOpenWorkbookXls:

    def test_open_from_file_contents_returns_book(self):
        bk = open_workbook_xls(file_contents=make_minimal_xls())
        assert isinstance(bk, Book)

    def test_open_from_file_contents_sheet_names(self):
        bk = open_workbook_xls(file_contents=make_minimal_xls(['Alpha', 'Beta']))
        assert bk.sheet_names() == ['Alpha', 'Beta']

    def test_open_invalid_contents_raises(self):
        with pytest.raises(Exception):
            open_workbook_xls(file_contents=b'\x00\x01\x02\x03\x04\x05\x06\x07GARBAGE')

    def test_open_empty_contents_raises(self):
        # Empty file_contents is falsy so it tries to open filename=None
        with pytest.raises((XLRDError, TypeError, Exception)):
            open_workbook_xls(file_contents=b'\x00' * 4)

    def test_open_on_demand_mode(self):
        bk = open_workbook_xls(
            file_contents=make_minimal_xls(['S1']),
            on_demand=True,
        )
        assert bk.on_demand is True
        assert bk.sheet_loaded(0) is False

    def test_open_from_file(self, tmp_path):
        xls_path = tmp_path / 'test.xls'
        xls_path.write_bytes(make_minimal_xls(['FileSheet']))
        bk = open_workbook_xls(filename=str(xls_path))
        assert bk.sheet_names() == ['FileSheet']

    def test_resources_released_after_open(self):
        # Resources should be released on normal (non-on_demand) open
        bk = open_workbook_xls(file_contents=make_minimal_xls())
        assert bk._resources_released == 1

    def test_resources_not_released_on_demand(self):
        bk = open_workbook_xls(
            file_contents=make_minimal_xls(),
            on_demand=True,
        )
        assert bk._resources_released == 0
        bk.release_resources()


# ---------------------------------------------------------------------------
# colname
# ---------------------------------------------------------------------------

class TestColname:

    def test_colname_single_letter_first(self):
        assert colname(0) == 'A'

    def test_colname_single_letter_last(self):
        assert colname(25) == 'Z'

    def test_colname_two_letter_first(self):
        assert colname(26) == 'AA'

    def test_colname_two_letter_az(self):
        assert colname(51) == 'AZ'

    def test_colname_two_letter_ba(self):
        assert colname(52) == 'BA'

    def test_colname_two_letter_zz(self):
        assert colname(701) == 'ZZ'

    def test_colname_three_letters(self):
        assert colname(702) == 'AAA'


# ---------------------------------------------------------------------------
# expand_cell_address
# ---------------------------------------------------------------------------

class TestExpandCellAddress:

    def test_absolute_row_absolute_col(self):
        outrow, outcol, relrow, relcol = expand_cell_address(5, 0x0003)
        assert outrow == 5
        assert outcol == 3
        assert relrow == 0
        assert relcol == 0

    def test_relative_row_positive(self):
        outrow, outcol, relrow, relcol = expand_cell_address(5, 0x8003)
        assert outrow == 5
        assert relrow == 1
        assert relcol == 0

    def test_relative_row_negative(self):
        # row >= 32768 wraps to negative
        outrow, outcol, relrow, relcol = expand_cell_address(40000, 0x8003)
        assert outrow == 40000 - 65536
        assert relrow == 1

    def test_relative_col_positive(self):
        outrow, outcol, relrow, relcol = expand_cell_address(5, 0x4003)
        assert outcol == 3
        assert relcol == 1
        assert relrow == 0

    def test_relative_col_negative(self):
        # col >= 128 wraps to negative
        outrow, outcol, relrow, relcol = expand_cell_address(5, 0x4080)
        assert outcol == 0x80 - 256
        assert relcol == 1

    def test_both_relative(self):
        outrow, outcol, relrow, relcol = expand_cell_address(5, 0xC003)
        assert relrow == 1
        assert relcol == 1


# ---------------------------------------------------------------------------
# display_cell_address
# ---------------------------------------------------------------------------

class TestDisplayCellAddress:

    def test_absolute_row_col(self):
        result = display_cell_address(0, 0, 0, 0)
        assert result == '$A$1'

    def test_absolute_row_col_offset(self):
        result = display_cell_address(2, 2, 0, 0)
        assert result == '$C$3'

    def test_relative_row_positive(self):
        result = display_cell_address(3, 0, 1, 0)
        assert '(*+3)' in result or '(*+' in result

    def test_relative_row_negative(self):
        result = display_cell_address(-2, 0, 1, 0)
        assert '(*-2)' in result

    def test_relative_row_zero(self):
        result = display_cell_address(0, 0, 1, 0)
        assert '(*+0)' in result

    def test_relative_col_positive(self):
        result = display_cell_address(0, 3, 0, 1)
        assert '(*+3)' in result

    def test_relative_col_negative(self):
        result = display_cell_address(0, -1, 0, 1)
        assert '(*-1)' in result

    def test_both_absolute_large_row_col(self):
        result = display_cell_address(9, 25, 0, 0)
        assert result == '$Z$10'


# ---------------------------------------------------------------------------
# Book.sheet_names
# ---------------------------------------------------------------------------

class TestBookSheetNames:

    def test_sheet_names_returns_list(self, two_sheet_book):
        names = two_sheet_book.sheet_names()
        assert names == ['First', 'Second']

    def test_sheet_names_returns_copy(self, two_sheet_book):
        names = two_sheet_book.sheet_names()
        names.append('Extra')
        assert 'Extra' not in two_sheet_book.sheet_names()


# ---------------------------------------------------------------------------
# Book.sheets
# ---------------------------------------------------------------------------

class TestBookSheets:

    def test_sheets_returns_all(self, two_sheet_book):
        sheets = two_sheet_book.sheets()
        assert len(sheets) == 2

    def test_sheets_returns_correct_names(self, two_sheet_book):
        sheets = two_sheet_book.sheets()
        assert sheets[0].name == 'First'
        assert sheets[1].name == 'Second'

    def test_sheets_on_demand_loads_all(self, on_demand_book):
        all_sheets = on_demand_book.sheets()
        assert len(all_sheets) == 2
        on_demand_book.release_resources()


# ---------------------------------------------------------------------------
# Book.sheet_by_index
# ---------------------------------------------------------------------------

class TestBookSheetByIndex:

    def test_sheet_by_index_first(self, two_sheet_book):
        sh = two_sheet_book.sheet_by_index(0)
        assert sh.name == 'First'

    def test_sheet_by_index_second(self, two_sheet_book):
        sh = two_sheet_book.sheet_by_index(1)
        assert sh.name == 'Second'

    def test_sheet_by_index_on_demand(self, on_demand_book):
        sh = on_demand_book.sheet_by_index(0)
        assert sh.name == 'SheetA'
        on_demand_book.release_resources()


# ---------------------------------------------------------------------------
# Book.__iter__
# ---------------------------------------------------------------------------

class TestBookIter:

    def test_iter_yields_all_sheets(self, two_sheet_book):
        names = [sh.name for sh in two_sheet_book]
        assert names == ['First', 'Second']

    def test_list_from_book(self, simple_book):
        sheets = list(simple_book)
        assert len(sheets) == 1
        assert sheets[0].name == 'Sheet1'


# ---------------------------------------------------------------------------
# Book.sheet_by_name
# ---------------------------------------------------------------------------

class TestBookSheetByName:

    def test_sheet_by_name_existing(self, two_sheet_book):
        sh = two_sheet_book.sheet_by_name('Second')
        assert sh.name == 'Second'

    def test_sheet_by_name_nonexistent_raises(self, two_sheet_book):
        with pytest.raises(XLRDError):
            two_sheet_book.sheet_by_name('NoSuchSheet')

    def test_sheet_by_name_error_message(self, simple_book):
        with pytest.raises(XLRDError, match='NoSheet'):
            simple_book.sheet_by_name('NoSheet')


# ---------------------------------------------------------------------------
# Book.__getitem__
# ---------------------------------------------------------------------------

class TestBookGetItem:

    def test_getitem_by_index(self, two_sheet_book):
        sh = two_sheet_book[0]
        assert sh.name == 'First'

    def test_getitem_by_name(self, two_sheet_book):
        sh = two_sheet_book['Second']
        assert sh.name == 'Second'

    def test_getitem_invalid_name_raises(self, simple_book):
        with pytest.raises(XLRDError):
            _ = simple_book['NonExistent']


# ---------------------------------------------------------------------------
# Book.sheet_loaded
# ---------------------------------------------------------------------------

class TestBookSheetLoaded:

    def test_sheet_loaded_false_before_access(self, on_demand_book):
        assert on_demand_book.sheet_loaded(0) is False
        on_demand_book.release_resources()

    def test_sheet_loaded_true_after_access(self, on_demand_book):
        on_demand_book.sheet_by_index(0)
        assert on_demand_book.sheet_loaded(0) is True
        on_demand_book.release_resources()

    def test_sheet_loaded_by_name(self, on_demand_book):
        assert on_demand_book.sheet_loaded('SheetA') is False
        on_demand_book.sheet_by_index(0)
        assert on_demand_book.sheet_loaded('SheetA') is True
        on_demand_book.release_resources()

    def test_sheet_loaded_invalid_name_raises(self, on_demand_book):
        with pytest.raises(XLRDError):
            on_demand_book.sheet_loaded('NoSuchSheet')
        on_demand_book.release_resources()


# ---------------------------------------------------------------------------
# Book.unload_sheet
# ---------------------------------------------------------------------------

class TestBookUnloadSheet:

    def test_unload_sheet_by_index(self, on_demand_book):
        on_demand_book.sheet_by_index(0)
        assert on_demand_book.sheet_loaded(0) is True
        on_demand_book.unload_sheet(0)
        assert on_demand_book.sheet_loaded(0) is False
        on_demand_book.release_resources()

    def test_unload_sheet_by_name(self, on_demand_book):
        on_demand_book.sheet_by_index(0)
        on_demand_book.unload_sheet('SheetA')
        assert on_demand_book.sheet_loaded('SheetA') is False
        on_demand_book.release_resources()

    def test_unload_invalid_name_raises(self, on_demand_book):
        with pytest.raises(XLRDError):
            on_demand_book.unload_sheet('DoesNotExist')
        on_demand_book.release_resources()


# ---------------------------------------------------------------------------
# Book.release_resources
# ---------------------------------------------------------------------------

class TestBookReleaseResources:

    def test_release_sets_flag(self, on_demand_book):
        assert on_demand_book._resources_released == 0
        on_demand_book.release_resources()
        assert on_demand_book._resources_released == 1

    def test_release_twice_no_error(self, on_demand_book):
        on_demand_book.release_resources()
        on_demand_book.release_resources()  # should not raise

    def test_release_makes_mem_none(self, on_demand_book):
        on_demand_book.release_resources()
        assert on_demand_book.mem is None

    def test_get_sheet_after_release_raises(self, on_demand_book):
        on_demand_book.release_resources()
        with pytest.raises(XLRDError):
            on_demand_book.get_sheet(0)


# ---------------------------------------------------------------------------
# Book.__enter__ / __exit__
# ---------------------------------------------------------------------------

class TestBookContextManager:

    def test_context_manager_returns_book(self):
        data = make_minimal_xls(['CM'])
        with xlrd.open_workbook(file_contents=data, on_demand=True) as bk:
            assert isinstance(bk, Book)
            assert bk.sheet_names() == ['CM']

    def test_context_manager_releases_resources(self):
        data = make_minimal_xls()
        with xlrd.open_workbook(file_contents=data, on_demand=True) as bk:
            pass
        assert bk._resources_released == 1

    def test_context_manager_releases_on_exception(self):
        data = make_minimal_xls()
        bk_ref = []
        try:
            with xlrd.open_workbook(file_contents=data, on_demand=True) as bk:
                bk_ref.append(bk)
                raise ValueError('test error')
        except ValueError:
            pass
        assert bk_ref[0]._resources_released == 1


# ---------------------------------------------------------------------------
# Name.cell (error path)
# ---------------------------------------------------------------------------

class TestNameCell:

    def test_name_cell_no_result_raises(self):
        n = Name()
        n.result = None
        n.book = xlrd.open_workbook(file_contents=make_minimal_xls())
        with pytest.raises(XLRDError, match='Not a constant absolute reference to a single cell'):
            n.cell()

    def test_name_cell_wrong_kind_raises(self):
        from xlrd.formula import Operand, oNUM
        n = Name()
        op = Operand()
        op.kind = oNUM
        op.value = 42
        n.result = op
        n.book = xlrd.open_workbook(file_contents=make_minimal_xls())
        with pytest.raises(XLRDError):
            n.cell()


# ---------------------------------------------------------------------------
# Name.area2d (error path)
# ---------------------------------------------------------------------------

class TestNameArea2d:

    def test_area2d_no_result_raises(self):
        n = Name()
        n.result = None
        n.book = xlrd.open_workbook(file_contents=make_minimal_xls())
        with pytest.raises(XLRDError, match='Not a constant absolute reference to a single area'):
            n.area2d()

    def test_area2d_wrong_kind_raises(self):
        from xlrd.formula import Operand, oNUM
        n = Name()
        op = Operand()
        op.kind = oNUM
        op.value = 42
        n.result = op
        n.book = xlrd.open_workbook(file_contents=make_minimal_xls())
        with pytest.raises(XLRDError):
            n.area2d()
