"""
Tests for Name.cell targeting uncovered lines (success and condition paths).
"""
import io
import struct
import pytest

import xlrd
from xlrd.book import Name
from xlrd.biffh import XLRDError
from xlrd.formula import Operand, Ref3D, oREF, oNUM


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_record(rec_type, data=b''):
    return struct.pack('<HH', rec_type, len(data)) + data


def make_minimal_xls_with_cell(sheet_names=None):
    """Build a minimal BIFF8 XLS stream with a NUMBER cell at row 0, col 0."""
    if sheet_names is None:
        sheet_names = ['Sheet1']

    bof_wbk = _make_record(0x0809, struct.pack('<HHHH', 0x0600, 0x0005, 0x0DBB, 0x07CC) + b'\x00\x00')
    codepage = _make_record(0x0042, struct.pack('<H', 1200))
    eof = _make_record(0x000A, b'')

    # Sheet stream: BOF + NUMBER cell + EOF
    sheet_streams = []
    for _name in sheet_names:
        bof_sheet = _make_record(0x0809, struct.pack('<HHHH', 0x0600, 0x0010, 0x0DBB, 0x07CC) + b'\x00\x00')
        number_data = struct.pack('<HHHd', 0, 0, 0, 42.0)
        number_record = _make_record(0x0203, number_data)
        sheet_eof = _make_record(0x000A, b'')
        sheet_streams.append(bof_sheet + number_record + sheet_eof)

    globals_without_eof = bof_wbk + codepage

    bs_placeholder = b''
    for name in sheet_names:
        name_bytes = name.encode('utf-16-le')
        bs_data = struct.pack('<IH', 0, 0) + struct.pack('<BB', len(name), 1) + name_bytes
        bs_placeholder += _make_record(0x0085, bs_data)

    globals_size = len(globals_without_eof) + len(bs_placeholder) + len(eof)

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


def _make_name_with_ref(book, shtxlo, shtxhi, rowxlo, rowxhi, colxlo, colxhi):
    """Create a Name object with an oREF result pointing to the given coords."""
    n = Name()
    n.book = book
    op = Operand()
    op.kind = oREF
    op.value = [Ref3D((shtxlo, shtxhi, rowxlo, rowxhi, colxlo, colxhi))]
    n.result = op
    return n


# ---------------------------------------------------------------------------
# Name.cell success path (lines 213, 214, 217, 218)
# ---------------------------------------------------------------------------

class TestNameCellSuccess:

    @pytest.fixture
    def book(self):
        return xlrd.open_workbook(file_contents=make_minimal_xls_with_cell(['Sheet1']))

    def test_cell_returns_cell_for_valid_single_cell_ref(self, book):
        """Name.cell returns a Cell when ref is a valid single cell."""
        n = _make_name_with_ref(book, 0, 1, 0, 1, 0, 1)

        result = n.cell()

        assert result is not None
        from xlrd.sheet import Cell
        assert isinstance(result, Cell)

    def test_cell_returns_correct_value(self, book):
        """Name.cell returns the cell with the value stored in the sheet."""
        n = _make_name_with_ref(book, 0, 1, 0, 1, 0, 1)

        result = n.cell()

        assert result.value == 42.0

    def test_cell_multi_row_ref_raises(self, book):
        """Name.cell raises XLRDError when ref spans multiple rows."""
        n = _make_name_with_ref(book, 0, 1, 0, 2, 0, 1)

        with pytest.raises(XLRDError):
            n.cell()

    def test_cell_multi_col_ref_raises(self, book):
        """Name.cell raises XLRDError when ref spans multiple columns."""
        n = _make_name_with_ref(book, 0, 1, 0, 1, 0, 2)

        with pytest.raises(XLRDError):
            n.cell()

    def test_cell_no_result_raises(self, book):
        """Name.cell raises XLRDError when result is None."""
        n = Name()
        n.book = book
        n.result = None

        with pytest.raises(XLRDError):
            n.cell()

    def test_cell_non_ref_kind_raises(self, book):
        """Name.cell raises XLRDError when result kind is not oREF."""
        n = Name()
        n.book = book
        op = Operand()
        op.kind = oNUM
        op.value = 42.0
        n.result = op

        with pytest.raises(XLRDError):
            n.cell()
