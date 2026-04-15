"""
Tests for Name.area2d targeting uncovered lines (success paths).
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


def make_minimal_xls(sheet_names=None):
    """Build a minimal BIFF8 XLS stream."""
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
# Name.area2d success paths (lines 250-261)
# ---------------------------------------------------------------------------

class TestNameArea2dSuccess:

    @pytest.fixture
    def book(self):
        return xlrd.open_workbook(file_contents=make_minimal_xls(['Sheet1']))

    def test_area2d_clipped_returns_sheet_and_bounds(self, book):
        """area2d with clipped=True returns clipped row/col bounds."""
        sh = book.sheet_by_index(0)
        n = _make_name_with_ref(book, 0, 1, 0, 5, 0, 3)

        result = n.area2d(clipped=True)

        assert result[0] is sh
        assert len(result) == 5
        rowxlo, rowxhi, colxlo, colxhi = result[1], result[2], result[3], result[4]
        assert 0 <= rowxlo <= rowxhi <= sh.nrows
        assert 0 <= colxlo <= colxhi <= sh.ncols

    def test_area2d_unclipped_returns_raw_bounds(self, book):
        """area2d with clipped=False returns the raw ref3d bounds."""
        sh = book.sheet_by_index(0)
        n = _make_name_with_ref(book, 0, 1, 2, 10, 1, 5)

        result = n.area2d(clipped=False)

        assert result[0] is sh
        assert result[1] == 2
        assert result[2] == 10
        assert result[3] == 1
        assert result[4] == 5

    def test_area2d_clipped_large_ref_clips_to_sheet(self, book):
        """area2d clipping reduces out-of-range refs to actual sheet size."""
        sh = book.sheet_by_index(0)
        n = _make_name_with_ref(book, 0, 1, 0, 65536, 0, 256)

        result = n.area2d(clipped=True)

        _, rowxlo, rowxhi, colxlo, colxhi = result
        assert rowxhi <= sh.nrows
        assert colxhi <= sh.ncols

    def test_area2d_default_clipped_true(self, book):
        """Default call to area2d uses clipped=True."""
        n = _make_name_with_ref(book, 0, 1, 0, 5, 0, 3)

        result_default = n.area2d()
        result_clipped = n.area2d(clipped=True)

        assert result_default == result_clipped

    def test_area2d_multiple_refs_raises(self, book):
        """area2d with multiple refs in value raises XLRDError."""
        n = Name()
        n.book = book
        op = Operand()
        op.kind = oREF
        op.value = [
            Ref3D((0, 1, 0, 5, 0, 3)),
            Ref3D((0, 1, 6, 10, 0, 3)),
        ]
        n.result = op

        with pytest.raises(XLRDError):
            n.area2d()

    def test_area2d_multi_sheet_ref_raises(self, book):
        """area2d with a ref spanning multiple sheets raises XLRDError."""
        n = _make_name_with_ref(book, 0, 2, 0, 5, 0, 3)

        with pytest.raises(XLRDError):
            n.area2d()
