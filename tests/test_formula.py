# -*- coding: utf-8 -*-
"""Unit tests for xlrd/formula.py utility functions."""
import io
import struct
from types import SimpleNamespace

import pytest

from xlrd.formula import (
    Operand,
    Ref3D,
    decompile_formula,
    dump_formula,
    FMLA_TYPE_CELL,
    _opr_eq,
    _opr_ge,
    _opr_gt,
    _opr_le,
    _opr_lt,
    _opr_ne,
    _opr_pow,
    adjust_cell_addr_biff8,
    adjust_cell_addr_biff_le7,
    cellname,
    cellnameabs,
    cellnamerel,
    colname,
    colnamerel,
    do_box_funcs,
    get_cell_addr,
    get_cell_range_addr,
    get_externsheet_local_range,
    get_externsheet_local_range_b57,
    nop,
    num2strg,
    quotedsheetname,
    rangename2d,
    rangename2drel,
    rangename3d,
    rangename3drel,
    rownamerel,
    sheetrange,
    sheetrangerel,
    tIsectFuncs,
    tRangeFuncs,
    oNUM,
    oSTRG,
    oUNK,
    oBOOL,
    oERR,
    oREF,
    oREL,
)


def make_bk(externsheet_info, all_sheets_map, locals_inx=0, addins_inx=999):
    bk = SimpleNamespace()
    bk.logfile = io.StringIO()
    bk._supbook_locals_inx = locals_inx
    bk._supbook_addins_inx = addins_inx
    bk._externsheet_info = externsheet_info
    bk._all_sheets_map = all_sheets_map
    return bk


def make_book_with_sheets(names):
    bk = SimpleNamespace()
    bk.sheet_names = lambda: names
    return bk


# ===========================================================================
# do_box_funcs
# ===========================================================================

class TestDoBoxFuncs:
    def test_tRangeFuncs(self):
        class Box:
            def __init__(self, coords):
                self.coords = coords

        boxa = Box((1, 5, 2, 6, 3, 7))
        boxb = Box((2, 3, 8, 4, 9, 5))
        result = do_box_funcs(tRangeFuncs, boxa, boxb)
        assert result == (min(1, 2), max(5, 3), min(2, 8), max(6, 4), min(3, 9), max(7, 5))

    def test_tIsectFuncs(self):
        class Box:
            def __init__(self, coords):
                self.coords = coords

        boxa = Box((1, 5, 2, 6, 3, 7))
        boxb = Box((2, 3, 8, 4, 9, 5))
        result = do_box_funcs(tIsectFuncs, boxa, boxb)
        assert result == (max(1, 2), min(5, 3), max(2, 8), min(6, 4), max(3, 9), min(7, 5))


# ===========================================================================
# adjust_cell_addr_biff8
# ===========================================================================

class TestAdjustCellAddrBiff8:
    def test_absolute_reldelta_true(self):
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(10, 5, True)
        assert rowx == 10
        assert colx == 5
        assert row_rel == 0
        assert col_rel == 0

    def test_both_relative_reldelta_true_no_overflow(self):
        colval = (1 << 15) | (1 << 14) | 5  # row_rel=1, col_rel=1, colx=5
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(10, colval, True)
        assert rowx == 10
        assert colx == 5
        assert row_rel == 1
        assert col_rel == 1

    def test_row_relative_overflow_reldelta_true(self):
        colval = (1 << 15) | 3  # row_rel=1, colx=3
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(40000, colval, True)
        assert rowx == 40000 - 65536
        assert col_rel == 0

    def test_col_relative_overflow_reldelta_true(self):
        colval = (1 << 14) | 200  # col_rel=1, colx=200 (>= 128)
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(5, colval, True)
        assert colx == 200 - 256
        assert row_rel == 0
        assert col_rel == 1

    def test_reldelta_false_absolute(self):
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(10, 5, False, browx=3, bcolx=2)
        assert rowx == 10
        assert colx == 5
        assert row_rel == 0
        assert col_rel == 0

    def test_reldelta_false_relative_subtracts_base(self):
        colval = (1 << 15) | (1 << 14) | 7  # both rel, colx=7
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(15, colval, False, browx=3, bcolx=2)
        assert rowx == 15 - 3
        assert colx == 7 - 2
        assert row_rel == 1
        assert col_rel == 1


# ===========================================================================
# adjust_cell_addr_biff_le7
# ===========================================================================

class TestAdjustCellAddrBiffLe7:
    def test_absolute(self):
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(10, 5, True)
        assert rowx == 10
        assert colx == 5
        assert row_rel == 0
        assert col_rel == 0

    def test_both_relative_no_overflow(self):
        rowval = (1 << 15) | (1 << 14) | 10  # row_rel=1, col_rel=1, rowx=10
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(rowval, 5, True)
        assert rowx == 10
        assert colx == 5
        assert row_rel == 1
        assert col_rel == 1

    def test_row_relative_overflow(self):
        rowval = (1 << 15) | 10000  # row_rel=1, rowx=10000 (>= 8192)
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(rowval, 5, True)
        assert rowx == 10000 - 16384

    def test_col_relative_overflow(self):
        rowval = (1 << 14) | 5  # col_rel=1, rowx=5
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(rowval, 200, True)
        assert colx == 200 - 256

    def test_reldelta_false_subtracts_base(self):
        rowval = (1 << 15) | (1 << 14) | 15
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(rowval, 7, False, browx=3, bcolx=2)
        assert rowx == 15 - 3
        assert colx == 7 - 2


# ===========================================================================
# get_cell_addr
# ===========================================================================

class TestGetCellAddr:
    def test_biff8_absolute(self):
        data = struct.pack("<HH", 5, 7)  # rowval=5, colval=7
        rowx, colx, row_rel, col_rel = get_cell_addr(data, 0, 80, True)
        assert rowx == 5
        assert colx == 7
        assert row_rel == 0
        assert col_rel == 0

    def test_biff_le7(self):
        data = struct.pack("<HB", 5, 7)  # rowval=5, colval=7
        rowx, colx, row_rel, col_rel = get_cell_addr(data, 0, 70, True)
        assert rowx == 5
        assert colx == 7
        assert row_rel == 0
        assert col_rel == 0


# ===========================================================================
# get_cell_range_addr
# ===========================================================================

class TestGetCellRangeAddr:
    def test_biff8(self):
        data = struct.pack("<HHHH", 0, 10, 0, 5)  # row1=0, row2=10, col1=0, col2=5
        res1, res2 = get_cell_range_addr(data, 0, 80, True)
        assert res1[0] == 0
        assert res2[0] == 10

    def test_biff_le7(self):
        data = struct.pack("<HHBB", 0, 10, 0, 5)
        res1, res2 = get_cell_range_addr(data, 0, 70, True)
        assert res1[0] == 0
        assert res2[0] == 10


# ===========================================================================
# get_externsheet_local_range
# ===========================================================================

class TestGetExternsheetLocalRange:
    def test_refx_out_of_range(self):
        bk = make_bk([], [])
        assert get_externsheet_local_range(bk, 5) == (-101, -101)

    def test_addins(self):
        bk = make_bk([(1, 0xFFFE, 0xFFFE)], [0], addins_inx=1)
        assert get_externsheet_local_range(bk, 0) == (-5, -5)

    def test_addins_with_blah(self):
        bk = make_bk([(1, 0xFFFE, 0xFFFE)], [0], addins_inx=1)
        assert get_externsheet_local_range(bk, 0, blah=1) == (-5, -5)

    def test_external_reference(self):
        bk = make_bk([(2, 0, 1)], [0, 1])
        assert get_externsheet_local_range(bk, 0) == (-4, -4)

    def test_external_reference_with_blah(self):
        bk = make_bk([(2, 0, 1)], [0, 1])
        assert get_externsheet_local_range(bk, 0, blah=1) == (-4, -4)

    def test_unspecified_sheet(self):
        bk = make_bk([(0, 0xFFFE, 0xFFFE)], [0])
        assert get_externsheet_local_range(bk, 0) == (-1, -1)

    def test_unspecified_sheet_with_blah(self):
        bk = make_bk([(0, 0xFFFE, 0xFFFE)], [0])
        assert get_externsheet_local_range(bk, 0, blah=1) == (-1, -1)

    def test_deleted_sheet(self):
        bk = make_bk([(0, 0xFFFF, 0xFFFF)], [0])
        assert get_externsheet_local_range(bk, 0) == (-2, -2)

    def test_deleted_sheet_with_blah(self):
        bk = make_bk([(0, 0xFFFF, 0xFFFF)], [0])
        assert get_externsheet_local_range(bk, 0, blah=1) == (-2, -2)

    def test_invalid_sheet_range(self):
        bk = make_bk([(0, 0, 5)], [0, 1])
        assert get_externsheet_local_range(bk, 0) == (-102, -102)

    def test_invalid_sheet_range_with_blah(self):
        bk = make_bk([(0, 0, 5)], [0, 1])
        assert get_externsheet_local_range(bk, 0, blah=1) == (-102, -102)

    def test_macro_sheet(self):
        bk = make_bk([(0, 0, 1)], [-1, 0])
        assert get_externsheet_local_range(bk, 0) == (-3, -3)

    def test_valid_single_sheet(self):
        bk = make_bk([(0, 0, 0)], [0, 1, 2])
        assert get_externsheet_local_range(bk, 0) == (0, 0)

    def test_valid_range(self):
        bk = make_bk([(0, 0, 1)], [0, 1, 2])
        assert get_externsheet_local_range(bk, 0) == (0, 1)


# ===========================================================================
# get_externsheet_local_range_b57
# ===========================================================================

class TestGetExternsheetLocalRangeB57:
    def test_external_reference(self):
        bk = make_bk([], [0, 1])
        assert get_externsheet_local_range_b57(bk, 1, 0, 1) == (-4, -4)

    def test_external_reference_with_blah(self):
        bk = make_bk([], [0, 1])
        assert get_externsheet_local_range_b57(bk, 1, 0, 1, blah=1) == (-4, -4)

    def test_deleted_sheets(self):
        bk = make_bk([], [0, 1])
        assert get_externsheet_local_range_b57(bk, 0, -1, -1) == (-2, -2)

    def test_invalid_range(self):
        bk = make_bk([], [0, 1])
        assert get_externsheet_local_range_b57(bk, 0, 0, 5) == (-103, -103)

    def test_invalid_range_with_blah(self):
        bk = make_bk([], [0, 1])
        assert get_externsheet_local_range_b57(bk, 0, 0, 5, blah=1) == (-103, -103)

    def test_macro_sheet(self):
        bk = make_bk([], [-1, 0])
        assert get_externsheet_local_range_b57(bk, 0, 0, 1) == (-3, -3)

    def test_valid_range(self):
        bk = make_bk([], [0, 1, 2])
        assert get_externsheet_local_range_b57(bk, 0, 0, 1) == (0, 1)


# ===========================================================================
# Operand
# ===========================================================================

class TestOperand:
    def test_default_init(self):
        op = Operand()
        assert op.kind == oUNK
        assert op.value is None
        assert op.rank == 0
        assert op.text == '?'

    def test_init_with_kind_and_value(self):
        op = Operand(oNUM, 42.0, 30, '42')
        assert op.kind == oNUM
        assert op.value == 42.0
        assert op.rank == 30
        assert op.text == '42'

    def test_init_kind_none_keeps_class_default(self):
        op = Operand(None, 5.0)
        assert op.kind == oUNK

    def test_repr_known_kind(self):
        op = Operand(oNUM, 3.14, 0, '3.14')
        r = repr(op)
        assert 'oNUM' in r
        assert '3.14' in r

    def test_repr_unknown_kind(self):
        op = Operand(99, None, 0, 'x')
        r = repr(op)
        assert '?Unknown kind?' in r

    def test_repr_string_operand(self):
        op = Operand(oSTRG, 'hello', 0, '"hello"')
        r = repr(op)
        assert 'oSTRG' in r
        assert 'hello' in r


# ===========================================================================
# Ref3D
# ===========================================================================

class TestRef3D:
    def test_init_without_relflags(self):
        r = Ref3D((0, 1, 2, 10, 3, 8))
        assert r.coords == (0, 1, 2, 10, 3, 8)
        assert r.relflags == (0, 0, 0, 0, 0, 0)
        assert r.shtxlo == 0
        assert r.shtxhi == 1
        assert r.rowxlo == 2
        assert r.rowxhi == 10
        assert r.colxlo == 3
        assert r.colxhi == 8

    def test_init_with_relflags(self):
        r = Ref3D((0, 1, 2, 10, 3, 8, 0, 0, 1, 1, 0, 0))
        assert r.coords == (0, 1, 2, 10, 3, 8)
        assert r.relflags == (0, 0, 1, 1, 0, 0)

    def test_repr_no_relflags(self):
        r = Ref3D((0, 1, 2, 10, 3, 8))
        result = repr(r)
        assert 'coords' in result
        assert 'relflags' not in result

    def test_repr_with_relflags(self):
        r = Ref3D((0, 1, 2, 10, 3, 8, 0, 0, 1, 1, 0, 0))
        result = repr(r)
        assert 'coords' in result
        assert 'relflags' in result


# ===========================================================================
# Simple operator functions
# ===========================================================================

class TestSimpleOperators:
    def test_nop(self):
        assert nop(5) == 5
        assert nop('hello') == 'hello'
        assert nop(None) is None

    def test_opr_pow(self):
        assert _opr_pow(2, 3) == 8
        assert _opr_pow(3, 2) == 9

    def test_opr_lt(self):
        assert _opr_lt(1, 2) is True
        assert _opr_lt(2, 1) is False

    def test_opr_le(self):
        assert _opr_le(1, 1) is True
        assert _opr_le(2, 1) is False

    def test_opr_eq(self):
        assert _opr_eq(1, 1) is True
        assert _opr_eq(1, 2) is False

    def test_opr_ge(self):
        assert _opr_ge(2, 1) is True
        assert _opr_ge(1, 2) is False

    def test_opr_gt(self):
        assert _opr_gt(2, 1) is True
        assert _opr_gt(1, 2) is False

    def test_opr_ne(self):
        assert _opr_ne(1, 2) is True
        assert _opr_ne(1, 1) is False


# ===========================================================================
# num2strg
# ===========================================================================

class TestNum2Strg:
    def test_integer_float(self):
        assert num2strg(3.0) == '3'

    def test_non_integer_float(self):
        assert num2strg(3.5) == '3.5'

    def test_large_integer(self):
        assert num2strg(1000.0) == '1000'

    def test_integer(self):
        assert num2strg(42) == '42'


# ===========================================================================
# colname
# ===========================================================================

class TestColname:
    def test_single_letter(self):
        assert colname(0) == 'A'
        assert colname(7) == 'H'
        assert colname(25) == 'Z'

    def test_two_letters(self):
        assert colname(26) == 'AA'
        assert colname(27) == 'AB'


# ===========================================================================
# cellname and cellnameabs
# ===========================================================================

class TestCellname:
    def test_cellname(self):
        assert cellname(5, 7) == 'H6'
        assert cellname(0, 0) == 'A1'

    def test_cellnameabs_default(self):
        assert cellnameabs(5, 7) == '$H$6'
        assert cellnameabs(0, 0) == '$A$1'

    def test_cellnameabs_r1c1(self):
        assert cellnameabs(5, 7, r1c1=1) == 'R6C8'


# ===========================================================================
# rownamerel
# ===========================================================================

class TestRownamerel:
    def test_absolute_r1c1(self):
        assert rownamerel(5, 0, browx=0, r1c1=1) == 'R6'

    def test_absolute_a1(self):
        assert rownamerel(5, 0, browx=0, r1c1=0) == '$6'

    def test_relative_with_base(self):
        assert rownamerel(5, 1, browx=10) == '16'

    def test_relative_no_base_nonzero(self):
        assert rownamerel(5, 1) == 'R[5]'

    def test_relative_no_base_zero(self):
        assert rownamerel(0, 1) == 'R'

    def test_absolute_no_base_forces_r1c1(self):
        # browx=None forces r1c1=True
        assert rownamerel(5, 0) == 'R6'


# ===========================================================================
# colnamerel
# ===========================================================================

class TestColnamerel:
    def test_absolute_r1c1(self):
        assert colnamerel(5, 0, bcolx=0, r1c1=1) == 'C6'

    def test_absolute_a1(self):
        assert colnamerel(5, 0, bcolx=0, r1c1=0) == '$F'

    def test_relative_with_base(self):
        assert colnamerel(5, 1, bcolx=10) == 'P'

    def test_relative_no_base_nonzero(self):
        assert colnamerel(5, 1) == 'C[5]'

    def test_relative_no_base_zero(self):
        assert colnamerel(0, 1) == 'C'

    def test_absolute_no_base_forces_r1c1(self):
        assert colnamerel(5, 0) == 'C6'


# ===========================================================================
# cellnamerel
# ===========================================================================

class TestCellnamerel:
    def test_both_absolute(self):
        result = cellnamerel(5, 7, 0, 0)
        assert result == '$H$6'

    def test_both_relative_no_base(self):
        result = cellnamerel(5, 7, 1, 1)
        assert 'R' in result and 'C' in result

    def test_both_relative_with_base(self):
        result = cellnamerel(5, 7, 1, 1, browx=10, bcolx=5)
        assert result == 'M16'

    def test_row_rel_no_base_forces_r1c1(self):
        result = cellnamerel(5, 7, 1, 0)
        assert 'R' in result


# ===========================================================================
# rangename2d
# ===========================================================================

class TestRangename2d:
    def test_single_cell(self):
        result = rangename2d(5, 6, 7, 8)
        assert result == '$H$6'

    def test_range(self):
        result = rangename2d(5, 20, 7, 10)
        assert result == '$H$6:$J$20'

    def test_r1c1_returns_none(self):
        result = rangename2d(5, 20, 7, 10, r1c1=1)
        assert result is None


# ===========================================================================
# rangename2drel
# ===========================================================================

class TestRangename2drel:
    def test_absolute_cells(self):
        result = rangename2drel((5, 20, 7, 10), (0, 0, 0, 0))
        assert '$H$6' in result
        assert '$J$20' in result

    def test_relative_no_base(self):
        result = rangename2drel((5, 20, 7, 10), (1, 1, 1, 1))
        assert 'R' in result and 'C' in result

    def test_relative_with_base(self):
        result = rangename2drel((5, 20, 7, 10), (1, 1, 1, 1), browx=10, bcolx=5)
        assert ':' in result


# ===========================================================================
# quotedsheetname
# ===========================================================================

class TestQuotedsheetname:
    def test_simple_name(self):
        assert quotedsheetname(['Sheet1'], 0) == 'Sheet1'

    def test_name_with_space(self):
        assert quotedsheetname(['My Sheet'], 0) == "'My Sheet'"

    def test_name_with_apostrophe(self):
        assert quotedsheetname(["Sheet's"], 0) == "'Sheet''s'"

    def test_negative_minus1(self):
        assert quotedsheetname([], -1) == "'?internal; any sheet?'"

    def test_negative_minus2(self):
        assert quotedsheetname([], -2) == "'internal; deleted sheet'"

    def test_negative_minus3(self):
        assert quotedsheetname([], -3) == "'internal; macro sheet'"

    def test_negative_minus4(self):
        assert quotedsheetname([], -4) == '<<external>>'

    def test_unknown_negative(self):
        result = quotedsheetname([], -99)
        assert '?error' in result


# ===========================================================================
# sheetrange
# ===========================================================================

class TestSheetrange:
    def test_single_sheet(self):
        bk = make_book_with_sheets(['Sheet1', 'Sheet2', 'Sheet3'])
        assert sheetrange(bk, 0, 1) == 'Sheet1'

    def test_multi_sheet(self):
        bk = make_book_with_sheets(['Sheet1', 'Sheet2', 'Sheet3'])
        assert sheetrange(bk, 0, 3) == 'Sheet1:Sheet3'


# ===========================================================================
# sheetrangerel
# ===========================================================================

class TestSheetrangerel:
    def test_absolute_sheets(self):
        bk = make_book_with_sheets(['Sheet1', 'Sheet2'])
        result = sheetrangerel(bk, (0, 1), (0, 0))
        assert result == 'Sheet1'

    def test_current_sheet_relative(self):
        bk = make_book_with_sheets(['Sheet1', 'Sheet2'])
        result = sheetrangerel(bk, (0, 1), (1, 1))
        assert result == ''


# ===========================================================================
# rangename3d
# ===========================================================================

class TestRangename3d:
    def test_basic(self):
        bk = make_book_with_sheets(['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4'])
        r = Ref3D((1, 4, 5, 20, 7, 10))
        result = rangename3d(bk, r)
        assert 'Sheet2' in result
        assert '$H$6' in result


# ===========================================================================
# rangename3drel
# ===========================================================================

class TestRangename3drel:
    def test_with_sheet_r1c1(self):
        bk = make_book_with_sheets(['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4'])
        r = Ref3D((0, 1, -32, -22, -13, 13, 0, 0, 1, 1, 1, 1))
        result = rangename3drel(bk, r, r1c1=1)
        assert 'Sheet1' in result
        assert 'R[-32]' in result

    def test_current_sheet_no_prefix(self):
        bk = make_book_with_sheets(['Sheet1', 'Sheet2'])
        r = Ref3D((0, 1, 5, 20, 7, 10, 1, 1, 0, 0, 0, 0))
        result = rangename3drel(bk, r)
        assert '!' not in result


# ===========================================================================
# dump_formula
# ===========================================================================

def make_dump_bk():
    """Create a minimal book object for dump_formula tests."""
    bk = make_bk(
        externsheet_info=[(0, 0, 0)],
        all_sheets_map=[0],
        locals_inx=0,
        addins_inx=999,
    )
    return bk


class TestDumpFormula:
    def test_empty_formula(self):
        bk = make_dump_bk()
        dump_formula(bk, b'', 0, 80, 0)

    def test_assertion_error_bv_lt_80(self):
        bk = make_dump_bk()
        with pytest.raises(AssertionError):
            dump_formula(bk, b'', 0, 70, 0)

    def test_blah_logging_prints_to_logfile(self):
        bk = make_dump_bk()
        data = b'\x03'  # opcode 0x03 (Add), optype=0, sz=1
        dump_formula(bk, data, 1, 80, 0, blah=1)
        output = bk.logfile.getvalue()
        assert 'dump_formula' in output

    def test_texp_opcode_optype0(self):
        bk = make_dump_bk()
        # opcode=0x01 (tExp), optype=0, sz=5: rowx, colx packed as <HH
        data = b'\x01' + struct.pack('<HH', 2, 3)
        dump_formula(bk, data, 5, 80, 0)

    def test_ttbl_opcode_optype0(self):
        bk = make_dump_bk()
        # opcode=0x02 (tTbl), optype=0, sz=5
        data = b'\x02' + struct.pack('<HH', 1, 4)
        dump_formula(bk, data, 5, 80, 0)

    def test_tattr_choose_subop(self):
        bk = make_dump_bk()
        # opcode=0x19 (tAttr), optype=0, subop=0x04 (Choose), nc=2 -> sz=2*2+6=10
        data = b'\x19' + struct.pack('<BH', 0x04, 2) + b'\x00' * 7
        dump_formula(bk, data, 10, 80, 0)

    def test_tattr_other_subop(self):
        bk = make_dump_bk()
        # opcode=0x19 (tAttr), optype=0, subop=0x01 (Volatile) -> sz=4
        data = b'\x19' + struct.pack('<BH', 0x01, 0)
        dump_formula(bk, data, 4, 80, 0)

    def test_tstr_biff8_path(self):
        bk = make_dump_bk()
        # opcode=0x17 (tStr), optype=0, bv=80 -> unpack_unicode_update_pos
        # lenlen=1: nchars byte + options byte (0=8-bit) + 3 chars
        data = b'\x17\x03\x00abc'  # nchars=3, options=0, chars='abc'
        dump_formula(bk, data, 6, 80, 0)

    def test_else_branch_dud_size_optype0(self):
        # opcode=0x1A, optype=0 -> opx=26, sztab4[26]=-2 -> dud size, returns early
        bk = make_dump_bk()
        data = b'\x1a'
        dump_formula(bk, data, 1, 80, 0)
        assert '**** Dud size' in bk.logfile.getvalue()

    def test_else_branch_valid_size_optype0(self):
        # opcode=0x03 (Add), optype=0 -> opx=3, sztab4[3]=1, advances normally
        bk = make_dump_bk()
        data = b'\x03'
        dump_formula(bk, data, 1, 80, 0)

    def test_tarray_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x20 -> optype=1, opcode=0x00 (tArray), opx=32, sz=9: just passes
        data = b'\x20' + b'\x00' * 8
        dump_formula(bk, data, 9, 80, 0)

    def test_tfunc_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x21 -> optype=1, opcode=0x01 (tFunc), bv=80 -> nb=2, sz=3
        data = b'\x21' + struct.pack('<H', 4)
        dump_formula(bk, data, 3, 80, 0)

    def test_tfuncvar_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x22 -> optype=1, opcode=0x02 (tFuncVar), bv=80 -> nb=2, sz=4
        data = b'\x22' + struct.pack('<BH', 2, 4)
        dump_formula(bk, data, 4, 80, 0)

    def test_tname_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x23 -> optype=1, opcode=0x03 (tName), opx=35, sz=5
        data = b'\x23' + struct.pack('<H', 1) + b'\x00\x00'
        dump_formula(bk, data, 5, 80, 0)

    def test_tref_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x24 -> optype=1, opcode=0x04 (tRef), opx=36, sz=5
        data = b'\x24' + struct.pack('<HH', 3, 5)
        dump_formula(bk, data, 5, 80, 0)

    def test_tarea_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x25 -> optype=1, opcode=0x05 (tArea), opx=37, sz=9
        data = b'\x25' + struct.pack('<HH', 3, 5) + struct.pack('<HH', 10, 15)
        dump_formula(bk, data, 9, 80, 0)

    def test_tmemfunc_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x29 -> optype=1, opcode=0x09 (tMemFunc), opx=41, sz=3
        data = b'\x29' + struct.pack('<H', 4)
        dump_formula(bk, data, 3, 80, 0)

    def test_trefn_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x2C -> optype=1, opcode=0x0C (tRefN), opx=44, sz=5; always reldelta=1
        data = b'\x2c' + struct.pack('<HH', 3, 5)
        dump_formula(bk, data, 5, 80, 0)

    def test_tarean_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x2D -> optype=1, opcode=0x0D (tAreaN), opx=45, sz=9; always reldelta=1
        data = b'\x2d' + b'\x00' * 8
        dump_formula(bk, data, 9, 80, 0)

    def test_tref3d_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x3A -> optype=1, opcode=0x1A (tRef3d), opx=58, sz=7
        # refx(2 bytes) + cell_addr(4 bytes for biff8)
        data = b'\x3a' + struct.pack('<H', 0) + struct.pack('<HH', 2, 3)
        dump_formula(bk, data, 7, 80, 0)

    def test_tarea3d_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x3B -> optype=1, opcode=0x1B (tArea3d), opx=59, sz=11
        data = b'\x3b' + struct.pack('<H', 0) + struct.pack('<HH', 2, 3) + struct.pack('<HH', 5, 7)
        dump_formula(bk, data, 11, 80, 0)

    def test_tnamex_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x39 -> optype=1, opcode=0x19 (tNameX), opx=57, sz=7
        data = b'\x39' + struct.pack('<HH', 1, 2) + b'\x00\x00'
        dump_formula(bk, data, 7, 80, 0)

    def test_error_opcode_sets_any_err(self):
        bk = make_dump_bk()
        # op=0x27 -> optype=1, opcode=0x07 (in error_opcodes), opx=39, sz=7
        data = b'\x27' + b'\x00' * 6
        dump_formula(bk, data, 7, 80, 0)

    def test_unknown_opcode_optype1(self):
        bk = make_dump_bk()
        # op=0x26 -> optype=1, opcode=0x06 (not handled), opx=38, sz=7 -> else branch
        data = b'\x26' + b'\x00' * 6
        dump_formula(bk, data, 7, 80, 0)

    def test_dud_size_optype1(self):
        # op=0x30 -> optype=1, opcode=0x10 (not handled), opx=48, sztab4[48]=-2 -> dud size
        bk = make_dump_bk()
        data = b'\x30'
        dump_formula(bk, data, 1, 80, 0)
        assert '**** Dud size' in bk.logfile.getvalue()

    def test_tlist_opcode_optype0(self):
        bk = make_dump_bk()
        # Two tRef3d push [coords] onto stack; then tList pops and concatenates
        tref3d = b'\x3a' + struct.pack('<H', 0) + struct.pack('<HH', 0, 0)
        data = tref3d + tref3d + b'\x10'
        dump_formula(bk, data, len(data), 80, 0)

    @pytest.mark.skip(reason="do_box_funcs expects objects with .coords attr; tRef3d pushes raw tuples")
    def test_trange_opcode_optype0(self):
        bk = make_dump_bk()
        # Two tRef3d push [coords]; then tRange (0x11) pops and applies do_box_funcs
        tref3d = b'\x3a' + struct.pack('<H', 0) + struct.pack('<HH', 0, 0)
        data = tref3d + tref3d + b'\x11'
        dump_formula(bk, data, len(data), 80, 0)

    @pytest.mark.skip(reason="do_box_funcs expects objects with .coords attr; tRef3d pushes raw tuples")
    def test_tisect_opcode_optype0(self):
        bk = make_dump_bk()
        # Two tRef3d push [coords]; then tIsect (0x0F) pops and applies do_box_funcs
        tref3d = b'\x3a' + struct.pack('<H', 0) + struct.pack('<HH', 0, 0)
        data = tref3d + tref3d + b'\x0f'
        dump_formula(bk, data, len(data), 80, 0)

    def test_blah_end_of_formula_output(self):
        bk = make_dump_bk()
        # opcode=0x03, sz=1, loop completes -> blah prints end-of-formula summary
        data = b'\x03'
        dump_formula(bk, data, 1, 80, 0, blah=1)
        output = bk.logfile.getvalue()
        assert 'End of formula' in output

    def test_blah_stack_unprocessed_warning(self):
        bk = make_dump_bk()
        # Two tRef3d (each pushes [coords]) without consuming -> stack len=2 at end
        tref3d = b'\x3a' + struct.pack('<H', 0) + struct.pack('<HH', 0, 0)
        data = tref3d + tref3d
        dump_formula(bk, data, len(data), 80, 0, blah=1)
        output = bk.logfile.getvalue()
        assert 'Stack has unprocessed args' in output

    def test_blah_tattr_logging(self):
        bk = make_dump_bk()
        # tAttr with blah=1 to cover the blah print inside tAttr handler
        data = b'\x19' + struct.pack('<BH', 0x01, 0)
        dump_formula(bk, data, 4, 80, 0, blah=1)
        output = bk.logfile.getvalue()
        assert 'subop' in output

    def test_blah_tstr_logging(self):
        bk = make_dump_bk()
        # tStr with blah=1 to cover the blah print inside tStr handler
        data = b'\x17\x02\x00hi'
        dump_formula(bk, data, 5, 80, 0, blah=1)
        output = bk.logfile.getvalue()
        assert 'sz=' in output


# ===========================================================================
# decompile_formula do_unaryop (lines 1400-1411)
# ===========================================================================

def make_decompile_bk():
    """Create a minimal book object for decompile_formula tests."""
    bk = SimpleNamespace()
    bk.logfile = io.StringIO()
    bk.biff_version = 80
    bk.encoding = 'ascii'
    return bk


class TestDecompileFormulaUnaryOp:
    """Tests for do_unaryop inside decompile_formula (tUplus, tUminus, tPercent)."""

    def test_unary_minus(self):
        """tUminus (0x13) negates the top-of-stack operand text."""
        bk = make_decompile_bk()
        # tInt 5 (3 bytes) followed by tUminus (1 byte)
        data = bytes([0x1e, 0x05, 0x00, 0x13])
        result = decompile_formula(bk, data, 4, FMLA_TYPE_CELL)
        assert result == '-5.0'

    def test_unary_plus(self):
        """tUplus (0x12) applies unary plus to the top-of-stack operand text."""
        bk = make_decompile_bk()
        # tInt 5 (3 bytes) followed by tUplus (1 byte)
        data = bytes([0x1e, 0x05, 0x00, 0x12])
        result = decompile_formula(bk, data, 4, FMLA_TYPE_CELL)
        assert result == '+5.0'

    def test_percent(self):
        """tPercent (0x14) appends '%' to the top-of-stack operand text."""
        bk = make_decompile_bk()
        # tInt 5 (3 bytes) followed by tPercent (1 byte)
        data = bytes([0x1e, 0x05, 0x00, 0x14])
        result = decompile_formula(bk, data, 4, FMLA_TYPE_CELL)
        assert result == '5.0%'


# ===========================================================================
# decompile_formula do_binop (lines 1383-1398)
# ===========================================================================

class TestDecompileFormulaBinOp:
    """Tests for do_binop inside decompile_formula."""

    def test_add(self):
        """tAdd (0x03) produces 'a+b' text with no parentheses for leaf operands."""
        bk = make_decompile_bk()
        # tInt 3, tInt 2, tAdd
        data = bytes([0x1e, 0x03, 0x00, 0x1e, 0x02, 0x00, 0x03])
        result = decompile_formula(bk, data, 7, FMLA_TYPE_CELL)
        assert result == '3.0+2.0'

    def test_sub(self):
        """tSub (0x04) produces 'a-b' text."""
        bk = make_decompile_bk()
        # tInt 5, tInt 3, tSub
        data = bytes([0x1e, 0x05, 0x00, 0x1e, 0x03, 0x00, 0x04])
        result = decompile_formula(bk, data, 7, FMLA_TYPE_CELL)
        assert result == '5.0-3.0'

    def test_mul(self):
        """tMul (0x05) produces 'a*b' text."""
        bk = make_decompile_bk()
        # tInt 4, tInt 5, tMul
        data = bytes([0x1e, 0x04, 0x00, 0x1e, 0x05, 0x00, 0x05])
        result = decompile_formula(bk, data, 7, FMLA_TYPE_CELL)
        assert result == '4.0*5.0'

    def test_div(self):
        """tDiv (0x06) produces 'a/b' text."""
        bk = make_decompile_bk()
        # tInt 10, tInt 2, tDiv
        data = bytes([0x1e, 0x0a, 0x00, 0x1e, 0x02, 0x00, 0x06])
        result = decompile_formula(bk, data, 7, FMLA_TYPE_CELL)
        assert result == '10.0/2.0'

    def test_parentheses_lower_rank_operand(self):
        """Operand with rank lower than operator rank gets wrapped in parentheses."""
        bk = make_decompile_bk()
        # (1+2)*3: tInt 1, tInt 2, tAdd (rank=30), tInt 3, tMul (rank=40)
        # After tAdd, result has rank=30; tMul has rank=40, so 30 < 40 → parentheses
        data = bytes([
            0x1e, 0x01, 0x00,  # tInt 1
            0x1e, 0x02, 0x00,  # tInt 2
            0x03,              # tAdd
            0x1e, 0x03, 0x00,  # tInt 3
            0x05,              # tMul
        ])
        result = decompile_formula(bk, data, 11, FMLA_TYPE_CELL)
        assert result == '(1.0+2.0)*3.0'

    def test_result_operand_has_correct_kind(self):
        """do_binop creates a result Operand with the correct result_kind (oNUM for tAdd)."""
        bk = make_decompile_bk()
        # tInt 2, tInt 3, tAdd — result should have text '2.0+3.0'
        data = bytes([0x1e, 0x02, 0x00, 0x1e, 0x03, 0x00, 0x03])
        result = decompile_formula(bk, data, 7, FMLA_TYPE_CELL)
        assert result == '2.0+3.0'

    def test_lt_comparison(self):
        """tLT (0x09) produces 'a<b' text with result_kind oBOOL."""
        bk = make_decompile_bk()
        # tInt 1, tInt 2, tLT
        data = bytes([0x1e, 0x01, 0x00, 0x1e, 0x02, 0x00, 0x09])
        result = decompile_formula(bk, data, 7, FMLA_TYPE_CELL)
        assert result == '1.0<2.0'
