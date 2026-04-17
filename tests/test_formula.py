# -*- coding: utf-8 -*-
"""
Tests for xlrd.formula — Operand, Ref3D, num2strg,
adjust_cell_addr_biff8, adjust_cell_addr_biff_le7, get_cell_addr,
get_cell_range_addr, do_box_funcs, and arithmetic helpers.
"""
import struct
import pytest

from xlrd.formula import (
    Operand,
    Ref3D,
    num2strg,
    adjust_cell_addr_biff8,
    adjust_cell_addr_biff_le7,
    get_cell_addr,
    get_cell_range_addr,
    do_box_funcs,
    tRangeFuncs,
    tIsectFuncs,
    nop,
    _opr_pow,
    _opr_lt,
    _opr_le,
    _opr_eq,
    _opr_ge,
    _opr_gt,
    _opr_ne,
    oUNK,
    oNUM,
    oSTRG,
    oBOOL,
    oREF,
    oREL,
    okind_dict,
    colname,
)


# ---------------------------------------------------------------------------
# Operand
# ---------------------------------------------------------------------------

class TestOperand:

    def test_default_kind_is_oUNK(self):
        op = Operand()
        assert op.kind == oUNK

    def test_default_text_is_question(self):
        op = Operand()
        assert op.text == '?'

    def test_init_with_kind_and_value(self):
        op = Operand(akind=oNUM, avalue=3.14)
        assert op.kind == oNUM
        assert op.value == 3.14

    def test_init_with_text(self):
        op = Operand(atext='A1+B1')
        assert op.text == 'A1+B1'

    def test_repr_contains_kind(self):
        op = Operand(akind=oNUM, avalue=42.0, atext='42')
        r = repr(op)
        assert 'oNUM' in r or 'NUM' in r.upper()

    def test_repr_contains_value(self):
        op = Operand(akind=oNUM, avalue=42.0)
        assert '42' in repr(op)

    def test_rank_default_zero(self):
        op = Operand()
        assert op.rank == 0

    def test_rank_set_via_init(self):
        op = Operand(arank=10)
        assert op.rank == 10

    def test_kind_string_operand(self):
        op = Operand(akind=oSTRG, avalue='hello')
        assert op.kind == oSTRG

    def test_kind_bool_operand(self):
        op = Operand(akind=oBOOL, avalue=True)
        assert op.kind == oBOOL


# ---------------------------------------------------------------------------
# Ref3D
# ---------------------------------------------------------------------------

class TestRef3D:

    def test_basic_coords(self):
        ref = Ref3D((0, 1, 2, 5, 3, 7))
        assert ref.shtxlo == 0
        assert ref.shtxhi == 1
        assert ref.rowxlo == 2
        assert ref.rowxhi == 5
        assert ref.colxlo == 3
        assert ref.colxhi == 7

    def test_coords_attribute(self):
        ref = Ref3D((1, 2, 3, 4, 5, 6))
        assert ref.coords == (1, 2, 3, 4, 5, 6)

    def test_default_relflags_all_zero(self):
        ref = Ref3D((0, 1, 0, 1, 0, 1))
        assert ref.relflags == (0, 0, 0, 0, 0, 0)

    def test_relflags_from_tuple(self):
        ref = Ref3D((0, 1, 0, 1, 0, 1, 1, 1, 0, 0, 0, 0))
        assert ref.relflags == (1, 1, 0, 0, 0, 0)

    def test_repr_absolute(self):
        ref = Ref3D((0, 1, 0, 1, 0, 1))
        r = repr(ref)
        assert 'Ref3D' in r
        assert 'coords' in r

    def test_repr_relative(self):
        ref = Ref3D((0, 1, 0, 1, 0, 1, 1, 0, 0, 0, 0, 0))
        r = repr(ref)
        assert 'relflags' in r

    def test_is_tuple_subclass(self):
        ref = Ref3D((0, 1, 0, 1, 0, 1))
        assert isinstance(ref, tuple)


# ---------------------------------------------------------------------------
# num2strg
# ---------------------------------------------------------------------------

class TestNum2strg:

    def test_integer_value_no_decimal(self):
        assert num2strg(5.0) == '5'

    def test_float_value_keeps_decimal(self):
        assert num2strg(3.14) == '3.14'

    def test_large_integer(self):
        assert num2strg(1000.0) == '1000'

    def test_negative_integer(self):
        assert num2strg(-7.0) == '-7'

    def test_negative_float(self):
        result = num2strg(-2.5)
        assert result == '-2.5'

    def test_zero(self):
        assert num2strg(0.0) == '0'


# ---------------------------------------------------------------------------
# adjust_cell_addr_biff8
# ---------------------------------------------------------------------------

class TestAdjustCellAddrBiff8:

    def test_absolute_row_and_col(self):
        # No rel flags → absolute row=5, col=3
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(5, 3, reldelta=0)
        assert rowx == 5
        assert colx == 3
        assert row_rel == 0
        assert col_rel == 0

    def test_relative_flags_extracted(self):
        # bit 15 = row_rel, bit 14 = col_rel
        colval = (1 << 15) | (1 << 14) | 4  # row_rel=1, col_rel=1, colx=4
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(10, colval, reldelta=1)
        assert row_rel == 1
        assert col_rel == 1
        assert colx == 4

    def test_relative_row_delta_positive(self):
        colval = (1 << 15) | 2  # row_rel=1, col_rel=0, colx=2
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(10, colval, reldelta=1)
        assert rowx == 10  # positive row, stays the same
        assert row_rel == 1

    def test_relative_row_wraps_negative(self):
        # row_rel=1 and rowx >= 32768 → rowx -= 65536
        colval = (1 << 15) | 0  # row_rel=1
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(32768, colval, reldelta=1)
        assert rowx == 32768 - 65536

    def test_relative_col_wraps_negative(self):
        # col_rel=1 and colx >= 128 → colx -= 256
        colval = (1 << 15) | (1 << 14) | 200  # row_rel=1, col_rel=1, colx=200
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(0, colval, reldelta=1)
        assert colx == 200 - 256

    def test_non_reldelta_row_relative(self):
        # reldelta=0, row_rel=1 → rowx -= browx
        colval = (1 << 15) | 0  # row_rel=1
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(
            10, colval, reldelta=0, browx=3, bcolx=0
        )
        assert rowx == 10 - 3

    def test_non_reldelta_col_relative(self):
        # reldelta=0, col_rel=1 → colx -= bcolx
        colval = (1 << 14) | 5  # col_rel=1, colx=5
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff8(
            0, colval, reldelta=0, browx=0, bcolx=2
        )
        assert colx == 5 - 2


# ---------------------------------------------------------------------------
# adjust_cell_addr_biff_le7
# ---------------------------------------------------------------------------

class TestAdjustCellAddrBiffLe7:

    def test_absolute_address(self):
        rowval = 5  # no rel flags
        colval = 3
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(
            rowval, colval, reldelta=0
        )
        assert rowx == 5
        assert colx == 3
        assert row_rel == 0
        assert col_rel == 0

    def test_relative_flags_in_rowval(self):
        # row_rel in bit 15, col_rel in bit 14
        rowval = (1 << 15) | (1 << 14) | 10  # row_rel=1, col_rel=1, rowx=10
        colval = 5
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(
            rowval, colval, reldelta=1
        )
        assert row_rel == 1
        assert col_rel == 1
        assert rowx == 10

    def test_relative_row_wraps(self):
        rowval = (1 << 15) | 8192  # row_rel=1, rowx=8192 → wraps
        colval = 0
        rowx, colx, row_rel, col_rel = adjust_cell_addr_biff_le7(
            rowval, colval, reldelta=1
        )
        assert rowx == 8192 - 16384


# ---------------------------------------------------------------------------
# get_cell_addr
# ---------------------------------------------------------------------------

class TestGetCellAddr:

    def test_biff8_absolute(self):
        # bv=80: rowval=5, colval=3 (no rel flags)
        data = struct.pack('<HH', 5, 3)
        rowx, colx, row_rel, col_rel = get_cell_addr(data, 0, bv=80, reldelta=0)
        assert rowx == 5
        assert colx == 3

    def test_biff_le7_absolute(self):
        # bv=50: rowval stored as <H, colval as <B
        data = struct.pack('<HB', 5, 3)
        rowx, colx, row_rel, col_rel = get_cell_addr(data, 0, bv=50, reldelta=0)
        assert rowx == 5
        assert colx == 3


# ---------------------------------------------------------------------------
# get_cell_range_addr
# ---------------------------------------------------------------------------

class TestGetCellRangeAddr:

    def test_biff8_range(self):
        # row1=0,row2=3,col1=0,col2=2 (no rel flags)
        data = struct.pack('<HHHH', 0, 3, 0, 2)
        res1, res2 = get_cell_range_addr(data, 0, bv=80, reldelta=0)
        assert res1[0] == 0  # rowx1
        assert res2[0] == 3  # rowx2
        assert res1[1] == 0  # colx1
        assert res2[1] == 2  # colx2

    def test_biff_le7_range(self):
        data = struct.pack('<HHBB', 0, 3, 0, 2)
        res1, res2 = get_cell_range_addr(data, 0, bv=50, reldelta=0)
        assert res1[0] == 0
        assert res2[0] == 3


# ---------------------------------------------------------------------------
# do_box_funcs
# ---------------------------------------------------------------------------

class TestDoBoxFuncs:

    def test_range_funcs_gives_bounding_box(self):
        # tRangeFuncs = (min, max, min, max, min, max)
        class Box:
            pass
        boxa = Box()
        boxa.coords = (0, 2, 1, 5, 0, 3)
        boxb = Box()
        boxb.coords = (1, 3, 0, 4, 2, 6)
        result = do_box_funcs(tRangeFuncs, boxa, boxb)
        assert result == (min(0, 1), max(2, 3), min(1, 0), max(5, 4), min(0, 2), max(3, 6))

    def test_isect_funcs_gives_intersection(self):
        # tIsectFuncs = (max, min, max, min, max, min)
        class Box:
            pass
        boxa = Box()
        boxa.coords = (0, 4, 1, 5, 0, 3)
        boxb = Box()
        boxb.coords = (1, 3, 0, 4, 2, 6)
        result = do_box_funcs(tIsectFuncs, boxa, boxb)
        assert result == (max(0, 1), min(4, 3), max(1, 0), min(5, 4), max(0, 2), min(3, 6))


# ---------------------------------------------------------------------------
# Arithmetic helper functions
# ---------------------------------------------------------------------------

class TestArithmeticHelpers:

    def test_nop_returns_value(self):
        assert nop(42) == 42
        assert nop('hello') == 'hello'

    def test_opr_pow(self):
        assert _opr_pow(2, 10) == 1024

    def test_opr_lt(self):
        assert _opr_lt(1, 2) is True
        assert _opr_lt(2, 1) is False

    def test_opr_le(self):
        assert _opr_le(2, 2) is True
        assert _opr_le(3, 2) is False

    def test_opr_eq(self):
        assert _opr_eq(5, 5) is True
        assert _opr_eq(5, 6) is False

    def test_opr_ge(self):
        assert _opr_ge(3, 3) is True
        assert _opr_ge(2, 3) is False

    def test_opr_gt(self):
        assert _opr_gt(4, 3) is True
        assert _opr_gt(3, 4) is False

    def test_opr_ne(self):
        assert _opr_ne(1, 2) is True
        assert _opr_ne(1, 1) is False


# ---------------------------------------------------------------------------
# formula.colname (separate from book.colname)
# ---------------------------------------------------------------------------

class TestFormulaColname:

    def test_first_col(self):
        assert colname(0) == 'A'

    def test_last_single(self):
        assert colname(25) == 'Z'

    def test_first_double(self):
        assert colname(26) == 'AA'


# ---------------------------------------------------------------------------
# okind_dict
# ---------------------------------------------------------------------------

class TestOkindDict:

    def test_oNUM_in_dict(self):
        assert oNUM in okind_dict

    def test_oSTRG_in_dict(self):
        assert oSTRG in okind_dict

    def test_oBOOL_in_dict(self):
        assert oBOOL in okind_dict

    def test_oREF_in_dict(self):
        assert oREF in okind_dict
