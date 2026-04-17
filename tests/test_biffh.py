# -*- coding: utf-8 -*-
"""
Tests for xlrd.biffh — low-level BIFF helper utilities.
"""
import io
import sys
import pytest

from xlrd.biffh import (
    XLRDError,
    BaseObject,
    biff_rec_name_dict,
    biff_text_from_num,
    error_text_from_code,
    is_cell_opcode,
    upkbits,
    upkbitsL,
    XL_CELL_EMPTY,
    XL_CELL_TEXT,
    XL_CELL_NUMBER,
    XL_CELL_DATE,
    XL_CELL_BOOLEAN,
    XL_CELL_ERROR,
    XL_CELL_BLANK,
    XL_BOOLERR,
    XL_FORMULA,
    XL_LABEL,
    XL_LABELSST,
    XL_MULRK,
    XL_NUMBER,
    XL_RK,
)


class TestCellTypeConstants:

    def test_cell_type_values_are_sequential(self):
        assert XL_CELL_EMPTY == 0
        assert XL_CELL_TEXT == 1
        assert XL_CELL_NUMBER == 2
        assert XL_CELL_DATE == 3
        assert XL_CELL_BOOLEAN == 4
        assert XL_CELL_ERROR == 5
        assert XL_CELL_BLANK == 6


class TestIsCellOpcode:

    def test_formula_opcode_is_cell(self):
        assert is_cell_opcode(XL_FORMULA) is True

    def test_boolerr_opcode_is_cell(self):
        assert is_cell_opcode(XL_BOOLERR) is True

    def test_label_opcode_is_cell(self):
        assert is_cell_opcode(XL_LABEL) is True

    def test_labelsst_opcode_is_cell(self):
        assert is_cell_opcode(XL_LABELSST) is True

    def test_mulrk_opcode_is_cell(self):
        assert is_cell_opcode(XL_MULRK) is True

    def test_number_opcode_is_cell(self):
        assert is_cell_opcode(XL_NUMBER) is True

    def test_rk_opcode_is_cell(self):
        assert is_cell_opcode(XL_RK) is True

    def test_unknown_opcode_is_not_cell(self):
        assert is_cell_opcode(0x9999) is False

    def test_zero_opcode_is_not_cell(self):
        assert is_cell_opcode(0x0000) is False

    def test_eof_opcode_is_not_cell(self):
        assert is_cell_opcode(0x000A) is False  # XL_EOF


class TestUpkbits:

    def test_single_bit_extraction(self):
        obj = type('Obj', (), {})()
        # Extract bit 0 (mask=1) into attr 'flag'
        upkbits(obj, 0b00000001, [(0, 0x01, 'flag')])
        assert obj.flag == 1

    def test_single_bit_extraction_zero(self):
        obj = type('Obj', (), {})()
        upkbits(obj, 0b00000000, [(0, 0x01, 'flag')])
        assert obj.flag == 0

    def test_multiple_fields(self):
        obj = type('Obj', (), {})()
        src = 0b10110010
        upkbits(obj, src, [
            (1, 0x02, 'bit1'),    # bit 1
            (4, 0x30, 'bits4_5'),  # bits 4-5
            (7, 0x80, 'bit7'),    # bit 7
        ])
        assert obj.bit1 == 1
        assert obj.bits4_5 == 0b11
        assert obj.bit7 == 1

    def test_nibble_extraction(self):
        obj = type('Obj', (), {})()
        upkbits(obj, 0xAB, [(4, 0xF0, 'high_nibble'), (0, 0x0F, 'low_nibble')])
        assert obj.high_nibble == 0xA
        assert obj.low_nibble == 0xB


class TestUpkbitsL:

    def test_returns_int_values(self):
        obj = type('Obj', (), {})()
        upkbitsL(obj, 0xFF, [(0, 0x0F, 'nibble')])
        assert obj.nibble == 15
        assert isinstance(obj.nibble, int)


class TestErrorTextFromCode:

    def test_null_error(self):
        assert error_text_from_code[0x00] == '#NULL!'

    def test_div0_error(self):
        assert error_text_from_code[0x07] == '#DIV/0!'

    def test_value_error(self):
        assert error_text_from_code[0x0F] == '#VALUE!'

    def test_ref_error(self):
        assert error_text_from_code[0x17] == '#REF!'

    def test_name_error(self):
        assert error_text_from_code[0x1D] == '#NAME?'

    def test_num_error(self):
        assert error_text_from_code[0x24] == '#NUM!'

    def test_na_error(self):
        assert error_text_from_code[0x2A] == '#N/A'


class TestBiffTextFromNum:

    def test_biff8(self):
        assert biff_text_from_num[80] == "8"

    def test_biff5(self):
        assert biff_text_from_num[50] == "5"

    def test_not_biff(self):
        assert biff_text_from_num[0] == "(not BIFF)"


class TestBiffRecNameDict:

    def test_eof_record(self):
        assert biff_rec_name_dict[0x000A] == "EOF"

    def test_bof_record(self):
        assert biff_rec_name_dict[0x0809] == "BOF"

    def test_sst_record(self):
        assert biff_rec_name_dict[0x00FC] == "SST"


class TestBaseObject:

    def test_dump_writes_attributes(self):
        obj = BaseObject()
        obj.x = 42
        obj.name = "test"
        buf = io.StringIO()
        obj.dump(f=buf)
        output = buf.getvalue()
        assert "x" in output
        assert "name" in output

    def test_dump_with_header_and_footer(self):
        obj = BaseObject()
        buf = io.StringIO()
        obj.dump(f=buf, header="=== HEADER ===", footer="=== FOOTER ===")
        output = buf.getvalue()
        assert "=== HEADER ===" in output
        assert "=== FOOTER ===" in output

    def test_xlrd_error_is_exception(self):
        with pytest.raises(XLRDError):
            raise XLRDError("test error")
