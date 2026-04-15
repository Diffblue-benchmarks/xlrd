# -*- coding: utf-8 -*-
"""Unit tests for Sheet.handle_obj."""

import sys
import struct
import pytest

import xlrd.sheet as sheet_module
from xlrd.sheet import Sheet
from xlrd.biffh import XLRDError


class MockBook:
    biff_version = 80
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    encoding = 'latin-1'
    _xf_index_to_xl_type_map = {}
    _sheet_visibility = [0, 0, 0, 0, 0]

    def get_record_parts(self):
        raise NotImplementedError("override in tests")

    def derive_encoding(self):
        return 'latin-1'


@pytest.fixture
def sheet():
    return Sheet(MockBook(), 0, 'TestSheet', 0)


def _make_ftcmo(obj_type=1, obj_id=1, option_flags=0):
    """Build a 22-byte ftCmo sub-record (ft=0x15, cb=18)."""
    header = struct.pack('<HH', 0x15, 18)
    body = struct.pack('<HHH', obj_type, obj_id, option_flags) + b'\x00' * 12
    return header + body  # 4 + 18 = 22 bytes


def _make_ft00_zeros(n=8):
    """Return all-zero bytes that satisfy the 'optional reserved' end-of-record check.

    The code checks data[pos:data_len] == b'\\0' * (data_len - pos), so every
    byte from pos onward (including the pseudo-header) must be zero.
    """
    return b'\x00' * n


def _make_ft0c_scrollbar(value=50, minimum=0, maximum=100, inc=1, page=10):
    """Build a ft=0x0C (Scrollbar) sub-record."""
    header = struct.pack('<HH', 0x0C, 14)
    body = b'\x00' * 4 + struct.pack('<5H', value, minimum, maximum, inc, page)
    return header + body  # 4 + 14 = 18 bytes


def _make_ft0d():
    """Build a ft=0x0D (notes structure) sub-record."""
    return struct.pack('<HH', 0x0D, 0)


def _make_ft13(n=0):
    """Build a ft=0x13 (list box) sub-record."""
    return struct.pack('<HH', 0x13, n) + b'\x00' * n


class TestHandleObjBiffVersion:
    def test_biff_lt80_returns_none(self):
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        result = sh.handle_obj(b'\x00' * 22)
        assert result is None

    def test_biff_ge80_does_not_return_none_for_valid_data(self, sheet):
        data = _make_ftcmo()
        result = sheet.handle_obj(data)
        assert result is not None


class TestHandleObjFirstRecord:
    def test_non_ftcmo_first_returns_none(self, sheet):
        # ft != 0x15 at pos=0 => return None
        bad_header = struct.pack('<HH', 0x01, 18) + b'\x00' * 18
        result = sheet.handle_obj(bad_header)
        assert result is None

    def test_wrong_cb_first_returns_none(self, sheet):
        # ft=0x15 but cb != 18 => return None
        bad_header = struct.pack('<HH', 0x15, 10) + b'\x00' * 10
        result = sheet.handle_obj(bad_header)
        assert result is None

    def test_non_ftcmo_first_with_verbosity_returns_none(self):
        book = MockBook()
        book.verbosity = 1
        sh = Sheet(book, 0, 'TestSheet', 0)
        bad_header = struct.pack('<HH', 0x01, 18) + b'\x00' * 18
        result = sh.handle_obj(bad_header)
        assert result is None

    def test_valid_ftcmo_returns_msobj(self, sheet):
        data = _make_ftcmo(obj_type=25, obj_id=7)
        result = sheet.handle_obj(data)
        assert result is not None
        assert result.type == 25
        assert result.id == 7

    def test_ftcmo_option_flags_locked(self, sheet):
        # bit 0 = locked
        data = _make_ftcmo(option_flags=0x0001)
        result = sheet.handle_obj(data)
        assert result.locked == 1

    def test_ftcmo_option_flags_printable(self, sheet):
        data = _make_ftcmo(option_flags=0x0010)
        result = sheet.handle_obj(data)
        assert result.printable == 1

    def test_ftcmo_option_flags_autofilter(self, sheet):
        data = _make_ftcmo(option_flags=0x0100)
        result = sheet.handle_obj(data)
        assert result.autofilter == 1

    def test_ftcmo_option_flags_scrollbar_flag(self, sheet):
        data = _make_ftcmo(option_flags=0x0200)
        result = sheet.handle_obj(data)
        assert result.scrollbar_flag == 1

    def test_ftcmo_option_flags_autofill(self, sheet):
        data = _make_ftcmo(option_flags=0x2000)
        result = sheet.handle_obj(data)
        assert result.autofill == 1

    def test_ftcmo_option_flags_autoline(self, sheet):
        data = _make_ftcmo(option_flags=0x4000)
        result = sheet.handle_obj(data)
        assert result.autoline == 1


class TestHandleObjFt00:
    def test_ft00_all_zeros_breaks_and_returns_msobj(self, sheet):
        # ft=0x00, all zeros => break out, return object
        data = _make_ftcmo() + _make_ft00_zeros()
        result = sheet.handle_obj(data)
        assert result is not None

    def test_ft00_non_zero_data_raises_xlrd_error(self, sheet):
        # ft=0x00 but non-zero data => XLRDError
        ft00_header = struct.pack('<HH', 0x00, 4)
        ft00_body = b'\x01\x02\x03\x04'
        data = _make_ftcmo() + ft00_header + ft00_body
        with pytest.raises(XLRDError):
            sheet.handle_obj(data)


class TestHandleObjFt0C:
    def test_scrollbar_values_parsed(self, sheet):
        data = _make_ftcmo() + _make_ft0c_scrollbar(
            value=50, minimum=0, maximum=100, inc=1, page=10
        )
        result = sheet.handle_obj(data)
        assert result.scrollbar_value == 50
        assert result.scrollbar_min == 0
        assert result.scrollbar_max == 100
        assert result.scrollbar_inc == 1
        assert result.scrollbar_page == 10

    def test_scrollbar_custom_values(self, sheet):
        data = _make_ftcmo() + _make_ft0c_scrollbar(
            value=25, minimum=5, maximum=200, inc=2, page=20
        )
        result = sheet.handle_obj(data)
        assert result.scrollbar_value == 25
        assert result.scrollbar_min == 5
        assert result.scrollbar_max == 200
        assert result.scrollbar_inc == 2
        assert result.scrollbar_page == 20


class TestHandleObjFt0D:
    def test_ft0d_notes_does_not_raise(self, sheet):
        data = _make_ftcmo() + _make_ft0d()
        result = sheet.handle_obj(data)
        assert result is not None


class TestHandleObjFt13:
    def test_ft13_with_autofilter_breaks(self, sheet):
        # autofilter=1 (option_flags=0x0100) + ft=0x13 => break
        data = _make_ftcmo(option_flags=0x0100) + _make_ft13()
        result = sheet.handle_obj(data)
        assert result is not None

    def test_ft13_without_autofilter_continues(self, sheet):
        # autofilter=0 => continues past ft=0x13 record
        data = _make_ftcmo(option_flags=0x0000) + _make_ft13()
        result = sheet.handle_obj(data)
        assert result is not None


class TestHandleObjDebugMode:
    def test_debug_mode_does_not_raise(self, sheet):
        original = sheet_module.OBJ_MSO_DEBUG
        try:
            sheet_module.OBJ_MSO_DEBUG = 1
            data = _make_ftcmo()
            result = sheet.handle_obj(data)
            assert result is not None
        finally:
            sheet_module.OBJ_MSO_DEBUG = original
