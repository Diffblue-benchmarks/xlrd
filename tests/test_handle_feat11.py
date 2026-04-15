# -*- coding: utf-8 -*-
"""Unit tests for Sheet.handle_feat11."""

import sys
import struct
import pytest

import xlrd.sheet as sheet_module
from xlrd.sheet import Sheet


class MockBook:
    biff_version = 80
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    _xf_index_to_xl_type_map = {}
    _sheet_visibility = [0, 0, 0, 0, 0]


@pytest.fixture
def sheet():
    return Sheet(MockBook(), 0, 'TestSheet', 0)


def _make_feat11_data():
    """Construct valid FEAT11 binary data (101 bytes)."""
    # First 35 bytes: rt, grbitFrt, Ref0, isf, fHdr, reserved0, cref, cbFeatData, reserved1, Ref1
    # Format: '<HH8sHBiHiH8s'
    ref0 = b'\x00' * 8
    part1 = struct.pack(
        '<HH8sHBiHiH8s',
        0x0872,  # rt
        0,       # grbitFrt
        ref0,    # Ref0
        5,       # isf
        0,       # fHdr
        0,       # reserved0
        1,       # cref
        66,      # cbFeatData
        0,       # reserved1
        ref0,    # Ref1 (must equal Ref0)
    )
    assert len(part1) == 35

    # Next 66 bytes: lt..cchName
    # Format: '<iiiiiiHHiiiii16sH'
    part2 = struct.pack(
        '<iiiiiiHHiiiii16sH',
        0,              # lt
        1,              # idList
        1,              # crwHeader
        0,              # crwTotals
        2,              # idFieldNext
        66,             # cbFSData
        0x0B00,         # rupBuild
        0,              # unusedShort
        0,              # listFlags
        0,              # lPosStmCache
        0,              # cbStmCache
        0,              # cchStmCache
        0,              # lem
        b'\x00' * 16,  # rgbHashParam
        0,              # cchName
    )
    assert len(part2) == 66

    return part1 + part2


class TestHandleFeat11:
    def test_early_return_when_debug_disabled(self, sheet):
        # OBJ_MSO_DEBUG is 0 by default; handle_feat11 should return immediately
        result = sheet.handle_feat11(b'\x00' * 101)
        assert result is None

    def test_debug_enabled_processes_data(self, sheet):
        # Enable debug mode to exercise the parsing branches
        original = sheet_module.OBJ_MSO_DEBUG
        try:
            sheet_module.OBJ_MSO_DEBUG = 1
            data = _make_feat11_data()
            # Should not raise; all assertions inside should pass
            sheet.handle_feat11(data)
        finally:
            sheet_module.OBJ_MSO_DEBUG = original

    def test_debug_enabled_reserved0_assertion(self, sheet):
        # reserved0 must be 0; a non-zero value should cause AssertionError
        original = sheet_module.OBJ_MSO_DEBUG
        try:
            sheet_module.OBJ_MSO_DEBUG = 1
            data = bytearray(_make_feat11_data())
            # reserved0 is at offset 12 (H+H+8s+H+B = 2+2+8+2+1=15... let me recalculate)
            # '<HH8sHBiHiH8s': H(2)+H(2)+8s(8)+H(2)+B(1)+i(4)+H(2)+i(4)+H(2)+8s(8) = 35
            # reserved0 is the 'i' at offset 2+2+8+2+1=15, 4 bytes
            offset = 2 + 2 + 8 + 2 + 1  # =15
            struct.pack_into('<i', data, offset, 1)  # reserved0 = 1 (invalid)
            with pytest.raises(AssertionError):
                sheet.handle_feat11(bytes(data))
        finally:
            sheet_module.OBJ_MSO_DEBUG = original
