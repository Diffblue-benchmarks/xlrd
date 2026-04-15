# -*- coding: utf-8 -*-
"""Unit tests for Sheet.handle_msodrawingetc."""

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


@pytest.fixture(autouse=True)
def restore_debug_flag():
    original = sheet_module.OBJ_MSO_DEBUG
    yield
    sheet_module.OBJ_MSO_DEBUG = original


def _make_record(fbt, ver, cb, body=b''):
    """Build an MSODrawing record header + body."""
    inst = 0
    tmp = (inst << 4) | (ver & 0xF)
    header = struct.pack('<HHI', tmp, fbt, cb)
    return header + body


def _make_container_record(fbt=0xF000):
    """Build a container record (ver=0xF, ndb=0)."""
    tmp = 0xF  # ver=0xF
    header = struct.pack('<HHI', tmp, fbt, 0)
    return header  # 8 bytes, ndb=0


def _make_client_anchor_record():
    """Build a Client Anchor record (fbt=0xF010), ndb=18."""
    fbt = 0xF010
    ver = 1
    cb = 18
    body = struct.pack('<Hiiii', 0, 0, 0, 1, 1)  # anchor_unk, colx_lo, rowx_lo, colx_hi, rowx_hi
    return _make_record(fbt, ver, cb, body)


def _make_client_data_record():
    """Build a Client Data record (fbt=0xF011), cb=0, must be last."""
    fbt = 0xF011
    ver = 1
    cb = 0
    return _make_record(fbt, ver, cb)


def _make_other_record():
    """Build a generic 'else' record (fbt != 0xF010 and != 0xF011)."""
    fbt = 0x0001
    ver = 1
    cb = 0
    return _make_record(fbt, ver, cb)


class TestHandleMsodrawingetcNoDebug:
    def test_returns_early_when_debug_disabled(self, sheet):
        sheet_module.OBJ_MSO_DEBUG = 0
        # Should return None immediately without processing
        data = _make_client_anchor_record()
        result = sheet.handle_msodrawingetc(0xEC, len(data), data)
        assert result is None

    def test_returns_early_when_debug_is_falsy(self, sheet):
        sheet_module.OBJ_MSO_DEBUG = False
        data = _make_client_anchor_record()
        result = sheet.handle_msodrawingetc(0xEC, len(data), data)
        assert result is None


class TestHandleMsodrawingetcBiffVersion:
    def test_biff_lt80_returns_early(self):
        sheet_module.OBJ_MSO_DEBUG = 1
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        data = _make_client_anchor_record()
        result = sh.handle_msodrawingetc(0xEC, len(data), data)
        assert result is None


class TestHandleMsodrawingetcContainerRecord:
    def test_container_record_parsed(self, sheet):
        sheet_module.OBJ_MSO_DEBUG = 1
        data = _make_container_record()
        # Should complete without error (pos advances by 0+8=8)
        sheet.handle_msodrawingetc(0xEC, len(data), data)

    def test_empty_data_completes(self, sheet):
        sheet_module.OBJ_MSO_DEBUG = 1
        data = b''
        sheet.handle_msodrawingetc(0xEC, 0, data)


class TestHandleMsodrawingetcClientAnchor:
    def test_client_anchor_parsed(self, sheet):
        sheet_module.OBJ_MSO_DEBUG = 1
        data = _make_client_anchor_record()
        # Should not raise; client anchor sets anchor fields on MSODrawing object
        sheet.handle_msodrawingetc(0xF010, len(data), data)

    def test_client_anchor_followed_by_client_data(self, sheet):
        sheet_module.OBJ_MSO_DEBUG = 1
        anchor = _make_client_anchor_record()
        client_data = _make_client_data_record()
        data = anchor + client_data
        sheet.handle_msodrawingetc(0xEC, len(data), data)


class TestHandleMsodrawingetcClientData:
    def test_client_data_at_end(self, sheet):
        sheet_module.OBJ_MSO_DEBUG = 1
        data = _make_client_data_record()
        sheet.handle_msodrawingetc(0xEC, len(data), data)


class TestHandleMsodrawingetcOtherRecord:
    def test_other_fbt_takes_else_branch(self, sheet):
        sheet_module.OBJ_MSO_DEBUG = 1
        data = _make_other_record()
        sheet.handle_msodrawingetc(0xEC, len(data), data)

    def test_multiple_other_records(self, sheet):
        sheet_module.OBJ_MSO_DEBUG = 1
        data = _make_other_record() + _make_other_record()
        sheet.handle_msodrawingetc(0xEC, len(data), data)
