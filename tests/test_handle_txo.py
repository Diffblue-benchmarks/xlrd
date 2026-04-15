# -*- coding: utf-8 -*-
"""Unit tests for Sheet.handle_txo."""

import sys
import struct
import pytest

import xlrd.sheet as sheet_module
from xlrd.sheet import Sheet
from xlrd.biffh import XL_CONTINUE


class MockBook:
    biff_version = 80
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    _xf_index_to_xl_type_map = {}
    _sheet_visibility = [0, 0, 0, 0, 0]

    def get_record_parts(self):
        raise NotImplementedError("override in tests")


@pytest.fixture
def sheet():
    return Sheet(MockBook(), 0, 'TestSheet', 0)


def _make_txo_data(option_flags=0, rot=0, cchText=0, cbRuns=0, ifntEmpty=0):
    """Pack a TXO binary header."""
    fmt = '<HH6sHHH'
    controlInfo = b'\x00' * 6
    data = struct.pack(fmt, option_flags, rot, controlInfo, cchText, cbRuns, ifntEmpty)
    return data


class TestHandleTxoBiffVersion:
    def test_returns_none_for_biff_version_less_than_80(self):
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        result = sh.handle_txo(_make_txo_data())
        assert result is None

    def test_returns_none_for_biff_version_50(self):
        book = MockBook()
        book.biff_version = 50
        sh = Sheet(book, 0, 'TestSheet', 0)
        result = sh.handle_txo(_make_txo_data())
        assert result is None


class TestHandleTxoEmptyText:
    def test_empty_text_no_runs_returns_mstxo(self, sheet):
        data = _make_txo_data(cchText=0, cbRuns=0)
        result = sheet.handle_txo(data)
        assert result is not None
        assert result.text == ''
        assert result.rich_text_runlist == []

    def test_option_flags_decoded(self, sheet):
        # lock_text in bit 9 (mask 0x0200), value 1 when bit 9 is set
        option_flags = 0x0200  # lock_text = 1
        data = _make_txo_data(option_flags=option_flags, cchText=0, cbRuns=0)
        result = sheet.handle_txo(data)
        assert result.lock_text == 1

    def test_rot_stored(self, sheet):
        data = _make_txo_data(rot=90, cchText=0, cbRuns=0)
        result = sheet.handle_txo(data)
        assert result.rot == 90

    def test_ifntEmpty_stored(self, sheet):
        data = _make_txo_data(ifntEmpty=5, cchText=0, cbRuns=0)
        result = sheet.handle_txo(data)
        assert result.ifntEmpty == 5

    def test_fmla_is_remainder_of_data(self, sheet):
        extra = b'\xAB\xCD'
        data = _make_txo_data(cchText=0, cbRuns=0) + extra
        result = sheet.handle_txo(data)
        assert result.fmla == extra

    def test_fmla_empty_when_no_extra(self, sheet):
        data = _make_txo_data(cchText=0, cbRuns=0)
        result = sheet.handle_txo(data)
        assert result.fmla == b''


class TestHandleTxoWithText:
    def test_latin1_text_read_from_continue_record(self, sheet, mocker):
        text = b'hello'
        # data2: options byte (0=latin1) followed by text bytes
        data2 = bytes([0]) + text
        mocker.patch.object(
            sheet.book, 'get_record_parts',
            return_value=(XL_CONTINUE, len(data2), data2)
        )
        data = _make_txo_data(cchText=5, cbRuns=0)
        result = sheet.handle_txo(data)
        assert result.text == 'hello'
        assert result.rich_text_runlist == []

    def test_multiple_continue_records_for_text(self, sheet, mocker):
        chunk1 = bytes([0]) + b'ab'
        chunk2 = bytes([0]) + b'cd'
        mocker.patch.object(
            sheet.book, 'get_record_parts',
            side_effect=[
                (XL_CONTINUE, len(chunk1), chunk1),
                (XL_CONTINUE, len(chunk2), chunk2),
            ]
        )
        data = _make_txo_data(cchText=4, cbRuns=0)
        result = sheet.handle_txo(data)
        assert result.text == 'abcd'


class TestHandleTxoWithRichText:
    def test_rich_text_runs_parsed(self, sheet, mocker):
        # cbRuns=8 means one run (8 bytes per run); offset 1 != cchText so not removed
        run_data = struct.pack('<HH4x', 1, 2)  # run: (char_offset=1, font_index=2)
        assert len(run_data) == 8
        data3 = run_data
        mocker.patch.object(
            sheet.book, 'get_record_parts',
            return_value=(XL_CONTINUE, len(data3), data3)
        )
        data = _make_txo_data(cchText=0, cbRuns=8)
        result = sheet.handle_txo(data)
        assert len(result.rich_text_runlist) == 1
        assert result.rich_text_runlist[0] == (1, 2)

    def test_trailing_run_pointing_to_end_removed(self, sheet, mocker):
        # cchText=3; a run at offset 3 should be removed as it points to end of string
        run1 = struct.pack('<HH4x', 0, 1)  # valid run at offset 0
        run2 = struct.pack('<HH4x', 3, 0)  # trailing run pointing to cchText=3
        data3 = run1 + run2
        text_data2 = bytes([0]) + b'abc'
        mocker.patch.object(
            sheet.book, 'get_record_parts',
            side_effect=[
                (XL_CONTINUE, len(text_data2), text_data2),
                (XL_CONTINUE, len(data3), data3),
            ]
        )
        data = _make_txo_data(cchText=3, cbRuns=16)
        result = sheet.handle_txo(data)
        assert len(result.rich_text_runlist) == 1
        assert result.rich_text_runlist[0] == (0, 1)

    def test_all_trailing_runs_removed(self, sheet, mocker):
        # All runs point to cchText, all should be removed
        run = struct.pack('<HH4x', 0, 0)  # offset 0 == cchText=0 => trailing
        data3 = run
        mocker.patch.object(
            sheet.book, 'get_record_parts',
            return_value=(XL_CONTINUE, len(data3), data3)
        )
        data = _make_txo_data(cchText=0, cbRuns=8)
        result = sheet.handle_txo(data)
        assert result.rich_text_runlist == []

    def test_multiple_run_continue_records(self, sheet, mocker):
        run1 = struct.pack('<HH4x', 0, 2)
        run2 = struct.pack('<HH4x', 2, 3)
        mocker.patch.object(
            sheet.book, 'get_record_parts',
            side_effect=[
                (XL_CONTINUE, 8, run1),
                (XL_CONTINUE, 8, run2),
            ]
        )
        data = _make_txo_data(cchText=0, cbRuns=16)
        result = sheet.handle_txo(data)
        assert len(result.rich_text_runlist) == 2
        assert result.rich_text_runlist[0] == (0, 2)
        assert result.rich_text_runlist[1] == (2, 3)


class TestHandleTxoDebugMode:
    def test_debug_mode_does_not_raise(self, sheet):
        original = sheet_module.OBJ_MSO_DEBUG
        try:
            sheet_module.OBJ_MSO_DEBUG = 1
            data = _make_txo_data(cchText=0, cbRuns=0)
            result = sheet.handle_txo(data)
            assert result is not None
        finally:
            sheet_module.OBJ_MSO_DEBUG = original
