# -*- coding: utf-8 -*-
"""Unit tests for Sheet.handle_note."""

import sys
import struct
import pytest

import xlrd.sheet as sheet_module
from xlrd.sheet import Sheet
from xlrd.biffh import XL_NOTE


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


def _make_note_data_biff8(rowx=0, colx=0, option_flags=0, object_id=1, author='Author'):
    """Pack a BIFF8 NOTE record with the given author string."""
    header = struct.pack('<4H', rowx, colx, option_flags, object_id)
    # Encode author as length-prefixed latin-1 string (lenlen=2, options byte=0)
    author_bytes = author.encode('latin-1')
    nchars = len(author_bytes)
    string_part = struct.pack('<H', nchars) + bytes([0]) + author_bytes
    return header + string_part


def _make_note_data_biff_old(rowx=0, colx=0, text=b'hello', encoding='latin-1'):
    """Pack a BIFF<80 NOTE record."""
    nb = len(text)
    data = struct.pack('<HHH', rowx, colx, nb) + text
    return data


class TestHandleNoteOldBiff:
    def test_biff_lt80_sets_basic_fields(self, mocker):
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        data = _make_note_data_biff_old(rowx=1, colx=2, text=b'hello')
        sh.handle_note(data, {})
        assert (1, 2) in sh.cell_note_map
        note = sh.cell_note_map[1, 2]
        assert note.rowx == 1
        assert note.colx == 2
        assert note.text == 'hello'

    def test_biff_lt80_rich_text_runlist(self, mocker):
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        data = _make_note_data_biff_old(rowx=0, colx=0, text=b'hi')
        sh.handle_note(data, {})
        note = sh.cell_note_map[0, 0]
        assert note.rich_text_runlist == [(0, 0)]

    def test_biff_lt80_show_and_hidden_flags(self, mocker):
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        data = _make_note_data_biff_old(rowx=0, colx=0, text=b'x')
        sh.handle_note(data, {})
        note = sh.cell_note_map[0, 0]
        assert note.show == 0
        assert note.row_hidden == 0
        assert note.col_hidden == 0

    def test_biff_lt80_author_is_empty_string(self, mocker):
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        data = _make_note_data_biff_old(rowx=0, colx=0, text=b'x')
        sh.handle_note(data, {})
        note = sh.cell_note_map[0, 0]
        assert note.author == ''

    def test_biff_lt80_object_id_is_none(self, mocker):
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        data = _make_note_data_biff_old(rowx=0, colx=0, text=b'test')
        sh.handle_note(data, {})
        note = sh.cell_note_map[0, 0]
        assert note._object_id is None

    def test_biff_lt80_continuation_records(self, mocker):
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        # expected_bytes = 10, but we only provide 5 here
        first_chunk = b'hello'
        second_chunk = b'world'
        header = struct.pack('<HHH', 0, 0, 10) + first_chunk
        # continuation record: rowx=0xFFFF, 2x bytes padding, nb=5, then 5 bytes of data
        cont_data = struct.pack('<H2xH', 0xFFFF, 5) + second_chunk
        mocker.patch.object(book, 'get_record_parts', return_value=(XL_NOTE, len(cont_data), cont_data))
        sh.handle_note(header, {})
        note = sh.cell_note_map[0, 0]
        assert note.text == 'helloworld'

    def test_biff_lt80_returns_none(self):
        book = MockBook()
        book.biff_version = 70
        sh = Sheet(book, 0, 'TestSheet', 0)
        data = _make_note_data_biff_old(rowx=0, colx=0, text=b'x')
        result = sh.handle_note(data, {})
        assert result is None


class TestHandleNoteBiff8:
    def test_biff8_fields_parsed(self, sheet):
        data = _make_note_data_biff8(rowx=2, colx=3, object_id=42, author='Alice')
        sheet.handle_note(data, {})
        # no txo means note not added to cell_note_map, but no exception raised
        assert (2, 3) not in sheet.cell_note_map

    def test_biff8_with_txo_adds_to_cell_note_map(self, sheet):
        data = _make_note_data_biff8(rowx=1, colx=1, object_id=7, author='Bob')

        class TXO:
            text = 'Note text here'
            rich_text_runlist = [(0, 0)]

        txos = {7: TXO()}
        sheet.handle_note(data, txos)
        assert (1, 1) in sheet.cell_note_map
        note = sheet.cell_note_map[1, 1]
        assert note.text == 'Note text here'
        assert note.rich_text_runlist == [(0, 0)]

    def test_biff8_option_flags_show(self, sheet):
        # show bit is bit 1 of option_flags
        option_flags = 0b00000010  # bit 1 set => show=1
        data = _make_note_data_biff8(option_flags=option_flags, object_id=1, author='A')

        class TXO:
            text = ''
            rich_text_runlist = []

        sheet.handle_note(data, {1: TXO()})
        note = sheet.cell_note_map[0, 0]
        assert note.show == 1

    def test_biff8_option_flags_row_hidden(self, sheet):
        # row_hidden bit is bit 7 of option_flags
        option_flags = 0b10000000  # bit 7 set => row_hidden=1
        data = _make_note_data_biff8(option_flags=option_flags, object_id=1, author='A')

        class TXO:
            text = ''
            rich_text_runlist = []

        sheet.handle_note(data, {1: TXO()})
        note = sheet.cell_note_map[0, 0]
        assert note.row_hidden == 1

    def test_biff8_option_flags_col_hidden(self, sheet):
        # col_hidden bit is bit 8 of option_flags
        option_flags = 0b100000000  # bit 8 set => col_hidden=1
        data = _make_note_data_biff8(option_flags=option_flags, object_id=1, author='A')

        class TXO:
            text = ''
            rich_text_runlist = []

        sheet.handle_note(data, {1: TXO()})
        note = sheet.cell_note_map[0, 0]
        assert note.col_hidden == 1

    def test_biff8_author_stored(self, sheet):
        data = _make_note_data_biff8(rowx=0, colx=0, object_id=5, author='TestAuthor')

        class TXO:
            text = 'x'
            rich_text_runlist = []

        sheet.handle_note(data, {5: TXO()})
        note = sheet.cell_note_map[0, 0]
        assert note.author == 'TestAuthor'

    def test_biff8_object_id_stored(self, sheet):
        data = _make_note_data_biff8(rowx=0, colx=0, object_id=99, author='A')

        class TXO:
            text = ''
            rich_text_runlist = []

        sheet.handle_note(data, {99: TXO()})
        note = sheet.cell_note_map[0, 0]
        assert note._object_id == 99

    def test_biff8_extra_padding_byte_allowed(self, sheet):
        data = _make_note_data_biff8(rowx=0, colx=0, object_id=1, author='Hi')
        # Add one extra byte (the optional undefined byte after the author)
        data = data + b'\x00'

        class TXO:
            text = 'abc'
            rich_text_runlist = []

        sheet.handle_note(data, {1: TXO()})
        note = sheet.cell_note_map[0, 0]
        assert note.text == 'abc'

    def test_biff8_no_txo_does_not_add_to_map(self, sheet):
        data = _make_note_data_biff8(rowx=3, colx=4, object_id=10, author='X')
        sheet.handle_note(data, {})
        assert (3, 4) not in sheet.cell_note_map

    def test_biff8_empty_author(self, sheet):
        # author string with nchars=0 and no options byte (edge case)
        header = struct.pack('<4H', 0, 0, 0, 1)
        # For zero-length string with no options byte, unpack_unicode_update_pos returns ''
        string_part = struct.pack('<H', 0)  # nchars=0, no options byte
        data = header + string_part

        class TXO:
            text = ''
            rich_text_runlist = []

        sheet.handle_note(data, {1: TXO()})
        note = sheet.cell_note_map[0, 0]
        assert note.author == ''


class TestHandleNoteDebugMode:
    def test_debug_mode_does_not_raise(self, sheet):
        original = sheet_module.OBJ_MSO_DEBUG
        try:
            sheet_module.OBJ_MSO_DEBUG = 1
            data = _make_note_data_biff8(rowx=0, colx=0, object_id=1, author='Debug')
            sheet.handle_note(data, {})
        finally:
            sheet_module.OBJ_MSO_DEBUG = original
