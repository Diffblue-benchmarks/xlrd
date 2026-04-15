# -*- coding: utf-8 -*-
"""Unit tests for Sheet.string_record_contents."""

import sys
import struct
import pytest

from xlrd.sheet import Sheet
from xlrd.biffh import XLRDError, XL_CONTINUE


class MockBook:
    biff_version = 80
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    encoding = ''
    _xf_index_to_xl_type_map = {}
    _sheet_visibility = [0, 0, 0, 0, 0]

    def get_record_parts(self):
        raise NotImplementedError("override in tests")

    def derive_encoding(self):
        return 'latin_1'


@pytest.fixture
def biff8_sheet():
    book = MockBook()
    book.biff_version = 80
    return Sheet(book, 0, 'TestSheet', 0)


@pytest.fixture
def biff7_sheet():
    book = MockBook()
    book.biff_version = 70
    return Sheet(book, 0, 'TestSheet', 0)


@pytest.fixture
def biff2_sheet():
    book = MockBook()
    book.biff_version = 20
    return Sheet(book, 0, 'TestSheet', 0)


class TestStringRecordContentsBiff8:
    def test_latin1_flag0_returns_correct_string(self, biff8_sheet):
        text = b'hello'
        data = struct.pack('<H', 5) + b'\x00' + text
        result = biff8_sheet.string_record_contents(data)
        assert result == 'hello'

    def test_utf16le_flag1_returns_correct_string(self, biff8_sheet):
        text = 'hi'.encode('utf_16_le')
        data = struct.pack('<H', 2) + b'\x01' + text
        result = biff8_sheet.string_record_contents(data)
        assert result == 'hi'

    def test_empty_string_biff8(self, biff8_sheet):
        data = struct.pack('<H', 0) + b'\x00' + b''
        result = biff8_sheet.string_record_contents(data)
        assert result == ''

    def test_too_many_chars_raises_xlrderror(self, biff8_sheet):
        # nchars_expected=2, but data has 5 chars -> nchars_found > nchars_expected
        text = b'hello'
        data = struct.pack('<H', 2) + b'\x00' + text
        with pytest.raises(XLRDError, match="STRING/CONTINUE"):
            biff8_sheet.string_record_contents(data)

    def test_continuation_record_latin1(self, mocker):
        book = MockBook()
        book.biff_version = 80
        sh = Sheet(book, 0, 'TestSheet', 0)
        # First chunk: 5 chars, but nchars_expected=10
        first_text = b'hello'
        data = struct.pack('<H', 10) + b'\x00' + first_text
        # Continuation: flag=0 (latin1), then 5 more chars
        cont_text = b'world'
        cont_data = b'\x00' + cont_text
        mocker.patch.object(book, 'get_record_parts', return_value=(XL_CONTINUE, len(cont_data), cont_data))
        result = sh.string_record_contents(data)
        assert result == 'helloworld'

    def test_continuation_record_utf16(self, mocker):
        book = MockBook()
        book.biff_version = 80
        sh = Sheet(book, 0, 'TestSheet', 0)
        # First chunk: 2 chars in utf16le, but nchars_expected=4
        first_text = 'ab'.encode('utf_16_le')
        data = struct.pack('<H', 4) + b'\x01' + first_text
        # Continuation: flag=1 (utf16le), 2 more chars
        cont_text = 'cd'.encode('utf_16_le')
        cont_data = b'\x01' + cont_text
        mocker.patch.object(book, 'get_record_parts', return_value=(XL_CONTINUE, len(cont_data), cont_data))
        result = sh.string_record_contents(data)
        assert result == 'abcd'

    def test_wrong_continuation_record_type_raises(self, mocker):
        book = MockBook()
        book.biff_version = 80
        sh = Sheet(book, 0, 'TestSheet', 0)
        first_text = b'hi'
        data = struct.pack('<H', 10) + b'\x00' + first_text
        mocker.patch.object(book, 'get_record_parts', return_value=(0x0809, 0, b''))
        with pytest.raises(XLRDError, match="Expected CONTINUE record"):
            sh.string_record_contents(data)


class TestStringRecordContentsBiffLt80:
    def test_biff7_lenlen2_with_encoding(self, biff7_sheet):
        text = b'world'
        data = struct.pack('<H', 5) + text
        result = biff7_sheet.string_record_contents(data)
        assert result == 'world'

    def test_biff7_uses_book_encoding(self, mocker):
        book = MockBook()
        book.biff_version = 70
        book.encoding = 'latin_1'
        sh = Sheet(book, 0, 'TestSheet', 0)
        text = b'test'
        data = struct.pack('<H', 4) + text
        result = sh.string_record_contents(data)
        assert result == 'test'

    def test_biff7_uses_derive_encoding_when_no_book_encoding(self, mocker):
        book = MockBook()
        book.biff_version = 70
        book.encoding = ''
        derive_mock = mocker.patch.object(book, 'derive_encoding', return_value='latin_1')
        sh = Sheet(book, 0, 'TestSheet', 0)
        text = b'foo'
        data = struct.pack('<H', 3) + text
        result = sh.string_record_contents(data)
        assert result == 'foo'
        derive_mock.assert_called_once()

    def test_biff7_too_many_chars_raises(self, biff7_sheet):
        text = b'hello'
        data = struct.pack('<H', 2) + text
        with pytest.raises(XLRDError, match="STRING/CONTINUE"):
            biff7_sheet.string_record_contents(data)

    def test_biff7_continuation_correct_rc(self, mocker):
        book = MockBook()
        book.biff_version = 70
        book.encoding = 'latin_1'
        sh = Sheet(book, 0, 'TestSheet', 0)
        first_text = b'abc'
        data = struct.pack('<H', 6) + first_text
        cont_text = b'def'
        mocker.patch.object(book, 'get_record_parts', return_value=(XL_CONTINUE, len(cont_text), cont_text))
        result = sh.string_record_contents(data)
        assert result == 'abcdef'

    def test_biff7_wrong_continuation_raises(self, mocker):
        book = MockBook()
        book.biff_version = 70
        book.encoding = 'latin_1'
        sh = Sheet(book, 0, 'TestSheet', 0)
        first_text = b'abc'
        data = struct.pack('<H', 6) + first_text
        mocker.patch.object(book, 'get_record_parts', return_value=(0x0809, 0, b''))
        with pytest.raises(XLRDError, match="Expected CONTINUE record"):
            sh.string_record_contents(data)


class TestStringRecordContentsBiffLt30:
    def test_biff2_lenlen1_single_byte_length(self, biff2_sheet):
        # biff_version < 30: lenlen = 1, uses 'B' format (1 byte for nchars)
        text = b'hi'
        data = struct.pack('<B', 2) + text
        result = biff2_sheet.string_record_contents(data)
        assert result == 'hi'
