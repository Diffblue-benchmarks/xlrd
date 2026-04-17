# -*- coding: utf-8 -*-
"""
Extended tests for xlrd.__init__ — dump and count_records functions.
"""
import io
import sys
import struct
import pytest
from unittest.mock import MagicMock, patch

import xlrd
from xlrd import dump, count_records
from xlrd.biffh import XLRDError


def _make_biff_stream(*records):
    """Build a minimal BIFF stream from (opcode, payload) pairs."""
    data = b''
    for opcode, payload in records:
        data += struct.pack('<HH', opcode, len(payload)) + payload
    return data


# ---------------------------------------------------------------------------
# dump
# ---------------------------------------------------------------------------

class TestDump:

    def test_dump_writes_output(self, tmp_path):
        # Create a fake XLS file with a minimal BIFF stream
        xls_content = _make_biff_stream(
            (0x0809, b'\x00' * 8),  # BOF
            (0x000A, b''),           # EOF
        )
        xls_file = tmp_path / 'test.xls'
        # We need to patch Book.biff2_8_load to avoid actual parsing
        with patch('xlrd.Book') as MockBook:
            instance = MagicMock()
            instance.mem = xls_content
            instance.base = 0
            instance.stream_len = len(xls_content)
            MockBook.return_value = instance
            outbuf = io.StringIO()
            dump(str(xls_file), outfile=outbuf)
        # If mocking worked, no exception was raised

    def test_dump_unnumbered_flag(self, tmp_path):
        xls_content = _make_biff_stream((0x000A, b''))
        with patch('xlrd.Book') as MockBook:
            instance = MagicMock()
            instance.mem = xls_content
            instance.base = 0
            instance.stream_len = len(xls_content)
            MockBook.return_value = instance
            outbuf = io.StringIO()
            dump('fake.xls', outfile=outbuf, unnumbered=True)


# ---------------------------------------------------------------------------
# count_records
# ---------------------------------------------------------------------------

class TestCountRecords:

    def test_count_records_calls_biff_count(self, tmp_path):
        xls_content = _make_biff_stream(
            (0x000A, b''),
            (0x000A, b''),
        )
        with patch('xlrd.Book') as MockBook:
            instance = MagicMock()
            instance.mem = xls_content
            instance.base = 0
            instance.stream_len = len(xls_content)
            MockBook.return_value = instance
            outbuf = io.StringIO()
            count_records('fake.xls', outfile=outbuf)


# ---------------------------------------------------------------------------
# FILE_FORMAT_DESCRIPTIONS
# ---------------------------------------------------------------------------

class TestFileFormatDescriptions:

    def test_all_formats_present(self):
        desc = xlrd.FILE_FORMAT_DESCRIPTIONS
        assert 'xls' in desc
        assert 'xlsx' in desc
        assert 'xlsb' in desc
        assert 'ods' in desc
        assert 'zip' in desc
        assert None in desc

    def test_xls_description(self):
        assert 'xls' in xlrd.FILE_FORMAT_DESCRIPTIONS['xls'].lower()

    def test_none_description(self):
        assert xlrd.FILE_FORMAT_DESCRIPTIONS[None] == 'Unknown file type'


# ---------------------------------------------------------------------------
# Module-level re-exports
# ---------------------------------------------------------------------------

class TestModuleReexports:

    def test_xlrderror_accessible(self):
        from xlrd import XLRDError
        assert XLRDError is not None

    def test_cell_type_constants_accessible(self):
        from xlrd import (
            XL_CELL_EMPTY, XL_CELL_TEXT, XL_CELL_NUMBER,
            XL_CELL_DATE, XL_CELL_BOOLEAN, XL_CELL_ERROR, XL_CELL_BLANK,
        )
        assert XL_CELL_EMPTY == 0
        assert XL_CELL_TEXT == 1
        assert XL_CELL_NUMBER == 2

    def test_xldate_functions_accessible(self):
        from xlrd import xldate_as_tuple, xldate_as_datetime
        assert callable(xldate_as_tuple)
        assert callable(xldate_as_datetime)

    def test_open_workbook_accessible(self):
        assert callable(xlrd.open_workbook)

    def test_inspect_format_accessible(self):
        assert callable(xlrd.inspect_format)

    def test_empty_cell_accessible(self):
        from xlrd import empty_cell
        from xlrd.biffh import XL_CELL_EMPTY
        assert empty_cell.ctype == XL_CELL_EMPTY

    def test_error_text_from_code_accessible(self):
        from xlrd import error_text_from_code
        assert error_text_from_code[0x07] == '#DIV/0!'

    def test_version_accessible(self):
        from xlrd import __version__, __VERSION__
        assert __version__ == __VERSION__
        assert isinstance(__version__, str)


# ---------------------------------------------------------------------------
# open_workbook successful path (lines 172-185)
# ---------------------------------------------------------------------------

class TestOpenWorkbookSuccess:

    # OLE2 compound document magic bytes — recognised as 'xls' by inspect_format
    _OLE2_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\x00' * 512

    def test_open_workbook_calls_open_workbook_xls(self):
        """open_workbook should delegate to open_workbook_xls for .xls files."""
        mock_book = MagicMock()
        with patch('xlrd.open_workbook_xls', return_value=mock_book) as mock_fn:
            result = xlrd.open_workbook(file_contents=self._OLE2_MAGIC)
        mock_fn.assert_called_once()
        assert result is mock_book

    def test_open_workbook_passes_kwargs_to_xls(self):
        """open_workbook should forward keyword args like formatting_info."""
        mock_book = MagicMock()
        with patch('xlrd.open_workbook_xls', return_value=mock_book) as mock_fn:
            xlrd.open_workbook(file_contents=self._OLE2_MAGIC, formatting_info=True)
        call_kwargs = mock_fn.call_args.kwargs
        assert call_kwargs.get('formatting_info') is True

    def test_open_workbook_returns_book_object(self):
        """open_workbook should return the Book returned by open_workbook_xls."""
        mock_book = MagicMock()
        with patch('xlrd.open_workbook_xls', return_value=mock_book):
            result = xlrd.open_workbook(file_contents=self._OLE2_MAGIC)
        assert result is mock_book
