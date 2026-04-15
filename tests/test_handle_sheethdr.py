"""
Tests targeting uncovered lines in xlrd/book.py Book.handle_sheethdr.
"""
import struct
import sys
import pytest

from xlrd.book import Book


def _make_book(sheet_names=None, position=100):
    """Create a minimal Book instance with attributes needed by handle_sheethdr."""
    if sheet_names is None:
        sheet_names = ['Sheet1']
    bk = Book()
    bk.logfile = sys.stdout
    bk.verbosity = 0
    bk.formatting_info = False
    bk.encoding_override = 'ascii'
    bk.encoding = 'ascii'
    bk.codepage = 1200
    bk.biff_version = 45
    bk._sheet_names = list(sheet_names)
    bk._sheet_list = []
    bk._sheethdr_count = 0
    bk._position = position
    bk.mem = b'\x00' * 512
    return bk


def _make_sheethdr_data(sheet_name, sheet_len, encoding='ascii'):
    """Build SHEETHDR record data for the given sheet name and content length."""
    name_bytes = sheet_name.encode(encoding)
    data = struct.pack('<i', sheet_len)
    data += struct.pack('<B', len(name_bytes))
    data += name_bytes
    return data


class TestHandleSheethdr:

    def test_sheethdr_increments_sheethdr_count(self, mocker):
        """handle_sheethdr increments _sheethdr_count."""
        bk = _make_book(['Sheet1'])
        mocker.patch.object(bk, 'get_sheet', return_value=None)
        data = _make_sheethdr_data('Sheet1', 200)
        bk.handle_sheethdr(data)
        assert bk._sheethdr_count == 1

    def test_sheethdr_appends_none_to_sheet_list(self, mocker):
        """handle_sheethdr appends None to _sheet_list before calling get_sheet."""
        bk = _make_book(['Sheet1'])
        appended = []
        original_get_sheet = lambda sh_number, update_pos=True: appended.append(
            len(bk._sheet_list)
        )
        mocker.patch.object(bk, 'get_sheet', side_effect=original_get_sheet)
        data = _make_sheethdr_data('Sheet1', 200)
        bk.handle_sheethdr(data)
        # _sheet_list should have had None appended before get_sheet was called
        assert appended[0] == 1

    def test_sheethdr_updates_position(self, mocker):
        """handle_sheethdr sets _position to BOF_posn + sheet_len."""
        bk = _make_book(['Sheet1'])
        bk._position = 50
        mocker.patch.object(bk, 'get_sheet', return_value=None)
        sheet_len = 300
        data = _make_sheethdr_data('Sheet1', sheet_len)
        bk.handle_sheethdr(data)
        assert bk._position == 50 + sheet_len

    def test_sheethdr_calls_get_sheet_with_correct_sheetno(self, mocker):
        """handle_sheethdr calls get_sheet with sheetno and update_pos=False."""
        bk = _make_book(['Sheet1'])
        mock_get_sheet = mocker.patch.object(bk, 'get_sheet', return_value=None)
        data = _make_sheethdr_data('Sheet1', 200)
        bk.handle_sheethdr(data)
        mock_get_sheet.assert_called_once_with(0, update_pos=False)

    def test_sheethdr_calls_initialise_format_info(self, mocker):
        """handle_sheethdr calls initialise_format_info to reset format state."""
        bk = _make_book(['Sheet1'])
        mocker.patch.object(bk, 'get_sheet', return_value=None)
        mock_fmt = mocker.patch.object(bk, 'initialise_format_info', wraps=bk.initialise_format_info)
        data = _make_sheethdr_data('Sheet1', 200)
        bk.handle_sheethdr(data)
        mock_fmt.assert_called_once()

    def test_sheethdr_second_sheet(self, mocker):
        """handle_sheethdr handles second sheet correctly with _sheethdr_count=1."""
        bk = _make_book(['Sheet1', 'Sheet2'])
        mock_get_sheet = mocker.patch.object(bk, 'get_sheet', return_value=None)
        data0 = _make_sheethdr_data('Sheet1', 100)
        bk.handle_sheethdr(data0)
        # Reset position for second call
        bk._position = 200
        data1 = _make_sheethdr_data('Sheet2', 150)
        bk.handle_sheethdr(data1)
        assert bk._sheethdr_count == 2
        assert mock_get_sheet.call_args_list[1][0] == (1,)
        assert bk._position == 200 + 150

    def test_sheethdr_asserts_name_match(self, mocker):
        """handle_sheethdr raises AssertionError when sheet_name doesn't match _sheet_names."""
        bk = _make_book(['CorrectName'])
        mocker.patch.object(bk, 'get_sheet', return_value=None)
        data = _make_sheethdr_data('WrongName', 200)
        with pytest.raises(AssertionError):
            bk.handle_sheethdr(data)

    def test_sheethdr_derive_encoding_called(self, mocker):
        """handle_sheethdr calls derive_encoding to set self.encoding."""
        bk = _make_book(['Sheet1'])
        mocker.patch.object(bk, 'get_sheet', return_value=None)
        mock_derive = mocker.patch.object(bk, 'derive_encoding', wraps=bk.derive_encoding)
        data = _make_sheethdr_data('Sheet1', 200)
        bk.handle_sheethdr(data)
        mock_derive.assert_called_once()
