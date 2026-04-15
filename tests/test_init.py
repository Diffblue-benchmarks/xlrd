import io
import zipfile

import pytest

import xlrd
from xlrd import count_records, dump, inspect_format, open_workbook
from xlrd.biffh import XLRDError
from xlrd.compdoc import SIGNATURE as XLS_SIGNATURE


def make_zip_content(names):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        for name in names:
            zf.writestr(name, b'')
    return buf.getvalue()


class TestInspectFormat:

    def test_xls_from_content(self):
        content = XLS_SIGNATURE + b'\x00' * 100
        assert inspect_format(content=content) == 'xls'

    def test_xlsx_from_content(self):
        content = make_zip_content(['xl/workbook.xml'])
        assert inspect_format(content=content) == 'xlsx'

    def test_xlsb_from_content(self):
        content = make_zip_content(['xl/workbook.bin'])
        assert inspect_format(content=content) == 'xlsb'

    def test_ods_from_content(self):
        content = make_zip_content(['content.xml'])
        assert inspect_format(content=content) == 'ods'

    def test_zip_unknown_from_content(self):
        content = make_zip_content(['some_other_file.txt'])
        assert inspect_format(content=content) == 'zip'

    def test_unknown_format_from_content(self):
        content = b'\x00\x01\x02\x03\x04\x05\x06\x07'
        assert inspect_format(content=content) is None

    def test_xls_from_path(self, tmp_path):
        f = tmp_path / 'test.xls'
        f.write_bytes(XLS_SIGNATURE + b'\x00' * 100)
        assert inspect_format(path=str(f)) == 'xls'

    def test_xlsx_from_path(self, tmp_path):
        content = make_zip_content(['xl/workbook.xml'])
        f = tmp_path / 'test.xlsx'
        f.write_bytes(content)
        assert inspect_format(path=str(f)) == 'xlsx'

    def test_unknown_from_path(self, tmp_path):
        f = tmp_path / 'test.bin'
        f.write_bytes(b'\x00\x01\x02\x03\x04\x05\x06\x07')
        assert inspect_format(path=str(f)) is None

    def test_zip_backslash_names_normalized(self):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w') as zf:
            zf.writestr('xl\\workbook.xml', b'')
        content = buf.getvalue()
        assert inspect_format(content=content) == 'xlsx'


class TestOpenWorkbook:

    def test_xlsx_raises_error(self):
        content = make_zip_content(['xl/workbook.xml'])
        with pytest.raises(XLRDError):
            open_workbook(file_contents=content)

    def test_xlsb_raises_error(self):
        content = make_zip_content(['xl/workbook.bin'])
        with pytest.raises(XLRDError):
            open_workbook(file_contents=content)

    def test_ods_raises_error(self):
        content = make_zip_content(['content.xml'])
        with pytest.raises(XLRDError):
            open_workbook(file_contents=content)

    def test_zip_raises_error(self):
        content = make_zip_content(['some_other_file.txt'])
        with pytest.raises(XLRDError):
            open_workbook(file_contents=content)


def _setup_book_attrs(self, **kwargs):
    self.mem = b''
    self.base = 0
    self.stream_len = 0


class TestDump:

    def test_dump_calls_biff_load(self, tmp_path, mocker):
        f = tmp_path / 'test.xls'
        f.write_bytes(b'\x00' * 10)
        mock_load = mocker.patch('xlrd.book.Book.biff2_8_load', autospec=True, side_effect=_setup_book_attrs)
        mocker.patch('xlrd.biffh.biff_dump')
        out = io.StringIO()
        dump(str(f), outfile=out)
        mock_load.assert_called_once()

    def test_dump_unnumbered(self, tmp_path, mocker):
        f = tmp_path / 'test.xls'
        f.write_bytes(b'\x00' * 10)
        mocker.patch('xlrd.book.Book.biff2_8_load', autospec=True, side_effect=_setup_book_attrs)
        mock_biff_dump = mocker.patch('xlrd.biffh.biff_dump')
        out = io.StringIO()
        dump(str(f), outfile=out, unnumbered=True)
        args = mock_biff_dump.call_args[0]
        assert args[-1] is True  # unnumbered passed as last positional arg


class TestCountRecords:

    def test_count_records_calls_biff_load(self, tmp_path, mocker):
        f = tmp_path / 'test.xls'
        f.write_bytes(b'\x00' * 10)
        mock_load = mocker.patch('xlrd.book.Book.biff2_8_load', autospec=True, side_effect=_setup_book_attrs)
        mocker.patch('xlrd.biffh.biff_count_records')
        out = io.StringIO()
        count_records(str(f), outfile=out)
        mock_load.assert_called_once()
