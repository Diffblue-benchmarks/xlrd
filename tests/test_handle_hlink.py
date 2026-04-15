# -*- coding: utf-8 -*-
"""Unit tests for Sheet.handle_hlink."""

import sys
import pytest
from struct import pack

from xlrd.sheet import Sheet
from xlrd.biffh import XLRDError

GUID0 = b"\xD0\xC9\xEA\x79\xF9\xBA\xCE\x11\x8C\x82\x00\xAA\x00\x4B\xA9\x0B"
DUMMY = b"\x02\x00\x00\x00"
URL_CLSID = b"\xE0\xC9\xEA\x79\xF9\xBA\xCE\x11\x8C\x82\x00\xAA\x00\x4B\xA9\x0B"
FILE_CLSID = b"\x03\x03\x00\x00\x00\x00\x00\x00\xC0\x00\x00\x00\x00\x00\x00\x46"
UNKNOWN_CLSID = b"\xAA\xBB\xCC\xDD\xEE\xFF\x00\x11\x22\x33\x44\x55\x66\x77\x88\x99"


class MockBook:
    biff_version = 80
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    _xf_index_to_xl_type_map = {}
    _sheet_visibility = [0, 0, 0, 0, 0]


def _make_header(frowx=0, lrowx=0, fcolx=0, lcolx=0, options=0):
    return pack('<HHHH16s4si', frowx, lrowx, fcolx, lcolx, GUID0, DUMMY, options)


def _encode_nul_unicode(text):
    """Encode text as nul-terminated UTF-16LE with 4-byte char count prefix."""
    encoded = (text + '\x00').encode('UTF-16le')
    count = len(text) + 1
    return pack('<L', count) + encoded


def _make_url_moniker(url):
    url_bytes = (url + '\x00').encode('UTF-16le')
    nbytes = len(url_bytes)
    return URL_CLSID + pack('<L', nbytes) + url_bytes


def _make_url_moniker_with_extra(url, extra_bytes=24):
    url_bytes = (url + '\x00').encode('UTF-16le')
    extra = b'\x00' * extra_bytes
    nbytes = len(url_bytes) + extra_bytes
    return URL_CLSID + pack('<L', nbytes) + url_bytes + extra


def _make_file_moniker(shortpath_str, extended_path=None, uplevels=0):
    shortpath = shortpath_str.encode('latin1') + b'\x00'
    nbytes = len(shortpath)
    unknown_24 = pack('<H', 0xDEAD) + b'\x00' * 22
    if extended_path:
        ext_bytes = extended_path.encode('UTF-16le')
        xl = len(ext_bytes)
        sz_value = 4 + 2 + xl
        return (
            FILE_CLSID
            + pack('<Hi', uplevels, nbytes)
            + shortpath
            + unknown_24
            + pack('<i', sz_value)
            + pack('<i', xl)
            + b'\x03\x00'
            + ext_bytes
        )
    else:
        return (
            FILE_CLSID
            + pack('<Hi', uplevels, nbytes)
            + shortpath
            + unknown_24
            + pack('<i', 0)
        )


@pytest.fixture
def sheet():
    return Sheet(MockBook(), 0, 'TestSheet', 0)


class TestHandleHlink:
    def test_url_hyperlink(self, sheet):
        options = 0x01
        header = _make_header(options=options)
        url = 'http://example.com'
        data = header + _make_url_moniker(url)

        sheet.handle_hlink(data)

        assert len(sheet.hyperlink_list) == 1
        h = sheet.hyperlink_list[0]
        assert h.type == 'url'
        assert h.url_or_path == url

    def test_url_hyperlink_with_extra_24_bytes(self, sheet):
        options = 0x01
        header = _make_header(options=options)
        url = 'http://example.org'
        data = header + _make_url_moniker_with_extra(url, extra_bytes=24)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.type == 'url'
        assert h.url_or_path == url

    def test_local_file_hyperlink_with_extended_path(self, sheet):
        options = 0x01
        header = _make_header(options=options)
        extended = 'C:\\folder\\file.xls'
        data = header + _make_file_moniker('shortname.xls', extended_path=extended)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.type == 'local file'
        assert h.url_or_path == extended

    def test_local_file_hyperlink_short_path_only(self, sheet):
        options = 0x01
        header = _make_header(options=options)
        data = header + _make_file_moniker('readme.txt', extended_path=None)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.type == 'local file'
        assert h.url_or_path == b'readme.txt'

    def test_local_file_hyperlink_with_uplevels(self, sheet):
        options = 0x01
        header = _make_header(options=options)
        data = header + _make_file_moniker('doc.xls', extended_path=None, uplevels=2)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.type == 'local file'
        assert h.url_or_path == b'..\\..\\doc.xls'

    def test_unknown_clsid_moniker(self, sheet, capsys):
        options = 0x01
        header = _make_header(options=options)
        data = header + UNKNOWN_CLSID

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.type is None

    def test_unc_hyperlink(self, sheet):
        # options & 0x163 == 0x103 -> UNC
        options = 0x103
        header = _make_header(options=options)
        unc_path = '\\\\server\\share'
        data = header + _encode_nul_unicode(unc_path)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.type == 'unc'
        assert h.url_or_path == unc_path

    def test_workbook_hyperlink(self, sheet):
        # options & 0x16B == 8 -> workbook; bit 3 (0x8) also triggers textmark
        options = 0x08
        header = _make_header(options=options)
        textmark = 'Sheet1!A1'
        data = header + _encode_nul_unicode(textmark)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.type == 'workbook'
        assert h.textmark == textmark

    def test_unknown_hyperlink_type(self, sheet):
        options = 0
        header = _make_header(options=options)
        data = header

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.type == 'unknown'

    def test_hyperlink_with_description(self, sheet):
        options = 0x14  # has description
        header = _make_header(options=options)
        desc = 'Click here'
        data = header + _encode_nul_unicode(desc)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.desc == desc

    def test_hyperlink_with_target(self, sheet):
        options = 0x80  # has target
        header = _make_header(options=options)
        target = '_blank'
        data = header + _encode_nul_unicode(target)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.target == target

    def test_hyperlink_with_textmark(self, sheet):
        # options = 0x08: workbook type + has textmark
        options = 0x08
        header = _make_header(options=options)
        textmark = 'Sheet1!A1'
        data = header + _encode_nul_unicode(textmark)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.textmark == textmark

    def test_hyperlink_with_description_and_textmark(self, sheet):
        # options = 0x1C: has description (0x14) + has textmark (0x08)
        options = 0x1C
        header = _make_header(options=options)
        desc = 'My Link'
        textmark = 'A1'
        data = header + _encode_nul_unicode(desc) + _encode_nul_unicode(textmark)

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.desc == desc
        assert h.textmark == textmark

    def test_hyperlink_updates_map_single_cell(self, sheet):
        options = 0
        header = _make_header(frowx=2, lrowx=2, fcolx=3, lcolx=3, options=options)
        data = header

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert sheet.hyperlink_map[2, 3] is h

    def test_hyperlink_updates_map_range(self, sheet):
        options = 0
        header = _make_header(frowx=1, lrowx=3, fcolx=0, lcolx=2, options=options)
        data = header

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        for rowx in range(1, 4):
            for colx in range(0, 3):
                assert sheet.hyperlink_map[rowx, colx] is h

    def test_hyperlink_stores_row_col_coords(self, sheet):
        options = 0
        header = _make_header(frowx=5, lrowx=7, fcolx=2, lcolx=4, options=options)
        data = header

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.frowx == 5
        assert h.lrowx == 7
        assert h.fcolx == 2
        assert h.lcolx == 4

    def test_hyperlink_extra_data_logs_warning(self, sheet, capsys):
        options = 0
        header = _make_header(options=options)
        extra = b'\x00\x00'
        data = header + extra

        sheet.handle_hlink(data)

        assert len(sheet.hyperlink_list) == 1

    def test_hyperlink_url_with_description_and_target(self, sheet):
        # options = 0x14 | 0x80 | 0x01 = 0x95: desc + target + HasMoniker
        options = 0x95
        header = _make_header(options=options)
        desc = 'My URL'
        target = '_self'
        url = 'http://python.org'
        data = (
            header
            + _encode_nul_unicode(desc)
            + _encode_nul_unicode(target)
            + _make_url_moniker(url)
        )

        sheet.handle_hlink(data)

        h = sheet.hyperlink_list[0]
        assert h.desc == desc
        assert h.target == target
        assert h.type == 'url'
        assert h.url_or_path == url
