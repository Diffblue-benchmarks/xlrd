"""
Tests targeting uncovered lines in xlrd/book.py Book.parse_globals.

Each test exercises a specific record handler branch dispatched inside
parse_globals by inserting the corresponding record into a minimal BIFF8
XLS byte stream.
"""
import struct
import pytest

import xlrd
from xlrd.biffh import XLRDError


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_record(rec_type, data=b''):
    return struct.pack('<HH', rec_type, len(data)) + data


def make_globals_xls(extra_records=b'', sheet_names=None, with_codepage=True):
    """Build a minimal BIFF8 XLS stream with optional extra globals records inserted
    before the BOUNDSHEET records and EOF."""
    if sheet_names is None:
        sheet_names = ['Sheet1']

    bof_wbk = _make_record(
        0x0809,
        struct.pack('<HHHH', 0x0600, 0x0005, 0x0DBB, 0x07CC) + b'\x00\x00',
    )
    codepage = _make_record(0x0042, struct.pack('<H', 1200)) if with_codepage else b''
    eof = _make_record(0x000A, b'')

    sheet_streams = []
    for _name in sheet_names:
        bof_sheet = _make_record(
            0x0809,
            struct.pack('<HHHH', 0x0600, 0x0010, 0x0DBB, 0x07CC) + b'\x00\x00',
        )
        sheet_eof = _make_record(0x000A, b'')
        sheet_streams.append(bof_sheet + sheet_eof)

    globals_without_eof = bof_wbk + codepage + extra_records

    # First pass: placeholder to measure size for offset calculation
    bs_placeholder = b''
    for name in sheet_names:
        name_bytes = name.encode('utf-16-le')
        bs_data = struct.pack('<IH', 0, 0) + struct.pack('<BB', len(name), 1) + name_bytes
        bs_placeholder += _make_record(0x0085, bs_data)

    globals_size = len(globals_without_eof) + len(bs_placeholder) + len(eof)

    # Second pass: build BOUNDSHEET records with correct file offsets
    offset = globals_size
    final_bs_records = b''
    for i, name in enumerate(sheet_names):
        name_bytes = name.encode('utf-16-le')
        bs_data = struct.pack('<IH', offset, 0) + struct.pack('<BB', len(name), 1) + name_bytes
        final_bs_records += _make_record(0x0085, bs_data)
        offset += len(sheet_streams[i])

    full_stream = globals_without_eof + final_bs_records + eof
    for ss in sheet_streams:
        full_stream += ss
    return full_stream


# ---------------------------------------------------------------------------
# Tests for parse_globals record dispatch (uncovered lines)
# ---------------------------------------------------------------------------

class TestParseGlobalsRecords:

    def test_sst_record_handled(self):
        """XL_SST (0xfc) dispatch: line 1211 – handle_sst called."""
        # Minimal SST: total_strings=0, unique_strings=0
        sst_data = struct.pack('<II', 0, 0)
        data = make_globals_xls(extra_records=_make_record(0xFC, sst_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.nsheets == 1

    def test_font_record_handled(self):
        """XL_FONT (0x31) dispatch: line 1213 – handle_font called."""
        # BIFF8 font: 13-byte header + 1 pad + 1 name_len=0 = 15 bytes
        font_data = struct.pack('<HHHHHBBB', 200, 0, 0x7FFF, 400, 0, 0, 0, 0) + b'\x00\x00'
        data = make_globals_xls(extra_records=_make_record(0x31, font_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.nsheets == 1

    def test_format_record_handled(self):
        """XL_FORMAT (0x41e) dispatch: line 1215 – handle_format called."""
        # BIFF8 FORMAT: 2-byte key + 2-byte nchars=0 (empty unicode string)
        fmt_data = struct.pack('<HH', 164, 0)
        data = make_globals_xls(extra_records=_make_record(0x041E, fmt_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.nsheets == 1

    def test_xf_record_handled(self):
        """XL_XF (0xe0) dispatch: line 1217 – handle_xf called."""
        # BIFF8 XF record: format '<HHHBBBBIiH' = 20 bytes
        xf_data = struct.pack('<HHHBBBBIiH', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        data = make_globals_xls(extra_records=_make_record(0xE0, xf_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.nsheets == 1

    def test_datemode_record_handled(self):
        """XL_DATEMODE (0x22) dispatch: line 1221 – handle_datemode called."""
        datemode_data = struct.pack('<H', 0)  # 1900 mode
        data = make_globals_xls(extra_records=_make_record(0x22, datemode_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.datemode == 0

    def test_country_record_handled(self):
        """XL_COUNTRY (0x8c) dispatch: line 1225 – handle_country called."""
        country_data = struct.pack('<HH', 1, 1)  # USA, USA
        data = make_globals_xls(extra_records=_make_record(0x8C, country_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.countries == (1, 1)

    def test_externname_record_handled(self):
        """XL_EXTERNNAME (0x23) dispatch: line 1227 – handle_externname called.

        EXTERNNAME requires a preceding SUPBOOK record so that
        _supbook_types[-1] is valid.
        """
        # Internal SUPBOOK: num_sheets=1 + magic b'\x01\x04'
        supbook_data = struct.pack('<H', 1) + b'\x01\x04'
        # BIFF8 EXTERNNAME: option_flags(H) + other_info(I) + name_len(B)=0 → name=""
        externname_data = struct.pack('<HIB', 0, 0, 0)
        extra = _make_record(0x01AE, supbook_data) + _make_record(0x23, externname_data)
        data = make_globals_xls(extra_records=extra)
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.nsheets == 1

    def test_externsheet_record_handled(self):
        """XL_EXTERNSHEET (0x17) dispatch: line 1229 – handle_externsheet called."""
        ext_data = struct.pack('<H', 0)  # num_refs=0
        data = make_globals_xls(extra_records=_make_record(0x17, ext_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.nsheets == 1

    def test_filepass_record_raises_encrypted_error(self):
        """XL_FILEPASS (0x2f) dispatch: line 1231 – handle_filepass raises XLRDError."""
        filepass_data = struct.pack('<H', 0) + b'\x00' * 4
        data = make_globals_xls(extra_records=_make_record(0x2F, filepass_data))
        with pytest.raises(XLRDError, match='encrypted'):
            xlrd.open_workbook(file_contents=data)

    def test_writeaccess_record_handled(self):
        """XL_WRITEACCESS (0x5c) dispatch: line 1233 – handle_writeaccess called."""
        # BIFF8: unpack_unicode(data, 0, lenlen=2) – nchars=0 returns "" immediately
        wa_data = struct.pack('<H', 0)
        data = make_globals_xls(extra_records=_make_record(0x5C, wa_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.user_name == ''

    def test_sheetsoffset_record_handled(self):
        """XL_SHEETSOFFSET (0x8e) dispatch: line 1235 – handle_sheetsoffset called."""
        offset_data = struct.pack('<i', 0)
        data = make_globals_xls(extra_records=_make_record(0x8E, offset_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.nsheets == 1

    @pytest.mark.skip(reason="BIFF4W SHEETHDR requires a full BIFF4W workbook stream which is not practical to construct in a unit test")
    def test_sheethdr_record_handled(self):
        """XL_SHEETHDR (0x8f) dispatch: line 1237 – handle_sheethdr called (BIFF4W only)."""

    def test_supbook_record_handled(self):
        """XL_SUPBOOK (0x1ae) dispatch: line 1239 – handle_supbook called."""
        # Internal SUPBOOK: num_sheets=1 + magic b'\x01\x04'
        supbook_data = struct.pack('<H', 1) + b'\x01\x04'
        data = make_globals_xls(extra_records=_make_record(0x01AE, supbook_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.nsheets == 1

    def test_name_record_handled(self):
        """XL_NAME (0x18) dispatch: line 1241 – handle_name called."""
        # 14-byte header: all zeros (option_flags=0, name_len=0, fmla_len=0, ...)
        name_data = struct.pack('<HBBHHH4B', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        data = make_globals_xls(extra_records=_make_record(0x18, name_data))
        bk = xlrd.open_workbook(file_contents=data)
        assert len(bk.name_obj_list) == 1

    def test_palette_record_handled(self):
        """XL_PALETTE (0x92) dispatch: line 1243 – handle_palette called."""
        # BIFF8 expects 56 colours; formatting_info=True required to process
        n_colours = 56
        palette_data = struct.pack('<H', n_colours) + b'\x00' * (n_colours * 4)
        data = make_globals_xls(extra_records=_make_record(0x92, palette_data))
        bk = xlrd.open_workbook(file_contents=data, formatting_info=True)
        assert bk.nsheets == 1

    def test_style_record_handled(self):
        """XL_STYLE (0x293) dispatch: line 1245 – handle_style called."""
        # Built-in style: 0x8000 flag set → Normal style, built_in_id=0, level=0xff
        style_data = struct.pack('<HBB', 0x8000, 0, 0xFF)
        data = make_globals_xls(extra_records=_make_record(0x0293, style_data))
        bk = xlrd.open_workbook(file_contents=data, formatting_info=True)
        assert bk.nsheets == 1

    def test_unexpected_bof_logged_with_verbosity(self):
        """Unexpected BOF branch (rc & 0xff == 9, verbosity): line 1247."""
        # 0x0109 has lower byte 9 and is not any known globals record
        data = make_globals_xls(
            extra_records=_make_record(0x0109, b'\x00' * 4),
        )
        bk = xlrd.open_workbook(file_contents=data, verbosity=1)
        assert bk.nsheets == 1

    def test_derive_encoding_called_when_no_codepage(self):
        """Line 1254: derive_encoding called after EOF when no CODEPAGE record."""
        data = make_globals_xls(with_codepage=False)
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.encoding is not None

    def test_unknown_record_code_silently_ignored(self):
        """Else branch (line 1265: pass) – unrecognised record codes are skipped."""
        data = make_globals_xls(extra_records=_make_record(0x9999, b'\x01\x02\x03\x04'))
        bk = xlrd.open_workbook(file_contents=data)
        assert bk.nsheets == 1
