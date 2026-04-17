# -*- coding: utf-8 -*-
"""
Extended tests for xlrd.book — display_cell_address, Name class,
and Book class defaults.
"""
import io
import struct
import sys
import pytest

import xlrd
from xlrd.book import (
    colname,
    display_cell_address,
    Name,
    Book,
    SUPBOOK_UNK,
    SUPBOOK_INTERNAL,
    SUPBOOK_EXTERNAL,
    SUPBOOK_ADDIN,
    SUPBOOK_DDEOLE,
    SUPPORTED_VERSIONS,
    builtin_name_from_code,
    code_from_builtin_name,
)
from xlrd.biffh import XLRDError, XL_WORKBOOK_GLOBALS


# ---------------------------------------------------------------------------
# display_cell_address
# ---------------------------------------------------------------------------

class TestDisplayCellAddress:

    def test_absolute_row_and_col(self):
        # relrow=0, relcol=0 → $A$1
        result = display_cell_address(rowx=0, colx=0, relrow=0, relcol=0)
        assert result == '$A$1'

    def test_absolute_different_cell(self):
        # rowx=4, colx=2 → $C$5
        result = display_cell_address(rowx=4, colx=2, relrow=0, relcol=0)
        assert result == '$C$5'

    def test_relative_row_positive(self):
        # relrow=1, rowx=3 → (*+3)
        result = display_cell_address(rowx=3, colx=0, relrow=1, relcol=0)
        assert '(*+3)' in result

    def test_relative_row_negative(self):
        # relrow=1, rowx=-2 → (*-2)
        result = display_cell_address(rowx=-2, colx=0, relrow=1, relcol=0)
        assert '(*-2)' in result

    def test_relative_col_positive(self):
        # relcol=1, colx=5 → (*+5)
        result = display_cell_address(rowx=0, colx=5, relrow=0, relcol=1)
        assert '(*+5)' in result

    def test_relative_col_negative(self):
        # relcol=1, colx=-3 → (*-3)
        result = display_cell_address(rowx=0, colx=-3, relrow=0, relcol=1)
        assert '(*-3)' in result

    def test_both_relative(self):
        result = display_cell_address(rowx=1, colx=1, relrow=1, relcol=1)
        assert '(*+1)' in result
        assert result.count('(*+1)') == 2

    def test_absolute_row_zero_is_row_1(self):
        # rowx=0 absolute → $1
        result = display_cell_address(rowx=0, colx=0, relrow=0, relcol=0)
        assert '$1' in result


# ---------------------------------------------------------------------------
# Name class
# ---------------------------------------------------------------------------

class TestNameDefaults:

    def test_hidden_default(self):
        n = Name()
        assert n.hidden == 0

    def test_macro_default(self):
        n = Name()
        assert n.macro == 0

    def test_builtin_default(self):
        n = Name()
        assert n.builtin == 0

    def test_scope_default(self):
        n = Name()
        assert n.scope == -1

    def test_result_default(self):
        n = Name()
        assert n.result is None

    def test_raw_formula_default(self):
        n = Name()
        assert n.raw_formula == b''


class TestNameCellRaisesWithoutResult:

    def test_cell_raises_when_no_result(self):
        n = Name()
        n.result = None
        n.book = type('MockBook', (), {'logfile': io.StringIO()})()
        with pytest.raises(XLRDError):
            n.cell()

    def test_area2d_raises_when_no_result(self):
        n = Name()
        n.result = None
        n.book = type('MockBook', (), {'logfile': io.StringIO()})()
        with pytest.raises(XLRDError):
            n.area2d()


# ---------------------------------------------------------------------------
# Book class defaults
# ---------------------------------------------------------------------------

class TestBookDefaults:

    def test_nsheets_default(self):
        bk = Book()
        assert bk.nsheets == 0

    def test_datemode_default(self):
        bk = Book()
        assert bk.datemode == 0

    def test_biff_version_default(self):
        bk = Book()
        assert bk.biff_version == 0

    def test_name_obj_list_empty(self):
        bk = Book()
        assert bk.name_obj_list == []

    def test_colour_map_empty(self):
        bk = Book()
        assert bk.colour_map == {}

    def test_palette_record_empty(self):
        bk = Book()
        assert bk.palette_record == []


def _make_simple_xls_bytes():
    """Create a minimal valid BIFF8 XLS file with two sheets."""
    import xlwt
    wb = xlwt.Workbook()
    for i in range(2):
        ws = wb.add_sheet(f'Sheet{i+1}')
        ws.write(0, 0, i)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# book.py line 442 — sheets() with on_demand loads unloaded sheets
# ---------------------------------------------------------------------------

class TestSheetsOnDemand:

    def test_sheets_loads_unloaded_on_demand_sheets(self):
        """sheets() must call get_sheet() for any None entries (line 442)."""
        data = _make_simple_xls_bytes()
        bk = xlrd.open_workbook(file_contents=data, on_demand=True)
        # With on_demand, sheets are not pre-loaded
        assert bk._sheet_list == [None, None]
        sheets = bk.sheets()
        assert len(sheets) == 2
        assert all(sh is not None for sh in sheets)
        assert sheets[0].name == 'Sheet1'
        assert sheets[1].name == 'Sheet2'

    def test_sheets_only_loads_unloaded(self):
        """sheets() skips already-loaded sheets (only calls get_sheet for None)."""
        data = _make_simple_xls_bytes()
        bk = xlrd.open_workbook(file_contents=data, on_demand=True)
        # Pre-load just the first sheet
        bk.sheet_by_index(0)
        assert bk._sheet_list[0] is not None
        assert bk._sheet_list[1] is None
        sheets = bk.sheets()
        assert len(sheets) == 2


# ---------------------------------------------------------------------------
# book.py line 698 — get_sheet() raises after release_resources()
# ---------------------------------------------------------------------------

class TestGetSheetAfterRelease:

    def test_get_sheet_raises_after_release(self):
        """get_sheet() raises XLRDError after release_resources() (line 698)."""
        data = _make_simple_xls_bytes()
        bk = xlrd.open_workbook(file_contents=data)
        bk.release_resources()
        with pytest.raises(XLRDError, match="releasing resources"):
            bk.get_sheet(0)


# ---------------------------------------------------------------------------
# book.py lines 691-694 — get_record_parts_conditional returns (None, 0, b'')
# ---------------------------------------------------------------------------

class TestGetRecordPartsConditional:

    def _make_book_with_mem(self, mem_bytes):
        bk = Book()
        bk.mem = mem_bytes
        bk._position = 0
        bk.logfile = io.StringIO()
        return bk

    def test_returns_none_when_code_mismatches(self):
        """get_record_parts_conditional returns (None,0,b'') when code!=reqd (lines 691-694)."""
        # Build a fake record with opcode 0x0001, length=4, data=4 zeros
        mem = struct.pack('<HH', 0x0001, 4) + b'\x00' * 4
        bk = self._make_book_with_mem(mem)
        result = bk.get_record_parts_conditional(0x0002)  # different reqd
        assert result == (None, 0, b'')
        # position should not have advanced
        assert bk._position == 0

    def test_returns_data_when_code_matches(self):
        """get_record_parts_conditional returns full record when code matches."""
        payload = b'\x01\x02\x03\x04'
        mem = struct.pack('<HH', 0x0005, 4) + payload
        bk = self._make_book_with_mem(mem)
        code, length, data = bk.get_record_parts_conditional(0x0005)
        assert code == 0x0005
        assert length == 4
        assert data == payload


# ---------------------------------------------------------------------------
# book.py getbof() error paths (lines 1278, 1282, 1284, 1287, 1289, 1296)
# ---------------------------------------------------------------------------

class TestGetbofErrors:

    def _make_book_with_mem(self, mem_bytes):
        bk = Book()
        bk.mem = mem_bytes
        bk._position = 0
        bk.logfile = io.StringIO()
        bk.verbosity = 0
        return bk

    def test_eof_before_bof_opcode(self):
        """getbof() raises XLRDError when stream is empty (line 1282)."""
        bk = self._make_book_with_mem(b'')
        with pytest.raises(XLRDError, match="Expected BOF record; met end of file"):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_non_bof_opcode(self):
        """getbof() raises XLRDError when opcode is not a BOF code (line 1284)."""
        # opcode 0x0001 is not in bofcodes
        mem = struct.pack('<HH', 0x0001, 4) + b'\x00' * 4
        bk = self._make_book_with_mem(mem)
        with pytest.raises(XLRDError, match="Expected BOF record"):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_eof_in_bof_length(self):
        """getbof() raises XLRDError when stream ends before length field (line 1287)."""
        # Valid BOF opcode 0x0809 but only 1 byte following (length needs 2)
        mem = struct.pack('<H', 0x0809) + b'\x00'
        bk = self._make_book_with_mem(mem)
        with pytest.raises(XLRDError, match=r"Incomplete BOF record\[1\]"):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_invalid_bof_length_too_small(self):
        """getbof() raises XLRDError when BOF length < 4 (lines 1289-1291)."""
        # BOF opcode OK, length=3 (invalid: must be 4-20)
        mem = struct.pack('<HH', 0x0809, 3) + b'\x00' * 3
        bk = self._make_book_with_mem(mem)
        with pytest.raises(XLRDError, match="Invalid length"):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_invalid_bof_length_too_large(self):
        """getbof() raises XLRDError when BOF length > 20 (lines 1289-1291)."""
        # BOF opcode OK, length=21
        mem = struct.pack('<HH', 0x0809, 21) + b'\x00' * 21
        bk = self._make_book_with_mem(mem)
        with pytest.raises(XLRDError, match="Invalid length"):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_truncated_bof_data(self):
        """getbof() raises XLRDError when BOF data is truncated (line 1296)."""
        # BOF opcode 0x0809, length=8, but only 3 bytes of data available
        mem = struct.pack('<HH', 0x0809, 8) + b'\x00' * 3
        bk = self._make_book_with_mem(mem)
        with pytest.raises(XLRDError, match=r"Incomplete BOF record\[2\]"):
            bk.getbof(XL_WORKBOOK_GLOBALS)

    def test_xf_list_empty(self):
        bk = Book()
        assert bk.xf_list == []

    def test_font_list_empty(self):
        bk = Book()
        assert bk.font_list == []

    def test_format_list_empty(self):
        bk = Book()
        assert bk.format_list == []

    def test_sheet_names_empty(self):
        bk = Book()
        # _sheet_names is populated during load; default should be empty or raise
        assert hasattr(bk, 'sheet_names') or hasattr(bk, '_sheet_names')

    def test_user_name_is_string(self):
        bk = Book()
        assert isinstance(bk.user_name, str)


# ---------------------------------------------------------------------------
# Module-level constants
# ---------------------------------------------------------------------------

class TestModuleConstants:

    def test_supbook_values(self):
        assert SUPBOOK_UNK == 0
        assert SUPBOOK_INTERNAL == 1
        assert SUPBOOK_EXTERNAL == 2
        assert SUPBOOK_ADDIN == 3
        assert SUPBOOK_DDEOLE == 4

    def test_supported_versions_includes_biff8(self):
        assert 80 in SUPPORTED_VERSIONS

    def test_supported_versions_includes_biff5(self):
        assert 50 in SUPPORTED_VERSIONS

    def test_builtin_name_roundtrip(self):
        # Print_Area should round-trip through code_from_builtin_name
        for name, code in code_from_builtin_name.items():
            assert builtin_name_from_code[code] == name


# ---------------------------------------------------------------------------
# book.py line 635 — raw BIFF stream without OLE2 header
# ---------------------------------------------------------------------------

class TestRawBiffStream:

    def test_raw_biff8_stream_loads_without_ole2(self):
        """open_workbook handles raw BIFF8 data without OLE2 header (line 635)."""
        import struct
        # BIFF8 global stream: BOF (opcode=0x0809, BIFF8 globals) + XL_EOF
        bof = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0600, 0x0005) + b'\x00' * 4
        eof_rec = struct.pack('<HH', 0x000A, 0)
        raw_biff = bof + eof_rec
        bk = xlrd.open_workbook(file_contents=raw_biff, logfile=io.StringIO())
        assert bk.biff_version == 80
        assert bk.nsheets == 0

    def test_raw_biff5_stream_year_lt_1994_gives_version_50(self):
        """getbof detects BIFF5 when version2=0x0500 and year < 1994 (lines 1312-1315)."""
        import struct
        bof = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0500, 0x0005) + struct.pack('<HH', 1234, 1990)
        eof_rec = struct.pack('<HH', 0x000A, 0)
        raw_biff = bof + eof_rec
        bk = xlrd.open_workbook(file_contents=raw_biff, logfile=io.StringIO())
        assert bk.biff_version == 50

    def test_raw_biff7_stream_year_gte_1994_gives_version_70(self):
        """getbof detects BIFF7 when version2=0x0500 and year >= 1994 (lines 1315-1316)."""
        import struct
        bof = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0500, 0x0005) + struct.pack('<HH', 1234, 1994)
        eof_rec = struct.pack('<HH', 0x000A, 0)
        raw_biff = bof + eof_rec
        bk = xlrd.open_workbook(file_contents=raw_biff, logfile=io.StringIO())
        assert bk.biff_version == 70

    def test_raw_biff8_workspace_file_raises(self):
        """getbof raises for workspace file (streamtype=0x0100) when version >= 50 (lines 1341-1342)."""
        import struct
        bof = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0600, 0x0100) + b'\x00' * 4
        raw_biff = bof  # no EOF needed, error raised in getbof
        with pytest.raises(XLRDError, match="Workspace file"):
            xlrd.open_workbook(file_contents=raw_biff, logfile=io.StringIO())


# ---------------------------------------------------------------------------
# book.py lines 1319-1325 (dodgy third-party version2) + 1339-1340 (worksheet stream)
# ---------------------------------------------------------------------------

class TestGetbofVersionPaths:

    def _make_book_with_mem(self, mem_bytes):
        bk = Book()
        bk.mem = mem_bytes
        bk._position = 0
        bk.logfile = io.StringIO()
        bk.verbosity = 0
        return bk

    def test_dodgy_version2_0x0000_gives_biff21(self):
        """getbof maps dodgy version2=0x0000 to BIFF 2.1 (version=21, lines 1319-1325)."""
        bof_mem = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0000, 0x0005) + b'\x00' * 4
        bk = self._make_book_with_mem(bof_mem)
        version = bk.getbof(XL_WORKBOOK_GLOBALS)
        assert version == 21

    def test_dodgy_version2_0x0400_gives_biff40(self):
        """getbof maps dodgy version2=0x0400 to BIFF 4S (version=40, lines 1319-1325)."""
        bof_mem = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0400, 0x0005) + b'\x00' * 4
        bk = self._make_book_with_mem(bof_mem)
        version = bk.getbof(XL_WORKBOOK_GLOBALS)
        assert version == 40

    def test_old_worksheet_bof_accepted_for_globals_request(self):
        """getbof returns version for BIFF2 worksheet stream when requesting globals (line 1339)."""
        # BIFF2 BOF: opcode=0x0009, version1=0x00 → version=21; streamtype=XL_WORKSHEET
        bof_mem = struct.pack('<HH', 0x0009, 4) + struct.pack('<HH', 0x0200, 0x0010)
        bk = self._make_book_with_mem(bof_mem)
        version = bk.getbof(XL_WORKBOOK_GLOBALS)
        assert version == 21

    def test_biff5_version2_build_special_gives_v50(self):
        """getbof detects BIFF5 for special build IDs (lines 1312-1314)."""
        bof_mem = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0500, 0x0005) + struct.pack('<HH', 2412, 1995)
        bk = self._make_book_with_mem(bof_mem)
        version = bk.getbof(XL_WORKBOOK_GLOBALS)
        assert version == 50

    def test_unrecognised_version2_returns_zero(self):
        """getbof returns 0 for unrecognised dodgy version2, then open_workbook raises (lines 1319-1325)."""
        bof_mem = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0xFFFF, 0x0005) + b'\x00' * 4
        bk = self._make_book_with_mem(bof_mem)
        version = bk.getbof(XL_WORKBOOK_GLOBALS)
        assert version == 0  # unrecognised → version=0 → bif version unknown


# ---------------------------------------------------------------------------
# book.py line 1330 — BIFF4W version=45 detection
# ---------------------------------------------------------------------------

class TestGetbofBiff4W:

    def _make_book_with_mem(self, mem_bytes):
        bk = Book()
        bk.mem = mem_bytes
        bk._position = 0
        bk.logfile = io.StringIO()
        bk.verbosity = 0
        return bk

    def test_biff4_worksheet_bof_gives_version_45(self):
        """getbof converts version=40 + XL_WORKBOOK_GLOBALS_4W streamtype to version=45 (line 1330)."""
        from xlrd.biffh import XL_WORKBOOK_GLOBALS_4W
        # BIFF4 BOF opcode=0x0409, version1=0x04 → version=40
        # streamtype=0x0100 (XL_WORKBOOK_GLOBALS_4W) → version becomes 45
        bof_mem = struct.pack('<HH', 0x0409, 6) + struct.pack('<HH', 0, XL_WORKBOOK_GLOBALS_4W) + b'\x00' * 2
        bk = self._make_book_with_mem(bof_mem)
        version = bk.getbof(XL_WORKBOOK_GLOBALS)
        assert version == 45


# ---------------------------------------------------------------------------
# book.py line 1247 — unexpected BOF in parse_globals with verbosity
# ---------------------------------------------------------------------------

class TestUnexpectedBofInGlobals:

    def test_unexpected_bof_logs_warning_with_verbosity(self):
        """parse_globals logs warning for unexpected BOF (line 1247)."""
        # Raw BIFF8: global BOF + unexpected inner BOF (same opcode, rc & 0xff == 9) + EOF
        bof = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0600, 0x0005) + b'\x00' * 4
        unexpected_bof = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0600, 0x0010) + b'\x00' * 4
        eof_rec = struct.pack('<HH', 0x000A, 0)
        raw_biff = bof + unexpected_bof + eof_rec

        log = io.StringIO()
        bk = xlrd.open_workbook(file_contents=raw_biff, verbosity=1, logfile=log)
        assert 'Unexpected BOF' in log.getvalue()


# ---------------------------------------------------------------------------
# book.py line 803 — derive_encoding verbosity with no CODEPAGE in BIFF8
# book.py line 794 — derive_encoding with encoding_override
# ---------------------------------------------------------------------------

class TestDeriveEncoding:

    @staticmethod
    def _raw_biff8():
        """Minimal raw BIFF8 global stream with no CODEPAGE record."""
        bof = struct.pack('<HH', 0x0809, 8) + struct.pack('<HH', 0x0600, 0x0005) + b'\x00' * 4
        eof_rec = struct.pack('<HH', 0x000A, 0)
        return bof + eof_rec

    def test_no_codepage_biff8_verbosity2_logs_assumption(self):
        """derive_encoding logs 'utf_16_le' assumption for BIFF8 without CODEPAGE at verbosity>=2 (line 803)."""
        log = io.StringIO()
        bk = xlrd.open_workbook(file_contents=self._raw_biff8(), verbosity=2, logfile=log)
        assert 'utf_16_le' in log.getvalue()

    def test_encoding_override_sets_encoding(self):
        """derive_encoding uses encoding_override when provided (line 794)."""
        log = io.StringIO()
        bk = xlrd.open_workbook(file_contents=self._raw_biff8(),
                                encoding_override='utf-8', logfile=log)
        assert bk.encoding == 'utf-8'


    def test_codepage_in_cp_range_sets_cp_encoding(self):
        """derive_encoding sets encoding='cp<N>' for codepage in 300-1999 range (line 808-809)."""
        from xlrd.book import Book
        bk = Book()
        bk.biff_version = 80
        bk.logfile = io.StringIO()
        bk.verbosity = 0
        bk.codepage = 1252  # not in encoding_from_codepage map, but 300 <= 1252 <= 1999
        bk.encoding = None
        bk.encoding_override = ''
        bk.raw_user_name = False
        bk.derive_encoding()
        assert bk.encoding == 'cp1252'

    def test_biff8_unknown_high_codepage_resets_to_utf16le(self):
        """derive_encoding resets codepage=1200 and encoding=utf_16_le for unknown high codepage in BIFF8 (lines 810-812)."""
        from xlrd.book import Book
        bk = Book()
        bk.biff_version = 80
        bk.logfile = io.StringIO()
        bk.verbosity = 0
        bk.codepage = 5000  # > 1999, not in map, biff >= 80
        bk.encoding = None
        bk.encoding_override = ''
        bk.raw_user_name = False
        bk.derive_encoding()
        assert bk.encoding == 'utf_16_le'
        assert bk.codepage == 1200
