"""
Tests for xlrd/book.py open_workbook_xls targeting uncovered lines.

Uncovered lines targeted:
  81  - raise XLRDError("Can't determine file's BIFF version")
  83  - raise XLRDError("BIFF version %s is not supported" ...)
  90  - if on_demand: (biff_version <= 40)
  91  - fprintf warning (biff_version <= 40, on_demand=True)
  94  - bk.on_demand = on_demand = False (biff_version <= 40)
  95  - bk.fake_globals_get_sheet()
  98  - bk.parse_globals() (biff_version == 45)
  99  - if on_demand: (biff_version == 45)
 100  - fprintf warning (biff_version == 45, on_demand=True)
 102  - bk.on_demand = on_demand = False (biff_version == 45)
 110  - fprintf warning (biff_version == 45, nsheets > 1)
"""
import io
import pytest

from xlrd.book import Book, open_workbook_xls
from xlrd.biffh import XLRDError

from tests.test_book import make_minimal_xls


# ---------------------------------------------------------------------------
# Tests for open_workbook_xls uncovered lines
# ---------------------------------------------------------------------------

class TestOpenWorkbookXlsUncoveredLines:

    def test_getbof_returns_zero_raises_xlrderror(self, mocker):
        """Line 81: getbof returning 0 (falsy) raises XLRDError."""
        mocker.patch.object(Book, 'getbof', return_value=0)
        with pytest.raises(XLRDError, match="Can't determine file's BIFF version"):
            open_workbook_xls(file_contents=make_minimal_xls())

    def test_unsupported_biff_version_raises_xlrderror(self, mocker):
        """Line 83: unsupported biff_version raises XLRDError mentioning the version."""
        # 85 is in biff_text_from_num but not in SUPPORTED_VERSIONS
        mocker.patch.object(Book, 'getbof', return_value=85)
        with pytest.raises(XLRDError, match="BIFF version"):
            open_workbook_xls(file_contents=make_minimal_xls())

    def test_biff_version_le40_no_on_demand(self, mocker):
        """Line 95: biff_version <= 40 without on_demand calls fake_globals_get_sheet."""
        mocker.patch.object(Book, 'getbof', return_value=40)

        def _fake_globals(self):
            self._sheet_list = [None]
            self._sheet_names = ['Sheet 1']

        mocker.patch.object(Book, 'fake_globals_get_sheet', _fake_globals)
        mocker.patch.object(Book, 'get_sheets', lambda self: None)

        bk = open_workbook_xls(file_contents=make_minimal_xls(), on_demand=False)
        assert bk.nsheets == 1

    def test_biff_version_le40_with_on_demand_warns_and_disables(self, mocker):
        """Lines 90, 91, 94, 95: biff_version <= 40 with on_demand=True warns and sets on_demand=False."""
        mocker.patch.object(Book, 'getbof', return_value=40)

        def _fake_globals(self):
            self._sheet_list = [None]
            self._sheet_names = ['Sheet 1']

        mocker.patch.object(Book, 'fake_globals_get_sheet', _fake_globals)
        mocker.patch.object(Book, 'get_sheets', lambda self: None)

        logfile = io.StringIO()
        bk = open_workbook_xls(
            file_contents=make_minimal_xls(),
            on_demand=True,
            logfile=logfile,
        )
        warning = logfile.getvalue()
        assert 'on_demand' in warning
        assert bk.on_demand is False

    def test_biff_version_45_parse_globals_called(self, mocker):
        """Line 98: biff_version == 45 calls parse_globals."""
        mocker.patch.object(Book, 'getbof', return_value=45)

        def _parse_globals(self):
            self._sheet_list = [None]
            self._sheet_names = ['Sheet1']

        mocker.patch.object(Book, 'parse_globals', _parse_globals)

        bk = open_workbook_xls(file_contents=make_minimal_xls(), on_demand=False)
        assert bk.nsheets == 1

    def test_biff_version_45_with_on_demand_warns_and_disables(self, mocker):
        """Lines 98, 99, 100, 102: biff_version == 45 with on_demand=True warns and disables it."""
        mocker.patch.object(Book, 'getbof', return_value=45)

        def _parse_globals(self):
            self._sheet_list = [None]
            self._sheet_names = ['Sheet1']

        mocker.patch.object(Book, 'parse_globals', _parse_globals)

        logfile = io.StringIO()
        bk = open_workbook_xls(
            file_contents=make_minimal_xls(),
            on_demand=True,
            logfile=logfile,
        )
        warning = logfile.getvalue()
        assert 'on_demand' in warning
        assert bk.on_demand is False

    def test_biff_version_45_multiple_sheets_warns(self, mocker):
        """Line 110: biff_version == 45 with nsheets > 1 emits a warning."""
        mocker.patch.object(Book, 'getbof', return_value=45)

        def _parse_globals(self):
            self._sheet_list = [None, None]
            self._sheet_names = ['Sheet1', 'Sheet2']

        mocker.patch.object(Book, 'parse_globals', _parse_globals)

        logfile = io.StringIO()
        bk = open_workbook_xls(
            file_contents=make_minimal_xls(),
            on_demand=False,
            logfile=logfile,
        )
        warning = logfile.getvalue()
        assert bk.nsheets == 2
        assert 'WARNING' in warning
