"""Tests for Book.names_epilogue targeting uncovered lines."""
import io

import pytest

from xlrd.book import Book, Name


def _make_book(biff_version=80, verbosity=0):
    """Create a Book with minimum state needed for names_epilogue."""
    bk = Book()
    bk.logfile = io.StringIO()
    bk.verbosity = verbosity
    bk.biff_version = biff_version
    bk._all_sheets_map = []
    bk._extnsht_name_from_num = {}
    bk._sheet_num_from_name = {}
    bk.name_obj_list = []
    return bk


def _make_name(name='test_name', excel_sheet_index=0, extn_sheet_num=0,
               macro=0, binary=0, evaluated=0, raw_formula=b'',
               basic_formula_len=0):
    """Create a Name object with the given attributes."""
    nobj = Name()
    nobj.name = name
    nobj.excel_sheet_index = excel_sheet_index
    nobj.extn_sheet_num = extn_sheet_num
    nobj.macro = macro
    nobj.binary = binary
    nobj.evaluated = evaluated
    nobj.raw_formula = raw_formula
    nobj.basic_formula_len = basic_formula_len
    nobj.scope = None
    nobj.result = None
    return nobj


# ---------------------------------------------------------------------------
# Verbose / debug output (lines 1009-1012, 1047-1052)
# ---------------------------------------------------------------------------

class TestNamesEpilogueVerbose:

    def test_verbose_prints_debug_info(self):
        bk = _make_book(biff_version=80, verbosity=2)
        bk.name_obj_list = []
        bk.names_epilogue()
        output = bk.logfile.getvalue()
        assert 'names_epilogue' in output
        assert '_all_sheets_map' in output
        assert '_extnsht_name_from_num' in output
        assert '_sheet_num_from_name' in output

    def test_verbose_dumps_name_objects(self):
        bk = _make_book(biff_version=80, verbosity=2)
        nobj = _make_name('myname', macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        output = bk.logfile.getvalue()
        assert 'name object dump' in output

    def test_non_verbose_produces_no_debug_output(self):
        bk = _make_book(biff_version=80, verbosity=0)
        bk.name_obj_list = []
        bk.names_epilogue()
        output = bk.logfile.getvalue()
        assert 'names_epilogue' not in output


# ---------------------------------------------------------------------------
# BIFF8 scope assignment (lines 1020-1030)
# ---------------------------------------------------------------------------

class TestNamesEpilogueBiff8Scope:

    def test_biff8_global_scope_when_sheet_index_zero(self):
        bk = _make_book(biff_version=80)
        nobj = _make_name('global_name', excel_sheet_index=0, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.scope == -1

    def test_biff8_local_scope_when_valid_sheet_index(self):
        bk = _make_book(biff_version=80)
        bk._all_sheets_map = [2]
        nobj = _make_name('local_name', excel_sheet_index=1, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.scope == 2

    def test_biff8_macro_sheet_maps_to_minus_two(self):
        bk = _make_book(biff_version=80)
        bk._all_sheets_map = [-1]
        nobj = _make_name('macro_sheet_name', excel_sheet_index=1, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.scope == -2

    def test_biff8_invalid_sheet_index_scope_minus_three(self):
        bk = _make_book(biff_version=80)
        bk._all_sheets_map = [0]
        nobj = _make_name('bad_name', excel_sheet_index=99, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.scope == -3


# ---------------------------------------------------------------------------
# BIFF5/7 scope assignment (lines 1031-1037)
# ---------------------------------------------------------------------------

class TestNamesEpilogueBiff7Scope:

    def test_biff7_global_scope_when_extn_sheet_zero(self):
        bk = _make_book(biff_version=70)
        nobj = _make_name('global_name', extn_sheet_num=0, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.scope == -1

    def test_biff7_local_scope_by_name_lookup(self):
        bk = _make_book(biff_version=70)
        bk._extnsht_name_from_num = {1: 'Sheet1'}
        bk._sheet_num_from_name = {'Sheet1': 0}
        nobj = _make_name('local_name', extn_sheet_num=1, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.scope == 0

    def test_biff7_unknown_sheet_name_scope_minus_two(self):
        bk = _make_book(biff_version=70)
        bk._extnsht_name_from_num = {1: 'UnknownSheet'}
        bk._sheet_num_from_name = {}
        nobj = _make_name('unknown_name', extn_sheet_num=1, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.scope == -2

    def test_biff5_version_uses_extn_sheet(self):
        bk = _make_book(biff_version=50)
        bk._extnsht_name_from_num = {1: 'Sheet1'}
        bk._sheet_num_from_name = {'Sheet1': 3}
        nobj = _make_name('biff5_name', extn_sheet_num=1, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.scope == 3


# ---------------------------------------------------------------------------
# Formula evaluation loop (lines 1040-1045)
# ---------------------------------------------------------------------------

class TestNamesEpilogueFormulaEval:

    def test_macro_name_skips_formula_evaluation(self):
        bk = _make_book(biff_version=80)
        nobj = _make_name('macro_name', macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.result is None

    def test_binary_name_skips_formula_evaluation(self):
        bk = _make_book(biff_version=80)
        nobj = _make_name('binary_name', binary=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.result is None

    def test_already_evaluated_name_skips_formula_evaluation(self):
        bk = _make_book(biff_version=80)
        nobj = _make_name('evaluated_name', evaluated=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.result is None

    def test_normal_name_gets_formula_evaluated(self):
        bk = _make_book(biff_version=80)
        nobj = _make_name('normal_name', evaluated=0, macro=0, binary=0,
                          basic_formula_len=0, raw_formula=b'')
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert nobj.evaluated == 1


# ---------------------------------------------------------------------------
# name_and_scope_map and name_map building (lines 1056-1077)
# ---------------------------------------------------------------------------

class TestNamesEpilogueNameMaps:

    def test_empty_name_list_produces_empty_maps(self):
        bk = _make_book(biff_version=80)
        bk.name_obj_list = []
        bk.names_epilogue()
        assert bk.name_and_scope_map == {}
        assert bk.name_map == {}

    def test_builds_name_and_scope_map(self):
        bk = _make_book(biff_version=80)
        nobj = _make_name('myname', excel_sheet_index=0, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert ('myname', -1) in bk.name_and_scope_map
        assert bk.name_and_scope_map[('myname', -1)] is nobj

    def test_builds_name_map(self):
        bk = _make_book(biff_version=80)
        nobj = _make_name('myname', excel_sheet_index=0, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert 'myname' in bk.name_map
        assert bk.name_map['myname'] == [nobj]

    def test_name_map_key_is_lowercase(self):
        bk = _make_book(biff_version=80)
        nobj = _make_name('MyName', excel_sheet_index=0, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert 'myname' in bk.name_map
        assert 'MyName' not in bk.name_map

    def test_multiple_names_same_name_appended_to_name_map(self):
        bk = _make_book(biff_version=80)
        bk._all_sheets_map = [0, 1]
        nobj1 = _make_name('repeated', excel_sheet_index=1, macro=1)
        nobj2 = _make_name('repeated', excel_sheet_index=2, macro=1)
        bk.name_obj_list = [nobj1, nobj2]
        bk.names_epilogue()
        assert 'repeated' in bk.name_map
        assert len(bk.name_map['repeated']) == 2

    def test_name_map_sorted_by_scope(self):
        bk = _make_book(biff_version=80)
        bk._all_sheets_map = [0, 1]
        nobj1 = _make_name('sorted_name', excel_sheet_index=1, macro=1)
        nobj2 = _make_name('sorted_name', excel_sheet_index=0, macro=1)
        bk.name_obj_list = [nobj1, nobj2]
        bk.names_epilogue()
        result_list = bk.name_map['sorted_name']
        assert result_list[0].scope <= result_list[1].scope

    def test_duplicate_key_warning_with_verbosity(self):
        bk = _make_book(biff_version=80, verbosity=1)
        nobj1 = _make_name('dup', excel_sheet_index=0, macro=1)
        nobj2 = _make_name('dup', excel_sheet_index=0, macro=1)
        bk.name_obj_list = [nobj1, nobj2]
        bk.names_epilogue()
        output = bk.logfile.getvalue()
        assert 'Duplicate' in output

    def test_final_maps_assigned_to_book(self):
        bk = _make_book(biff_version=80)
        nobj = _make_name('aname', excel_sheet_index=0, macro=1)
        bk.name_obj_list = [nobj]
        bk.names_epilogue()
        assert hasattr(bk, 'name_and_scope_map')
        assert hasattr(bk, 'name_map')
        assert isinstance(bk.name_and_scope_map, dict)
        assert isinstance(bk.name_map, dict)
