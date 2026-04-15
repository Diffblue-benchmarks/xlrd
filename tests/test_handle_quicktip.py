# -*- coding: utf-8 -*-
"""Unit tests for Sheet.handle_quicktip."""

import sys
import pytest
from struct import pack

from xlrd.sheet import Sheet
from xlrd.biffh import BaseObject

XL_QUICKTIP = 0x0800


class MockBook:
    biff_version = 80
    logfile = sys.stdout
    verbosity = 0
    formatting_info = False
    ragged_rows = False
    _xf_index_to_xl_type_map = {}
    _sheet_visibility = [0, 0, 0, 0, 0]


def _make_hyperlink(frowx, lrowx, fcolx, lcolx):
    h = BaseObject()
    h.frowx = frowx
    h.lrowx = lrowx
    h.fcolx = fcolx
    h.lcolx = lcolx
    return h


def _make_quicktip_data(frowx, lrowx, fcolx, lcolx, tip_text):
    header = pack('<5H', XL_QUICKTIP, frowx, lrowx, fcolx, lcolx)
    encoded = tip_text.encode('utf_16_le')
    return header + encoded + b'\x00\x00'


@pytest.fixture
def sheet():
    return Sheet(MockBook(), 0, 'TestSheet', 0)


class TestHandleQuicktip:
    def test_quicktip_sets_tip_on_hyperlink(self, sheet):
        h = _make_hyperlink(0, 0, 0, 0)
        sheet.hyperlink_list.append(h)
        data = _make_quicktip_data(0, 0, 0, 0, 'My Tooltip')

        sheet.handle_quicktip(data)

        assert h.quicktip == 'My Tooltip'

    def test_quicktip_with_multirow_multicol_hyperlink(self, sheet):
        h = _make_hyperlink(1, 3, 2, 5)
        sheet.hyperlink_list.append(h)
        data = _make_quicktip_data(1, 3, 2, 5, 'Another tip')

        sheet.handle_quicktip(data)

        assert h.quicktip == 'Another tip'

    def test_quicktip_updates_last_hyperlink(self, sheet):
        h1 = _make_hyperlink(0, 0, 0, 0)
        h2 = _make_hyperlink(1, 1, 1, 1)
        sheet.hyperlink_list.append(h1)
        sheet.hyperlink_list.append(h2)
        data = _make_quicktip_data(1, 1, 1, 1, 'Second link tip')

        sheet.handle_quicktip(data)

        assert h2.quicktip == 'Second link tip'
        assert not hasattr(h1, 'quicktip')

    def test_quicktip_empty_string(self, sheet):
        h = _make_hyperlink(0, 0, 0, 0)
        sheet.hyperlink_list.append(h)
        data = _make_quicktip_data(0, 0, 0, 0, '')

        sheet.handle_quicktip(data)

        assert h.quicktip == ''

    def test_quicktip_assertion_fails_wrong_rcx(self, sheet):
        h = _make_hyperlink(0, 0, 0, 0)
        sheet.hyperlink_list.append(h)
        # Use wrong rcx value
        data = pack('<5H', 0x0900, 0, 0, 0, 0) + b'\x00\x00'

        with pytest.raises(AssertionError):
            sheet.handle_quicktip(data)

    def test_quicktip_assertion_fails_empty_hyperlink_list(self, sheet):
        data = _make_quicktip_data(0, 0, 0, 0, 'tip')

        with pytest.raises(AssertionError):
            sheet.handle_quicktip(data)

    def test_quicktip_assertion_fails_coord_mismatch(self, sheet):
        h = _make_hyperlink(0, 0, 0, 0)
        sheet.hyperlink_list.append(h)
        # Use mismatched coords in data
        data = _make_quicktip_data(1, 1, 1, 1, 'tip')

        with pytest.raises(AssertionError):
            sheet.handle_quicktip(data)

    def test_quicktip_assertion_fails_wrong_terminator(self, sheet):
        h = _make_hyperlink(0, 0, 0, 0)
        sheet.hyperlink_list.append(h)
        header = pack('<5H', XL_QUICKTIP, 0, 0, 0, 0)
        data = header + 'tip'.encode('utf_16_le') + b'\x01\x00'

        with pytest.raises(AssertionError):
            sheet.handle_quicktip(data)
