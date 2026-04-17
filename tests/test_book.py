# -*- coding: utf-8 -*-
"""
Tests for xlrd.book — Book-level helpers: colname, display_cell_address.
"""
import pytest

from xlrd.book import colname


class TestColname:

    def test_first_column_is_A(self):
        assert colname(0) == "A"

    def test_second_column_is_B(self):
        assert colname(1) == "B"

    def test_last_single_letter_is_Z(self):
        assert colname(25) == "Z"

    def test_first_double_letter_is_AA(self):
        assert colname(26) == "AA"

    def test_AZ(self):
        assert colname(51) == "AZ"

    def test_BA(self):
        assert colname(52) == "BA"

    def test_ZZ(self):
        assert colname(701) == "ZZ"

    def test_AAA(self):
        assert colname(702) == "AAA"

    def test_column_255_is_IV(self):
        # Excel 97-2003 max column is IV (255, 0-indexed)
        assert colname(255) == "IV"

    def test_column_256_is_IW(self):
        assert colname(256) == "IW"

    def test_negative_column_raises(self):
        with pytest.raises(AssertionError):
            colname(-1)
