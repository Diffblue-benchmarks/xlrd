# -*- coding: utf-8 -*-
"""
Tests for xlrd.__init__ — inspect_format and open_workbook public API.
"""
import io
import zipfile
import pytest

import xlrd
from xlrd import inspect_format
from xlrd.biffh import XLRDError
from xlrd.compdoc import SIGNATURE as XLS_SIGNATURE


ZIP_SIGNATURE = b"PK\x03\x04"


def _make_zip_bytes(names):
    """Return in-memory zip file bytes containing empty files with given names."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for name in names:
            zf.writestr(name, "")
    return buf.getvalue()


class TestInspectFormat:

    def test_xls_signature_from_content(self):
        content = XLS_SIGNATURE + b"\x00" * 512
        assert inspect_format(content=content) == "xls"

    def test_xlsx_detected_from_content(self):
        content = _make_zip_bytes(["xl/workbook.xml"])
        assert inspect_format(content=content) == "xlsx"

    def test_xlsb_detected_from_content(self):
        content = _make_zip_bytes(["xl/workbook.bin"])
        assert inspect_format(content=content) == "xlsb"

    def test_ods_detected_from_content(self):
        content = _make_zip_bytes(["content.xml"])
        assert inspect_format(content=content) == "ods"

    def test_unknown_zip_returns_zip(self):
        content = _make_zip_bytes(["something_else.txt"])
        assert inspect_format(content=content) == "zip"

    def test_unknown_content_returns_none(self):
        content = b"\x00\x01\x02\x03\x04\x05\x06\x07\x08"
        assert inspect_format(content=content) is None

    def test_xlsx_case_insensitive_component_names(self):
        # Some third-party tools use lowercase or backslash paths
        content = _make_zip_bytes(["XL\\Workbook.XML"])
        assert inspect_format(content=content) == "xlsx"

    def test_from_file(self, tmp_path):
        xls_file = tmp_path / "test.xls"
        xls_file.write_bytes(XLS_SIGNATURE + b"\x00" * 512)
        assert inspect_format(path=str(xls_file)) == "xls"

    def test_from_file_unknown(self, tmp_path):
        unknown_file = tmp_path / "test.bin"
        unknown_file.write_bytes(b"\x00\x01\x02\x03\x04\x05\x06\x07\x08")
        assert inspect_format(path=str(unknown_file)) is None


class TestOpenWorkbook:

    def test_open_xlsx_raises_xlrderror(self):
        content = _make_zip_bytes(["xl/workbook.xml"])
        with pytest.raises(XLRDError, match="xlsx"):
            xlrd.open_workbook(file_contents=content)

    def test_open_xlsb_raises_xlrderror(self):
        content = _make_zip_bytes(["xl/workbook.bin"])
        with pytest.raises(XLRDError, match="xlsb"):
            xlrd.open_workbook(file_contents=content)

    def test_open_ods_raises_xlrderror(self):
        content = _make_zip_bytes(["content.xml"])
        with pytest.raises(XLRDError, match="not supported"):
            xlrd.open_workbook(file_contents=content)

    def test_open_zip_raises_xlrderror(self):
        content = _make_zip_bytes(["random_file.txt"])
        with pytest.raises(XLRDError):
            xlrd.open_workbook(file_contents=content)
