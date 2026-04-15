import io
import sys
import pytest

from xlrd.timemachine import fprintf


def test_fprintf_with_newline():
    f = io.StringIO()
    fprintf(f, "hello %s\n", "world")
    assert f.getvalue() == "hello world\n"


def test_fprintf_without_newline():
    f = io.StringIO()
    fprintf(f, "value: %d", 42)
    assert f.getvalue() == "value: 42 "


def test_fprintf_replaces_percent_r_with_percent_a():
    f = io.StringIO()
    fprintf(f, "%r\n", "test")
    assert f.getvalue() == "'test'\n"


def test_fprintf_no_args():
    f = io.StringIO()
    fprintf(f, "no args\n")
    assert f.getvalue() == "no args\n"


@pytest.mark.skip(reason="Lines 40-44 are in the Python 2 branch and unreachable in Python 3")
def test_fprintf_python2_branch():
    pass
