"""
Tests targeting uncovered lines in xlrd/book.py Book.handle_filepass.
"""
import io
import struct
import sys
import pytest

from xlrd.book import Book
from xlrd.biffh import XLRDError


def _make_book(biff_version=80, verbosity=0):
    """Create a minimal Book instance with essential attributes set."""
    bk = Book()
    bk.biff_version = biff_version
    bk.verbosity = verbosity
    bk.logfile = io.StringIO()
    bk.encoding = 'ascii'
    return bk


class TestHandleFilepass:

    def test_raises_xlrd_error_low_verbosity(self):
        """handle_filepass always raises XLRDError regardless of verbosity."""
        bk = _make_book(verbosity=0)
        with pytest.raises(XLRDError, match="encrypted"):
            bk.handle_filepass(b'\x00' * 10)

    def test_raises_xlrd_error_high_verbosity(self):
        """handle_filepass raises XLRDError even when verbosity >= 2."""
        bk = _make_book(verbosity=2)
        data = struct.pack('<H', 0) + struct.pack('<HH', 0xABCD, 0x1234)
        with pytest.raises(XLRDError, match="encrypted"):
            bk.handle_filepass(data)

    def test_verbosity_high_biff8_weak_xor(self):
        """BIFF8, verbosity>=2, kind1==0 (weak XOR): logs key and hash, then raises."""
        bk = _make_book(biff_version=80, verbosity=2)
        # kind1=0 (weak XOR), key=0x1234, hash=0x5678
        data = struct.pack('<H', 0) + struct.pack('<HH', 0x1234, 0x5678)
        with pytest.raises(XLRDError):
            bk.handle_filepass(data)
        log_output = bk.logfile.getvalue()
        assert 'FILEPASS' in log_output

    def test_verbosity_high_biff8_kind1_std(self):
        """BIFF8, verbosity>=2, kind1==1, kind2==1 (BIFF8 standard): logs caption, then raises."""
        bk = _make_book(biff_version=80, verbosity=2)
        # kind1=1, reserved=0, kind2=1
        data = struct.pack('<HHH', 1, 0, 1) + b'\x00' * 10
        with pytest.raises(XLRDError):
            bk.handle_filepass(data)
        log_output = bk.logfile.getvalue()
        assert 'BIFF8 std' in log_output

    def test_verbosity_high_biff8_kind1_strong(self):
        """BIFF8, verbosity>=2, kind1==1, kind2==2 (BIFF8 strong): logs caption, then raises."""
        bk = _make_book(biff_version=80, verbosity=2)
        # kind1=1, reserved=0, kind2=2
        data = struct.pack('<HHH', 1, 0, 2) + b'\x00' * 10
        with pytest.raises(XLRDError):
            bk.handle_filepass(data)
        log_output = bk.logfile.getvalue()
        assert 'BIFF8 strong' in log_output

    def test_verbosity_high_biff8_kind1_unknown(self):
        """BIFF8, verbosity>=2, kind1==1, kind2 unknown: logs unknown caption, then raises."""
        bk = _make_book(biff_version=80, verbosity=2)
        # kind1=1, reserved=0, kind2=99 (unknown)
        data = struct.pack('<HHH', 1, 0, 99) + b'\x00' * 10
        with pytest.raises(XLRDError):
            bk.handle_filepass(data)
        log_output = bk.logfile.getvalue()
        assert 'UNKNOWN' in log_output

    def test_verbosity_high_biff_pre8(self):
        """Pre-BIFF8, verbosity>=2: logs FILEPASS header but skips kind1 check, then raises."""
        bk = _make_book(biff_version=70, verbosity=2)
        data = b'\x00' * 10
        with pytest.raises(XLRDError):
            bk.handle_filepass(data)
        log_output = bk.logfile.getvalue()
        assert 'FILEPASS' in log_output

    def test_verbosity_one_does_not_log(self):
        """verbosity==1 does not log (branch not taken), but still raises."""
        bk = _make_book(verbosity=1)
        data = b'\x00' * 10
        with pytest.raises(XLRDError):
            bk.handle_filepass(data)
        assert bk.logfile.getvalue() == ''
