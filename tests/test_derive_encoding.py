"""
Tests targeting uncovered lines in xlrd/book.py Book.derive_encoding.
"""
import sys
import struct
import pytest

from xlrd.book import Book


def _make_book(biff_version=80, codepage=None, encoding=None,
               encoding_override=None, verbosity=0, raw_user_name=False,
               user_name=b''):
    """Create a minimal Book instance with essential attributes set."""
    bk = Book()
    bk.biff_version = biff_version
    bk.codepage = codepage
    bk.encoding = encoding
    bk.encoding_override = encoding_override
    bk.verbosity = verbosity
    bk.logfile = sys.stdout
    bk.raw_user_name = raw_user_name
    bk.user_name = user_name
    return bk


class TestDeriveEncoding:

    def test_encoding_override_sets_encoding(self):
        """When encoding_override is set, self.encoding is set to it."""
        bk = _make_book(encoding_override='utf-8', codepage=1200)
        result = bk.derive_encoding()
        assert result == 'utf-8'
        assert bk.encoding == 'utf-8'

    def test_no_codepage_biff_lt_80_sets_iso_8859_1(self):
        """codepage is None + biff_version < 80 -> encoding set to iso-8859-1."""
        bk = _make_book(biff_version=70, codepage=None)
        result = bk.derive_encoding()
        assert result == 'iso-8859-1'
        assert bk.encoding == 'iso-8859-1'

    def test_no_codepage_biff_gte_80_sets_codepage_1200(self):
        """codepage is None + biff_version >= 80 -> codepage set to 1200 (utf_16_le assumed)."""
        bk = _make_book(biff_version=80, codepage=None, encoding='utf_16_le')
        bk.derive_encoding()
        assert bk.codepage == 1200

    def test_no_codepage_biff_gte_80_verbosity_2(self, capsys):
        """codepage is None + biff_version >= 80 + verbosity >= 2 -> log message."""
        bk = _make_book(biff_version=80, codepage=None, verbosity=2)
        bk.derive_encoding()
        assert bk.codepage == 1200

    def test_codepage_in_encoding_from_codepage_map(self):
        """codepage 10000 is in encoding_from_codepage -> mac_roman."""
        bk = _make_book(codepage=10000)
        result = bk.derive_encoding()
        assert result == 'mac_roman'

    def test_codepage_in_range_300_to_1999(self):
        """codepage in 300..1999 range uses cp<N> encoding."""
        bk = _make_book(codepage=1252)
        result = bk.derive_encoding()
        assert result == 'cp1252'
        assert bk.encoding == 'cp1252'

    def test_codepage_not_known_biff_gte_80_sets_utf16le(self):
        """codepage not in map, not 300-1999, biff >= 80 -> utf_16_le."""
        bk = _make_book(biff_version=80, codepage=5000)
        result = bk.derive_encoding()
        assert result == 'utf_16_le'
        assert bk.codepage == 1200

    def test_codepage_not_known_biff_lt_80_sets_unknown_raises(self):
        """codepage not in map, not 300-1999, biff < 80 -> unknown_codepage_N, then raises on decode."""
        bk = _make_book(biff_version=70, codepage=5000)
        with pytest.raises(BaseException):
            bk.derive_encoding()

    def test_verbosity_logs_codepage_when_encoding_differs(self, capsys):
        """verbosity set + encoding differs from current -> log line printed."""
        bk = _make_book(codepage=1252, encoding='ascii', verbosity=1)
        bk.derive_encoding()
        assert bk.encoding == 'cp1252'

    def test_invalid_encoding_raises(self):
        """Non-1200 codepage with an invalid encoding raises an exception."""
        bk = _make_book(codepage=9999, biff_version=70)
        # unknown_codepage_9999 is not a valid codec
        with pytest.raises(BaseException):
            bk.derive_encoding()

    def test_raw_user_name_decoded_and_cleared(self):
        """raw_user_name=True: user_name is decoded from bytes and raw_user_name set False."""
        name = b'Author'
        data = struct.pack('B', len(name)) + name
        bk = _make_book(codepage=1252, raw_user_name=True, user_name=data)
        bk.derive_encoding()
        assert bk.user_name == 'Author'
        assert bk.raw_user_name is False

    def test_raw_user_name_strips_trailing_whitespace(self):
        """raw_user_name=True: decoded user_name is rstripped."""
        name = b'Author   '
        data = struct.pack('B', len(name)) + name
        bk = _make_book(codepage=1252, raw_user_name=True, user_name=data)
        bk.derive_encoding()
        assert bk.user_name == 'Author'
        assert bk.raw_user_name is False
