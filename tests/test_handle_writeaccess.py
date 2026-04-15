"""
Tests targeting uncovered lines in xlrd/book.py Book.handle_writeaccess.
"""
import struct
import sys
import pytest

from xlrd.book import Book


def _make_book(biff_version=80, encoding=None):
    """Create a minimal Book instance with essential attributes set."""
    bk = Book()
    bk.biff_version = biff_version
    bk.encoding = encoding
    bk.raw_user_name = False
    bk.logfile = sys.stdout
    return bk


class TestHandleWriteaccess:

    def test_biff8_valid_compressed_string(self):
        """BIFF8: valid compressed (latin-1) unicode string sets user_name."""
        bk = _make_book(biff_version=80)
        name = b'TestUser'
        # lenlen=2 → 2-byte count, then 1-byte options (0x00 = compressed latin-1)
        data = struct.pack('<H', len(name)) + b'\x00' + name
        bk.handle_writeaccess(data)
        assert bk.user_name == 'TestUser'

    def test_biff8_valid_utf16_string(self):
        """BIFF8: valid UTF-16-LE unicode string sets user_name correctly."""
        bk = _make_book(biff_version=80)
        name = 'TestUser'
        encoded = name.encode('utf-16-le')
        # options=0x01 means UTF-16-LE uncompressed
        data = struct.pack('<H', len(name)) + b'\x01' + encoded
        bk.handle_writeaccess(data)
        assert bk.user_name == 'TestUser'

    def test_biff8_strips_trailing_whitespace(self):
        """BIFF8: trailing whitespace in user_name is stripped."""
        bk = _make_book(biff_version=80)
        name = b'User   '
        data = struct.pack('<H', len(name)) + b'\x00' + name
        bk.handle_writeaccess(data)
        assert bk.user_name == 'User'

    def test_biff8_unicode_decode_error_falls_back_to_strip(self, mocker):
        """BIFF8: UnicodeDecodeError causes fallback with data.strip()."""
        bk = _make_book(biff_version=80)
        call_count = [0]

        def mock_unpack_unicode(data, pos, lenlen=2):
            call_count[0] += 1
            if call_count[0] == 1:
                raise UnicodeDecodeError('utf_16_le', b'', 0, 1, 'invalid data')
            return 'FallbackUser'

        mocker.patch('xlrd.book.unpack_unicode', side_effect=mock_unpack_unicode)
        bk.handle_writeaccess(b'  some data  ')
        assert bk.user_name == 'FallbackUser'
        assert call_count[0] == 2

    def test_biff_pre8_no_encoding_sets_raw_user_name(self):
        """BIFF < 80 with no encoding: raw_user_name=True and user_name=data."""
        bk = _make_book(biff_version=70, encoding=None)
        data = b'\x04Test'
        bk.handle_writeaccess(data)
        assert bk.raw_user_name is True
        assert bk.user_name == data

    def test_biff_pre8_no_encoding_returns_early(self):
        """BIFF < 80 with no encoding: method returns early (no rstrip called)."""
        bk = _make_book(biff_version=50, encoding=None)
        data = b'\x07MyUser '
        bk.handle_writeaccess(data)
        # user_name should be the raw data, not stripped
        assert bk.user_name == data
        assert bk.raw_user_name is True

    def test_biff_pre8_with_encoding_unpacks_string(self):
        """BIFF < 80 with encoding: unpack_string decodes the name."""
        bk = _make_book(biff_version=70, encoding='cp1252')
        name = b'MyWorkbook'
        # unpack_string with lenlen=1: 1-byte length prefix
        data = struct.pack('B', len(name)) + name
        bk.handle_writeaccess(data)
        assert bk.user_name == 'MyWorkbook'

    def test_biff_pre8_with_encoding_strips_trailing_whitespace(self):
        """BIFF < 80 with encoding: trailing whitespace in decoded name is stripped."""
        bk = _make_book(biff_version=70, encoding='cp1252')
        name = b'Author   '
        data = struct.pack('B', len(name)) + name
        bk.handle_writeaccess(data)
        assert bk.user_name == 'Author'
