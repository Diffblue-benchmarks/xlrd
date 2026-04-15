"""
Tests targeting uncovered lines in xlrd/book.py Book.handle_sst.
"""
import struct
import sys
import pytest

from xlrd.book import Book
from xlrd.biffh import XL_CONTINUE


def _make_book(verbosity=0, formatting_info=False):
    """Create a minimal Book instance with attributes needed by handle_sst."""
    bk = Book()
    bk.logfile = sys.stdout
    bk.verbosity = verbosity
    bk.formatting_info = formatting_info
    # Set up mem/position so get_record_parts_conditional returns None (no CONTINUE)
    # Use record code 0x0000 which is not XL_CONTINUE (0x3c)
    bk.mem = struct.pack('<HH', 0x0000, 0)
    bk._position = 0
    return bk


def _make_sst_data(strings):
    """Build a minimal SST record body with given list of strings (compressed latin-1)."""
    total = len(strings)
    unique = len(strings)
    body = struct.pack('<II', total, unique)
    for s in strings:
        encoded = s.encode('latin_1')
        body += struct.pack('<H', len(s))
        body += b'\x00'  # options: compressed
        body += encoded
    return body


class TestHandleSst:

    def test_sst_no_strings_sets_empty_shared_strings(self):
        """SST with 0 unique strings: _sharedstrings is an empty list."""
        bk = _make_book()
        data = struct.pack('<II', 0, 0)
        bk.handle_sst(data)
        assert bk._sharedstrings == []

    def test_sst_with_strings_populates_shared_strings(self):
        """SST with compressed strings: _sharedstrings contains the decoded strings."""
        bk = _make_book()
        data = _make_sst_data(['Hello', 'World'])
        bk.handle_sst(data)
        assert bk._sharedstrings == ['Hello', 'World']

    def test_sst_formatting_info_sets_rich_text_runlist_map(self):
        """formatting_info=True: _rich_text_runlist_map is set from SST rt_runlist."""
        bk = _make_book(formatting_info=True)
        data = _make_sst_data(['Test'])
        bk.handle_sst(data)
        assert hasattr(bk, '_rich_text_runlist_map')

    def test_sst_without_formatting_info_does_not_overwrite_rich_text_runlist_map(self):
        """formatting_info=False: _rich_text_runlist_map is not updated by handle_sst."""
        bk = _make_book(formatting_info=False)
        bk._rich_text_runlist_map = None  # set to sentinel
        data = _make_sst_data(['Test'])
        bk.handle_sst(data)
        assert bk._rich_text_runlist_map is None

    def test_sst_verbosity_high_logs_unique_strings(self, capsys):
        """verbosity >= 2: unique string count is printed to logfile."""
        bk = _make_book(verbosity=2)
        data = _make_sst_data(['One', 'Two'])
        bk.handle_sst(data)
        captured = capsys.readouterr()
        assert '2' in captured.out

    def test_sst_continue_record_appends_to_strlist(self, mocker):
        """While loop: CONTINUE records are consumed and appended to strlist."""
        bk = _make_book()
        # First call returns a CONTINUE record, second call returns None (stop)
        call_count = [0]

        def fake_get_record_parts_conditional(reqd_record):
            call_count[0] += 1
            if call_count[0] == 1:
                # Return a continuation with one more compressed string char
                cont_data = b'\x00' + 'X'.encode('latin_1')  # options + 1 char
                return (XL_CONTINUE, len(cont_data), cont_data)
            return (None, 0, b'')

        mocker.patch.object(
            bk, 'get_record_parts_conditional',
            side_effect=fake_get_record_parts_conditional
        )
        # 2-char string split across SST + CONTINUE: options=0 compressed
        # SST body: total=1, unique=1, nchars=2, options=0, first char 'A'
        body = struct.pack('<II', 1, 1)
        body += struct.pack('<H', 2)  # nchars=2
        body += b'\x00'              # options: compressed
        body += b'A'                 # first char only
        bk.handle_sst(body)
        assert bk._sharedstrings == ['AX']
        assert call_count[0] == 2
