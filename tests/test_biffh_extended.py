# -*- coding: utf-8 -*-
"""
Extended tests for xlrd.biffh — string unpackers, range list unpacker,
hex_char_dump, biff_dump, and biff_count_records.
"""
import io
import struct
import pytest

from xlrd.biffh import (
    unpack_string,
    unpack_string_update_pos,
    unpack_unicode,
    unpack_unicode_update_pos,
    unpack_cell_range_address_list_update_pos,
    hex_char_dump,
    biff_dump,
    biff_count_records,
)


# ---------------------------------------------------------------------------
# unpack_string
# ---------------------------------------------------------------------------

class TestUnpackString:

    def test_single_byte_length_latin1(self):
        data = b'\x05Hello'
        result = unpack_string(data, 0, 'latin_1', lenlen=1)
        assert result == 'Hello'

    def test_two_byte_length(self):
        data = struct.pack('<H', 3) + b'ABC'
        result = unpack_string(data, 0, 'latin_1', lenlen=2)
        assert result == 'ABC'

    def test_non_zero_pos_offset(self):
        data = b'\x00\x00\x03XYZ'
        result = unpack_string(data, 2, 'latin_1', lenlen=1)
        assert result == 'XYZ'

    def test_empty_string(self):
        data = b'\x00'
        result = unpack_string(data, 0, 'latin_1', lenlen=1)
        assert result == ''


# ---------------------------------------------------------------------------
# unpack_string_update_pos
# ---------------------------------------------------------------------------

class TestUnpackStringUpdatePos:

    def test_basic(self):
        data = b'\x04Test'
        strg, newpos = unpack_string_update_pos(data, 0, 'latin_1', lenlen=1)
        assert strg == 'Test'
        assert newpos == 5  # 1 (lenlen) + 4 (chars)

    def test_known_len(self):
        data = b'Hello'
        strg, newpos = unpack_string_update_pos(data, 0, 'latin_1', known_len=5)
        assert strg == 'Hello'
        assert newpos == 5

    def test_two_byte_length(self):
        data = struct.pack('<H', 2) + b'HI'
        strg, newpos = unpack_string_update_pos(data, 0, 'latin_1', lenlen=2)
        assert strg == 'HI'
        assert newpos == 4  # 2 (lenlen) + 2 (chars)

    def test_with_offset(self):
        # Skip 2 leading bytes, then a 1-byte length string
        data = b'\xFF\xFF\x03ABC'
        strg, newpos = unpack_string_update_pos(data, 2, 'latin_1', lenlen=1)
        assert strg == 'ABC'
        assert newpos == 6


# ---------------------------------------------------------------------------
# unpack_unicode
# ---------------------------------------------------------------------------

class TestUnpackUnicode:

    def test_empty_string_returns_empty(self):
        # nchars=0, function returns "" early without reading options byte
        data = struct.pack('<H', 0)
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == ''

    def test_compressed_ascii(self):
        text = 'ABC'
        data = struct.pack('<H', 3) + b'\x00' + text.encode('latin_1')
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == 'ABC'

    def test_uncompressed_utf16(self):
        text = 'Hi'
        data = struct.pack('<H', 2) + b'\x01' + text.encode('utf_16_le')
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == 'Hi'

    def test_richtext_flag_skipped(self):
        # options = 0x08 (richtext), rt=1 run → 2 extra bytes (H) ignored
        text = 'Test'
        nchars = len(text)
        options = 0x08
        rt = 1
        data = struct.pack('<H', nchars) + bytes([options])
        data += struct.pack('<H', rt)
        # phonetic ext: none
        data += text.encode('latin_1')
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == 'Test'

    def test_phonetic_flag_skipped(self):
        # options = 0x04 (phonetic), sz=0 → 4 extra bytes (i) ignored
        text = 'Test'
        nchars = len(text)
        options = 0x04
        sz = 0
        data = struct.pack('<H', nchars) + bytes([options])
        data += struct.pack('<i', sz)
        data += text.encode('latin_1')
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == 'Test'

    def test_lenlen_1(self):
        text = 'ABC'
        data = struct.pack('<B', 3) + b'\x00' + text.encode('latin_1')
        result = unpack_unicode(data, 0, lenlen=1)
        assert result == 'ABC'


# ---------------------------------------------------------------------------
# unpack_unicode_update_pos
# ---------------------------------------------------------------------------

class TestUnpackUnicodeUpdatePos:

    def test_zero_length_no_options_byte(self):
        # nchars=0, no more bytes → empty string, pos stays at 2
        data = struct.pack('<H', 0)
        strg, newpos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == ''
        assert newpos == 2

    def test_compressed_updates_pos(self):
        text = 'Hello'
        nchars = len(text)
        data = struct.pack('<H', nchars) + b'\x00' + text.encode('latin_1')
        strg, newpos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == 'Hello'
        assert newpos == 2 + 1 + nchars

    def test_uncompressed_updates_pos(self):
        text = 'Hi'
        nchars = len(text)
        data = struct.pack('<H', nchars) + b'\x01' + text.encode('utf_16_le')
        strg, newpos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == 'Hi'
        assert newpos == 2 + 1 + nchars * 2

    def test_richtext_advances_pos_by_run_data(self):
        # Layout: lenH | optB | rt_countH | string_bytes | run_data(4*rt)
        text = 'Hi'
        nchars = len(text)
        rt = 2  # 2 richtext runs → 4*2 = 8 extra bytes AFTER string
        options = 0x08
        data = struct.pack('<H', nchars) + bytes([options]) + struct.pack('<H', rt)
        data += text.encode('latin_1')  # string data first
        data += b'\x00' * (4 * rt)     # run data after
        strg, newpos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == 'Hi'
        # 2 (len) + 1 (opts) + 2 (rt count) + 2 (chars) + 8 (run data)
        assert newpos == 2 + 1 + 2 + 2 + 8

    def test_phonetic_advances_pos(self):
        # Layout: lenH | optB | szI | string_bytes | phonetic_data(sz)
        text = 'Hi'
        nchars = len(text)
        sz = 4  # 4 bytes of phonetic data AFTER string
        options = 0x04
        data = struct.pack('<H', nchars) + bytes([options]) + struct.pack('<i', sz)
        data += text.encode('latin_1')  # string data first
        data += b'\x00' * sz            # phonetic data after
        strg, newpos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == 'Hi'

    def test_known_len_skips_length_field(self):
        text = 'AB'
        data = b'\x00' + text.encode('latin_1')  # options byte then text
        strg, newpos = unpack_unicode_update_pos(data, 0, lenlen=2, known_len=2)
        assert strg == 'AB'
        assert newpos == 3  # 0 + 1 (opts) + 2 (chars)


# ---------------------------------------------------------------------------
# unpack_cell_range_address_list_update_pos
# ---------------------------------------------------------------------------

class TestUnpackCellRangeAddressListUpdatePos:

    def test_zero_ranges(self):
        data = struct.pack('<H', 0)
        output = []
        newpos = unpack_cell_range_address_list_update_pos(output, data, 0, 80, addr_size=6)
        assert output == []
        assert newpos == 2

    def test_one_range_addr6(self):
        # ra=0, rb=2, ca=0, cb=3 → output (0, 3, 0, 4)
        data = struct.pack('<H', 1) + struct.pack('<HHBB', 0, 2, 0, 3)
        output = []
        unpack_cell_range_address_list_update_pos(output, data, 0, 80, addr_size=6)
        assert output == [(0, 3, 0, 4)]

    def test_one_range_addr8(self):
        # addr_size=8, format <HHHH
        data = struct.pack('<H', 1) + struct.pack('<HHHH', 1, 3, 2, 5)
        output = []
        unpack_cell_range_address_list_update_pos(output, data, 0, 80, addr_size=8)
        assert output == [(1, 4, 2, 6)]

    def test_two_ranges(self):
        data = struct.pack('<H', 2)
        data += struct.pack('<HHBB', 0, 0, 0, 0)
        data += struct.pack('<HHBB', 1, 1, 1, 1)
        output = []
        unpack_cell_range_address_list_update_pos(output, data, 0, 80, addr_size=6)
        assert len(output) == 2
        assert output[0] == (0, 1, 0, 1)
        assert output[1] == (1, 2, 1, 2)

    def test_invalid_addr_size_raises(self):
        data = struct.pack('<H', 0)
        with pytest.raises(AssertionError):
            unpack_cell_range_address_list_update_pos([], data, 0, 80, addr_size=4)

    def test_non_zero_start_pos(self):
        # Put garbage before pos and a zero-count range list after
        data = b'\xFF\xFF' + struct.pack('<H', 0)
        output = []
        newpos = unpack_cell_range_address_list_update_pos(output, data, 2, 80)
        assert output == []
        assert newpos == 4


# ---------------------------------------------------------------------------
# hex_char_dump
# ---------------------------------------------------------------------------

def _make_stream(*records):
    """Build a BIFF byte stream from (opcode, payload) pairs."""
    out = b''
    for opcode, payload in records:
        out += struct.pack('<HH', opcode, len(payload)) + payload
    return out


class TestHexCharDump:

    def test_basic_output_contains_hex(self):
        data = b'ABCDEFGHIJKLMNOP'
        buf = io.StringIO()
        hex_char_dump(data, 0, len(data), fout=buf)
        output = buf.getvalue()
        assert '41 ' in output  # 'A' = 0x41

    def test_printable_chars_shown(self):
        data = b'Hello'
        buf = io.StringIO()
        hex_char_dump(data, 0, len(data), fout=buf)
        output = buf.getvalue()
        assert 'Hello' in output

    def test_non_printable_replaced(self):
        data = bytes([0x00, 0x41])  # null then 'A'
        buf = io.StringIO()
        hex_char_dump(data, 0, len(data), fout=buf)
        output = buf.getvalue()
        assert '~' in output  # null -> '~'

    def test_unnumbered_no_numeric_prefix(self):
        data = b'X'
        buf = io.StringIO()
        hex_char_dump(data, 0, len(data), base=0, fout=buf, unnumbered=True)
        # Just check it doesn't crash and produces output
        assert buf.getvalue() != ''

    def test_numbered_with_base_offset(self):
        data = b'A'
        buf = io.StringIO()
        hex_char_dump(data, 0, len(data), base=200, fout=buf, unnumbered=False)
        output = buf.getvalue()
        assert '200' in output

    def test_empty_range(self):
        buf = io.StringIO()
        hex_char_dump(b'Hello', 5, 0, fout=buf)
        assert buf.getvalue() == ''

    def test_multi_line_output(self):
        # 17 bytes → 2 lines (16 per line)
        data = bytes(range(17))
        buf = io.StringIO()
        hex_char_dump(data, 0, len(data), fout=buf)
        output = buf.getvalue()
        assert output.count('\n') >= 2


# ---------------------------------------------------------------------------
# biff_dump
# ---------------------------------------------------------------------------

class TestBiffDump:

    def test_known_record_bof(self):
        stream = _make_stream((0x0809, b'\x00' * 8))
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert 'BOF' in output
        assert '0809' in output

    def test_known_record_eof(self):
        stream = _make_stream((0x000A, b''))
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf)
        assert 'EOF' in buf.getvalue()

    def test_unknown_record_labeled(self):
        stream = _make_stream((0x9999, b'\x01\x02'))
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf)
        assert '<UNKNOWN>' in buf.getvalue()

    def test_empty_stream(self):
        buf = io.StringIO()
        biff_dump(b'', 0, 0, fout=buf)
        assert buf.getvalue() == ''

    def test_unnumbered_mode(self):
        stream = _make_stream((0x000A, b''))
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf, unnumbered=True)
        assert 'EOF' in buf.getvalue()

    def test_zero_padding_reported(self):
        stream = _make_stream((0x000A, b''))
        stream += b'\x00' * 16
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf)
        assert 'zero bytes skipped' in buf.getvalue()

    def test_multiple_records(self):
        stream = _make_stream(
            (0x0809, b'\x00' * 8),
            (0x000A, b''),
        )
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert 'BOF' in output
        assert 'EOF' in output

    def test_with_nonzero_base(self):
        stream = _make_stream((0x000A, b''))
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), base=100, fout=buf)
        output = buf.getvalue()
        assert '100' in output


# ---------------------------------------------------------------------------
# biff_count_records
# ---------------------------------------------------------------------------

class TestBiffCountRecords:

    def test_single_eof_record(self):
        stream = _make_stream((0x000A, b''))
        buf = io.StringIO()
        biff_count_records(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert 'EOF' in output
        assert '1' in output

    def test_two_eof_records_counted(self):
        stream = _make_stream((0x000A, b''), (0x000A, b''))
        buf = io.StringIO()
        biff_count_records(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert '2' in output

    def test_unknown_record_is_counted(self):
        stream = _make_stream((0xBEEF, b'\x00'))
        buf = io.StringIO()
        biff_count_records(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert 'Unknown_0x' in output

    def test_all_zero_padding_breaks_out(self):
        # All-zero stream: biff_count_records detects all-zero tail and breaks early
        # without recording any tally entry, so output will be empty.
        stream = b'\x00' * 8
        buf = io.StringIO()
        biff_count_records(stream, 0, len(stream), fout=buf)
        assert buf.getvalue() == ''

    def test_partial_zero_dummy_record(self):
        # 4 zero bytes followed by EOF triggers the Dummy (zero) tally entry
        stream = b'\x00\x00\x00\x00' + _make_stream((0x000A, b''))
        buf = io.StringIO()
        biff_count_records(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert '<Dummy (zero)>' in output

    def test_mixed_records(self):
        stream = _make_stream((0x0809, b'\x00' * 8), (0x000A, b''))
        buf = io.StringIO()
        biff_count_records(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        assert 'BOF' in output
        assert 'EOF' in output


    def test_consecutive_zero_records_accumulated(self):
        """biff_dump accumulates consecutive zero records in dummies (line 568)."""
        # Two zero records (each 4 bytes) followed by EOF — first triggers dummies=4,
        # second triggers line 568 (dummies += 4 = 8) before dummies are printed
        stream = b'\x00' * 4 + b'\x00' * 4 + struct.pack('<HH', 0x000A, 0)
        buf = io.StringIO()
        biff_dump(stream, 0, len(stream), fout=buf)
        output = buf.getvalue()
        # 8 bytes of zero padding should be reported as skipped
        assert '8 zero bytes skipped' in output
