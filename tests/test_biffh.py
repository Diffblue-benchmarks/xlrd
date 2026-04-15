# -*- coding: utf-8 -*-
"""Tests for xlrd/biffh.py utility functions."""
from __future__ import print_function

import io
import sys
from struct import pack

import pytest

from xlrd.biffh import (
    BaseObject,
    biff_count_records,
    biff_dump,
    biff_rec_name_dict,
    hex_char_dump,
    is_cell_opcode,
    unpack_cell_range_address_list_update_pos,
    unpack_string,
    unpack_string_update_pos,
    unpack_unicode,
    unpack_unicode_update_pos,
    upkbits,
    upkbitsL,
    XL_BOOLERR,
    XL_FORMULA,
    XL_LABELSST,
    XL_NUMBER,
    XL_RK,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sio():
    """Return a text-mode StringIO so fprintf / print can write to it."""
    return io.StringIO()


# ---------------------------------------------------------------------------
# BaseObject.dump
# ---------------------------------------------------------------------------

class TestBaseObjectDump:

    def test_dump_writes_attributes_to_file(self):
        obj = BaseObject()
        obj.name = "test"
        obj.value = 42
        f = _sio()
        obj.dump(f=f)
        output = f.getvalue()
        assert "name" in output
        assert "value" in output

    def test_dump_uses_stderr_when_f_is_none(self, capsys):
        obj = BaseObject()
        obj.x = 1
        # Should not raise; writes to stderr
        obj.dump(f=None)
        captured = capsys.readouterr()
        assert "x" in captured.err

    def test_dump_writes_header(self):
        obj = BaseObject()
        f = _sio()
        obj.dump(f=f, header="=== HEADER ===")
        assert "=== HEADER ===" in f.getvalue()

    def test_dump_writes_footer(self):
        obj = BaseObject()
        f = _sio()
        obj.dump(f=f, footer="=== FOOTER ===")
        assert "=== FOOTER ===" in f.getvalue()

    def test_dump_indent_affects_output(self):
        obj = BaseObject()
        obj.attr = "hello"
        f = _sio()
        obj.dump(f=f, indent=4)
        output = f.getvalue()
        # indent of 4 means 4 spaces before attribute name
        assert "    attr" in output

    def test_dump_list_attribute_shows_len(self):
        obj = BaseObject()
        obj.items = [1, 2, 3]
        f = _sio()
        obj.dump(f=f)
        output = f.getvalue()
        assert "len = 3" in output

    def test_dump_dict_attribute_shows_len(self):
        obj = BaseObject()
        obj.mapping = {"a": 1, "b": 2}
        f = _sio()
        obj.dump(f=f)
        output = f.getvalue()
        assert "len = 2" in output

    def test_dump_repr_these_shows_full_value(self):
        class MyObj(BaseObject):
            _repr_these = ["items"]
        obj = MyObj()
        obj.items = [1, 2, 3]
        f = _sio()
        obj.dump(f=f)
        output = f.getvalue()
        # Should print repr, not "len = ..."
        assert "[1, 2, 3]" in output

    def test_dump_nested_dump_object(self):
        parent = BaseObject()
        child = BaseObject()
        child.val = 99
        parent.child = child
        f = _sio()
        parent.dump(f=f)
        output = f.getvalue()
        assert "child" in output
        assert "val" in output

    def test_dump_with_slots(self):
        class SlottedObj(BaseObject):
            __slots__ = ["x", "y"]
            def __init__(self):
                self.x = 10
                self.y = 20
        obj = SlottedObj()
        f = _sio()
        obj.dump(f=f)
        output = f.getvalue()
        assert "x" in output
        assert "y" in output


# ---------------------------------------------------------------------------
# is_cell_opcode
# ---------------------------------------------------------------------------

class TestIsCellOpcode:

    @pytest.mark.parametrize("opcode", [
        XL_BOOLERR, XL_FORMULA, XL_LABELSST, XL_NUMBER, XL_RK,
    ])
    def test_known_opcodes_return_true(self, opcode):
        assert is_cell_opcode(opcode)

    @pytest.mark.parametrize("opcode", [0x0000, 0x0001, 0xFFFF, 0x0809])
    def test_unknown_opcodes_return_false(self, opcode):
        assert not is_cell_opcode(opcode)


# ---------------------------------------------------------------------------
# upkbits / upkbitsL
# ---------------------------------------------------------------------------

class TestUpkbits:

    def test_upkbits_sets_attribute(self):
        obj = BaseObject()
        manifest = [(0, 0x01, "flag0"), (4, 0xF0, "nibble")]
        upkbits(obj, 0x52, manifest)
        assert obj.flag0 == 0
        assert obj.nibble == 5

    def test_upkbits_all_bits_zero(self):
        obj = BaseObject()
        manifest = [(0, 0xFF, "byte")]
        upkbits(obj, 0x00, manifest)
        assert obj.byte == 0

    def test_upkbitsL_returns_int(self):
        obj = BaseObject()
        manifest = [(0, 0x0F, "low")]
        upkbitsL(obj, 0xFF, manifest)
        assert obj.low == 15
        assert isinstance(obj.low, int)

    def test_upkbitsL_multiple_fields(self):
        obj = BaseObject()
        manifest = [(0, 0x01, "bit0"), (1, 0x02, "bit1"), (4, 0x10, "bit4")]
        upkbitsL(obj, 0x13, manifest)
        assert obj.bit0 == 1
        assert obj.bit1 == 1
        assert obj.bit4 == 1


# ---------------------------------------------------------------------------
# unpack_string
# ---------------------------------------------------------------------------

class TestUnpackString:

    def test_lenlen1_ascii(self):
        data = b'\x05hello world'
        result = unpack_string(data, 0, 'latin_1', lenlen=1)
        assert result == 'hello'

    def test_lenlen2_ascii(self):
        data = pack('<H', 3) + b'abc'
        result = unpack_string(data, 0, 'latin_1', lenlen=2)
        assert result == 'abc'

    def test_empty_string_lenlen1(self):
        data = b'\x00'
        result = unpack_string(data, 0, 'latin_1', lenlen=1)
        assert result == ''

    def test_with_offset(self):
        data = b'\x00\x00\x03abc'
        result = unpack_string(data, 2, 'latin_1', lenlen=1)
        assert result == 'abc'


# ---------------------------------------------------------------------------
# unpack_string_update_pos
# ---------------------------------------------------------------------------

class TestUnpackStringUpdatePos:

    def test_without_known_len(self):
        data = b'\x03abc'
        strg, newpos = unpack_string_update_pos(data, 0, 'latin_1', lenlen=1)
        assert strg == 'abc'
        assert newpos == 4

    def test_with_known_len(self):
        data = b'hello'
        strg, newpos = unpack_string_update_pos(data, 0, 'latin_1', known_len=3)
        assert strg == 'hel'
        assert newpos == 3

    def test_lenlen2_without_known_len(self):
        data = pack('<H', 4) + b'test'
        strg, newpos = unpack_string_update_pos(data, 0, 'latin_1', lenlen=2)
        assert strg == 'test'
        assert newpos == 6

    def test_empty_string_with_known_len(self):
        data = b'abc'
        strg, newpos = unpack_string_update_pos(data, 0, 'latin_1', known_len=0)
        assert strg == ''
        assert newpos == 0


# ---------------------------------------------------------------------------
# unpack_unicode
# ---------------------------------------------------------------------------

class TestUnpackUnicode:

    def _make_unicode_data(self, text, compressed=True, richtext=False, phonetic=False):
        nchars = len(text)
        options = 0
        if not compressed:
            options |= 0x01
        if richtext:
            options |= 0x08
        if phonetic:
            options |= 0x04
        header = pack('<H', nchars) + bytes([options])
        extra = b''
        if richtext:
            extra += pack('<H', 0)  # rt = 0 richtext runs
        if phonetic:
            extra += pack('<i', 0)  # sz = 0
        if compressed:
            body = text.encode('latin_1')
        else:
            body = text.encode('utf_16_le')
        return header + extra + body

    def test_compressed_latin1_string(self):
        data = self._make_unicode_data('hello')
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == 'hello'

    def test_uncompressed_utf16_string(self):
        data = self._make_unicode_data('hi', compressed=False)
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == 'hi'

    def test_empty_string_returns_empty(self):
        data = pack('<H', 0)
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == ''

    def test_richtext_flag_skips_bytes(self):
        data = self._make_unicode_data('abc', richtext=True)
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == 'abc'

    def test_phonetic_flag_skips_bytes(self):
        data = self._make_unicode_data('xyz', phonetic=True)
        result = unpack_unicode(data, 0, lenlen=2)
        assert result == 'xyz'

    def test_lenlen1(self):
        nchars = 3
        options = 0  # compressed
        data = pack('<B', nchars) + bytes([options]) + b'cat'
        result = unpack_unicode(data, 0, lenlen=1)
        assert result == 'cat'


# ---------------------------------------------------------------------------
# unpack_unicode_update_pos
# ---------------------------------------------------------------------------

class TestUnpackUnicodeUpdatePos:

    def test_compressed_string_updates_pos(self):
        text = 'hello'
        nchars = len(text)
        options = 0x00  # compressed
        data = pack('<H', nchars) + bytes([options]) + text.encode('latin_1')
        strg, pos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == 'hello'
        assert pos == 2 + 1 + 5  # lenlen + options byte + 5 chars

    def test_uncompressed_string_updates_pos(self):
        text = 'hi'
        nchars = len(text)
        options = 0x01  # uncompressed UTF-16-LE
        data = pack('<H', nchars) + bytes([options]) + text.encode('utf_16_le')
        strg, pos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == 'hi'
        assert pos == 2 + 1 + 4  # lenlen + options + 2*nchars

    def test_zero_length_empty_data_returns_empty(self):
        data = pack('<H', 0)
        strg, pos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == ''
        assert pos == 2

    def test_known_len(self):
        text = 'abc'
        options = 0x00  # compressed
        data = bytes([options]) + text.encode('latin_1')
        strg, pos = unpack_unicode_update_pos(data, 0, lenlen=2, known_len=3)
        assert strg == 'abc'
        assert pos == 1 + 3

    def test_richtext_adjusts_pos(self):
        text = 'test'
        nchars = len(text)
        options = 0x08  # richtext
        rt = 2
        data = pack('<H', nchars) + bytes([options]) + pack('<H', rt) + text.encode('latin_1')
        strg, pos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == 'test'
        # lenlen(2) + options(1) + richtext_count(2) + nchars(4) + richtext_runs(2*4=8)
        assert pos == 2 + 1 + 2 + 4 + 4 * rt

    def test_phonetic_adjusts_pos(self):
        text = 'ab'
        nchars = len(text)
        options = 0x04  # phonetic
        sz = 6
        data = pack('<H', nchars) + bytes([options]) + pack('<i', sz) + text.encode('latin_1')
        strg, pos = unpack_unicode_update_pos(data, 0, lenlen=2)
        assert strg == 'ab'
        # lenlen(2) + options(1) + phonetic_size(4) + nchars(2) + sz(6)
        assert pos == 2 + 1 + 4 + 2 + sz


# ---------------------------------------------------------------------------
# unpack_cell_range_address_list_update_pos
# ---------------------------------------------------------------------------

class TestUnpackCellRangeAddressListUpdatePos:

    def test_empty_list_n_zero(self):
        output = []
        data = pack('<H', 0)
        pos = unpack_cell_range_address_list_update_pos(output, data, 0, biff_version=8)
        assert output == []
        assert pos == 2

    def test_single_range_addr_size_6(self):
        output = []
        n = 1
        ra, rb, ca, cb = 0, 2, 1, 3
        data = pack('<H', n) + pack('<HHBB', ra, rb, ca, cb)
        pos = unpack_cell_range_address_list_update_pos(output, data, 0, biff_version=5, addr_size=6)
        assert len(output) == 1
        assert output[0] == (ra, rb + 1, ca, cb + 1)
        assert pos == 2 + 6

    def test_single_range_addr_size_8(self):
        output = []
        n = 1
        ra, rb, ca, cb = 0, 5, 0, 5
        data = pack('<H', n) + pack('<HHHH', ra, rb, ca, cb)
        pos = unpack_cell_range_address_list_update_pos(output, data, 0, biff_version=8, addr_size=8)
        assert len(output) == 1
        assert output[0] == (ra, rb + 1, ca, cb + 1)
        assert pos == 2 + 8

    def test_multiple_ranges(self):
        output = []
        n = 2
        data = pack('<H', n)
        data += pack('<HHBB', 0, 1, 0, 1)
        data += pack('<HHBB', 2, 3, 2, 3)
        pos = unpack_cell_range_address_list_update_pos(output, data, 0, biff_version=5, addr_size=6)
        assert len(output) == 2
        assert pos == 2 + 2 * 6

    def test_invalid_addr_size_raises(self):
        output = []
        data = pack('<H', 0)
        with pytest.raises(AssertionError):
            unpack_cell_range_address_list_update_pos(output, data, 0, biff_version=8, addr_size=7)


# ---------------------------------------------------------------------------
# hex_char_dump
# ---------------------------------------------------------------------------

class TestHexCharDump:

    def test_basic_dump(self):
        data = b'Hello World'
        f = _sio()
        hex_char_dump(data, 0, len(data), fout=f)
        output = f.getvalue()
        assert '48' in output  # 'H' in hex
        assert 'Hello' in output

    def test_unnumbered_dump(self):
        data = b'ABCD'
        f = _sio()
        hex_char_dump(data, 0, len(data), fout=f, unnumbered=True)
        output = f.getvalue()
        assert '41' in output  # 'A' in hex

    def test_null_bytes_replaced_with_tilde(self):
        data = b'A\x00B'
        f = _sio()
        hex_char_dump(data, 0, len(data), fout=f, unnumbered=True)
        output = f.getvalue()
        assert '~' in output

    def test_non_printable_replaced_with_question_mark(self):
        data = b'\x01\x02\x03'
        f = _sio()
        hex_char_dump(data, 0, len(data), fout=f, unnumbered=True)
        output = f.getvalue()
        assert '?' in output

    def test_offset_and_dlen(self):
        data = b'XXABCXX'
        f = _sio()
        hex_char_dump(data, 2, 3, fout=f, unnumbered=True)
        output = f.getvalue()
        assert 'ABC' in output

    def test_numbered_includes_position(self):
        data = b'ABCDE'
        f = _sio()
        hex_char_dump(data, 0, len(data), base=0, fout=f, unnumbered=False)
        output = f.getvalue()
        assert '0:' in output or '    0:' in output

    def test_long_data_multiple_lines(self):
        data = bytes(range(32))
        f = _sio()
        hex_char_dump(data, 0, len(data), fout=f, unnumbered=True)
        output = f.getvalue()
        lines = [l for l in output.splitlines() if l.strip()]
        assert len(lines) == 2  # 32 bytes / 16 per line


# ---------------------------------------------------------------------------
# biff_dump
# ---------------------------------------------------------------------------

class TestBiffDump:

    def _make_record(self, rc, data=b''):
        return pack('<HH', rc, len(data)) + data

    def test_basic_record_dump(self):
        # Use a known record code: XL_EOF = 0x000a
        rec = self._make_record(0x000a)
        f = _sio()
        biff_dump(rec, 0, len(rec), fout=f)
        output = f.getvalue()
        assert 'EOF' in output or '000a' in output

    def test_unknown_record(self):
        rec = self._make_record(0xABCD)
        f = _sio()
        biff_dump(rec, 0, len(rec), fout=f)
        output = f.getvalue()
        assert 'UNKNOWN' in output or 'abcd' in output

    def test_zero_bytes_skipped_message(self):
        # A block of zeros followed by a real record
        zeros = b'\x00' * 8
        rec = self._make_record(0x000a)
        mem = zeros + rec
        f = _sio()
        biff_dump(mem, 0, len(mem), fout=f)
        output = f.getvalue()
        assert 'zero bytes skipped' in output

    def test_unnumbered_output(self):
        rec = self._make_record(0x000a)
        f = _sio()
        biff_dump(rec, 0, len(rec), fout=f, unnumbered=True)
        output = f.getvalue()
        assert output  # should produce some output

    def test_misc_bytes_at_end(self):
        rec = self._make_record(0x000a) + b'\xAB\xCD'  # 2 trailing bytes
        f = _sio()
        biff_dump(rec, 0, len(rec), fout=f)
        output = f.getvalue()
        assert 'Misc bytes at end' in output

    def test_all_zeros_stream(self):
        mem = b'\x00' * 16
        f = _sio()
        biff_dump(mem, 0, len(mem), fout=f)
        output = f.getvalue()
        assert 'zero bytes skipped' in output

    def test_record_with_data(self):
        data = b'ABCDEF'
        rc = 0x000a
        rec = self._make_record(rc, data)
        f = _sio()
        biff_dump(rec, 0, len(rec), fout=f)
        output = f.getvalue()
        assert 'ABCDEF' in output


# ---------------------------------------------------------------------------
# biff_count_records
# ---------------------------------------------------------------------------

class TestBiffCountRecords:

    def _make_record(self, rc, data=b''):
        return pack('<HH', rc, len(data)) + data

    def test_counts_single_record(self):
        rec = self._make_record(0x000a)
        f = _sio()
        biff_count_records(rec, 0, len(rec), fout=f)
        output = f.getvalue()
        assert 'EOF' in output
        assert '1' in output

    def test_counts_multiple_same_records(self):
        eof = self._make_record(0x000a)
        mem = eof + eof + eof
        f = _sio()
        biff_count_records(mem, 0, len(mem), fout=f)
        output = f.getvalue()
        assert '3' in output

    def test_counts_different_records(self):
        rec1 = self._make_record(0x000a)  # EOF
        rec2 = self._make_record(0x0042)  # CODEPAGE
        mem = rec1 + rec2
        f = _sio()
        biff_count_records(mem, 0, len(mem), fout=f)
        output = f.getvalue()
        assert 'EOF' in output
        assert 'CODEPAGE' in output

    def test_unknown_record_code(self):
        rec = self._make_record(0xABCD)
        f = _sio()
        biff_count_records(rec, 0, len(rec), fout=f)
        output = f.getvalue()
        assert 'Unknown_0xABCD' in output

    def test_all_zeros_stream(self):
        mem = b'\x00' * 16
        f = _sio()
        biff_count_records(mem, 0, len(mem), fout=f)
        # Should not raise; zero-filled stream is handled
        # (outputs nothing or dummy record counts)

    def test_dummy_zero_record(self):
        # rc=0, length=0 but not all zeros after
        dummy = pack('<HH', 0, 0) + self._make_record(0x000a)
        f = _sio()
        biff_count_records(dummy, 0, len(dummy), fout=f)
        output = f.getvalue()
        assert 'Dummy' in output or 'EOF' in output

    def test_output_is_sorted(self):
        rec_a = self._make_record(0x000a)  # EOF
        rec_b = self._make_record(0x0042)  # CODEPAGE
        mem = rec_b + rec_a
        f = _sio()
        biff_count_records(mem, 0, len(mem), fout=f)
        output = f.getvalue()
        lines = [l for l in output.splitlines() if l.strip()]
        # Output should be alphabetically sorted
        names = [l.split()[-1] for l in lines]
        assert names == sorted(names)
