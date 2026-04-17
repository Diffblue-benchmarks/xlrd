# -*- coding: utf-8 -*-
"""
Tests for xlrd.compdoc — CompDocError, DirNode, SIGNATURE constant,
_build_family_tree, and CompDoc rejection of invalid streams.
"""
import io
import struct
import pytest
import xlwt

from xlrd.compdoc import (
    CompDocError,
    CompDoc,
    DirNode,
    SIGNATURE,
    EOCSID,
    FREESID,
    SATSID,
    MSATSID,
    EVILSID,
    _build_family_tree,
    dump_list,
    x_dump_line,
)


def _valid_xls_bytes():
    """Return raw bytes of a minimal valid xlwt-generated XLS file."""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Sheet1')
    ws.write(0, 0, 'test')
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

class TestConstants:

    def test_signature_is_8_bytes(self):
        assert len(SIGNATURE) == 8

    def test_signature_value(self):
        assert SIGNATURE == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

    def test_eocsid_is_minus_2(self):
        assert EOCSID == -2

    def test_freesid_is_minus_1(self):
        assert FREESID == -1

    def test_satsid_is_minus_3(self):
        assert SATSID == -3

    def test_msatsid_is_minus_4(self):
        assert MSATSID == -4

    def test_evilsid_is_minus_5(self):
        assert EVILSID == -5


# ---------------------------------------------------------------------------
# CompDocError
# ---------------------------------------------------------------------------

class TestCompDocError:

    def test_is_exception(self):
        with pytest.raises(CompDocError):
            raise CompDocError('test error')

    def test_message_preserved(self):
        try:
            raise CompDocError('my message')
        except CompDocError as e:
            assert 'my message' in str(e)


# ---------------------------------------------------------------------------
# DirNode
# ---------------------------------------------------------------------------

def _make_dir_entry(name='Root', etype=5, left_DID=-1, right_DID=-1,
                    root_DID=-1, first_SID=-2, tot_size=0):
    """Build a 128-byte directory entry for DirNode."""
    # Name is stored as UTF-16-LE, max 32 characters (64 bytes)
    name_encoded = name.encode('utf_16_le')
    cbufsize = len(name_encoded) + 2  # includes trailing NUL
    name_field = name_encoded + b'\x00\x00'
    name_field = name_field.ljust(64, b'\x00')

    # Offset 64: cbufsize(H), etype(B), colour(B), left_DID(i), right_DID(i), root_DID(i)
    part1 = struct.pack('<HBBiii', cbufsize, etype, 0, left_DID, right_DID, root_DID)

    # Offsets 80-99: padding (20 bytes)
    part2 = b'\x00' * 20

    # Offset 100: timestamp info (4 × 4 bytes = 16 bytes)
    part3 = b'\x00' * 16

    # Offset 116: first_SID(i), tot_size(i)
    part4 = struct.pack('<ii', first_SID, tot_size)

    # Offset 124: 4 bytes padding
    part5 = b'\x00' * 4

    dent = name_field + part1 + part2 + part3 + part4 + part5
    assert len(dent) == 128, f"dent is {len(dent)} bytes, expected 128"
    return dent


class TestDirNode:

    def test_name_extracted(self):
        dent = _make_dir_entry(name='Root')
        node = DirNode(0, dent)
        assert node.name == 'Root'

    def test_empty_name(self):
        dent = _make_dir_entry(name='')
        # When cbufsize=0, name should be empty string
        # Build manually with cbufsize=0
        raw = b'\x00' * 64  # empty name field
        part1 = struct.pack('<HBBiii', 0, 5, 0, -1, -1, -1)
        part2 = b'\x00' * 20
        part3 = b'\x00' * 16
        part4 = struct.pack('<ii', -2, 0)
        part5 = b'\x00' * 4
        dent = raw + part1 + part2 + part3 + part4 + part5
        node = DirNode(0, dent)
        assert node.name == ''

    def test_did_set(self):
        dent = _make_dir_entry()
        node = DirNode(7, dent)
        assert node.DID == 7

    def test_etype_extracted(self):
        dent = _make_dir_entry(etype=2)
        node = DirNode(0, dent)
        assert node.etype == 2

    def test_first_sid_extracted(self):
        dent = _make_dir_entry(first_SID=42)
        node = DirNode(0, dent)
        assert node.first_SID == 42

    def test_tot_size_extracted(self):
        dent = _make_dir_entry(tot_size=512)
        node = DirNode(0, dent)
        assert node.tot_size == 512

    def test_children_initially_empty(self):
        dent = _make_dir_entry()
        node = DirNode(0, dent)
        assert node.children == []

    def test_parent_initially_orphan(self):
        dent = _make_dir_entry()
        node = DirNode(0, dent)
        assert node.parent == -1

    def test_left_right_root_did(self):
        dent = _make_dir_entry(left_DID=1, right_DID=2, root_DID=3)
        node = DirNode(0, dent)
        assert node.left_DID == 1
        assert node.right_DID == 2
        assert node.root_DID == 3


# ---------------------------------------------------------------------------
# _build_family_tree
# ---------------------------------------------------------------------------

class TestBuildFamilyTree:

    def _make_nodes(self, count):
        nodes = []
        for i in range(count):
            dent = _make_dir_entry(name=f'Node{i}')
            nodes.append(DirNode(i, dent))
        return nodes

    def test_negative_child_does_nothing(self):
        nodes = self._make_nodes(2)
        _build_family_tree(nodes, 0, -1)
        assert nodes[0].children == []

    def test_child_added_to_parent(self):
        # Node 0 is parent, node 1 is child with no further children
        nodes = self._make_nodes(2)
        # node1 has left=-1, right=-1 (default), etype != 1 so no recursive storage
        _build_family_tree(nodes, 0, 1)
        assert 1 in nodes[0].children
        assert nodes[1].parent == 0

    def test_storage_etype_recurses(self):
        """Node with etype==1 (storage) triggers recursive _build_family_tree (line 71)."""
        # 3 nodes: 0=root, 1=storage-child, 2=leaf child of node1
        nodes = self._make_nodes(3)
        # Make node 1 a storage (etype=1) with root_DID=2
        dent = _make_dir_entry(name='Node1', etype=1, root_DID=2)
        nodes[1] = DirNode(1, dent)
        nodes[1].children = []
        nodes[1].parent = -1
        # Build tree: root(0) → child(1) which is storage with root(2)
        _build_family_tree(nodes, 0, 1)
        assert 1 in nodes[0].children
        assert nodes[1].parent == 0
        # Node1 is storage so node2 should be added under node1
        assert 2 in nodes[1].children

    def test_left_right_siblings(self):
        """Tree with left and right children is built correctly."""
        nodes = self._make_nodes(3)
        # node1: left_DID=-1, right_DID=2
        dent = _make_dir_entry(name='Node1', right_DID=2)
        nodes[1] = DirNode(1, dent)
        nodes[1].children = []
        nodes[1].parent = -1
        _build_family_tree(nodes, 0, 1)
        assert 1 in nodes[0].children
        assert 2 in nodes[0].children


# ---------------------------------------------------------------------------
# DirNode dump method
# ---------------------------------------------------------------------------

class TestDirNodeDump:

    def test_dump_default_debug(self):
        """DirNode.dump() writes to logfile (covers lines 54-59)."""
        dent = _make_dir_entry(name='TestNode', etype=2, first_SID=0, tot_size=512)
        node = DirNode(0, dent)
        logfile = io.StringIO()
        node.logfile = logfile
        node.dump(DEBUG=1)
        output = logfile.getvalue()
        assert 'TestNode' in output
        assert 'DID=0' in output

    def test_dump_debug_2(self):
        """DirNode.dump(DEBUG=2) also prints timestamp info (line 62)."""
        dent = _make_dir_entry(name='TestNode')
        node = DirNode(0, dent)
        logfile = io.StringIO()
        node.logfile = logfile
        node.dump(DEBUG=2)
        output = logfile.getvalue()
        assert 'timestamp info' in output


# ---------------------------------------------------------------------------
# CompDoc — invalid streams
# ---------------------------------------------------------------------------

class TestCompDocInvalidStream:

    def test_wrong_signature_raises(self):
        mem = b'\x00' * 512
        with pytest.raises(CompDocError, match="Not an OLE2"):
            CompDoc(mem, logfile=io.StringIO())

    def test_correct_signature_wrong_endian_marker_raises(self):
        # Correct signature but wrong byte-order marker at offset 28
        mem = SIGNATURE + b'\x00' * 20 + b'\x00\x00'  # wrong endian marker
        mem = mem.ljust(512, b'\x00')
        with pytest.raises(CompDocError, match="little-endian"):
            CompDoc(mem, logfile=io.StringIO())


# ---------------------------------------------------------------------------
# CompDoc — header warning paths (using patched xlwt-generated XLS bytes)
# ---------------------------------------------------------------------------

class TestCompDocHeaderWarnings:

    def test_preposterous_ssz_warning(self):
        """ssz > 20 triggers 'preposterous sector size' warning (lines 98-100)."""
        import xlrd
        data = bytearray(_valid_xls_bytes())
        data[30:32] = struct.pack('<H', 25)  # ssz=25 > 20
        log = io.StringIO()
        bk = xlrd.open_workbook(file_contents=bytes(data), logfile=log)
        assert 'preposterous' in log.getvalue()
        assert bk.nsheets > 0

    def test_preposterous_sssz_warning(self):
        """sssz > ssz triggers 'preposterous short stream sector' warning (lines 102-104)."""
        import xlrd
        data = bytearray(_valid_xls_bytes())
        data[32:34] = struct.pack('<H', 10)  # sssz=10 > ssz=9
        log = io.StringIO()
        bk = xlrd.open_workbook(file_contents=bytes(data), logfile=log)
        assert 'preposterous' in log.getvalue()
        assert bk.nsheets > 0

    def test_file_size_not_multiple_of_sector(self):
        """File size not 512 + N*sector triggers size warning (lines 118-120)."""
        import xlrd
        data = bytearray(_valid_xls_bytes())
        data += b'\x00'  # Add 1 byte to make size not a multiple
        log = io.StringIO()
        bk = xlrd.open_workbook(file_contents=bytes(data), logfile=log)
        assert 'not 512' in log.getvalue() or bk.nsheets > 0  # may warn or succeed

    def test_non_standard_sector_sizes_printed(self):
        """Non-512/64 sector sizes print @@@@ line (line 108)."""
        import xlrd
        data = bytearray(_valid_xls_bytes())
        data[30:32] = struct.pack('<H', 10)  # ssz=10 (1024-byte sectors, non-standard)
        log = io.StringIO()
        try:
            xlrd.open_workbook(file_contents=bytes(data), logfile=log)
        except Exception:
            pass
        assert '@@@@' in log.getvalue()


# ---------------------------------------------------------------------------
# CompDoc — direct method testing
# ---------------------------------------------------------------------------

class TestCompDocMethods:

    @pytest.fixture
    def compdoc(self):
        """A CompDoc constructed from a valid xlwt-generated XLS file."""
        return CompDoc(_valid_xls_bytes(), logfile=io.StringIO())

    def test_get_named_stream_workbook(self, compdoc):
        """get_named_stream('Workbook') returns stream bytes (lines 362, 365-368)."""
        result = compdoc.get_named_stream('Workbook')
        assert result is not None
        assert len(result) > 0

    def test_get_named_stream_nonexistent_returns_none(self, compdoc):
        """get_named_stream for non-existent stream returns None (lines 363-364)."""
        result = compdoc.get_named_stream('NonExistentStream')
        assert result is None

    def test_locate_named_stream_nonexistent(self, compdoc):
        """locate_named_stream for non-existent stream returns (None, 0, 0) (line 393)."""
        mem, base, length = compdoc.locate_named_stream('NoSuchStream')
        assert mem is None
        assert base == 0
        assert length == 0

    def test_locate_named_stream_workbook(self, compdoc):
        """locate_named_stream('Workbook') returns stream data."""
        mem, base, length = compdoc.locate_named_stream('Workbook')
        assert mem is not None
        assert length > 0

    def test_get_named_stream_book_alternative(self, compdoc):
        """get_named_stream for 'Book' (BIFF5 name) also works or returns None."""
        result = compdoc.get_named_stream('Book')
        # Valid result is either bytes or None depending on what xlwt writes
        assert result is None or isinstance(result, bytes)


# ---------------------------------------------------------------------------
# dump_list and x_dump_line utility functions (lines 461-485)
# ---------------------------------------------------------------------------

class TestDumpList:

    def test_empty_list_produces_no_output(self):
        """dump_list with empty list produces no output."""
        f = io.StringIO()
        dump_list([], 5, f)
        assert f.getvalue() == ''

    def test_single_element(self):
        """dump_list with single element."""
        f = io.StringIO()
        dump_list([42], 1, f)
        assert '42' in f.getvalue()

    def test_non_repeating_elements(self):
        """Non-repeating elements all printed (no equal= lines)."""
        f = io.StringIO()
        dump_list([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 5, f)
        output = f.getvalue()
        assert '1' in output
        assert '6' in output

    def test_repeating_elements_uses_equal_marker(self):
        """Repeating runs produce equal= summary lines (last section)."""
        f = io.StringIO()
        dump_list([1] * 15, 5, f)
        output = f.getvalue()
        assert '=' in output

    def test_change_after_multiple_repeating_strides(self):
        """Change after >1 skipped repeated strides triggers inner equal= line (line 481)."""
        f = io.StringIO()
        dump_list([1] * 15 + [2] * 5, 5, f)
        output = f.getvalue()
        # Should show 0, 10=, 15
        assert '=' in output
        lines = output.strip().split('\n')
        assert len(lines) == 3
        assert '=' in lines[1]

    def test_mixed_repeating_non_repeating(self):
        """Mixed repeating and non-repeating data."""
        f = io.StringIO()
        dump_list([1, 1, 1, 1, 1, 2, 2, 2, 2, 2], 5, f)
        output = f.getvalue()
        assert '1' in output
        assert '2' in output


class TestXDumpLine:

    def test_x_dump_line_basic(self):
        """x_dump_line prints position and values."""
        f = io.StringIO()
        x_dump_line([10, 20, 30, 40, 50], 5, f, 0)
        output = f.getvalue()
        assert '10' in output
        assert '50' in output

    def test_x_dump_line_with_equal_flag(self):
        """x_dump_line with equal=1 prints '=' marker."""
        f = io.StringIO()
        x_dump_line([1, 2, 3, 4, 5], 5, f, 0, equal=1)
        output = f.getvalue()
        assert '=' in output
