# -*- coding: utf-8 -*-
import io
import struct
import pytest

from xlrd.compdoc import (
    CompDoc, CompDocError, DirNode, SIGNATURE, EOCSID, FREESID, SATSID,
    _build_family_tree, dump_list, x_dump_line,
)


def make_dir_entry(name='', etype=0, colour=1, left_DID=-1, right_DID=-1,
                   root_DID=-1, first_SID=-2, tot_size=0):
    """Build a 128-byte raw directory entry."""
    buf = bytearray(128)
    if name:
        name_utf16 = (name + '\0').encode('utf_16_le')
        cbufsize = len(name_utf16)
        buf[0:cbufsize] = name_utf16
    else:
        cbufsize = 0
    struct.pack_into('<HBBiii', buf, 64, cbufsize, etype, colour, left_DID, right_DID, root_DID)
    struct.pack_into('<IIII', buf, 100, 0, 0, 0, 0)
    struct.pack_into('<ii', buf, 116, first_SID, tot_size)
    return bytes(buf)


def make_ole2_mem(stream_data=None):
    """
    Create a minimal valid OLE2 compound document in memory.

    Layout:
      - Header   (512 bytes)
      - Sector 0: SAT
      - Sector 1: Directory (4 entries: Root Entry, Workbook, empty, empty)
      - Sector 2: Workbook stream data
    """
    if stream_data is None:
        stream_data = b'X' * 512

    sec_size = 512
    nent = sec_size // 4  # 128 SAT entries per sector

    # SAT (sector 0)
    sat = bytearray(sec_size)
    struct.pack_into('<i', sat, 0, SATSID)   # sector 0 is the SAT itself
    struct.pack_into('<i', sat, 4, EOCSID)   # sector 1 (directory) end
    struct.pack_into('<i', sat, 8, EOCSID)   # sector 2 (stream) end
    for i in range(3, nent):
        struct.pack_into('<i', sat, i * 4, FREESID)

    # Directory (sector 1): 4 entries of 128 bytes each
    dir_sector = (
        make_dir_entry('Root Entry', etype=5, root_DID=1, first_SID=EOCSID, tot_size=0) +
        make_dir_entry('Workbook', etype=2, first_SID=2, tot_size=len(stream_data)) +
        make_dir_entry('', etype=0) +
        make_dir_entry('', etype=0)
    )

    # Stream data sector (padded to sector size)
    stream_sector = (stream_data + b'\x00' * sec_size)[:sec_size]

    # Header (512 bytes)
    header = bytearray(512)
    header[0:8] = SIGNATURE
    struct.pack_into('<HH', header, 24, 0x003E, 0x0003)   # revision=0x3E, version=0x03
    header[28:30] = b'\xFE\xFF'                            # little-endian marker
    struct.pack_into('<HH', header, 30, 9, 6)              # ssz=9 (512 B), sssz=6 (64 B)
    # SAT_tot_secs=1, dir_first_sec_sid=1, _unused=0, min_size_std_stream=0,
    # SSAT_first_sec_sid=EOCSID, SSAT_tot_secs=0, MSATX_first_sec_sid=EOCSID, MSATX_tot_secs=0
    struct.pack_into('<iiiiiiii', header, 44, 1, 1, 0, 0, EOCSID, 0, EOCSID, 0)
    struct.pack_into('<i', header, 76, 0)                  # MSAT[0] = 0 (SAT at sector 0)
    for i in range(1, 109):
        struct.pack_into('<i', header, 76 + i * 4, FREESID)

    return bytes(header) + bytes(sat) + bytes(dir_sector) + bytes(stream_sector)


@pytest.fixture
def valid_mem():
    return make_ole2_mem()


@pytest.fixture
def compdoc(valid_mem):
    return CompDoc(valid_mem)


# === CompDoc error conditions ===

def test_not_ole2_raises_error():
    mem = b'\x00' * 512
    with pytest.raises(CompDocError, match='Not an OLE2 compound document'):
        CompDoc(mem)


def test_wrong_endian_marker_raises_error():
    mem = bytearray(512)
    mem[0:8] = SIGNATURE
    mem[28:30] = b'\xFF\xFE'
    with pytest.raises(CompDocError, match='Expected "little-endian" marker'):
        CompDoc(bytes(mem))


# === CompDoc initialization ===

def test_compdoc_init_success(valid_mem):
    doc = CompDoc(valid_mem)
    assert doc.sec_size == 512
    assert doc.short_sec_size == 64
    assert len(doc.dirlist) == 4
    assert doc.dirlist[0].name == 'Root Entry'
    assert doc.dirlist[1].name == 'Workbook'


def test_compdoc_init_with_debug(valid_mem):
    f = io.StringIO()
    doc = CompDoc(valid_mem, logfile=f, DEBUG=1)
    output = f.getvalue()
    assert 'CompDoc format' in output


def test_compdoc_has_workbook_child(compdoc):
    root = compdoc.dirlist[0]
    assert 1 in root.children


# === DirNode ===

def test_dirnode_init():
    dent = make_dir_entry('TestStream', etype=2, first_SID=5, tot_size=1024)
    node = DirNode(0, dent)
    assert node.DID == 0
    assert node.name == 'TestStream'
    assert node.etype == 2
    assert node.first_SID == 5
    assert node.tot_size == 1024
    assert node.left_DID == -1
    assert node.right_DID == -1


def test_dirnode_init_empty_name():
    dent = make_dir_entry('', etype=0)
    node = DirNode(3, dent)
    assert node.name == ''
    assert node.etype == 0


def test_dirnode_dump():
    dent = make_dir_entry('TestStream', etype=2, first_SID=5, tot_size=100)
    node = DirNode(0, dent)
    f = io.StringIO()
    node.logfile = f
    node.dump(1)
    output = f.getvalue()
    assert 'DID=0' in output
    assert 'TestStream' in output


def test_dirnode_dump_debug2():
    dent = make_dir_entry('TestStream', etype=2)
    node = DirNode(0, dent)
    f = io.StringIO()
    node.logfile = f
    node.dump(2)
    output = f.getvalue()
    assert 'timestamp info' in output


def test_dirnode_init_with_debug():
    dent = make_dir_entry('TestStream', etype=2)
    f = io.StringIO()
    node = DirNode(0, dent, DEBUG=1, logfile=f)
    output = f.getvalue()
    assert 'DID=0' in output


# === _build_family_tree ===

def test_build_family_tree_negative_child():
    dent0 = make_dir_entry('Root', etype=5, root_DID=-1)
    node0 = DirNode(0, dent0)
    dirlist = [node0]
    _build_family_tree(dirlist, 0, -1)
    assert dirlist[0].children == []


def test_build_family_tree_single_child():
    dent0 = make_dir_entry('Root', etype=5, root_DID=1)
    dent1 = make_dir_entry('Child', etype=2, left_DID=-1, right_DID=-1)
    node0 = DirNode(0, dent0)
    node1 = DirNode(1, dent1)
    dirlist = [node0, node1]
    _build_family_tree(dirlist, 0, 1)
    assert 1 in dirlist[0].children
    assert dirlist[1].parent == 0


def test_build_family_tree_storage_recursion():
    dent0 = make_dir_entry('Root', etype=5, root_DID=1)
    dent1 = make_dir_entry('Storage', etype=1, left_DID=-1, right_DID=-1, root_DID=2)
    dent2 = make_dir_entry('Stream', etype=2, left_DID=-1, right_DID=-1)
    node0 = DirNode(0, dent0)
    node1 = DirNode(1, dent1)
    node2 = DirNode(2, dent2)
    dirlist = [node0, node1, node2]
    _build_family_tree(dirlist, 0, 1)
    assert 1 in dirlist[0].children
    assert 2 in dirlist[1].children
    assert dirlist[2].parent == 1


# === get_named_stream ===

def test_get_named_stream_found(valid_mem):
    doc = CompDoc(valid_mem)
    data = doc.get_named_stream('Workbook')
    assert data == b'X' * 512


def test_get_named_stream_not_found(valid_mem):
    doc = CompDoc(valid_mem)
    data = doc.get_named_stream('NonExistent')
    assert data is None


# === locate_named_stream ===

def test_locate_named_stream_found(valid_mem):
    doc = CompDoc(valid_mem)
    mem, offset, size = doc.locate_named_stream('Workbook')
    assert mem is not None
    assert size == 512
    assert mem[offset:offset + size] == b'X' * 512


def test_locate_named_stream_not_found(valid_mem):
    doc = CompDoc(valid_mem)
    result = doc.locate_named_stream('NonExistent')
    assert result == (None, 0, 0)


# === _dir_search edge cases ===

def test_dir_search_not_found(compdoc):
    result = compdoc._dir_search(['Missing'])
    assert result is None


def test_dir_search_storage_no_tail(valid_mem):
    doc = CompDoc(valid_mem)
    dent0 = make_dir_entry('Root', etype=5, root_DID=1)
    dent1 = make_dir_entry('Storage', etype=1, left_DID=-1, right_DID=-1)
    node0 = DirNode(0, dent0)
    node1 = DirNode(1, dent1)
    node0.children = [1]
    node1.parent = 0
    doc.dirlist = [node0, node1]
    with pytest.raises(CompDocError, match="Requested component is a 'storage'"):
        doc._dir_search(['Storage'])


def test_dir_search_invalid_etype(valid_mem):
    doc = CompDoc(valid_mem)
    dent0 = make_dir_entry('Root', etype=5)
    dent1 = make_dir_entry('Weird', etype=3)
    node0 = DirNode(0, dent0)
    node1 = DirNode(1, dent1)
    node0.children = [1]
    node1.parent = 0
    f = io.StringIO()
    node1.logfile = f
    doc.dirlist = [node0, node1]
    with pytest.raises(CompDocError, match="Requested stream is not a 'user stream'"):
        doc._dir_search(['Weird'])


# === dump_list ===

def test_dump_list_empty():
    f = io.StringIO()
    dump_list([], 10, f)
    assert f.getvalue() == ''


def test_dump_list_single_row():
    f = io.StringIO()
    dump_list([1, 2, 3], 3, f)
    output = f.getvalue()
    assert '1' in output
    assert '2' in output
    assert '3' in output


def test_dump_list_multiple_rows():
    f = io.StringIO()
    dump_list(list(range(10)), 3, f)
    output = f.getvalue()
    assert '0' in output
    assert '9' in output


def test_dump_list_equal_rows():
    f = io.StringIO()
    alist = [1, 2, 3] * 4
    dump_list(alist, 3, f)
    output = f.getvalue()
    assert '=' in output


# === x_dump_line ===

def test_x_dump_line():
    f = io.StringIO()
    x_dump_line([10, 20, 30, 40], 2, f, 0)
    output = f.getvalue()
    assert '10' in output
    assert '20' in output


def test_x_dump_line_equal():
    f = io.StringIO()
    x_dump_line([10, 20, 30], 2, f, 0, equal=1)
    output = f.getvalue()
    assert '=' in output
