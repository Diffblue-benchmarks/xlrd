"""
Tests for xlrd/book.py Book.biff2_8_load targeting uncovered lines.
"""
import io
import struct
import sys
import pytest

from xlrd.book import Book
from xlrd.biffh import XLRDError
from xlrd import compdoc


# ---------------------------------------------------------------------------
# Helpers to build minimal OLE2 compound documents
# ---------------------------------------------------------------------------

SECTOR_SIZE = 512
EOCSID = -2
FREESID = -1
SATSID = -3


def _pack_dir_entry(name, etype, colour, left_did, right_did, root_did, first_sid, tot_size):
    """Build a 128-byte OLE2 directory entry."""
    entry = bytearray(128)
    if name:
        name_encoded = name.encode('utf-16-le') + b'\x00\x00'
        cbufsize = len(name_encoded)
        entry[0:len(name_encoded)] = name_encoded
        struct.pack_into('<H', entry, 64, cbufsize)
    struct.pack_into('<B', entry, 66, etype)
    struct.pack_into('<B', entry, 67, colour)
    struct.pack_into('<i', entry, 68, left_did)
    struct.pack_into('<i', entry, 72, right_did)
    struct.pack_into('<i', entry, 76, root_did)
    struct.pack_into('<i', entry, 116, first_sid)
    struct.pack_into('<i', entry, 120, tot_size)
    return bytes(entry)


def make_ole2_with_workbook(fragmented=False):
    """
    Build a minimal OLE2 compound document containing a 'Workbook' stream.

    If fragmented=True, the Workbook stream occupies non-contiguous sectors so
    that locate_named_stream returns a new bytes object (self.mem is not
    self.filestr after biff2_8_load).
    """
    # Sector layout (contiguous):
    #   Sector 0: SAT
    #   Sector 1: Directory
    #   Sector 2: Workbook data (sector 1 of stream)
    #
    # Sector layout (fragmented):
    #   Sector 0: SAT
    #   Sector 1: Directory
    #   Sector 2: Workbook data (sector 1 of stream)
    #   Sector 3: padding (unused)
    #   Sector 4: Workbook data (sector 2 of stream)

    if fragmented:
        workbook_first_sid = 2
        workbook_tot_size = SECTOR_SIZE + 100  # spans 2 sectors
        sat_ints = [SATSID, EOCSID, 4, FREESID, EOCSID] + [FREESID] * (SECTOR_SIZE // 4 - 5)
        num_data_sectors = 5  # SAT + Dir + WB1 + padding + WB2
    else:
        workbook_first_sid = 2
        workbook_tot_size = SECTOR_SIZE
        sat_ints = [SATSID, EOCSID, EOCSID] + [FREESID] * (SECTOR_SIZE // 4 - 3)
        num_data_sectors = 3  # SAT + Dir + WB

    # Header
    header = bytearray(512)
    header[0:8] = compdoc.SIGNATURE
    struct.pack_into('<H', header, 24, 0x003E)   # revision
    struct.pack_into('<H', header, 26, 0x0003)   # version
    header[28:30] = b'\xFE\xFF'                  # little-endian marker
    struct.pack_into('<H', header, 30, 9)         # ssz = 9 → sector size = 512
    struct.pack_into('<H', header, 32, 6)         # sssz = 6 → short sector size = 64
    struct.pack_into('<i', header, 44, 1)         # SAT_tot_secs = 1
    struct.pack_into('<i', header, 48, 1)         # dir_first_sec_sid = 1
    struct.pack_into('<i', header, 56, 0)         # min_size_std_stream = 0 (use SAT for all)
    struct.pack_into('<i', header, 60, EOCSID)    # SSAT_first_sec_sid
    struct.pack_into('<i', header, 64, 0)         # SSAT_tot_secs
    struct.pack_into('<i', header, 68, EOCSID)    # MSATX_first_sec_sid
    struct.pack_into('<i', header, 72, 0)         # MSATX_tot_secs
    # MSAT: sector 0 is the SAT; remaining 108 entries are FREESID
    struct.pack_into('<i', header, 76, 0)
    for i in range(1, 109):
        struct.pack_into('<i', header, 76 + i * 4, FREESID)

    # SAT sector (sector 0)
    sat_sector = bytearray(SECTOR_SIZE)
    for i, sid in enumerate(sat_ints):
        struct.pack_into('<i', sat_sector, i * 4, sid)

    # Directory sector (sector 1): root entry + Workbook entry
    dir_sector = bytearray(SECTOR_SIZE)
    root = _pack_dir_entry('Root Entry', 5, 1, -1, -1, 1, EOCSID, 0)
    wb   = _pack_dir_entry('Workbook',   2, 1, -1, -1, -1, workbook_first_sid, workbook_tot_size)
    dir_sector[0:128]   = root
    dir_sector[128:256] = wb

    # Workbook data (zeroed; biff2_8_load only sets self.mem – does not parse)
    data = bytes(header) + bytes(sat_sector) + bytes(dir_sector)
    if fragmented:
        data += bytes(SECTOR_SIZE)  # sector 2: WB part 1
        data += bytes(SECTOR_SIZE)  # sector 3: padding
        data += bytes(SECTOR_SIZE)  # sector 4: WB part 2
    else:
        data += bytes(SECTOR_SIZE)  # sector 2: WB data

    return data


def make_ole2_no_workbook():
    """
    Build a minimal OLE2 compound document with no 'Workbook' or 'Book' stream.
    The root entry has no children (root_DID = -1).
    """
    header = bytearray(512)
    header[0:8] = compdoc.SIGNATURE
    struct.pack_into('<H', header, 24, 0x003E)
    struct.pack_into('<H', header, 26, 0x0003)
    header[28:30] = b'\xFE\xFF'
    struct.pack_into('<H', header, 30, 9)
    struct.pack_into('<H', header, 32, 6)
    struct.pack_into('<i', header, 44, 1)
    struct.pack_into('<i', header, 48, 1)
    struct.pack_into('<i', header, 56, 0)
    struct.pack_into('<i', header, 60, EOCSID)
    struct.pack_into('<i', header, 64, 0)
    struct.pack_into('<i', header, 68, EOCSID)
    struct.pack_into('<i', header, 72, 0)
    struct.pack_into('<i', header, 76, 0)
    for i in range(1, 109):
        struct.pack_into('<i', header, 76 + i * 4, FREESID)

    sat_ints = [SATSID, EOCSID] + [FREESID] * (SECTOR_SIZE // 4 - 2)
    sat_sector = bytearray(SECTOR_SIZE)
    for i, sid in enumerate(sat_ints):
        struct.pack_into('<i', sat_sector, i * 4, sid)

    dir_sector = bytearray(SECTOR_SIZE)
    # Root entry with root_DID = -1 (no children)
    root = _pack_dir_entry('Root Entry', 5, 1, -1, -1, -1, EOCSID, 0)
    dir_sector[0:128] = root

    return bytes(header) + bytes(sat_sector) + bytes(dir_sector)


# ---------------------------------------------------------------------------
# Tests for biff2_8_load uncovered lines
# ---------------------------------------------------------------------------

class TestBiff28LoadUncoveredLines:

    def test_empty_file_raises_xlrderror(self, tmp_path):
        """Line 621: file of size 0 raises XLRDError."""
        empty = tmp_path / 'empty.xls'
        empty.write_bytes(b'')

        bk = Book()
        with pytest.raises(XLRDError, match="File size is 0 bytes"):
            bk.biff2_8_load(filename=str(empty))

    def test_load_from_file_no_mmap(self, tmp_path):
        """Lines 626-627: use_mmap=False reads file contents via f.read()."""
        from tests.test_book import make_minimal_xls
        xls_path = tmp_path / 'test.xls'
        xls_path.write_bytes(make_minimal_xls(['Sheet1']))

        bk = Book()
        bk.biff2_8_load(filename=str(xls_path), use_mmap=False)

        assert isinstance(bk.filestr, bytes)
        assert bk.stream_len == len(bk.filestr)

    def test_ole2_with_workbook_stream(self):
        """Lines 637, 639-643: OLE2 compound doc with Workbook stream is loaded."""
        ole2_data = make_ole2_with_workbook(fragmented=False)

        bk = Book()
        bk.biff2_8_load(file_contents=ole2_data, logfile=io.StringIO())

        # After a successful OLE2 parse, self.mem should be set
        assert bk.mem is not None
        assert bk.stream_len > 0

    def test_ole2_no_workbook_stream_raises(self):
        """Lines 644-645: OLE2 without Workbook/Book stream raises XLRDError."""
        ole2_data = make_ole2_no_workbook()

        bk = Book()
        with pytest.raises(XLRDError, match="Can't find workbook in OLE2 compound document"):
            bk.biff2_8_load(file_contents=ole2_data, logfile=io.StringIO())

    def test_ole2_fragmented_stream_from_file_contents(self):
        """Lines 646-647, 650: fragmented OLE2 stream yields self.mem != self.filestr."""
        ole2_data = make_ole2_with_workbook(fragmented=True)

        bk = Book()
        bk.biff2_8_load(file_contents=ole2_data, logfile=io.StringIO())

        # The stream was fragmented, so mem must be a new bytes object
        assert bk.mem is not ole2_data
        # filestr is reset to b'' (line 650)
        assert bk.filestr == b''

    def test_ole2_fragmented_stream_with_mmap_closes_filestr(self, tmp_path):
        """Lines 647-650: mmap filestr is closed when mem comes from fragmented OLE2."""
        import mmap
        ole2_data = make_ole2_with_workbook(fragmented=True)
        xls_path = tmp_path / 'fragmented.xls'
        xls_path.write_bytes(ole2_data)

        bk = Book()
        bk.biff2_8_load(filename=str(xls_path), use_mmap=True, logfile=io.StringIO())

        # filestr (the mmap) should have been closed and replaced with b''
        assert bk.filestr == b''
        assert bk.mem is not None
