#!/usr/bin/env python3
"""
build_vba.py — BajaTax v4
Crea un nuevo vbaProject.bin desde cero con todos los módulos VBA.
No requiere Excel abierto ni permisos especiales.
Requiere: pip install olefile oletools openpyxl
"""

from pathlib import Path
import json

ROOT     = Path(__file__).parent.parent
config   = json.loads((ROOT / "bajatax.config.json").read_text(encoding="utf-8"))
SRC_FILE = str(ROOT / config["xlsm_source"])
DST_FILE = str(ROOT / config["xlsm_output"])
VBA_DIR  = str(ROOT / config["vba_modules_dir"])
import struct, zipfile, io, os, sys, shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VBA_DIR  = os.path.join(BASE_DIR, "VBA_CODIGO")
SRC_FILE = os.path.join(BASE_DIR, config["xlsm_source"])
DST_FILE = os.path.join(BASE_DIR, config["xlsm_output"])

# ── OLE2 Constants ──────────────────────────────────────────────────
SECTOR_SIZE        = 512
MINI_SECTOR_SIZE   = 64
MINI_CUTOFF        = 4096
ENDOFCHAIN         = 0xFFFFFFFE
FREESECT           = 0xFFFFFFFF
FATSECT            = 0xFFFFFFFD
DIFSECT            = 0xFFFFFFFC
STGTY_EMPTY        = 0
STGTY_STORAGE      = 1
STGTY_STREAM       = 2
STGTY_ROOT         = 5
NOSTREAM           = 0xFFFFFFFF

# ── MS-OVBA Compression (raw chunks only) ──────────────────────────
def ovba_compress(source: bytes) -> bytes:
    """Compress using MS-OVBA raw chunks (uncompressed, always valid).
    Raw chunk header format per MS-OVBA spec §2.4.2:
      bit15=0 (raw, CompressedChunkFlag), bits[14:12]=0b011 (signature), bits[11:0]=0xFFF (4095)
      → 0_011_1111_1111_1111 = 0x3FFF  (little-endian: FF 3F)
    """
    out = bytearray([0x01])   # CompressedContainer Signature byte
    pos = 0
    while pos < len(source):
        chunk = source[pos:pos + 4096]
        padded = chunk + b'\x00' * (4096 - len(chunk))
        # Raw chunk header: 0x3FFF  (bit15=0=raw, bits14:12=011, bits11:0=4095)
        out += struct.pack('<H', 0x3FFF)
        out += padded
        pos += 4096
    return bytes(out)

# ── dir stream builder ──────────────────────────────────────────────
def get_src_dir_proj_info() -> bytes:
    """
    Extrae el prefijo del dir stream de v2 (INFO + REFERENCES section),
    que son todos los bytes ANTES del record PROJMODULES (0x000F).
    Esto incluye: SYSKIND, LCID, VERSION, CONSTANTS, y los REFERENCE records
    criticos (Excel Object Library, Office Library, etc.).
    Sin estos references, Excel rechaza el VBA project completo.
    """
    try:
        from oletools.olevba import decompress_stream as _dc
        ole = _open_src_ole()
        raw = _dc(ole.openstream('VBA/dir').read())
        ole.close()
        # Walk records to find byte position of PROJMODULES (0x000F)
        pos = 0
        while pos < len(raw) - 5:
            rid = struct.unpack_from('<H', raw, pos)[0]
            if rid == 0x000F:   # PROJMODULES – stop here
                info_bytes = raw[:pos]
                print(f'  dir INFO+REFS section: {len(info_bytes)} bytes (del fuente)')
                return info_bytes
            if rid == 0x0009:   # VERSION: no size field, special 4+2 bytes
                pos += 2 + 4 + 4 + 2
            else:
                sz = struct.unpack_from('<I', raw, pos + 2)[0]
                pos += 2 + 4 + sz
    except Exception as e:
        print(f'  AVISO get_src_dir_proj_info: {e}')
    return None   # None → fallback to built-in info section


def build_dir_stream(modules, src_proj_info: bytes = None) -> bytes:
    """
    Build the MS-OVBA dir stream (uncompressed) describing all VBA modules.

    src_proj_info: bytes from get_src_dir_proj_info() – the INFO+REFERENCES
    section copied verbatim from v2.  This is CRITICAL: it contains the
    REFERENCEREGISTERED/REFERENCECONTROL records for the Excel Object Library
    and Office Library.  Without them Excel rejects the VBA project entirely.

    modules = list of dict with keys: name, stream_name, type ('std'|'doc')
    """
    def w16(v): return struct.pack('<H', v)
    def w32(v): return struct.pack('<I', v)
    def vrec(rid, data):
        return struct.pack('<HI', rid, len(data)) + data

    out = bytearray()

    if src_proj_info:
        # Use v2's project info section verbatim (SYSKIND → ... → just before PROJMODULES)
        out += src_proj_info
    else:
        # Fallback: minimal project info (no library references — may not work)
        out += vrec(0x0001, w32(0x00000003))   # SYSKIND=Mac
        out += vrec(0x0002, w32(0x00000409))   # LCID
        out += vrec(0x0014, w32(0x00000409))   # LCIDINVOKE
        out += vrec(0x0003, w16(0x04E4))       # CODEPAGE=1252
        out += vrec(0x0004, b'VBAProject')     # NAME
        out += vrec(0x0005, b'')               # DOCSTRING
        out += vrec(0x0040, b'')               # DOCSTRINGUNICODE
        out += vrec(0x0006, b'')               # HELPFILEPATH1
        out += vrec(0x003D, b'')               # HELPFILEPATH2
        out += vrec(0x0007, w32(0))            # HELPCONTEXT
        out += vrec(0x0008, w32(0))            # LIBFLAGS
        out += struct.pack('<HI', 0x0009, 4) + w32(0x61)   # VERSION major
        out += struct.pack('<H', 0x5F4F)                   # VERSION minor
        out += vrec(0x000C, b'')               # CONSTANTS
        out += vrec(0x003C, b'')               # CONSTANTS_UNI

    # PROJECTMODULES section
    mod_count = len(modules)
    out += struct.pack('<HI', 0x000F, 2) + w16(mod_count)  # PROJMODULES count
    out += struct.pack('<HI', 0x0013, 2) + w16(0xFFFF)     # PROJCOOKIE

    for m in modules:
        name_bytes   = m['name'].encode('latin-1')
        name_utf16   = m['name'].encode('utf-16-le')
        stream_bytes = m.get('stream_name', m['name']).encode('latin-1')
        stream_utf16 = m.get('stream_name', m['name']).encode('utf-16-le')

        out += vrec(0x0019, name_bytes)         # MODULENAME
        out += vrec(0x0047, name_utf16)         # MODULENAMEUNICODE
        out += vrec(0x001A, stream_bytes)       # MODULESTREAMNAME
        out += vrec(0x0032, stream_utf16)       # MODULESTREAMNAMEUNICODE
        out += vrec(0x001C, b'')                # MODULEDOCSTRING
        out += vrec(0x0048, b'')                # MODULEDOCSTRINGUNICODE
        out += vrec(0x0031, w32(0))             # MODULEOFFSET (0 = source at byte 0)
        out += vrec(0x001E, w32(0))             # MODULEHELPCONTEXT
        out += struct.pack('<HI', 0x002C, 2) + w16(0xFFFF)  # MODULECOOKIE
        if m.get('type', 'std') == 'std':
            out += struct.pack('<HI', 0x0021, 0)  # MODULETYPE standard
        else:
            out += struct.pack('<HI', 0x0022, 0)  # MODULETYPE document
        out += struct.pack('<HI', 0x002B, 0)   # MODULETERMINATOR

    out += struct.pack('<HI', 0x0010, 0)       # PROJECTTERMINATOR
    return bytes(out)

# ── PROJECT stream ─────────────────────────────────────────────────
def build_project_stream(modules, proj_id="{12345678-ABCD-EF01-2345-6789ABCDEF01}") -> bytes:
    lines = [f'ID="{proj_id}"']
    for m in modules:
        entry_type = 'Document' if m.get('type') == 'doc' else 'Module'
        if entry_type == 'Document':
            lines.append(f"Document={m['name']}/&H00000000")
        else:
            lines.append(f"Module={m['name']}")
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG=""',
        'DPB=""',
        'GC=""',
        '',
        '[Host Extender Info]',
        '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
    ]
    return ('\r\n'.join(lines) + '\r\n').encode('latin-1')

# ── OLE2 Writer ─────────────────────────────────────────────────────
class OLE2Writer:
    """Minimal Compound File Binary (OLE2) writer."""

    def __init__(self):
        self._streams = {}   # path -> bytes  (path = list of str)
        self._storages = []  # list of list of str

    def add_storage(self, path):
        self._storages.append(list(path))

    def add_stream(self, path, data: bytes):
        self._streams[tuple(path)] = data

    # ── helpers ────────────────────────────────────────────────────
    def _sectors(self, data):
        """Split data into 512-byte sectors, padded."""
        sectors = []
        for i in range(0, len(data), SECTOR_SIZE):
            chunk = data[i:i + SECTOR_SIZE]
            sectors.append(chunk + b'\x00' * (SECTOR_SIZE - len(chunk)))
        return sectors

    def write(self, output_path):
        # ── 1. Flatten all entries (root, storages, streams) ───────
        # We build a flat directory list:
        # [0] = root
        # [1] = VBA storage (if any)
        # [2..] = streams

        # Collect streams in a predictable order
        # Root entry + storages come first, then streams sorted
        all_entries = []  # list of (name, type, data_or_none, parent_path)

        # Root (type 5)
        all_entries.append(('Root Entry', STGTY_ROOT, None, []))

        # Gather unique storages
        unique_storages = []
        seen = set()
        for p in self._storages:
            for i in range(1, len(p) + 1):
                key = tuple(p[:i])
                if key not in seen:
                    unique_storages.append(list(p[:i]))
                    seen.add(key)
        # Also add storages implied by stream paths
        for path in self._streams:
            for i in range(1, len(path)):
                key = tuple(path[:i])
                if key not in seen:
                    unique_storages.append(list(path[:i]))
                    seen.add(key)
        for s in unique_storages:
            all_entries.append((s[-1], STGTY_STORAGE, None, s[:-1]))

        # Add streams
        for path, data in sorted(self._streams.items()):
            all_entries.append((path[-1], STGTY_STREAM, data, list(path[:-1])))

        # ── 2. Determine data placement ────────────────────────────
        # Small streams (< MINI_CUTOFF) go into mini-stream
        # Large streams go into regular sectors

        # First allocate regular sector space for all large streams
        # Sector 0 is the header (not a real sector), sectors start at 0
        regular_data_sectors = []  # (entry_idx, [sector_list])
        mini_data = bytearray()    # concatenated mini-stream data
        mini_stream_entries = []   # (entry_idx, mini_start, size)

        # We'll build everything in passes
        # Pass 1: figure out which streams go where
        stream_info = {}  # entry_idx -> dict with placement info
        for idx, (name, etype, data, parent) in enumerate(all_entries):
            if etype != STGTY_STREAM or data is None:
                continue
            if len(data) >= MINI_CUTOFF:
                stream_info[idx] = {'large': True, 'size': len(data), 'data': data}
            else:
                stream_info[idx] = {'large': False, 'size': len(data), 'data': data}

        # Assign mini-stream positions for small streams
        # IMPORTANT: zero-byte streams must NOT get a mini-sector (use ENDOFCHAIN)
        mini_offset = 0
        for idx in sorted(stream_info.keys()):
            info = stream_info[idx]
            if not info['large']:
                if info['size'] == 0:
                    info['mini_start'] = ENDOFCHAIN  # zero-byte stream → ENDOFCHAIN
                    continue  # don't allocate any mini-sector
                info['mini_start'] = mini_offset // MINI_SECTOR_SIZE
                padded_size = ((info['size'] + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE) * MINI_SECTOR_SIZE
                mini_data += info['data'] + b'\x00' * (padded_size - info['size'])
                mini_offset += padded_size

        # ── 3. Layout sectors ───────────────────────────────────────
        # Sector layout: [FAT sectors] [Dir sectors] [mini-FAT sectors]
        # [mini-stream container sectors] [large stream sectors]

        # Count directory entries (4 per sector)
        n_dir_entries = len(all_entries)
        n_dir_sectors = (n_dir_entries + 3) // 4

        # We'll do 2 passes: estimate FAT size, then actually build
        # For simplicity, do a fixed layout:
        # Sectors: [FAT_0][DIR_0..n][MINIFAT_0..m][MINISTREAM_0..k][LARGE_0..N]

        # First, estimate total sectors
        def estimate_sectors():
            # Large streams
            large_sectors = sum(
                (info['size'] + SECTOR_SIZE - 1) // SECTOR_SIZE
                for info in stream_info.values() if info['large']
            )
            # Mini stream container (mini_data in regular sectors)
            mini_container = (len(mini_data) + SECTOR_SIZE - 1) // SECTOR_SIZE
            # Mini FAT
            total_mini_sectors = (len(mini_data) + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE
            n_minifat = (total_mini_sectors * 4 + SECTOR_SIZE - 1) // SECTOR_SIZE
            # Dir sectors
            total_known = large_sectors + mini_container + n_minifat + n_dir_sectors
            # FAT sectors (each covers 128 sectors)
            n_fat = max(1, (total_known + 128) // 128 + 1)
            return n_fat, n_dir_sectors, n_minifat, mini_container, large_sectors

        n_fat, n_dir, n_minifat, n_mini_container, n_large = estimate_sectors()

        # Refine: with FAT sectors included
        total_sectors = n_fat + n_dir + n_minifat + n_mini_container + n_large
        n_fat = max(1, (total_sectors * 4 + SECTOR_SIZE - 1) // SECTOR_SIZE // 1)
        n_fat = (total_sectors + 127) // 128 + 1

        # Assign sector numbers
        fat_start    = 0
        dir_start    = fat_start + n_fat
        minifat_start = dir_start + n_dir
        ministream_start = minifat_start + n_minifat
        large_start  = ministream_start + n_mini_container

        # Assign sectors to large streams
        cur_sector = large_start
        for idx in sorted(stream_info.keys()):
            info = stream_info[idx]
            if not info['large']:
                continue
            n = (info['size'] + SECTOR_SIZE - 1) // SECTOR_SIZE
            info['sector_start'] = cur_sector
            info['sectors'] = list(range(cur_sector, cur_sector + n))
            cur_sector += n

        total_sectors_final = cur_sector

        # Build FAT
        fat = [FREESECT] * ((total_sectors_final + 127) // 128 * 128)
        # Mark FAT sectors
        for s in range(fat_start, fat_start + n_fat):
            fat[s] = FATSECT
        # Mark directory sectors
        for i, s in enumerate(range(dir_start, dir_start + n_dir)):
            fat[s] = dir_start + i + 1 if i < n_dir - 1 else ENDOFCHAIN
        # Mark mini-FAT sectors
        if n_minifat > 0:
            for i, s in enumerate(range(minifat_start, minifat_start + n_minifat)):
                fat[s] = minifat_start + i + 1 if i < n_minifat - 1 else ENDOFCHAIN
        # Mark mini-stream container sectors
        if n_mini_container > 0:
            for i, s in enumerate(range(ministream_start, ministream_start + n_mini_container)):
                fat[s] = ministream_start + i + 1 if i < n_mini_container - 1 else ENDOFCHAIN
        # Mark large stream sectors
        for idx in sorted(stream_info.keys()):
            info = stream_info[idx]
            if not info['large']:
                continue
            secs = info['sectors']
            for i, s in enumerate(secs):
                fat[s] = secs[i + 1] if i < len(secs) - 1 else ENDOFCHAIN

        # Build mini-FAT
        total_mini_sectors = (len(mini_data) + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE if mini_data else 0
        minifat = [FREESECT] * max(128, ((total_mini_sectors + 127) // 128) * 128)
        for idx in sorted(stream_info.keys()):
            info = stream_info[idx]
            if info['large']:
                continue
            n_ms = (info['size'] + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE
            ms_start = info['mini_start']
            for i in range(n_ms):
                minifat[ms_start + i] = (ms_start + i + 1) if i < n_ms - 1 else ENDOFCHAIN

        # ── 4. Build directory entries ─────────────────────────────
        # Build parent_path → children mapping for tree building
        from collections import defaultdict
        children_map = defaultdict(list)
        entry_path = {}
        entry_path[0] = []  # root

        for idx, (name, etype, data, parent) in enumerate(all_entries):
            path = list(parent) + [name] if parent else [name] if idx > 0 else []
            entry_path[idx] = path
            parent_key = tuple(parent) if parent else ()
            if idx > 0:
                # Find parent entry index
                parent_idx = 0  # default root
                for pidx, (pname, petype, pdata, pparent) in enumerate(all_entries):
                    if pparent == [] and parent == []:
                        break
                    ep = list(pparent) + [pname] if pparent else [pname] if pidx > 0 else []
                    if ep == parent:
                        parent_idx = pidx
                        break
                children_map[parent_idx].append(idx)

        def build_red_black(children):
            """Simple: just chain siblings left→right."""
            if not children:
                return NOSTREAM, {}
            # Sort by name (uppercase) for compatibility
            children_sorted = sorted(children, key=lambda i: all_entries[i][0].upper())
            mid = len(children_sorted) // 2
            node_left  = {children_sorted[mid]: NOSTREAM}
            node_right = {children_sorted[mid]: NOSTREAM}
            # Left chain
            for i in range(mid - 1, -1, -1):
                node_left[children_sorted[i]] = NOSTREAM
                node_right[children_sorted[i]] = children_sorted[i + 1] if i < mid - 1 else children_sorted[i + 1]
            # Right chain
            for i in range(mid + 1, len(children_sorted)):
                node_left[children_sorted[i]] = NOSTREAM
                node_right[children_sorted[i]] = NOSTREAM
                if i > mid + 1:
                    node_right[children_sorted[i - 1]] = children_sorted[i]
            return children_sorted[mid], {}

        # Simpler: just chain all children as a linear right-sibling chain from the first child
        def assign_siblings(children):
            """Returns (first_child_id, {idx: (left, right)})"""
            if not children:
                return NOSTREAM, {}
            children_sorted = sorted(children, key=lambda i: all_entries[i][0].upper())
            siblings = {}
            for i, cidx in enumerate(children_sorted):
                right = children_sorted[i + 1] if i + 1 < len(children_sorted) else NOSTREAM
                siblings[cidx] = (NOSTREAM, right)
            return children_sorted[0], siblings

        # Build directory entries bytes
        dir_bytes = bytearray()
        all_siblings = {}
        all_first_child = {}
        for pidx in range(len(all_entries)):
            ch = children_map.get(pidx, [])
            first, sibs = assign_siblings(ch)
            all_first_child[pidx] = first
            all_siblings.update(sibs)

        # VBA storage CLSID: real Excel files use all-zeros (empty CLSID)
        VBA_CLSID = b'\x00' * 16

        for idx, (name, etype, data, parent) in enumerate(all_entries):
            left_sib, right_sib = all_siblings.get(idx, (NOSTREAM, NOSTREAM))
            child_id = all_first_child.get(idx, NOSTREAM)

            # Determine start sector and size
            if etype == STGTY_ROOT:
                start = ministream_start if n_mini_container > 0 else ENDOFCHAIN
                size = len(mini_data)
                clsid = b'\x00' * 16
            elif etype == STGTY_STORAGE:
                start = 0          # real Excel files use 0, not ENDOFCHAIN, for storage start
                size = 0
                clsid = VBA_CLSID  # same empty CLSID for all storages
            elif etype == STGTY_STREAM:
                info = stream_info.get(idx)
                if info is None:
                    start, size = ENDOFCHAIN, 0
                elif info['large']:
                    start = info['sector_start']
                    size = info['size']
                elif info['size'] == 0:
                    start, size = ENDOFCHAIN, 0  # zero-byte stream must use ENDOFCHAIN
                else:
                    start = info['mini_start']
                    size = info['size']
                clsid = b'\x00' * 16
            else:
                start, size, clsid = ENDOFCHAIN, 0, b'\x00' * 16

            name_enc = name.encode('utf-16-le')
            name_bytes_raw = (name_enc + b'\x00\x00')[:64]
            name_bytes_raw = name_bytes_raw + b'\x00' * (64 - len(name_bytes_raw))
            name_len = min(len(name_enc) + 2, 64)

            entry = bytearray(128)
            entry[0:64]   = name_bytes_raw
            struct.pack_into('<H', entry, 64, name_len)
            entry[66] = etype
            entry[67] = 1  # black node
            struct.pack_into('<I', entry, 68, left_sib)
            struct.pack_into('<I', entry, 72, right_sib)
            struct.pack_into('<I', entry, 76, child_id)
            entry[80:96]  = clsid
            struct.pack_into('<I', entry, 96, 0)
            struct.pack_into('<Q', entry, 100, 0)
            struct.pack_into('<Q', entry, 108, 0)
            struct.pack_into('<I', entry, 116, start)
            struct.pack_into('<I', entry, 120, size)
            struct.pack_into('<I', entry, 124, 0)
            dir_bytes += entry

        # Pad directory to full sectors with proper STGTY_EMPTY entries
        # (sibling/child fields MUST be NOSTREAM=0xFFFFFFFF, not 0)
        while len(dir_bytes) < n_dir * SECTOR_SIZE:
            empty = bytearray(128)
            # name_len=0, etype=0 (STGTY_EMPTY) already zero
            struct.pack_into('<I', empty, 68, NOSTREAM)   # left sib
            struct.pack_into('<I', empty, 72, NOSTREAM)   # right sib
            struct.pack_into('<I', empty, 76, NOSTREAM)   # child
            struct.pack_into('<I', empty, 116, ENDOFCHAIN) # start
            dir_bytes += bytes(empty)

        # ── 5. Build FAT sector bytes ──────────────────────────────
        fat_bytes = bytearray()
        for v in fat:
            fat_bytes += struct.pack('<I', v)
        fat_bytes = fat_bytes[:n_fat * SECTOR_SIZE]
        fat_bytes += b'\xff' * (n_fat * SECTOR_SIZE - len(fat_bytes))

        # ── 6. Build mini-FAT bytes ────────────────────────────────
        minifat_bytes = bytearray()
        for v in minifat:
            minifat_bytes += struct.pack('<I', v)
        minifat_bytes = minifat_bytes[:max(n_minifat, 1) * SECTOR_SIZE]
        minifat_bytes += b'\xff' * (max(n_minifat, 1) * SECTOR_SIZE - len(minifat_bytes))

        # ── 7. Build mini-stream container ────────────────────────
        mini_container_bytes = bytearray(mini_data)
        target = n_mini_container * SECTOR_SIZE
        mini_container_bytes += b'\x00' * max(0, target - len(mini_container_bytes))

        # ── 8. Build large stream data ────────────────────────────
        large_stream_bytes = bytearray()
        for idx in sorted(stream_info.keys()):
            info = stream_info[idx]
            if not info['large']:
                continue
            data = info['data']
            padded_size = ((len(data) + SECTOR_SIZE - 1) // SECTOR_SIZE) * SECTOR_SIZE
            large_stream_bytes += data + b'\x00' * (padded_size - len(data))

        # ── 9. Build header (512 bytes) ───────────────────────────
        header = bytearray(512)
        # Magic
        header[0:8] = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
        # CLSID = 0
        # Minor version = 0x003E
        struct.pack_into('<H', header, 24, 0x003E)
        # Major version = 0x0003
        struct.pack_into('<H', header, 26, 0x0003)
        # Byte order = 0xFFFE (little-endian)
        struct.pack_into('<H', header, 28, 0xFFFE)
        # Sector size shift = 9 (2^9 = 512)
        struct.pack_into('<H', header, 30, 0x0009)
        # Mini-sector size shift = 6 (2^6 = 64)
        struct.pack_into('<H', header, 32, 0x0006)
        # Reserved: 6 bytes (already zero)
        # Number of FAT sectors
        struct.pack_into('<I', header, 44, n_fat)
        # First directory sector
        struct.pack_into('<I', header, 48, dir_start)
        # Transaction signature (0)
        struct.pack_into('<I', header, 52, 0)
        # Mini stream cutoff
        struct.pack_into('<I', header, 56, MINI_CUTOFF)
        # First mini-FAT sector
        struct.pack_into('<I', header, 60, minifat_start if n_minifat > 0 else ENDOFCHAIN)
        # Number of mini-FAT sectors
        struct.pack_into('<I', header, 64, n_minifat)
        # First DIFAT sector
        struct.pack_into('<I', header, 68, ENDOFCHAIN)
        # Number of DIFAT sectors
        struct.pack_into('<I', header, 72, 0)
        # DIFAT array in header (109 entries × 4 bytes)
        # First entries point to FAT sectors
        for i in range(min(n_fat, 109)):
            struct.pack_into('<I', header, 76 + i * 4, fat_start + i)
        # Remaining DIFAT entries = FREESECT
        for i in range(n_fat, 109):
            struct.pack_into('<I', header, 76 + i * 4, FREESECT)

        # ── 10. Concatenate everything ────────────────────────────
        # Order: header | FAT | Dir | MiniFAT | MiniStream | LargeStreams
        result  = bytes(header)
        result += bytes(fat_bytes)
        result += bytes(dir_bytes)
        result += bytes(minifat_bytes) if n_minifat > 0 else b''
        result += bytes(mini_container_bytes) if n_mini_container > 0 else b''
        result += bytes(large_stream_bytes)

        with open(output_path, 'wb') as f:
            f.write(result)
        return output_path


# ── Streams copiados del archivo fuente (v2) ────────────────────────
def _open_src_ole():
    """Abre el OleFileIO del vbaProject.bin del archivo fuente v2."""
    import zipfile as _zf, olefile as _ole, io as _io
    with _zf.ZipFile(SRC_FILE) as z:
        vba_bin = z.read('xl/vbaProject.bin')
    return _ole.OleFileIO(_io.BytesIO(vba_bin))


def build_vba_project_stream() -> bytes:
    """
    Retorna stub de _VBA_PROJECT (8 bytes).
    No copiamos el p-code de v2 porque fue compilado para 14 módulos con
    MODULE.OFFSET distintos.  Con el stub Excel recompila desde el código fuente.
    """
    print(f'  _VBA_PROJECT: stub 8 bytes (fuerza recompilación)')
    return b'\xCC\x61\xFF\xFF\x00\x00\x00\x00'


def get_src_extra_streams() -> dict:
    """
    Copia del archivo fuente:
    - PROJECTwm  (DEBE estar presente según MS-OVBA §2.3.4.16)
    - __SRP_*    (caché de bytecode; Excel los regenera si no coinciden)
    Retorna dict: stream_name -> bytes (todos bajo VBA storage)
    """
    result = {}
    try:
        ole = _open_src_ole()
        # PROJECTwm (bajo Root, no bajo VBA)
        try:
            result['__projectwm__'] = ole.openstream('PROJECTwm').read()
            print(f'  PROJECTwm: {len(result["__projectwm__"])} bytes (del fuente)')
        except Exception:
            result['__projectwm__'] = b'\x00'   # fallback: solo el terminador
            print(f'  PROJECTwm: fallback = 1 byte')
        # __SRP_* (bajo VBA)
        for entry in ole.listdir(streams=True):
            name = entry[-1]
            if name.startswith('__SRP_'):
                data = ole.openstream(entry).read()
                result[name] = data
                print(f'  {name}: {len(data)} bytes (del fuente)')
        ole.close()
    except Exception as e:
        print(f'  AVISO get_src_extra_streams: {e}')
        if '__projectwm__' not in result:
            result['__projectwm__'] = b'\x00'
    return result


def get_src_guid() -> str:
    """Extrae el ID (GUID) del PROJECT stream del archivo fuente."""
    import re
    try:
        ole = _open_src_ole()
        proj = ole.openstream('PROJECT').read().decode('latin-1')
        ole.close()
        m = re.search(r'ID="(\{[^}]+\})"', proj)
        if m:
            return m.group(1)
    except Exception:
        pass
    return '{12345678-ABCD-EF01-2345-6789ABCDEF01}'


# ── Main ────────────────────────────────────────────────────────────
def leer_bas(fname):
    path = os.path.join(VBA_DIR, fname)
    if not os.path.exists(path):
        print(f"  ADVERTENCIA: No se encuentra {fname}")
        return None
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()

def strip_attribute_vb_name(src):
    """Remove the first Attribute VB_Name line (used in .cls/.bas headers)."""
    lines = src.split('\n')
    clean = [l for l in lines if not l.startswith('Attribute VB_Name')]
    return '\n'.join(clean)

def main():
    print("=" * 60)
    print("  BajaTax v4 — Constructor OLE2 VBA (Python puro)")
    print("=" * 60)

    # ── 1. Copiar base ─────────────────────────────────────────────
    if not os.path.exists(SRC_FILE):
        print(f"ERROR: No se encuentra {SRC_FILE}"); sys.exit(1)
    shutil.copy2(SRC_FILE, DST_FILE)
    print(f"✓ Copia creada: {os.path.basename(DST_FILE)}")

    # ── 2. Crear carpetas ──────────────────────────────────────────
    for d in ['IMPORTAR', 'SALIDA_PDF', 'LOGOS']:
        os.makedirs(os.path.join(BASE_DIR, d), exist_ok=True)
    print("✓ Carpetas IMPORTAR / SALIDA_PDF / LOGOS")

    # ── 3. Leer módulos .bas ───────────────────────────────────────
    print("\n► Leyendo módulos VBA...")

    # Standard modules: (filename, module_name, internal_name)
    # internal_name = name used in OLE2 stream
    std_mods = [
        ("01_Mod_Sistema.bas",         "Mod_Sistema",          "Mod_Sistema"),
        ("02_Mod_ImportarArchivos.bas","Mod_ImportarArchivos",  "Mod_ImportarArchivos"),
        ("03_Mod_WhatsApp.bas",        "WhatsApp",              "WhatsApp"),
        ("04_Mod_PDF.bas",             "PDF",                   "PDF"),
        ("07_Mod_MasivoPDF.bas",       "Mod_MasivoPDF",         "Mod_MasivoPDF"),
        ("08_Mod_BuscadorCliente.bas", "Mod_BuscadorCliente",   "BuscadorCliente"),
        ("09_Mod_FormatoGlobal.bas",     "Mod_FormatoGlobal",       "Mod_FormatoGlobal"),
    ]
    # Sheet modules: (filename, module_name, stream_name)
    # IMPORTANT: module_name = codeName from sheet XML (must match for event binding)
    # From xl/worksheets/sheet*.xml codeName analysis of v2:
    #   Hoja2 = OPERACIONES (sheet3.xml), Hoja5 = DIRECTORIO (sheet5.xml),
    #   Hoja10 = BUSCADOR CLIENTE (sheet6.xml)
    #   Hoja1=Soportes, Hoja6=CONFIGURACION, Hoja3=REGISTROS, Hoja4=REPORTES CXC
    #   Hoja7=LOG ENVÍOS, Hoja8=LOG ENVIOS
    # For document modules, name MUST = stream_name = codeName for correct binding
    hoja_mods = [
        ("05_Hoja_OPERACIONES.bas",     "OPERACIONES",    "OPERACIONES"),
        ("06_Hoja_DIRECTORIO.bas",      "DIRECTORIO",    "DIRECTORIO"),
        ("10_Hoja_BuscadorCliente.bas", "BUSCADOR CLIENTE",   "BUSCADOR CLIENTE"),
        ("11_Hoja_REGISTROS.bas", "REGISTROS", "REGISTROS"),
        # Stubs for remaining sheets (name=codeName=stream_name)
        (None, "ThisWorkbook",  "ThisWorkbook"),
        (None, "Hoja1",         "Hoja1"),
        (None, "Hoja3",         "Hoja3"),
        (None, "Hoja4",         "Hoja4"),
        (None, "Hoja6",         "Hoja6"),
        (None, "Hoja7",         "Hoja7"),
        (None, "Hoja8",         "Hoja8"),
    ]

    modules_meta = []  # for dir stream and PROJECT stream
    module_streams = {}  # stream_name -> compressed_bytes

    # Process standard modules
    for fname, mod_name, stream_name in std_mods:
        src = leer_bas(fname)
        if src is None:
            print(f"  ✗ Falta {fname}, se omite")
            continue
        # Strip Attribute VB_Name line (it's part of the module metadata, not the code)
        src_clean = src
        compressed = ovba_compress(src_clean.encode('latin-1', errors='replace'))
        module_streams[stream_name] = compressed
        modules_meta.append({'name': mod_name, 'stream_name': stream_name, 'type': 'std'})
        print(f"  ✓ {mod_name} ({len(src)} bytes src → {len(compressed)} compressed)")

    # Process sheet modules
    for fname, mod_name, stream_name in hoja_mods:
        if fname:
            src = leer_bas(fname)
            if src is None:
                src = f'Attribute VB_Name = "{mod_name}"\n'
        else:
            src = f'Attribute VB_Name = "{mod_name}"\n'

        # For sheet modules, strip VB_Name attribute line from code
        # (it's stored separately in VBA dir stream)
        src_code = '\n'.join([l for l in src.split('\n')
                               if not l.startswith('Attribute VB_Name')])

        compressed = ovba_compress(src_code.encode('latin-1', errors='replace'))
        module_streams[stream_name] = compressed
        modules_meta.append({'name': mod_name, 'stream_name': stream_name, 'type': 'doc'})
        if fname:
            print(f"  ✓ Hoja CodeName={stream_name} ({len(src)} bytes)")
        else:
            print(f"  ○ Hoja CodeName={stream_name} (stub)")

    # ── 4. Build dir stream ────────────────────────────────────────
    print("\n► Construyendo dir stream...")
    # CRITICO: copiar la sección INFO+REFS del archivo fuente v2.
    # Sin los REFERENCE records (Excel Object Library, Office Library, etc.)
    # Excel rechaza el VBA project completo con el mensaje de reparación.
    src_proj_info = get_src_dir_proj_info()
    dir_raw = build_dir_stream(modules_meta, src_proj_info=src_proj_info)
    dir_compressed = ovba_compress(dir_raw)
    print(f"  dir: {len(dir_raw)} bytes → {len(dir_compressed)} compressed")

    # ── 5. Obtener streams extra del archivo fuente ────────────────
    print("\n► Copiando streams del fuente (v2)...")
    src_guid = get_src_guid()
    extra = get_src_extra_streams()
    projectwm_bytes = extra.pop('__projectwm__', b'\x00')
    srp_streams = extra   # keys = '__SRP_0', '__SRP_1', ...
    print(f"  GUID del fuente: {src_guid}")

    # ── 6. Build PROJECT stream ────────────────────────────────────
    project_bytes = build_project_stream(modules_meta, proj_id=src_guid)
    print(f"  PROJECT: {len(project_bytes)} bytes")

    # ── 7. Build OLE2 file ─────────────────────────────────────────
    print("\n► Construyendo vbaProject.bin...")
    writer = OLE2Writer()
    writer.add_storage(['VBA'])
    writer.add_stream(['VBA', '_VBA_PROJECT'], build_vba_project_stream())
    writer.add_stream(['VBA', 'dir'], dir_compressed)
    for stream_name, data in module_streams.items():
        writer.add_stream(['VBA', stream_name], data)
    # __SRP_* del fuente (caché de bytecode; Excel los regenera si son obsoletos)
    for srp_name, srp_data in srp_streams.items():
        writer.add_stream(['VBA', srp_name], srp_data)
    writer.add_stream(['PROJECT'], project_bytes)
    # PROJECTwm del fuente (MS-OVBA §2.3.4.16: MUST be present)
    writer.add_stream(['PROJECTwm'], projectwm_bytes)

    tmp_vba = os.path.join(BASE_DIR, "_tmp_vbaProject.bin")
    writer.write(tmp_vba)
    print(f"  ✓ vbaProject.bin temporal: {os.path.getsize(tmp_vba):,} bytes")

    # ── 8. Inject into destination xlsm ───────────────────────────
    print("\n► Inyectando en xlsm...")
    # Read the destination xlsm, replace vbaProject.bin, write back
    with open(DST_FILE, 'rb') as f:
        dst_bytes = f.read()

    with zipfile.ZipFile(io.BytesIO(dst_bytes), 'r') as zin:
        names = zin.namelist()
        files = {}
        for n in names:
            files[n] = zin.read(n)

    with open(tmp_vba, 'rb') as f:
        new_vba = f.read()

    files['xl/vbaProject.bin'] = new_vba

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            if name == 'xl/vbaProject.bin':
                # Usar DEFLATED igual que Excel (v2 también es DEFLATED)
                zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
            elif name.endswith('.bin'):
                zout.writestr(name, data, compress_type=zipfile.ZIP_STORED)
            else:
                zout.writestr(name, data)

    with open(DST_FILE, 'wb') as f:
        f.write(out_buf.getvalue())

    os.remove(tmp_vba)
    print(f"  ✓ {os.path.basename(DST_FILE)} actualizado")
    print(f"    Tamaño final: {os.path.getsize(DST_FILE):,} bytes")

    # ── 9. Verify with olefile ─────────────────────────────────────
    print("\n► Verificando estructura OLE2...")
    try:
        import olefile
        with zipfile.ZipFile(DST_FILE, 'r') as z:
            vba_check = z.read('xl/vbaProject.bin')
        ole = olefile.OleFileIO(io.BytesIO(vba_check))
        streams = ole.listdir(streams=True)
        print(f"  Streams encontrados: {len(streams)}")
        for s in streams:
            path = '/'.join(s)
            try:
                sz = ole.get_size(s)
                print(f"    {path} ({sz} bytes)")
            except:
                print(f"    {path}")
        ole.close()
        print("\n✓ ESTRUCTURA OLE2 VÁLIDA")
    except Exception as e:
        print(f"\n✗ Error en verificación: {e}")
        import traceback; traceback.print_exc()

    print("\n" + "=" * 60)
    print("  ✓ CONSTRUCCIÓN COMPLETADA")
    print(f"  Archivo: {os.path.basename(DST_FILE)}")
    print("=" * 60)
    print("\nPRÓXIMOS PASOS:")
    print("1. Abre AUTOMATIZACION_v4_FINAL.xlsm en Excel")
    print("2. Habilita macros cuando Excel lo solicite")
    print("3. En DIRECTORIO ejecuta: InicializarEncabezadosDirectorio()")
    print("4. En REPORTES CXC asigna botón → ActualizarReportesCXC")

if __name__ == '__main__':
    main()

