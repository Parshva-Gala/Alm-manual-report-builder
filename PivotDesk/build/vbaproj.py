"""
Read and write an Office VBA project (vbaProject.bin) without Office.

The project is rewritten the way the MS-OVBA spec says a writer should:
source only, no compiled p-code.

  * each module stream holds just its compressed source (MODULEOFFSET = 0)
  * _VBA_PROJECT is the 7-byte stub with version 0xFFFF, so the host
    recompiles from source on first open
  * the __SRP_* caches are dropped; Office rebuilds them on save

Everything else in the project - the references, the protection fields in
PROJECT, the host extender info - is carried over byte for byte.

Two formats are implemented here, both from their published specs:

  MS-OVBA  2.4.1   the RLE/LZ77 compression used for dir and module streams
  MS-CFB           the compound file container (v3, 512-byte sectors)
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

# ============================================================================
#  MS-OVBA 2.4.1 compression
# ============================================================================

CHUNK = 4096


def _copy_token_help(difference: int):
    bit_count = 4
    while (1 << bit_count) < difference:
        bit_count += 1
    length_mask = 0xFFFF >> bit_count
    offset_mask = (~length_mask) & 0xFFFF
    max_length = length_mask + 3
    return length_mask, offset_mask, bit_count, max_length


def decompress(data: bytes) -> bytes:
    if not data or data[0] != 0x01:
        raise ValueError("not a compressed container")
    out = bytearray()
    pos = 1
    n = len(data)
    while pos < n:
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        flag = (header >> 15) & 1
        chunk_end = min(n, pos + size)
        pos += 2
        start = len(out)
        if not flag:
            out += data[pos:pos + CHUNK]
            pos += CHUNK
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                else:
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    lmask, omask, bits, _ = _copy_token_help(len(out) - start)
                    length = (token & lmask) + 3
                    offset = ((token & omask) >> (16 - bits)) + 1
                    src = len(out) - offset
                    for i in range(length):
                        out.append(out[src + i])
    return bytes(out)


def _compress_chunk(buf: bytes, start: int, end: int) -> bytes:
    """One CompressedChunk (header + data) for buf[start:end]."""
    out = bytearray()
    pos = start
    heads: dict[bytes, list[int]] = {}

    def remember(p: int):
        if p + 3 <= end:
            heads.setdefault(buf[p:p + 3], []).append(p)

    while pos < end:
        flag_index = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= end:
                break
            difference = pos - start
            best_len, best_off = 0, 0
            if difference > 0 and pos + 3 <= end:
                _, _, _, max_len = _copy_token_help(difference)
                cands = heads.get(buf[pos:pos + 3])
                if cands:
                    limit = min(max_len, end - pos)
                    tried = 0
                    for cand in reversed(cands):
                        tried += 1
                        if tried > 96:
                            break
                        ln = 0
                        while ln < limit and buf[cand + ln] == buf[pos + ln]:
                            ln += 1
                        if ln > best_len:
                            best_len, best_off = ln, pos - cand
                            if ln == limit:
                                break
            if best_len >= 3:
                _, _, bits, _ = _copy_token_help(difference)
                token = ((best_off - 1) << (16 - bits)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
                for p in range(pos, pos + best_len):
                    remember(p)
                pos += best_len
            else:
                out.append(buf[pos])
                remember(pos)
                pos += 1
        out[flag_index] = flags

    if len(out) > CHUNK:
        raise ValueError("incompressible chunk; raw chunks would pad the source with NULs")
    header = 0xB000 | ((len(out) + 2 - 3) & 0x0FFF)   # flag=1, signature=0b011
    return struct.pack("<H", header) + bytes(out)


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    for start in range(0, len(data), CHUNK):
        out += _compress_chunk(data, start, min(len(data), start + CHUNK))
    return bytes(out)


# ============================================================================
#  MS-CFB compound files
# ============================================================================

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF
SECTOR = 512
MINI = 64
MINI_CUTOFF = 4096


def read_cfb(path_or_bytes) -> dict[str, bytes]:
    """Every stream in a compound file, keyed by its '/'-joined path."""
    import olefile  # reading is the solved half; olefile does it well
    ole = olefile.OleFileIO(path_or_bytes)
    out = {}
    for entry in ole.listdir(streams=True, storages=False):
        p = "/".join(entry)
        out[p] = ole.openstream(p).read()
    ole.close()
    return out


def _cfb_key(name: str):
    # MS-CFB 2.6.4: shorter names sort first, then by upper-cased code units.
    return (len(name), name.upper())


@dataclass
class _Node:
    name: str
    kind: int                      # 1 storage, 2 stream, 5 root
    data: bytes = b""
    children: list = field(default_factory=list)
    sid: int = 0
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    color: int = 1
    start: int = ENDOFCHAIN
    size: int = 0


def write_cfb(streams: dict[str, bytes]) -> bytes:
    root = _Node("Root Entry", 5)
    storages = {"": root}
    for path in sorted(streams):
        parts = path.split("/")
        parent = root
        for i, part in enumerate(parts[:-1]):
            key = "/".join(parts[:i + 1])
            if key not in storages:
                node = _Node(part, 1)
                storages[key] = node
                parent.children.append(node)
            parent = storages[key]
        parent.children.append(_Node(parts[-1], 2, data=streams[path]))

    # Directory ids, root first, breadth first.
    order = [root]
    i = 0
    while i < len(order):
        order.extend(sorted(order[i].children, key=lambda n: _cfb_key(n.name)))
        i += 1
    for sid, node in enumerate(order):
        node.sid = sid

    # Each storage's children as a balanced red-black tree: median split, and
    # the deepest level red (never the root). Every root-to-leaf path then
    # carries the same number of black nodes, which is the invariant.
    def build(nodes, depth, out_depths):
        if not nodes:
            return NOSTREAM
        mid = len(nodes) // 2
        n = nodes[mid]
        out_depths.append((n, depth))
        n.left = build(nodes[:mid], depth + 1, out_depths)
        n.right = build(nodes[mid + 1:], depth + 1, out_depths)
        return n.sid

    for node in order:
        if node.kind in (1, 5) and node.children:
            kids = sorted(node.children, key=lambda n: _cfb_key(n.name))
            depths = []
            node.child = build(kids, 0, depths)
            maxd = max(d for _, d in depths)
            for n, d in depths:
                n.color = 0 if (d == maxd and d > 0) else 1

    # --- the mini stream: every stream under the cutoff -------------------
    mini = bytearray()
    minifat: list[int] = []
    for node in order:
        if node.kind == 2 and len(node.data) < MINI_CUTOFF:
            node.size = len(node.data)
            if node.size == 0:
                node.start = ENDOFCHAIN
                continue
            first = len(mini) // MINI
            count = (node.size + MINI - 1) // MINI
            mini += node.data + b"\x00" * (count * MINI - node.size)
            for k in range(count):
                minifat.append(first + k + 1 if k < count - 1 else ENDOFCHAIN)
            node.start = first

    # --- lay out the regular sectors --------------------------------------
    body = bytearray()
    fat: list[int] = []

    def alloc(data: bytes) -> int:
        if not data:
            return ENDOFCHAIN
        first = len(fat)
        count = (len(data) + SECTOR - 1) // SECTOR
        body.extend(data + b"\x00" * (count * SECTOR - len(data)))
        for k in range(count):
            fat.append(first + k + 1 if k < count - 1 else ENDOFCHAIN)
        return first

    for node in order:
        if node.kind == 2 and len(node.data) >= MINI_CUTOFF:
            node.size = len(node.data)
            node.start = alloc(node.data)

    root.size = len(mini)
    root.start = alloc(bytes(mini)) if mini else ENDOFCHAIN

    minifat_bytes = b"".join(struct.pack("<I", v) for v in minifat)
    if minifat_bytes:
        pad = (-len(minifat_bytes)) % SECTOR
        minifat_bytes += struct.pack("<I", FREESECT) * (pad // 4)
    minifat_start = alloc(minifat_bytes) if minifat_bytes else ENDOFCHAIN
    n_minifat = len(minifat_bytes) // SECTOR

    # directory
    entries = bytearray()
    for node in order:
        nm = node.name.encode("utf-16-le")
        if len(nm) > 62:
            raise ValueError("name too long: " + node.name)
        e = bytearray(128)
        e[0:len(nm)] = nm
        struct.pack_into("<H", e, 64, len(nm) + 2)
        e[66] = node.kind
        e[67] = node.color
        struct.pack_into("<III", e, 68, node.left, node.right, node.child)
        struct.pack_into("<I", e, 116, node.start if node.kind != 1 else 0)
        struct.pack_into("<Q", e, 120, node.size if node.kind != 1 else 0)
        entries += e
    while len(entries) % SECTOR:
        e = bytearray(128)
        struct.pack_into("<III", e, 68, NOSTREAM, NOSTREAM, NOSTREAM)
        entries += e
    dir_start = alloc(bytes(entries))

    # FAT sectors describe themselves too, so size it to a fixed point.
    n_fat = 1
    while True:
        total = len(fat) + n_fat
        need = (total + 127) // 128
        if need <= n_fat:
            break
        n_fat = need
    if n_fat > 109:
        raise ValueError("file too large for a header-only DIFAT")
    fat_first = len(fat)
    fat.extend([FATSECT] * n_fat)
    fat.extend([FREESECT] * (n_fat * 128 - len(fat)))
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    body.extend(fat_bytes)

    header = bytearray(SECTOR)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", header, 24, 0x003E, 0x0003, 0xFFFE, 9, 6)
    struct.pack_into("<I", header, 40, 0)              # dir sectors (v3: 0)
    struct.pack_into("<I", header, 44, n_fat)
    struct.pack_into("<I", header, 48, dir_start)
    struct.pack_into("<I", header, 52, 0)
    struct.pack_into("<I", header, 56, MINI_CUTOFF)
    struct.pack_into("<I", header, 60, minifat_start)
    struct.pack_into("<I", header, 64, n_minifat)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)     # no DIFAT sectors
    struct.pack_into("<I", header, 72, 0)
    for k in range(109):
        struct.pack_into("<I", header, 76 + 4 * k, fat_first + k if k < n_fat else FREESECT)
    return bytes(header) + bytes(body)


# ============================================================================
#  The VBA project
# ============================================================================

@dataclass
class Module:
    name: str
    stream: str
    kind: str                      # "std" | "class" | "doc"
    source: str                    # full text, attribute lines included, CRLF
    docstring: bytes = b""
    docstring_u: bytes = b""
    help_context: int = 0
    readonly: bool = False
    private: bool = False
    name_u: bytes = b""
    stream_u: bytes = b""


class VBAProject:
    def __init__(self, streams: dict[str, bytes]):
        self.streams = dict(streams)
        self.codepage = 1252
        raw = decompress(self.streams["VBA/dir"])
        self._parse_dir(raw)
        self.project_text = self.streams["PROJECT"].decode(self._enc())

    def _enc(self):
        return "cp%d" % self.codepage

    # --- dir stream -------------------------------------------------------
    def _parse_dir(self, d: bytes):
        pos = 0
        head_end = None
        while pos < len(d):
            rid, size = struct.unpack_from("<HI", d, pos)
            if rid == 0x000F:                           # PROJECTMODULES
                head_end = pos
                break
            if rid == 0x0003:
                self.codepage = struct.unpack_from("<H", d, pos + 6)[0]
            if rid == 0x0009:                           # PROJECTVERSION: size lies
                pos += 6 + 4 + 2
                continue
            pos += 6 + size
        if head_end is None:
            raise ValueError("no PROJECTMODULES record")
        self.dir_head = d[:head_end]

        pos = head_end
        _, size = struct.unpack_from("<HI", d, pos)
        count = struct.unpack_from("<H", d, pos + 6)[0]
        pos += 6 + size
        rid, size = struct.unpack_from("<HI", d, pos)   # PROJECTCOOKIE
        assert rid == 0x0013
        self.project_cookie = d[pos + 6:pos + 6 + size]
        pos += 6 + size

        self.modules: list[Module] = []
        for _ in range(count):
            m = {"readonly": False, "private": False}
            while True:
                rid, size = struct.unpack_from("<HI", d, pos)
                body = d[pos + 6:pos + 6 + size]
                pos += 6 + size
                if rid == 0x0019: m["name"] = body.decode(self._enc())
                elif rid == 0x0047: m["name_u"] = body
                elif rid == 0x001A: m["stream"] = body.decode(self._enc())
                elif rid == 0x0032: m["stream_u"] = body
                elif rid == 0x001C: m["doc"] = body
                elif rid == 0x0048: m["doc_u"] = body
                elif rid == 0x0031: m["offset"] = struct.unpack("<I", body)[0]
                elif rid == 0x001E: m["help"] = struct.unpack("<I", body)[0]
                elif rid == 0x002C: pass                              # cookie
                elif rid == 0x0021: m["type"] = "std"
                elif rid == 0x0022: m["type"] = "cls"
                elif rid == 0x0025: m["readonly"] = True
                elif rid == 0x0028: m["private"] = True
                elif rid == 0x002B: break
                else:
                    raise ValueError("unexpected module record 0x%04X" % rid)
            data = self.streams["VBA/" + m["stream"]]
            src = decompress(data[m["offset"]:]).decode(self._enc())
            kind = m["type"]
            if kind == "cls":
                # Workbook and Worksheet bases mark a document module; anything
                # else with a VB_Base is an ordinary class.
                is_doc = "{00020819-" in src or "{00020820-" in src
                kind = "doc" if is_doc else "class"
            self.modules.append(Module(
                name=m["name"], stream=m["stream"], kind=kind, source=src,
                docstring=m.get("doc", b""), docstring_u=m.get("doc_u", b""),
                help_context=m.get("help", 0), readonly=m["readonly"],
                private=m["private"], name_u=m.get("name_u", b""),
                stream_u=m.get("stream_u", b"")))

    def _emit_dir(self) -> bytes:
        enc = self._enc()
        out = bytearray(self.dir_head)
        out += struct.pack("<HIH", 0x000F, 2, len(self.modules))
        out += struct.pack("<HI", 0x0013, len(self.project_cookie)) + self.project_cookie

        def rec(rid, body: bytes):
            out.extend(struct.pack("<HI", rid, len(body)) + body)

        for m in self.modules:
            rec(0x0019, m.name.encode(enc))
            rec(0x0047, m.name.encode("utf-16-le"))
            rec(0x001A, m.stream.encode(enc))
            rec(0x0032, m.stream.encode("utf-16-le"))
            rec(0x001C, m.docstring)
            rec(0x0048, m.docstring_u)
            rec(0x0031, struct.pack("<I", 0))
            rec(0x001E, struct.pack("<I", m.help_context))
            rec(0x002C, struct.pack("<H", 0xFFFF))
            rec(0x0021 if m.kind == "std" else 0x0022, b"")
            if m.readonly:
                rec(0x0025, b"")
            if m.private:
                rec(0x0028, b"")
            rec(0x002B, b"")
        out += struct.pack("<HI", 0x0010, 0)
        return bytes(out)

    # --- editing ----------------------------------------------------------
    def module(self, name: str) -> Module:
        for m in self.modules:
            if m.name.lower() == name.lower():
                return m
        raise KeyError(name)

    def has(self, name: str) -> bool:
        return any(m.name.lower() == name.lower() for m in self.modules)

    def code(self, name: str) -> str:
        """The module's code without its Attribute header, LF line ends."""
        return split_attributes(self.module(name).source)[1]

    def set_code(self, name: str, code: str):
        m = self.module(name)
        attrs, _ = split_attributes(m.source)
        m.source = join_source(attrs, code)

    def add_std(self, name: str, code: str):
        if self.has(name):
            self.set_code(name, code)
            return
        attrs = ['Attribute VB_Name = "%s"' % name]
        m = Module(name=name, stream=name, kind="std", source=join_source(attrs, code))
        # Standard modules sit before the first class/document module that
        # follows them in the original order, which keeps PROJECT tidy.
        self.modules.append(m)
        self._project_add("Module=" + name, name)

    def remove(self, name: str):
        m = self.module(name)
        if m.kind == "doc":
            raise ValueError("document modules belong to a sheet; blank them instead")
        self.modules.remove(m)
        self.streams.pop("VBA/" + m.stream, None)
        lines = self.project_text.split("\r\n")
        lines = [ln for ln in lines
                 if ln not in ("Module=" + name, "Class=" + name)
                 and not ln.startswith(name + "=")]
        self.project_text = "\r\n".join(lines)

    def _project_add(self, decl: str, name: str):
        lines = self.project_text.split("\r\n")
        # after the last Module= line
        idx = max(i for i, ln in enumerate(lines) if ln.startswith(("Module=", "Class=", "Document=")))
        last_mod = max((i for i, ln in enumerate(lines) if ln.startswith("Module=")), default=idx)
        lines.insert(last_mod + 1, decl)
        if "[Workspace]" in lines:
            lines.insert(lines.index("[Workspace]") + 1, "%s=0, 0, 0, 0, C" % name)
        self.project_text = "\r\n".join(lines)

    # --- writing ----------------------------------------------------------
    def to_streams(self) -> dict[str, bytes]:
        enc = self._enc()
        out = {k: v for k, v in self.streams.items()
               if not k.startswith("VBA/__SRP_")}
        for m in self.modules:
            out["VBA/" + m.stream] = compress(m.source.encode(enc))
        out["VBA/dir"] = compress(self._emit_dir())
        out["VBA/_VBA_PROJECT"] = bytes.fromhex("CC61FFFF000000")
        out["PROJECT"] = self.project_text.encode(enc)
        out["PROJECTwm"] = self._projectwm()
        return out

    def _projectwm(self) -> bytes:
        enc = self._enc()
        b = bytearray()
        for m in self.modules:
            b += m.name.encode(enc) + b"\x00" + m.name.encode("utf-16-le") + b"\x00\x00"
        b += b"\x00\x00"
        return bytes(b)

    def to_bin(self) -> bytes:
        return write_cfb(self.to_streams())


def split_attributes(source: str):
    lines = source.replace("\r\n", "\n").split("\n")
    i = 0
    while i < len(lines) and lines[i].startswith("Attribute VB_"):
        i += 1
    return lines[:i], "\n".join(lines[i:])


def join_source(attrs, code: str) -> str:
    code = code.replace("\r\n", "\n").strip("\n")
    return "\r\n".join(list(attrs) + code.split("\n")) + "\r\n"


def load(path: str) -> VBAProject:
    return VBAProject(read_cfb(path))
