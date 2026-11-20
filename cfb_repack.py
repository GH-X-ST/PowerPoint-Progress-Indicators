"""Rebuild a small version-3 CFB file while preserving opaque form streams.

The caller supplies VBA stream edits. This module does not compile VBA, change
macro security, or rewrite UserForm content. It rewrites allocation tables and
directory trees because olefile only supports same-size stream replacement.

References:
  [MS-CFB] sections 2.2, 2.3, 2.4, 2.6, and 2.6.4.
  https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
"""

from __future__ import annotations

import io
import struct
import uuid
from dataclasses import dataclass, field
from typing import Iterable, Mapping

import olefile


FREE = 0xFFFFFFFF
END = 0xFFFFFFFE
FAT = 0xFFFFFFFD
SECTOR = 512
MINISECTOR = 64
CUTOFF = 4096


@dataclass
class Node:
    name: str
    kind: int
    path: tuple[str, ...]
    clsid: str = ""
    flags: int = 0
    created: int = 0
    modified: int = 0
    data: bytes = b""
    children: list["Node"] = field(default_factory=list)
    sid: int = 0
    left: int = FREE
    right: int = FREE
    child: int = FREE
    start: int = END
    size: int = 0


def _path(value: str | tuple[str, ...]) -> tuple[str, ...]:
    return tuple(value.split("/")) if isinstance(value, str) else tuple(value)


def _key(name: str) -> tuple[int, tuple[int, ...]]:
    """CFB length-first comparison with one-code-unit uppercase mapping."""
    raw = name.encode("utf-16-le")
    points = struct.unpack("<" + "H" * (len(raw) // 2), raw)
    upper = []
    for point in points:
        char = chr(point)
        mapped = char.upper()
        # Simple uppercase does not expand a character into multiple units.
        upper.append(ord(mapped) if len(mapped) == 1 and ord(mapped) <= 0xFFFF else point)
    return len(raw), tuple(upper)


def _snapshot(data: bytes) -> tuple[Node, dict[tuple[str, ...], Node]]:
    nodes: dict[tuple[str, ...], Node] = {}
    with olefile.OleFileIO(io.BytesIO(data)) as ole:
        def visit(entry, path):
            node = Node(
                name=entry.name,
                kind=entry.entry_type,
                path=path,
                clsid=entry.clsid,
                flags=entry.dwUserFlags,
                created=entry.createTime,
                modified=entry.modifyTime,
                data=ole.openstream(path).read() if entry.entry_type == 2 else b"",
            )
            nodes[path] = node
            for kid in entry.kids:
                node.children.append(visit(kid, path + (kid.name,)))
            return node

        root = visit(ole.root, ())
    return root, nodes


def rebuild_cfb(
    original: bytes,
    replacements: Mapping[str | tuple[str, ...], bytes] | None = None,
    removals: Iterable[str | tuple[str, ...]] = (),
) -> bytes:
    """Return a rebuilt CFB, applying existing-stream edits or adding streams.

    A replacement for a new stream is allowed only under an existing storage.
    Removal of a storage removes its entire subtree. The source is never edited.
    Output is deliberately limited to 109 FAT sectors (about 7 MiB), sufficient
    for this add-in, so no unsupported DIFAT path can silently create bad files.
    """
    root, by_path = _snapshot(original)
    deleted = {_path(p) for p in removals}
    if () in deleted:
        raise ValueError("The root storage cannot be removed")

    def prune(node):
        node.children = [child for child in node.children if child.path not in deleted]
        for child in node.children:
            prune(child)

    prune(root)
    present = {}
    def collect(node):
        present[node.path] = node
        for child in node.children:
            collect(child)
    collect(root)

    for path, new_data in (replacements or {}).items():
        path = _path(path)
        if not isinstance(new_data, bytes):
            raise TypeError(f"Replacement must be bytes: {path!r}")
        if path in present:
            if present[path].kind != 2:
                raise ValueError(f"Replacement target is not a stream: {path!r}")
            present[path].data = new_data
        else:
            parent = present.get(path[:-1])
            if parent is None or parent.kind not in (1, 5):
                raise ValueError(f"No existing parent storage: {path!r}")
            if any(_key(n.name) == _key(path[-1]) for n in parent.children):
                raise ValueError(f"Duplicate case-insensitive stream name: {path!r}")
            node = Node(path[-1], 2, path, data=new_data)
            parent.children.append(node)
            present[path] = node

    flat = []
    def enumerate_nodes(node):
        node.sid = len(flat)
        flat.append(node)
        for child in node.children:
            enumerate_nodes(child)
    enumerate_nodes(root)

    # MS-CFB 2.6.4 explicitly permits an all-black binary search tree. Use a
    # balanced shape to keep lookups short while avoiding recoloring errors.
    def make_tree(items):
        if not items:
            return FREE
        middle = len(items) // 2
        node = items[middle]
        node.left = make_tree(items[:middle])
        node.right = make_tree(items[middle + 1:])
        return node.sid

    for node in flat:
        node.child = make_tree(sorted(node.children, key=lambda child: _key(child.name)))

    sectors: list[bytes] = []
    fat: list[int] = []

    def allocate(data: bytes) -> tuple[int, int]:
        if not data:
            return END, 0
        count = (len(data) + SECTOR - 1) // SECTOR
        start = len(sectors)
        padded = data.ljust(count * SECTOR, b"\0")
        sectors.extend(padded[i * SECTOR:(i + 1) * SECTOR] for i in range(count))
        fat.extend(start + i + 1 if i + 1 < count else END for i in range(count))
        return start, count

    mini = bytearray()
    minifat: list[int] = []
    for node in flat:
        if node.kind != 2:
            continue
        node.size = len(node.data)
        if node.size == 0:
            node.start = END
        elif node.size < CUTOFF:
            node.start = len(minifat)
            count = (node.size + MINISECTOR - 1) // MINISECTOR
            mini.extend(node.data.ljust(count * MINISECTOR, b"\0"))
            minifat.extend(node.start + i + 1 if i + 1 < count else END for i in range(count))
        else:
            node.start, _ = allocate(node.data)

    root.start, _ = allocate(bytes(mini))
    root.size = len(mini)

    if minifat:
        minifat.extend([FREE] * ((-len(minifat)) % (SECTOR // 4)))
        minifat_start, minifat_sectors = allocate(struct.pack("<" + "I" * len(minifat), *minifat))
    else:
        minifat_start, minifat_sectors = END, 0

    directory_count = (len(flat) + 3) // 4
    directory_start, _ = allocate(b"\0" * (directory_count * SECTOR))
    entries = bytearray()
    for node in flat:
        name = (node.name + "\0").encode("utf-16-le")
        if len(name) > 64:
            raise ValueError(f"CFB stream name too long: {node.name!r}")
        if any(c in node.name for c in "/\\:!"):
            raise ValueError(f"Illegal CFB stream name: {node.name!r}")
        guid = uuid.UUID(node.clsid).bytes_le if node.clsid else bytes(16)
        entry = bytearray(name.ljust(64, b"\0"))
        entry.extend(struct.pack("<HBBIII", len(name), node.kind, 1, node.left, node.right, node.child))
        entry.extend(guid)
        entry.extend(struct.pack("<IQQIQ", node.flags, node.created, node.modified,
                                 node.start if node.kind in (2, 5) else 0,
                                 node.size if node.kind in (2, 5) else 0))
        if len(entry) != 128:
            raise AssertionError("Directory entry must be 128 bytes")
        entries.extend(entry)
    entries.extend(bytes(directory_count * SECTOR - len(entries)))
    for i in range(directory_count):
        sectors[directory_start + i] = bytes(entries[i * SECTOR:(i + 1) * SECTOR])

    count_without_fat = len(sectors)
    fat_count = 0
    while (count_without_fat + fat_count + 127) // 128 != fat_count:
        fat_count = (count_without_fat + fat_count + 127) // 128
    if fat_count > 109:
        raise ValueError("CFB exceeds this writer's 109-FAT-sector size limit")
    fat_indices = list(range(count_without_fat, count_without_fat + fat_count))
    fat.extend([FAT] * fat_count)
    fat.extend([FREE] * (fat_count * 128 - len(fat)))
    fat_bytes = struct.pack("<" + "I" * len(fat), *fat)
    sectors.extend(fat_bytes[i * SECTOR:(i + 1) * SECTOR] for i in range(fat_count))

    header = bytearray(bytes.fromhex("D0CF11E0A1B11AE1"))
    header.extend(bytes(16))
    header.extend(struct.pack("<HHHHH", 0x003E, 3, 0xFFFE, 9, 6))
    header.extend(bytes(6))
    header.extend(struct.pack("<IIIIIIIII", 0, fat_count, directory_start, 0,
                              CUTOFF, minifat_start, minifat_sectors, END, 0))
    header.extend(struct.pack("<" + "I" * 109, *(fat_indices + [FREE] * (109 - len(fat_indices)))))
    if len(header) != SECTOR:
        raise AssertionError("CFB header must be 512 bytes")
    result = bytes(header) + b"".join(sectors)

    # Verify every stream and storage metadata against what was requested.
    check_root, check_nodes = _snapshot(result)
    if set(check_nodes) != set(present):
        raise AssertionError("CFB round-trip changed the storage tree")
    for path, expected in present.items():
        actual = check_nodes[path]
        for attr in ("kind", "clsid", "flags", "created", "modified", "data"):
            if getattr(actual, attr) != getattr(expected, attr):
                raise AssertionError(f"CFB round-trip changed {attr}: {path!r}")
    return result


def validate_cfb(data: bytes) -> dict[str, int]:
    """Check parser defects and allocation-chain consistency without execution."""
    with olefile.OleFileIO(io.BytesIO(data), raise_defects=olefile.DEFECT_INCORRECT) as ole:
        streams = ole.listdir(streams=True, storages=False)
        used_large: set[int] = set()
        used_mini: set[int] = set()

        def check_chain(start, size, table, unit, used, label):
            count = (size + unit - 1) // unit
            current = start
            for _ in range(count):
                if current >= len(table) or current in used:
                    raise AssertionError(f"Invalid or overlapping allocation: {label}")
                used.add(current)
                current = table[current]
            if current != END:
                raise AssertionError(f"Incorrect allocation-chain length: {label}")

        # Loading a small stream ensures ole.minifat is initialized.
        for path in streams:
            if 0 < ole.get_size(path) < CUTOFF:
                ole.openstream(path).read()
                break

        for path in streams:
            entry = ole.direntries[ole._find(path)]
            payload = ole.openstream(path).read()
            if len(payload) != entry.size:
                raise AssertionError(f"Truncated stream: {path!r}")
            if entry.size:
                if entry.size < CUTOFF:
                    check_chain(entry.isectStart, entry.size, ole.minifat, MINISECTOR, used_mini, path)
                else:
                    check_chain(entry.isectStart, entry.size, ole.fat, SECTOR, used_large, path)

        if ole.root.size:
            check_chain(ole.root.isectStart, ole.root.size, ole.fat, SECTOR, used_large, "root mini stream")
        if ole.num_mini_fat_sectors:
            check_chain(ole.first_mini_fat_sector, ole.num_mini_fat_sectors * SECTOR,
                        ole.fat, SECTOR, used_large, "mini FAT")
        root, nodes = _snapshot(data)
        dir_size = ((len(nodes) + 3) // 4) * SECTOR
        check_chain(ole.first_dir_sector, dir_size, ole.fat, SECTOR, used_large, "directory")
        if ole.parsing_issues:
            raise AssertionError(f"CFB parsing warnings: {ole.parsing_issues!r}")
        return {"streams": len(streams), "storages": len(nodes) - len(streams),
                "large_sectors": len(used_large), "mini_sectors": len(used_mini)}


if __name__ == "__main__":
    import argparse
    import json
    from pathlib import Path
    import zipfile

    parser = argparse.ArgumentParser(description="Round-trip an Office VBA CFB and verify all opaque streams.")
    parser.add_argument("input")
    args = parser.parse_args()
    path = Path(args.input)
    original = path.read_bytes()
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as package:
            original = package.read("ppt/vbaProject.bin")
    rebuilt = rebuild_cfb(original)
    print(json.dumps({"input_bytes": len(original), "rebuilt_bytes": len(rebuilt),
                      **validate_cfb(rebuilt)}))
