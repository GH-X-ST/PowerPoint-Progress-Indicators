"""Build source-based PowerPoint add-ins without retaining stale VBA caches.

The official .pptm and .ppam are independent package bases. Their original
slides, images, UserForms, project protection fields, references, and form
identity attributes are preserved. Only VBA source and Ribbon XML change.

PowerPoint must compile and execute the resulting project for native-runtime
verification. The checks here cover the package, CFB, VBA compression, source,
form identities, Ribbon callbacks, and unchanged resources, not Office UI.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import io
import json
import posixpath
import re
import struct
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent
RESOURCE_ROOT = ROOT
sys.path.insert(0, str(RESOURCE_ROOT))
sys.path.insert(0, str(RESOURCE_ROOT / "deps"))

import olefile
from lxml import etree
from oletools.olevba import VBA_Project, decompress_stream

from cfb_repack import rebuild_cfb, validate_cfb
from vba_source_checks import validate_vba_sources


def compress_vba(data: bytes) -> bytes:
    """Encode MS-OVBA 4096-byte chunks using literal and greedy copy tokens.

    Unlike short all-literal chunk encoders, every non-final decompressed chunk
    has exactly 4096 bytes. For an incompressible short final chunk that would
    need raw NUL padding, fail rather than silently alter the caller's bytes.
    VBA source and the add-in's dir stream fit the compressed representation.
    """
    result = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        length = len(chunk)
        current = 0
        candidates: dict[bytes, list[int]] = defaultdict(list)
        encoded = bytearray()
        while current < length:
            flags_at = len(encoded)
            encoded.append(0)
            flags = 0
            for token_index in range(8):
                if current >= length:
                    break
                offset_bits = max(4, (current - 1).bit_length())
                length_bits = 16 - offset_bits
                max_length = min(length - current, ((1 << length_bits) - 1) + 3)
                best_length = 0
                best_offset = 0
                if max_length >= 3:
                    for previous in reversed(candidates.get(chunk[current:current + 3], ())):
                        offset = current - previous
                        if offset > (1 << offset_bits):
                            break
                        count = 3
                        while count < max_length and chunk[previous + count] == chunk[current + count]:
                            count += 1
                        if count > best_length:
                            best_length = count
                            best_offset = offset
                            if count == max_length:
                                break
                old_current = current
                if best_length >= 3:
                    token = ((best_offset - 1) << length_bits) | (best_length - 3)
                    encoded.extend(struct.pack("<H", token))
                    flags |= 1 << token_index
                    current += best_length
                else:
                    encoded.append(chunk[current])
                    current += 1
                for pos in range(old_current, current):
                    if pos + 3 <= length:
                        candidates[chunk[pos:pos + 3]].append(pos)
            encoded[flags_at] = flags
        if len(encoded) <= 4096:
            header = 0xB000 | (len(encoded) - 1)
            result.extend(struct.pack("<H", header))
            result.extend(encoded)
        elif length == 4096:
            result.extend(b"\xff\x3f")
            result.extend(chunk)
        else:
            raise ValueError("An incompressible final VBA chunk would require altering source with NUL padding")
    packed = bytes(result)
    if decompress_stream(packed) != data:
        raise AssertionError("MS-OVBA compression round-trip failed")
    return packed


def parse_project(cfb: bytes):
    ole = olefile.OleFileIO(io.BytesIO(cfb), raise_defects=olefile.DEFECT_INCORRECT)
    project = VBA_Project(ole, "", "PROJECT", "VBA/dir", relaxed=False)
    module_metadata_start = project.dir_stream.tell()
    project.parse_project_stream()
    list(project.parse_modules())
    return ole, project, module_metadata_start


def normalized_code(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def apply_form_identity(text: str, original: str) -> str:
    """Keep VB_Base tied to the corresponding base package's designer storage."""
    pattern = r"(?m)^Attribute VB_Base = [^\r\n]+"
    original_line = re.search(pattern, original)
    changed_line = re.search(pattern, text)
    if bool(original_line) != bool(changed_line):
        raise ValueError("A form's VB_Base attribute was added or removed")
    if original_line:
        text, count = re.subn(pattern, lambda _: original_line.group(0), text)
        if count != 1:
            raise ValueError("A form must have exactly one VB_Base attribute")
    return text


def build_vba_project(original: bytes, source_dir: Path) -> tuple[bytes, dict]:
    source_checks = validate_vba_sources(source_dir)
    ole, project, metadata_start = parse_project(original)
    expected_files = {module.filename_str for module in project.modules}
    actual_files = {p.name for p in source_dir.iterdir() if p.suffix.lower() in (".bas", ".frm", ".cls")}
    if expected_files != actual_files:
        raise ValueError(f"VBA module inventory differs: missing={expected_files - actual_files}, extra={actual_files - expected_files}")

    directory = bytearray(decompress_stream(ole.openstream("VBA/dir").read()))
    offset_record = b"\x31\x00\x04\x00\x00\x00"
    positions = [m.start() for m in re.finditer(re.escape(offset_record), directory)]
    positions = [p for p in positions if p >= metadata_start]
    if len(positions) != len(project.modules):
        raise ValueError("MODULEOFFSET records cannot be mapped one-to-one to modules")

    replacements: dict[str, bytes] = {}
    expected_sources: dict[str, bytes] = {}
    identities = 0
    for module, position in zip(project.modules, positions):
        old_offset = struct.unpack_from("<I", directory, position + 6)[0]
        if old_offset != module.textoffset:
            raise ValueError(f"MODULEOFFSET metadata mismatch: {module.name_str}")
        struct.pack_into("<I", directory, position + 6, 0)
        text = (source_dir / module.filename_str).read_text(encoding="utf-8-sig")
        text = apply_form_identity(text, module.code_str)
        identities += bool(re.search(r"(?m)^Attribute VB_Base =", text))
        name = re.search(r'^Attribute VB_Name = "([^\"]+)"', text, re.MULTILINE)
        if name is None or name.group(1) != module.name_str:
            raise ValueError(f"VBA module identity changed: {module.filename_str}")
        text = normalized_code(text)
        code = text.encode(project.codec, errors="strict")
        expected_sources[module.code_path] = code
        replacements[module.code_path] = compress_vba(code)

    replacements["VBA/dir"] = compress_vba(bytes(directory))
    original_header = ole.openstream("VBA/_VBA_PROJECT").read()
    if original_header[:2] != b"\xcc\x61":
        raise ValueError("Unexpected _VBA_PROJECT signature")
    replacements["VBA/_VBA_PROJECT"] = b"\xcc\x61\xff\xff\x00" + original_header[5:7]
    removals = [tuple(path) for path in ole.listdir() if len(path) == 2 and path[0] == "VBA" and re.fullmatch(r"__SRP_[0-9a-f]+", path[1], re.I)]
    rebuilt = rebuild_cfb(original, replacements, removals)
    allocation = validate_cfb(rebuilt)

    check_ole, check_project, _ = parse_project(rebuilt)
    actual_sources = {module.code_path: module.code_raw for module in check_project.modules}
    if actual_sources != expected_sources:
        raise AssertionError("Repacked VBA source differs from supplied source")
    if any(module.textoffset != 0 for module in check_project.modules):
        raise AssertionError("A module retains a performance-cache prefix")
    if len(check_ole.openstream("VBA/_VBA_PROJECT").read()) != 7:
        raise AssertionError("Project retains a performance cache")
    if any(path[-1].startswith("__SRP_") for path in check_ole.listdir()):
        raise AssertionError("Project retains SRP performance-cache streams")
    for path in ole.listdir():
        key = "/".join(path)
        if key not in replacements and tuple(path) not in removals:
            if check_ole.openstream(path).read() != ole.openstream(path).read():
                raise AssertionError(f"An unrelated stream changed: {path!r}")
    if check_ole.openstream("PROJECT").read() != ole.openstream("PROJECT").read():
        raise AssertionError("PROJECT or its protection fields changed")
    report = {
        "source_checks": source_checks,
        "source_codepage": project.codepage,
        "module_count": len(project.modules),
        "form_identity_count": identities,
        "removed_srp_streams": len(removals),
        "module_offsets": [module.textoffset for module in check_project.modules],
        "vba_bytes_before": len(original),
        "vba_bytes_after": len(rebuilt),
        "cfb": allocation,
        "sources": [{"module": module.name_str, "filename": module.filename_str,
                     "bytes": len(module.code_raw), "sha256": hashlib.sha256(module.code_raw).hexdigest()}
                    for module in check_project.modules],
    }
    check_ole.close()
    ole.close()
    return rebuilt, report


def verify_ribbon(xml: bytes, source_dir: Path) -> dict:
    tree = etree.fromstring(xml)
    procedures = set()
    for path in source_dir.iterdir():
        if path.suffix.lower() in (".bas", ".frm", ".cls"):
            procedures.update(name.lower() for name in re.findall(
                r"(?im)^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)", path.read_text(encoding="utf-8-sig")))
    callbacks = set()
    ids = []
    for node in tree.iter():
        if "id" in node.attrib:
            ids.append(node.get("id"))
        for attribute, value in node.attrib.items():
            if attribute.startswith("get") or attribute.startswith("on"):
                callbacks.add(value)
    absent = sorted(name for name in callbacks if name.lower() not in procedures)
    if absent:
        raise ValueError(f"Ribbon refers to missing callbacks: {absent!r}")
    if len(ids) != len(set(ids)):
        raise ValueError("Ribbon XML contains duplicate control IDs")
    return {"control_ids": len(ids), "callbacks": len(callbacks)}


def verify_package(path: Path, original_path: Path, expected_changes: set[str]) -> dict:
    xml_count = 0
    relation_count = 0
    unchanged = 0
    with zipfile.ZipFile(path) as current, zipfile.ZipFile(original_path) as original:
        if current.testzip() is not None:
            raise AssertionError("ZIP CRC verification failed")
        names = set(current.namelist())
        if names != set(original.namelist()):
            raise AssertionError("Package parts were added or lost")
        for name in names:
            data = current.read(name)
            if name not in expected_changes:
                if data != original.read(name):
                    raise AssertionError(f"Unrelated package part changed: {name}")
                unchanged += 1
            if data and (name.endswith(".xml") or name.endswith(".rels")):
                root = etree.fromstring(data)
                xml_count += 1
                if name.endswith(".rels"):
                    source_dir = posixpath.dirname(name).removesuffix("/_rels")
                    if name == "_rels/.rels":
                        source_dir = ""
                    for relationship in root:
                        relation_count += 1
                        if relationship.get("TargetMode") == "External":
                            continue
                        target = relationship.get("Target", "")
                        resolved = posixpath.normpath(posixpath.join(source_dir, target)).lstrip("/")
                        if resolved not in names:
                            raise AssertionError(f"Missing package relationship target: {name}: {target}")
        main = "ppt/presentation.xml"
        if not current.read(main) and path.suffix != ".ppam":
            raise AssertionError("Only the original add-in may have an empty presentation part")
    return {"parts": len(names), "unchanged_parts": unchanged,
            "parsed_xml_parts": xml_count, "relationships": relation_count}


def build_package(base: Path, source_dir: Path, output: Path) -> dict:
    ribbon = (source_dir / "customUI14.xml").read_bytes()
    ribbon_report = verify_ribbon(ribbon, source_dir)
    with zipfile.ZipFile(base) as original:
        if any("signature" in name.lower() for name in original.namelist()):
            raise ValueError("A signed package needs explicit signature handling")
        vba, vba_report = build_vba_project(original.read("ppt/vbaProject.bin"), source_dir)
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w") as target:
            target.comment = original.comment
            for info in original.infolist():
                if info.filename == "ppt/vbaProject.bin":
                    data = vba
                elif info.filename == "customUI/customUI14.xml":
                    data = ribbon
                else:
                    data = original.read(info.filename)
                packed_info = copy.copy(info)
                target.writestr(packed_info, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    package_report = verify_package(output, base, {"ppt/vbaProject.bin", "customUI/customUI14.xml"})
    return {"file": str(output.resolve()), "bytes": output.stat().st_size,
            "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
            "base": str(base.resolve()), "ribbon": ribbon_report,
            "vba": vba_report, "package": package_report,
            "native_powerpoint_execution_tested": False}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-dir", type=Path, default=ROOT / "source")
    parser.add_argument("--output-dir", type=Path, default=ROOT / "output")
    parser.add_argument("--language", choices=("zh-CN", "en"), default="zh-CN")
    parser.add_argument("--version-token",
                        help="Override the edition's default filename suffix.")
    parser.add_argument("--report", type=Path, default=ROOT / "package-validation.json")
    args = parser.parse_args()
    source_dir = args.source_dir
    if args.language == "en":
        from english_sources import prepare_english

        source_dir = prepare_english(source_dir, ROOT, args.output_dir / "en")
    version = args.version_token or (
        "EN11_1" if args.language == "en" else "CN11_1"
    )
    base_dir = RESOURCE_ROOT / "original/ProgressIndicators_v1960_3264b"
    reports = []
    for extension in ("ppam", "pptm"):
        base = base_dir / f"ProgressIndicators_v1960_3264b.{extension}"
        output = args.output_dir / (
            f"ProgressIndicators_v1960_3264b_SB_{version}.{extension}"
        )
        reports.append(build_package(base, source_dir, output))
    args.report.write_text(json.dumps(reports, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps([{key: value for key, value in item.items() if key != "vba"} | {
        "modules": item["vba"]["module_count"], "source_codepage": item["vba"]["source_codepage"]}
        for item in reports], indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
