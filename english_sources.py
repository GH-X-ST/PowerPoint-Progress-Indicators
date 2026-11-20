"""Prepare the English VBA, Ribbon, and preset catalog from shared sources."""

import json
import re
from pathlib import Path

from lxml import etree


HEX_STRING = r'"[0-9A-Fa-f]*"'
ENCODED_TEXT = re.compile(
    rf'\bPI_Zh\(\s*({HEX_STRING}(?:\s*&\s*(?:_\s*)?{HEX_STRING})*)\s*\)'
)
HAN = re.compile(r"[\u3400-\u9fff]")


def vba_string(text: str) -> str:
    """Represent English text as a VBA expression, preserving form line breaks."""
    lines = re.split(r"\r\n|\n|\r", text)
    quoted = ['"' + line.replace('"', '""') + '"' for line in lines]
    if len(quoted) == 1:
        return quoted[0]
    return "(" + " & vbCrLf & _\n            ".join(quoted) + ")"


def prepare_english(
    source_dir: Path, project_dir: Path, output: Path
) -> Path:
    locale = project_dir / "locales" / "en"
    target = output / "source"
    target.mkdir(parents=True, exist_ok=True)
    ribbon_text = json.loads((locale / "ribbon.json").read_text("utf-8"))
    vba_text = {}
    for name in ("vba-interface.json", "vba-navigation.json"):
        vba_text.update(json.loads((locale / name).read_text("utf-8")))

    def translate_vba(match: re.Match) -> str:
        units = "".join(re.findall(r'"([0-9A-Fa-f]*)"', match[1]))
        original = bytes.fromhex(units).decode("utf-16-be")
        return vba_string(vba_text[original]) if original else '""'

    for module in sorted(source_dir.iterdir()):
        if module.suffix not in {".bas", ".cls", ".frm"}:
            continue
        source = ENCODED_TEXT.sub(translate_vba, module.read_text("utf-8-sig"))
        if module.name == "mdlOptionsHandling.bas":
            source = source.replace('" CN.11.2 32b/64b"', '" EN.11.2 32b/64b"')
        if module.name == "basGeneral.bas":
            source = source.replace('family = "Microsoft YaHei"',
                                    'family = "Segoe UI"')
            source = source.replace('family = "PingFang SC"',
                                    'family = "Helvetica Neue"')
        (target / module.name).write_text(source, encoding="utf-8")

    xml = etree.parse(source_dir / "customUI14.xml")
    for element in xml.iter():
        for attribute, value in element.attrib.items():
            if HAN.search(value):
                element.set(attribute, ribbon_text[value])
    xml.write(target / "customUI14.xml", encoding="UTF-8", xml_declaration=True)

    catalog = json.loads((project_dir / "styles.json").read_text("utf-8"))
    catalog["version"] = "EN.11.2"
    for style in catalog["styles"]:
        for field in ("name", "family", "description"):
            style[field] = ribbon_text[style[field]]
        if "reference" in style:
            reference = style["reference"]
            for field in ("interpretation", "imageDetail"):
                reference[field] = ribbon_text[reference[field]]
    (output / "styles.json").write_text(
        json.dumps(catalog, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return target
