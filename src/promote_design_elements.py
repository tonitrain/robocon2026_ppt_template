"""Make PowerPoint master/layout decorations editable on individual slides."""

from __future__ import annotations

import argparse
import copy
import posixpath
import sys
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

DRAWABLES = {"sp", "pic", "graphicFrame", "cxnSp", "grpSp"}

for prefix, uri in {
    "a": A,
    "p": P,
    "r": R,
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
}.items():
    ET.register_namespace(prefix, uri)
ET.register_namespace("", PKG_REL)


def qname(namespace: str, local_name: str) -> str:
    return f"{{{namespace}}}{local_name}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def serialize_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def part_from_relative(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relative_target(source_part: str, target_part: str) -> str:
    return posixpath.relpath(target_part, posixpath.dirname(source_part))


def relationship_map(root: ET.Element) -> dict[str, ET.Element]:
    return {
        relationship.attrib["Id"]: relationship
        for relationship in root.findall(qname(PKG_REL, "Relationship"))
    }


def next_relationship_id(root: ET.Element) -> str:
    used = set(relationship_map(root))
    number = 1
    while f"rId{number}" in used:
        number += 1
    return f"rId{number}"


def has_placeholder(element: ET.Element) -> bool:
    return any(child.tag == qname(P, "ph") for child in element.iter())


def max_shape_id(root: ET.Element) -> int:
    values = []
    for element in root.iter(qname(P, "cNvPr")):
        try:
            values.append(int(element.attrib.get("id", "0")))
        except ValueError:
            continue
    return max(values, default=0)


def remap_relationships(
    element: ET.Element,
    source_part: str,
    target_part: str,
    source_rels: ET.Element,
    target_rels: ET.Element,
) -> None:
    source_map = relationship_map(source_rels)
    remapped: dict[str, str] = {}

    for node in element.iter():
        for attribute_name, value in list(node.attrib.items()):
            if not attribute_name.startswith(f"{{{R}}}") or not value.startswith("rId"):
                continue
            if value in remapped:
                node.attrib[attribute_name] = remapped[value]
                continue
            if value not in source_map:
                raise ValueError(f"Relationship {value} is missing from {source_part}")

            source_relationship = source_map[value]
            new_id = next_relationship_id(target_rels)
            new_relationship = ET.SubElement(target_rels, qname(PKG_REL, "Relationship"))
            new_relationship.attrib["Id"] = new_id
            new_relationship.attrib["Type"] = source_relationship.attrib["Type"]
            if "TargetMode" in source_relationship.attrib:
                new_relationship.attrib["TargetMode"] = source_relationship.attrib["TargetMode"]
                new_relationship.attrib["Target"] = source_relationship.attrib["Target"]
            else:
                absolute_target = part_from_relative(
                    source_part, source_relationship.attrib["Target"]
                )
                new_relationship.attrib["Target"] = relative_target(
                    target_part, absolute_target
                )
            remapped[value] = new_id
            node.attrib[attribute_name] = new_id


def promote_from_source(
    source_root: ET.Element,
    source_part: str,
    source_rels: ET.Element,
    slide_roots: list[tuple[ET.Element, ET.Element, str]],
) -> int:
    """Copy non-placeholder drawable children to slides, then remove the originals."""
    source_tree = source_root.find(f".//{qname(P, 'cSld')}/{qname(P, 'spTree')}")
    if source_tree is None:
        return 0

    movable = [
        child
        for child in list(source_tree)
        if local_name(child.tag) in DRAWABLES and not has_placeholder(child)
    ]
    if not movable:
        return 0

    for slide_root, slide_rels, slide_part in slide_roots:
        slide_tree = slide_root.find(f".//{qname(P, 'cSld')}/{qname(P, 'spTree')}")
        if slide_tree is None:
            raise ValueError(f"Slide {slide_part} has no shape tree")

        insert_at = 2
        next_id = max_shape_id(slide_root) + 1
        for original in movable:
            promoted = copy.deepcopy(original)
            remap_relationships(
                promoted, source_part, slide_part, source_rels, slide_rels
            )
            for c_nv_pr in promoted.iter(qname(P, "cNvPr")):
                c_nv_pr.attrib["id"] = str(next_id)
                next_id += 1
            slide_tree.insert(insert_at, promoted)
            insert_at += 1

    for original in movable:
        source_tree.remove(original)
    return len(movable)


def slide_layout_and_master(
    entries: dict[str, bytes], slide_number: int
) -> tuple[str, str]:
    slide_part = f"ppt/slides/slide{slide_number}.xml"
    rel_part = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
    rel_root = parse_xml(entries[rel_part])
    layout_relationship = next(
        relationship
        for relationship in rel_root.findall(qname(PKG_REL, "Relationship"))
        if relationship.attrib["Type"].endswith("/slideLayout")
    )
    layout_part = part_from_relative(slide_part, layout_relationship.attrib["Target"])

    layout_rel_part = (
        f"{posixpath.dirname(layout_part)}/_rels/{posixpath.basename(layout_part)}.rels"
    )
    layout_rel_root = parse_xml(entries[layout_rel_part])
    master_relationship = next(
        relationship
        for relationship in layout_rel_root.findall(qname(PKG_REL, "Relationship"))
        if relationship.attrib["Type"].endswith("/slideMaster")
    )
    master_part = part_from_relative(layout_part, master_relationship.attrib["Target"])
    return layout_part, master_part


def write_archive(path: Path, entries: dict[str, tuple[zipfile.ZipInfo, bytes]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        dir=path.parent, prefix=f".{path.name}.", suffix=".tmp", delete=False
    ) as temporary:
        temporary_path = Path(temporary.name)
    try:
        with zipfile.ZipFile(temporary_path, "w") as archive:
            for info, data in entries.values():
                archive.writestr(info, data)
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument(
        "--slides",
        default="all",
        help="comma-separated slide numbers, or all (default: all)",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    input_path = args.input.resolve()
    output_path = args.output.resolve()
    if input_path == output_path:
        print("Output must be different from input", file=sys.stderr)
        return 2
    if not input_path.is_file() or not zipfile.is_zipfile(input_path):
        print(f"Input is not a valid PPTX/POTX file: {input_path}", file=sys.stderr)
        return 2

    with zipfile.ZipFile(input_path, "r") as archive:
        entries = {
            info.filename: (info, archive.read(info))
            for info in archive.infolist()
        }

    slide_numbers = sorted(
        int(Path(name).stem.removeprefix("slide"))
        for name in entries
        if name.startswith("ppt/slides/slide") and name.endswith(".xml")
    )
    if args.slides.lower() != "all":
        try:
            slide_numbers = [int(value.strip()) for value in args.slides.split(",")]
        except ValueError:
            print("--slides must be 'all' or a comma-separated list of numbers", file=sys.stderr)
            return 2

    layouts: dict[str, ET.Element] = {}
    layout_rels: dict[str, ET.Element] = {}
    masters: dict[str, ET.Element] = {}
    master_rels: dict[str, ET.Element] = {}
    slides: dict[int, tuple[ET.Element, ET.Element]] = {}
    groups_by_layout: dict[str, list[tuple[ET.Element, ET.Element, str]]] = defaultdict(list)
    groups_by_master: dict[str, list[tuple[ET.Element, ET.Element, str]]] = defaultdict(list)

    for slide_number in slide_numbers:
        slide_part = f"ppt/slides/slide{slide_number}.xml"
        slide_rel_part = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
        if slide_part not in entries or slide_rel_part not in entries:
            raise ValueError(f"Slide {slide_number} does not exist")
        slide_root = parse_xml(entries[slide_part][1])
        slide_rel_root = parse_xml(entries[slide_rel_part][1])
        layout_part, master_part = slide_layout_and_master(
            {name: data for name, (_, data) in entries.items()}, slide_number
        )
        layout_rel_part = (
            f"{posixpath.dirname(layout_part)}/_rels/{posixpath.basename(layout_part)}.rels"
        )
        master_rel_part = (
            f"{posixpath.dirname(master_part)}/_rels/{posixpath.basename(master_part)}.rels"
        )
        if layout_part not in entries or layout_rel_part not in entries:
            raise ValueError(f"Missing layout package entries for slide {slide_number}")
        if master_part not in entries or master_rel_part not in entries:
            raise ValueError(f"Missing master package entries for slide {slide_number}")

        slides[slide_number] = (slide_root, slide_rel_root)
        layouts.setdefault(layout_part, parse_xml(entries[layout_part][1]))
        layout_rels.setdefault(layout_part, parse_xml(entries[layout_rel_part][1]))
        masters.setdefault(master_part, parse_xml(entries[master_part][1]))
        master_rels.setdefault(master_part, parse_xml(entries[master_rel_part][1]))
        groups_by_layout[layout_part].append(
            (slide_root, slide_rel_root, slide_part)
        )
        groups_by_master[master_part].append((slide_root, slide_rel_root, slide_part))

    promoted = 0
    for master_part, slide_roots in groups_by_master.items():
        promoted += promote_from_source(
            masters[master_part], master_part, master_rels[master_part], slide_roots
        )
    for layout_part, slide_roots in groups_by_layout.items():
        promoted += promote_from_source(
            layouts[layout_part], layout_part, layout_rels[layout_part], slide_roots
        )

    for slide_number, (root, rels) in slides.items():
        slide_part = f"ppt/slides/slide{slide_number}.xml"
        rel_part = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
        entries[slide_part] = (entries[slide_part][0], serialize_xml(root))
        entries[rel_part] = (entries[rel_part][0], serialize_xml(rels))
    for layout_part, root in layouts.items():
        entries[layout_part] = (entries[layout_part][0], serialize_xml(root))
    for layout_part, rels in layout_rels.items():
        rel_part = (
            f"{posixpath.dirname(layout_part)}/_rels/{posixpath.basename(layout_part)}.rels"
        )
        entries[rel_part] = (entries[rel_part][0], serialize_xml(rels))
    for master_part, root in masters.items():
        entries[master_part] = (entries[master_part][0], serialize_xml(root))

    write_archive(output_path, entries)
    print(f"Promoted {promoted} master/layout elements into {len(slides)} slides")
    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (ET.ParseError, KeyError, StopIteration, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
