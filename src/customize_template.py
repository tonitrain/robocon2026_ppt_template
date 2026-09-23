"""Replace text in a PowerPoint PPTX or POTX without changing the input file."""

from __future__ import annotations

import argparse
import html
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path


TEXT_NODE = re.compile(rb"(<a:t(?:\s[^>]*)?>)(.*?)(</a:t>)", re.DOTALL)


def parse_replacement(value: str) -> tuple[str, str]:
    """Parse one OLD=NEW command-line replacement."""
    if "=" not in value:
        raise argparse.ArgumentTypeError("replacement must use OLD=NEW")
    old, new = value.split("=", 1)
    if not old:
        raise argparse.ArgumentTypeError("replacement text cannot be empty")
    return old, new


def load_replacements(args: argparse.Namespace) -> list[tuple[str, str]]:
    if args.replace_file:
        data = json.loads(args.replace_file.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("replacement JSON must contain an object")
        replacements = []
        for old, new in data.items():
            if not isinstance(old, str) or not isinstance(new, str) or not old:
                raise ValueError("replacement JSON keys and values must be non-empty strings")
            replacements.append((old, new))
        return replacements
    return args.replace or []


def xml_text_values(data: bytes) -> list[str]:
    values = []
    for match in TEXT_NODE.finditer(data):
        raw_value = match.group(2).decode("utf-8")
        values.append(html.unescape(raw_value))
    return values


def replace_xml_text(
    data: bytes,
    replacements: list[tuple[str, str]],
    found: set[str],
) -> bytes:
    def replace_node(match: re.Match[bytes]) -> bytes:
        raw_value = match.group(2).decode("utf-8")
        value = html.unescape(raw_value)
        updated = value
        for old, new in replacements:
            if old in updated:
                found.add(old)
                updated = updated.replace(old, new)
        if updated == value:
            return match.group(0)
        encoded = html.escape(updated, quote=False).encode("utf-8")
        return match.group(1) + encoded + match.group(3)

    return TEXT_NODE.sub(replace_node, data)


def read_xml_entries(path: Path) -> list[tuple[zipfile.ZipInfo, bytes]]:
    with zipfile.ZipFile(path, "r") as archive:
        return [
            (info, archive.read(info))
            for info in archive.infolist()
        ]


def write_archive(path: Path, entries: list[tuple[zipfile.ZipInfo, bytes]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        dir=path.parent,
        prefix=f".{path.name}.",
        suffix=".tmp",
        delete=False,
    ) as temporary:
        temporary_path = Path(temporary.name)
    try:
        with zipfile.ZipFile(temporary_path, "w") as archive:
            for info, data in entries:
                archive.writestr(info, data)
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path, help="input PPTX or POTX")
    parser.add_argument("--output", type=Path, help="new output PPTX or POTX")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--list-text", action="store_true", help="list text nodes")
    mode.add_argument(
        "--replace",
        action="append",
        type=parse_replacement,
        help="replace text using OLD=NEW; repeat for multiple replacements",
    )
    mode.add_argument("--replace-file", type=Path, help="JSON object containing OLD: NEW pairs")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    input_path = args.input.resolve()
    if not input_path.is_file() or not zipfile.is_zipfile(input_path):
        print(f"Input is not a valid PPTX/POTX file: {input_path}", file=sys.stderr)
        return 2

    entries = read_xml_entries(input_path)
    xml_entries = [
        (index, info, data)
        for index, (info, data) in enumerate(entries)
        if info.filename.startswith("ppt/") and info.filename.endswith(".xml")
    ]

    if args.list_text:
        seen = set()
        for _, _, data in xml_entries:
            for value in xml_text_values(data):
                if value and value not in seen:
                    print(value)
                    seen.add(value)
        return 0

    if args.output is None:
        print("--output is required when replacing text", file=sys.stderr)
        return 2
    output_path = args.output.resolve()
    if output_path == input_path:
        print("Output must be different from input", file=sys.stderr)
        return 2

    try:
        replacements = load_replacements(args)
        if not replacements:
            raise ValueError("at least one replacement is required")
    except (OSError, json.JSONDecodeError, ValueError) as error:
        print(f"Invalid replacements: {error}", file=sys.stderr)
        return 2

    found: set[str] = set()
    updated_entries = list(entries)
    for index, info, data in xml_entries:
        updated_entries[index] = (info, replace_xml_text(data, replacements, found))

    missing = [old for old, _ in replacements if old not in found]
    if missing:
        print("Text not found: " + ", ".join(repr(value) for value in missing), file=sys.stderr)
        return 1

    write_archive(output_path, updated_entries)
    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
