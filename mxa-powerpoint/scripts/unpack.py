#!/usr/bin/env python3
"""Unpack a .pptx file into a directory of pretty-printed XML files.

Extracts the ZIP archive and reformats all XML files for human-readable
editing. Non-XML files (images, media) are extracted as-is.

Usage:
    python unpack.py input.pptx output_dir/
"""

import sys
import zipfile
from pathlib import Path

import defusedxml.minidom as minidom


def _pretty_print_xml(raw_bytes):
    """Parse and pretty-print XML bytes. Returns formatted string."""
    try:
        dom = minidom.parseString(raw_bytes)
        pretty = dom.toprettyxml(indent="  ", encoding="UTF-8")
        # Remove extra blank lines that toprettyxml tends to add
        lines = pretty.decode("utf-8").split("\n")
        cleaned = "\n".join(line for line in lines if line.strip())
        return cleaned.encode("utf-8")
    except Exception:
        # If XML parsing fails, return raw bytes
        return raw_bytes


def unpack(pptx_path, output_dir):
    """Unpack a .pptx into a directory with pretty-printed XML.

    Args:
        pptx_path: Path to the .pptx file
        output_dir: Directory to extract into (created if needed)

    Returns:
        Tuple of (xml_count, other_count)
    """
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)

    if output_dir.exists():
        print(f"Warning: {output_dir} already exists, files may be overwritten", file=sys.stderr)
    output_dir.mkdir(parents=True, exist_ok=True)

    xml_count = 0
    other_count = 0

    with zipfile.ZipFile(pptx_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                (output_dir / info.filename).mkdir(parents=True, exist_ok=True)
                continue

            data = zf.read(info.filename)
            out_path = output_dir / info.filename
            out_path.parent.mkdir(parents=True, exist_ok=True)

            # Pretty-print XML files
            if info.filename.endswith((".xml", ".rels")):
                data = _pretty_print_xml(data)
                xml_count += 1
            else:
                other_count += 1

            out_path.write_bytes(data)

    return xml_count, other_count


def main():
    if len(sys.argv) < 3:
        print("Usage: python unpack.py input.pptx output_dir/", file=sys.stderr)
        sys.exit(1)

    pptx_path = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])

    if not pptx_path.exists():
        print(f"Error: {pptx_path} not found", file=sys.stderr)
        sys.exit(1)

    xml_count, other_count = unpack(pptx_path, output_dir)
    print(f"Unpacked {pptx_path} → {output_dir}")
    print(f"  {xml_count} XML files (pretty-printed)")
    print(f"  {other_count} other files (media, images)")


if __name__ == "__main__":
    main()
