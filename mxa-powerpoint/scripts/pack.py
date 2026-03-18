#!/usr/bin/env python3
"""Repack a directory of XML/media files into a .pptx archive.

Condenses pretty-printed XML back to single-line format and creates a
valid .pptx (ZIP) file. Pair with unpack.py.

Usage:
    python pack.py unpacked_dir/ output.pptx
"""

import sys
import zipfile
from pathlib import Path

import defusedxml.minidom as minidom


def _condense_xml(content_bytes):
    """Remove pretty-printing whitespace from XML, returning compact bytes."""
    try:
        dom = minidom.parseString(content_bytes)
        # Serialize without extra whitespace
        compact = dom.toxml(encoding="UTF-8")
        return compact
    except Exception:
        return content_bytes


def pack(source_dir, output_path):
    """Pack a directory into a .pptx file.

    XML files are condensed (whitespace stripped). Other files are stored as-is.

    Args:
        source_dir: Directory containing unpacked presentation files
        output_path: Path for the output .pptx file

    Returns:
        Tuple of (xml_count, other_count)
    """
    source_dir = Path(source_dir)
    output_path = Path(output_path)

    xml_count = 0
    other_count = 0

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(source_dir.rglob("*")):
            if file_path.is_dir():
                continue

            arcname = file_path.relative_to(source_dir).as_posix()
            data = file_path.read_bytes()

            if file_path.suffix in (".xml", ".rels"):
                data = _condense_xml(data)
                xml_count += 1
            else:
                other_count += 1

            zf.writestr(arcname, data)

    return xml_count, other_count


def main():
    if len(sys.argv) < 3:
        print("Usage: python pack.py unpacked_dir/ output.pptx", file=sys.stderr)
        sys.exit(1)

    source_dir = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    if not source_dir.exists():
        print(f"Error: {source_dir} not found", file=sys.stderr)
        sys.exit(1)

    xml_count, other_count = pack(source_dir, output_path)
    print(f"Packed {source_dir} → {output_path}")
    print(f"  {xml_count} XML files (condensed)")
    print(f"  {other_count} other files")


if __name__ == "__main__":
    main()
