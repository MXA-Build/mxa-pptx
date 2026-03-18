#!/usr/bin/env python3
"""Clean orphaned files from an unpacked PowerPoint directory.

Scans the presentation structure, identifies slides/media/layouts that are
not referenced by any relationship file, and removes them. Run after manual
XML edits, before repacking.

Usage:
    python clean.py unpacked_dir/
    python clean.py unpacked_dir/ --dry-run
"""

import sys
from pathlib import Path

import defusedxml.minidom as minidom


def _parse_xml(path):
    """Safely parse an XML file, returning a DOM or None."""
    try:
        return minidom.parse(str(path))
    except Exception as e:
        print(f"  Warning: could not parse {path}: {e}", file=sys.stderr)
        return None


def _collect_referenced_targets(rels_dir):
    """Collect all Target values from .rels files under a directory."""
    targets = set()
    if not rels_dir.exists():
        return targets

    for rels_file in rels_dir.rglob("*.rels"):
        dom = _parse_xml(rels_file)
        if dom is None:
            continue
        for rel in dom.getElementsByTagName("Relationship"):
            target = rel.getAttribute("Target")
            if target:
                # Resolve relative paths
                if target.startswith("../"):
                    # Resolve relative to the rels file's parent's parent
                    resolved = (rels_file.parent.parent / target).resolve()
                else:
                    resolved = (rels_file.parent.parent / target).resolve()
                targets.add(resolved)
    return targets


def _collect_content_type_refs(content_types_path):
    """Collect all PartName values from [Content_Types].xml."""
    parts = set()
    dom = _parse_xml(content_types_path)
    if dom is None:
        return parts
    for override in dom.getElementsByTagName("Override"):
        part_name = override.getAttribute("PartName")
        if part_name:
            # PartName is like "/ppt/slides/slide1.xml"
            parts.add(part_name.lstrip("/"))
    return parts


def clean(unpacked_dir, dry_run=False):
    """Remove orphaned files from an unpacked presentation directory.

    Args:
        unpacked_dir: Path to the unpacked directory
        dry_run: If True, only report what would be deleted

    Returns:
        List of removed (or would-be-removed) file paths
    """
    unpacked_dir = Path(unpacked_dir).resolve()
    removed = []

    # Collect all referenced targets from all .rels files
    referenced = set()
    for rels_file in unpacked_dir.rglob("*.rels"):
        dom = _parse_xml(rels_file)
        if dom is None:
            continue
        rels_parent = rels_file.parent.parent  # _rels/ is always a child
        for rel in dom.getElementsByTagName("Relationship"):
            target = rel.getAttribute("Target")
            if target:
                resolved = (rels_parent / target).resolve()
                referenced.add(resolved)

    # Also consider content types
    ct_path = unpacked_dir / "[Content_Types].xml"
    content_parts = set()
    if ct_path.exists():
        content_parts = _collect_content_type_refs(ct_path)

    # Check slides directory for orphans
    slides_dir = unpacked_dir / "ppt" / "slides"
    if slides_dir.exists():
        for slide_file in sorted(slides_dir.glob("slide*.xml")):
            if slide_file.resolve() not in referenced:
                rel_path = slide_file.relative_to(unpacked_dir)
                # Double-check via content types
                if str(rel_path).replace("\\", "/") not in content_parts:
                    removed.append(slide_file)
                    if not dry_run:
                        slide_file.unlink()
                        # Also remove corresponding .rels
                        rels = slides_dir / "_rels" / (slide_file.name + ".rels")
                        if rels.exists():
                            rels.unlink()
                            removed.append(rels)

    # Check media directory for orphaned images
    media_dir = unpacked_dir / "ppt" / "media"
    if media_dir.exists():
        for media_file in sorted(media_dir.iterdir()):
            if media_file.is_file() and media_file.resolve() not in referenced:
                removed.append(media_file)
                if not dry_run:
                    media_file.unlink()

    # Check slide layouts for orphans
    layouts_dir = unpacked_dir / "ppt" / "slideLayouts"
    if layouts_dir.exists():
        for layout_file in sorted(layouts_dir.glob("slideLayout*.xml")):
            if layout_file.resolve() not in referenced:
                removed.append(layout_file)
                if not dry_run:
                    layout_file.unlink()
                    rels = layouts_dir / "_rels" / (layout_file.name + ".rels")
                    if rels.exists():
                        rels.unlink()
                        removed.append(rels)

    return removed


def main():
    if len(sys.argv) < 2:
        print("Usage: python clean.py unpacked_dir/ [--dry-run]", file=sys.stderr)
        sys.exit(1)

    unpacked_dir = Path(sys.argv[1])
    dry_run = "--dry-run" in sys.argv

    if not unpacked_dir.exists():
        print(f"Error: {unpacked_dir} not found", file=sys.stderr)
        sys.exit(1)

    removed = clean(unpacked_dir, dry_run=dry_run)

    if dry_run:
        print(f"Dry run: {len(removed)} orphan(s) would be removed:")
    else:
        print(f"Removed {len(removed)} orphan(s):")

    for f in removed:
        print(f"  {f}")

    if not removed:
        print("No orphaned files found.")


if __name__ == "__main__":
    main()
