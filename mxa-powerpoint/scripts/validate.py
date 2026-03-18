#!/usr/bin/env python3
"""Validate structural integrity of an unpacked PowerPoint directory.

Checks XML well-formedness, relationship integrity, content type consistency,
and common structural issues. Run after manual XML edits, before repacking.

Usage:
    python validate.py unpacked_dir/
    python validate.py unpacked_dir/ --strict

Exit codes:
    0  All checks passed
    1  Fatal error
    2  Warnings found (non-fatal issues)
"""

import sys
from pathlib import Path

import defusedxml.minidom as minidom


# ---------------------------------------------------------------------------
# Validation checks
# ---------------------------------------------------------------------------

class ValidationResult:
    def __init__(self):
        self.errors = []  # Fatal
        self.warnings = []  # Non-fatal

    def error(self, msg):
        self.errors.append(msg)

    def warn(self, msg):
        self.warnings.append(msg)

    @property
    def ok(self):
        return len(self.errors) == 0


def _parse_xml_safe(path):
    """Parse XML and return (dom, error_string)."""
    try:
        dom = minidom.parse(str(path))
        return dom, None
    except Exception as e:
        return None, str(e)


def check_xml_wellformedness(unpacked_dir, result):
    """Verify all XML and .rels files parse without errors."""
    for xml_file in sorted(unpacked_dir.rglob("*")):
        if xml_file.is_file() and xml_file.suffix in (".xml", ".rels"):
            dom, err = _parse_xml_safe(xml_file)
            if err:
                rel = xml_file.relative_to(unpacked_dir)
                result.error(f"XML parse error in {rel}: {err}")


def check_content_types(unpacked_dir, result):
    """Verify [Content_Types].xml exists and references valid files."""
    ct_path = unpacked_dir / "[Content_Types].xml"
    if not ct_path.exists():
        result.error("[Content_Types].xml missing")
        return

    dom, err = _parse_xml_safe(ct_path)
    if err:
        result.error(f"[Content_Types].xml parse error: {err}")
        return

    # Check Override entries point to existing files
    for override in dom.getElementsByTagName("Override"):
        part_name = override.getAttribute("PartName")
        if part_name:
            # PartName starts with /
            rel_path = part_name.lstrip("/")
            full_path = unpacked_dir / rel_path
            if not full_path.exists():
                result.warn(f"[Content_Types].xml references missing file: {rel_path}")


def check_relationships(unpacked_dir, result):
    """Verify all relationship targets exist."""
    for rels_file in sorted(unpacked_dir.rglob("*.rels")):
        dom, err = _parse_xml_safe(rels_file)
        if err:
            continue  # Already caught by wellformedness check

        rels_parent = rels_file.parent.parent
        rel_file_display = rels_file.relative_to(unpacked_dir)

        for rel in dom.getElementsByTagName("Relationship"):
            target = rel.getAttribute("Target")
            rel_type = rel.getAttribute("Type")
            target_mode = rel.getAttribute("TargetMode")

            # Skip external targets (URLs)
            if target_mode == "External" or target.startswith("http"):
                continue

            # Resolve target path
            resolved = (rels_parent / target).resolve()
            if not resolved.exists():
                result.warn(
                    f"{rel_file_display}: target not found: {target}"
                )


def check_presentation_structure(unpacked_dir, result):
    """Verify presentation.xml has valid slide references."""
    pres_path = unpacked_dir / "ppt" / "presentation.xml"
    if not pres_path.exists():
        result.error("ppt/presentation.xml missing")
        return

    dom, err = _parse_xml_safe(pres_path)
    if err:
        return

    # Check sldIdLst entries
    slide_ids = dom.getElementsByTagName("p:sldId")
    if not slide_ids:
        # Try without namespace prefix
        slide_ids = [
            el for el in dom.getElementsByTagName("*")
            if el.localName == "sldId"
        ]

    if not slide_ids:
        result.warn("No slide references found in presentation.xml")
        return

    # Verify each slide ID references a valid relationship
    pres_rels = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"
    rel_ids = set()
    if pres_rels.exists():
        rels_dom, _ = _parse_xml_safe(pres_rels)
        if rels_dom:
            for rel in rels_dom.getElementsByTagName("Relationship"):
                rel_ids.add(rel.getAttribute("Id"))

    for sld_id in slide_ids:
        r_id = None
        for attr_name in sld_id.attributes.keys():
            if "id" in attr_name.lower() and attr_name != "id":
                r_id = sld_id.getAttribute(attr_name)
                break
        if r_id and r_id not in rel_ids:
            result.warn(f"Slide reference {r_id} not found in presentation.xml.rels")


def check_duplicate_shape_ids(unpacked_dir, result):
    """Check for duplicate shape IDs within individual slides."""
    slides_dir = unpacked_dir / "ppt" / "slides"
    if not slides_dir.exists():
        return

    for slide_file in sorted(slides_dir.glob("slide*.xml")):
        dom, err = _parse_xml_safe(slide_file)
        if err:
            continue

        ids_seen = set()
        # Look for cNvPr elements which carry shape IDs
        for el in dom.getElementsByTagName("*"):
            if el.localName == "cNvPr":
                shape_id = el.getAttribute("id")
                if shape_id:
                    if shape_id in ids_seen:
                        rel = slide_file.relative_to(unpacked_dir)
                        result.warn(f"{rel}: duplicate shape id={shape_id}")
                    ids_seen.add(shape_id)


def check_slide_size(unpacked_dir, result):
    """Verify slide dimensions are set (common issue with corrupt files)."""
    pres_path = unpacked_dir / "ppt" / "presentation.xml"
    if not pres_path.exists():
        return

    dom, _ = _parse_xml_safe(pres_path)
    if not dom:
        return

    size_elements = [el for el in dom.getElementsByTagName("*") if el.localName == "sldSz"]
    if not size_elements:
        result.warn("No slide size element found in presentation.xml")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def validate(unpacked_dir, strict=False):
    """Run all validation checks.

    Args:
        unpacked_dir: Path to unpacked presentation directory
        strict: If True, treat warnings as errors

    Returns:
        ValidationResult
    """
    result = ValidationResult()

    check_xml_wellformedness(unpacked_dir, result)
    check_content_types(unpacked_dir, result)
    check_relationships(unpacked_dir, result)
    check_presentation_structure(unpacked_dir, result)
    check_duplicate_shape_ids(unpacked_dir, result)
    check_slide_size(unpacked_dir, result)

    if strict:
        # Promote warnings to errors
        result.errors.extend(result.warnings)
        result.warnings = []

    return result


def main():
    if len(sys.argv) < 2:
        print("Usage: python validate.py unpacked_dir/ [--strict]", file=sys.stderr)
        sys.exit(1)

    unpacked_dir = Path(sys.argv[1])
    strict = "--strict" in sys.argv

    if not unpacked_dir.exists():
        print(f"Error: {unpacked_dir} not found", file=sys.stderr)
        sys.exit(1)

    result = validate(unpacked_dir, strict=strict)

    if result.errors:
        print(f"ERRORS ({len(result.errors)}):")
        for e in result.errors:
            print(f"  ✗ {e}")

    if result.warnings:
        print(f"WARNINGS ({len(result.warnings)}):")
        for w in result.warnings:
            print(f"  ⚠ {w}")

    if result.ok and not result.warnings:
        print("All checks passed.")
        sys.exit(0)
    elif result.ok:
        print(f"\nPassed with {len(result.warnings)} warning(s).")
        sys.exit(2)
    else:
        print(f"\nFailed with {len(result.errors)} error(s).")
        sys.exit(1)


if __name__ == "__main__":
    main()
