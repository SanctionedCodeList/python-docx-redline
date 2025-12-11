#!/usr/bin/env python3
"""Unpack and format XML contents of Office files (.docx, .pptx, .xlsx)"""

import random
import sys
import zipfile
from pathlib import Path

import defusedxml.minidom


def unpack_document(input_file, output_dir):
    """Unpack an Office document and pretty-print all XML files.

    Args:
        input_file: Path to the Office file (.docx, .pptx, .xlsx)
        output_dir: Path to the output directory

    Returns:
        str: Suggested RSID for .docx files, None otherwise
    """
    input_file = Path(input_file)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    zipfile.ZipFile(input_file).extractall(output_path)

    # Pretty print all XML files
    xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
    for xml_file in xml_files:
        content = xml_file.read_text(encoding="utf-8")
        dom = defusedxml.minidom.parseString(content)
        xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="ascii"))

    # For .docx files, suggest an RSID for tracked changes
    if input_file.suffix.lower() == ".docx":
        return "".join(random.choices("0123456789ABCDEF", k=8))
    return None


def main():
    """Command-line interface for unpacking Office documents."""
    assert len(sys.argv) == 3, "Usage: python unpack.py <office_file> <output_dir>"
    input_file, output_dir = sys.argv[1], sys.argv[2]

    suggested_rsid = unpack_document(input_file, output_dir)
    if suggested_rsid:
        print(f"Suggested RSID for edit session: {suggested_rsid}")


if __name__ == "__main__":
    main()
