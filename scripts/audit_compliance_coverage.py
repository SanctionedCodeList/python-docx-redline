#!/usr/bin/env python3
"""
Audit OOXML spec compliance test coverage.

This script analyzes the Document class public methods and compares them
against what's tested in the OOXML spec compliance tests.

Usage:
    python scripts/audit_compliance_coverage.py [--generate-stubs]

Options:
    --generate-stubs    Generate test stub code for missing methods
"""

import ast
import re
import sys
from pathlib import Path


def get_document_public_methods(doc_file: Path) -> list[tuple[str, str, list[str]]]:
    """Extract public methods from Document class with their signatures."""
    doc_source = doc_file.read_text()
    tree = ast.parse(doc_source)

    public_methods = []
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef) and node.name == "Document":
            for item in node.body:
                if isinstance(item, ast.FunctionDef):
                    if not item.name.startswith("_"):
                        docstring = ast.get_docstring(item) or ""
                        first_line = docstring.split("\n")[0] if docstring else "No description"

                        # Get parameter names (excluding self)
                        params = [arg.arg for arg in item.args.args if arg.arg != "self"]

                        public_methods.append((item.name, first_line, params))

    return public_methods


def get_tested_methods(test_file: Path) -> set[str]:
    """Get methods called in the compliance test file."""
    test_source = test_file.read_text()

    # Find all doc.method_name( patterns
    called_methods = set(re.findall(r"doc\d*\.(\w+)\s*\(", test_source))
    # Also check for Document.method patterns
    called_methods.update(re.findall(r"Document\([^)]+\)\.(\w+)\s*\(", test_source))

    return called_methods


# Feature categories for organization
CATEGORIES = {
    "Core Tracked Changes": [
        "insert_tracked",
        "delete_tracked",
        "replace_tracked",
        "move_tracked",
    ],
    "Accept/Reject Changes": [
        "accept_all_changes",
        "accept_by_author",
        "accept_change",
        "accept_changes",
        "accept_deletions",
        "accept_format_changes",
        "accept_insertions",
        "reject_all_changes",
        "reject_by_author",
        "reject_change",
        "reject_changes",
        "reject_deletions",
        "reject_format_changes",
        "reject_insertions",
    ],
    "Comments": [
        "add_comment",
        "comments",
        "get_comments",
        "delete_all_comments",
    ],
    "Headers/Footers": [
        "headers",
        "footers",
        "insert_in_header",
        "insert_in_footer",
        "replace_in_header",
        "replace_in_footer",
    ],
    "Tables": [
        "tables",
        "find_table",
        "replace_in_table",
        "update_cell",
        "insert_table_row",
        "insert_table_column",
        "delete_table_row",
        "delete_table_column",
    ],
    "Footnotes/Endnotes": [
        "footnotes",
        "endnotes",
        "insert_footnote",
        "insert_endnote",
    ],
    "Format Changes": [
        "format_tracked",
        "format_text",
        "format_paragraph_tracked",
        "copy_format",
        "apply_style",
    ],
    "Paragraphs/Sections": [
        "paragraphs",
        "sections",
        "insert_paragraph",
        "insert_paragraphs",
        "delete_section",
    ],
    "Document Comparison": [
        "compare_to",
        "comparison_stats",
    ],
    "Export/Reports": [
        "export_changes_json",
        "export_changes_markdown",
        "export_changes_html",
        "generate_change_report",
    ],
    "Query/Inspection": [
        "tracked_changes",
        "get_tracked_changes",
        "has_tracked_changes",
        "get_text",
    ],
    "Pattern Operations": [
        "normalize_dates",
        "normalize_currency",
        "update_section_references",
    ],
    "Batch/File Operations": [
        "apply_edits",
        "apply_edit_file",
    ],
    "Save/Validate": [
        "save",
        "save_to_bytes",
        "validate",
    ],
}


def categorize_method(method_name: str) -> str:
    """Find the category for a method."""
    for category, methods in CATEGORIES.items():
        if method_name in methods:
            return category
    return "Uncategorized"


def generate_test_stub(method_name: str, description: str, params: list[str]) -> str:
    """Generate a test stub for a method."""
    param_hints = ", ".join(params) if params else "no params"
    return f'''
    def test_{method_name}(self) -> None:
        """Test {method_name} produces valid OOXML.

        Method: {description}
        Params: {param_hints}
        """
        doc_path = create_test_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        try:
            doc = Document(doc_path)
            # TODO: Call doc.{method_name}(...) with appropriate arguments
            doc.save(output_path)

            errors = validate_document(output_path)
            assert errors == [], f"Validation errors: {{errors}}"
        finally:
            doc_path.unlink()
            output_path.unlink(missing_ok=True)
'''


def main():
    generate_stubs = "--generate-stubs" in sys.argv

    # Paths
    root = Path(__file__).parent.parent
    doc_file = root / "src" / "python_docx_redline" / "document.py"
    test_file = root / "tests" / "test_ooxml_spec_compliance.py"

    if not doc_file.exists():
        print(f"Error: Document file not found: {doc_file}")
        sys.exit(1)

    if not test_file.exists():
        print(f"Error: Test file not found: {test_file}")
        sys.exit(1)

    # Get methods
    public_methods = get_document_public_methods(doc_file)
    tested_methods = get_tested_methods(test_file)

    # Analyze
    tested = []
    not_tested = []

    for method, desc, params in public_methods:
        if method in tested_methods:
            tested.append((method, desc, params))
        else:
            not_tested.append((method, desc, params))

    # Report
    print("=" * 80)
    print("OOXML SPEC COMPLIANCE TEST COVERAGE AUDIT")
    print("=" * 80)
    print()

    # By category
    print("COVERAGE BY CATEGORY:")
    print("-" * 60)

    for category in CATEGORIES:
        category_methods = CATEGORIES[category]
        tested_in_category = [m for m in category_methods if m in tested_methods]
        coverage = len(tested_in_category) / len(category_methods) * 100 if category_methods else 0

        status = "✓" if coverage == 100 else "○" if coverage > 0 else "✗"
        tested_count = len(tested_in_category)
        total_count = len(category_methods)
        print(f"  {status} {category}: {tested_count}/{total_count} ({coverage:.0f}%)")

        if coverage < 100 and not generate_stubs:
            missing = [m for m in category_methods if m not in tested_methods]
            for m in missing[:3]:  # Show first 3 missing
                print(f"      - {m}")
            if len(missing) > 3:
                print(f"      ... and {len(missing) - 3} more")

    print()
    print("-" * 60)
    total_methods = len(public_methods)
    total_tested = len(tested)
    pct = 100 * total_tested / total_methods
    print(f"TOTAL: {total_tested}/{total_methods} methods tested ({pct:.1f}%)")
    print()

    # Summary
    if not_tested:
        print(f"Missing coverage for {len(not_tested)} methods.")

        if generate_stubs:
            print()
            print("=" * 80)
            print("GENERATED TEST STUBS")
            print("=" * 80)
            print()

            # Group by category
            by_category: dict[str, list] = {}
            for method, desc, params in not_tested:
                cat = categorize_method(method)
                if cat not in by_category:
                    by_category[cat] = []
                by_category[cat].append((method, desc, params))

            for category, methods in sorted(by_category.items()):
                print(f"\n# === {category} ===\n")
                print(f"class Test{category.replace('/', '').replace(' ', '')}:")
                print(f'    """Test {category} produce valid OOXML."""')
                for method, desc, params in methods:
                    print(generate_test_stub(method, desc, params))

    return 0 if len(not_tested) == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
