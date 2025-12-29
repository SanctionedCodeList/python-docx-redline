"""
Tests for batch operations (apply_edits).

These tests verify that multiple edits can be applied in sequence.
"""

import tempfile
from pathlib import Path

from python_docx_redline import Document, EditResult


def create_test_document() -> Path:
    """Create a test Word document for batch operations."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Create a simple XML document
    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph with target text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph with old text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph to delete.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")
    return doc_path


def test_apply_edits_basic():
    """Test basic apply_edits with multiple operations."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "insert_tracked", "text": " inserted", "after": "target"},
            {"type": "replace_tracked", "find": "old", "replace": "new"},
            {"type": "delete_tracked", "text": "to delete"},
        ]

        results = doc.apply_edits(edits)

        # Should have 3 results
        assert len(results) == 3

        # All should succeed
        assert all(r.success for r in results)

        # Check each result
        assert results[0].edit_type == "insert_tracked"
        assert results[1].edit_type == "replace_tracked"
        assert results[2].edit_type == "delete_tracked"

    finally:
        doc_path.unlink()


def test_apply_edits_with_scope():
    """Test apply_edits with scope parameters."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {
                "type": "insert_tracked",
                "text": " added",
                "after": "paragraph",
                "scope": "First",
            }
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert "inserted" in results[0].message.lower() or "added" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_apply_edits_missing_type():
    """Test apply_edits with missing type field."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [{"text": "some text", "after": "target"}]  # Missing 'type'

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert not results[0].success
        assert results[0].edit_type == "unknown"
        assert "missing" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_apply_edits_missing_parameters():
    """Test apply_edits with missing required parameters."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "insert_tracked", "after": "target"},  # Missing 'text'
            {"type": "delete_tracked"},  # Missing 'text'
            {"type": "replace_tracked", "find": "old"},  # Missing 'replace'
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 3
        assert not any(r.success for r in results)
        assert all("missing" in r.message.lower() for r in results)

    finally:
        doc_path.unlink()


def test_apply_edits_text_not_found():
    """Test apply_edits when text is not found."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {
                "type": "insert_tracked",
                "text": "new text",
                "after": "nonexistent text",
            }
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert not results[0].success
        assert "not found" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_apply_edits_stop_on_error():
    """Test apply_edits with stop_on_error=True."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "insert_tracked", "text": " first", "after": "target"},
            {
                "type": "insert_tracked",
                "text": " error",
                "after": "nonexistent",
            },  # This will fail
            {"type": "insert_tracked", "text": " third", "after": "target"},
        ]

        results = doc.apply_edits(edits, stop_on_error=True)

        # Should only have 2 results (stops after first error)
        assert len(results) == 2
        assert results[0].success
        assert not results[1].success

    finally:
        doc_path.unlink()


def test_apply_edits_continue_on_error():
    """Test apply_edits with stop_on_error=False (default)."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "insert_tracked", "text": " first", "after": "target"},
            {
                "type": "insert_tracked",
                "text": " error",
                "after": "nonexistent",
            },  # This will fail
            {"type": "insert_tracked", "text": " third", "after": "target"},
        ]

        results = doc.apply_edits(edits, stop_on_error=False)

        # Should have all 3 results
        assert len(results) == 3
        assert results[0].success
        assert not results[1].success
        assert results[2].success

    finally:
        doc_path.unlink()


def test_apply_edits_unknown_type():
    """Test apply_edits with unknown edit type."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [{"type": "unknown_operation", "some_param": "value"}]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert not results[0].success
        assert "unknown" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_edit_result_string_representation():
    """Test EditResult string representation."""
    success_result = EditResult(success=True, edit_type="insert_tracked", message="Inserted text")
    assert "✓" in str(success_result)
    assert "insert_tracked" in str(success_result)

    failure_result = EditResult(success=False, edit_type="delete_tracked", message="Text not found")
    assert "✗" in str(failure_result)
    assert "delete_tracked" in str(failure_result)


def test_apply_edits_empty_list():
    """Test apply_edits with empty list."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        results = doc.apply_edits([])

        assert len(results) == 0

    finally:
        doc_path.unlink()


def test_apply_edits_with_author():
    """Test apply_edits with custom author."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {
                "type": "insert_tracked",
                "text": " custom",
                "after": "target",
                "author": "Custom Author",
            }
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success

    finally:
        doc_path.unlink()


# Tests for Phase 2 structural operations


def create_structured_document() -> Path:
    """Create a test document with headings for structural operations."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Introduction</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Introduction content</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Methods</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Methods content</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")
    return doc_path


def test_apply_edits_insert_paragraph():
    """Test batch insert_paragraph operation."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [
            {
                "type": "insert_paragraph",
                "text": "New paragraph text",
                "after": "Introduction content",
            }
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert results[0].edit_type == "insert_paragraph"
        assert "inserted paragraph" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_apply_edits_insert_paragraph_with_style():
    """Test batch insert_paragraph with style."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [
            {
                "type": "insert_paragraph",
                "text": "New heading",
                "after": "Introduction content",
                "style": "Heading2",
            }
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success

    finally:
        doc_path.unlink()


def test_apply_edits_insert_paragraphs():
    """Test batch insert_paragraphs operation."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [
            {
                "type": "insert_paragraphs",
                "texts": ["First new para", "Second new para", "Third new para"],
                "after": "Methods content",
            }
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert results[0].edit_type == "insert_paragraphs"
        assert "3 paragraphs" in results[0].message

    finally:
        doc_path.unlink()


def test_apply_edits_delete_section():
    """Test batch delete_section operation."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [{"type": "delete_section", "heading": "Methods"}]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert results[0].edit_type == "delete_section"
        assert "deleted section" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_apply_edits_mixed_phase_1_and_2():
    """Test mixing Phase 1 and Phase 2 operations."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "insert_tracked", "text": " updated", "after": "Introduction content"},
            {
                "type": "insert_paragraph",
                "text": "New paragraph between sections",
                "after": "Introduction content",
            },
            {"type": "replace_tracked", "find": "Methods content", "replace": "New methods"},
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 3
        # Debug: print results if any failed
        if not all(r.success for r in results):
            for i, r in enumerate(results):
                print(f"Result {i}: success={r.success}, type={r.edit_type}, message={r.message}")

        assert all(
            r.success for r in results
        ), f"Failed results: {[r.message for r in results if not r.success]}"
        assert results[0].edit_type == "insert_tracked"
        assert results[1].edit_type == "insert_paragraph"
        assert results[2].edit_type == "replace_tracked"

    finally:
        doc_path.unlink()


def test_apply_edits_insert_paragraph_missing_params():
    """Test insert_paragraph with missing parameters."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "insert_paragraph", "after": "Introduction"},  # Missing text
            {"type": "insert_paragraph", "text": "New text"},  # Missing after/before
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 2
        assert not any(r.success for r in results)
        assert all("missing" in r.message.lower() for r in results)

    finally:
        doc_path.unlink()


def test_apply_edits_insert_paragraphs_missing_params():
    """Test insert_paragraphs with missing parameters."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "insert_paragraphs", "after": "Introduction"},  # Missing texts
            {"type": "insert_paragraphs", "texts": ["Text 1", "Text 2"]},  # Missing after/before
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 2
        assert not any(r.success for r in results)
        assert all("missing" in r.message.lower() for r in results)

    finally:
        doc_path.unlink()


def test_apply_edits_delete_section_missing_params():
    """Test delete_section with missing parameters."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [{"type": "delete_section"}]  # Missing heading

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert not results[0].success
        assert "missing" in results[0].message.lower()

    finally:
        doc_path.unlink()


# Tests for Phase 5: Per-edit track support


def test_apply_edits_generic_insert():
    """Test generic insert edit type."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [{"type": "insert", "text": " inserted", "after": "target"}]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert results[0].edit_type == "insert"
        assert "inserted" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_apply_edits_generic_delete():
    """Test generic delete edit type."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [{"type": "delete", "text": "to delete"}]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert results[0].edit_type == "delete"
        assert "deleted" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_apply_edits_generic_replace():
    """Test generic replace edit type."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [{"type": "replace", "find": "old", "replace": "new"}]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert results[0].edit_type == "replace"
        assert "replaced" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_apply_edits_with_track_field():
    """Test per-edit track field."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "replace", "find": "old", "replace": "new1", "track": False},
            {"type": "replace", "find": "target", "replace": "target2", "track": True},
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 2
        assert all(r.success for r in results)

        # First edit should not be tracked (message should not say tracked)
        assert "(tracked)" not in results[0].message

        # Second edit should be tracked (message should say tracked)
        assert "(tracked)" in results[1].message

    finally:
        doc_path.unlink()


def test_apply_edits_default_track_true():
    """Test default_track=True parameter."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "replace", "find": "old", "replace": "new"},  # No track field
        ]

        # With default_track=True
        results = doc.apply_edits(edits, default_track=True)

        assert len(results) == 1
        assert results[0].success
        # Should be tracked
        assert "(tracked)" in results[0].message

    finally:
        doc_path.unlink()


def test_apply_edits_default_track_false():
    """Test default_track=False parameter (default behavior)."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "replace", "find": "old", "replace": "new"},  # No track field
        ]

        # With default_track=False (default)
        results = doc.apply_edits(edits, default_track=False)

        assert len(results) == 1
        assert results[0].success
        # Should NOT be tracked
        assert "(tracked)" not in results[0].message

    finally:
        doc_path.unlink()


def test_apply_edits_track_field_overrides_default():
    """Test that per-edit track field overrides default_track."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "replace", "find": "old", "replace": "new", "track": True},
        ]

        # Even with default_track=False, explicit track=True should win
        results = doc.apply_edits(edits, default_track=False)

        assert len(results) == 1
        assert results[0].success
        assert "(tracked)" in results[0].message

    finally:
        doc_path.unlink()


def test_apply_edits_backwards_compat():
    """Test that existing edits still work without track field (backwards compat)."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        # These are the old-style tracked operations
        edits = [
            {"type": "insert_tracked", "text": " inserted", "after": "target"},
            {"type": "replace_tracked", "find": "old", "replace": "new"},
            {"type": "delete_tracked", "text": "to delete"},
        ]

        # Should work exactly as before
        results = doc.apply_edits(edits)

        assert len(results) == 3
        assert all(r.success for r in results)

    finally:
        doc_path.unlink()


def test_apply_edits_mixed_types():
    """Test mixing generic and tracked operations."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            # Generic with track=False
            {"type": "insert", "text": " untracked", "after": "First"},
            # Legacy tracked
            {"type": "replace_tracked", "find": "old", "replace": "new"},
            # Generic with track=True
            {"type": "delete", "text": "paragraph", "track": True},
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 3
        assert results[0].success
        assert results[1].success
        assert results[2].success

    finally:
        doc_path.unlink()


def test_apply_edits_generic_insert_missing_params():
    """Test generic insert with missing parameters."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "insert", "after": "target"},  # Missing text
            {"type": "insert", "text": "new text"},  # Missing after/before
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 2
        assert not any(r.success for r in results)
        assert all("missing" in r.message.lower() for r in results)

    finally:
        doc_path.unlink()


def test_apply_edits_generic_delete_missing_params():
    """Test generic delete with missing parameters."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "delete"},  # Missing text
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert not results[0].success
        assert "missing" in results[0].message.lower()

    finally:
        doc_path.unlink()


def test_apply_edits_generic_replace_missing_params():
    """Test generic replace with missing parameters."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "replace", "replace": "new"},  # Missing find
            {"type": "replace", "find": "old"},  # Missing replace
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 2
        assert not any(r.success for r in results)
        assert all("missing" in r.message.lower() for r in results)

    finally:
        doc_path.unlink()


def test_apply_edit_file_with_default_track():
    """Test apply_edit_file with default_track in YAML."""
    import tempfile

    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create a YAML file with default_track
        yaml_content = """
default_track: true

edits:
  - type: replace
    find: "old"
    replace: "new"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path)

        assert len(results) == 1
        assert results[0].success
        # Should be tracked because default_track: true in file
        assert "(tracked)" in results[0].message

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_with_per_edit_track():
    """Test apply_edit_file with per-edit track in YAML."""
    import tempfile

    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create a YAML file with per-edit track
        yaml_content = """
default_track: false

edits:
  - type: replace
    find: "old"
    replace: "new"
    track: true
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path)

        assert len(results) == 1
        assert results[0].success
        # Should be tracked because track: true overrides default_track: false
        assert "(tracked)" in results[0].message

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_caller_override():
    """Test that caller-provided default_track overrides file default."""
    import tempfile

    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create a YAML file with default_track: false
        yaml_content = """
default_track: false

edits:
  - type: replace
    find: "old"
    replace: "new"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)
        # Caller overrides with default_track=True
        results = doc.apply_edit_file(yaml_path, default_track=True)

        assert len(results) == 1
        assert results[0].success
        # Should be tracked because caller's default_track=True overrides file
        assert "(tracked)" in results[0].message

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_json_with_track():
    """Test apply_edit_file with track fields in JSON."""
    import json
    import tempfile

    doc_path = create_test_document()
    json_path = Path(tempfile.mktemp(suffix=".json"))

    try:
        # Create a JSON file with track fields
        json_content = {
            "default_track": False,
            "edits": [{"type": "replace", "find": "old", "replace": "new", "track": True}],
        }
        json_path.write_text(json.dumps(json_content), encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(json_path, format="json")

        assert len(results) == 1
        assert results[0].success
        assert "(tracked)" in results[0].message

    finally:
        doc_path.unlink()
        json_path.unlink()


def test_apply_edit_file_no_default_track():
    """Test apply_edit_file when no default_track is specified."""
    import tempfile

    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create a YAML file without default_track
        yaml_content = """
edits:
  - type: replace
    find: "old"
    replace: "new"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path)

        assert len(results) == 1
        assert results[0].success
        # Should NOT be tracked (default is False)
        assert "(tracked)" not in results[0].message

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edits_insert_with_before():
    """Test generic insert with 'before' parameter."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [{"type": "insert", "text": "prefix ", "before": "First"}]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert "before" in results[0].message

    finally:
        doc_path.unlink()


def test_apply_edits_insert_paragraph_track_false():
    """Test insert_paragraph with track=False."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [
            {
                "type": "insert_paragraph",
                "text": "New untracked paragraph",
                "after": "Introduction content",
                "track": False,
            }
        ]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert results[0].edit_type == "insert_paragraph"

    finally:
        doc_path.unlink()


def test_apply_edits_delete_section_track_false():
    """Test delete_section with track=False."""
    doc_path = create_structured_document()
    try:
        doc = Document(doc_path)

        edits = [{"type": "delete_section", "heading": "Methods", "track": False}]

        results = doc.apply_edits(edits)

        assert len(results) == 1
        assert results[0].success
        assert results[0].edit_type == "delete_section"

    finally:
        doc_path.unlink()


# Run tests with: pytest tests/test_batch_operations.py -v
