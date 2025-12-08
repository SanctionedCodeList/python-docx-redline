"""
Tests for batch operations (apply_edits).

These tests verify that multiple edits can be applied in sequence.
"""

import tempfile
from pathlib import Path

from docx_redline import Document, EditResult


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


# Run tests with: pytest tests/test_batch_operations.py -v
