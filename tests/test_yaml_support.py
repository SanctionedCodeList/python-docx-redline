"""
Tests for YAML file support (apply_edit_file).

These tests verify that edits can be loaded from YAML files and applied.
"""

import json
import tempfile
from pathlib import Path

import pytest

from python_docx_redline import Document, ValidationError


def create_test_document() -> Path:
    """Create a test Word document."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

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
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")
    return doc_path


def test_apply_edit_file_yaml_basic():
    """Test basic YAML file loading and application."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create YAML file
        yaml_content = """
edits:
  - type: insert_tracked
    text: " inserted"
    after: "target"
  - type: replace_tracked
    find: "old"
    replace: "new"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path)

        assert len(results) == 2
        assert all(r.success for r in results)

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_json():
    """Test JSON file loading and application."""
    doc_path = create_test_document()
    json_path = Path(tempfile.mktemp(suffix=".json"))

    try:
        # Create JSON file
        json_data = {
            "edits": [
                {"type": "insert_tracked", "text": " inserted", "after": "target"},
                {"type": "replace_tracked", "find": "old", "replace": "new"},
            ]
        }
        json_path.write_text(json.dumps(json_data), encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(json_path, format="json")

        assert len(results) == 2
        assert all(r.success for r in results)

    finally:
        doc_path.unlink()
        json_path.unlink()


def test_apply_edit_file_with_scope():
    """Test YAML file with scope specifications."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        yaml_content = """
edits:
  - type: insert_tracked
    text: " added"
    after: "paragraph"
    scope: "First"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path)

        assert len(results) == 1
        assert results[0].success

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_not_found():
    """Test error when file doesn't exist."""
    doc_path = create_test_document()

    try:
        doc = Document(doc_path)

        with pytest.raises(FileNotFoundError):
            doc.apply_edit_file("nonexistent.yaml")

    finally:
        doc_path.unlink()


def test_apply_edit_file_invalid_yaml():
    """Test error when YAML is invalid."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create invalid YAML
        yaml_path.write_text("invalid: yaml: content: [", encoding="utf-8")

        doc = Document(doc_path)

        with pytest.raises(ValidationError, match="parse"):
            doc.apply_edit_file(yaml_path)

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_missing_edits_key():
    """Test error when 'edits' key is missing."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create YAML without 'edits' key
        yaml_content = """
document: test.docx
author: Test Author
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)

        with pytest.raises(ValidationError, match="'edits'"):
            doc.apply_edit_file(yaml_path)

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_edits_not_list():
    """Test error when 'edits' is not a list."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        yaml_content = """
edits:
  type: insert_tracked
  text: "bad format"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)

        with pytest.raises(ValidationError, match="list"):
            doc.apply_edit_file(yaml_path)

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_stop_on_error():
    """Test stop_on_error parameter."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        yaml_content = """
edits:
  - type: insert_tracked
    text: " first"
    after: "target"
  - type: insert_tracked
    text: " error"
    after: "nonexistent"
  - type: insert_tracked
    text: " third"
    after: "target"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path, stop_on_error=True)

        # Should stop after first error
        assert len(results) == 2
        assert results[0].success
        assert not results[1].success

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_unsupported_format():
    """Test error with unsupported format."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".txt"))

    try:
        yaml_path.write_text("edits: []", encoding="utf-8")

        doc = Document(doc_path)

        with pytest.raises(ValidationError, match="Unsupported format"):
            doc.apply_edit_file(yaml_path, format="xml")

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_not_dict():
    """Test error when file doesn't contain a dictionary."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create YAML that's just a list
        yaml_content = """
- item1
- item2
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)

        with pytest.raises(ValidationError, match="dictionary"):
            doc.apply_edit_file(yaml_path)

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_with_metadata():
    """Test YAML file with metadata (ignored)."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        yaml_content = """
document: test.docx
author: Test Author
output: output.docx

edits:
  - type: insert_tracked
    text: " inserted"
    after: "target"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path)

        assert len(results) == 1
        assert results[0].success

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_apply_edit_file_empty_edits():
    """Test YAML file with empty edits list."""
    doc_path = create_test_document()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        yaml_content = """
edits: []
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path)

        assert len(results) == 0

    finally:
        doc_path.unlink()
        yaml_path.unlink()


# Run tests with: pytest tests/test_yaml_support.py -v
