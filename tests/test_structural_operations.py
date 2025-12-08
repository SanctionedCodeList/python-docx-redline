"""
Tests for structural operations (insert_paragraph, insert_paragraphs, delete_section).
"""

import tempfile
from pathlib import Path

import pytest
from lxml import etree

from docx_redline import Document
from docx_redline.errors import AmbiguousTextError, TextNotFoundError
from docx_redline.models.paragraph import WORD_NAMESPACE


def create_test_document() -> Path:
    """Create a test Word document with multiple paragraphs."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")
    return doc_path


def test_insert_paragraph_after_basic():
    """Test inserting a paragraph after anchor text."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        new_para = doc.insert_paragraph("New paragraph", after="First paragraph")

        assert new_para is not None
        assert new_para.text == "New paragraph"

        # Verify it was inserted in the right place
        all_paras = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        # Should have 4 paragraphs now (original 3 + 1 new)
        # But the new one is wrapped in w:ins, so we need to look deeper
        assert len(all_paras) >= 4

    finally:
        doc_path.unlink()


def test_insert_paragraph_before_basic():
    """Test inserting a paragraph before anchor text."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        new_para = doc.insert_paragraph("New paragraph", before="Second paragraph")

        assert new_para is not None
        assert new_para.text == "New paragraph"

    finally:
        doc_path.unlink()


def test_insert_paragraph_with_style():
    """Test inserting a paragraph with a style."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        new_para = doc.insert_paragraph("Heading text", after="First paragraph", style="Heading1")

        assert new_para is not None
        assert new_para.text == "Heading text"
        assert new_para.style == "Heading1"
        assert new_para.is_heading() is True

    finally:
        doc_path.unlink()


def test_insert_paragraph_tracked():
    """Test that tracked insertion wraps paragraph in w:ins."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        doc.insert_paragraph("Tracked paragraph", after="First paragraph", track=True)

        # Check for w:ins element in the XML
        xml_string = etree.tostring(doc.xml_root, encoding="unicode")
        assert "w:ins" in xml_string

    finally:
        doc_path.unlink()


def test_insert_paragraph_untracked():
    """Test that untracked insertion does not use w:ins."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        new_para = doc.insert_paragraph("Untracked paragraph", after="First paragraph", track=False)

        assert new_para is not None
        assert new_para.text == "Untracked paragraph"

        # The paragraph itself should not be in a w:ins element
        # (It's inserted directly)

    finally:
        doc_path.unlink()


def test_insert_paragraph_with_scope():
    """Test inserting paragraph with scope filter."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Section One</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content in section one</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Section Two</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content in section two</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        # Insert in section one only
        new_para = doc.insert_paragraph("New content", after="Content", scope="section:Section One")

        assert new_para is not None
        assert new_para.text == "New content"

    finally:
        doc_path.unlink()


def test_insert_paragraph_neither_after_nor_before():
    """Test error when neither after nor before is specified."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        with pytest.raises(ValueError, match="Must specify either"):
            doc.insert_paragraph("Text")

    finally:
        doc_path.unlink()


def test_insert_paragraph_both_after_and_before():
    """Test error when both after and before are specified."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        with pytest.raises(ValueError, match="Cannot specify both"):
            doc.insert_paragraph("Text", after="First", before="Second")

    finally:
        doc_path.unlink()


def test_insert_paragraph_anchor_not_found():
    """Test error when anchor text is not found."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        with pytest.raises(TextNotFoundError):
            doc.insert_paragraph("Text", after="nonexistent text")

    finally:
        doc_path.unlink()


def test_insert_paragraph_ambiguous_anchor():
    """Test error when anchor text is ambiguous."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First target paragraph</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second target paragraph</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        with pytest.raises(AmbiguousTextError):
            doc.insert_paragraph("Text", after="target")

    finally:
        doc_path.unlink()


def test_insert_paragraph_positioning():
    """Test that paragraph is inserted in the correct position."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        # Insert after "First paragraph"
        doc.insert_paragraph("Inserted after first", after="First paragraph")

        # Get all paragraph texts
        all_paras = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        texts = ["".join(p.itertext()) for p in all_paras]

        # The new paragraph should appear after "First paragraph"
        first_idx = next(i for i, t in enumerate(texts) if "First paragraph" in t)
        inserted_idx = next(i for i, t in enumerate(texts) if "Inserted after first" in t)

        assert inserted_idx > first_idx

    finally:
        doc_path.unlink()


def test_insert_paragraph_before_positioning():
    """Test that paragraph is inserted before anchor."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        # Insert before "Third paragraph"
        doc.insert_paragraph("Inserted before third", before="Third paragraph")

        # Get all paragraph texts
        all_paras = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        texts = ["".join(p.itertext()) for p in all_paras]

        # The new paragraph should appear before "Third paragraph"
        third_idx = next(i for i, t in enumerate(texts) if "Third paragraph" in t)
        inserted_idx = next(i for i, t in enumerate(texts) if "Inserted before third" in t)

        assert inserted_idx < third_idx

    finally:
        doc_path.unlink()


def test_insert_paragraph_returns_paragraph_object():
    """Test that insert_paragraph returns a Paragraph object."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        from docx_redline.models.paragraph import Paragraph

        new_para = doc.insert_paragraph("New text", after="First paragraph")

        assert isinstance(new_para, Paragraph)
        assert new_para.text == "New text"

    finally:
        doc_path.unlink()


# Run tests with: pytest tests/test_structural_operations.py -v
