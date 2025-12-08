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


# Tests for insert_paragraphs()


def test_insert_paragraphs_basic():
    """Test inserting multiple paragraphs at once."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        texts = ["Para 1", "Para 2", "Para 3"]
        paras = doc.insert_paragraphs(texts, after="First paragraph")

        assert len(paras) == 3
        assert paras[0].text == "Para 1"
        assert paras[1].text == "Para 2"
        assert paras[2].text == "Para 3"

    finally:
        doc_path.unlink()


def test_insert_paragraphs_ordering():
    """Test that paragraphs are inserted in correct order."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        texts = ["First new", "Second new", "Third new"]
        doc.insert_paragraphs(texts, after="First paragraph")

        # Get all paragraph texts
        all_paras = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        all_texts = ["".join(p.itertext()) for p in all_paras]

        # Find the inserted paragraphs
        first_idx = next(i for i, t in enumerate(all_texts) if "First new" in t)
        second_idx = next(i for i, t in enumerate(all_texts) if "Second new" in t)
        third_idx = next(i for i, t in enumerate(all_texts) if "Third new" in t)

        # They should be in order
        assert first_idx < second_idx < third_idx

    finally:
        doc_path.unlink()


def test_insert_paragraphs_with_style():
    """Test inserting multiple paragraphs with a style."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        texts = ["Heading A", "Heading B"]
        paras = doc.insert_paragraphs(texts, after="First paragraph", style="Heading2")

        assert paras[0].style == "Heading2"
        assert paras[1].style == "Heading2"

    finally:
        doc_path.unlink()


def test_insert_paragraphs_empty_list():
    """Test inserting empty list returns empty list."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        paras = doc.insert_paragraphs([], after="First paragraph")

        assert paras == []

    finally:
        doc_path.unlink()


def test_insert_paragraphs_single():
    """Test inserting single paragraph via insert_paragraphs."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        paras = doc.insert_paragraphs(["Single para"], after="First paragraph")

        assert len(paras) == 1
        assert paras[0].text == "Single para"

    finally:
        doc_path.unlink()


def test_insert_paragraphs_before():
    """Test inserting paragraphs before anchor."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        texts = ["Before 1", "Before 2"]
        paras = doc.insert_paragraphs(texts, before="Third paragraph")

        assert len(paras) == 2

        # Verify ordering
        all_paras = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        all_texts = ["".join(p.itertext()) for p in all_paras]

        third_idx = next(i for i, t in enumerate(all_texts) if "Third paragraph" in t)
        before1_idx = next(i for i, t in enumerate(all_texts) if "Before 1" in t)
        before2_idx = next(i for i, t in enumerate(all_texts) if "Before 2" in t)

        assert before1_idx < before2_idx < third_idx

    finally:
        doc_path.unlink()


def test_insert_paragraphs_untracked():
    """Test inserting untracked paragraphs."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        texts = ["Untracked 1", "Untracked 2"]
        paras = doc.insert_paragraphs(texts, after="First paragraph", track=False)

        assert len(paras) == 2
        assert paras[0].text == "Untracked 1"
        assert paras[1].text == "Untracked 2"

    finally:
        doc_path.unlink()


def test_insert_paragraphs_returns_paragraph_objects():
    """Test that insert_paragraphs returns Paragraph objects."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        from docx_redline.models.paragraph import Paragraph

        texts = ["Para A", "Para B"]
        paras = doc.insert_paragraphs(texts, after="First paragraph")

        assert all(isinstance(p, Paragraph) for p in paras)

    finally:
        doc_path.unlink()


# Tests for delete_section()


def test_delete_section_basic_tracked():
    """Test deleting a section with tracked changes."""
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
        <w:t>Intro content</w:t>
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

    try:
        doc = Document(doc_path)

        deleted_section = doc.delete_section("Methods", track=True)

        assert deleted_section is not None
        assert deleted_section.heading_text.strip() == "Methods"
        assert len(deleted_section.paragraphs) == 2

        # Check that w:del elements were created
        xml_string = etree.tostring(doc.xml_root, encoding="unicode")
        assert "w:del" in xml_string

    finally:
        doc_path.unlink()


def test_delete_section_untracked():
    """Test deleting a section without tracked changes."""
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
        <w:t>Content one</w:t>
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
        <w:t>Content two</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        # Count paragraphs before deletion
        paras_before = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        assert len(paras_before) == 4

        deleted_section = doc.delete_section("Section One", track=False)

        assert deleted_section is not None
        assert deleted_section.heading_text.strip() == "Section One"

        # Paragraphs should be completely removed (not wrapped in w:del)
        paras_after = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        assert len(paras_after) == 2  # Only Section Two and its content remain

        # No w:del elements should exist
        xml_string = etree.tostring(doc.xml_root, encoding="unicode")
        assert "w:del" not in xml_string

    finally:
        doc_path.unlink()


def test_delete_section_multiple_paragraphs():
    """Test deleting section with multiple content paragraphs."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Results</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Result paragraph 1</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Result paragraph 2</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Result paragraph 3</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        deleted_section = doc.delete_section("Results", track=True)

        assert len(deleted_section.paragraphs) == 4  # Heading + 3 content paragraphs

    finally:
        doc_path.unlink()


def test_delete_section_not_found():
    """Test error when section heading not found."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        with pytest.raises(TextNotFoundError):
            doc.delete_section("Nonexistent Section")

    finally:
        doc_path.unlink()


def test_delete_section_ambiguous():
    """Test error when multiple sections match."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Results Summary</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content 1</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Results Analysis</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content 2</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        # Both sections contain "Results"
        with pytest.raises(AmbiguousTextError):
            doc.delete_section("Results")

    finally:
        doc_path.unlink()


def test_delete_section_case_insensitive():
    """Test that section matching is case insensitive."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Discussion</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Discussion content</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        # Should match despite different case
        deleted_section = doc.delete_section("discussion", track=True)

        assert deleted_section.heading_text.strip() == "Discussion"

    finally:
        doc_path.unlink()


def test_delete_section_with_scope():
    """Test deleting section with scope filter."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Chapter One</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content in chapter one</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Summary</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Summary of chapter one</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Chapter Two</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content in chapter two</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Summary</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Summary of chapter two</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        # Delete only the first Summary section (in chapter one)
        deleted_section = doc.delete_section("Summary", track=True, scope="chapter one")

        assert deleted_section.heading_text.strip() == "Summary"

        # Verify the second Summary (in chapter two) still exists
        remaining_paras = [p for p in doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p")]
        texts = ["".join(p.itertext()) for p in remaining_paras]

        # Should still have "Summary of chapter two"
        assert any("chapter two" in text.lower() for text in texts)

    finally:
        doc_path.unlink()


def test_delete_section_returns_section_object():
    """Test that delete_section returns a Section object."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Conclusion</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Final thoughts</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        from docx_redline.models.section import Section

        deleted_section = doc.delete_section("Conclusion", track=True)

        assert isinstance(deleted_section, Section)
        assert deleted_section.heading_text.strip() == "Conclusion"
        assert len(deleted_section) == 2  # Heading + 1 content paragraph

    finally:
        doc_path.unlink()


# Run tests with: pytest tests/test_structural_operations.py -v
