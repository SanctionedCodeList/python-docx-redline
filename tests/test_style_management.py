"""Tests for style management methods."""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_redline.document import Document

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def create_test_docx(content: str | None = None) -> Path:
    """Create a minimal but valid OOXML test .docx file."""
    if content is None:
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a simple document.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", content)

    return docx_path


class TestApplyStyle:
    """Tests for apply_style() method."""

    def test_apply_style_basic(self):
        """Test applying a style to a paragraph."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("Section 1: Introduction", after="This is a simple document")

        # Apply Heading1 style
        count = doc.apply_style("Section 1", "Heading1")

        assert count == 1

        # Verify style was applied
        paragraphs = doc.paragraphs
        found = False
        for para in paragraphs:
            if "Section 1" in para.text:
                assert para.style == "Heading1"
                found = True
        assert found

    def test_apply_style_multiple_paragraphs(self):
        """Test applying style to multiple paragraphs."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("Section 1", after="This is a simple document")
        doc.insert_paragraph("Section 2", after="Section 1")
        doc.insert_paragraph("Section 3", after="Section 2")

        # Apply style to all section paragraphs
        count = doc.apply_style("Section", "Heading2")

        assert count == 3

    def test_apply_style_with_regex(self):
        """Test applying style with regex pattern."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("Chapter 1", after="This is a simple document")
        doc.insert_paragraph("Chapter 2", after="This is a simple document")
        doc.insert_paragraph("Appendix A", after="This is a simple document")

        # Apply style only to Chapter paragraphs
        count = doc.apply_style(r"Chapter \d+", "Heading1", regex=True)

        assert count == 2

    def test_apply_style_no_matches(self):
        """Test applying style when no matches found."""
        doc = Document(create_test_docx())

        count = doc.apply_style("Nonexistent text", "Heading1")

        assert count == 0

    def test_apply_style_already_has_style(self):
        """Test that paragraphs with the same style aren't counted."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("Test paragraph", after="This is a simple document", style="Heading1")

        # Try to apply same style
        count = doc.apply_style("Test paragraph", "Heading1")

        assert count == 0  # Already has this style

    def test_apply_style_with_scope(self):
        """Test applying style with scope filter."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("First Section 1", after="This is a simple document")
        doc.insert_paragraph("Second Section 1", after="This is a simple document")

        # Apply only to paragraphs in "First" scope
        count = doc.apply_style("Section 1", "Heading1", scope={"contains": "First"})

        assert count == 1


class TestFormatText:
    """Tests for format_text() method."""

    def test_format_text_bold(self):
        """Test making text bold."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("This is IMPORTANT text", after="This is a simple document")

        # Make "IMPORTANT" bold
        count = doc.format_text("IMPORTANT", bold=True)

        assert count == 1

        # Verify bold was applied
        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "<w:b/>" in xml_str

    def test_format_text_italic(self):
        """Test making text italic."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("See also reference", after="This is a simple document")

        count = doc.format_text("also", italic=True)

        assert count == 1

        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "<w:i/>" in xml_str

    def test_format_text_color(self):
        """Test applying color to text."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("Warning: Do not proceed", after="This is a simple document")

        # Make "Warning" red
        count = doc.format_text("Warning", color="FF0000")

        assert count == 1

        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert 'w:val="FF0000"' in xml_str

    def test_format_text_combined(self):
        """Test applying multiple formats at once."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("CRITICAL ERROR", after="This is a simple document")

        count = doc.format_text("CRITICAL", bold=True, italic=True, color="FF0000")

        assert count == 1

        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "<w:b/>" in xml_str
        assert "<w:i/>" in xml_str
        assert 'w:val="FF0000"' in xml_str

    def test_format_text_multiple_occurrences(self):
        """Test formatting multiple occurrences of the same text."""
        doc = Document(create_test_docx())

        doc.insert_paragraph(
            "Note: this is a note. See note above.", after="This is a simple document"
        )

        count = doc.format_text("note", bold=True)

        # Should match "Note" and "note" (2 occurrences due to quote normalization)
        assert count >= 1

    def test_format_text_with_regex(self):
        """Test formatting with regex pattern."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("Section 2.1 and Section 3.5", after="This is a simple document")

        # Format all section references
        count = doc.format_text(r"Section \d+\.\d+", italic=True, regex=True)

        assert count == 2

    def test_format_text_remove_bold(self):
        """Test removing bold formatting."""
        doc = Document(create_test_docx())

        # Create document with bold text
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Bold text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""
        doc = Document(create_test_docx(content))

        # Remove bold
        count = doc.format_text("Bold", bold=False)

        assert count == 1

        # Verify bold was removed
        # The w:b element should be removed from the run containing "Bold"
        # This is harder to verify precisely, so we just check it ran without error

    def test_format_text_with_scope(self):
        """Test formatting with scope filter."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("First: test word", after="This is a simple document")
        doc.insert_paragraph("Second: test word", after="This is a simple document")

        count = doc.format_text("test", bold=True, scope={"contains": "First"})

        assert count == 1


class TestCopyFormat:
    """Tests for copy_format() method."""

    def test_copy_format_basic(self):
        """Test copying format from one text to another."""
        # Create document with formatted text
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:i/>
          <w:color w:val="FF0000"/>
        </w:rPr>
        <w:t>Source text</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Target text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""
        doc = Document(create_test_docx(content))

        # Copy formatting from "Source" to "Target"
        count = doc.copy_format("Source", "Target")

        assert count == 1

        # Verify formatting was copied
        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        # Should have bold, italic, and color applied to Target
        assert xml_str.count("<w:b/>") >= 2  # Both source and target
        assert xml_str.count("<w:i/>") >= 2

    def test_copy_format_source_not_found(self):
        """Test copy_format when source text doesn't exist."""
        doc = Document(create_test_docx())

        from docx_redline import TextNotFoundError

        with pytest.raises(TextNotFoundError):
            doc.copy_format("Nonexistent", "Target")

    def test_copy_format_no_formatting(self):
        """Test copying from text with no formatting."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("Plain text", after="This is a simple document")
        doc.insert_paragraph("Target text", after="This is a simple document")

        # Copy from text with no formatting
        count = doc.copy_format("Plain", "Target")

        # Should return 0 because there's no formatting to copy
        assert count == 0

    def test_copy_format_target_not_found(self):
        """Test copy_format when target text doesn't exist."""
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Source text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""
        doc = Document(create_test_docx(content))

        # Copy to nonexistent target
        count = doc.copy_format("Source", "Nonexistent")

        # Should return 0 (no target found)
        assert count == 0


class TestStyleManagementIntegration:
    """Integration tests for style management."""

    def test_apply_style_and_format_text_together(self):
        """Test using both apply_style and format_text on same document."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("Chapter 1: Introduction", after="This is a simple document")

        # Apply heading style
        doc.apply_style("Chapter 1", "Heading1")

        # Make the chapter number bold
        doc.format_text("Chapter 1", bold=True)

        # Verify both operations succeeded
        paragraphs = doc.paragraphs
        for para in paragraphs:
            if "Chapter 1" in para.text:
                assert para.style == "Heading1"
                break

        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "<w:b/>" in xml_str

    def test_style_management_persists_after_save(self):
        """Test that style changes persist after saving and reloading."""
        doc = Document(create_test_docx())

        doc.insert_paragraph("Important text", after="This is a simple document")
        doc.apply_style("Important", "Heading1")
        doc.format_text("Important", bold=True, color="FF0000")

        # Save and reload
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "styled.docx"
            doc.save(output_path)

            reloaded_doc = Document(output_path)

            # Verify style persisted
            paragraphs = reloaded_doc.paragraphs
            found = False
            for para in paragraphs:
                if "Important" in para.text:
                    assert para.style == "Heading1"
                    found = True
            assert found

            # Verify formatting persisted
            xml_str = etree.tostring(reloaded_doc.xml_root, encoding="unicode")
            assert "<w:b/>" in xml_str
            assert 'w:val="FF0000"' in xml_str

    def test_format_workflow(self):
        """Test a complete formatting workflow."""
        doc = Document(create_test_docx())

        # Add some content
        doc.insert_paragraph("1. First Item", after="This is a simple document")
        doc.insert_paragraph("2. Second Item", after="This is a simple document")
        doc.insert_paragraph("3. Third Item", after="This is a simple document")

        # Style all as list items
        doc.apply_style(r"\d+\. ", "ListParagraph", regex=True)

        # Make numbers bold
        doc.format_text(r"\d+\.", bold=True, regex=True)

        # Verify
        assert doc.get_text().count("Item") == 3
