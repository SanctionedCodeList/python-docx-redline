"""
Tests for format tracking functionality.

These tests verify the format_tracked() and format_paragraph_tracked() methods,
which apply formatting changes with proper OOXML revision tracking.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest

from python_docx_redline import Document, TextNotFoundError

# Word namespace
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    """Create a fully qualified Word namespace tag."""
    return f"{{{WORD_NS}}}{tag}"


# Minimal Word document XML structure
MINIMAL_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a test document with IMPORTANT text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Section 2.1: Introduction</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Some content here with a note.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

# Document with existing formatting
FORMATTED_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:i/>
        </w:rPr>
        <w:t>This italic text needs bold.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:jc w:val="left"/>
      </w:pPr>
      <w:r>
        <w:t>Left aligned paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


def create_test_docx(content: str = MINIMAL_DOCUMENT_XML) -> Path:
    """Create a minimal but valid OOXML test .docx file."""
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


class TestFormatTracked:
    """Tests for format_tracked() method."""

    def test_format_tracked_bold(self):
        """Test applying bold formatting with tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path, author="Test Author")
            result = doc.format_tracked("IMPORTANT", bold=True)

            assert result.success
            assert result.text_matched == "IMPORTANT"
            assert result.changes_applied == {"bold": True}
            assert result.runs_affected >= 1
            assert result.change_id > 0

            # Verify XML structure
            rpr_changes = list(doc.xml_root.iter(_w("rPrChange")))
            assert len(rpr_changes) == 1

            # Verify the change has proper attributes
            change = rpr_changes[0]
            assert change.get(_w("id")) is not None
            assert change.get(_w("author")) == "Test Author"
            assert change.get(_w("date")) is not None

            # Verify the parent rPr has bold
            parent_rpr = change.getparent()
            assert parent_rpr.find(_w("b")) is not None
        finally:
            docx_path.unlink()

    def test_format_tracked_multiple_properties(self):
        """Test applying multiple formatting properties."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            result = doc.format_tracked(
                "IMPORTANT",
                bold=True,
                italic=True,
                color="#FF0000",
            )

            assert result.success
            assert result.changes_applied == {
                "bold": True,
                "italic": True,
                "color": "#FF0000",
            }

            # Verify XML has all properties
            rpr_change = list(doc.xml_root.iter(_w("rPrChange")))[0]
            parent_rpr = rpr_change.getparent()
            assert parent_rpr.find(_w("b")) is not None
            assert parent_rpr.find(_w("i")) is not None
            assert parent_rpr.find(_w("color")) is not None
        finally:
            docx_path.unlink()

    def test_format_tracked_preserve_existing(self):
        """Test that existing formatting is preserved when adding new."""
        docx_path = create_test_docx(FORMATTED_DOCUMENT_XML)
        try:
            doc = Document(docx_path)
            result = doc.format_tracked("italic text", bold=True)

            assert result.success

            # Verify both italic (existing) and bold (new) are present
            rpr_change = list(doc.xml_root.iter(_w("rPrChange")))[0]
            parent_rpr = rpr_change.getparent()
            assert parent_rpr.find(_w("b")) is not None
            assert parent_rpr.find(_w("i")) is not None

            # Verify previous state only had italic
            prev_rpr = rpr_change.find(_w("rPr"))
            assert prev_rpr.find(_w("i")) is not None
            assert prev_rpr.find(_w("b")) is None
        finally:
            docx_path.unlink()

    def test_format_tracked_text_not_found(self):
        """Test that TextNotFoundError is raised for missing text."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            with pytest.raises(TextNotFoundError):
                doc.format_tracked("nonexistent text", bold=True)
        finally:
            docx_path.unlink()

    def test_format_tracked_no_formatting(self):
        """Test that ValueError is raised when no formatting specified."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            with pytest.raises(ValueError, match="formatting property"):
                doc.format_tracked("IMPORTANT")
        finally:
            docx_path.unlink()

    def test_format_tracked_font_size(self):
        """Test applying font size with tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            result = doc.format_tracked("IMPORTANT", font_size=14)

            assert result.success

            # Verify sz element with correct half-point value (14pt = 28 half-points)
            rpr_change = list(doc.xml_root.iter(_w("rPrChange")))[0]
            parent_rpr = rpr_change.getparent()
            sz = parent_rpr.find(_w("sz"))
            assert sz is not None
            assert sz.get(_w("val")) == "28"
        finally:
            docx_path.unlink()

    def test_format_tracked_font_name(self):
        """Test applying font name with tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            result = doc.format_tracked("IMPORTANT", font_name="Arial")

            assert result.success

            rpr_change = list(doc.xml_root.iter(_w("rPrChange")))[0]
            parent_rpr = rpr_change.getparent()
            rfonts = parent_rpr.find(_w("rFonts"))
            assert rfonts is not None
            assert rfonts.get(_w("ascii")) == "Arial"
        finally:
            docx_path.unlink()

    def test_format_tracked_occurrence_all(self):
        """Test formatting all occurrences."""
        # Create document with duplicate text
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>First note here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second note here.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            result = doc.format_tracked("note", bold=True, occurrence="all")

            assert result.success
            assert result.runs_affected == 2

            # Verify two rPrChange elements
            rpr_changes = list(doc.xml_root.iter(_w("rPrChange")))
            assert len(rpr_changes) == 2
        finally:
            docx_path.unlink()

    def test_format_tracked_no_change_when_already_formatted(self):
        """Test no change created when formatting already matches."""
        # Create document with already bold text
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Already bold text.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            result = doc.format_tracked("Already bold", bold=True)

            # Should succeed but with no runs affected (no actual change)
            assert result.runs_affected == 0
        finally:
            docx_path.unlink()


class TestFormatParagraphTracked:
    """Tests for format_paragraph_tracked() method."""

    def test_format_paragraph_alignment(self):
        """Test applying paragraph alignment with tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            result = doc.format_paragraph_tracked(
                containing="test document",
                alignment="center",
            )

            assert result.success
            assert result.changes_applied == {"alignment": "center"}
            assert result.change_id > 0

            # Verify XML structure
            ppr_changes = list(doc.xml_root.iter(_w("pPrChange")))
            assert len(ppr_changes) == 1

            # Verify the parent pPr has center alignment
            change = ppr_changes[0]
            parent_ppr = change.getparent()
            jc = parent_ppr.find(_w("jc"))
            assert jc is not None
            assert jc.get(_w("val")) == "center"
        finally:
            docx_path.unlink()

    def test_format_paragraph_spacing(self):
        """Test applying paragraph spacing with tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            result = doc.format_paragraph_tracked(
                containing="test document",
                spacing_before=12,
                spacing_after=6,
            )

            assert result.success

            ppr_change = list(doc.xml_root.iter(_w("pPrChange")))[0]
            parent_ppr = ppr_change.getparent()
            spacing = parent_ppr.find(_w("spacing"))
            assert spacing is not None
            # 12pt = 240 twips, 6pt = 120 twips
            assert spacing.get(_w("before")) == "240"
            assert spacing.get(_w("after")) == "120"
        finally:
            docx_path.unlink()

    def test_format_paragraph_by_index(self):
        """Test formatting paragraph by index."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            result = doc.format_paragraph_tracked(
                index=1,  # Second paragraph (0-indexed)
                alignment="right",
            )

            assert result.success
            assert "Section 2.1" in result.text_matched
        finally:
            docx_path.unlink()

    def test_format_paragraph_preserve_existing(self):
        """Test that existing paragraph formatting is preserved."""
        docx_path = create_test_docx(FORMATTED_DOCUMENT_XML)
        try:
            doc = Document(docx_path)
            result = doc.format_paragraph_tracked(
                containing="Left aligned",
                spacing_after=12,
            )

            assert result.success

            # Verify previous state had left alignment
            ppr_change = list(doc.xml_root.iter(_w("pPrChange")))[0]
            prev_ppr = ppr_change.find(_w("pPr"))
            prev_jc = prev_ppr.find(_w("jc"))
            assert prev_jc is not None
            assert prev_jc.get(_w("val")) == "left"
        finally:
            docx_path.unlink()

    def test_format_paragraph_not_found(self):
        """Test error when paragraph not found."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            with pytest.raises(TextNotFoundError):
                doc.format_paragraph_tracked(
                    containing="nonexistent paragraph",
                    alignment="center",
                )
        finally:
            docx_path.unlink()

    def test_format_paragraph_no_target(self):
        """Test error when no targeting parameter provided."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            with pytest.raises(ValueError, match="targeting parameter"):
                doc.format_paragraph_tracked(alignment="center")
        finally:
            docx_path.unlink()

    def test_format_paragraph_no_formatting(self):
        """Test error when no formatting parameter provided."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            with pytest.raises(ValueError, match="formatting property"):
                doc.format_paragraph_tracked(containing="test")
        finally:
            docx_path.unlink()


class TestAcceptRejectFormatChanges:
    """Tests for accepting and rejecting format changes."""

    def test_accept_format_changes(self):
        """Test accepting format changes removes rPrChange elements."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Apply formatting
            doc.format_tracked("IMPORTANT", bold=True)
            assert len(list(doc.xml_root.iter(_w("rPrChange")))) == 1

            # Accept changes
            count = doc.accept_format_changes()
            assert count == 1

            # Verify rPrChange removed but bold remains
            assert len(list(doc.xml_root.iter(_w("rPrChange")))) == 0
            # Find the run with bold
            for r in doc.xml_root.iter(_w("r")):
                rpr = r.find(_w("rPr"))
                if rpr is not None and rpr.find(_w("b")) is not None:
                    break
            else:
                pytest.fail("Bold formatting should be preserved after accept")
        finally:
            docx_path.unlink()

    def test_reject_format_changes(self):
        """Test rejecting format changes restores previous formatting."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Apply formatting
            doc.format_tracked("IMPORTANT", bold=True)

            # Reject changes
            count = doc.reject_format_changes()
            assert count == 1

            # Verify rPrChange removed and bold removed
            assert len(list(doc.xml_root.iter(_w("rPrChange")))) == 0

            # Verify no bold in the document (reverted)
            for r in doc.xml_root.iter(_w("r")):
                rpr = r.find(_w("rPr"))
                if rpr is not None:
                    assert rpr.find(_w("b")) is None
        finally:
            docx_path.unlink()

    def test_accept_all_changes_includes_format(self):
        """Test accept_all_changes() handles format changes."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Apply text and format changes
            doc.insert_tracked(" [ADDED]", after="IMPORTANT")
            doc.format_tracked("test document", italic=True)

            # Accept all
            doc.accept_all_changes()

            # Verify no tracked changes remain
            assert len(list(doc.xml_root.iter(_w("ins")))) == 0
            assert len(list(doc.xml_root.iter(_w("rPrChange")))) == 0
        finally:
            docx_path.unlink()

    def test_reject_all_changes_includes_format(self):
        """Test reject_all_changes() handles format changes."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Apply format change
            doc.format_tracked("IMPORTANT", bold=True)
            doc.format_paragraph_tracked(containing="test", alignment="center")

            # Reject all
            doc.reject_all_changes()

            # Verify no tracked changes remain
            assert len(list(doc.xml_root.iter(_w("rPrChange")))) == 0
            assert len(list(doc.xml_root.iter(_w("pPrChange")))) == 0
        finally:
            docx_path.unlink()


class TestBatchFormatOperations:
    """Tests for batch format operations via apply_edits()."""

    def test_apply_edits_format_tracked(self):
        """Test format_tracked via apply_edits()."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            edits = [
                {
                    "type": "format_tracked",
                    "text": "IMPORTANT",
                    "bold": True,
                    "color": "#FF0000",
                }
            ]
            results = doc.apply_edits(edits)

            assert len(results) == 1
            assert results[0].success
            assert results[0].edit_type == "format_tracked"
        finally:
            docx_path.unlink()

    def test_apply_edits_format_paragraph_tracked(self):
        """Test format_paragraph_tracked via apply_edits()."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            edits = [
                {
                    "type": "format_paragraph_tracked",
                    "containing": "test document",
                    "alignment": "center",
                }
            ]
            results = doc.apply_edits(edits)

            assert len(results) == 1
            assert results[0].success
            assert results[0].edit_type == "format_paragraph_tracked"
        finally:
            docx_path.unlink()

    def test_apply_edits_format_missing_text(self):
        """Test format_tracked error when text missing."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            edits = [
                {
                    "type": "format_tracked",
                    "bold": True,
                    # Missing "text" parameter
                }
            ]
            results = doc.apply_edits(edits)

            assert len(results) == 1
            assert not results[0].success
            assert "Missing required parameter" in results[0].message
        finally:
            docx_path.unlink()

    def test_apply_edits_format_missing_formatting(self):
        """Test format_tracked error when no formatting specified."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            edits = [
                {
                    "type": "format_tracked",
                    "text": "IMPORTANT",
                    # No formatting parameters
                }
            ]
            results = doc.apply_edits(edits)

            assert len(results) == 1
            assert not results[0].success
            assert "formatting parameter" in results[0].message
        finally:
            docx_path.unlink()


class TestFormatBuilders:
    """Tests for RunPropertyBuilder and ParagraphPropertyBuilder."""

    def test_run_property_builder_build(self):
        """Test building rPr from scratch."""
        from python_docx_redline.format_builder import RunPropertyBuilder

        rpr = RunPropertyBuilder.build(bold=True, italic=True, font_size=12)

        assert rpr.find(_w("b")) is not None
        assert rpr.find(_w("i")) is not None
        assert rpr.find(_w("sz")).get(_w("val")) == "24"  # 12pt = 24 half-points

    def test_run_property_builder_merge(self):
        """Test merging new properties into existing rPr."""
        from python_docx_redline.format_builder import RunPropertyBuilder

        # Create existing rPr with italic
        existing = RunPropertyBuilder.build(italic=True)

        # Merge bold into it
        merged = RunPropertyBuilder.merge(existing, {"bold": True})

        # Both should be present
        assert merged.find(_w("b")) is not None
        assert merged.find(_w("i")) is not None

    def test_run_property_builder_extract(self):
        """Test extracting properties from rPr."""
        from python_docx_redline.format_builder import RunPropertyBuilder

        rpr = RunPropertyBuilder.build(bold=True, font_size=14)
        extracted = RunPropertyBuilder.extract(rpr)

        assert extracted["bold"] is True
        assert extracted["font_size"] == 14.0

    def test_run_property_builder_diff(self):
        """Test diffing two rPr elements."""
        from python_docx_redline.format_builder import RunPropertyBuilder

        old = RunPropertyBuilder.build(italic=True)
        new = RunPropertyBuilder.build(italic=True, bold=True)

        diff = RunPropertyBuilder.diff(old, new)

        assert "bold" in diff
        assert diff["bold"] == (None, True)  # Was None, now True
        assert "italic" not in diff  # Unchanged

    def test_paragraph_property_builder_build(self):
        """Test building pPr from scratch."""
        from python_docx_redline.format_builder import ParagraphPropertyBuilder

        ppr = ParagraphPropertyBuilder.build(alignment="center", spacing_after=12)

        jc = ppr.find(_w("jc"))
        assert jc is not None
        assert jc.get(_w("val")) == "center"

        spacing = ppr.find(_w("spacing"))
        assert spacing is not None
        assert spacing.get(_w("after")) == "240"  # 12pt = 240 twips

    def test_paragraph_property_builder_extract(self):
        """Test extracting properties from pPr."""
        from python_docx_redline.format_builder import ParagraphPropertyBuilder

        ppr = ParagraphPropertyBuilder.build(alignment="right", indent_left=0.5)
        extracted = ParagraphPropertyBuilder.extract(ppr)

        assert extracted["alignment"] == "right"
        assert extracted["indent_left"] == 0.5


class TestTrackedXMLGeneratorFormats:
    """Tests for TrackedXMLGenerator format change methods."""

    def test_create_run_property_change(self):
        """Test creating rPrChange element."""
        from python_docx_redline.format_builder import RunPropertyBuilder
        from python_docx_redline.tracked_xml import TrackedXMLGenerator

        gen = TrackedXMLGenerator(author="Test Author")
        prev_rpr = RunPropertyBuilder.build(italic=True)

        change = gen.create_run_property_change(prev_rpr)

        assert change.tag == _w("rPrChange")
        assert change.get(_w("id")) == "1"
        assert change.get(_w("author")) == "Test Author"
        assert change.get(_w("date")) is not None

        # Verify previous rPr is included
        inner_rpr = change.find(_w("rPr"))
        assert inner_rpr is not None
        assert inner_rpr.find(_w("i")) is not None

    def test_create_paragraph_property_change(self):
        """Test creating pPrChange element."""
        from python_docx_redline.format_builder import ParagraphPropertyBuilder
        from python_docx_redline.tracked_xml import TrackedXMLGenerator

        gen = TrackedXMLGenerator(author="Test Author")
        prev_ppr = ParagraphPropertyBuilder.build(alignment="left")

        change = gen.create_paragraph_property_change(prev_ppr)

        assert change.tag == _w("pPrChange")
        assert change.get(_w("id")) == "1"
        assert change.get(_w("author")) == "Test Author"

        # Verify previous pPr is included
        inner_ppr = change.find(_w("pPr"))
        assert inner_ppr is not None
        assert inner_ppr.find(_w("jc")) is not None

    def test_change_id_sequencing(self):
        """Test that change IDs are properly sequenced."""
        from python_docx_redline.tracked_xml import TrackedXMLGenerator

        gen = TrackedXMLGenerator()

        change1 = gen.create_run_property_change(None)
        change2 = gen.create_paragraph_property_change(None)
        change3 = gen.create_run_property_change(None)

        assert change1.get(_w("id")) == "1"
        assert change2.get(_w("id")) == "2"
        assert change3.get(_w("id")) == "3"
