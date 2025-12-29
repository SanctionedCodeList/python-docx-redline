"""Integration tests for footnote/endnote style creation.

These tests verify that inserting footnotes and endnotes properly creates
required styles in the document's styles.xml file.
"""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from python_docx_redline import Document
from python_docx_redline.models.style import StyleType

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def create_minimal_docx_without_styles() -> Path:
    """Create a minimal .docx file without any footnote/endnote styles.

    This creates a document with only the bare minimum required structure,
    without FootnoteReference, FootnoteText, EndnoteReference, or EndnoteText styles.
    """
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

    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a test document with some text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Another paragraph for testing.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document)

    return docx_path


class TestFootnoteStyleCreation:
    """Tests for footnote style auto-creation."""

    def test_insert_footnote_creates_footnote_reference_style(self):
        """Test that inserting a footnote creates FootnoteReference style."""
        doc = Document(create_minimal_docx_without_styles())

        # Verify style doesn't exist yet
        assert doc.styles.get("FootnoteReference") is None

        # Insert a footnote
        doc.insert_footnote("Test footnote", at="test document")

        # Verify FootnoteReference style was created
        style = doc.styles.get("FootnoteReference")
        assert style is not None
        assert style.style_type == StyleType.CHARACTER
        assert style.run_formatting.superscript is True

    def test_insert_footnote_creates_footnote_text_style(self):
        """Test that inserting a footnote creates FootnoteText style."""
        doc = Document(create_minimal_docx_without_styles())

        # Verify style doesn't exist yet
        assert doc.styles.get("FootnoteText") is None

        # Insert a footnote
        doc.insert_footnote("Test footnote", at="test document")

        # Verify FootnoteText style was created
        style = doc.styles.get("FootnoteText")
        assert style is not None
        assert style.style_type == StyleType.PARAGRAPH
        assert style.run_formatting.font_size == 10.0

    def test_insert_footnote_creates_footnote_text_char_style(self):
        """Test that inserting a footnote creates FootnoteTextChar style."""
        doc = Document(create_minimal_docx_without_styles())

        # Verify style doesn't exist yet
        assert doc.styles.get("FootnoteTextChar") is None

        # Insert a footnote
        doc.insert_footnote("Test footnote", at="test document")

        # Verify FootnoteTextChar style was created
        style = doc.styles.get("FootnoteTextChar")
        assert style is not None
        assert style.style_type == StyleType.CHARACTER
        assert style.linked_style == "FootnoteText"

    def test_insert_footnote_creates_all_three_styles(self):
        """Test that inserting a footnote creates all required styles."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert a footnote
        doc.insert_footnote("Test footnote", at="test document")

        # Verify all three styles exist
        assert "FootnoteReference" in doc.styles
        assert "FootnoteText" in doc.styles
        assert "FootnoteTextChar" in doc.styles


class TestEndnoteStyleCreation:
    """Tests for endnote style auto-creation."""

    def test_insert_endnote_creates_endnote_reference_style(self):
        """Test that inserting an endnote creates EndnoteReference style."""
        doc = Document(create_minimal_docx_without_styles())

        # Verify style doesn't exist yet
        assert doc.styles.get("EndnoteReference") is None

        # Insert an endnote
        doc.insert_endnote("Test endnote", at="test document")

        # Verify EndnoteReference style was created
        style = doc.styles.get("EndnoteReference")
        assert style is not None
        assert style.style_type == StyleType.CHARACTER
        assert style.run_formatting.superscript is True

    def test_insert_endnote_creates_endnote_text_style(self):
        """Test that inserting an endnote creates EndnoteText style."""
        doc = Document(create_minimal_docx_without_styles())

        # Verify style doesn't exist yet
        assert doc.styles.get("EndnoteText") is None

        # Insert an endnote
        doc.insert_endnote("Test endnote", at="test document")

        # Verify EndnoteText style was created
        style = doc.styles.get("EndnoteText")
        assert style is not None
        assert style.style_type == StyleType.PARAGRAPH
        assert style.run_formatting.font_size == 10.0

    def test_insert_endnote_creates_endnote_text_char_style(self):
        """Test that inserting an endnote creates EndnoteTextChar style."""
        doc = Document(create_minimal_docx_without_styles())

        # Verify style doesn't exist yet
        assert doc.styles.get("EndnoteTextChar") is None

        # Insert an endnote
        doc.insert_endnote("Test endnote", at="test document")

        # Verify EndnoteTextChar style was created
        style = doc.styles.get("EndnoteTextChar")
        assert style is not None
        assert style.style_type == StyleType.CHARACTER
        assert style.linked_style == "EndnoteText"

    def test_insert_endnote_creates_all_three_styles(self):
        """Test that inserting an endnote creates all required styles."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert an endnote
        doc.insert_endnote("Test endnote", at="test document")

        # Verify all three styles exist
        assert "EndnoteReference" in doc.styles
        assert "EndnoteText" in doc.styles
        assert "EndnoteTextChar" in doc.styles


class TestStylePersistence:
    """Tests for style persistence after save/reload."""

    def test_footnote_styles_persist_after_save(self):
        """Test that footnote styles persist after save and reload."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert a footnote to trigger style creation
        doc.insert_footnote("Test footnote", at="test document")

        # Save and reload
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "footnote_styles.docx"
            doc.save(output_path)

            # Reload the document
            reloaded_doc = Document(output_path)

            # Verify all styles still exist
            assert "FootnoteReference" in reloaded_doc.styles
            assert "FootnoteText" in reloaded_doc.styles
            assert "FootnoteTextChar" in reloaded_doc.styles

            # Verify style properties are preserved
            fn_ref = reloaded_doc.styles.get("FootnoteReference")
            assert fn_ref is not None
            assert fn_ref.run_formatting.superscript is True

            fn_text = reloaded_doc.styles.get("FootnoteText")
            assert fn_text is not None
            assert fn_text.run_formatting.font_size == 10.0

    def test_endnote_styles_persist_after_save(self):
        """Test that endnote styles persist after save and reload."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert an endnote to trigger style creation
        doc.insert_endnote("Test endnote", at="test document")

        # Save and reload
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "endnote_styles.docx"
            doc.save(output_path)

            # Reload the document
            reloaded_doc = Document(output_path)

            # Verify all styles still exist
            assert "EndnoteReference" in reloaded_doc.styles
            assert "EndnoteText" in reloaded_doc.styles
            assert "EndnoteTextChar" in reloaded_doc.styles

            # Verify style properties are preserved
            en_ref = reloaded_doc.styles.get("EndnoteReference")
            assert en_ref is not None
            assert en_ref.run_formatting.superscript is True

            en_text = reloaded_doc.styles.get("EndnoteText")
            assert en_text is not None
            assert en_text.run_formatting.font_size == 10.0

    def test_both_footnote_and_endnote_styles_persist(self):
        """Test that both footnote and endnote styles persist after save."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert both footnote and endnote
        doc.insert_footnote("Test footnote", at="test document")
        doc.insert_endnote("Test endnote", at="Another paragraph")

        # Save and reload
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "both_styles.docx"
            doc.save(output_path)

            # Reload the document
            reloaded_doc = Document(output_path)

            # Verify all six styles exist
            assert "FootnoteReference" in reloaded_doc.styles
            assert "FootnoteText" in reloaded_doc.styles
            assert "FootnoteTextChar" in reloaded_doc.styles
            assert "EndnoteReference" in reloaded_doc.styles
            assert "EndnoteText" in reloaded_doc.styles
            assert "EndnoteTextChar" in reloaded_doc.styles


class TestStyleIdempotence:
    """Tests for style creation idempotence (no duplication)."""

    def test_multiple_footnotes_dont_duplicate_styles(self):
        """Test that inserting multiple footnotes doesn't duplicate styles."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert multiple footnotes
        doc.insert_footnote("First footnote", at="test document")
        doc.insert_footnote("Second footnote", at="Another paragraph")

        # Count occurrences of each style ID
        fn_ref_count = sum(1 for s in doc.styles if s.style_id == "FootnoteReference")
        fn_text_count = sum(1 for s in doc.styles if s.style_id == "FootnoteText")
        fn_char_count = sum(1 for s in doc.styles if s.style_id == "FootnoteTextChar")

        # Should have exactly one of each
        assert fn_ref_count == 1
        assert fn_text_count == 1
        assert fn_char_count == 1

    def test_multiple_endnotes_dont_duplicate_styles(self):
        """Test that inserting multiple endnotes doesn't duplicate styles."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert multiple endnotes
        doc.insert_endnote("First endnote", at="test document")
        doc.insert_endnote("Second endnote", at="Another paragraph")

        # Count occurrences of each style ID
        en_ref_count = sum(1 for s in doc.styles if s.style_id == "EndnoteReference")
        en_text_count = sum(1 for s in doc.styles if s.style_id == "EndnoteText")
        en_char_count = sum(1 for s in doc.styles if s.style_id == "EndnoteTextChar")

        # Should have exactly one of each
        assert en_ref_count == 1
        assert en_text_count == 1
        assert en_char_count == 1

    def test_existing_styles_are_reused(self):
        """Test that existing styles are reused, not recreated."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert first footnote to create styles
        doc.insert_footnote("First footnote", at="test document")

        # Get references to the style objects
        original_fn_ref = doc.styles.get("FootnoteReference")
        assert original_fn_ref is not None

        # Insert second footnote
        doc.insert_footnote("Second footnote", at="Another paragraph")

        # Style should be the same object (or at least same content)
        current_fn_ref = doc.styles.get("FootnoteReference")
        assert current_fn_ref is not None
        assert current_fn_ref.style_id == original_fn_ref.style_id
        assert current_fn_ref.name == original_fn_ref.name


class TestStyleXmlStructure:
    """Tests for correct XML structure in styles.xml."""

    def test_footnote_reference_xml_has_superscript(self):
        """Test that FootnoteReference style XML has w:vertAlign superscript."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert a footnote to create styles
        doc.insert_footnote("Test footnote", at="test document")

        # Save the document
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "styles_test.docx"
            doc.save(output_path)

            # Extract and parse styles.xml
            with zipfile.ZipFile(output_path, "r") as docx:
                styles_xml = docx.read("word/styles.xml").decode("utf-8")

            # Parse and verify
            root = etree.fromstring(styles_xml.encode("utf-8"))
            nsmap = {"w": WORD_NAMESPACE}

            # Find FootnoteReference style
            fn_ref_style = root.find(".//w:style[@w:styleId='FootnoteReference']", namespaces=nsmap)
            assert fn_ref_style is not None

            # Verify it has vertAlign=superscript
            vert_align = fn_ref_style.find(".//w:vertAlign", namespaces=nsmap)
            assert vert_align is not None
            assert vert_align.get(f"{{{WORD_NAMESPACE}}}val") == "superscript"

    def test_endnote_reference_xml_has_superscript(self):
        """Test that EndnoteReference style XML has w:vertAlign superscript."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert an endnote to create styles
        doc.insert_endnote("Test endnote", at="test document")

        # Save the document
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "styles_test.docx"
            doc.save(output_path)

            # Extract and parse styles.xml
            with zipfile.ZipFile(output_path, "r") as docx:
                styles_xml = docx.read("word/styles.xml").decode("utf-8")

            # Parse and verify
            root = etree.fromstring(styles_xml.encode("utf-8"))
            nsmap = {"w": WORD_NAMESPACE}

            # Find EndnoteReference style
            en_ref_style = root.find(".//w:style[@w:styleId='EndnoteReference']", namespaces=nsmap)
            assert en_ref_style is not None

            # Verify it has vertAlign=superscript
            vert_align = en_ref_style.find(".//w:vertAlign", namespaces=nsmap)
            assert vert_align is not None
            assert vert_align.get(f"{{{WORD_NAMESPACE}}}val") == "superscript"

    def test_footnote_text_style_xml_structure(self):
        """Test that FootnoteText style XML has correct structure."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert a footnote to create styles
        doc.insert_footnote("Test footnote", at="test document")

        # Save the document
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "styles_test.docx"
            doc.save(output_path)

            # Extract and parse styles.xml
            with zipfile.ZipFile(output_path, "r") as docx:
                styles_xml = docx.read("word/styles.xml").decode("utf-8")

            # Parse and verify
            root = etree.fromstring(styles_xml.encode("utf-8"))
            nsmap = {"w": WORD_NAMESPACE}

            # Find FootnoteText style
            fn_text_style = root.find(".//w:style[@w:styleId='FootnoteText']", namespaces=nsmap)
            assert fn_text_style is not None

            # Verify it's a paragraph style
            assert fn_text_style.get(f"{{{WORD_NAMESPACE}}}type") == "paragraph"

            # Verify it has w:name element
            name_elem = fn_text_style.find(".//w:name", namespaces=nsmap)
            assert name_elem is not None
            assert name_elem.get(f"{{{WORD_NAMESPACE}}}val") == "footnote text"

    def test_styles_xml_contains_all_required_styles(self):
        """Test that styles.xml contains all required styles after save."""
        doc = Document(create_minimal_docx_without_styles())

        # Insert both footnote and endnote
        doc.insert_footnote("Test footnote", at="test document")
        doc.insert_endnote("Test endnote", at="Another paragraph")

        # Save the document
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "all_styles.docx"
            doc.save(output_path)

            # Extract and parse styles.xml
            with zipfile.ZipFile(output_path, "r") as docx:
                styles_xml = docx.read("word/styles.xml").decode("utf-8")

            # Verify all style IDs are present in the XML
            assert "FootnoteReference" in styles_xml
            assert "FootnoteText" in styles_xml
            assert "FootnoteTextChar" in styles_xml
            assert "EndnoteReference" in styles_xml
            assert "EndnoteText" in styles_xml
            assert "EndnoteTextChar" in styles_xml


class TestStyleAttributes:
    """Tests for specific style attributes."""

    def test_footnote_reference_based_on_default_paragraph_font(self):
        """Test that FootnoteReference is based on DefaultParagraphFont."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_footnote("Test footnote", at="test document")

        style = doc.styles.get("FootnoteReference")
        assert style is not None
        assert style.based_on == "DefaultParagraphFont"

    def test_footnote_text_based_on_normal(self):
        """Test that FootnoteText is based on Normal style."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_footnote("Test footnote", at="test document")

        style = doc.styles.get("FootnoteText")
        assert style is not None
        assert style.based_on == "Normal"

    def test_footnote_text_linked_to_footnote_text_char(self):
        """Test that FootnoteText is linked to FootnoteTextChar."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_footnote("Test footnote", at="test document")

        fn_text = doc.styles.get("FootnoteText")
        fn_text_char = doc.styles.get("FootnoteTextChar")

        assert fn_text is not None
        assert fn_text_char is not None
        assert fn_text.linked_style == "FootnoteTextChar"
        assert fn_text_char.linked_style == "FootnoteText"

    def test_endnote_reference_based_on_default_paragraph_font(self):
        """Test that EndnoteReference is based on DefaultParagraphFont."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_endnote("Test endnote", at="test document")

        style = doc.styles.get("EndnoteReference")
        assert style is not None
        assert style.based_on == "DefaultParagraphFont"

    def test_endnote_text_based_on_normal(self):
        """Test that EndnoteText is based on Normal style."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_endnote("Test endnote", at="test document")

        style = doc.styles.get("EndnoteText")
        assert style is not None
        assert style.based_on == "Normal"

    def test_endnote_text_linked_to_endnote_text_char(self):
        """Test that EndnoteText is linked to EndnoteTextChar."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_endnote("Test endnote", at="test document")

        en_text = doc.styles.get("EndnoteText")
        en_text_char = doc.styles.get("EndnoteTextChar")

        assert en_text is not None
        assert en_text_char is not None
        assert en_text.linked_style == "EndnoteTextChar"
        assert en_text_char.linked_style == "EndnoteText"


class TestStyleUIProperties:
    """Tests for style UI properties (hidden, unhideWhenUsed, etc.)."""

    def test_footnote_reference_is_semi_hidden(self):
        """Test that FootnoteReference style is semi-hidden."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_footnote("Test footnote", at="test document")

        style = doc.styles.get("FootnoteReference")
        assert style is not None
        assert style.semi_hidden is True

    def test_footnote_reference_unhide_when_used(self):
        """Test that FootnoteReference has unhideWhenUsed property."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_footnote("Test footnote", at="test document")

        style = doc.styles.get("FootnoteReference")
        assert style is not None
        assert style.unhide_when_used is True

    def test_footnote_text_ui_priority(self):
        """Test that FootnoteText has ui_priority set."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_footnote("Test footnote", at="test document")

        style = doc.styles.get("FootnoteText")
        assert style is not None
        assert style.ui_priority == 99

    def test_endnote_reference_is_semi_hidden(self):
        """Test that EndnoteReference style is semi-hidden."""
        doc = Document(create_minimal_docx_without_styles())
        doc.insert_endnote("Test endnote", at="test document")

        style = doc.styles.get("EndnoteReference")
        assert style is not None
        assert style.semi_hidden is True
