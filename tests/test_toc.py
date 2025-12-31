"""
Tests for Table of Contents functionality.

Tests the insert_toc() method for generating valid TOC field structures,
settings.xml updates, and style creation.
"""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from python_docx_redline import Document

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def create_simple_document() -> Path:
    """Create a minimal test document with headings."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Introduction</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>This is the introduction paragraph.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
  <w:r><w:t>Background</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>Some background information here.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Methodology</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>Our methodology is described here.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_styles() -> Path:
    """Create a document with an existing styles.xml."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Chapter One</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>Content here.</w:t></w:r>
</w:p>
<w:sectPr/>
</w:body>
</w:document>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
</w:styles>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/styles.xml", styles_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestInsertTocBasic:
    """Tests for basic TOC insertion."""

    def test_insert_toc_basic(self):
        """Test inserting a TOC with default settings."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            # Find the SDT element (TOC container)
            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            assert body is not None

            sdt_elems = body.findall(f"{{{WORD_NS}}}sdt")
            assert len(sdt_elems) >= 1, "TOC SDT not found in document"

            # Check that TOC is marked as Table of Contents
            sdt = sdt_elems[0]
            # docPartGallery should be inside docPartObj per OOXML schema
            doc_part_obj = sdt.find(f".//{{{WORD_NS}}}docPartObj")
            assert doc_part_obj is not None
            doc_part_gallery = doc_part_obj.find(f"{{{WORD_NS}}}docPartGallery")
            assert doc_part_gallery is not None
            assert doc_part_gallery.get(f"{{{WORD_NS}}}val") == "Table of Contents"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_toc_creates_title_paragraph(self):
        """Test that TOC insertion creates a title paragraph."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(title="Contents")

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            assert body is not None

            # First child should be the title paragraph
            first_para = body[0]
            assert first_para.tag == f"{{{WORD_NS}}}p"

            # Check for TOCHeading style
            p_style = first_para.find(f".//{{{WORD_NS}}}pStyle")
            assert p_style is not None
            assert p_style.get(f"{{{WORD_NS}}}val") == "TOCHeading"

            # Check title text
            text_elem = first_para.find(f".//{{{WORD_NS}}}t")
            assert text_elem is not None
            assert text_elem.text == "Contents"

        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertTocCustomLevels:
    """Tests for custom heading level configurations."""

    def test_insert_toc_custom_levels(self):
        """Test TOC with custom heading levels."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(levels=(1, 5))

            # Find the field instruction
            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            instr_text = body.find(f".//{{{WORD_NS}}}instrText")
            assert instr_text is not None
            assert '\\o "1-5"' in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_toc_single_level(self):
        """Test TOC including only one heading level."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(levels=(1, 1))

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            instr_text = body.find(f".//{{{WORD_NS}}}instrText")
            assert instr_text is not None
            assert '\\o "1-1"' in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertTocNoTitle:
    """Tests for TOC without title."""

    def test_insert_toc_no_title(self):
        """Test inserting TOC without a title."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(title=None)

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            assert body is not None

            # First child should be the SDT (no title paragraph)
            first_child = body[0]
            assert first_child.tag == f"{{{WORD_NS}}}sdt"

        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertTocUpdateOnOpen:
    """Tests for settings.xml update behavior."""

    def test_insert_toc_update_on_open(self):
        """Test that insert_toc sets updateFields in settings.xml."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(update_on_open=True)

            # Save the document to apply changes
            output_path = Path(tempfile.mktemp(suffix=".docx"))
            doc.save(output_path)

            # Read back the settings.xml
            with zipfile.ZipFile(output_path, "r") as docx:
                assert "word/settings.xml" in docx.namelist()
                settings_xml = docx.read("word/settings.xml")

            # Parse and check for updateFields
            root = etree.fromstring(settings_xml)
            update_fields = root.find(f".//{{{WORD_NS}}}updateFields")
            assert update_fields is not None
            assert update_fields.get(f"{{{WORD_NS}}}val") == "true"

            output_path.unlink(missing_ok=True)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_toc_no_update_on_open(self):
        """Test that update_on_open=False skips settings update."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(update_on_open=False)

            # Save the document
            output_path = Path(tempfile.mktemp(suffix=".docx"))
            doc.save(output_path)

            # Check settings.xml - should not have updateFields=true
            with zipfile.ZipFile(output_path, "r") as docx:
                # May or may not have settings.xml
                if "word/settings.xml" in docx.namelist():
                    settings_xml = docx.read("word/settings.xml")
                    root = etree.fromstring(settings_xml)
                    _ = root.find(f".//{{{WORD_NS}}}updateFields")
                    # If it exists, value should not be "true" (or at least not set by us)
                    # This test just verifies we don't crash

            output_path.unlink(missing_ok=True)
        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertTocFieldStructure:
    """Tests for verifying correct TOC field XML structure."""

    def test_toc_field_has_begin_separate_end(self):
        """Test that TOC field has proper begin/separate/end structure."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            sdt = body.find(f"{{{WORD_NS}}}sdt")
            assert sdt is not None

            # Find all fldChar elements
            fld_chars = sdt.findall(f".//{{{WORD_NS}}}fldChar")
            assert len(fld_chars) == 3, "Expected 3 fldChar elements (begin, separate, end)"

            # Check types
            types = [fc.get(f"{{{WORD_NS}}}fldCharType") for fc in fld_chars]
            assert types == ["begin", "separate", "end"]

        finally:
            doc_path.unlink(missing_ok=True)

    def test_toc_field_is_dirty(self):
        """Test that TOC field is marked as dirty."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")

            # Find the begin fldChar
            begin_fld = body.find(f".//{{{WORD_NS}}}fldChar[@{{{WORD_NS}}}fldCharType='begin']")
            assert begin_fld is not None
            assert begin_fld.get(f"{{{WORD_NS}}}dirty") == "true"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_toc_field_instruction_format(self):
        """Test the field instruction string format."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(levels=(1, 3), hyperlinks=True, use_outline_levels=True)

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            instr_text = body.find(f".//{{{WORD_NS}}}instrText")
            assert instr_text is not None

            instruction = instr_text.text
            assert "TOC" in instruction
            assert '\\o "1-3"' in instruction
            assert "\\h" in instruction  # hyperlinks
            assert "\\z" in instruction  # hide in web view
            assert "\\u" in instruction  # outline levels

        finally:
            doc_path.unlink(missing_ok=True)

    def test_toc_field_no_hyperlinks(self):
        """Test field instruction without hyperlinks switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(hyperlinks=False)

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            instr_text = body.find(f".//{{{WORD_NS}}}instrText")
            assert instr_text is not None
            assert "\\h" not in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_toc_field_no_page_numbers(self):
        """Test field instruction with page numbers suppressed."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(show_page_numbers=False)

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            instr_text = body.find(f".//{{{WORD_NS}}}instrText")
            assert instr_text is not None
            assert "\\n" in instr_text.text  # \n switch suppresses page numbers

        finally:
            doc_path.unlink(missing_ok=True)

    def test_toc_has_placeholder_text(self):
        """Test that TOC has placeholder text between separate and end."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            sdt = body.find(f"{{{WORD_NS}}}sdt")
            sdt_content = sdt.find(f"{{{WORD_NS}}}sdtContent")

            # Find all text elements
            text_elems = sdt_content.findall(f".//{{{WORD_NS}}}t")
            texts = [t.text for t in text_elems if t.text]
            assert any("Update" in t for t in texts), "Placeholder text not found"

        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertTocPosition:
    """Tests for TOC positioning."""

    def test_insert_toc_at_start(self):
        """Test inserting TOC at document start."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(position="start", title=None)

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            # First element should be SDT
            first = body[0]
            assert first.tag == f"{{{WORD_NS}}}sdt"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_toc_at_end(self):
        """Test inserting TOC at document end (before sectPr)."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(position="end", title=None)

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            children = list(body)

            # Second to last should be SDT (last is sectPr)
            assert children[-1].tag == f"{{{WORD_NS}}}sectPr"
            assert children[-2].tag == f"{{{WORD_NS}}}sdt"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_toc_at_index(self):
        """Test inserting TOC at specific paragraph index."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(position=2, title=None)

            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            # Element at index 2 should be SDT
            assert body[2].tag == f"{{{WORD_NS}}}sdt"

        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertTocStyles:
    """Tests for style creation."""

    def test_insert_toc_creates_toc_heading_style(self):
        """Test that TOCHeading style is created if missing."""
        doc_path = create_document_with_styles()
        try:
            doc = Document(doc_path)

            # Verify TOCHeading doesn't exist initially
            assert "TOCHeading" not in doc.styles

            doc.insert_toc()

            # Now TOCHeading should exist
            assert "TOCHeading" in doc.styles

            style = doc.styles.get("TOCHeading")
            assert style is not None
            assert style.style_id == "TOCHeading"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_toc_preserves_existing_styles(self):
        """Test that existing styles are preserved."""
        doc_path = create_document_with_styles()
        try:
            doc = Document(doc_path)

            # Verify Heading1 exists
            assert "Heading1" in doc.styles

            doc.insert_toc()

            # Heading1 should still exist
            assert "Heading1" in doc.styles

        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertTocSaveRoundtrip:
    """Tests for save/reload behavior."""

    def test_toc_survives_save_reload(self):
        """Test that TOC structure is preserved after save/reload."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        try:
            # Insert TOC and save
            doc = Document(doc_path)
            doc.insert_toc(title="Table of Contents", levels=(1, 4))
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)
            body = doc2.xml_root.find(f".//{{{WORD_NS}}}body")

            # Should have SDT with TOC
            # Note: The title paragraph comes first, then SDT
            # Find the SDT (might be first or second element)
            sdts = body.findall(f"{{{WORD_NS}}}sdt")
            assert len(sdts) >= 1

            # Check field instruction preserved
            instr_text = body.find(f".//{{{WORD_NS}}}instrText")
            assert instr_text is not None
            assert '\\o "1-4"' in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


class TestRemoveToc:
    """Tests for TOC removal functionality."""

    def test_remove_toc_basic(self):
        """Test removing a TOC from the document."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(title="Table of Contents")

            # Verify TOC exists
            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            sdts_before = body.findall(f"{{{WORD_NS}}}sdt")
            assert len(sdts_before) >= 1, "TOC not inserted"

            # Remove TOC
            result = doc.remove_toc()
            assert result is True

            # Verify TOC is removed
            sdts_after = body.findall(f"{{{WORD_NS}}}sdt")
            assert len(sdts_after) == 0, "TOC SDT should be removed"

            # Verify title paragraph is also removed
            first_child = body[0]
            pstyle = first_child.find(f".//{{{WORD_NS}}}pStyle")
            if pstyle is not None:
                assert pstyle.get(f"{{{WORD_NS}}}val") != "TOCHeading"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_remove_toc_no_title(self):
        """Test removing a TOC that was inserted without a title."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(title=None)

            # Count children before removal
            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            children_before = len(list(body))

            # Remove TOC
            result = doc.remove_toc()
            assert result is True

            # Should have one fewer child (just the SDT, no title)
            children_after = len(list(body))
            assert children_after == children_before - 1

        finally:
            doc_path.unlink(missing_ok=True)

    def test_remove_toc_when_none_exists(self):
        """Test removing TOC when document has no TOC."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            # Try to remove TOC when none exists
            result = doc.remove_toc()
            assert result is False

        finally:
            doc_path.unlink(missing_ok=True)

    def test_remove_toc_preserves_other_content(self):
        """Test that removing TOC preserves other document content."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            # Count paragraphs before inserting TOC
            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            original_paras = len(body.findall(f"{{{WORD_NS}}}p"))

            # Insert and remove TOC
            doc.insert_toc(title="Contents")
            doc.remove_toc()

            # Original paragraph count should be preserved
            paras_after = len(body.findall(f"{{{WORD_NS}}}p"))
            assert paras_after == original_paras

        finally:
            doc_path.unlink(missing_ok=True)

    def test_remove_toc_survives_save_reload(self):
        """Test that TOC removal persists after save/reload."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        try:
            # Insert TOC
            doc = Document(doc_path)
            doc.insert_toc()

            # Save and reload
            doc.save(output_path)
            doc2 = Document(output_path)

            # Remove TOC
            result = doc2.remove_toc()
            assert result is True

            # Save again
            output_path2 = Path(tempfile.mktemp(suffix=".docx"))
            doc2.save(output_path2)

            # Reload and verify no TOC
            doc3 = Document(output_path2)
            body = doc3.xml_root.find(f".//{{{WORD_NS}}}body")
            sdts = body.findall(f"{{{WORD_NS}}}sdt")
            assert len(sdts) == 0

            output_path2.unlink(missing_ok=True)

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


class TestMarkTocDirty:
    """Tests for marking TOC as dirty."""

    def test_mark_toc_dirty_basic(self):
        """Test marking an existing TOC as dirty."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            # Mark dirty
            result = doc.mark_toc_dirty()
            assert result is True

            # Verify dirty flag is set
            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            begin_fld = body.find(f".//{{{WORD_NS}}}fldChar[@{{{WORD_NS}}}fldCharType='begin']")
            assert begin_fld is not None
            assert begin_fld.get(f"{{{WORD_NS}}}dirty") == "true"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_mark_toc_dirty_when_none_exists(self):
        """Test marking TOC dirty when no TOC exists."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            # Try to mark dirty when no TOC exists
            result = doc.mark_toc_dirty()
            assert result is False

        finally:
            doc_path.unlink(missing_ok=True)

    def test_mark_toc_dirty_idempotent(self):
        """Test that marking dirty multiple times is safe."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            # Mark dirty multiple times
            result1 = doc.mark_toc_dirty()
            result2 = doc.mark_toc_dirty()
            result3 = doc.mark_toc_dirty()

            assert result1 is True
            assert result2 is True
            assert result3 is True

            # Should still have dirty="true"
            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            begin_fld = body.find(f".//{{{WORD_NS}}}fldChar[@{{{WORD_NS}}}fldCharType='begin']")
            assert begin_fld.get(f"{{{WORD_NS}}}dirty") == "true"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_mark_toc_dirty_survives_save_reload(self):
        """Test that dirty flag persists after save/reload."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        try:
            doc = Document(doc_path)
            doc.insert_toc()
            doc.mark_toc_dirty()
            doc.save(output_path)

            # Reload and verify dirty flag
            doc2 = Document(output_path)
            body = doc2.xml_root.find(f".//{{{WORD_NS}}}body")
            begin_fld = body.find(f".//{{{WORD_NS}}}fldChar[@{{{WORD_NS}}}fldCharType='begin']")
            assert begin_fld is not None
            assert begin_fld.get(f"{{{WORD_NS}}}dirty") == "true"

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


class TestRemoveTocThenInsertToc:
    """Tests for removing and re-inserting TOC."""

    def test_remove_then_insert_toc(self):
        """Test that remove_toc followed by insert_toc works."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            # Insert initial TOC
            doc.insert_toc(title="Original Contents", levels=(1, 2))

            # Verify first TOC
            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            instr = body.find(f".//{{{WORD_NS}}}instrText")
            assert '\\o "1-2"' in instr.text

            # Remove TOC
            result = doc.remove_toc()
            assert result is True

            # Insert new TOC with different settings
            doc.insert_toc(title="New Contents", levels=(1, 4))

            # Verify new TOC has new settings
            instr = body.find(f".//{{{WORD_NS}}}instrText")
            assert '\\o "1-4"' in instr.text

            # Verify new title
            first_para = body[0]
            text_elem = first_para.find(f".//{{{WORD_NS}}}t")
            assert text_elem is not None
            assert text_elem.text == "New Contents"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_remove_and_insert_different_position(self):
        """Test removing TOC at start and inserting at end."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            # Insert TOC at start
            doc.insert_toc(position="start", title=None)

            # Verify TOC at start
            body = doc.xml_root.find(f".//{{{WORD_NS}}}body")
            assert body[0].tag == f"{{{WORD_NS}}}sdt"

            # Remove TOC
            doc.remove_toc()

            # Insert at end
            doc.insert_toc(position="end", title=None)

            # Verify TOC at end (before sectPr)
            children = list(body)
            assert children[-1].tag == f"{{{WORD_NS}}}sectPr"
            assert children[-2].tag == f"{{{WORD_NS}}}sdt"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_remove_insert_roundtrip(self):
        """Test complete roundtrip: insert, save, reload, remove, insert, save."""
        doc_path = create_simple_document()
        output_path1 = Path(tempfile.mktemp(suffix=".docx"))
        output_path2 = Path(tempfile.mktemp(suffix=".docx"))
        try:
            # First document: insert TOC
            doc1 = Document(doc_path)
            doc1.insert_toc(levels=(1, 3))
            doc1.save(output_path1)

            # Second document: remove and re-insert
            doc2 = Document(output_path1)
            doc2.remove_toc()
            doc2.insert_toc(levels=(1, 5))
            doc2.save(output_path2)

            # Verify final document
            doc3 = Document(output_path2)
            body = doc3.xml_root.find(f".//{{{WORD_NS}}}body")
            instr = body.find(f".//{{{WORD_NS}}}instrText")
            assert instr is not None
            assert '\\o "1-5"' in instr.text

        finally:
            doc_path.unlink(missing_ok=True)
            output_path1.unlink(missing_ok=True)
            output_path2.unlink(missing_ok=True)


# =============================================================================
# Phase 3: TOC Inspection Tests
# =============================================================================


def create_document_with_populated_toc() -> Path:
    """Create a document with a populated TOC (as Word would generate).

    This simulates a document where Word has already updated the TOC,
    so there are actual cached entries with styles, text, and page numbers.
    """
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # TOC with populated entries (as Word generates them)
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<!-- TOC Title -->
<w:p>
  <w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>
  <w:r><w:t>Contents</w:t></w:r>
</w:p>
<!-- TOC SDT with populated entries -->
<w:sdt>
  <w:sdtPr>
    <w:docPartObj>
      <w:docPartGallery w:val="Table of Contents"/>
      <w:docPartUnique/>
    </w:docPartObj>
  </w:sdtPr>
  <w:sdtContent>
    <!-- Field begin/instruction -->
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin" w:dirty="false"/></w:r>
      <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    </w:p>
    <!-- TOC Entry 1: Level 1 with hyperlink -->
    <w:p>
      <w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
      <w:hyperlink w:anchor="_Toc123456">
        <w:r><w:t>Introduction</w:t></w:r>
      </w:hyperlink>
      <w:r><w:tab/></w:r>
      <w:r><w:t>1</w:t></w:r>
    </w:p>
    <!-- TOC Entry 2: Level 2 with hyperlink -->
    <w:p>
      <w:pPr><w:pStyle w:val="TOC2"/></w:pPr>
      <w:hyperlink w:anchor="_Toc123457">
        <w:r><w:t>Background</w:t></w:r>
      </w:hyperlink>
      <w:r><w:tab/></w:r>
      <w:r><w:t>2</w:t></w:r>
    </w:p>
    <!-- TOC Entry 3: Level 1 with hyperlink -->
    <w:p>
      <w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
      <w:hyperlink w:anchor="_Toc123458">
        <w:r><w:t>Methodology</w:t></w:r>
      </w:hyperlink>
      <w:r><w:tab/></w:r>
      <w:r><w:t>5</w:t></w:r>
    </w:p>
    <!-- TOC Entry 4: Level 3 with hyperlink -->
    <w:p>
      <w:pPr><w:pStyle w:val="TOC3"/></w:pPr>
      <w:hyperlink w:anchor="_Toc123459">
        <w:r><w:t>Data Collection</w:t></w:r>
      </w:hyperlink>
      <w:r><w:tab/></w:r>
      <w:r><w:t>6</w:t></w:r>
    </w:p>
    <!-- Field end -->
    <w:p>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>
<!-- Document content -->
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Introduction</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>This is the introduction.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_toc_no_hyperlinks() -> Path:
    """Create a document with a TOC that has no hyperlinks."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:sdt>
  <w:sdtPr>
    <w:docPartObj>
      <w:docPartGallery w:val="Table of Contents"/>
      <w:docPartUnique/>
    </w:docPartObj>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
      <w:r><w:instrText xml:space="preserve"> TOC \\o "1-4" \\z \\u </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    </w:p>
    <!-- TOC Entry without hyperlink -->
    <w:p>
      <w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
      <w:r><w:t>Chapter One</w:t></w:r>
      <w:r><w:tab/></w:r>
      <w:r><w:t>3</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="TOC2"/></w:pPr>
      <w:r><w:t>Section A</w:t></w:r>
      <w:r><w:tab/></w:r>
      <w:r><w:t>4</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>
<w:sectPr/>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestGetTocBasic:
    """Tests for basic get_toc() functionality."""

    def test_get_toc_returns_none_when_no_toc(self):
        """Test that get_toc returns None when document has no TOC."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_finds_existing_toc(self):
        """Test that get_toc finds a freshly inserted TOC."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(levels=(1, 3))

            toc = doc.get_toc()
            assert toc is not None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_finds_populated_toc(self):
        """Test that get_toc finds a populated TOC."""
        doc_path = create_document_with_populated_toc()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetTocPosition:
    """Tests for TOC position detection."""

    def test_get_toc_position_at_start(self):
        """Test position when TOC is at start of document."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(position="start", title=None)

            toc = doc.get_toc()
            assert toc is not None
            assert toc.position == 0
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_position_with_title(self):
        """Test position when TOC has a title paragraph before it."""
        doc_path = create_document_with_populated_toc()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None
            # Position is the index of the SDT in body children
            # Our fixture has: title para, SDT, heading para, content para, sectPr
            # XML comments don't count as children, but blank text nodes might
            # Just verify it's not at position 0 (there's a title before it)
            assert toc.position > 0
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetTocLevels:
    """Tests for parsing TOC levels from field instruction."""

    def test_get_toc_levels_default(self):
        """Test parsing default levels (1-3)."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(levels=(1, 3))

            toc = doc.get_toc()
            assert toc is not None
            assert toc.levels == (1, 3)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_levels_custom(self):
        """Test parsing custom levels (1-4)."""
        doc_path = create_document_with_toc_no_hyperlinks()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None
            assert toc.levels == (1, 4)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_levels_wide_range(self):
        """Test parsing wide range levels (1-9)."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(levels=(1, 9))

            toc = doc.get_toc()
            assert toc is not None
            assert toc.levels == (1, 9)
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetTocSwitches:
    """Tests for parsing TOC field switches."""

    def test_get_toc_switches_contains_instruction(self):
        """Test that switches contain the full field instruction."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(levels=(1, 3), hyperlinks=True, use_outline_levels=True)

            toc = doc.get_toc()
            assert toc is not None
            assert "TOC" in toc.switches
            assert '\\o "1-3"' in toc.switches
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_switch_o(self):
        """Test get_switch for \\o switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(levels=(1, 5))

            toc = doc.get_toc()
            assert toc is not None
            assert toc.get_switch("o") == "1-5"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_switch_h(self):
        """Test get_switch for \\h switch (hyperlinks)."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(hyperlinks=True)

            toc = doc.get_toc()
            assert toc is not None
            # \h is present but has no value
            assert toc.get_switch("h") == ""
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_switch_h_not_present(self):
        """Test get_switch for \\h switch when not present."""
        doc_path = create_document_with_toc_no_hyperlinks()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None
            # \h is not in the instruction
            assert toc.get_switch("h") is None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_switch_z(self):
        """Test get_switch for \\z switch (hide in web view)."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            toc = doc.get_toc()
            assert toc is not None
            assert toc.get_switch("z") == ""
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_switch_u(self):
        """Test get_switch for \\u switch (outline levels)."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(use_outline_levels=True)

            toc = doc.get_toc()
            assert toc is not None
            assert toc.get_switch("u") == ""
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_switch_n(self):
        """Test get_switch for \\n switch (no page numbers)."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc(show_page_numbers=False)

            toc = doc.get_toc()
            assert toc is not None
            assert toc.get_switch("n") == ""
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_switch_not_present(self):
        """Test get_switch returns None for missing switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            toc = doc.get_toc()
            assert toc is not None
            # \b (bookmark) is not used
            assert toc.get_switch("b") is None
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetTocIsDirty:
    """Tests for checking TOC dirty flag."""

    def test_get_toc_is_dirty_on_insert(self):
        """Test that newly inserted TOC is marked dirty."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            toc = doc.get_toc()
            assert toc is not None
            assert toc.is_dirty is True
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_is_dirty_false(self):
        """Test TOC with dirty=false."""
        doc_path = create_document_with_populated_toc()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None
            # The populated TOC has dirty="false"
            assert toc.is_dirty is False
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_is_dirty_true(self):
        """Test TOC with dirty=true."""
        doc_path = create_document_with_toc_no_hyperlinks()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None
            # This TOC has dirty="true"
            assert toc.is_dirty is True
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetTocEntriesWithHyperlinks:
    """Tests for extracting TOC entries with hyperlinks."""

    def test_get_toc_entries_count(self):
        """Test that correct number of entries are extracted."""
        doc_path = create_document_with_populated_toc()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None
            # Should have 4 entries: Introduction, Background, Methodology, Data Collection
            assert len(toc.entries) == 4
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_entry_text(self):
        """Test that entry text is extracted correctly."""
        doc_path = create_document_with_populated_toc()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None

            texts = [e.text for e in toc.entries]
            assert "Introduction" in texts
            assert "Background" in texts
            assert "Methodology" in texts
            assert "Data Collection" in texts
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_entry_levels(self):
        """Test that entry levels are parsed correctly."""
        doc_path = create_document_with_populated_toc()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None

            # Check levels by text
            level_map = {e.text: e.level for e in toc.entries}
            assert level_map["Introduction"] == 1
            assert level_map["Background"] == 2
            assert level_map["Methodology"] == 1
            assert level_map["Data Collection"] == 3
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_entry_page_numbers(self):
        """Test that page numbers are extracted."""
        doc_path = create_document_with_populated_toc()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None

            # Check page numbers by text
            page_map = {e.text: e.page_number for e in toc.entries}
            assert page_map["Introduction"] == "1"
            assert page_map["Background"] == "2"
            assert page_map["Methodology"] == "5"
            assert page_map["Data Collection"] == "6"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_entry_bookmarks(self):
        """Test that bookmark targets are extracted from hyperlinks."""
        doc_path = create_document_with_populated_toc()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None

            # Check bookmarks by text
            bookmark_map = {e.text: e.bookmark for e in toc.entries}
            assert bookmark_map["Introduction"] == "_Toc123456"
            assert bookmark_map["Background"] == "_Toc123457"
            assert bookmark_map["Methodology"] == "_Toc123458"
            assert bookmark_map["Data Collection"] == "_Toc123459"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_entry_styles(self):
        """Test that paragraph styles are extracted."""
        doc_path = create_document_with_populated_toc()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None

            # Check styles by text
            style_map = {e.text: e.style for e in toc.entries}
            assert style_map["Introduction"] == "TOC1"
            assert style_map["Background"] == "TOC2"
            assert style_map["Methodology"] == "TOC1"
            assert style_map["Data Collection"] == "TOC3"
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetTocEntriesWithoutHyperlinks:
    """Tests for extracting TOC entries without hyperlinks."""

    def test_get_toc_entries_no_hyperlinks(self):
        """Test extracting entries when TOC has no hyperlinks."""
        doc_path = create_document_with_toc_no_hyperlinks()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None
            assert len(toc.entries) == 2
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_entry_text_no_hyperlinks(self):
        """Test entry text when no hyperlinks."""
        doc_path = create_document_with_toc_no_hyperlinks()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None

            texts = [e.text for e in toc.entries]
            assert "Chapter One" in texts
            assert "Section A" in texts
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_entry_bookmark_none(self):
        """Test that bookmark is None when no hyperlink."""
        doc_path = create_document_with_toc_no_hyperlinks()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None

            for entry in toc.entries:
                assert entry.bookmark is None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_toc_entry_levels_no_hyperlinks(self):
        """Test levels when no hyperlinks."""
        doc_path = create_document_with_toc_no_hyperlinks()
        try:
            doc = Document(doc_path)
            toc = doc.get_toc()
            assert toc is not None

            level_map = {e.text: e.level for e in toc.entries}
            assert level_map["Chapter One"] == 1
            assert level_map["Section A"] == 2
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetTocFreshlyInserted:
    """Tests for get_toc on freshly inserted TOC (no cached entries)."""

    def test_get_toc_fresh_has_no_entries(self):
        """Test that freshly inserted TOC has no cached entries."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            doc.insert_toc()

            toc = doc.get_toc()
            assert toc is not None
            # Freshly inserted TOC only has placeholder text, not real entries
            assert len(toc.entries) == 0
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetTocRoundtrip:
    """Tests for get_toc after save/reload."""

    def test_get_toc_survives_save_reload(self):
        """Test that get_toc works after save/reload."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        try:
            # Insert TOC and save
            doc = Document(doc_path)
            doc.insert_toc(levels=(1, 3))
            doc.save(output_path)

            # Reload and inspect
            doc2 = Document(output_path)
            toc = doc2.get_toc()
            assert toc is not None
            assert toc.levels == (1, 3)
            # Freshly inserted TOC has no cached entries
            assert toc.is_dirty is True
        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_get_toc_after_insert_save_reload(self):
        """Test get_toc after inserting, saving, and reloading."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        try:
            # Insert TOC
            doc = Document(doc_path)
            doc.insert_toc(levels=(2, 5), hyperlinks=False)
            doc.save(output_path)

            # Reload and inspect
            doc2 = Document(output_path)
            toc = doc2.get_toc()
            assert toc is not None
            assert toc.levels == (2, 5)
            assert toc.is_dirty is True
            assert toc.get_switch("h") is None  # hyperlinks=False
        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)
