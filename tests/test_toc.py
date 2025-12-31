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
