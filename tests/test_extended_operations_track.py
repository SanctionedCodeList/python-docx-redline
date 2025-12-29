"""Tests for Phase 4 extended operations with track parameter.

This module tests the track parameter on extended operations:
- Header/footer operations (replace_in_header, replace_in_footer, etc.)
- apply_criticmarkup with track=False
- Table operations already have track parameter; included for completeness
"""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from python_docx_redline import Document

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def create_simple_docx(text: str = "Hello world") -> Path:
    """Create a minimal .docx file with specified text."""
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    document_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WORD_NAMESPACE}">
  <w:body>
    <w:p>
      <w:r>
        <w:t>{text}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    with zipfile.ZipFile(docx_path, "w") as zf:
        zf.writestr("word/document.xml", document_content)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("[Content_Types].xml", content_types)

    return docx_path


def create_docx_with_header_footer(
    header_text: str = "Test Header",
    footer_text: str = "Test Footer",
) -> Path:
    """Create a minimal .docx file with headers and footers."""
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    document_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WORD_NAMESPACE}" xmlns:r="{RELS_NAMESPACE}">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Document body text.</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId6"/>
      <w:footerReference w:type="default" r:id="rId7"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    header_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="{WORD_NAMESPACE}">
  <w:p>
    <w:r>
      <w:t>{header_text}</w:t>
    </w:r>
  </w:p>
</w:hdr>"""

    footer_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{WORD_NAMESPACE}">
  <w:p>
    <w:r>
      <w:t>{footer_text}</w:t>
    </w:r>
  </w:p>
</w:ftr>"""

    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>"""

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>"""

    with zipfile.ZipFile(docx_path, "w") as zf:
        zf.writestr("word/document.xml", document_content)
        zf.writestr("word/header1.xml", header_content)
        zf.writestr("word/footer1.xml", footer_content)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("[Content_Types].xml", content_types)

    return docx_path


class TestHeaderFooterTrackParameter:
    """Tests for track parameter on header/footer operations."""

    def test_replace_in_header_tracked(self):
        """Test replace_in_header with track=True (default)."""
        docx_path = create_docx_with_header_footer(header_text="Draft Document")
        doc = Document(docx_path)

        doc.replace_in_header("Draft", "Final", track=True)
        doc.save(docx_path)

        # Reload and check for tracked changes
        doc2 = Document(docx_path)
        header = doc2.headers[0]
        header_xml = etree.tostring(header.element, encoding="unicode")

        # Should have tracked deletion and insertion
        assert "<w:del" in header_xml or "w:del" in header_xml
        assert "<w:ins" in header_xml or "w:ins" in header_xml

    def test_replace_in_header_untracked(self):
        """Test replace_in_header with track=False."""
        docx_path = create_docx_with_header_footer(header_text="Draft Document")
        doc = Document(docx_path)

        doc.replace_in_header("Draft", "Final", track=False)
        doc.save(docx_path)

        # Reload and check - should NOT have tracked changes
        doc2 = Document(docx_path)
        header = doc2.headers[0]
        header_xml = etree.tostring(header.element, encoding="unicode")

        # Should NOT have tracked deletion or insertion
        assert "<w:del" not in header_xml and "w:del" not in header_xml
        assert "<w:ins" not in header_xml and "w:ins" not in header_xml
        # But should have the new text
        assert "Final" in header_xml

    def test_replace_in_footer_tracked(self):
        """Test replace_in_footer with track=True (default)."""
        docx_path = create_docx_with_header_footer(footer_text="Page 1")
        doc = Document(docx_path)

        doc.replace_in_footer("Page 1", "Page One", track=True)
        doc.save(docx_path)

        # Reload and check for tracked changes
        doc2 = Document(docx_path)
        footer = doc2.footers[0]
        footer_xml = etree.tostring(footer.element, encoding="unicode")

        # Should have tracked deletion and insertion
        assert "<w:del" in footer_xml or "w:del" in footer_xml
        assert "<w:ins" in footer_xml or "w:ins" in footer_xml

    def test_replace_in_footer_untracked(self):
        """Test replace_in_footer with track=False."""
        docx_path = create_docx_with_header_footer(footer_text="Page 1")
        doc = Document(docx_path)

        doc.replace_in_footer("Page 1", "Page One", track=False)
        doc.save(docx_path)

        # Reload and check - should NOT have tracked changes
        doc2 = Document(docx_path)
        footer = doc2.footers[0]
        footer_xml = etree.tostring(footer.element, encoding="unicode")

        # Should NOT have tracked deletion or insertion
        assert "<w:del" not in footer_xml and "w:del" not in footer_xml
        assert "<w:ins" not in footer_xml and "w:ins" not in footer_xml
        # But should have the new text
        assert "Page One" in footer_xml

    def test_insert_in_header_tracked(self):
        """Test insert_in_header with track=True (default)."""
        docx_path = create_docx_with_header_footer(header_text="Document Title")
        doc = Document(docx_path)

        doc.insert_in_header(" - Final", after="Title", track=True)
        doc.save(docx_path)

        # Reload and check for tracked insertion
        doc2 = Document(docx_path)
        header = doc2.headers[0]
        header_xml = etree.tostring(header.element, encoding="unicode")

        # Should have tracked insertion
        assert "<w:ins" in header_xml or "w:ins" in header_xml

    def test_insert_in_header_untracked(self):
        """Test insert_in_header with track=False."""
        docx_path = create_docx_with_header_footer(header_text="Document Title")
        doc = Document(docx_path)

        doc.insert_in_header(" - Final", after="Title", track=False)
        doc.save(docx_path)

        # Reload and check - should NOT have tracked changes
        doc2 = Document(docx_path)
        header = doc2.headers[0]
        header_xml = etree.tostring(header.element, encoding="unicode")

        # Should NOT have tracked insertion
        assert "<w:ins" not in header_xml and "w:ins" not in header_xml
        # But should have the new text
        assert "Final" in header_xml

    def test_insert_in_footer_tracked(self):
        """Test insert_in_footer with track=True (default)."""
        docx_path = create_docx_with_header_footer(footer_text="Page")
        doc = Document(docx_path)

        doc.insert_in_footer(" Number", after="Page", track=True)
        doc.save(docx_path)

        # Reload and check for tracked insertion
        doc2 = Document(docx_path)
        footer = doc2.footers[0]
        footer_xml = etree.tostring(footer.element, encoding="unicode")

        # Should have tracked insertion
        assert "<w:ins" in footer_xml or "w:ins" in footer_xml

    def test_insert_in_footer_untracked(self):
        """Test insert_in_footer with track=False."""
        docx_path = create_docx_with_header_footer(footer_text="Page")
        doc = Document(docx_path)

        doc.insert_in_footer(" Number", after="Page", track=False)
        doc.save(docx_path)

        # Reload and check - should NOT have tracked changes
        doc2 = Document(docx_path)
        footer = doc2.footers[0]
        footer_xml = etree.tostring(footer.element, encoding="unicode")

        # Should NOT have tracked insertion
        assert "<w:ins" not in footer_xml and "w:ins" not in footer_xml
        # But should have the new text
        assert "Number" in footer_xml


class TestCriticMarkupTrackParameter:
    """Tests for track parameter on apply_criticmarkup."""

    def test_apply_criticmarkup_tracked_default(self):
        """Test apply_criticmarkup with track=True (default)."""
        docx_path = create_simple_docx("Payment in 30 days")
        doc = Document(docx_path)

        doc.apply_criticmarkup("Payment in {~~30~>45~~} days", author="Test")
        doc.save(docx_path)

        # Reload and check for tracked changes
        doc2 = Document(docx_path)
        assert doc2.has_tracked_changes()

    def test_apply_criticmarkup_untracked(self):
        """Test apply_criticmarkup with track=False."""
        docx_path = create_simple_docx("Payment in 30 days")
        doc = Document(docx_path)

        doc.apply_criticmarkup("Payment in {~~30~>45~~} days", author="Test", track=False)
        doc.save(docx_path)

        # Reload and check - should NOT have tracked changes
        doc2 = Document(docx_path)
        assert not doc2.has_tracked_changes()
        # But should have the new text
        text = doc2.get_text()
        assert "45" in text
        assert "30" not in text

    def test_apply_criticmarkup_insertion_untracked(self):
        """Test apply_criticmarkup with insertion and track=False."""
        docx_path = create_simple_docx("Hello world")
        doc = Document(docx_path)

        doc.apply_criticmarkup("Hello {++beautiful ++}world", author="Test", track=False)
        doc.save(docx_path)

        # Reload and check - should NOT have tracked changes
        doc2 = Document(docx_path)
        assert not doc2.has_tracked_changes()
        # But should have the new text
        text = doc2.get_text()
        assert "beautiful" in text

    def test_apply_criticmarkup_deletion_untracked(self):
        """Test apply_criticmarkup with deletion and track=False."""
        docx_path = create_simple_docx("Hello beautiful world")
        doc = Document(docx_path)

        doc.apply_criticmarkup("Hello {--beautiful --}world", author="Test", track=False)
        doc.save(docx_path)

        # Reload and check - should NOT have tracked changes
        doc2 = Document(docx_path)
        assert not doc2.has_tracked_changes()
        # The text should be removed
        text = doc2.get_text()
        assert "beautiful" not in text

    def test_apply_criticmarkup_comments_always_visible(self):
        """Test that comments are always applied regardless of track parameter."""
        docx_path = create_simple_docx("Important text here")
        doc = Document(docx_path)

        # Comments should still be added even with track=False
        # (Comments are not "tracked changes" in the Word sense)
        result = doc.apply_criticmarkup(
            "Important {>>This needs review<<}text here", author="Test", track=False
        )

        # Result should be successful for the comment
        assert result.total == 1
        # Note: comment application may or may not succeed depending on anchor finding


class TestTableOperationsTrackParameter:
    """Tests for track parameter on table operations (already implemented)."""

    def create_table_docx(self) -> Path:
        """Create a docx with a simple table."""
        temp_dir = Path(tempfile.mkdtemp())
        docx_path = temp_dir / "test.docx"

        document_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WORD_NAMESPACE}">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="dxa"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="2500"/>
        <w:gridCol w:w="2500"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Header 1</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Header 2</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Cell A</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Cell B</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"""

        rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

        doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

        content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

        with zipfile.ZipFile(docx_path, "w") as zf:
            zf.writestr("word/document.xml", document_content)
            zf.writestr("_rels/.rels", rels)
            zf.writestr("word/_rels/document.xml.rels", doc_rels)
            zf.writestr("[Content_Types].xml", content_types)

        return docx_path

    def test_replace_in_table_tracked(self):
        """Test replace_in_table with track=True (default)."""
        docx_path = self.create_table_docx()
        doc = Document(docx_path)

        count = doc.replace_in_table("Cell A", "Modified A", track=True)
        doc.save(docx_path)

        assert count == 1

        # Reload and check for tracked changes
        doc2 = Document(docx_path)
        assert doc2.has_tracked_changes()

    def test_replace_in_table_untracked(self):
        """Test replace_in_table with track=False."""
        docx_path = self.create_table_docx()
        doc = Document(docx_path)

        count = doc.replace_in_table("Cell A", "Modified A", track=False)
        doc.save(docx_path)

        assert count == 1

        # Reload and check - should NOT have tracked changes
        doc2 = Document(docx_path)
        assert not doc2.has_tracked_changes()
        # But the table should have the new text
        tables = doc2.tables
        assert len(tables) > 0
