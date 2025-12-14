"""Tests for header and footer functionality."""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from python_docx_redline import Document

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def create_test_docx_with_header_footer(
    header_text: str = "Test Header",
    footer_text: str = "Test Footer",
    first_header_text: str | None = None,
    first_footer_text: str | None = None,
) -> Path:
    """Create a minimal .docx file with headers and footers."""
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    # Document content
    document_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WORD_NAMESPACE}" xmlns:r="{RELS_NAMESPACE}">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a test document with headers and footers.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Another paragraph for testing.</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId6"/>
      <w:footerReference w:type="default" r:id="rId7"/>
      {"<w:headerReference w:type='first' r:id='rId8'/>" if first_header_text else ""}
      {"<w:footerReference w:type='first' r:id='rId9'/>" if first_footer_text else ""}
      <w:titlePg/>
    </w:sectPr>
  </w:body>
</w:document>"""

    # Default header
    header1_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="{WORD_NAMESPACE}">
  <w:p>
    <w:r>
      <w:t>{header_text}</w:t>
    </w:r>
  </w:p>
</w:hdr>"""

    # Default footer
    footer1_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{WORD_NAMESPACE}">
  <w:p>
    <w:r>
      <w:t>{footer_text}</w:t>
    </w:r>
  </w:p>
</w:ftr>"""

    # First page header (optional)
    header2_content = (
        f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="{WORD_NAMESPACE}">
  <w:p>
    <w:r>
      <w:t>{first_header_text or "First Header"}</w:t>
    </w:r>
  </w:p>
</w:hdr>"""
        if first_header_text
        else None
    )

    # First page footer (optional)
    footer2_content = (
        f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{WORD_NAMESPACE}">
  <w:p>
    <w:r>
      <w:t>{first_footer_text or "First Footer"}</w:t>
    </w:r>
  </w:p>
</w:ftr>"""
        if first_footer_text
        else None
    )

    # Relationships with all needed entries
    rels = f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  {"<Relationship Id='rId8' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/header' Target='header2.xml'/>" if first_header_text else ""}
  {"<Relationship Id='rId9' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer' Target='footer2.xml'/>" if first_footer_text else ""}
</Relationships>"""

    # Content types
    content_types = f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  {"<Override PartName='/word/header2.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'/>" if first_header_text else ""}
  {"<Override PartName='/word/footer2.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'/>" if first_footer_text else ""}
</Types>"""

    # Root relationships
    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_content)
        docx.writestr("word/_rels/document.xml.rels", rels)
        docx.writestr("word/header1.xml", header1_content)
        docx.writestr("word/footer1.xml", footer1_content)
        if header2_content:
            docx.writestr("word/header2.xml", header2_content)
        if footer2_content:
            docx.writestr("word/footer2.xml", footer2_content)

    return docx_path


def create_simple_docx_no_headers_footers() -> Path:
    """Create a minimal .docx file without headers and footers."""
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    document_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WORD_NAMESPACE}">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a test document without headers or footers.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

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
        docx.writestr("word/document.xml", document_content)

    return docx_path


class TestHeadersProperty:
    """Tests for headers property."""

    def test_headers_property_returns_list(self):
        """Test that headers property returns a list."""
        doc = Document(create_test_docx_with_header_footer())
        headers = doc.headers
        assert isinstance(headers, list)

    def test_headers_property_finds_default_header(self):
        """Test that headers property finds the default header."""
        doc = Document(create_test_docx_with_header_footer(header_text="My Header"))
        headers = doc.headers
        assert len(headers) >= 1
        assert any(h.type == "default" for h in headers)

    def test_headers_property_empty_when_no_headers(self):
        """Test that headers property returns empty list when no headers."""
        doc = Document(create_simple_docx_no_headers_footers())
        headers = doc.headers
        assert headers == []

    def test_headers_property_with_multiple_types(self):
        """Test headers property with multiple header types."""
        doc = Document(
            create_test_docx_with_header_footer(
                header_text="Default Header",
                first_header_text="First Page Header",
            )
        )
        headers = doc.headers
        assert len(headers) >= 2
        types = {h.type for h in headers}
        assert "default" in types
        assert "first" in types


class TestFootersProperty:
    """Tests for footers property."""

    def test_footers_property_returns_list(self):
        """Test that footers property returns a list."""
        doc = Document(create_test_docx_with_header_footer())
        footers = doc.footers
        assert isinstance(footers, list)

    def test_footers_property_finds_default_footer(self):
        """Test that footers property finds the default footer."""
        doc = Document(create_test_docx_with_header_footer(footer_text="My Footer"))
        footers = doc.footers
        assert len(footers) >= 1
        assert any(f.type == "default" for f in footers)

    def test_footers_property_empty_when_no_footers(self):
        """Test that footers property returns empty list when no footers."""
        doc = Document(create_simple_docx_no_headers_footers())
        footers = doc.footers
        assert footers == []

    def test_footers_property_with_multiple_types(self):
        """Test footers property with multiple footer types."""
        doc = Document(
            create_test_docx_with_header_footer(
                footer_text="Default Footer",
                first_footer_text="First Page Footer",
            )
        )
        footers = doc.footers
        assert len(footers) >= 2
        types = {f.type for f in footers}
        assert "default" in types
        assert "first" in types


class TestHeaderModel:
    """Tests for Header model class."""

    def test_header_text_property(self):
        """Test header text property."""
        doc = Document(create_test_docx_with_header_footer(header_text="Sample Header"))
        header = doc.headers[0]
        assert header.text == "Sample Header"

    def test_header_type_property(self):
        """Test header type property."""
        doc = Document(create_test_docx_with_header_footer())
        header = doc.headers[0]
        assert header.type == "default"

    def test_header_paragraphs_property(self):
        """Test header paragraphs property."""
        doc = Document(create_test_docx_with_header_footer(header_text="Header Text"))
        header = doc.headers[0]
        assert len(header.paragraphs) >= 1
        assert header.paragraphs[0].text == "Header Text"

    def test_header_contains(self):
        """Test header contains method."""
        doc = Document(create_test_docx_with_header_footer(header_text="Header Content"))
        header = doc.headers[0]
        assert header.contains("Content")
        assert not header.contains("Missing")

    def test_header_contains_case_insensitive(self):
        """Test header contains with case insensitive search."""
        doc = Document(create_test_docx_with_header_footer(header_text="HEADER CONTENT"))
        header = doc.headers[0]
        assert header.contains("content", case_sensitive=False)
        assert not header.contains("content", case_sensitive=True)

    def test_header_repr(self):
        """Test header string representation."""
        doc = Document(create_test_docx_with_header_footer(header_text="Short"))
        header = doc.headers[0]
        repr_str = repr(header)
        assert "Header" in repr_str
        assert "default" in repr_str


class TestFooterModel:
    """Tests for Footer model class."""

    def test_footer_text_property(self):
        """Test footer text property."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Sample Footer"))
        footer = doc.footers[0]
        assert footer.text == "Sample Footer"

    def test_footer_type_property(self):
        """Test footer type property."""
        doc = Document(create_test_docx_with_header_footer())
        footer = doc.footers[0]
        assert footer.type == "default"

    def test_footer_paragraphs_property(self):
        """Test footer paragraphs property."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Footer Text"))
        footer = doc.footers[0]
        assert len(footer.paragraphs) >= 1
        assert footer.paragraphs[0].text == "Footer Text"

    def test_footer_contains(self):
        """Test footer contains method."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Footer Content"))
        footer = doc.footers[0]
        assert footer.contains("Content")
        assert not footer.contains("Missing")

    def test_footer_contains_case_insensitive(self):
        """Test footer contains with case insensitive search."""
        doc = Document(create_test_docx_with_header_footer(footer_text="FOOTER CONTENT"))
        footer = doc.footers[0]
        assert footer.contains("content", case_sensitive=False)
        assert not footer.contains("content", case_sensitive=True)

    def test_footer_repr(self):
        """Test footer string representation."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Short"))
        footer = doc.footers[0]
        repr_str = repr(footer)
        assert "Footer" in repr_str
        assert "default" in repr_str


class TestReplaceInHeader:
    """Tests for replace_in_header method."""

    def test_replace_in_header_basic(self):
        """Test basic header replacement."""
        doc = Document(create_test_docx_with_header_footer(header_text="Draft Document"))
        doc.replace_in_header("Draft", "Final")

        # Verify XML contains tracked changes
        header = doc.headers[0]
        xml_str = etree.tostring(header.element, encoding="unicode")
        assert "del" in xml_str.lower() or "ins" in xml_str.lower()

    def test_replace_in_header_with_author(self):
        """Test header replacement with custom author."""
        doc = Document(create_test_docx_with_header_footer(header_text="Original Text"))
        doc.replace_in_header("Original", "Updated", author="John Doe")

        # Verify the replacement was made
        header = doc.headers[0]
        xml_str = etree.tostring(header.element, encoding="unicode")
        assert "John Doe" in xml_str or "ins" in xml_str.lower()

    def test_replace_in_header_not_found(self):
        """Test header replacement with text not found."""
        from python_docx_redline import TextNotFoundError

        doc = Document(create_test_docx_with_header_footer(header_text="Header Text"))
        with pytest.raises(TextNotFoundError):
            doc.replace_in_header("Nonexistent", "New")

    def test_replace_in_header_no_header(self):
        """Test header replacement when no header exists."""
        doc = Document(create_simple_docx_no_headers_footers())
        with pytest.raises(ValueError, match="No header"):
            doc.replace_in_header("Text", "New")

    def test_replace_in_header_specific_type(self):
        """Test header replacement in specific header type."""
        doc = Document(
            create_test_docx_with_header_footer(
                header_text="Default Header",
                first_header_text="First Page",
            )
        )
        doc.replace_in_header("First", "Title", header_type="first")

        # Verify only first header was modified
        first_header = None
        for h in doc.headers:
            if h.type == "first":
                first_header = h
                break
        assert first_header is not None
        xml_str = etree.tostring(first_header.element, encoding="unicode")
        assert "del" in xml_str.lower() or "ins" in xml_str.lower()


class TestReplaceInFooter:
    """Tests for replace_in_footer method."""

    def test_replace_in_footer_basic(self):
        """Test basic footer replacement."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Page 1 of N"))
        doc.replace_in_footer("Page 1", "Page A")

        # Verify XML contains tracked changes
        footer = doc.footers[0]
        xml_str = etree.tostring(footer.element, encoding="unicode")
        assert "del" in xml_str.lower() or "ins" in xml_str.lower()

    def test_replace_in_footer_with_author(self):
        """Test footer replacement with custom author."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Confidential"))
        doc.replace_in_footer("Confidential", "Public", author="Jane Smith")

        # Verify the replacement was made
        footer = doc.footers[0]
        xml_str = etree.tostring(footer.element, encoding="unicode")
        assert "Jane Smith" in xml_str or "ins" in xml_str.lower()

    def test_replace_in_footer_not_found(self):
        """Test footer replacement with text not found."""
        from python_docx_redline import TextNotFoundError

        doc = Document(create_test_docx_with_header_footer(footer_text="Footer Text"))
        with pytest.raises(TextNotFoundError):
            doc.replace_in_footer("Nonexistent", "New")

    def test_replace_in_footer_no_footer(self):
        """Test footer replacement when no footer exists."""
        doc = Document(create_simple_docx_no_headers_footers())
        with pytest.raises(ValueError, match="No footer"):
            doc.replace_in_footer("Text", "New")


class TestInsertInHeader:
    """Tests for insert_in_header method."""

    def test_insert_in_header_after(self):
        """Test inserting text after anchor in header."""
        doc = Document(create_test_docx_with_header_footer(header_text="Document Title"))
        doc.insert_in_header(" - Final", after="Title")

        # Verify XML contains insertion
        header = doc.headers[0]
        xml_str = etree.tostring(header.element, encoding="unicode")
        assert "ins" in xml_str.lower()

    def test_insert_in_header_before(self):
        """Test inserting text before anchor in header."""
        doc = Document(create_test_docx_with_header_footer(header_text="Document Title"))
        doc.insert_in_header("Draft ", before="Document")

        # Verify XML contains insertion
        header = doc.headers[0]
        xml_str = etree.tostring(header.element, encoding="unicode")
        assert "ins" in xml_str.lower()

    def test_insert_in_header_both_after_and_before(self):
        """Test error when both after and before specified."""
        doc = Document(create_test_docx_with_header_footer(header_text="Header"))
        with pytest.raises(ValueError, match="Cannot specify both"):
            doc.insert_in_header("Text", after="Header", before="Header")

    def test_insert_in_header_neither_after_nor_before(self):
        """Test error when neither after nor before specified."""
        doc = Document(create_test_docx_with_header_footer(header_text="Header"))
        with pytest.raises(ValueError, match="Must specify"):
            doc.insert_in_header("Text")

    def test_insert_in_header_not_found(self):
        """Test insert in header when anchor text not found."""
        from python_docx_redline import TextNotFoundError

        doc = Document(create_test_docx_with_header_footer(header_text="Header"))
        with pytest.raises(TextNotFoundError):
            doc.insert_in_header("Text", after="Nonexistent")

    def test_insert_in_header_no_header(self):
        """Test insert in header when no header exists."""
        doc = Document(create_simple_docx_no_headers_footers())
        with pytest.raises(ValueError, match="No header"):
            doc.insert_in_header("Text", after="Anchor")


class TestInsertInFooter:
    """Tests for insert_in_footer method."""

    def test_insert_in_footer_after(self):
        """Test inserting text after anchor in footer."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Page Number"))
        doc.insert_in_footer(" of Total", after="Number")

        # Verify XML contains insertion
        footer = doc.footers[0]
        xml_str = etree.tostring(footer.element, encoding="unicode")
        assert "ins" in xml_str.lower()

    def test_insert_in_footer_before(self):
        """Test inserting text before anchor in footer."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Footer Content"))
        doc.insert_in_footer("© 2024 ", before="Footer")

        # Verify XML contains insertion
        footer = doc.footers[0]
        xml_str = etree.tostring(footer.element, encoding="unicode")
        assert "ins" in xml_str.lower()

    def test_insert_in_footer_both_after_and_before(self):
        """Test error when both after and before specified."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Footer"))
        with pytest.raises(ValueError, match="Cannot specify both"):
            doc.insert_in_footer("Text", after="Footer", before="Footer")

    def test_insert_in_footer_neither_after_nor_before(self):
        """Test error when neither after nor before specified."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Footer"))
        with pytest.raises(ValueError, match="Must specify"):
            doc.insert_in_footer("Text")

    def test_insert_in_footer_not_found(self):
        """Test insert in footer when anchor text not found."""
        from python_docx_redline import TextNotFoundError

        doc = Document(create_test_docx_with_header_footer(footer_text="Footer"))
        with pytest.raises(TextNotFoundError):
            doc.insert_in_footer("Text", after="Nonexistent")


class TestHeaderFooterPersistence:
    """Tests for header/footer persistence across save/load."""

    def test_header_persists_after_save(self):
        """Test that headers persist after save and reload."""
        doc = Document(create_test_docx_with_header_footer(header_text="Original Header"))

        # Modify header
        doc.replace_in_header("Original", "Modified")

        # Save and reload
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "modified.docx"
            doc.save(output_path)

            reloaded = Document(output_path)
            headers = reloaded.headers
            assert len(headers) >= 1

    def test_footer_persists_after_save(self):
        """Test that footers persist after save and reload."""
        doc = Document(create_test_docx_with_header_footer(footer_text="Original Footer"))

        # Modify footer
        doc.replace_in_footer("Original", "Modified")

        # Save and reload
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "modified.docx"
            doc.save(output_path)

            reloaded = Document(output_path)
            footers = reloaded.footers
            assert len(footers) >= 1


class TestHeaderFooterTypes:
    """Tests for HeaderFooterType enum."""

    def test_header_footer_type_values(self):
        """Test HeaderFooterType enum values."""
        from python_docx_redline import HeaderFooterType

        assert HeaderFooterType.DEFAULT.value == "default"
        assert HeaderFooterType.FIRST.value == "first"
        assert HeaderFooterType.EVEN.value == "even"


class TestHeaderFooterImports:
    """Tests for correct imports."""

    def test_header_import(self):
        """Test Header class can be imported."""
        from python_docx_redline import Header

        assert Header is not None

    def test_footer_import(self):
        """Test Footer class can be imported."""
        from python_docx_redline import Footer

        assert Footer is not None

    def test_header_footer_type_import(self):
        """Test HeaderFooterType enum can be imported."""
        from python_docx_redline import HeaderFooterType

        assert HeaderFooterType is not None
