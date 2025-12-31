"""
Tests for Phase 1 hyperlink functionality.

Tests the insert_hyperlink() method for external URLs and internal bookmarks,
the doc.hyperlinks property for reading existing hyperlinks, and error handling.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from python_docx_redline import Document
from python_docx_redline.errors import AmbiguousTextError, TextNotFoundError
from python_docx_redline.operations.hyperlinks import HyperlinkInfo

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def create_simple_document() -> Path:
    """Create a minimal test document."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>This is a test document with some text.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>For more information, please visit our website.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>Section 2: Additional details are available.</w:t></w:r>
</w:p>
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

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)

    return doc_path


def create_document_with_duplicate_text() -> Path:
    """Create a document with duplicate text for AmbiguousTextError testing."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>Click here for details.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>Click here for more info.</w:t></w:r>
</w:p>
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

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)

    return doc_path


def create_document_with_bookmark() -> Path:
    """Create a document with a bookmark for testing internal hyperlinks."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>This is a test document with some text.</w:t></w:r>
</w:p>
<w:p>
  <w:bookmarkStart w:id="0" w:name="TestBookmark"/>
  <w:r><w:t>This paragraph contains a bookmark.</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
<w:p>
  <w:r><w:t>More content after the bookmark.</w:t></w:r>
</w:p>
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

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)

    return doc_path


def create_document_with_hyperlinks() -> Path:
    """Create a document that already contains hyperlinks."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:r><w:t>Visit our </w:t></w:r>
  <w:hyperlink r:id="rId2" w:tooltip="Company Website">
    <w:r>
      <w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>
      <w:t>website</w:t>
    </w:r>
  </w:hyperlink>
  <w:r><w:t> for more info.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>See the </w:t></w:r>
  <w:hyperlink w:anchor="DefinitionsSection">
    <w:r>
      <w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>
      <w:t>definitions</w:t>
    </w:r>
  </w:hyperlink>
  <w:r><w:t> section.</w:t></w:r>
</w:p>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestInsertHyperlink:
    """Tests for Document.insert_hyperlink() method."""

    def test_insert_external_hyperlink_after(self) -> None:
        """Test inserting external hyperlink after anchor text."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink(
                url="https://example.com",
                text="Example Link",
                after="visit our website",
            )

            # Should return relationship ID for external links
            assert r_id is not None
            assert r_id.startswith("rId")

            # Verify hyperlink element was created
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            # Verify r:id attribute is set
            hyperlink = hyperlinks[0]
            r_id_attr = hyperlink.get(f"{{{REL_NS}}}id")
            assert r_id_attr == r_id

            # Verify hyperlink contains the display text
            text_elements = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elements)
            assert "Example Link" in text

        finally:
            doc_path.unlink()

    def test_insert_external_hyperlink_before(self) -> None:
        """Test inserting external hyperlink before anchor text."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink(
                url="https://example.com",
                text="[Link]",
                before="For more information",
            )

            # Should return relationship ID
            assert r_id is not None

            # Verify hyperlink element was created
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            # Verify display text
            hyperlink = hyperlinks[0]
            text_elements = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elements)
            assert "[Link]" in text

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_with_tooltip(self) -> None:
        """Test inserting hyperlink with tooltip."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            doc.insert_hyperlink(
                url="https://example.com/docs",
                text="Documentation",
                after="test document",
                tooltip="Click to view documentation",
            )

            # Verify tooltip attribute
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]
            tooltip = hyperlink.get(f"{{{WORD_NS}}}tooltip")
            assert tooltip == "Click to view documentation"

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_creates_relationship(self) -> None:
        """Test that external hyperlink creates proper relationship entry."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink(
                url="https://test-url.com/page",
                text="Test Link",
                after="some text",
            )

            doc.save(output_path)

            # Read the relationship file from saved document
            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/document.xml.rels").decode("utf-8")

            # Parse and verify relationship
            rels_tree = etree.fromstring(rels_content.encode())
            relationships = rels_tree.findall(f".//{{{PKG_REL_NS}}}Relationship")

            # Find the hyperlink relationship
            hyperlink_rel = None
            for rel in relationships:
                if rel.get("Id") == r_id:
                    hyperlink_rel = rel
                    break

            assert hyperlink_rel is not None
            assert hyperlink_rel.get("Target") == "https://test-url.com/page"
            assert hyperlink_rel.get("TargetMode") == "External"
            assert "hyperlink" in hyperlink_rel.get("Type", "").lower()

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_hyperlink_applies_hyperlink_style(self) -> None:
        """Test that Hyperlink character style is applied."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            doc.insert_hyperlink(
                url="https://example.com",
                text="Styled Link",
                after="test document",
            )

            # Find hyperlink element
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            # Find run inside hyperlink
            hyperlink = hyperlinks[0]
            runs = hyperlink.findall(f".//{{{WORD_NS}}}r")
            assert len(runs) > 0

            # Check for rStyle with Hyperlink value
            run = runs[0]
            rpr = run.find(f"{{{WORD_NS}}}rPr")
            assert rpr is not None

            rstyle = rpr.find(f"{{{WORD_NS}}}rStyle")
            assert rstyle is not None
            assert rstyle.get(f"{{{WORD_NS}}}val") == "Hyperlink"

        finally:
            doc_path.unlink()

    def test_insert_internal_hyperlink_with_anchor(self) -> None:
        """Test inserting internal hyperlink with anchor parameter."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            # Insert internal link to bookmark
            r_id = doc.insert_hyperlink(
                anchor="DefinitionsSection",
                text="See Definitions",
                after="Additional details",
            )

            # Internal links should return None (no relationship ID)
            assert r_id is None

            # Verify hyperlink element
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]

            # Verify w:anchor attribute is set (not r:id)
            anchor_attr = hyperlink.get(f"{{{WORD_NS}}}anchor")
            assert anchor_attr == "DefinitionsSection"

            # Verify r:id is NOT set
            r_id_attr = hyperlink.get(f"{{{REL_NS}}}id")
            assert r_id_attr is None

        finally:
            doc_path.unlink()

    def test_insert_internal_hyperlink_no_relationship_created(self) -> None:
        """Test that internal hyperlinks don't create relationship entries."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            doc.insert_hyperlink(
                anchor="BookmarkName",
                text="Internal Link",
                after="some text",
            )

            doc.save(output_path)

            # Check relationships file
            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/document.xml.rels").decode("utf-8")

            # Parse and check for hyperlink relationships
            rels_tree = etree.fromstring(rels_content.encode())
            relationships = rels_tree.findall(f".//{{{PKG_REL_NS}}}Relationship")

            # Count hyperlink relationships
            hyperlink_count = sum(
                1 for rel in relationships if "hyperlink" in rel.get("Type", "").lower()
            )

            # Should be no hyperlink relationships for internal links
            assert hyperlink_count == 0

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_hyperlink_with_scope(self) -> None:
        """Test inserting hyperlink with scope parameter to limit search."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            # Use scope to limit search to paragraph containing "Section 2"
            r_id = doc.insert_hyperlink(
                url="https://example.com",
                text="Scoped Link",
                after="Additional details",
                scope="Section 2",  # Limit to paragraph containing this text
            )

            assert r_id is not None

            # Verify hyperlink was created
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_preserves_whitespace(self) -> None:
        """Test that hyperlink text with leading/trailing spaces is preserved."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            doc.insert_hyperlink(
                url="https://example.com",
                text=" spaced text ",
                after="test document",
            )

            # Find the text element
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            hyperlink = hyperlinks[0]
            text_elem = hyperlink.find(f".//{{{WORD_NS}}}t")

            # Check xml:space="preserve" attribute
            space_attr = text_elem.get("{http://www.w3.org/XML/1998/namespace}space")
            assert space_attr == "preserve"
            assert text_elem.text == " spaced text "

        finally:
            doc_path.unlink()

    def test_insert_multiple_hyperlinks(self) -> None:
        """Test inserting multiple hyperlinks in the same document."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            r_id1 = doc.insert_hyperlink(
                url="https://first-link.com",
                text="First",
                after="test document",
            )

            r_id2 = doc.insert_hyperlink(
                url="https://second-link.com",
                text="Second",
                after="visit our website",
            )

            # Both should have unique relationship IDs
            assert r_id1 is not None
            assert r_id2 is not None
            assert r_id1 != r_id2

            # Verify both hyperlinks exist
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 2

        finally:
            doc_path.unlink()


class TestHyperlinksProperty:
    """Tests for Document.hyperlinks property."""

    def test_hyperlinks_empty_document(self) -> None:
        """Test hyperlinks property returns empty list for document without hyperlinks."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            # hyperlinks property returns empty list for document without hyperlinks
            hyperlinks = doc.hyperlinks
            assert hyperlinks == []

        finally:
            doc_path.unlink()

    def test_hyperlinks_returns_list(self) -> None:
        """Test that hyperlinks property returns a list when implemented."""
        doc_path = create_document_with_hyperlinks()
        try:
            doc = Document(doc_path)

            # hyperlinks property returns a list of HyperlinkInfo objects
            hyperlinks = doc.hyperlinks
            assert isinstance(hyperlinks, list)
            # The document created by create_document_with_hyperlinks should have hyperlinks
            assert len(hyperlinks) > 0

        finally:
            doc_path.unlink()


class TestHyperlinkErrors:
    """Tests for hyperlink error handling."""

    def test_anchor_not_found_raises_error(self) -> None:
        """Test TextNotFoundError when anchor text is not found."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(TextNotFoundError) as exc_info:
                doc.insert_hyperlink(
                    url="https://example.com",
                    text="Link",
                    after="nonexistent anchor text",
                )

            assert "nonexistent anchor text" in str(exc_info.value)

        finally:
            doc_path.unlink()

    def test_neither_url_nor_anchor_raises_error(self) -> None:
        """Test ValueError when neither url nor anchor is provided."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.insert_hyperlink(
                    text="Orphan Link",
                    after="some text",
                )

            assert "url" in str(exc_info.value).lower() or "anchor" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_both_url_and_anchor_raises_error(self) -> None:
        """Test ValueError when both url and anchor are provided."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.insert_hyperlink(
                    url="https://example.com",
                    anchor="BookmarkName",
                    text="Conflicting Link",
                    after="some text",
                )

            assert "url" in str(exc_info.value).lower() or "anchor" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_neither_after_nor_before_raises_error(self) -> None:
        """Test ValueError when neither after nor before is specified."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.insert_hyperlink(
                    url="https://example.com",
                    text="No Position Link",
                )

            assert "after" in str(exc_info.value).lower() or "before" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_both_after_and_before_raises_error(self) -> None:
        """Test ValueError when both after and before are specified."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.insert_hyperlink(
                    url="https://example.com",
                    text="Dual Position Link",
                    after="some text",
                    before="other text",
                )

            assert "after" in str(exc_info.value).lower() or "before" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_ambiguous_text_raises_error(self) -> None:
        """Test AmbiguousTextError when anchor text appears multiple times."""
        doc_path = create_document_with_duplicate_text()
        try:
            doc = Document(doc_path)

            with pytest.raises(AmbiguousTextError) as exc_info:
                doc.insert_hyperlink(
                    url="https://example.com",
                    text="Link",
                    after="Click here",
                )

            assert "Click here" in str(exc_info.value)
            # Should mention number of occurrences
            assert "2" in str(exc_info.value)

        finally:
            doc_path.unlink()


class TestHyperlinkInfoDataclass:
    """Tests for the HyperlinkInfo dataclass."""

    def test_hyperlink_info_creation(self) -> None:
        """Test HyperlinkInfo can be created with all fields."""
        info = HyperlinkInfo(
            ref="lnk:1",
            text="Example Link",
            target="https://example.com",
            is_external=True,
            tooltip="Click here",
            location="body",
            r_id="rId5",
        )

        assert info.ref == "lnk:1"
        assert info.text == "Example Link"
        assert info.target == "https://example.com"
        assert info.is_external is True
        assert info.tooltip == "Click here"
        assert info.location == "body"
        assert info.r_id == "rId5"

    def test_hyperlink_info_defaults(self) -> None:
        """Test HyperlinkInfo has proper defaults."""
        info = HyperlinkInfo(
            ref="lnk:2",
            text="Internal Link",
            target="BookmarkName",
            is_external=False,
        )

        assert info.tooltip is None
        assert info.location == "body"
        assert info.r_id is None


class TestHyperlinkPersistence:
    """Tests for hyperlink persistence after save/reload."""

    def test_hyperlink_persists_after_save_reload(self) -> None:
        """Test that inserted hyperlink persists after save and reload."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Create and save
            doc = Document(doc_path)
            doc.insert_hyperlink(
                url="https://persistent-link.com",
                text="Persistent Link",
                after="test document",
                tooltip="Hover text",
            )
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)

            hyperlinks = list(doc2.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]

            # Verify r:id is set
            assert hyperlink.get(f"{{{REL_NS}}}id") is not None

            # Verify tooltip
            assert hyperlink.get(f"{{{WORD_NS}}}tooltip") == "Hover text"

            # Verify text
            text_elems = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elems)
            assert "Persistent Link" in text

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_internal_hyperlink_persists(self) -> None:
        """Test that internal hyperlink persists after save and reload."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.insert_hyperlink(
                anchor="MyBookmark",
                text="Jump to Section",
                after="test document",
            )
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)

            hyperlinks = list(doc2.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]

            # Verify w:anchor is set
            assert hyperlink.get(f"{{{WORD_NS}}}anchor") == "MyBookmark"

            # Verify r:id is NOT set
            assert hyperlink.get(f"{{{REL_NS}}}id") is None

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()


class TestInternalHyperlinks:
    """Tests for Phase 2 internal hyperlink functionality."""

    def test_insert_internal_hyperlink_with_anchor_parameter(self) -> None:
        """Test inserting hyperlink to bookmark using anchor parameter."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            # Insert internal link to bookmark
            r_id = doc.insert_hyperlink(
                anchor="SectionOne",
                text="Jump to Section One",
                after="test document",
            )

            # Internal links should return None (no relationship ID)
            assert r_id is None

            # Verify hyperlink element was created
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]

            # Verify display text
            text_elements = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elements)
            assert "Jump to Section One" in text

        finally:
            doc_path.unlink()

    def test_internal_hyperlink_has_anchor_attribute_not_rid(self) -> None:
        """Test that internal hyperlinks use w:anchor attribute, not r:id."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            doc.insert_hyperlink(
                anchor="MyBookmark",
                text="Internal Link",
                after="some text",
            )

            # Find the hyperlink
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]

            # Verify w:anchor attribute is set
            anchor_attr = hyperlink.get(f"{{{WORD_NS}}}anchor")
            assert anchor_attr == "MyBookmark"

            # Verify r:id is NOT set
            r_id_attr = hyperlink.get(f"{{{REL_NS}}}id")
            assert r_id_attr is None

        finally:
            doc_path.unlink()

    def test_internal_hyperlink_missing_bookmark_warning(self) -> None:
        """Test warning issued when bookmark doesn't exist."""
        import warnings

        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                doc.insert_hyperlink(
                    anchor="NonExistentBookmark",
                    text="Broken Link",
                    after="some text",
                )

                # Should have issued exactly one warning
                assert len(w) == 1
                assert issubclass(w[0].category, UserWarning)
                assert "NonExistentBookmark" in str(w[0].message)
                assert "does not exist" in str(w[0].message)

            # Link should still be created (just broken)
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

        finally:
            doc_path.unlink()

    def test_internal_hyperlink_existing_bookmark_no_warning(self) -> None:
        """Test no warning when bookmark exists."""
        import warnings

        doc_path = create_document_with_bookmark()
        try:
            doc = Document(doc_path)

            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                doc.insert_hyperlink(
                    anchor="TestBookmark",
                    text="Valid Link",
                    after="some text",
                )

                # Should NOT have issued any warning
                assert len(w) == 0

            # Link should be created successfully
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

        finally:
            doc_path.unlink()

    def test_internal_hyperlink_no_relationship_in_saved_file(self) -> None:
        """Test that internal hyperlinks don't create relationship entries after save."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            doc.insert_hyperlink(
                anchor="InternalTarget",
                text="Internal Link",
                after="some text",
            )

            doc.save(output_path)

            # Check relationships file
            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/document.xml.rels").decode("utf-8")

            # Parse and check for hyperlink relationships
            rels_tree = etree.fromstring(rels_content.encode())
            relationships = rels_tree.findall(f".//{{{PKG_REL_NS}}}Relationship")

            # Count hyperlink relationships
            hyperlink_count = sum(
                1 for rel in relationships if "hyperlink" in rel.get("Type", "").lower()
            )

            # Should be no hyperlink relationships for internal links
            assert hyperlink_count == 0

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_hyperlinks_property_identifies_internal_vs_external(self) -> None:
        """Test that hyperlinks property correctly identifies internal vs external links."""
        doc_path = create_document_with_hyperlinks()
        try:
            doc = Document(doc_path)

            # hyperlinks property returns list with is_external flag
            # distinguishing internal (anchor) from external (URL) links
            hyperlinks = doc.hyperlinks
            assert isinstance(hyperlinks, list)
            assert len(hyperlinks) > 0

            # Each HyperlinkInfo should have an is_external attribute
            for link in hyperlinks:
                assert hasattr(link, "is_external")
                assert isinstance(link.is_external, bool)

        finally:
            doc_path.unlink()

    def test_internal_hyperlink_with_tooltip(self) -> None:
        """Test inserting internal hyperlink with tooltip."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            doc.insert_hyperlink(
                anchor="MySection",
                text="See Section",
                after="test document",
                tooltip="Jump to My Section",
            )

            # Verify tooltip attribute
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]
            tooltip = hyperlink.get(f"{{{WORD_NS}}}tooltip")
            assert tooltip == "Jump to My Section"

        finally:
            doc_path.unlink()

    def test_internal_hyperlink_applies_style(self) -> None:
        """Test that internal hyperlinks also get Hyperlink character style."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)

            doc.insert_hyperlink(
                anchor="BookmarkTarget",
                text="Styled Internal Link",
                after="test document",
            )

            # Find hyperlink element
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            # Find run inside hyperlink
            hyperlink = hyperlinks[0]
            runs = hyperlink.findall(f".//{{{WORD_NS}}}r")
            assert len(runs) > 0

            # Check for rStyle with Hyperlink value
            run = runs[0]
            rpr = run.find(f"{{{WORD_NS}}}rPr")
            assert rpr is not None

            rstyle = rpr.find(f"{{{WORD_NS}}}rStyle")
            assert rstyle is not None
            assert rstyle.get(f"{{{WORD_NS}}}val") == "Hyperlink"

        finally:
            doc_path.unlink()


class TestHyperlinkStyleCreation:
    """Tests for Hyperlink style creation in styles.xml."""

    def test_hyperlink_style_created_if_missing(self) -> None:
        """Test that Hyperlink style is created in styles.xml if it doesn't exist."""
        doc_path = create_simple_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.insert_hyperlink(
                url="https://example.com",
                text="Styled Link",
                after="test document",
            )
            doc.save(output_path)

            # Check styles.xml exists and contains Hyperlink style
            with zipfile.ZipFile(output_path, "r") as docx:
                assert "word/styles.xml" in docx.namelist()
                styles_content = docx.read("word/styles.xml").decode("utf-8")

            # Parse and find Hyperlink style
            styles_tree = etree.fromstring(styles_content.encode())
            styles = styles_tree.findall(f".//{{{WORD_NS}}}style")

            hyperlink_style = None
            for style in styles:
                style_id = style.get(f"{{{WORD_NS}}}styleId")
                if style_id == "Hyperlink":
                    hyperlink_style = style
                    break

            assert hyperlink_style is not None
            assert hyperlink_style.get(f"{{{WORD_NS}}}type") == "character"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()


def create_document_with_footnote() -> Path:
    """Create a document with an existing footnote for testing."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:xml="http://www.w3.org/XML/1998/namespace">
<w:body>
<w:p>
  <w:r><w:t>This document has a footnote reference here</w:t></w:r>
  <w:r>
    <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
    <w:footnoteReference w:id="1"/>
  </w:r>
  <w:r><w:t xml:space="preserve"> and continues after.</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

    footnotes_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:xml="http://www.w3.org/XML/1998/namespace">
  <w:footnote w:id="-1" w:type="separator">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="0" w:type="continuationSeparator">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p>
      <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
      <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
      <w:r><w:t xml:space="preserve"> </w:t></w:r>
      <w:r><w:t>See the reference for more details.</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/footnotes.xml", footnotes_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_endnote() -> Path:
    """Create a document with an existing endnote for testing."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:xml="http://www.w3.org/XML/1998/namespace">
<w:body>
<w:p>
  <w:r><w:t>This document has an endnote reference here</w:t></w:r>
  <w:r>
    <w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>
    <w:endnoteReference w:id="1"/>
  </w:r>
  <w:r><w:t xml:space="preserve"> and continues after.</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

    endnotes_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:xml="http://www.w3.org/XML/1998/namespace">
  <w:endnote w:id="-1" w:type="separator">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:id="0" w:type="continuationSeparator">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:id="1">
    <w:p>
      <w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>
      <w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>
      <w:r><w:t xml:space="preserve"> </w:t></w:r>
      <w:r><w:t>See the bibliography for sources.</w:t></w:r>
    </w:p>
  </w:endnote>
</w:endnotes>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/endnotes.xml", endnotes_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestInsertHyperlinkInFootnote:
    """Tests for insert_hyperlink_in_footnote method."""

    def test_insert_external_hyperlink_in_footnote(self) -> None:
        """Test inserting external hyperlink after text in a footnote."""
        doc_path = create_document_with_footnote()
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_footnote(
                note_id=1,
                url="https://example.com/reference",
                text="online reference",
                after="See the",
            )

            # Should return relationship ID
            assert r_id is not None
            assert r_id.startswith("rId")

            # Verify hyperlink was inserted in the footnote
            footnotes_path = doc._temp_dir / "word" / "footnotes.xml"
            tree = etree.parse(str(footnotes_path))
            hyperlinks = tree.findall(f".//{{{WORD_NS}}}hyperlink")
            assert len(hyperlinks) == 1

            # Verify the hyperlink has the relationship ID
            hyperlink = hyperlinks[0]
            assert hyperlink.get(f"{{{REL_NS}}}id") == r_id

            # Verify the display text
            text_elems = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elems)
            assert "online reference" in text

        finally:
            doc_path.unlink()

    def test_insert_internal_hyperlink_in_footnote(self) -> None:
        """Test inserting internal hyperlink in a footnote."""
        doc_path = create_document_with_footnote()
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_footnote(
                note_id=1,
                anchor="DefinitionsSection",
                text="definitions",
                after="See the",
            )

            # Internal links should return None (no rId)
            assert r_id is None

            # Verify hyperlink was inserted
            footnotes_path = doc._temp_dir / "word" / "footnotes.xml"
            tree = etree.parse(str(footnotes_path))
            hyperlinks = tree.findall(f".//{{{WORD_NS}}}hyperlink")
            assert len(hyperlinks) == 1

            # Verify w:anchor attribute
            hyperlink = hyperlinks[0]
            assert hyperlink.get(f"{{{WORD_NS}}}anchor") == "DefinitionsSection"

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_footnote_before(self) -> None:
        """Test inserting hyperlink before text in a footnote."""
        doc_path = create_document_with_footnote()
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_footnote(
                note_id=1,
                url="https://example.com",
                text="Link:",
                before="See the",
            )

            assert r_id is not None

            # Verify hyperlink was inserted
            footnotes_path = doc._temp_dir / "word" / "footnotes.xml"
            tree = etree.parse(str(footnotes_path))
            hyperlinks = tree.findall(f".//{{{WORD_NS}}}hyperlink")
            assert len(hyperlinks) == 1

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_footnote_creates_relationship_file(self) -> None:
        """Test that inserting hyperlink creates footnotes.xml.rels."""
        doc_path = create_document_with_footnote()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            doc.insert_hyperlink_in_footnote(
                note_id=1,
                url="https://test-url.com",
                text="Test Link",
                after="reference",
            )

            doc.save(output_path)

            # Verify the footnotes rels file exists and contains the hyperlink
            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/footnotes.xml.rels").decode("utf-8")

            rels_tree = etree.fromstring(rels_content.encode())
            relationships = rels_tree.findall(f".//{{{PKG_REL_NS}}}Relationship")

            # Find the hyperlink relationship
            hyperlink_rel = None
            for rel in relationships:
                if "hyperlink" in rel.get("Type", "").lower():
                    hyperlink_rel = rel
                    break

            assert hyperlink_rel is not None
            assert hyperlink_rel.get("Target") == "https://test-url.com"
            assert hyperlink_rel.get("TargetMode") == "External"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_hyperlink_in_footnote_not_found_raises_error(self) -> None:
        """Test NoteNotFoundError when footnote doesn't exist."""
        from python_docx_redline.errors import NoteNotFoundError

        doc_path = create_document_with_footnote()
        try:
            doc = Document(doc_path)

            with pytest.raises(NoteNotFoundError) as exc_info:
                doc.insert_hyperlink_in_footnote(
                    note_id=999,
                    url="https://example.com",
                    text="Link",
                    after="text",
                )

            assert "999" in str(exc_info.value)

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_footnote_text_not_found(self) -> None:
        """Test TextNotFoundError when anchor text not in footnote."""
        doc_path = create_document_with_footnote()
        try:
            doc = Document(doc_path)

            with pytest.raises(TextNotFoundError) as exc_info:
                doc.insert_hyperlink_in_footnote(
                    note_id=1,
                    url="https://example.com",
                    text="Link",
                    after="nonexistent text",
                )

            assert "nonexistent text" in str(exc_info.value)

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_footnote_validation_errors(self) -> None:
        """Test validation errors for invalid parameters."""
        doc_path = create_document_with_footnote()
        try:
            doc = Document(doc_path)

            # Neither url nor anchor
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_footnote(
                    note_id=1,
                    text="Link",
                    after="text",
                )

            # Both url and anchor
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_footnote(
                    note_id=1,
                    url="https://example.com",
                    anchor="Bookmark",
                    text="Link",
                    after="text",
                )

            # Neither after nor before
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_footnote(
                    note_id=1,
                    url="https://example.com",
                    text="Link",
                )

            # Both after and before
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_footnote(
                    note_id=1,
                    url="https://example.com",
                    text="Link",
                    after="text",
                    before="other",
                )

        finally:
            doc_path.unlink()


class TestInsertHyperlinkInEndnote:
    """Tests for insert_hyperlink_in_endnote method."""

    def test_insert_external_hyperlink_in_endnote(self) -> None:
        """Test inserting external hyperlink after text in an endnote."""
        doc_path = create_document_with_endnote()
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_endnote(
                note_id=1,
                url="https://example.com/bibliography",
                text="full bibliography",
                after="See the",
            )

            # Should return relationship ID
            assert r_id is not None
            assert r_id.startswith("rId")

            # Verify hyperlink was inserted in the endnote
            endnotes_path = doc._temp_dir / "word" / "endnotes.xml"
            tree = etree.parse(str(endnotes_path))
            hyperlinks = tree.findall(f".//{{{WORD_NS}}}hyperlink")
            assert len(hyperlinks) == 1

            # Verify the hyperlink has the relationship ID
            hyperlink = hyperlinks[0]
            assert hyperlink.get(f"{{{REL_NS}}}id") == r_id

        finally:
            doc_path.unlink()

    def test_insert_internal_hyperlink_in_endnote(self) -> None:
        """Test inserting internal hyperlink in an endnote."""
        doc_path = create_document_with_endnote()
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_endnote(
                note_id=1,
                anchor="BibliographySection",
                text="bibliography section",
                after="See the",
            )

            # Internal links should return None
            assert r_id is None

            # Verify hyperlink was inserted
            endnotes_path = doc._temp_dir / "word" / "endnotes.xml"
            tree = etree.parse(str(endnotes_path))
            hyperlinks = tree.findall(f".//{{{WORD_NS}}}hyperlink")
            assert len(hyperlinks) == 1

            # Verify w:anchor attribute
            hyperlink = hyperlinks[0]
            assert hyperlink.get(f"{{{WORD_NS}}}anchor") == "BibliographySection"

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_endnote_creates_relationship_file(self) -> None:
        """Test that inserting hyperlink creates endnotes.xml.rels."""
        doc_path = create_document_with_endnote()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            doc.insert_hyperlink_in_endnote(
                note_id=1,
                url="https://test-url.com/sources",
                text="Test Link",
                after="bibliography",
            )

            doc.save(output_path)

            # Verify the endnotes rels file exists and contains the hyperlink
            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/endnotes.xml.rels").decode("utf-8")

            rels_tree = etree.fromstring(rels_content.encode())
            relationships = rels_tree.findall(f".//{{{PKG_REL_NS}}}Relationship")

            # Find the hyperlink relationship
            hyperlink_rel = None
            for rel in relationships:
                if "hyperlink" in rel.get("Type", "").lower():
                    hyperlink_rel = rel
                    break

            assert hyperlink_rel is not None
            assert hyperlink_rel.get("Target") == "https://test-url.com/sources"
            assert hyperlink_rel.get("TargetMode") == "External"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_hyperlink_in_endnote_not_found_raises_error(self) -> None:
        """Test NoteNotFoundError when endnote doesn't exist."""
        from python_docx_redline.errors import NoteNotFoundError

        doc_path = create_document_with_endnote()
        try:
            doc = Document(doc_path)

            with pytest.raises(NoteNotFoundError) as exc_info:
                doc.insert_hyperlink_in_endnote(
                    note_id=999,
                    url="https://example.com",
                    text="Link",
                    after="text",
                )

            assert "999" in str(exc_info.value)

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_endnote_text_not_found(self) -> None:
        """Test TextNotFoundError when anchor text not in endnote."""
        doc_path = create_document_with_endnote()
        try:
            doc = Document(doc_path)

            with pytest.raises(TextNotFoundError) as exc_info:
                doc.insert_hyperlink_in_endnote(
                    note_id=1,
                    url="https://example.com",
                    text="Link",
                    after="nonexistent text",
                )

            assert "nonexistent text" in str(exc_info.value)

        finally:
            doc_path.unlink()


def create_document_with_header_footer(
    header_text: str = "Test Header",
    footer_text: str = "Test Footer",
    first_header_text: str | None = None,
    first_footer_text: str | None = None,
) -> Path:
    """Create a minimal .docx file with headers and footers for hyperlink testing."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a test document with headers and footers.</w:t>
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

    header1_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:t>{header_text}</w:t>
    </w:r>
  </w:p>
</w:hdr>"""

    footer1_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:t>{footer_text}</w:t>
    </w:r>
  </w:p>
</w:ftr>"""

    header2_xml = (
        f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:t>{first_header_text}</w:t>
    </w:r>
  </w:p>
</w:hdr>"""
        if first_header_text
        else None
    )

    footer2_xml = (
        f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:t>{first_footer_text}</w:t>
    </w:r>
  </w:p>
</w:ftr>"""
        if first_footer_text
        else None
    )

    content_types_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  {"<Override PartName='/word/header2.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'/>" if first_header_text else ""}
  {"<Override PartName='/word/footer2.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'/>" if first_footer_text else ""}
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  {"<Relationship Id='rId8' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/header' Target='header2.xml'/>" if first_header_text else ""}
  {"<Relationship Id='rId9' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer' Target='footer2.xml'/>" if first_footer_text else ""}
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)
        docx.writestr("word/header1.xml", header1_xml)
        docx.writestr("word/footer1.xml", footer1_xml)
        if header2_xml:
            docx.writestr("word/header2.xml", header2_xml)
        if footer2_xml:
            docx.writestr("word/footer2.xml", footer2_xml)

    return doc_path


class TestHeaderHyperlinks:
    """Tests for insert_hyperlink_in_header method."""

    def test_insert_external_hyperlink_in_header(self) -> None:
        """Test inserting external hyperlink in header default type."""
        doc_path = create_document_with_header_footer(header_text="Visit our website")
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_header(
                url="https://example.com",
                text="Example",
                after="Visit our",
                header_type="default",
            )

            # Should return relationship ID
            assert r_id is not None
            assert r_id.startswith("rId")

            # Verify hyperlink was inserted in header
            headers = doc.headers
            assert len(headers) >= 1
            header = headers[0]
            xml_str = etree.tostring(header.element, encoding="unicode")
            assert "hyperlink" in xml_str.lower()

        finally:
            doc_path.unlink()

    def test_insert_internal_hyperlink_in_header(self) -> None:
        """Test inserting internal hyperlink (anchor) in header."""
        doc_path = create_document_with_header_footer(header_text="Go to Section One")
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_header(
                anchor="SectionOne",
                text="Jump",
                after="Go to",
                header_type="default",
            )

            # Internal links should return None
            assert r_id is None

            # Verify hyperlink was inserted with w:anchor attribute
            headers = doc.headers
            header = headers[0]
            hyperlinks = list(header.element.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]
            assert hyperlink.get(f"{{{WORD_NS}}}anchor") == "SectionOne"

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_first_header(self) -> None:
        """Test inserting hyperlink in first page header."""
        doc_path = create_document_with_header_footer(
            header_text="Default Header",
            first_header_text="First Page Title",
        )
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_header(
                url="https://first-page.com",
                text="Link",
                after="First Page",
                header_type="first",
            )

            assert r_id is not None

            # Verify hyperlink was inserted in the first header
            first_header = None
            for h in doc.headers:
                if h.type == "first":
                    first_header = h
                    break
            assert first_header is not None
            xml_str = etree.tostring(first_header.element, encoding="unicode")
            assert "hyperlink" in xml_str.lower()

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_header_creates_relationship_file(self) -> None:
        """Test that inserting hyperlink creates header1.xml.rels."""
        doc_path = create_document_with_header_footer(header_text="Click here")
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            doc.insert_hyperlink_in_header(
                url="https://test-url.com",
                text="Test Link",
                after="Click",
            )

            doc.save(output_path)

            # Verify the header rels file exists and contains the hyperlink
            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/header1.xml.rels").decode("utf-8")

            rels_tree = etree.fromstring(rels_content.encode())
            relationships = rels_tree.findall(f".//{{{PKG_REL_NS}}}Relationship")

            # Find the hyperlink relationship
            hyperlink_rel = None
            for rel in relationships:
                if "hyperlink" in rel.get("Type", "").lower():
                    hyperlink_rel = rel
                    break

            assert hyperlink_rel is not None
            assert hyperlink_rel.get("Target") == "https://test-url.com"
            assert hyperlink_rel.get("TargetMode") == "External"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_hyperlink_in_header_text_not_found(self) -> None:
        """Test TextNotFoundError when anchor text not in header."""
        doc_path = create_document_with_header_footer(header_text="Test Header")
        try:
            doc = Document(doc_path)

            with pytest.raises(TextNotFoundError) as exc_info:
                doc.insert_hyperlink_in_header(
                    url="https://example.com",
                    text="Link",
                    after="nonexistent text",
                )

            assert "nonexistent text" in str(exc_info.value)

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_header_invalid_type(self) -> None:
        """Test ValueError when invalid header_type specified."""
        doc_path = create_document_with_header_footer(header_text="Test Header")
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.insert_hyperlink_in_header(
                    url="https://example.com",
                    text="Link",
                    after="Test",
                    header_type="invalid",
                )

            assert "invalid" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_header_validation_errors(self) -> None:
        """Test validation errors for invalid parameters."""
        doc_path = create_document_with_header_footer(header_text="Test Header")
        try:
            doc = Document(doc_path)

            # Neither url nor anchor
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_header(
                    text="Link",
                    after="Test",
                )

            # Both url and anchor
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_header(
                    url="https://example.com",
                    anchor="Bookmark",
                    text="Link",
                    after="Test",
                )

            # Neither after nor before
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_header(
                    url="https://example.com",
                    text="Link",
                )

            # Both after and before
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_header(
                    url="https://example.com",
                    text="Link",
                    after="Test",
                    before="Header",
                )

        finally:
            doc_path.unlink()


class TestFooterHyperlinks:
    """Tests for insert_hyperlink_in_footer method."""

    def test_insert_external_hyperlink_in_footer(self) -> None:
        """Test inserting external hyperlink in footer default type."""
        doc_path = create_document_with_header_footer(footer_text="Contact us for more info")
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_footer(
                url="https://contact.example.com",
                text="contact form",
                after="Contact us",
                footer_type="default",
            )

            # Should return relationship ID
            assert r_id is not None
            assert r_id.startswith("rId")

            # Verify hyperlink was inserted in footer
            footers = doc.footers
            assert len(footers) >= 1
            footer = footers[0]
            xml_str = etree.tostring(footer.element, encoding="unicode")
            assert "hyperlink" in xml_str.lower()

        finally:
            doc_path.unlink()

    def test_insert_internal_hyperlink_in_footer(self) -> None:
        """Test inserting internal hyperlink (anchor) in footer."""
        doc_path = create_document_with_header_footer(footer_text="See the appendix for details")
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_footer(
                anchor="AppendixA",
                text="Appendix A",
                after="See the",
                footer_type="default",
            )

            # Internal links should return None
            assert r_id is None

            # Verify hyperlink was inserted with w:anchor attribute
            footers = doc.footers
            footer = footers[0]
            hyperlinks = list(footer.element.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]
            assert hyperlink.get(f"{{{WORD_NS}}}anchor") == "AppendixA"

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_first_footer(self) -> None:
        """Test inserting hyperlink in first page footer."""
        doc_path = create_document_with_header_footer(
            footer_text="Default Footer",
            first_footer_text="Cover Page Footer",
        )
        try:
            doc = Document(doc_path)

            r_id = doc.insert_hyperlink_in_footer(
                url="https://cover-page.com",
                text="Link",
                after="Cover Page",
                footer_type="first",
            )

            assert r_id is not None

            # Verify hyperlink was inserted in the first footer
            first_footer = None
            for f in doc.footers:
                if f.type == "first":
                    first_footer = f
                    break
            assert first_footer is not None
            xml_str = etree.tostring(first_footer.element, encoding="unicode")
            assert "hyperlink" in xml_str.lower()

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_footer_creates_relationship_file(self) -> None:
        """Test that inserting hyperlink creates footer1.xml.rels."""
        doc_path = create_document_with_header_footer(footer_text="Click here")
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            doc.insert_hyperlink_in_footer(
                url="https://footer-link.com",
                text="Footer Link",
                after="Click",
            )

            doc.save(output_path)

            # Verify the footer rels file exists and contains the hyperlink
            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/footer1.xml.rels").decode("utf-8")

            rels_tree = etree.fromstring(rels_content.encode())
            relationships = rels_tree.findall(f".//{{{PKG_REL_NS}}}Relationship")

            # Find the hyperlink relationship
            hyperlink_rel = None
            for rel in relationships:
                if "hyperlink" in rel.get("Type", "").lower():
                    hyperlink_rel = rel
                    break

            assert hyperlink_rel is not None
            assert hyperlink_rel.get("Target") == "https://footer-link.com"
            assert hyperlink_rel.get("TargetMode") == "External"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_hyperlink_in_footer_text_not_found(self) -> None:
        """Test TextNotFoundError when anchor text not in footer."""
        doc_path = create_document_with_header_footer(footer_text="Test Footer")
        try:
            doc = Document(doc_path)

            with pytest.raises(TextNotFoundError) as exc_info:
                doc.insert_hyperlink_in_footer(
                    url="https://example.com",
                    text="Link",
                    after="nonexistent text",
                )

            assert "nonexistent text" in str(exc_info.value)

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_footer_invalid_type(self) -> None:
        """Test ValueError when invalid footer_type specified."""
        doc_path = create_document_with_header_footer(footer_text="Test Footer")
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.insert_hyperlink_in_footer(
                    url="https://example.com",
                    text="Link",
                    after="Test",
                    footer_type="invalid",
                )

            assert "invalid" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_insert_hyperlink_in_footer_validation_errors(self) -> None:
        """Test validation errors for invalid parameters."""
        doc_path = create_document_with_header_footer(footer_text="Test Footer")
        try:
            doc = Document(doc_path)

            # Neither url nor anchor
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_footer(
                    text="Link",
                    after="Test",
                )

            # Both url and anchor
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_footer(
                    url="https://example.com",
                    anchor="Bookmark",
                    text="Link",
                    after="Test",
                )

            # Neither after nor before
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_footer(
                    url="https://example.com",
                    text="Link",
                )

            # Both after and before
            with pytest.raises(ValueError):
                doc.insert_hyperlink_in_footer(
                    url="https://example.com",
                    text="Link",
                    after="Test",
                    before="Footer",
                )

        finally:
            doc_path.unlink()


# ==================== Phase 4 Tests: Edit and Remove Hyperlinks ====================


def create_document_with_external_hyperlink() -> Path:
    """Create a document with an external hyperlink for testing edit/remove."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:r><w:t>Click on this </w:t></w:r>
  <w:hyperlink r:id="rId2" w:tooltip="Original tooltip">
    <w:r>
      <w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>
      <w:t>link text</w:t>
    </w:r>
  </w:hyperlink>
  <w:r><w:t> to visit the site.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>More content here.</w:t></w:r>
</w:p>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://original-url.com" TargetMode="External"/>
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_internal_hyperlink() -> Path:
    """Create a document with an internal hyperlink for testing edit/remove."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:r><w:t>See the </w:t></w:r>
  <w:hyperlink w:anchor="OriginalBookmark">
    <w:r>
      <w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>
      <w:t>definitions section</w:t>
    </w:r>
  </w:hyperlink>
  <w:r><w:t> for details.</w:t></w:r>
</w:p>
<w:p>
  <w:bookmarkStart w:id="0" w:name="OriginalBookmark"/>
  <w:r><w:t>Definitions start here.</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
<w:p>
  <w:bookmarkStart w:id="1" w:name="NewBookmark"/>
  <w:r><w:t>New section here.</w:t></w:r>
  <w:bookmarkEnd w:id="1"/>
</w:p>
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

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)

    return doc_path


class TestEditHyperlinkUrl:
    """Tests for Document.edit_hyperlink_url() method."""

    def test_edit_hyperlink_url_success(self) -> None:
        """Test successfully changing external link URL."""
        doc_path = create_document_with_external_hyperlink()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            # Edit the URL using relationship ID
            doc.edit_hyperlink_url("rId2", "https://new-url.com/page")

            # Save and verify the relationship was updated (skip validation for minimal test doc)
            doc.save(output_path, validate=False)

            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/document.xml.rels").decode("utf-8")

            rels_tree = etree.fromstring(rels_content.encode())
            relationships = rels_tree.findall(f".//{{{PKG_REL_NS}}}Relationship")

            hyperlink_rel = None
            for rel in relationships:
                if rel.get("Id") == "rId2":
                    hyperlink_rel = rel
                    break

            assert hyperlink_rel is not None
            assert hyperlink_rel.get("Target") == "https://new-url.com/page"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_edit_hyperlink_url_by_lnk_ref(self) -> None:
        """Test editing URL using lnk:N ref format."""
        doc_path = create_document_with_external_hyperlink()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            # Edit using lnk:0 (first hyperlink)
            doc.edit_hyperlink_url("lnk:0", "https://updated-link.com")

            doc.save(output_path, validate=False)

            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/document.xml.rels").decode("utf-8")

            rels_tree = etree.fromstring(rels_content.encode())
            relationships = rels_tree.findall(f".//{{{PKG_REL_NS}}}Relationship")

            for rel in relationships:
                if "hyperlink" in rel.get("Type", "").lower():
                    assert rel.get("Target") == "https://updated-link.com"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_edit_hyperlink_url_ref_not_found(self) -> None:
        """Test error when ref not found."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_url("rIdNonExistent", "https://new-url.com")

            assert "not found" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_url_on_internal_link_raises_error(self) -> None:
        """Test error when trying to edit URL of internal link."""
        doc_path = create_document_with_internal_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_url("lnk:0", "https://new-url.com")

            assert "internal" in str(exc_info.value).lower()
            assert "edit_hyperlink_anchor" in str(exc_info.value)

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_url_empty_url_raises_error(self) -> None:
        """Test error when new_url is empty."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_url("rId2", "")

            assert "empty" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()


class TestEditHyperlinkText:
    """Tests for Document.edit_hyperlink_text() method."""

    def test_edit_hyperlink_text_success(self) -> None:
        """Test successfully changing display text."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            doc.edit_hyperlink_text("lnk:0", "Updated Link Text")

            # Verify text was updated
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]
            text_elements = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elements)
            assert text == "Updated Link Text"

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_text_preserves_hyperlink_style(self) -> None:
        """Test that Hyperlink style is preserved after text edit."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            doc.edit_hyperlink_text("lnk:0", "New Text")

            # Verify Hyperlink style is still applied
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            hyperlink = hyperlinks[0]
            runs = hyperlink.findall(f".//{{{WORD_NS}}}r")

            assert len(runs) >= 1
            run = runs[0]
            rpr = run.find(f"{{{WORD_NS}}}rPr")
            assert rpr is not None

            rstyle = rpr.find(f"{{{WORD_NS}}}rStyle")
            assert rstyle is not None
            assert rstyle.get(f"{{{WORD_NS}}}val") == "Hyperlink"

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_text_persists_after_save_reload(self) -> None:
        """Test text change persists after save/reload."""
        doc_path = create_document_with_external_hyperlink()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.edit_hyperlink_text("lnk:0", "Persistent Text Change")
            doc.save(output_path, validate=False)

            # Reload and verify
            doc2 = Document(output_path)
            hyperlinks = list(doc2.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]
            text_elements = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elements)
            assert text == "Persistent Text Change"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_edit_hyperlink_text_ref_not_found(self) -> None:
        """Test error when ref not found."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_text("lnk:999", "New Text")

            assert "not found" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_text_empty_text_raises_error(self) -> None:
        """Test error when new_text is empty."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_text("lnk:0", "")

            assert "empty" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_text_preserves_whitespace(self) -> None:
        """Test that leading/trailing whitespace in text is preserved."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            doc.edit_hyperlink_text("lnk:0", " spaced text ")

            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            hyperlink = hyperlinks[0]
            t_elem = hyperlink.find(f".//{{{WORD_NS}}}t")

            # Check xml:space="preserve" attribute
            space_attr = t_elem.get("{http://www.w3.org/XML/1998/namespace}space")
            assert space_attr == "preserve"
            assert t_elem.text == " spaced text "

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_text_on_internal_link(self) -> None:
        """Test editing text on internal hyperlink works."""
        doc_path = create_document_with_internal_hyperlink()

        try:
            doc = Document(doc_path)

            doc.edit_hyperlink_text("lnk:0", "Updated internal link text")

            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            hyperlink = hyperlinks[0]
            text_elements = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elements)
            assert text == "Updated internal link text"

        finally:
            doc_path.unlink()


class TestEditHyperlinkAnchor:
    """Tests for Document.edit_hyperlink_anchor() method."""

    def test_edit_hyperlink_anchor_success(self) -> None:
        """Test successfully changing internal link target."""
        doc_path = create_document_with_internal_hyperlink()

        try:
            doc = Document(doc_path)

            doc.edit_hyperlink_anchor("lnk:0", "NewBookmark")

            # Verify anchor was updated
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]
            anchor = hyperlink.get(f"{{{WORD_NS}}}anchor")
            assert anchor == "NewBookmark"

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_anchor_ref_not_found(self) -> None:
        """Test error when ref not found."""
        doc_path = create_document_with_internal_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_anchor("lnk:999", "SomeBookmark")

            assert "not found" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_anchor_on_external_link_raises_error(self) -> None:
        """Test error when trying to edit anchor of external link."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_anchor("lnk:0", "SomeBookmark")

            assert "external" in str(exc_info.value).lower()
            assert "edit_hyperlink_url" in str(exc_info.value)

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_anchor_empty_anchor_raises_error(self) -> None:
        """Test error when new_anchor is empty."""
        doc_path = create_document_with_internal_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_anchor("lnk:0", "")

            assert "empty" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_anchor_nonexistent_bookmark_warning(self) -> None:
        """Test warning when new bookmark doesn't exist."""
        import warnings

        doc_path = create_document_with_internal_hyperlink()

        try:
            doc = Document(doc_path)

            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                doc.edit_hyperlink_anchor("lnk:0", "NonExistentBookmark")

                # Should have issued a warning
                assert len(w) == 1
                assert issubclass(w[0].category, UserWarning)
                assert "NonExistentBookmark" in str(w[0].message)
                assert "does not exist" in str(w[0].message)

            # But the anchor should still be changed
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            hyperlink = hyperlinks[0]
            anchor = hyperlink.get(f"{{{WORD_NS}}}anchor")
            assert anchor == "NonExistentBookmark"

        finally:
            doc_path.unlink()

    def test_edit_hyperlink_anchor_persists_after_save_reload(self) -> None:
        """Test anchor change persists after save/reload."""
        import warnings

        doc_path = create_document_with_internal_hyperlink()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            # Suppress warning about existing bookmark
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                doc.edit_hyperlink_anchor("lnk:0", "NewBookmark")

            doc.save(output_path, validate=False)

            # Reload and verify
            doc2 = Document(output_path)
            hyperlinks = list(doc2.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]
            anchor = hyperlink.get(f"{{{WORD_NS}}}anchor")
            assert anchor == "NewBookmark"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()


class TestRemoveHyperlink:
    """Tests for Document.remove_hyperlink() method."""

    def test_remove_hyperlink_keep_text(self) -> None:
        """Test removing hyperlink while keeping the text."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            doc.remove_hyperlink("lnk:0", keep_text=True)

            # Hyperlink element should be gone
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 0

            # But the text should still exist in the document
            all_text = []
            for t_elem in doc.xml_root.iter(f"{{{WORD_NS}}}t"):
                if t_elem.text:
                    all_text.append(t_elem.text)

            full_text = "".join(all_text)
            assert "link text" in full_text

        finally:
            doc_path.unlink()

    def test_remove_hyperlink_keep_text_removes_style(self) -> None:
        """Test that removing hyperlink removes Hyperlink character style from text."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            doc.remove_hyperlink("lnk:0", keep_text=True)

            # Find all runs containing "link text"
            for run in doc.xml_root.iter(f"{{{WORD_NS}}}r"):
                t_elem = run.find(f"{{{WORD_NS}}}t")
                if t_elem is not None and t_elem.text and "link text" in t_elem.text:
                    # Check that Hyperlink style is removed
                    rpr = run.find(f"{{{WORD_NS}}}rPr")
                    if rpr is not None:
                        rstyle = rpr.find(f"{{{WORD_NS}}}rStyle")
                        if rstyle is not None:
                            # If there's a style, it shouldn't be "Hyperlink"
                            style_val = rstyle.get(f"{{{WORD_NS}}}val")
                            assert style_val != "Hyperlink"

        finally:
            doc_path.unlink()

    def test_remove_hyperlink_remove_text(self) -> None:
        """Test removing hyperlink and its text entirely."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            doc.remove_hyperlink("lnk:0", keep_text=False)

            # Hyperlink element should be gone
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 0

            # The text "link text" should also be gone
            all_text = []
            for t_elem in doc.xml_root.iter(f"{{{WORD_NS}}}t"):
                if t_elem.text:
                    all_text.append(t_elem.text)

            full_text = "".join(all_text)
            assert "link text" not in full_text

            # But surrounding text should remain
            assert "Click on this" in full_text
            assert "to visit the site" in full_text

        finally:
            doc_path.unlink()

    def test_remove_hyperlink_ref_not_found(self) -> None:
        """Test error when ref not found."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.remove_hyperlink("lnk:999")

            assert "not found" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_remove_hyperlink_persists_after_save_reload(self) -> None:
        """Test that removal persists after save/reload."""
        doc_path = create_document_with_external_hyperlink()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.remove_hyperlink("lnk:0", keep_text=True)
            doc.save(output_path, validate=False)

            # Reload and verify
            doc2 = Document(output_path)
            hyperlinks = list(doc2.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 0

            # Text should still be there
            all_text = []
            for t_elem in doc2.xml_root.iter(f"{{{WORD_NS}}}t"):
                if t_elem.text:
                    all_text.append(t_elem.text)

            full_text = "".join(all_text)
            assert "link text" in full_text

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_remove_internal_hyperlink_keep_text(self) -> None:
        """Test removing internal hyperlink while keeping text."""
        doc_path = create_document_with_internal_hyperlink()

        try:
            doc = Document(doc_path)

            doc.remove_hyperlink("lnk:0", keep_text=True)

            # Hyperlink element should be gone
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 0

            # But the text should still exist
            all_text = []
            for t_elem in doc.xml_root.iter(f"{{{WORD_NS}}}t"):
                if t_elem.text:
                    all_text.append(t_elem.text)

            full_text = "".join(all_text)
            assert "definitions section" in full_text

        finally:
            doc_path.unlink()

    def test_remove_internal_hyperlink_remove_text(self) -> None:
        """Test removing internal hyperlink and its text."""
        doc_path = create_document_with_internal_hyperlink()

        try:
            doc = Document(doc_path)

            doc.remove_hyperlink("lnk:0", keep_text=False)

            # Hyperlink element should be gone
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 0

            # The text should also be gone
            all_text = []
            for t_elem in doc.xml_root.iter(f"{{{WORD_NS}}}t"):
                if t_elem.text:
                    all_text.append(t_elem.text)

            full_text = "".join(all_text)
            assert "definitions section" not in full_text

            # But surrounding text should remain
            assert "See the" in full_text
            assert "for details" in full_text

        finally:
            doc_path.unlink()

    def test_remove_hyperlink_with_track_true(self) -> None:
        """Test removing hyperlink with tracked changes."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            # Remove with track=True and keep_text=False
            doc.remove_hyperlink("lnk:0", keep_text=False, track=True)

            # Hyperlink element should be gone
            hyperlinks = list(doc.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 0

            # A tracked deletion should be present
            deletions = list(doc.xml_root.iter(f"{{{WORD_NS}}}del"))
            assert len(deletions) >= 1

            # The deleted text should be "link text"
            del_text = []
            for del_elem in deletions:
                for t_elem in del_elem.iter(f"{{{WORD_NS}}}delText"):
                    if t_elem.text:
                        del_text.append(t_elem.text)

            assert "link text" in "".join(del_text)

        finally:
            doc_path.unlink()


class TestEditHyperlinkInvalidRef:
    """Tests for invalid ref format handling."""

    def test_invalid_ref_format_raises_error(self) -> None:
        """Test error for invalid ref format."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_url("invalid:format", "https://new-url.com")

            assert "invalid" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()

    def test_lnk_ref_with_non_integer_raises_error(self) -> None:
        """Test error when lnk:N has non-integer N."""
        doc_path = create_document_with_external_hyperlink()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError) as exc_info:
                doc.edit_hyperlink_text("lnk:abc", "New Text")

            assert "invalid" in str(exc_info.value).lower()

        finally:
            doc_path.unlink()


# === Tests for Issue #6: edit_hyperlink_text preserving style and markdown support ===


class TestEditHyperlinkTextIssue6:
    """Tests for Issue #6: edit_hyperlink_text preserving Hyperlink style and markdown."""

    def test_edit_hyperlink_text_untracked_preserves_style(self) -> None:
        """Test that untracked edit preserves Hyperlink rStyle."""
        doc_path = create_document_with_external_hyperlink()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            # Edit hyperlink text without tracking
            doc.edit_hyperlink_text("lnk:0", "Replaced Text", track=False)
            # Use validate=False since test document doesn't have styles.xml
            doc.save(output_path, validate=False)

            # Reload and verify
            doc2 = Document(output_path)
            hyperlinks = list(doc2.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            assert len(hyperlinks) == 1

            hyperlink = hyperlinks[0]
            runs = hyperlink.findall(f".//{{{WORD_NS}}}r")
            assert len(runs) >= 1

            # All runs should have Hyperlink rStyle
            for run in runs:
                rpr = run.find(f"{{{WORD_NS}}}rPr")
                assert rpr is not None, "Run should have rPr"
                rstyle = rpr.find(f"{{{WORD_NS}}}rStyle")
                assert rstyle is not None, "Run should have rStyle"
                assert rstyle.get(f"{{{WORD_NS}}}val") == "Hyperlink"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_edit_hyperlink_text_untracked_markdown_bold(self) -> None:
        """Test that untracked edit supports **bold** markdown."""
        doc_path = create_document_with_external_hyperlink()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            # Edit with bold markdown
            doc.edit_hyperlink_text("lnk:0", "Click **here**", track=False)
            # Use validate=False since test document doesn't have styles.xml
            doc.save(output_path, validate=False)

            # Reload and verify
            doc2 = Document(output_path)
            hyperlinks = list(doc2.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            hyperlink = hyperlinks[0]

            # Text should read correctly
            text_elems = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elems)
            assert text == "Click here"  # No markdown markers

            # Should have multiple runs
            runs = hyperlink.findall(f".//{{{WORD_NS}}}r")
            assert len(runs) >= 2

            # Find bold run
            bold_found = False
            for run in runs:
                rpr = run.find(f"{{{WORD_NS}}}rPr")
                if rpr is not None:
                    if rpr.find(f"{{{WORD_NS}}}b") is not None:
                        bold_found = True
                        # Should also have Hyperlink style
                        rstyle = rpr.find(f"{{{WORD_NS}}}rStyle")
                        assert rstyle is not None
                        assert rstyle.get(f"{{{WORD_NS}}}val") == "Hyperlink"
                        break

            assert bold_found, "Expected bold formatting in hyperlink"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_edit_hyperlink_text_untracked_markdown_italic(self) -> None:
        """Test that untracked edit supports *italic* markdown."""
        doc_path = create_document_with_external_hyperlink()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)

            # Edit with italic markdown
            doc.edit_hyperlink_text("lnk:0", "Click *here*", track=False)
            # Use validate=False since test document doesn't have styles.xml
            doc.save(output_path, validate=False)

            # Reload and verify
            doc2 = Document(output_path)
            hyperlinks = list(doc2.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            hyperlink = hyperlinks[0]

            # Find italic run
            runs = hyperlink.findall(f".//{{{WORD_NS}}}r")
            italic_found = False
            for run in runs:
                rpr = run.find(f"{{{WORD_NS}}}rPr")
                if rpr is not None:
                    if rpr.find(f"{{{WORD_NS}}}i") is not None:
                        italic_found = True
                        break

            assert italic_found, "Expected italic formatting in hyperlink"

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_edit_hyperlink_text_multi_run_to_single_preserves_style(self) -> None:
        """Test that editing multi-run hyperlink to single text preserves style."""
        # Create document with multi-run hyperlink
        doc_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:hyperlink r:id="rId5">
    <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>First</w:t></w:r>
    <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t> Part</w:t></w:r>
  </w:hyperlink>
</w:p>
</w:body>
</w:document>"""

        doc_path = Path(tempfile.mktemp(suffix=".docx"))
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        # Create docx file
        content_types = """<?xml version="1.0" encoding="UTF-8"?>
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
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"""

        with zipfile.ZipFile(doc_path, "w") as docx:
            docx.writestr("[Content_Types].xml", content_types)
            docx.writestr("_rels/.rels", root_rels)
            docx.writestr("word/document.xml", doc_xml)
            docx.writestr("word/_rels/document.xml.rels", doc_rels)

        try:
            doc = Document(doc_path)

            # Edit to single simple text
            doc.edit_hyperlink_text("lnk:0", "Simple Link", track=False)
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)
            hyperlinks = list(doc2.xml_root.iter(f"{{{WORD_NS}}}hyperlink"))
            hyperlink = hyperlinks[0]

            # Verify text
            text_elems = hyperlink.findall(f".//{{{WORD_NS}}}t")
            text = "".join(t.text or "" for t in text_elems)
            assert text == "Simple Link"

            # Verify Hyperlink style preserved
            runs = hyperlink.findall(f".//{{{WORD_NS}}}r")
            for run in runs:
                rpr = run.find(f"{{{WORD_NS}}}rPr")
                assert rpr is not None
                rstyle = rpr.find(f"{{{WORD_NS}}}rStyle")
                assert rstyle is not None
                assert rstyle.get(f"{{{WORD_NS}}}val") == "Hyperlink"

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)
