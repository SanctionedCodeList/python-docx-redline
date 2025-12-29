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

            # hyperlinks property currently raises NotImplementedError
            # Once implemented, this should return empty list
            with pytest.raises(NotImplementedError):
                _ = doc.hyperlinks

        finally:
            doc_path.unlink()

    def test_hyperlinks_returns_list(self) -> None:
        """Test that hyperlinks property returns a list when implemented."""
        doc_path = create_document_with_hyperlinks()
        try:
            doc = Document(doc_path)

            # hyperlinks property currently raises NotImplementedError
            with pytest.raises(NotImplementedError):
                _ = doc.hyperlinks

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

            # hyperlinks property currently raises NotImplementedError
            # This test documents expected behavior once implemented
            with pytest.raises(NotImplementedError):
                _ = doc.hyperlinks

            # Future implementation should return list with is_external flag
            # distinguishing internal (anchor) from external (URL) links

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
