"""
Tests for bookmark and hyperlink extraction and management.

These tests verify:
- Bookmark extraction from document XML
- Hyperlink extraction (internal and external)
- Bidirectional reference tracking
- Broken reference detection
- Bookmark management functions (add, rename)
- YAML output for bookmarks and links
"""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from python_docx_redline import Document
from python_docx_redline.accessibility import (
    AccessibilityTree,
    BookmarkRegistry,
    LinkType,
    ViewMode,
    add_bookmark,
    rename_bookmark,
)

# Test XML documents

DOCUMENT_WITH_BOOKMARKS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="Introduction"/>
      <w:r>
        <w:t>Introduction section text.</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Regular paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="Definitions"/>
      <w:r>
        <w:t>Definitions section.</w:t>
      </w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_INTERNAL_LINKS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="TargetSection"/>
      <w:r>
        <w:t>Target section content.</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:r>
        <w:t>See also </w:t>
      </w:r>
      <w:hyperlink w:anchor="TargetSection">
        <w:r>
          <w:t>the target section</w:t>
        </w:r>
      </w:hyperlink>
      <w:r>
        <w:t> for more details.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_EXTERNAL_LINKS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Visit </w:t>
      </w:r>
      <w:hyperlink r:id="rId1">
        <w:r>
          <w:t>our website</w:t>
        </w:r>
      </w:hyperlink>
      <w:r>
        <w:t> for more information.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:hyperlink r:id="rId2">
        <w:r>
          <w:t>Documentation</w:t>
        </w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_BROKEN_LINKS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="ValidBookmark"/>
      <w:r>
        <w:t>Valid bookmark.</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:hyperlink w:anchor="NonExistentBookmark">
        <w:r>
          <w:t>Broken link to missing bookmark</w:t>
        </w:r>
      </w:hyperlink>
    </w:p>
    <w:p>
      <w:hyperlink r:id="rId99">
        <w:r>
          <w:t>Broken external link</w:t>
        </w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_SPAN_BOOKMARK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="SpanningBookmark"/>
      <w:r>
        <w:t>First paragraph of span.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph of span.</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:r>
        <w:t>After the bookmark.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_INTERNAL_BOOKMARKS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="_GoBack"/>
      <w:bookmarkEnd w:id="0"/>
      <w:r>
        <w:t>Document with internal bookmark.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="UserBookmark"/>
      <w:r>
        <w:t>User bookmark content.</w:t>
      </w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
  </w:body>
</w:document>"""


MINIMAL_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Simple paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


def create_registry_from_xml(
    xml_content: str,
    relationships: dict[str, str] | None = None,
) -> BookmarkRegistry:
    """Create a BookmarkRegistry from raw XML content."""
    root = etree.fromstring(xml_content.encode("utf-8"))
    return BookmarkRegistry.from_xml(root, relationships=relationships)


def create_tree_from_xml(
    xml_content: str,
    view_mode: ViewMode | None = None,
) -> AccessibilityTree:
    """Create an AccessibilityTree from raw XML content."""
    root = etree.fromstring(xml_content.encode("utf-8"))
    return AccessibilityTree.from_xml(root, view_mode=view_mode)


def create_test_docx(content: str) -> Path:
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


class TestBookmarkExtraction:
    """Tests for bookmark extraction."""

    def test_extract_bookmarks_basic(self) -> None:
        """Test basic bookmark extraction."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BOOKMARKS)

        assert len(registry.bookmarks) == 2
        assert "Introduction" in registry.bookmarks
        assert "Definitions" in registry.bookmarks

    def test_bookmark_has_correct_ref(self) -> None:
        """Test that bookmarks have correct refs."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BOOKMARKS)

        intro = registry.bookmarks["Introduction"]
        assert intro.ref == "bk:Introduction"

        defs = registry.bookmarks["Definitions"]
        assert defs.ref == "bk:Definitions"

    def test_bookmark_has_location(self) -> None:
        """Test that bookmarks have correct location refs."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BOOKMARKS)

        intro = registry.bookmarks["Introduction"]
        assert intro.location == "p:0"

        defs = registry.bookmarks["Definitions"]
        assert defs.location == "p:2"

    def test_bookmark_has_text_preview(self) -> None:
        """Test that bookmarks extract text preview."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BOOKMARKS)

        intro = registry.bookmarks["Introduction"]
        assert "Introduction section text" in intro.text_preview

    def test_bookmark_id_stored(self) -> None:
        """Test that bookmark ID is stored."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BOOKMARKS)

        intro = registry.bookmarks["Introduction"]
        assert intro.bookmark_id == "0"

    def test_internal_bookmarks_skipped(self) -> None:
        """Test that internal Word bookmarks (starting with _) are skipped."""
        registry = create_registry_from_xml(DOCUMENT_WITH_INTERNAL_BOOKMARKS)

        # Should only have UserBookmark, not _GoBack
        assert len(registry.bookmarks) == 1
        assert "UserBookmark" in registry.bookmarks
        assert "_GoBack" not in registry.bookmarks

    def test_span_bookmark_has_end_location(self) -> None:
        """Test that span bookmarks have end location."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SPAN_BOOKMARK)

        span = registry.bookmarks["SpanningBookmark"]
        assert span.location == "p:0"
        assert span.span_end_location == "p:1"

    def test_get_bookmark_by_name(self) -> None:
        """Test getting a bookmark by name."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BOOKMARKS)

        bookmark = registry.get_bookmark("Introduction")
        assert bookmark is not None
        assert bookmark.name == "Introduction"

        missing = registry.get_bookmark("NonExistent")
        assert missing is None


class TestHyperlinkExtraction:
    """Tests for hyperlink extraction."""

    def test_extract_internal_links(self) -> None:
        """Test extraction of internal links to bookmarks."""
        registry = create_registry_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        assert len(registry.hyperlinks) == 1

        link = registry.hyperlinks[0]
        assert link.link_type == LinkType.INTERNAL
        assert link.anchor == "TargetSection"
        assert link.text == "the target section"

    def test_internal_link_has_target_location(self) -> None:
        """Test that internal links resolve target location."""
        registry = create_registry_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        link = registry.hyperlinks[0]
        assert link.target_location == "p:0"
        assert not link.is_broken

    def test_extract_external_links(self) -> None:
        """Test extraction of external links."""
        relationships = {
            "rId1": "https://example.com",
            "rId2": "https://docs.example.com",
        }
        registry = create_registry_from_xml(DOCUMENT_WITH_EXTERNAL_LINKS, relationships)

        assert len(registry.hyperlinks) == 2

        link1 = registry.hyperlinks[0]
        assert link1.link_type == LinkType.EXTERNAL
        assert link1.target == "https://example.com"
        assert link1.text == "our website"
        assert not link1.is_broken

    def test_hyperlink_refs(self) -> None:
        """Test that hyperlinks get sequential refs."""
        relationships = {
            "rId1": "https://example.com",
            "rId2": "https://docs.example.com",
        }
        registry = create_registry_from_xml(DOCUMENT_WITH_EXTERNAL_LINKS, relationships)

        assert registry.hyperlinks[0].ref == "lnk:0"
        assert registry.hyperlinks[1].ref == "lnk:1"

    def test_hyperlink_from_location(self) -> None:
        """Test that hyperlinks have correct from_location."""
        registry = create_registry_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        link = registry.hyperlinks[0]
        assert link.from_location == "p:1"

    def test_get_internal_links(self) -> None:
        """Test filtering for internal links only."""
        relationships = {"rId1": "https://example.com"}

        # Create a document with both internal and external links
        doc = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="Section1"/>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:hyperlink w:anchor="Section1">
        <w:r><w:t>Internal</w:t></w:r>
      </w:hyperlink>
    </w:p>
    <w:p>
      <w:hyperlink r:id="rId1">
        <w:r><w:t>External</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>"""

        registry = create_registry_from_xml(doc, relationships)

        internal = registry.get_internal_links()
        external = registry.get_external_links()

        assert len(internal) == 1
        assert len(external) == 1
        assert internal[0].text == "Internal"
        assert external[0].text == "External"


class TestBidirectionalTracking:
    """Tests for bidirectional reference tracking."""

    def test_bookmark_tracks_referencing_links(self) -> None:
        """Test that bookmarks track which links reference them."""
        registry = create_registry_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        bookmark = registry.bookmarks["TargetSection"]
        assert len(bookmark.referenced_by) == 1
        assert "lnk:0" in bookmark.referenced_by

    def test_orphan_bookmarks_detected(self) -> None:
        """Test detection of bookmarks with no references."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BOOKMARKS)

        orphans = registry.get_orphan_bookmarks()

        # Both bookmarks are orphans (no links pointing to them)
        assert len(orphans) == 2

    def test_referenced_bookmark_not_orphan(self) -> None:
        """Test that referenced bookmarks are not in orphan list."""
        registry = create_registry_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        orphans = registry.get_orphan_bookmarks()
        assert len(orphans) == 0


class TestBrokenLinkDetection:
    """Tests for broken link detection."""

    def test_detect_broken_internal_link(self) -> None:
        """Test detection of internal link to missing bookmark."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BROKEN_LINKS)

        broken = registry.get_broken_links()

        # Should have 2 broken links
        assert len(broken) == 2

        # Find the internal broken link
        internal_broken = [link for link in broken if link.link_type == LinkType.INTERNAL]
        assert len(internal_broken) == 1
        assert internal_broken[0].anchor == "NonExistentBookmark"
        assert internal_broken[0].is_broken
        assert internal_broken[0].error == "Bookmark not found"

    def test_detect_broken_external_link(self) -> None:
        """Test detection of external link with missing relationship."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BROKEN_LINKS)

        broken = registry.get_broken_links()

        external_broken = [link for link in broken if link.link_type == LinkType.EXTERNAL]
        assert len(external_broken) == 1
        assert external_broken[0].relationship_id == "rId99"
        assert external_broken[0].is_broken
        assert external_broken[0].error == "Relationship not found"


class TestReferenceValidation:
    """Tests for reference validation."""

    def test_validate_valid_document(self) -> None:
        """Test validation of document with valid references."""
        registry = create_registry_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        result = registry.validate_references()

        assert result.is_valid
        assert len(result.broken_links) == 0

    def test_validate_document_with_broken_links(self) -> None:
        """Test validation of document with broken links."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BROKEN_LINKS)

        result = registry.validate_references()

        assert not result.is_valid
        assert len(result.broken_links) == 2
        assert "NonExistentBookmark" in result.missing_bookmarks

    def test_validation_includes_orphan_bookmarks(self) -> None:
        """Test that validation reports orphan bookmarks."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BOOKMARKS)

        result = registry.validate_references()

        # Document is valid (no broken links) but has orphan bookmarks
        assert result.is_valid
        assert len(result.orphan_bookmarks) == 2

    def test_validation_warnings(self) -> None:
        """Test that validation includes appropriate warnings."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BROKEN_LINKS)

        result = registry.validate_references()

        assert len(result.warnings) >= 1
        assert any("broken link" in w.lower() for w in result.warnings)


class TestAccessibilityTreeBookmarks:
    """Tests for bookmarks in AccessibilityTree."""

    def test_tree_has_bookmark_stats(self) -> None:
        """Test that tree stats include bookmark count."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BOOKMARKS)

        assert tree.stats.bookmarks == 2

    def test_tree_has_hyperlink_stats(self) -> None:
        """Test that tree stats include hyperlink count."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        assert tree.stats.hyperlinks == 1

    def test_tree_bookmarks_property(self) -> None:
        """Test accessing bookmarks via tree property."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BOOKMARKS)

        bookmarks = tree.bookmarks
        assert len(bookmarks) == 2
        assert "Introduction" in bookmarks

    def test_tree_hyperlinks_property(self) -> None:
        """Test accessing hyperlinks via tree property."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        links = tree.hyperlinks
        assert len(links) == 1

    def test_tree_get_bookmark(self) -> None:
        """Test getting bookmark via tree method."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BOOKMARKS)

        bookmark = tree.get_bookmark("Introduction")
        assert bookmark is not None
        assert bookmark.name == "Introduction"

    def test_tree_validate_references(self) -> None:
        """Test validate_references via tree method."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BROKEN_LINKS)

        result = tree.validate_references()
        assert not result.is_valid


class TestYamlBookmarkOutput:
    """Tests for YAML output of bookmarks and links."""

    def test_yaml_includes_bookmark_stats(self) -> None:
        """Test that YAML output includes bookmark stats."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BOOKMARKS)

        yaml = tree.to_yaml()

        assert "bookmarks: 2" in yaml

    def test_yaml_includes_bookmarks_section(self) -> None:
        """Test that YAML includes bookmarks section."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BOOKMARKS)

        yaml = tree.to_yaml()

        assert "\nbookmarks:" in yaml
        assert "- ref: bk:Introduction" in yaml
        assert "name: Introduction" in yaml
        assert "location: p:0" in yaml

    def test_yaml_includes_bookmark_text_preview(self) -> None:
        """Test that YAML includes bookmark text preview."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BOOKMARKS)

        yaml = tree.to_yaml()

        assert "text_preview:" in yaml

    def test_yaml_includes_links_section(self) -> None:
        """Test that YAML includes links section."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        yaml = tree.to_yaml()

        assert "\nlinks:" in yaml
        assert "internal:" in yaml
        assert "- ref: lnk:0" in yaml

    def test_yaml_internal_link_format(self) -> None:
        """Test format of internal links in YAML."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        yaml = tree.to_yaml()

        assert "from: p:1" in yaml
        assert "to: bk:TargetSection" in yaml

    def test_yaml_external_link_format(self) -> None:
        """Test format of external links in YAML."""
        root = etree.fromstring(DOCUMENT_WITH_EXTERNAL_LINKS.encode("utf-8"))
        relationships = {
            "rId1": "https://example.com",
            "rId2": "https://docs.example.com",
        }
        registry = BookmarkRegistry.from_xml(root, relationships=relationships)

        yaml_dict = registry.to_yaml_dict()

        assert "links" in yaml_dict
        assert "external" in yaml_dict["links"]
        assert len(yaml_dict["links"]["external"]) == 2

    def test_yaml_broken_links_section(self) -> None:
        """Test that broken links appear in YAML."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BROKEN_LINKS)

        yaml = tree.to_yaml()

        assert "broken:" in yaml

    def test_yaml_no_bookmarks_section_when_empty(self) -> None:
        """Test that bookmarks section is omitted when no bookmarks."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        yaml = tree.to_yaml()

        assert "\nbookmarks:" not in yaml

    def test_yaml_referenced_by_section(self) -> None:
        """Test that referenced_by appears for linked bookmarks."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        yaml = tree.to_yaml()

        assert "referenced_by:" in yaml
        assert "- lnk:0" in yaml


class TestAddBookmark:
    """Tests for add_bookmark function."""

    def test_add_bookmark_basic(self) -> None:
        """Test adding a bookmark to a paragraph."""
        root = etree.fromstring(MINIMAL_DOCUMENT_XML.encode("utf-8"))

        bookmark = add_bookmark(root, "NewBookmark", "p:0")

        assert bookmark.name == "NewBookmark"
        assert bookmark.ref == "bk:NewBookmark"
        assert bookmark.location == "p:0"

    def test_add_bookmark_creates_xml_elements(self) -> None:
        """Test that add_bookmark creates proper XML elements."""
        root = etree.fromstring(MINIMAL_DOCUMENT_XML.encode("utf-8"))

        add_bookmark(root, "TestBookmark", "p:0")

        # Verify elements were created
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        bookmark_start = root.find(f".//{ns}bookmarkStart[@{ns}name='TestBookmark']")
        bookmark_end = root.find(f".//{ns}bookmarkEnd")

        assert bookmark_start is not None
        assert bookmark_end is not None

    def test_add_bookmark_duplicate_name_raises(self) -> None:
        """Test that adding duplicate bookmark name raises error."""
        root = etree.fromstring(DOCUMENT_WITH_BOOKMARKS.encode("utf-8"))

        try:
            add_bookmark(root, "Introduction", "p:1")
            assert False, "Should have raised ValueError"
        except ValueError as e:
            assert "already exists" in str(e)

    def test_add_bookmark_invalid_ref_raises(self) -> None:
        """Test that invalid paragraph ref raises error."""
        root = etree.fromstring(MINIMAL_DOCUMENT_XML.encode("utf-8"))

        try:
            add_bookmark(root, "Test", "p:99")
            assert False, "Should have raised ValueError"
        except ValueError as e:
            assert "not found" in str(e)


class TestRenameBookmark:
    """Tests for rename_bookmark function."""

    def test_rename_bookmark_basic(self) -> None:
        """Test basic bookmark renaming."""
        root = etree.fromstring(DOCUMENT_WITH_BOOKMARKS.encode("utf-8"))

        rename_bookmark(root, "Introduction", "Intro")

        # Verify the rename
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        old = root.find(f".//{ns}bookmarkStart[@{ns}name='Introduction']")
        new = root.find(f".//{ns}bookmarkStart[@{ns}name='Intro']")

        assert old is None
        assert new is not None

    def test_rename_bookmark_updates_references(self) -> None:
        """Test that renaming updates hyperlink references."""
        root = etree.fromstring(DOCUMENT_WITH_INTERNAL_LINKS.encode("utf-8"))

        count = rename_bookmark(root, "TargetSection", "NewTarget")

        # Verify hyperlink was updated
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        hyperlink = root.find(f".//{ns}hyperlink")

        assert hyperlink is not None
        assert hyperlink.get(f"{ns}anchor") == "NewTarget"
        assert count == 1

    def test_rename_bookmark_not_found_raises(self) -> None:
        """Test that renaming non-existent bookmark raises error."""
        root = etree.fromstring(DOCUMENT_WITH_BOOKMARKS.encode("utf-8"))

        try:
            rename_bookmark(root, "NonExistent", "NewName")
            assert False, "Should have raised ValueError"
        except ValueError as e:
            assert "not found" in str(e)

    def test_rename_bookmark_to_existing_raises(self) -> None:
        """Test that renaming to existing name raises error."""
        root = etree.fromstring(DOCUMENT_WITH_BOOKMARKS.encode("utf-8"))

        try:
            rename_bookmark(root, "Introduction", "Definitions")
            assert False, "Should have raised ValueError"
        except ValueError as e:
            assert "already exists" in str(e)


class TestBookmarkRegistryYamlDict:
    """Tests for BookmarkRegistry.to_yaml_dict()."""

    def test_yaml_dict_structure(self) -> None:
        """Test the structure of YAML dict output."""
        registry = create_registry_from_xml(DOCUMENT_WITH_INTERNAL_LINKS)

        result = registry.to_yaml_dict()

        assert "bookmarks" in result
        assert "links" in result
        assert "internal" in result["links"]

    def test_yaml_dict_empty_when_no_bookmarks(self) -> None:
        """Test that empty document produces empty dict."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        result = registry.to_yaml_dict()

        assert "bookmarks" not in result
        assert "links" not in result


class TestDocumentIntegration:
    """Integration tests with Document class."""

    def test_tree_from_document_includes_bookmarks(self) -> None:
        """Test that tree built from Document includes bookmarks."""
        docx_path = create_test_docx(DOCUMENT_WITH_BOOKMARKS)

        try:
            doc = Document(docx_path)
            tree = AccessibilityTree.from_document(doc)

            assert tree.stats.bookmarks == 2
            assert len(tree.bookmarks) == 2

        finally:
            docx_path.unlink()

    def test_yaml_from_document(self) -> None:
        """Test YAML output from Document."""
        docx_path = create_test_docx(DOCUMENT_WITH_BOOKMARKS)

        try:
            doc = Document(docx_path)
            tree = AccessibilityTree.from_document(doc)
            yaml = tree.to_yaml()

            assert "bookmarks:" in yaml
            assert "bk:Introduction" in yaml

        finally:
            docx_path.unlink()
