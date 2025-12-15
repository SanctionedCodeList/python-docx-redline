"""Tests for ContentTypeManager class."""

from pathlib import Path

import pytest

from python_docx_redline.content_types import (
    ContentTypeManager,
    ContentTypes,
)
from python_docx_redline.package import OOXMLPackage


@pytest.fixture
def sample_docx(tmp_path: Path) -> Path:
    """Create a minimal valid docx file for testing."""
    import zipfile

    docx_path = tmp_path / "test.docx"

    # Create minimal docx structure
    with zipfile.ZipFile(docx_path, "w") as zf:
        # [Content_Types].xml
        content_types = b"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""
        zf.writestr("[Content_Types].xml", content_types)

        # _rels/.rels
        root_rels = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
        zf.writestr("_rels/.rels", root_rels)

        # word/document.xml
        document = b"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Test document</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        zf.writestr("word/document.xml", document)

        # word/_rels/document.xml.rels
        doc_rels = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
        zf.writestr("word/_rels/document.xml.rels", doc_rels)

        # word/styles.xml (minimal)
        styles = b"""<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"""
        zf.writestr("word/styles.xml", styles)

    return docx_path


@pytest.fixture
def package(sample_docx: Path) -> OOXMLPackage:
    """Create an OOXMLPackage from the sample docx."""
    return OOXMLPackage.open(sample_docx)


class TestContentTypeManagerInit:
    """Test ContentTypeManager initialization."""

    def test_init(self, package: OOXMLPackage) -> None:
        """Test basic initialization."""
        ct_mgr = ContentTypeManager(package)
        assert ct_mgr._package is package
        assert ct_mgr.is_modified is False

    def test_loads_existing_content_types(self, package: OOXMLPackage) -> None:
        """Test that existing content types file is loaded."""
        ct_mgr = ContentTypeManager(package)
        # Getting a content type should load the file
        ct = ct_mgr.get_content_type("/word/document.xml")
        assert ct is not None
        assert "document.main+xml" in ct


class TestContentTypeManagerGet:
    """Test getting content types."""

    def test_get_existing_content_type(self, package: OOXMLPackage) -> None:
        """Test getting an existing content type."""
        ct_mgr = ContentTypeManager(package)
        ct = ct_mgr.get_content_type("/word/document.xml")
        assert ct == ContentTypes.DOCUMENT

    def test_get_nonexistent_content_type(self, package: OOXMLPackage) -> None:
        """Test getting a content type that doesn't exist."""
        ct_mgr = ContentTypeManager(package)
        ct = ct_mgr.get_content_type("/word/comments.xml")
        assert ct is None

    def test_has_override(self, package: OOXMLPackage) -> None:
        """Test checking if an override exists."""
        ct_mgr = ContentTypeManager(package)
        assert ct_mgr.has_override("/word/document.xml") is True
        assert ct_mgr.has_override("/word/comments.xml") is False


class TestContentTypeManagerAdd:
    """Test adding content type overrides."""

    def test_add_new_override(self, package: OOXMLPackage) -> None:
        """Test adding a new content type override."""
        ct_mgr = ContentTypeManager(package)
        result = ct_mgr.add_override("/word/comments.xml", ContentTypes.COMMENTS)

        assert result is True
        assert ct_mgr.is_modified is True
        assert ct_mgr.get_content_type("/word/comments.xml") == ContentTypes.COMMENTS

    def test_add_existing_override_returns_false(self, package: OOXMLPackage) -> None:
        """Test that adding an existing override returns False."""
        ct_mgr = ContentTypeManager(package)

        # First add
        result1 = ct_mgr.add_override("/word/comments.xml", ContentTypes.COMMENTS)
        assert result1 is True

        # Try to add again
        result2 = ct_mgr.add_override("/word/comments.xml", ContentTypes.COMMENTS)
        assert result2 is False

    def test_add_multiple_overrides(self, package: OOXMLPackage) -> None:
        """Test adding multiple different overrides."""
        ct_mgr = ContentTypeManager(package)

        ct_mgr.add_override("/word/comments.xml", ContentTypes.COMMENTS)
        ct_mgr.add_override("/word/footnotes.xml", ContentTypes.FOOTNOTES)
        ct_mgr.add_override("/word/endnotes.xml", ContentTypes.ENDNOTES)

        assert ct_mgr.get_content_type("/word/comments.xml") == ContentTypes.COMMENTS
        assert ct_mgr.get_content_type("/word/footnotes.xml") == ContentTypes.FOOTNOTES
        assert ct_mgr.get_content_type("/word/endnotes.xml") == ContentTypes.ENDNOTES


class TestContentTypeManagerRemove:
    """Test removing content type overrides."""

    def test_remove_existing_override(self, package: OOXMLPackage) -> None:
        """Test removing an existing override."""
        ct_mgr = ContentTypeManager(package)

        # First verify it exists
        assert ct_mgr.has_override("/word/styles.xml") is True

        # Remove it
        result = ct_mgr.remove_override("/word/styles.xml")
        assert result is True
        assert ct_mgr.is_modified is True
        assert ct_mgr.has_override("/word/styles.xml") is False

    def test_remove_nonexistent_override(self, package: OOXMLPackage) -> None:
        """Test removing an override that doesn't exist."""
        ct_mgr = ContentTypeManager(package)
        result = ct_mgr.remove_override("/word/comments.xml")
        assert result is False

    def test_remove_multiple_overrides(self, package: OOXMLPackage) -> None:
        """Test removing multiple overrides at once."""
        ct_mgr = ContentTypeManager(package)

        # Add some overrides first
        ct_mgr.add_override("/word/comments.xml", ContentTypes.COMMENTS)
        ct_mgr.add_override("/word/commentsExtended.xml", ContentTypes.COMMENTS_EXTENDED)
        ct_mgr.add_override("/word/footnotes.xml", ContentTypes.FOOTNOTES)

        # Remove comment-related overrides
        part_names = ["/word/comments.xml", "/word/commentsExtended.xml"]
        removed = ct_mgr.remove_overrides(part_names)

        assert removed == 2
        assert ct_mgr.has_override("/word/comments.xml") is False
        assert ct_mgr.has_override("/word/commentsExtended.xml") is False
        # Footnotes should still exist
        assert ct_mgr.has_override("/word/footnotes.xml") is True


class TestContentTypeManagerSave:
    """Test saving content type changes."""

    def test_save_persists_additions(self, package: OOXMLPackage) -> None:
        """Test that added overrides are persisted."""
        ct_mgr = ContentTypeManager(package)
        ct_mgr.add_override("/word/comments.xml", ContentTypes.COMMENTS)
        ct_mgr.save()

        # Create a new manager and verify persistence
        ct_mgr2 = ContentTypeManager(package)
        assert ct_mgr2.get_content_type("/word/comments.xml") == ContentTypes.COMMENTS

    def test_save_persists_removals(self, package: OOXMLPackage) -> None:
        """Test that removed overrides are persisted."""
        ct_mgr = ContentTypeManager(package)
        ct_mgr.remove_override("/word/styles.xml")
        ct_mgr.save()

        # Create a new manager and verify persistence
        ct_mgr2 = ContentTypeManager(package)
        assert ct_mgr2.has_override("/word/styles.xml") is False

    def test_save_does_nothing_if_not_modified(self, package: OOXMLPackage) -> None:
        """Test that save doesn't write if nothing changed."""
        ct_mgr = ContentTypeManager(package)

        # Just read, don't modify
        _ = ct_mgr.get_content_type("/word/document.xml")

        # is_modified should be False
        assert ct_mgr.is_modified is False

        # Save should be a no-op
        ct_mgr.save()


class TestContentTypes:
    """Test ContentTypes constants."""

    def test_common_content_types(self) -> None:
        """Test that common content type constants are defined."""
        assert ContentTypes.COMMENTS is not None
        assert "comments+xml" in ContentTypes.COMMENTS.lower()

        assert ContentTypes.COMMENTS_EXTENDED is not None
        assert "commentsextended+xml" in ContentTypes.COMMENTS_EXTENDED.lower()

        assert ContentTypes.FOOTNOTES is not None
        assert "footnotes+xml" in ContentTypes.FOOTNOTES.lower()

        assert ContentTypes.ENDNOTES is not None
        assert "endnotes+xml" in ContentTypes.ENDNOTES.lower()

        assert ContentTypes.DOCUMENT is not None
        assert "document.main+xml" in ContentTypes.DOCUMENT.lower()
