"""Tests for RelationshipManager class."""

from pathlib import Path

import pytest
from lxml import etree

from python_docx_redline.package import OOXMLPackage
from python_docx_redline.relationships import (
    RELS_NAMESPACE,
    RelationshipManager,
    RelationshipTypes,
)


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


class TestRelationshipManagerInit:
    """Test RelationshipManager initialization."""

    def test_init_existing_rels_file(self, package: OOXMLPackage) -> None:
        """Test initializing with an existing .rels file."""
        rel_mgr = RelationshipManager(package, "word/document.xml")
        assert rel_mgr._part_name == "word/document.xml"
        assert rel_mgr._rels_path.name == "document.xml.rels"
        assert "word/_rels" in str(rel_mgr._rels_path)

    def test_init_nonexistent_rels_file(self, package: OOXMLPackage) -> None:
        """Test initializing for a part without an existing .rels file."""
        rel_mgr = RelationshipManager(package, "word/comments.xml")
        assert rel_mgr._part_name == "word/comments.xml"
        assert rel_mgr._rels_path.name == "comments.xml.rels"

    def test_rels_path_computation(self, package: OOXMLPackage) -> None:
        """Test that .rels paths are computed correctly."""
        # Standard part
        rel_mgr = RelationshipManager(package, "word/document.xml")
        assert rel_mgr._rels_path.name == "document.xml.rels"
        assert "_rels" in str(rel_mgr._rels_path)

        # Nested part
        rel_mgr2 = RelationshipManager(package, "word/theme/theme1.xml")
        assert rel_mgr2._rels_path.name == "theme1.xml.rels"


class TestRelationshipManagerGet:
    """Test getting relationships."""

    def test_get_existing_relationship(self, package: OOXMLPackage) -> None:
        """Test getting an existing relationship."""
        rel_mgr = RelationshipManager(package, "word/document.xml")
        rel_id = rel_mgr.get_relationship(RelationshipTypes.STYLES)
        assert rel_id == "rId1"

    def test_get_nonexistent_relationship(self, package: OOXMLPackage) -> None:
        """Test getting a relationship that doesn't exist."""
        rel_mgr = RelationshipManager(package, "word/document.xml")
        rel_id = rel_mgr.get_relationship(RelationshipTypes.COMMENTS)
        assert rel_id is None

    def test_get_relationship_target(self, package: OOXMLPackage) -> None:
        """Test getting the target of a relationship."""
        rel_mgr = RelationshipManager(package, "word/document.xml")
        target = rel_mgr.get_relationship_target(RelationshipTypes.STYLES)
        assert target == "styles.xml"

    def test_has_relationship(self, package: OOXMLPackage) -> None:
        """Test checking if a relationship exists."""
        rel_mgr = RelationshipManager(package, "word/document.xml")
        assert rel_mgr.has_relationship(RelationshipTypes.STYLES) is True
        assert rel_mgr.has_relationship(RelationshipTypes.COMMENTS) is False


class TestRelationshipManagerAdd:
    """Test adding relationships."""

    def test_add_new_relationship(self, package: OOXMLPackage) -> None:
        """Test adding a new relationship."""
        rel_mgr = RelationshipManager(package, "word/document.xml")
        rel_id = rel_mgr.add_relationship(RelationshipTypes.COMMENTS, "comments.xml")

        # Should get the next available rId
        assert rel_id == "rId2"
        assert rel_mgr.is_modified is True

        # Should be able to retrieve it
        assert rel_mgr.get_relationship(RelationshipTypes.COMMENTS) == "rId2"
        assert rel_mgr.get_relationship_target(RelationshipTypes.COMMENTS) == "comments.xml"

    def test_add_existing_relationship_returns_same_id(self, package: OOXMLPackage) -> None:
        """Test that adding an existing relationship returns the same ID."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # Add a new relationship
        rel_id1 = rel_mgr.add_relationship(RelationshipTypes.COMMENTS, "comments.xml")

        # Try to add the same type again
        rel_id2 = rel_mgr.add_relationship(RelationshipTypes.COMMENTS, "comments.xml")

        assert rel_id1 == rel_id2

    def test_add_multiple_relationships(self, package: OOXMLPackage) -> None:
        """Test adding multiple different relationships."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        rel_id1 = rel_mgr.add_relationship(RelationshipTypes.COMMENTS, "comments.xml")
        rel_id2 = rel_mgr.add_relationship(RelationshipTypes.FOOTNOTES, "footnotes.xml")
        rel_id3 = rel_mgr.add_relationship(RelationshipTypes.ENDNOTES, "endnotes.xml")

        # Should have unique IDs
        assert rel_id1 == "rId2"
        assert rel_id2 == "rId3"
        assert rel_id3 == "rId4"

    def test_add_to_new_rels_file(self, package: OOXMLPackage) -> None:
        """Test adding a relationship when no .rels file exists."""
        rel_mgr = RelationshipManager(package, "word/comments.xml")

        # This part doesn't have a .rels file
        assert not rel_mgr._rels_path.exists()

        # Add a relationship
        rel_id = rel_mgr.add_relationship("http://example.com/relationship", "target.xml")

        assert rel_id == "rId1"  # First relationship
        assert rel_mgr.is_modified is True


class TestRelationshipManagerRemove:
    """Test removing relationships."""

    def test_remove_existing_relationship(self, package: OOXMLPackage) -> None:
        """Test removing an existing relationship."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # First verify it exists
        assert rel_mgr.has_relationship(RelationshipTypes.STYLES) is True

        # Remove it
        result = rel_mgr.remove_relationship(RelationshipTypes.STYLES)
        assert result is True
        assert rel_mgr.is_modified is True

        # Verify it's gone
        assert rel_mgr.has_relationship(RelationshipTypes.STYLES) is False

    def test_remove_nonexistent_relationship(self, package: OOXMLPackage) -> None:
        """Test removing a relationship that doesn't exist."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        result = rel_mgr.remove_relationship(RelationshipTypes.COMMENTS)
        assert result is False

    def test_remove_multiple_relationships(self, package: OOXMLPackage) -> None:
        """Test removing multiple relationships at once."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # Add some relationships first
        rel_mgr.add_relationship(RelationshipTypes.COMMENTS, "comments.xml")
        rel_mgr.add_relationship(RelationshipTypes.COMMENTS_EXTENDED, "commentsExtended.xml")
        rel_mgr.add_relationship(RelationshipTypes.FOOTNOTES, "footnotes.xml")

        # Remove comment-related relationships
        comment_types = [
            RelationshipTypes.COMMENTS,
            RelationshipTypes.COMMENTS_EXTENDED,
        ]
        removed = rel_mgr.remove_relationships(comment_types)

        assert removed == 2
        assert rel_mgr.has_relationship(RelationshipTypes.COMMENTS) is False
        assert rel_mgr.has_relationship(RelationshipTypes.COMMENTS_EXTENDED) is False
        # Footnotes should still exist
        assert rel_mgr.has_relationship(RelationshipTypes.FOOTNOTES) is True


class TestRelationshipManagerSave:
    """Test saving relationship changes."""

    def test_save_creates_rels_file(self, package: OOXMLPackage) -> None:
        """Test that save creates a new .rels file when needed."""
        rel_mgr = RelationshipManager(package, "word/comments.xml")

        # Add a relationship
        rel_mgr.add_relationship("http://example.com/rel", "target.xml")

        # Verify file doesn't exist yet
        assert not rel_mgr._rels_path.exists()

        # Save
        rel_mgr.save()

        # Verify file now exists
        assert rel_mgr._rels_path.exists()
        assert rel_mgr.is_modified is False

    def test_save_updates_existing_file(self, package: OOXMLPackage) -> None:
        """Test that save updates an existing .rels file."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # Add a relationship
        rel_mgr.add_relationship(RelationshipTypes.COMMENTS, "comments.xml")
        rel_mgr.save()

        # Read the file and verify
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(rel_mgr._rels_path), parser)
        root = tree.getroot()

        # Should have 2 relationships now (styles + comments)
        rels = root.findall(f"{{{RELS_NAMESPACE}}}Relationship")
        assert len(rels) == 2

        # Find the comments relationship
        comment_rel = None
        for rel in rels:
            if rel.get("Type") == RelationshipTypes.COMMENTS:
                comment_rel = rel
                break

        assert comment_rel is not None
        assert comment_rel.get("Target") == "comments.xml"

    def test_save_does_nothing_if_not_modified(self, package: OOXMLPackage) -> None:
        """Test that save doesn't write if nothing changed."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # Just read, don't modify
        _ = rel_mgr.get_relationship(RelationshipTypes.STYLES)

        # is_modified should be False
        assert rel_mgr.is_modified is False

        # Save should be a no-op
        rel_mgr.save()

    def test_save_persists_removals(self, package: OOXMLPackage) -> None:
        """Test that removed relationships are persisted."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # Remove the styles relationship
        rel_mgr.remove_relationship(RelationshipTypes.STYLES)
        rel_mgr.save()

        # Create a new manager and verify removal persisted
        rel_mgr2 = RelationshipManager(package, "word/document.xml")
        assert rel_mgr2.has_relationship(RelationshipTypes.STYLES) is False


class TestRelationshipTargetMode:
    """Test target_mode parameter for external relationships."""

    def test_add_relationship_with_target_mode(self, package: OOXMLPackage) -> None:
        """Test adding a relationship with TargetMode attribute."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # Add a hyperlink relationship with External target mode
        rel_id = rel_mgr.add_relationship(
            RelationshipTypes.HYPERLINK, "https://example.com", target_mode="External"
        )

        assert rel_id == "rId2"
        rel_mgr.save()

        # Read the file and verify TargetMode attribute
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(rel_mgr._rels_path), parser)
        root = tree.getroot()

        hyperlink_rel = None
        for rel in root.findall(f"{{{RELS_NAMESPACE}}}Relationship"):
            if rel.get("Type") == RelationshipTypes.HYPERLINK:
                hyperlink_rel = rel
                break

        assert hyperlink_rel is not None
        assert hyperlink_rel.get("Target") == "https://example.com"
        assert hyperlink_rel.get("TargetMode") == "External"

    def test_add_relationship_without_target_mode(self, package: OOXMLPackage) -> None:
        """Test adding a relationship without TargetMode (default behavior)."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # Add a relationship without target_mode
        rel_id = rel_mgr.add_relationship(RelationshipTypes.COMMENTS, "comments.xml")

        assert rel_id == "rId2"
        rel_mgr.save()

        # Read the file and verify no TargetMode attribute
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(rel_mgr._rels_path), parser)
        root = tree.getroot()

        comments_rel = None
        for rel in root.findall(f"{{{RELS_NAMESPACE}}}Relationship"):
            if rel.get("Type") == RelationshipTypes.COMMENTS:
                comments_rel = rel
                break

        assert comments_rel is not None
        assert comments_rel.get("TargetMode") is None

    def test_add_unique_relationship_with_target_mode(self, package: OOXMLPackage) -> None:
        """Test adding multiple unique relationships with TargetMode."""
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # Add multiple hyperlinks (same type, different targets)
        rel_id1 = rel_mgr.add_unique_relationship(
            RelationshipTypes.HYPERLINK, "https://example.com", target_mode="External"
        )
        rel_id2 = rel_mgr.add_unique_relationship(
            RelationshipTypes.HYPERLINK, "https://another.com", target_mode="External"
        )

        assert rel_id1 == "rId2"
        assert rel_id2 == "rId3"
        rel_mgr.save()

        # Read the file and verify both have TargetMode
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(rel_mgr._rels_path), parser)
        root = tree.getroot()

        hyperlink_rels = [
            rel
            for rel in root.findall(f"{{{RELS_NAMESPACE}}}Relationship")
            if rel.get("Type") == RelationshipTypes.HYPERLINK
        ]

        assert len(hyperlink_rels) == 2
        for rel in hyperlink_rels:
            assert rel.get("TargetMode") == "External"


class TestRelationshipTypes:
    """Test RelationshipTypes constants."""

    def test_common_relationship_types(self) -> None:
        """Test that common relationship type constants are defined."""
        assert RelationshipTypes.COMMENTS is not None
        assert "comments" in RelationshipTypes.COMMENTS.lower()

        assert RelationshipTypes.FOOTNOTES is not None
        assert "footnotes" in RelationshipTypes.FOOTNOTES.lower()

        assert RelationshipTypes.ENDNOTES is not None
        assert "endnotes" in RelationshipTypes.ENDNOTES.lower()

        assert RelationshipTypes.COMMENTS_EXTENDED is not None
        assert "commentsExtended" in RelationshipTypes.COMMENTS_EXTENDED

    def test_hyperlink_relationship_type(self) -> None:
        """Test that HYPERLINK relationship type is defined."""
        assert RelationshipTypes.HYPERLINK is not None
        assert "hyperlink" in RelationshipTypes.HYPERLINK.lower()
