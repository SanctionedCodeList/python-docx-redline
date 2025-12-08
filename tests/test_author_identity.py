"""
Tests for MS365 author identity integration (Phase 4).

Tests the AuthorIdentity dataclass and its integration with Document
and TrackedXMLGenerator to embed MS365 identity information in tracked changes.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_redline import AuthorIdentity, Document


def create_test_document() -> Path:
    """Create a simple test document."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>This is a test document.</w:t></w:r></w:p>
<w:p><w:r><w:t>It has multiple paragraphs.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


# AuthorIdentity dataclass tests


def test_author_identity_creation() -> None:
    """Test creating an AuthorIdentity with all fields."""
    identity = AuthorIdentity(
        author="Hancock, Parker",
        email="parker.hancock@company.com",
        provider_id="AD",
        guid="c5c513d2-1f51-4d69-ae91-17e5787f9bfc",
    )

    assert identity.author == "Hancock, Parker"
    assert identity.email == "parker.hancock@company.com"
    assert identity.provider_id == "AD"
    assert identity.guid == "c5c513d2-1f51-4d69-ae91-17e5787f9bfc"


def test_author_identity_defaults() -> None:
    """Test AuthorIdentity with default values."""
    identity = AuthorIdentity(author="John Doe", email="john@example.com")

    assert identity.provider_id == "AD"  # Default
    assert identity.guid == ""  # Default


def test_author_identity_display_name() -> None:
    """Test display_name property."""
    identity = AuthorIdentity(author="Hancock, Parker", email="parker@company.com")

    assert identity.display_name == "Hancock, Parker"


def test_author_identity_str() -> None:
    """Test string representation."""
    identity = AuthorIdentity(author="John Doe", email="john@example.com")

    assert str(identity) == "John Doe <john@example.com>"


def test_author_identity_repr() -> None:
    """Test repr representation."""
    identity = AuthorIdentity(
        author="John Doe", email="john@example.com", provider_id="AD", guid="abc123"
    )

    repr_str = repr(identity)
    assert "AuthorIdentity" in repr_str
    assert "John Doe" in repr_str
    assert "john@example.com" in repr_str


def test_author_identity_validation_empty_author() -> None:
    """Test that empty author name raises ValueError."""
    with pytest.raises(ValueError, match="Author name cannot be empty"):
        AuthorIdentity(author="", email="test@example.com")


def test_author_identity_validation_empty_email() -> None:
    """Test that empty email raises ValueError."""
    with pytest.raises(ValueError, match="Email cannot be empty"):
        AuthorIdentity(author="John Doe", email="")


def test_author_identity_validation_invalid_email() -> None:
    """Test that invalid email format raises ValueError."""
    with pytest.raises(ValueError, match="Invalid email format"):
        AuthorIdentity(author="John Doe", email="notanemail")


# Document integration tests


def test_document_with_string_author() -> None:
    """Test Document with simple string author (backward compatibility)."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path, author="Claude")

        assert doc.author == "Claude"
        assert doc._author_identity is None

    finally:
        doc_path.unlink()


def test_document_with_author_identity() -> None:
    """Test Document with AuthorIdentity."""
    doc_path = create_test_document()
    try:
        identity = AuthorIdentity(
            author="Hancock, Parker",
            email="parker@company.com",
            provider_id="AD",
            guid="test-guid-123",
        )
        doc = Document(doc_path, author=identity)

        assert doc.author == "Hancock, Parker"
        assert doc._author_identity == identity
        assert doc._author_identity.email == "parker@company.com"

    finally:
        doc_path.unlink()


def test_document_default_author() -> None:
    """Test Document with default author."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        assert doc.author == "Claude"
        assert doc._author_identity is None

    finally:
        doc_path.unlink()


# Tracked changes with identity


def test_insert_tracked_with_identity() -> None:
    """Test that insertions include MS365 identity attributes."""
    doc_path = create_test_document()
    try:
        identity = AuthorIdentity(
            author="Hancock, Parker",
            email="parker@company.com",
            provider_id="AD",
            guid="c5c513d2-1f51-4d69-ae91-17e5787f9bfc",
        )
        doc = Document(doc_path, author=identity)

        doc.insert_tracked(" with identity", after="test document")

        # Read the XML to verify identity attributes are present
        xml_content = etree.tostring(doc.xml_root, encoding="unicode")

        # Check for w:ins element with identity attributes
        assert (
            'w15:userId="c5c513d2-1f51-4d69-ae91-17e5787f9bfc"' in xml_content
            or 'userId="c5c513d2-1f51-4d69-ae91-17e5787f9bfc"' in xml_content
        )
        assert 'w15:providerId="AD"' in xml_content or 'providerId="AD"' in xml_content
        assert 'w:author="Hancock, Parker"' in xml_content

    finally:
        doc_path.unlink()


def test_delete_tracked_with_identity() -> None:
    """Test that deletions include MS365 identity attributes."""
    doc_path = create_test_document()
    try:
        identity = AuthorIdentity(
            author="Smith, John",
            email="john.smith@company.com",
            provider_id="AD",
            guid="test-guid-456",
        )
        doc = Document(doc_path, author=identity)

        doc.delete_tracked("multiple")

        # Read the XML to verify identity attributes are present
        xml_content = etree.tostring(doc.xml_root, encoding="unicode")

        # Check for w:del element with identity attributes
        assert 'userId="test-guid-456"' in xml_content
        assert 'providerId="AD"' in xml_content
        assert 'w:author="Smith, John"' in xml_content

    finally:
        doc_path.unlink()


def test_replace_tracked_with_identity() -> None:
    """Test that replacements (delete + insert) include identity."""
    doc_path = create_test_document()
    try:
        identity = AuthorIdentity(
            author="Doe, Jane", email="jane.doe@company.com", provider_id="AD", guid="jane-guid-789"
        )
        doc = Document(doc_path, author=identity)

        doc.replace_tracked("test", "sample")

        xml_content = etree.tostring(doc.xml_root, encoding="unicode")

        # Should have both insertion and deletion with identity
        assert 'userId="jane-guid-789"' in xml_content
        assert 'providerId="AD"' in xml_content
        assert 'w:author="Doe, Jane"' in xml_content

    finally:
        doc_path.unlink()


def test_identity_without_guid() -> None:
    """Test that identity works even without GUID."""
    doc_path = create_test_document()
    try:
        identity = AuthorIdentity(
            author="No GUID User",
            email="noguid@company.com",
            provider_id="AD",
            # guid defaults to empty string
        )
        doc = Document(doc_path, author=identity)

        doc.insert_tracked(" test", after="document")

        xml_content = etree.tostring(doc.xml_root, encoding="unicode")

        # Should have providerId but not userId
        assert 'providerId="AD"' in xml_content
        assert 'w:author="No GUID User"' in xml_content
        # Should not have empty userId attribute
        assert 'userId=""' not in xml_content

    finally:
        doc_path.unlink()


def test_multiple_edits_same_identity() -> None:
    """Test multiple edits with same identity."""
    doc_path = create_test_document()
    try:
        identity = AuthorIdentity(
            author="Consistent Author",
            email="consistent@company.com",
            provider_id="AD",
            guid="consistent-guid",
        )
        doc = Document(doc_path, author=identity)

        doc.insert_tracked(" first", after="document")
        doc.insert_tracked(" second", after="paragraphs")
        doc.delete_tracked("is a")

        xml_content = etree.tostring(doc.xml_root, encoding="unicode")

        # All changes should have same identity
        guid_count = xml_content.count('userId="consistent-guid"')
        assert guid_count >= 3  # At least 3 changes with this GUID

    finally:
        doc_path.unlink()


def test_batch_operations_with_identity() -> None:
    """Test batch operations respect identity."""
    doc_path = create_test_document()
    try:
        identity = AuthorIdentity(
            author="Batch Author",
            email="batch@company.com",
            provider_id="AD",
            guid="batch-guid-123",
        )
        doc = Document(doc_path, author=identity)

        edits = [
            {"type": "insert_tracked", "text": " A", "after": "document"},
            {"type": "delete_tracked", "text": "test"},
            {"type": "replace_tracked", "find": "multiple", "replace": "many"},
        ]

        results = doc.apply_edits(edits)

        assert all(r.success for r in results)

        xml_content = etree.tostring(doc.xml_root, encoding="unicode")
        assert 'userId="batch-guid-123"' in xml_content
        assert 'w:author="Batch Author"' in xml_content

    finally:
        doc_path.unlink()


def test_structural_operations_with_identity() -> None:
    """Test that structural operations work with identity."""
    doc_path = create_test_document()
    try:
        identity = AuthorIdentity(
            author="Structure Author",
            email="structure@company.com",
            provider_id="AD",
            guid="structure-guid",
        )
        doc = Document(doc_path, author=identity)

        # Insert paragraph with tracked changes
        doc.insert_paragraph("New paragraph with identity", after="test document", track=True)

        xml_content = etree.tostring(doc.xml_root, encoding="unicode")

        # Tracked paragraph insertion should have identity
        assert (
            'userId="structure-guid"' in xml_content or 'w:author="Structure Author"' in xml_content
        )

    finally:
        doc_path.unlink()
