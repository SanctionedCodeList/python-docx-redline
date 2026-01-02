"""
Tests for the tracked changes listing functionality (Phase 9).

These tests verify:
- TrackedChange model class
- get_tracked_changes() method with filters
- accept_changes() and reject_changes() bulk operations
- has_tracked_changes and tracked_changes properties
"""

import tempfile
import zipfile
from pathlib import Path

import pytest

from python_docx_redline import ChangeType, Document, TrackedChange

# XML with various tracked changes for testing
DOCUMENT_WITH_TRACKED_CHANGES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is the original text.</w:t>
      </w:r>
      <w:ins w:id="1" w:author="Alice" w:date="2024-01-15T10:30:00Z">
        <w:r>
          <w:t> Added by Alice.</w:t>
        </w:r>
      </w:ins>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph </w:t>
      </w:r>
      <w:del w:id="2" w:author="Bob" w:date="2024-01-16T14:00:00Z">
        <w:r>
          <w:delText>removed text</w:delText>
        </w:r>
      </w:del>
      <w:r>
        <w:t> more text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph with </w:t>
      </w:r>
      <w:ins w:id="3" w:author="Alice" w:date="2024-01-17T09:00:00Z">
        <w:r>
          <w:t>more insertions</w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t> here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:pPrChange w:id="4" w:author="Carol" w:date="2024-01-18T11:00:00Z">
          <w:pPr/>
        </w:pPrChange>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:rPrChange w:id="5" w:author="Bob" w:date="2024-01-18T12:00:00Z">
            <w:rPr/>
          </w:rPrChange>
        </w:rPr>
        <w:t>Formatted paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_NO_CHANGES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a document with no tracked changes.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


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


class TestTrackedChangeModel:
    """Tests for the TrackedChange model class."""

    def test_change_type_enum_values(self):
        """Test that ChangeType enum has expected values."""
        assert ChangeType.INSERTION.value == "insertion"
        assert ChangeType.DELETION.value == "deletion"
        assert ChangeType.MOVE_FROM.value == "move_from"
        assert ChangeType.MOVE_TO.value == "move_to"
        assert ChangeType.FORMAT_RUN.value == "format_run"
        assert ChangeType.FORMAT_PARAGRAPH.value == "format_paragraph"

    def test_tracked_change_properties(self):
        """Test TrackedChange helper properties."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            changes = doc.get_tracked_changes()

            # Find an insertion
            insertion = next(c for c in changes if c.change_type == ChangeType.INSERTION)
            assert insertion.is_insertion
            assert not insertion.is_deletion
            assert not insertion.is_move
            assert not insertion.is_format_change

            # Find a deletion
            deletion = next(c for c in changes if c.change_type == ChangeType.DELETION)
            assert not deletion.is_insertion
            assert deletion.is_deletion
            assert not deletion.is_move
            assert not deletion.is_format_change

            # Find a format change
            format_change = next(c for c in changes if c.change_type == ChangeType.FORMAT_RUN)
            assert not format_change.is_insertion
            assert not format_change.is_deletion
            assert not format_change.is_move
            assert format_change.is_format_change

        finally:
            docx_path.unlink()

    def test_tracked_change_repr(self):
        """Test TrackedChange string representation."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            changes = doc.get_tracked_changes()
            insertion = next(c for c in changes if c.change_type == ChangeType.INSERTION)

            repr_str = repr(insertion)
            assert "TrackedChange" in repr_str
            assert "id=" in repr_str
            assert "type=insertion" in repr_str
            assert "author=" in repr_str

        finally:
            docx_path.unlink()

    def test_tracked_change_equality(self):
        """Test TrackedChange equality based on ID."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            changes1 = doc.get_tracked_changes()
            changes2 = doc.get_tracked_changes()

            # Same ID should be equal
            assert changes1[0] == changes2[0]

            # Different IDs should not be equal
            assert changes1[0] != changes1[1]

            # Can be used in sets
            change_set = set(changes1)
            assert len(change_set) == len(changes1)

        finally:
            docx_path.unlink()


class TestGetTrackedChanges:
    """Tests for get_tracked_changes() method."""

    def test_get_all_changes(self):
        """Test getting all tracked changes without filters."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            changes = doc.get_tracked_changes()

            # Should have 5 changes: 2 insertions, 1 deletion, 2 format changes
            assert len(changes) == 5

            # Verify we have all types
            types = {c.change_type for c in changes}
            assert ChangeType.INSERTION in types
            assert ChangeType.DELETION in types
            assert ChangeType.FORMAT_RUN in types
            assert ChangeType.FORMAT_PARAGRAPH in types

        finally:
            docx_path.unlink()

    def test_filter_by_type_insertion(self):
        """Test filtering changes by insertion type."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            insertions = doc.get_tracked_changes(change_type="insertion")

            assert len(insertions) == 2
            assert all(c.change_type == ChangeType.INSERTION for c in insertions)

        finally:
            docx_path.unlink()

    def test_filter_by_type_deletion(self):
        """Test filtering changes by deletion type."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            deletions = doc.get_tracked_changes(change_type="deletion")

            assert len(deletions) == 1
            assert deletions[0].change_type == ChangeType.DELETION
            assert deletions[0].text == "removed text"

        finally:
            docx_path.unlink()

    def test_filter_by_type_format_run(self):
        """Test filtering changes by run format type."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            format_changes = doc.get_tracked_changes(change_type="format_run")

            assert len(format_changes) == 1
            assert format_changes[0].change_type == ChangeType.FORMAT_RUN

        finally:
            docx_path.unlink()

    def test_filter_by_type_format_paragraph(self):
        """Test filtering changes by paragraph format type."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            format_changes = doc.get_tracked_changes(change_type="format_paragraph")

            assert len(format_changes) == 1
            assert format_changes[0].change_type == ChangeType.FORMAT_PARAGRAPH

        finally:
            docx_path.unlink()

    def test_filter_by_author(self):
        """Test filtering changes by author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            alice_changes = doc.get_tracked_changes(author="Alice")
            assert len(alice_changes) == 2
            assert all(c.author == "Alice" for c in alice_changes)

            bob_changes = doc.get_tracked_changes(author="Bob")
            assert len(bob_changes) == 2
            assert all(c.author == "Bob" for c in bob_changes)

            carol_changes = doc.get_tracked_changes(author="Carol")
            assert len(carol_changes) == 1
            assert carol_changes[0].author == "Carol"

        finally:
            docx_path.unlink()

    def test_filter_by_type_and_author(self):
        """Test filtering changes by both type and author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Alice's insertions only
            alice_insertions = doc.get_tracked_changes(change_type="insertion", author="Alice")
            assert len(alice_insertions) == 2
            assert all(c.change_type == ChangeType.INSERTION for c in alice_insertions)
            assert all(c.author == "Alice" for c in alice_insertions)

            # Bob's deletions only
            bob_deletions = doc.get_tracked_changes(change_type="deletion", author="Bob")
            assert len(bob_deletions) == 1
            assert bob_deletions[0].author == "Bob"
            assert bob_deletions[0].change_type == ChangeType.DELETION

        finally:
            docx_path.unlink()

    def test_filter_with_no_matches(self):
        """Test filtering that returns no matches."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Non-existent author
            no_changes = doc.get_tracked_changes(author="NonExistent")
            assert len(no_changes) == 0

            # Carol doesn't have insertions
            no_changes = doc.get_tracked_changes(change_type="insertion", author="Carol")
            assert len(no_changes) == 0

        finally:
            docx_path.unlink()

    def test_invalid_change_type_raises_error(self):
        """Test that invalid change_type raises ValueError."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            with pytest.raises(ValueError) as exc_info:
                doc.get_tracked_changes(change_type="invalid_type")

            assert "Invalid change_type" in str(exc_info.value)
            assert "invalid_type" in str(exc_info.value)

        finally:
            docx_path.unlink()

    def test_change_metadata(self):
        """Test that change metadata is correctly extracted."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            changes = doc.get_tracked_changes()

            # Find Alice's first insertion
            alice_insertion = next(
                c for c in changes if c.author == "Alice" and c.change_type == ChangeType.INSERTION
            )

            assert alice_insertion.id == "1"
            assert alice_insertion.author == "Alice"
            assert alice_insertion.date is not None
            assert alice_insertion.date.year == 2024
            assert alice_insertion.date.month == 1
            assert " Added by Alice." in alice_insertion.text

        finally:
            docx_path.unlink()

    def test_document_with_no_changes(self):
        """Test getting changes from document with no tracked changes."""
        docx_path = create_test_docx(DOCUMENT_NO_CHANGES)
        try:
            doc = Document(docx_path)
            changes = doc.get_tracked_changes()

            assert len(changes) == 0
            assert changes == []

        finally:
            docx_path.unlink()


class TestTrackedChangesProperty:
    """Tests for the tracked_changes property."""

    def test_tracked_changes_property(self):
        """Test tracked_changes property returns all changes."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Property should return same as get_tracked_changes()
            prop_changes = doc.tracked_changes
            method_changes = doc.get_tracked_changes()

            assert len(prop_changes) == len(method_changes)
            assert prop_changes == method_changes

        finally:
            docx_path.unlink()


class TestHasTrackedChanges:
    """Tests for the has_tracked_changes method."""

    def test_has_tracked_changes_true(self):
        """Test has_tracked_changes returns True when changes exist."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            assert doc.has_tracked_changes() is True

        finally:
            docx_path.unlink()

    def test_has_tracked_changes_false(self):
        """Test has_tracked_changes returns False when no changes."""
        docx_path = create_test_docx(DOCUMENT_NO_CHANGES)
        try:
            doc = Document(docx_path)
            assert doc.has_tracked_changes() is False

        finally:
            docx_path.unlink()

    def test_has_tracked_changes_after_accept_all(self):
        """Test has_tracked_changes after accepting all changes."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            assert doc.has_tracked_changes() is True
            doc.accept_all_changes()
            assert doc.has_tracked_changes() is False

        finally:
            docx_path.unlink()


class TestBulkAcceptChanges:
    """Tests for accept_changes() bulk operation."""

    def test_accept_changes_by_type(self):
        """Test accepting changes filtered by type."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Count insertions before
            insertions_before = len(doc.get_tracked_changes(change_type="insertion"))
            assert insertions_before == 2

            # Accept all insertions
            count = doc.accept_changes(change_type="insertion")
            assert count == 2

            # Verify insertions are gone
            insertions_after = len(doc.get_tracked_changes(change_type="insertion"))
            assert insertions_after == 0

            # Other changes should still exist
            deletions = doc.get_tracked_changes(change_type="deletion")
            assert len(deletions) == 1

        finally:
            docx_path.unlink()

    def test_accept_changes_by_author(self):
        """Test accepting changes filtered by author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Count Alice's changes before
            alice_before = len(doc.get_tracked_changes(author="Alice"))
            assert alice_before == 2

            # Accept Alice's changes
            count = doc.accept_changes(author="Alice")
            assert count == 2

            # Verify Alice's changes are gone
            alice_after = len(doc.get_tracked_changes(author="Alice"))
            assert alice_after == 0

            # Bob's changes should still exist
            bob_changes = doc.get_tracked_changes(author="Bob")
            assert len(bob_changes) == 2

        finally:
            docx_path.unlink()

    def test_accept_changes_by_type_and_author(self):
        """Test accepting changes filtered by both type and author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Accept Bob's deletions only (not his format change)
            count = doc.accept_changes(change_type="deletion", author="Bob")
            assert count == 1

            # Bob's deletion should be gone
            bob_deletions = doc.get_tracked_changes(change_type="deletion", author="Bob")
            assert len(bob_deletions) == 0

            # But Bob's format change should remain
            bob_formats = doc.get_tracked_changes(change_type="format_run", author="Bob")
            assert len(bob_formats) == 1

        finally:
            docx_path.unlink()

    def test_accept_changes_no_filter_calls_accept_all(self):
        """Test that accept_changes with no filter accepts all."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Initial state has changes
            assert doc.has_tracked_changes() is True

            # Accept all (no filters)
            doc.accept_changes()

            # All changes should be gone
            assert doc.has_tracked_changes() is False

        finally:
            docx_path.unlink()


class TestBulkRejectChanges:
    """Tests for reject_changes() bulk operation."""

    def test_reject_changes_by_type(self):
        """Test rejecting changes filtered by type."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Count insertions before
            insertions_before = len(doc.get_tracked_changes(change_type="insertion"))
            assert insertions_before == 2

            # Reject all insertions (removes inserted text)
            count = doc.reject_changes(change_type="insertion")
            assert count == 2

            # Verify insertions are gone
            insertions_after = len(doc.get_tracked_changes(change_type="insertion"))
            assert insertions_after == 0

            # Deletions should still exist
            deletions = doc.get_tracked_changes(change_type="deletion")
            assert len(deletions) == 1

        finally:
            docx_path.unlink()

    def test_reject_changes_by_author(self):
        """Test rejecting changes filtered by author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Count Bob's changes before
            bob_before = len(doc.get_tracked_changes(author="Bob"))
            assert bob_before == 2

            # Reject Bob's changes
            count = doc.reject_changes(author="Bob")
            assert count == 2

            # Verify Bob's changes are gone
            bob_after = len(doc.get_tracked_changes(author="Bob"))
            assert bob_after == 0

            # Alice's changes should still exist
            alice_changes = doc.get_tracked_changes(author="Alice")
            assert len(alice_changes) == 2

        finally:
            docx_path.unlink()

    def test_reject_changes_by_type_and_author(self):
        """Test rejecting changes filtered by both type and author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Reject Bob's format change only
            count = doc.reject_changes(change_type="format_run", author="Bob")
            assert count == 1

            # Bob's format change should be gone
            bob_formats = doc.get_tracked_changes(change_type="format_run", author="Bob")
            assert len(bob_formats) == 0

            # But Bob's deletion should remain
            bob_deletions = doc.get_tracked_changes(change_type="deletion", author="Bob")
            assert len(bob_deletions) == 1

        finally:
            docx_path.unlink()

    def test_reject_changes_no_filter_calls_reject_all(self):
        """Test that reject_changes with no filter rejects all."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Initial state has changes
            assert doc.has_tracked_changes() is True

            # Reject all (no filters)
            doc.reject_changes()

            # All changes should be gone
            assert doc.has_tracked_changes() is False

        finally:
            docx_path.unlink()


class TestTrackedChangeAcceptReject:
    """Tests for accept/reject on individual TrackedChange objects."""

    def test_accept_individual_change(self):
        """Test accepting a change via TrackedChange.accept()."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            changes = doc.get_tracked_changes()

            # Get an insertion and accept it
            insertion = next(c for c in changes if c.change_type == ChangeType.INSERTION)
            insertion_id = insertion.id

            insertion.accept()

            # Verify it's gone
            remaining = doc.get_tracked_changes()
            remaining_ids = {c.id for c in remaining}
            assert insertion_id not in remaining_ids

        finally:
            docx_path.unlink()

    def test_reject_individual_change(self):
        """Test rejecting a change via TrackedChange.reject()."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            changes = doc.get_tracked_changes()

            # Get a deletion and reject it (restores deleted text)
            deletion = next(c for c in changes if c.change_type == ChangeType.DELETION)
            deletion_id = deletion.id

            deletion.reject()

            # Verify it's gone
            remaining = doc.get_tracked_changes()
            remaining_ids = {c.id for c in remaining}
            assert deletion_id not in remaining_ids

        finally:
            docx_path.unlink()

    def test_accept_without_document_raises_error(self):
        """Test that accept() without document reference raises error."""
        # Create a TrackedChange without document reference
        from lxml import etree

        element = etree.fromstring(
            '<w:ins xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z"/>'
        )
        change = TrackedChange.from_element(element, ChangeType.INSERTION, document=None)

        with pytest.raises(ValueError) as exc_info:
            change.accept()

        assert "no document reference" in str(exc_info.value)

    def test_reject_without_document_raises_error(self):
        """Test that reject() without document reference raises error."""
        from lxml import etree

        element = etree.fromstring(
            '<w:ins xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z"/>'
        )
        change = TrackedChange.from_element(element, ChangeType.INSERTION, document=None)

        with pytest.raises(ValueError) as exc_info:
            change.reject()

        assert "no document reference" in str(exc_info.value)


class TestIntegrationWithExistingMethods:
    """Tests verifying integration with existing accept/reject methods."""

    def test_accept_by_author_via_accept_changes(self):
        """Test that accept_changes(author=...) uses existing accept_by_author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Both should produce the same result
            count = doc.accept_changes(author="Alice")
            assert count == 2

            # Alice's changes should be gone
            assert len(doc.get_tracked_changes(author="Alice")) == 0

        finally:
            docx_path.unlink()

    def test_reject_by_author_via_reject_changes(self):
        """Test that reject_changes(author=...) uses existing reject_by_author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Both should produce the same result
            count = doc.reject_changes(author="Bob")
            assert count == 2

            # Bob's changes should be gone
            assert len(doc.get_tracked_changes(author="Bob")) == 0

        finally:
            docx_path.unlink()

    def test_accept_change_by_id(self):
        """Test accept_change(id) works with IDs from get_tracked_changes."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            changes = doc.get_tracked_changes()

            # Get first change's ID
            first_id = changes[0].id

            # Accept by ID
            doc.accept_change(first_id)

            # Verify it's gone
            remaining = doc.get_tracked_changes()
            remaining_ids = {c.id for c in remaining}
            assert first_id not in remaining_ids

        finally:
            docx_path.unlink()

    def test_reject_change_by_id(self):
        """Test reject_change(id) works with IDs from get_tracked_changes."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            changes = doc.get_tracked_changes()

            # Get first change's ID
            first_id = changes[0].id

            # Reject by ID
            doc.reject_change(first_id)

            # Verify it's gone
            remaining = doc.get_tracked_changes()
            remaining_ids = {c.id for c in remaining}
            assert first_id not in remaining_ids

        finally:
            docx_path.unlink()


class TestTrackedChangeWithInsertions:
    """Tests for creating and listing insertions via API."""

    def test_insert_tracked_appears_in_get_tracked_changes(self):
        """Test that insert_tracked creates change visible in get_tracked_changes."""
        docx_path = create_test_docx(DOCUMENT_NO_CHANGES)
        try:
            doc = Document(docx_path, author="TestAuthor")

            # Initially no changes
            assert len(doc.get_tracked_changes()) == 0

            # Insert text
            doc.insert_tracked(" new text", after="document")

            # Now should have one insertion
            changes = doc.get_tracked_changes()
            assert len(changes) == 1
            assert changes[0].change_type == ChangeType.INSERTION
            assert changes[0].author == "TestAuthor"
            assert "new text" in changes[0].text

        finally:
            docx_path.unlink()

    def test_delete_tracked_appears_in_get_tracked_changes(self):
        """Test that delete_tracked creates change visible in get_tracked_changes."""
        docx_path = create_test_docx(DOCUMENT_NO_CHANGES)
        try:
            doc = Document(docx_path, author="TestAuthor")

            # Initially no changes
            assert len(doc.get_tracked_changes()) == 0

            # Delete text
            doc.delete_tracked("no tracked")

            # Now should have one deletion
            changes = doc.get_tracked_changes()
            assert len(changes) == 1
            assert changes[0].change_type == ChangeType.DELETION
            assert changes[0].author == "TestAuthor"
            assert "no tracked" in changes[0].text

        finally:
            docx_path.unlink()


class TestRejectChangesContaining:
    """Tests for reject_changes_containing() method."""

    def test_reject_changes_containing_basic(self):
        """Test basic text-based rejection of tracked changes."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # The first insertion contains " Added by Alice."
            # Reject changes containing "Added by Alice"
            count = doc.reject_changes_containing("Added by Alice")
            assert count == 1

            # Verify the change was rejected - should have 4 remaining
            remaining = doc.get_tracked_changes()
            assert len(remaining) == 4

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_case_insensitive(self):
        """Test case-insensitive text search for rejection."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # The deletion contains "removed text" - test case insensitive search
            count_lower = doc.reject_changes_containing("REMOVED")
            # With case-insensitive search (default), should find it
            assert count_lower == 1

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_case_sensitive(self):
        """Test case-sensitive text search for rejection."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # The deletion contains "removed text" - uppercase should NOT match
            count = doc.reject_changes_containing("REMOVED", match_case=True)
            assert count == 0

            # Correct case should match
            count = doc.reject_changes_containing("removed", match_case=True)
            assert count == 1

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_with_type_filter(self):
        """Test rejection with change_type filter."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Count deletions before
            deletions_before = len(doc.get_tracked_changes(change_type="deletion"))
            assert deletions_before == 1

            # Reject deletions containing "removed"
            count = doc.reject_changes_containing("removed", change_type="deletion")
            assert count == 1

            # Deletions should be gone
            deletions_after = len(doc.get_tracked_changes(change_type="deletion"))
            assert deletions_after == 0

            # Insertions should still exist
            insertions = doc.get_tracked_changes(change_type="insertion")
            assert len(insertions) == 2

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_with_author_filter(self):
        """Test rejection with author filter."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Bob has a deletion with "removed text"
            count = doc.reject_changes_containing("removed", author="Bob")
            assert count == 1

            # Alice doesn't have any changes with "removed"
            # Reload doc to test
            doc2 = Document(docx_path)
            count = doc2.reject_changes_containing("removed", author="Alice")
            assert count == 0

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_no_matches(self):
        """Test rejection when no changes match the text."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Search for text that doesn't exist in any change
            count = doc.reject_changes_containing("nonexistent text xyz")
            assert count == 0

            # All original changes should still exist
            assert len(doc.get_tracked_changes()) == 5

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_insertion_text(self):
        """Test rejection of insertions by text content."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Find insertions containing "Alice" in the text
            # The first insertion has " Added by Alice."
            count = doc.reject_changes_containing("Added by Alice", change_type="insertion")
            assert count == 1

            # Should have one fewer insertion
            insertions = doc.get_tracked_changes(change_type="insertion")
            assert len(insertions) == 1

        finally:
            docx_path.unlink()


class TestAcceptChangesContaining:
    """Tests for accept_changes_containing() method."""

    def test_accept_changes_containing_basic(self):
        """Test basic text-based acceptance of tracked changes."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Accept changes containing "insertions" (the third insertion)
            count = doc.accept_changes_containing("insertions")
            assert count == 1

            # The insertion should be accepted (unwrapped)
            remaining_insertions = doc.get_tracked_changes(change_type="insertion")
            assert len(remaining_insertions) == 1  # Was 2, now 1

        finally:
            docx_path.unlink()

    def test_accept_changes_containing_with_type_filter(self):
        """Test acceptance with change_type filter."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Accept insertions containing "Alice"
            count = doc.accept_changes_containing("Alice", change_type="insertion")
            assert count == 1

            # Should have one fewer insertion
            insertions = doc.get_tracked_changes(change_type="insertion")
            assert len(insertions) == 1

            # Other change types should be unaffected
            deletions = doc.get_tracked_changes(change_type="deletion")
            assert len(deletions) == 1

        finally:
            docx_path.unlink()

    def test_accept_changes_containing_case_insensitive(self):
        """Test case-insensitive acceptance."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Search for uppercase version - should still match
            count = doc.accept_changes_containing("MORE INSERTIONS")
            assert count == 1

        finally:
            docx_path.unlink()

    def test_accept_changes_containing_no_matches(self):
        """Test acceptance when no changes match."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Search for non-existent text
            count = doc.accept_changes_containing("xyz123nonexistent")
            assert count == 0

            # All changes should still exist
            assert len(doc.get_tracked_changes()) == 5

        finally:
            docx_path.unlink()

    def test_accept_changes_containing_with_author_and_type(self):
        """Test acceptance with both author and type filters."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Accept Bob's deletion containing "removed"
            count = doc.accept_changes_containing("removed", change_type="deletion", author="Bob")
            assert count == 1

            # Bob's deletion should be gone
            bob_deletions = doc.get_tracked_changes(change_type="deletion", author="Bob")
            assert len(bob_deletions) == 0

        finally:
            docx_path.unlink()


class TestChangesContainingRegex:
    """Tests for regex support in accept/reject_changes_containing() methods."""

    def test_reject_changes_containing_regex_basic(self):
        """Test basic regex pattern matching for rejection."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Match "Added by Alice" using regex pattern
            count = doc.reject_changes_containing(r"Added by \w+", regex=True)
            assert count == 1

            # Verify the change was rejected
            remaining = doc.get_tracked_changes()
            assert len(remaining) == 4

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_regex_case_insensitive(self):
        """Test regex with case-insensitive matching (default)."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Match "removed" with uppercase pattern - should work with case insensitive
            count = doc.reject_changes_containing(r"REMOVED\s+\w+", regex=True)
            assert count == 1

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_regex_case_sensitive(self):
        """Test regex with case-sensitive matching."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Uppercase pattern should NOT match when case-sensitive
            count = doc.reject_changes_containing(r"REMOVED\s+\w+", regex=True, match_case=True)
            assert count == 0

            # Correct case pattern should match
            count = doc.reject_changes_containing(r"removed\s+\w+", regex=True, match_case=True)
            assert count == 1

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_regex_with_type_filter(self):
        """Test regex matching with change_type filter."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Match insertions containing any word with "tion" suffix
            count = doc.reject_changes_containing(r"\w+tions?", regex=True, change_type="insertion")
            # "more insertions" should match
            assert count == 1

            # Verify only insertion was affected
            remaining_insertions = doc.get_tracked_changes(change_type="insertion")
            assert len(remaining_insertions) == 1

        finally:
            docx_path.unlink()

    def test_accept_changes_containing_regex_basic(self):
        """Test basic regex pattern matching for acceptance."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Match insertions with "more" followed by any word
            count = doc.accept_changes_containing(r"more\s+\w+", regex=True)
            assert count == 1

            # Verify change was accepted
            remaining = doc.get_tracked_changes()
            assert len(remaining) == 4

        finally:
            docx_path.unlink()

    def test_accept_changes_containing_regex_no_matches(self):
        """Test regex that matches nothing."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Pattern that won't match any change content
            count = doc.accept_changes_containing(r"\d{10}", regex=True)
            assert count == 0

            # All changes should still exist
            assert len(doc.get_tracked_changes()) == 5

        finally:
            docx_path.unlink()

    def test_reject_changes_containing_regex_with_author_filter(self):
        """Test regex with author filter combined."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            # Reject Bob's changes matching pattern
            count = doc.reject_changes_containing(r"removed.*", regex=True, author="Bob")
            assert count == 1

            # Alice should not be affected
            doc2 = Document(docx_path)
            count = doc2.reject_changes_containing(r"removed.*", regex=True, author="Alice")
            assert count == 0

        finally:
            docx_path.unlink()
