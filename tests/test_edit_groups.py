"""
Tests for edit group functionality.

These tests verify the edit_group context manager and reject_edit_group method
for batch rejection of related tracked changes.
"""

import tempfile
from pathlib import Path

import pytest

from python_docx_redline import Document
from python_docx_redline.operations.edit_groups import EditGroup, EditGroupRegistry


def create_test_document() -> Path:
    """Create a test Word document for edit group testing."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph with some long text to replace.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph with another section to edit.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph for additional testing.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")
    return doc_path


# ============================================================================
# EditGroupRegistry Unit Tests
# ============================================================================


class TestEditGroupRegistry:
    """Tests for EditGroupRegistry class."""

    def test_registry_init(self):
        """Test registry initializes with empty state."""
        registry = EditGroupRegistry()
        assert registry.active_group is None
        assert registry.list_groups() == []

    def test_start_group(self):
        """Test starting a new group."""
        registry = EditGroupRegistry()
        registry.start_group("round1")

        assert registry.active_group == "round1"
        assert registry.group_exists("round1")
        assert registry.get_group_status("round1") == "active"

    def test_end_group(self):
        """Test ending an active group."""
        registry = EditGroupRegistry()
        registry.start_group("round1")
        registry.end_group()

        assert registry.active_group is None
        assert registry.get_group_status("round1") == "completed"

    def test_add_change_id(self):
        """Test adding change IDs to active group."""
        registry = EditGroupRegistry()
        registry.start_group("round1")
        registry.add_change_id(1)
        registry.add_change_id(2)
        registry.add_change_id(3)
        registry.end_group()

        ids = registry.get_group_ids("round1")
        assert ids == [1, 2, 3]

    def test_add_change_id_no_active_group(self):
        """Test that add_change_id does nothing without active group."""
        registry = EditGroupRegistry()
        registry.add_change_id(1)  # Should not raise

        # Verify no group was created
        assert registry.list_groups() == []

    def test_get_group_ids_returns_copy(self):
        """Test that get_group_ids returns a copy."""
        registry = EditGroupRegistry()
        registry.start_group("round1")
        registry.add_change_id(1)
        registry.end_group()

        ids1 = registry.get_group_ids("round1")
        ids1.append(999)  # Modify the copy

        ids2 = registry.get_group_ids("round1")
        assert ids2 == [1]  # Original unchanged

    def test_start_group_already_active_error(self):
        """Test error when starting group while another is active."""
        registry = EditGroupRegistry()
        registry.start_group("round1")

        with pytest.raises(ValueError, match="round1.*already active"):
            registry.start_group("round2")

    def test_start_group_duplicate_name_error(self):
        """Test error when starting group with existing name."""
        registry = EditGroupRegistry()
        registry.start_group("round1")
        registry.end_group()

        with pytest.raises(ValueError, match="round1.*already exists"):
            registry.start_group("round1")

    def test_get_group_ids_nonexistent_error(self):
        """Test error when getting IDs for nonexistent group."""
        registry = EditGroupRegistry()

        with pytest.raises(ValueError, match="No group 'nonexistent' found"):
            registry.get_group_ids("nonexistent")

    def test_mark_rejected(self):
        """Test marking a group as rejected."""
        registry = EditGroupRegistry()
        registry.start_group("round1")
        registry.end_group()
        registry.mark_rejected("round1")

        assert registry.get_group_status("round1") == "rejected"

    def test_mark_rejected_nonexistent(self):
        """Test that mark_rejected doesn't raise for nonexistent group."""
        registry = EditGroupRegistry()
        registry.mark_rejected("nonexistent")  # Should not raise

    def test_list_groups(self):
        """Test listing all groups."""
        registry = EditGroupRegistry()
        registry.start_group("round1")
        registry.end_group()
        registry.start_group("round2")
        registry.end_group()

        groups = registry.list_groups()
        assert "round1" in groups
        assert "round2" in groups

    def test_get_group_status_nonexistent(self):
        """Test get_group_status returns None for nonexistent group."""
        registry = EditGroupRegistry()
        assert registry.get_group_status("nonexistent") is None


# ============================================================================
# EditGroup Dataclass Tests
# ============================================================================


class TestEditGroup:
    """Tests for EditGroup dataclass."""

    def test_edit_group_creation(self):
        """Test creating an EditGroup."""
        group = EditGroup(name="test", status="active")
        assert group.name == "test"
        assert group.status == "active"
        assert group.change_ids == []
        assert group.created_at is not None

    def test_edit_group_with_change_ids(self):
        """Test creating EditGroup with change IDs."""
        group = EditGroup(name="test", status="completed", change_ids=[1, 2, 3])
        assert group.change_ids == [1, 2, 3]


# ============================================================================
# Document Integration Tests
# ============================================================================


class TestEditGroupContextManager:
    """Tests for the edit_group context manager on Document."""

    def test_edit_group_basic(self):
        """Test basic edit group context manager usage."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                doc.insert(text=" inserted", after="First", track=True)

            # Verify the group was created and completed
            assert doc._edit_groups.group_exists("round1")
            assert doc._edit_groups.get_group_status("round1") == "completed"

            # Verify change IDs were captured
            ids = doc._edit_groups.get_group_ids("round1")
            assert len(ids) > 0

        finally:
            doc_path.unlink()

    def test_edit_group_multiple_changes(self):
        """Test edit group captures multiple change IDs."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path, minimal_edits=False)

            with doc.edit_group("condensing"):
                doc.replace_tracked("long text", "short")
                # Insert after a different anchor to avoid the tracked change wrapper
                doc.insert_tracked(" more", after="Second")

            ids = doc._edit_groups.get_group_ids("condensing")
            # Should have at least 3 changes (replace creates del+ins, insert creates ins)
            assert len(ids) >= 3

        finally:
            doc_path.unlink()

    def test_edit_group_exception_handling(self):
        """Test that edit group closes even if exception occurs."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            try:
                with doc.edit_group("round1"):
                    doc.insert(text=" ok", after="First", track=True)
                    raise RuntimeError("Simulated error")
            except RuntimeError:
                pass

            # Group should still be completed (not left active)
            assert doc._edit_groups.active_group is None
            assert doc._edit_groups.get_group_status("round1") == "completed"

        finally:
            doc_path.unlink()

    def test_edit_group_already_active_error(self):
        """Test error when nesting edit groups."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError, match="already active"):
                with doc.edit_group("round1"):
                    with doc.edit_group("round2"):
                        pass

        finally:
            doc_path.unlink()

    def test_edit_group_duplicate_name_error(self):
        """Test error when reusing group name."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                doc.insert(text=" text", after="First", track=True)

            with pytest.raises(ValueError, match="already exists"):
                with doc.edit_group("round1"):
                    pass

        finally:
            doc_path.unlink()


class TestRejectEditGroup:
    """Tests for reject_edit_group method."""

    def test_reject_edit_group_basic(self):
        """Test basic rejection of an edit group."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            # Make some tracked changes
            with doc.edit_group("round1"):
                doc.insert(text=" INSERTED", after="First", track=True)

            # Verify the insertion exists
            changes_before = doc.get_tracked_changes()
            assert len(changes_before) > 0

            # Reject the edit group
            count = doc.reject_edit_group("round1")
            assert count > 0

            # Verify the changes are rejected
            changes_after = doc.get_tracked_changes()
            assert len(changes_after) == 0

        finally:
            doc_path.unlink()

    def test_reject_edit_group_multiple_changes(self):
        """Test rejecting a group with multiple changes."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path, minimal_edits=False)

            with doc.edit_group("condensing"):
                doc.replace_tracked("long text", "short")
                doc.delete_tracked("another")

            changes_before = doc.get_tracked_changes()
            # Should have changes for both operations
            assert len(changes_before) >= 2

            count = doc.reject_edit_group("condensing")
            assert count >= 2

            # All changes should be rejected
            changes_after = doc.get_tracked_changes()
            assert len(changes_after) == 0

        finally:
            doc_path.unlink()

    def test_reject_edit_group_marks_rejected(self):
        """Test that reject_edit_group marks the group as rejected."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                doc.insert(text=" text", after="First", track=True)

            doc.reject_edit_group("round1")

            assert doc._edit_groups.get_group_status("round1") == "rejected"

        finally:
            doc_path.unlink()

    def test_reject_edit_group_nonexistent_error(self):
        """Test error when rejecting nonexistent group."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError, match="No group 'nonexistent' found"):
                doc.reject_edit_group("nonexistent")

        finally:
            doc_path.unlink()

    def test_reject_edit_group_already_rejected(self):
        """Test rejecting an already-rejected group."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                doc.insert(text=" text", after="First", track=True)

            # First rejection
            count1 = doc.reject_edit_group("round1")
            assert count1 > 0

            # Second rejection (changes already gone)
            count2 = doc.reject_edit_group("round1")
            assert count2 == 0

        finally:
            doc_path.unlink()


class TestAcceptEditGroup:
    """Tests for accept_edit_group method."""

    def test_accept_edit_group_basic(self):
        """Test basic acceptance of an edit group."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                doc.insert_tracked(" INSERTED", after="First")

            # Verify changes exist
            changes_before = doc.get_tracked_changes()
            assert len(changes_before) > 0

            # Accept the edit group
            count = doc.accept_edit_group("round1")
            assert count > 0

            # Verify changes are accepted (removed as tracked)
            changes_after = doc.get_tracked_changes()
            assert len(changes_after) == 0

            # But content should still be there
            text = doc.get_text()
            assert "INSERTED" in text

        finally:
            doc_path.unlink()

    def test_accept_edit_group_nonexistent_error(self):
        """Test error when accepting nonexistent group."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError, match="No group 'nonexistent' found"):
                doc.accept_edit_group("nonexistent")

        finally:
            doc_path.unlink()


class TestEditGroupTracking:
    """Tests for edit group change ID tracking."""

    def test_change_ids_captured_for_insert(self):
        """Test that insert operations capture change IDs."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                doc.insert(text=" new", after="First", track=True)

            ids = doc._edit_groups.get_group_ids("round1")
            assert len(ids) == 1  # One insertion

        finally:
            doc_path.unlink()

    def test_change_ids_captured_for_delete(self):
        """Test that delete operations capture change IDs."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                doc.delete(text="some", track=True)

            ids = doc._edit_groups.get_group_ids("round1")
            assert len(ids) == 1  # One deletion

        finally:
            doc_path.unlink()

    def test_change_ids_captured_for_replace(self):
        """Test that replace operations capture change IDs."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path, minimal_edits=False)

            with doc.edit_group("round1"):
                doc.replace_tracked("long text", "short")

            ids = doc._edit_groups.get_group_ids("round1")
            # Replace creates del + ins = 2 change IDs
            assert len(ids) == 2

        finally:
            doc_path.unlink()

    def test_no_ids_captured_for_untracked(self):
        """Test that untracked operations don't capture change IDs."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                # Untracked edit
                doc.insert(text=" new", after="First", track=False)

            ids = doc._edit_groups.get_group_ids("round1")
            assert len(ids) == 0

        finally:
            doc_path.unlink()

    def test_mixed_tracked_untracked(self):
        """Test group with mix of tracked and untracked edits."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("mixed"):
                doc.insert(text=" tracked1", after="First", track=True)
                doc.insert(text=" untracked", after="Second", track=False)
                doc.insert(text=" tracked2", after="Third", track=True)

            ids = doc._edit_groups.get_group_ids("mixed")
            assert len(ids) == 2  # Only tracked changes

        finally:
            doc_path.unlink()


class TestMultipleGroups:
    """Tests for multiple edit groups in the same document."""

    def test_multiple_sequential_groups(self):
        """Test creating multiple groups sequentially."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                doc.insert(text=" r1", after="First", track=True)

            with doc.edit_group("round2"):
                doc.insert(text=" r2", after="Second", track=True)

            # Both groups should exist
            assert doc._edit_groups.group_exists("round1")
            assert doc._edit_groups.group_exists("round2")

            # Each should have its own IDs
            ids1 = doc._edit_groups.get_group_ids("round1")
            ids2 = doc._edit_groups.get_group_ids("round2")
            assert len(ids1) == 1
            assert len(ids2) == 1
            assert ids1[0] != ids2[0]

        finally:
            doc_path.unlink()

    def test_reject_one_group_not_other(self):
        """Test rejecting one group doesn't affect another."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("round1"):
                doc.insert(text=" R1", after="First", track=True)

            with doc.edit_group("round2"):
                doc.insert(text=" R2", after="Second", track=True)

            # Reject only round1
            doc.reject_edit_group("round1")

            # round2 changes should still exist
            changes = doc.get_tracked_changes()
            assert len(changes) == 1
            assert "R2" in changes[0].text

        finally:
            doc_path.unlink()


class TestEdgeCases:
    """Tests for edge cases and error conditions."""

    def test_empty_group(self):
        """Test group with no edits."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with doc.edit_group("empty"):
                pass  # No edits

            ids = doc._edit_groups.get_group_ids("empty")
            assert len(ids) == 0

            # Rejecting empty group should work
            count = doc.reject_edit_group("empty")
            assert count == 0

        finally:
            doc_path.unlink()

    def test_group_before_any_edits(self):
        """Test accessing edit groups before any edits."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            # Accessing _edit_groups before any edits should work
            with doc.edit_group("first"):
                pass

            assert doc._edit_groups is not None

        finally:
            doc_path.unlink()


# Run tests with: pytest tests/test_edit_groups.py -v
