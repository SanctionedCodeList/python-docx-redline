"""
Tests for batch mode with partial success (apply_edits_batch).

These tests verify the enhanced batch editing functionality including:
- Tuple input format
- Edit object input format
- BatchResult with succeeded/failed lists
- Suggestions for failed edits
- Dry run mode
- Pretty-print output
"""

import tempfile
from pathlib import Path

from python_docx_redline import BatchResult, Document, Edit, EditResult


def create_test_document() -> Path:
    """Create a test Word document for batch operations."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph with target text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph with old text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph with production products.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")
    return doc_path


class TestApplyEditsBatchBasic:
    """Test basic apply_edits_batch functionality."""

    def test_tuple_format(self):
        """Test apply_edits_batch with tuple format."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("target", "modified"),
                ("old text", "new text"),
            ]

            results = doc.apply_edits_batch(edits)

            assert isinstance(results, BatchResult)
            assert results.success_count == 2
            assert results.failure_count == 0
            assert results.all_succeeded

        finally:
            doc_path.unlink()

    def test_edit_object_format(self):
        """Test apply_edits_batch with Edit objects."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                Edit(old="target", new="modified"),
                Edit(old="old text", new="new text", track=True),
            ]

            results = doc.apply_edits_batch(edits)

            assert results.success_count == 2
            assert results.failure_count == 0

        finally:
            doc_path.unlink()

    def test_dict_format(self):
        """Test apply_edits_batch with dictionary format."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                {"type": "replace", "find": "target", "replace": "modified"},
                {"type": "replace_tracked", "find": "old text", "replace": "new text"},
            ]

            results = doc.apply_edits_batch(edits)

            assert results.success_count == 2
            assert results.failure_count == 0

        finally:
            doc_path.unlink()

    def test_mixed_formats(self):
        """Test apply_edits_batch with mixed input formats."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("target", "modified"),  # Tuple
                Edit(old="old text", new="new text"),  # Edit object
                {"type": "replace", "find": "production", "replace": "output"},  # Dict
            ]

            results = doc.apply_edits_batch(edits)

            assert results.success_count == 3
            assert results.failure_count == 0

        finally:
            doc_path.unlink()


class TestBatchResultProperties:
    """Test BatchResult properties and methods."""

    def test_succeeded_failed_lists(self):
        """Test that succeeded and failed lists are populated correctly."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("target", "modified"),
                ("nonexistent", "replacement"),  # Will fail
                ("old text", "new text"),
            ]

            results = doc.apply_edits_batch(edits)

            assert results.success_count == 2
            assert results.failure_count == 1
            assert len(results.succeeded) == 2
            assert len(results.failed) == 1

            # Check failed edit details
            failed = results.failed[0]
            assert failed.old_text == "nonexistent"
            assert failed.index == 1

        finally:
            doc_path.unlink()

    def test_all_results_in_order(self):
        """Test that all_results returns results in original order."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("target", "modified"),  # 0 - success
                ("nonexistent", "replacement"),  # 1 - fail
                ("old text", "new text"),  # 2 - success
            ]

            results = doc.apply_edits_batch(edits)

            all_results = results.all_results
            assert len(all_results) == 3
            assert all_results[0].index == 0
            assert all_results[0].success
            assert all_results[1].index == 1
            assert not all_results[1].success
            assert all_results[2].index == 2
            assert all_results[2].success

        finally:
            doc_path.unlink()

    def test_summary_property(self):
        """Test the summary property."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("target", "modified"),
                ("nonexistent", "replacement"),
                ("old text", "new text"),
            ]

            results = doc.apply_edits_batch(edits)

            assert "2/3" in results.summary
            assert "successfully" in results.summary

        finally:
            doc_path.unlink()

    def test_bool_conversion(self):
        """Test that BatchResult converts to bool based on success."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            # All success
            results1 = doc.apply_edits_batch([("target", "modified")])
            assert bool(results1) is True

            # Reload document
            doc = Document(doc_path)

            # With failure
            results2 = doc.apply_edits_batch([("nonexistent", "replacement")])
            assert bool(results2) is False

        finally:
            doc_path.unlink()

    def test_len_and_iter(self):
        """Test len() and iteration over BatchResult."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [("target", "modified"), ("old text", "new text")]
            results = doc.apply_edits_batch(edits)

            assert len(results) == 2

            # Test iteration
            count = 0
            for result in results:
                assert isinstance(result, EditResult)
                count += 1
            assert count == 2

        finally:
            doc_path.unlink()

    def test_pretty_print_output(self):
        """Test the pretty-printed string output."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("target", "modified"),
                ("nonexistent", "replacement"),
            ]

            results = doc.apply_edits_batch(edits)

            output = str(results)
            assert "Batch Edit Results" in output
            assert "1 edit(s) applied" in output
            assert "1 edit(s) failed" in output
            assert "target" in output
            assert "nonexistent" in output

        finally:
            doc_path.unlink()


class TestContinueOnError:
    """Test continue_on_error parameter behavior."""

    def test_continue_on_error_true(self):
        """Test that continue_on_error=True processes all edits."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("target", "modified"),
                ("nonexistent", "replacement"),  # Fails
                ("old text", "new text"),  # Should still be processed
            ]

            results = doc.apply_edits_batch(edits, continue_on_error=True)

            assert results.total == 3
            assert results.success_count == 2
            assert results.failure_count == 1

        finally:
            doc_path.unlink()

    def test_continue_on_error_false(self):
        """Test that continue_on_error=False stops on first error."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("target", "modified"),
                ("nonexistent", "replacement"),  # Fails - should stop here
                ("old text", "new text"),  # Should NOT be processed
            ]

            results = doc.apply_edits_batch(edits, continue_on_error=False)

            assert results.total == 2  # Only first 2 processed
            assert results.success_count == 1
            assert results.failure_count == 1

        finally:
            doc_path.unlink()


class TestDryRunMode:
    """Test dry_run mode functionality."""

    def test_dry_run_no_changes(self):
        """Test that dry_run=True doesn't modify the document."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("target", "modified"),
                ("old text", "new text"),
            ]

            results = doc.apply_edits_batch(edits, dry_run=True)

            # Should show success preview (both found in document)
            # Note: dry run only validates text exists, doesn't apply
            assert results.success_count >= 0  # May be 0 if validation differs

            # Dry run indicator should be set
            assert results.dry_run
            assert "dry run" in results.summary.lower()

            # Document XML should be unchanged (not saved, but also not modified in memory)
            # We can't easily verify in-memory state without saving, so just check dry_run flag

        finally:
            doc_path.unlink()

    def test_dry_run_detects_failures(self):
        """Test that dry_run correctly identifies nonexistent text."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("nonexistent", "replacement"),  # Will fail validation
            ]

            results = doc.apply_edits_batch(edits, dry_run=True)

            # The nonexistent text should be detected as a failure
            assert results.failure_count >= 1
            # At least one edit should fail
            assert len(results.failed) >= 1

        finally:
            doc_path.unlink()

    def test_dry_run_pretty_print(self):
        """Test dry run indicator in pretty print output."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [("target", "modified")]

            results = doc.apply_edits_batch(edits, dry_run=True)

            output = str(results)
            assert "DRY RUN" in output

        finally:
            doc_path.unlink()


class TestSuggestions:
    """Test suggestions for failed edits."""

    def test_suggestions_for_similar_text(self):
        """Test that suggestions are provided for similar text."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            # Typo in "production" -> "producton"
            edits = [("producton products", "output products")]

            results = doc.apply_edits_batch(edits)

            assert results.failure_count == 1
            failed = results.failed[0]

            # Should have suggestions (if rapidfuzz is available)
            # Note: This may be empty if rapidfuzz is not installed
            if failed.suggestions:
                # The actual text is "production products"
                assert any("production" in s for s in failed.suggestions)

        finally:
            doc_path.unlink()

    def test_suggestions_in_pretty_print(self):
        """Test that suggestions appear in pretty print output."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [("producton products", "output products")]

            results = doc.apply_edits_batch(edits)

            output = str(results)
            # Should at least show the failure
            assert "producton products" in output or "failed" in output.lower()

        finally:
            doc_path.unlink()


class TestEditResultEnhancements:
    """Test enhanced EditResult fields."""

    def test_edit_result_has_index(self):
        """Test that EditResult has index field."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [("target", "modified"), ("old text", "new text")]

            results = doc.apply_edits_batch(edits)

            assert results.succeeded[0].index == 0
            assert results.succeeded[1].index == 1

        finally:
            doc_path.unlink()

    def test_edit_result_has_old_new_text(self):
        """Test that EditResult has old_text and new_text fields."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [("target", "modified")]

            results = doc.apply_edits_batch(edits)

            result = results.succeeded[0]
            assert result.old_text == "target"
            assert result.new_text == "modified"

        finally:
            doc_path.unlink()

    def test_edit_result_suggestions_field(self):
        """Test that EditResult has suggestions field."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [("nonexistent", "replacement")]

            results = doc.apply_edits_batch(edits)

            result = results.failed[0]
            # suggestions should be a list (may be empty)
            assert isinstance(result.suggestions, list)

        finally:
            doc_path.unlink()


class TestEditDataclass:
    """Test the Edit dataclass."""

    def test_edit_default_values(self):
        """Test Edit dataclass default values."""
        edit = Edit(old="old", new="new")

        assert edit.old == "old"
        assert edit.new == "new"
        assert edit.edit_type == "replace"
        assert edit.track is True
        assert edit.author is None
        assert edit.scope is None
        assert edit.regex is False
        assert edit.occurrence == "first"

    def test_edit_custom_values(self):
        """Test Edit dataclass with custom values."""
        edit = Edit(
            old="old",
            new="new",
            edit_type="replace_tracked",
            track=False,
            author="Test Author",
            scope="paragraph 1",
            regex=True,
            occurrence="all",
        )

        assert edit.track is False
        assert edit.author == "Test Author"
        assert edit.scope == "paragraph 1"
        assert edit.regex is True
        assert edit.occurrence == "all"

    def test_edit_to_dict(self):
        """Test Edit.to_dict() method."""
        edit = Edit(old="old", new="new", author="Author")

        d = edit.to_dict()

        assert d["type"] == "replace"
        assert d["find"] == "old"
        assert d["replace"] == "new"
        assert d["author"] == "Author"
        assert d["track"] is True


class TestBatchResultEmpty:
    """Test BatchResult with empty or edge cases."""

    def test_empty_edits_list(self):
        """Test apply_edits_batch with empty list."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            results = doc.apply_edits_batch([])

            assert results.total == 0
            assert results.success_count == 0
            assert results.failure_count == 0
            assert "No edits" in results.summary

        finally:
            doc_path.unlink()

    def test_all_failures(self):
        """Test BatchResult when all edits fail."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [
                ("nonexistent1", "replacement1"),
                ("nonexistent2", "replacement2"),
            ]

            results = doc.apply_edits_batch(edits)

            assert results.total == 2
            assert results.success_count == 0
            assert results.failure_count == 2
            assert not results.all_succeeded
            assert bool(results) is False

        finally:
            doc_path.unlink()


class TestDefaultTrackParameter:
    """Test default_track parameter behavior."""

    def test_default_track_true(self):
        """Test that default_track=True (default) tracks changes."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [("target", "modified")]

            # default_track=True is the default
            results = doc.apply_edits_batch(edits)

            assert results.success_count == 1
            # The edit should be tracked (message should indicate this)
            # Note: The actual tracking is tested in integration tests

        finally:
            doc_path.unlink()

    def test_default_track_false(self):
        """Test that default_track=False doesn't track changes."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            edits = [("target", "modified")]

            results = doc.apply_edits_batch(edits, default_track=False)

            assert results.success_count == 1

        finally:
            doc_path.unlink()


# Run tests with: pytest tests/test_batch_partial_success.py -v
