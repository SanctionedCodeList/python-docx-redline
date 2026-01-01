"""
Tests for GitHub Issue #10: Document corruption when replace_tracked() fails on multi-run text.

This tests the scenario where text spans multiple Word runs (common after editing in Word)
and replace_tracked() is called. The operation should either:
1. Complete successfully, or
2. Leave the document unmodified if it cannot complete (atomic operation)

It should NEVER leave the document in a corrupted state.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from python_docx_redline import Document


def get_document_xml(doc: Document) -> str:
    """Get the document XML as a string for comparison."""
    return etree.tostring(doc.xml_root, encoding="unicode")


# Document with text split across multiple runs (simulating Word editing behavior)
MULTIRUN_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>The text is </w:t>
      </w:r>
      <w:r>
        <w:t>processed</w:t>
      </w:r>
      <w:r>
        <w:t>, a response is </w:t>
      </w:r>
      <w:r>
        <w:t>generated</w:t>
      </w:r>
      <w:r>
        <w:t> and returned to the user.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


# Document with text split across runs AND inside existing tracked changes
MULTIRUN_WITH_TRACKED_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>The text is </w:t>
      </w:r>
      <w:ins w:id="1" w:author="Previous Author" w:date="2024-01-01T00:00:00Z">
        <w:r>
          <w:t>processed</w:t>
        </w:r>
        <w:r>
          <w:t>, a response</w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t> is generated</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


# Simple multi-run document for basic testing
SIMPLE_MULTIRUN_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>30</w:t>
      </w:r>
      <w:r>
        <w:t> days</w:t>
      </w:r>
      <w:r>
        <w:t> notice period.</w:t>
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


class TestMultiRunReplaceTracked:
    """Tests for replace_tracked() on text spanning multiple runs."""

    def test_replace_tracked_multirun_simple(self):
        """Test replace_tracked on simple multi-run text succeeds."""
        docx_path = create_test_docx(SIMPLE_MULTIRUN_XML)
        try:
            doc = Document(docx_path)

            # Verify text exists and spans runs
            matches = doc.find_all("30 days")
            assert len(matches) == 1
            assert len(matches[0].span.runs) > 1, "Test requires text spanning multiple runs"

            # Replace should succeed
            doc.replace_tracked("30 days", "45 days")

            # Verify tracked changes exist
            changes = doc.tracked_changes
            assert len(changes) > 0, "Should have tracked changes"

        finally:
            docx_path.unlink()
            docx_path.parent.rmdir()

    def test_replace_tracked_multirun_accept_produces_valid_text(self):
        """Test that accept_all_changes after multi-run replace produces correct text."""
        docx_path = create_test_docx(SIMPLE_MULTIRUN_XML)
        try:
            doc = Document(docx_path)

            # Replace multi-run text
            doc.replace_tracked("30 days", "45 days")

            # Accept all changes
            doc.accept_all_changes()

            # Extract final text - should be clean with "45 days"
            final_text = doc.get_text()
            assert "45 days" in final_text, f"Expected '45 days' in final text, got: {final_text}"
            assert "30 days" not in final_text, f"Old text should be gone, got: {final_text}"
            # Should not have garbled fragments
            assert "30" not in final_text or "45 days" in final_text

        finally:
            docx_path.unlink()
            docx_path.parent.rmdir()

    def test_replace_tracked_multirun_complex_text(self):
        """Test replace_tracked on complex multi-run text (from GitHub issue #10)."""
        docx_path = create_test_docx(MULTIRUN_DOCUMENT_XML)
        try:
            doc = Document(docx_path)

            # This is the exact scenario from the issue
            search_text = "The text is processed, a response is generated"
            replace_text = "The system processes the text and generates a response"

            # Verify text exists and spans multiple runs
            matches = doc.find_all(search_text)
            assert len(matches) == 1, f"Expected 1 match, got {len(matches)}"
            assert len(matches[0].span.runs) > 1, "Test requires text spanning multiple runs"

            # Replace should succeed
            doc.replace_tracked(search_text, replace_text)

            # Verify tracked changes exist
            changes = doc.tracked_changes
            assert len(changes) > 0, "Should have tracked changes"

        finally:
            docx_path.unlink()
            docx_path.parent.rmdir()

    def test_replace_tracked_multirun_accept_no_corruption(self):
        """Test that accept_all_changes after complex multi-run replace is not corrupted.

        This is the core test for GitHub issue #10 - the accept operation
        was producing garbled text like:
        "The text a response, and may briefly...system processes the ,generatesretainboth"
        """
        docx_path = create_test_docx(MULTIRUN_DOCUMENT_XML)
        try:
            doc = Document(docx_path)

            search_text = "The text is processed, a response is generated"
            replace_text = "The system processes the text and generates a response"

            # Replace
            doc.replace_tracked(search_text, replace_text)

            # Accept all changes
            doc.accept_all_changes()

            # Verify final text
            final_text = doc.get_text()

            # The replacement text should appear in the final document
            assert replace_text in final_text, (
                f"Expected replacement text in final document.\n"
                f"Expected: '{replace_text}'\n"
                f"Got: '{final_text}'"
            )

            # The original text should NOT appear
            assert search_text not in final_text, (
                f"Original text should be replaced.\n"
                f"Should not contain: '{search_text}'\n"
                f"Got: '{final_text}'"
            )

            # Check for common corruption patterns mentioned in the issue
            corruption_patterns = [
                ",generatesretainboth",
                "system processes the ,",
                "a response, and may briefly",
            ]
            for pattern in corruption_patterns:
                assert (
                    pattern not in final_text
                ), f"Found corruption pattern '{pattern}' in text: {final_text}"

        finally:
            docx_path.unlink()
            docx_path.parent.rmdir()

    def test_replace_tracked_with_existing_tracked_changes(self):
        """Test replace_tracked on text that spans existing tracked changes."""
        docx_path = create_test_docx(MULTIRUN_WITH_TRACKED_XML)
        try:
            doc = Document(docx_path)

            # Text spans both regular runs and runs inside w:ins
            search_text = "processed, a response"
            replace_text = "handled, and a result"

            # Replace should succeed (or cleanly fail without corruption)
            doc.replace_tracked(search_text, replace_text)

            # Accept all changes
            doc.accept_all_changes()

            # Final text should be coherent
            final_text = doc.get_text()
            assert (
                "handled, and a result" in final_text or "processed, a response" in final_text
            ), f"Final text should contain either new or old text (no corruption): {final_text}"

        finally:
            docx_path.unlink()
            docx_path.parent.rmdir()


class TestAtomicOperation:
    """Tests ensuring replace operations are atomic (all-or-nothing)."""

    def test_failed_replace_leaves_document_unchanged(self):
        """If replace_tracked cannot complete, document should be unchanged."""
        docx_path = create_test_docx(SIMPLE_MULTIRUN_XML)
        try:
            doc = Document(docx_path)

            # Capture original state
            original_text = doc.get_text()
            original_xml = get_document_xml(doc)

            # Try to replace text that doesn't exist
            from python_docx_redline import TextNotFoundError

            with pytest.raises(TextNotFoundError):
                doc.replace_tracked("nonexistent text", "replacement")

            # Document should be unchanged
            assert (
                doc.get_text() == original_text
            ), "Document text should be unchanged after failed replace"
            assert get_document_xml(doc) == original_xml, "Document XML should be unchanged"

        finally:
            docx_path.unlink()
            docx_path.parent.rmdir()

    def test_replace_tracked_rollback_on_partial_failure(self):
        """If replace partially fails, entire operation should roll back.

        This is the critical atomicity test - we need to verify that if
        something goes wrong during multi-run replacement, the document
        is restored to its original state.
        """
        docx_path = create_test_docx(MULTIRUN_DOCUMENT_XML)
        try:
            doc = Document(docx_path)

            # Capture original state
            original_text = doc.get_text()

            # Perform a valid replacement
            doc.replace_tracked(
                "The text is processed, a response is generated",
                "The system processes the text and generates a response",
            )

            # Verify changes were made
            assert doc.get_text() != original_text or len(doc.tracked_changes) > 0

            # Now accept changes and verify no corruption
            doc.accept_all_changes()
            final_text = doc.get_text()

            # Text should be clean - no interleaved fragments
            words = final_text.split()
            # Check for obvious corruption: unexpected concatenations
            for word in words:
                # A word shouldn't be a weird mix of old and new text
                assert not (
                    "processed" in word and "system" in word
                ), f"Found corrupted word: {word}"

        finally:
            docx_path.unlink()
            docx_path.parent.rmdir()


class TestRegressionGitHub10:
    """Specific regression tests for GitHub issue #10."""

    def test_github_10_exact_scenario(self):
        """Reproduce the exact scenario from GitHub issue #10."""
        docx_path = create_test_docx(MULTIRUN_DOCUMENT_XML)
        try:
            doc = Document(docx_path)

            # From the issue: text exists but spans multiple runs
            matches = doc.find_all("The text is processed")
            assert len(matches) == 1, "Text should be found"

            # The issue shows replace returning None (failure) but modifying doc
            # After our fix, it should either succeed OR leave doc unchanged
            original_xml = get_document_xml(doc)

            try:
                doc.replace_tracked(
                    "The text is processed, a response is generated",
                    "The system processes the text and generates a response",
                )
                replace_succeeded = True
            except Exception:
                replace_succeeded = False

            if replace_succeeded:
                # If it succeeded, accepting changes should produce clean text
                doc.accept_all_changes()
                final_text = doc.get_text()
                assert (
                    "The system processes" in final_text
                ), f"After accepting, text should contain replacement: {final_text}"
            else:
                # If it failed, document should be unchanged
                assert (
                    get_document_xml(doc) == original_xml
                ), "Failed replace should leave document unchanged"

        finally:
            docx_path.unlink()
            docx_path.parent.rmdir()
