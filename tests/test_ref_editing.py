"""
Tests for ref-based editing operations.

These tests verify the Document methods that use refs for editing:
- resolve_ref()
- get_ref()
- insert_at_ref()
- delete_ref()
- replace_at_ref()
- add_comment_at_ref()
"""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from python_docx_redline import Document
from python_docx_redline.accessibility.types import ElementType
from python_docx_redline.constants import WORD_NAMESPACE
from python_docx_redline.errors import RefNotFoundError

# ============================================================================
# Test fixtures and helpers
# ============================================================================

MINIMAL_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph content.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Second paragraph is a heading.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph for testing.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_TABLE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Intro paragraph before table.</w:t>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 0,0 content</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 0,1 content</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 1,0 content</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 1,1 content</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:t>Final paragraph after table.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


def create_test_docx(content: str = MINIMAL_DOCUMENT_XML) -> Path:
    """Create a minimal but valid OOXML test .docx file.

    Args:
        content: The document.xml content

    Returns:
        Path to the created .docx file
    """
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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


def get_text_content(element: etree._Element) -> str:
    """Extract text content from an element."""
    text_parts = []
    for t_elem in element.iter(f"{{{WORD_NAMESPACE}}}t"):
        if t_elem.text:
            text_parts.append(t_elem.text)
    return "".join(text_parts)


# ============================================================================
# Test resolve_ref()
# ============================================================================


class TestResolveRef:
    """Tests for Document.resolve_ref()."""

    def test_resolve_paragraph_ref(self) -> None:
        """Test resolving a paragraph ref."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            element = doc.resolve_ref("p:0")

            assert element is not None
            assert element.tag == f"{{{WORD_NAMESPACE}}}p"
            text = get_text_content(element)
            assert "First paragraph" in text
        finally:
            docx_path.unlink()

    def test_resolve_second_paragraph(self) -> None:
        """Test resolving the second paragraph ref."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            element = doc.resolve_ref("p:1")

            text = get_text_content(element)
            assert "Second paragraph" in text
        finally:
            docx_path.unlink()

    def test_resolve_out_of_bounds_raises(self) -> None:
        """Test that out of bounds ref raises RefNotFoundError."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(RefNotFoundError, match="out of bounds"):
                doc.resolve_ref("p:99")
        finally:
            docx_path.unlink()

    def test_resolve_table_ref(self) -> None:
        """Test resolving a table ref."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            element = doc.resolve_ref("tbl:0")

            assert element is not None
            assert element.tag == f"{{{WORD_NAMESPACE}}}tbl"
        finally:
            docx_path.unlink()

    def test_resolve_table_cell_ref(self) -> None:
        """Test resolving a table cell ref."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            element = doc.resolve_ref("tbl:0/row:0/cell:1")

            assert element is not None
            assert element.tag == f"{{{WORD_NAMESPACE}}}tc"
            text = get_text_content(element)
            assert "Cell 0,1" in text
        finally:
            docx_path.unlink()

    def test_resolve_invalid_ref_format(self) -> None:
        """Test that invalid ref format raises error."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(ValueError, match="Invalid"):
                doc.resolve_ref("invalid_ref_format")
        finally:
            docx_path.unlink()


# ============================================================================
# Test get_ref()
# ============================================================================


class TestGetRef:
    """Tests for Document.get_ref()."""

    def test_get_ref_for_paragraph(self) -> None:
        """Test getting ref for a paragraph element."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Get the second paragraph element
            element = doc.resolve_ref("p:1")
            ref = doc.get_ref(element)

            assert ref.path == "p:1"
            assert ref.element_type == ElementType.PARAGRAPH
        finally:
            docx_path.unlink()

    def test_get_ref_roundtrip(self) -> None:
        """Test that get_ref and resolve_ref are inverses."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Resolve a ref
            original_ref = "p:2"
            element = doc.resolve_ref(original_ref)

            # Get the ref back
            ref = doc.get_ref(element)

            # Verify it matches
            assert ref.path == original_ref
        finally:
            docx_path.unlink()

    def test_get_ref_with_fingerprint(self) -> None:
        """Test getting a fingerprint-based ref."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            element = doc.resolve_ref("p:0")
            ref = doc.get_ref(element, use_fingerprint=True)

            assert ref.is_fingerprint
            assert ref.path.startswith("p:~")
        finally:
            docx_path.unlink()

    def test_fingerprint_ref_resolves(self) -> None:
        """Test that fingerprint ref can be resolved back."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Get element and fingerprint ref
            element = doc.resolve_ref("p:1")
            fp_ref = doc.get_ref(element, use_fingerprint=True)

            # Resolve the fingerprint ref
            resolved = doc.resolve_ref(fp_ref.path)

            assert resolved is element
        finally:
            docx_path.unlink()


# ============================================================================
# Test get_text_at_ref()
# ============================================================================


class TestGetTextAtRef:
    """Tests for Document.get_text_at_ref()."""

    def test_get_text_from_paragraph_ref(self) -> None:
        """Test getting text from a paragraph ref."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            text = doc.get_text_at_ref("p:0")

            assert text == "First paragraph content."
        finally:
            docx_path.unlink()

    def test_get_text_from_second_paragraph(self) -> None:
        """Test getting text from the second paragraph."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            text = doc.get_text_at_ref("p:1")

            assert text == "Second paragraph is a heading."
        finally:
            docx_path.unlink()

    def test_get_text_from_table_cell(self) -> None:
        """Test getting text from a table cell ref."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            text = doc.get_text_at_ref("tbl:0/row:0/cell:1")

            assert text == "Cell 0,1 content"
        finally:
            docx_path.unlink()

    def test_get_text_from_paragraph_in_table_cell(self) -> None:
        """Test getting text from a paragraph inside a table cell."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            text = doc.get_text_at_ref("tbl:0/row:1/cell:0/p:0")

            assert text == "Cell 1,0 content"
        finally:
            docx_path.unlink()

    def test_get_text_from_table_ref(self) -> None:
        """Test getting text from entire table ref returns all cell text."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            text = doc.get_text_at_ref("tbl:0")

            # Should include all text from all cells
            assert "Cell 0,0 content" in text
            assert "Cell 0,1 content" in text
            assert "Cell 1,0 content" in text
            assert "Cell 1,1 content" in text
        finally:
            docx_path.unlink()

    def test_get_text_invalid_ref_raises(self) -> None:
        """Test that invalid ref raises RefNotFoundError."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(RefNotFoundError, match="out of bounds"):
                doc.get_text_at_ref("p:99")
        finally:
            docx_path.unlink()

    def test_get_text_invalid_format_raises(self) -> None:
        """Test that invalid ref format raises error."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(ValueError, match="Invalid"):
                doc.get_text_at_ref("invalid_ref")
        finally:
            docx_path.unlink()


# ============================================================================
# Test insert_at_ref() - untracked mode
# ============================================================================


class TestInsertAtRefUntracked:
    """Tests for Document.insert_at_ref() without tracking."""

    def test_insert_after_paragraph_untracked(self) -> None:
        """Test inserting a new paragraph after a ref (untracked)."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_at_ref("p:0", "Inserted paragraph.", position="after", track=False)

            assert result.success
            # Verify the paragraph was inserted
            new_element = doc.resolve_ref("p:1")
            text = get_text_content(new_element)
            assert text == "Inserted paragraph."
        finally:
            docx_path.unlink()

    def test_insert_before_paragraph_untracked(self) -> None:
        """Test inserting a new paragraph before a ref (untracked)."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_at_ref("p:1", "Before heading.", position="before", track=False)

            assert result.success
            # The new paragraph should now be at p:1
            element = doc.resolve_ref("p:1")
            text = get_text_content(element)
            assert text == "Before heading."
        finally:
            docx_path.unlink()

    def test_insert_at_start_untracked(self) -> None:
        """Test inserting text at the start of a paragraph (untracked)."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_at_ref("p:0", "PREFIX: ", position="start", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert text.startswith("PREFIX: ")
        finally:
            docx_path.unlink()

    def test_insert_at_end_untracked(self) -> None:
        """Test inserting text at the end of a paragraph (untracked)."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_at_ref("p:0", " SUFFIX", position="end", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert text.endswith(" SUFFIX")
        finally:
            docx_path.unlink()

    def test_insert_invalid_position_raises(self) -> None:
        """Test that invalid position raises ValueError."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(ValueError, match="position must be one of"):
                doc.insert_at_ref("p:0", "text", position="invalid")
        finally:
            docx_path.unlink()


# ============================================================================
# Test insert_at_ref() - tracked mode
# ============================================================================


class TestInsertAtRefTracked:
    """Tests for Document.insert_at_ref() with tracking."""

    def test_insert_after_paragraph_tracked(self) -> None:
        """Test inserting a tracked paragraph after a ref."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_at_ref(
                "p:0", "New tracked paragraph.", position="after", track=True, author="TestAgent"
            )

            assert result.success
            assert doc.has_tracked_changes()

            # Verify the insertion element exists
            new_element = doc.resolve_ref("p:1")
            ins_elem = new_element.find(f".//{{{WORD_NAMESPACE}}}ins")
            assert ins_elem is not None
            assert ins_elem.get(f"{{{WORD_NAMESPACE}}}author") == "TestAgent"
        finally:
            docx_path.unlink()

    def test_insert_at_end_tracked(self) -> None:
        """Test inserting tracked text at end of paragraph."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_at_ref("p:0", " (AMENDED)", position="end", track=True)

            assert result.success
            assert doc.has_tracked_changes()

            # Verify the insertion is at the end
            element = doc.resolve_ref("p:0")
            ins_elem = element.find(f"./{{{WORD_NAMESPACE}}}ins")
            assert ins_elem is not None
        finally:
            docx_path.unlink()

    def test_insert_at_start_tracked(self) -> None:
        """Test inserting tracked text at start of paragraph."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_at_ref("p:0", "IMPORTANT: ", position="start", track=True)

            assert result.success
            # First child should be the insertion
            element = doc.resolve_ref("p:0")
            first_child = element[0]
            assert first_child.tag == f"{{{WORD_NAMESPACE}}}ins"
        finally:
            docx_path.unlink()


# ============================================================================
# Test delete_ref() - untracked mode
# ============================================================================


class TestDeleteRefUntracked:
    """Tests for Document.delete_ref() without tracking."""

    def test_delete_paragraph_untracked(self) -> None:
        """Test deleting a paragraph without tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Get text of p:1 before deletion
            element_before = doc.resolve_ref("p:1")
            original_text = get_text_content(element_before)
            assert "Second paragraph" in original_text

            result = doc.delete_ref("p:1", track=False)

            assert result.success
            # Now p:1 should be what was p:2
            element_after = doc.resolve_ref("p:1")
            new_text = get_text_content(element_after)
            assert "Third paragraph" in new_text
        finally:
            docx_path.unlink()

    def test_delete_paragraph_reduces_count(self) -> None:
        """Test that deleting reduces paragraph count."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Count paragraphs before
            count_before = len(doc.paragraphs)

            doc.delete_ref("p:0", track=False)

            count_after = len(doc.paragraphs)
            assert count_after == count_before - 1
        finally:
            docx_path.unlink()


# ============================================================================
# Test delete_ref() - tracked mode
# ============================================================================


class TestDeleteRefTracked:
    """Tests for Document.delete_ref() with tracking."""

    def test_delete_paragraph_tracked(self) -> None:
        """Test deleting a paragraph with tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.delete_ref("p:0", track=True, author="TestAgent")

            assert result.success
            assert doc.has_tracked_changes()

            # The paragraph should still exist but with deletion markers
            element = doc.resolve_ref("p:0")
            del_elem = element.find(f".//{{{WORD_NAMESPACE}}}del")
            assert del_elem is not None
            assert del_elem.get(f"{{{WORD_NAMESPACE}}}author") == "TestAgent"
        finally:
            docx_path.unlink()

    def test_delete_tracked_has_deltext(self) -> None:
        """Test that tracked deletion uses w:delText."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.delete_ref("p:0", track=True)

            element = doc.resolve_ref("p:0")
            del_text = element.find(f".//{{{WORD_NAMESPACE}}}delText")
            assert del_text is not None
        finally:
            docx_path.unlink()


# ============================================================================
# Test replace_at_ref() - untracked mode
# ============================================================================


class TestReplaceAtRefUntracked:
    """Tests for Document.replace_at_ref() without tracking."""

    def test_replace_paragraph_untracked(self) -> None:
        """Test replacing paragraph content without tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.replace_at_ref("p:0", "Completely new content.", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert text == "Completely new content."
        finally:
            docx_path.unlink()

    def test_replace_removes_old_content(self) -> None:
        """Test that replace removes old content."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.replace_at_ref("p:0", "New text only.", track=False)

            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert "First paragraph" not in text
            assert text == "New text only."
        finally:
            docx_path.unlink()


# ============================================================================
# Test replace_at_ref() - tracked mode
# ============================================================================


class TestReplaceAtRefTracked:
    """Tests for Document.replace_at_ref() with tracking."""

    def test_replace_paragraph_tracked(self) -> None:
        """Test replacing paragraph content with tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.replace_at_ref("p:0", "Replacement text.", track=True, author="Reviewer")

            assert result.success
            assert doc.has_tracked_changes()

            element = doc.resolve_ref("p:0")
            # Should have both deletion and insertion
            del_elem = element.find(f".//{{{WORD_NAMESPACE}}}del")
            ins_elem = element.find(f".//{{{WORD_NAMESPACE}}}ins")
            assert del_elem is not None
            assert ins_elem is not None
        finally:
            docx_path.unlink()

    def test_replace_tracked_preserves_old_as_deleted(self) -> None:
        """Test that tracked replace shows old text as deleted."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.replace_at_ref("p:0", "New content.", track=True)

            element = doc.resolve_ref("p:0")
            del_text = element.find(f".//{{{WORD_NAMESPACE}}}delText")
            assert del_text is not None
            assert "First paragraph" in (del_text.text or "")
        finally:
            docx_path.unlink()

    def test_replace_tracked_shows_new_as_inserted(self) -> None:
        """Test that tracked replace shows new text as inserted."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.replace_at_ref("p:0", "Brand new content.", track=True)

            element = doc.resolve_ref("p:0")
            ins_elem = element.find(f".//{{{WORD_NAMESPACE}}}ins")
            assert ins_elem is not None
            t_elem = ins_elem.find(f".//{{{WORD_NAMESPACE}}}t")
            assert t_elem is not None
            assert t_elem.text == "Brand new content."
        finally:
            docx_path.unlink()


# ============================================================================
# Test add_comment_at_ref()
# ============================================================================


class TestAddCommentAtRef:
    """Tests for Document.add_comment_at_ref()."""

    def test_add_comment_at_ref_basic(self) -> None:
        """Test adding a comment at a ref."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            comment = doc.add_comment_at_ref("p:0", "Please review this.", author="Reviewer")

            assert comment is not None
            assert comment.text == "Please review this."
        finally:
            docx_path.unlink()

    def test_add_comment_appears_in_comments_list(self) -> None:
        """Test that added comment appears in document comments."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.add_comment_at_ref("p:1", "Check this heading.", author="Editor")

            comments = doc.comments
            assert len(comments) >= 1
            comment_texts = [c.text for c in comments]
            assert "Check this heading." in comment_texts
        finally:
            docx_path.unlink()


# ============================================================================
# Test table cell operations
# ============================================================================


class TestRefEditingWithTables:
    """Tests for ref-based editing with table elements."""

    def test_resolve_table_cell_content(self) -> None:
        """Test resolving and reading table cell content."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            cell = doc.resolve_ref("tbl:0/row:1/cell:1")
            text = get_text_content(cell)

            assert "Cell 1,1" in text
        finally:
            docx_path.unlink()


# ============================================================================
# Test integration scenarios
# ============================================================================


class TestRefEditingIntegration:
    """Integration tests for ref-based editing workflows."""

    def test_find_and_edit_by_ref(self) -> None:
        """Test finding text, getting ref, then editing by ref."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Find text
            matches = doc.find_all("Second paragraph")
            assert len(matches) > 0

            # Get ref for the paragraph containing the match
            para_element = matches[0].span.paragraph
            ref = doc.get_ref(para_element)

            # Edit by ref
            result = doc.insert_at_ref(ref.path, " (Updated)", position="end", track=True)

            assert result.success
            element = doc.resolve_ref(ref.path)
            text = get_text_content(element)
            assert "(Updated)" in text
        finally:
            docx_path.unlink()

    def test_multiple_edits_with_cache_invalidation(self) -> None:
        """Test that multiple edits work correctly with cache invalidation."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Insert a paragraph
            doc.insert_at_ref("p:0", "New first paragraph.", position="before", track=False)

            # The old p:0 should now be p:1
            element = doc.resolve_ref("p:1")
            text = get_text_content(element)
            assert "First paragraph" in text

            # Insert another paragraph
            doc.insert_at_ref("p:2", "After second.", position="after", track=False)

            # Verify all paragraphs are accessible
            p0 = doc.resolve_ref("p:0")
            p1 = doc.resolve_ref("p:1")
            p2 = doc.resolve_ref("p:2")
            p3 = doc.resolve_ref("p:3")

            assert get_text_content(p0) == "New first paragraph."
            assert "First paragraph" in get_text_content(p1)
            assert "Second paragraph" in get_text_content(p2)
            assert get_text_content(p3) == "After second."
        finally:
            docx_path.unlink()

    def test_edit_save_and_reload(self) -> None:
        """Test that edits persist after save and reload."""
        docx_path = create_test_docx()
        try:
            # Make edits
            doc = Document(docx_path)
            doc.insert_at_ref("p:0", " [EDITED]", position="end", track=True)
            doc.save(docx_path)

            # Reload and verify
            doc2 = Document(docx_path)
            element = doc2.resolve_ref("p:0")
            text = get_text_content(element)
            assert "[EDITED]" in text
            assert doc2.has_tracked_changes()
        finally:
            docx_path.unlink()


# ============================================================================
# Test error handling
# ============================================================================


class TestRefEditingErrors:
    """Tests for error handling in ref-based editing."""

    def test_resolve_nonexistent_ref(self) -> None:
        """Test that resolving nonexistent ref raises error."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(RefNotFoundError):
                doc.resolve_ref("p:100")
        finally:
            docx_path.unlink()

    def test_insert_at_invalid_ref(self) -> None:
        """Test that inserting at invalid ref raises error."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(RefNotFoundError):
                doc.insert_at_ref("p:999", "text", position="end")
        finally:
            docx_path.unlink()

    def test_delete_invalid_ref(self) -> None:
        """Test that deleting invalid ref raises error."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(RefNotFoundError):
                doc.delete_ref("p:999")
        finally:
            docx_path.unlink()

    def test_replace_invalid_ref(self) -> None:
        """Test that replacing at invalid ref raises error."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(RefNotFoundError):
                doc.replace_at_ref("p:999", "new text")
        finally:
            docx_path.unlink()


# ============================================================================
# Test insert_in_ref() - untracked mode
# ============================================================================


class TestInsertInRefUntracked:
    """Tests for Document.insert_in_ref() without tracking."""

    def test_insert_after_anchor_untracked(self) -> None:
        """Test inserting text after an anchor within an element."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_in_ref("p:0", " [ADDED]", after="First paragraph", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            # The anchor "First paragraph" is found, and text is inserted after it
            assert "[ADDED]" in text
            assert "First paragraph" in text
        finally:
            docx_path.unlink()

    def test_insert_before_anchor_untracked(self) -> None:
        """Test inserting text before an anchor within an element."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_in_ref("p:0", "[PREFIX] ", before="First", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert "[PREFIX] First paragraph" in text
        finally:
            docx_path.unlink()

    def test_insert_in_ref_partial_anchor(self) -> None:
        """Test inserting after a partial match of the paragraph text."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Insert after just "content" (partial match)
            result = doc.insert_in_ref("p:0", " [SUFFIX]", after="content", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            # The text "content" is part of "First paragraph content."
            # Insertion happens after that word
            assert "[SUFFIX]" in text
            assert "content" in text
        finally:
            docx_path.unlink()


# ============================================================================
# Test insert_in_ref() - tracked mode
# ============================================================================


class TestInsertInRefTracked:
    """Tests for Document.insert_in_ref() with tracking."""

    def test_insert_after_anchor_tracked(self) -> None:
        """Test inserting tracked text after an anchor within an element."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_in_ref(
                "p:0", " (amended)", after="content", track=True, author="TestAgent"
            )

            assert result.success
            assert doc.has_tracked_changes()

            # Verify the insertion marker exists
            element = doc.resolve_ref("p:0")
            ins_elem = element.find(f".//{{{WORD_NAMESPACE}}}ins")
            assert ins_elem is not None
            assert ins_elem.get(f"{{{WORD_NAMESPACE}}}author") == "TestAgent"
        finally:
            docx_path.unlink()

    def test_insert_before_anchor_tracked(self) -> None:
        """Test inserting tracked text before an anchor within an element."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.insert_in_ref(
                "p:1", "Important: ", before="Second", track=True, author="Editor"
            )

            assert result.success
            assert doc.has_tracked_changes()

            element = doc.resolve_ref("p:1")
            text = get_text_content(element)
            # The inserted text should appear in the document
            assert "Important:" in text or "Second" in text
        finally:
            docx_path.unlink()


# ============================================================================
# Test insert_in_ref() with tables
# ============================================================================


class TestInsertInRefWithTables:
    """Tests for insert_in_ref() with table elements."""

    def test_insert_in_table_cell(self) -> None:
        """Test inserting text within a table cell."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            result = doc.insert_in_ref(
                "tbl:0/row:0/cell:0", " [EDITED]", after="Cell 0,0", track=False
            )

            assert result.success
            cell = doc.resolve_ref("tbl:0/row:0/cell:0")
            text = get_text_content(cell)
            # The anchor "Cell 0,0" is found, and text is inserted after it
            assert "[EDITED]" in text
            assert "Cell 0,0" in text
        finally:
            docx_path.unlink()

    def test_insert_in_table_cell_tracked(self) -> None:
        """Test inserting tracked text within a table cell."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            result = doc.insert_in_ref(
                "tbl:0/row:1/cell:1", " (updated)", after="content", track=True
            )

            assert result.success
            assert doc.has_tracked_changes()
        finally:
            docx_path.unlink()


# ============================================================================
# Test insert_in_ref() error handling
# ============================================================================


class TestInsertInRefErrors:
    """Tests for error handling in insert_in_ref()."""

    def test_insert_in_ref_no_anchor_params(self) -> None:
        """Test that missing both before and after raises ValueError."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(ValueError, match="Must specify either"):
                doc.insert_in_ref("p:0", "text")
        finally:
            docx_path.unlink()

    def test_insert_in_ref_both_anchor_params(self) -> None:
        """Test that specifying both before and after raises ValueError."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(ValueError, match="Cannot specify both"):
                doc.insert_in_ref("p:0", "text", before="a", after="b")
        finally:
            docx_path.unlink()

    def test_insert_in_ref_anchor_not_found(self) -> None:
        """Test that anchor text not in element raises TextNotFoundError."""
        from python_docx_redline.errors import TextNotFoundError

        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(TextNotFoundError):
                doc.insert_in_ref("p:0", "text", after="nonexistent anchor")
        finally:
            docx_path.unlink()

    def test_insert_in_ref_invalid_ref(self) -> None:
        """Test that invalid ref raises RefNotFoundError."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(RefNotFoundError):
                doc.insert_in_ref("p:999", "text", after="anchor")
        finally:
            docx_path.unlink()

    def test_insert_in_ref_ambiguous_anchor(self) -> None:
        """Test that ambiguous anchor raises AmbiguousTextError."""
        from python_docx_redline.errors import AmbiguousTextError

        # Create a document with repeated text in a single paragraph
        repeated_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>word word word</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(repeated_xml)
        try:
            doc = Document(docx_path)

            with pytest.raises(AmbiguousTextError):
                doc.insert_in_ref("p:0", " [insert]", after="word")
        finally:
            docx_path.unlink()


# ============================================================================
# Test insert_in_ref() integration scenarios
# ============================================================================


class TestInsertInRefIntegration:
    """Integration tests for insert_in_ref()."""

    def test_insert_in_ref_save_and_reload(self) -> None:
        """Test that insert_in_ref edits persist after save and reload."""
        docx_path = create_test_docx()
        try:
            # Make edit
            doc = Document(docx_path)
            doc.insert_in_ref("p:0", " [EDITED]", after="content", track=True)
            doc.save(docx_path)

            # Reload and verify
            doc2 = Document(docx_path)
            element = doc2.resolve_ref("p:0")
            text = get_text_content(element)
            assert "[EDITED]" in text
            assert doc2.has_tracked_changes()
        finally:
            docx_path.unlink()

    def test_insert_in_ref_combined_with_insert_at_ref(self) -> None:
        """Test using both insert_in_ref and insert_at_ref together."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # First, insert within a paragraph (after "paragraph" word)
            doc.insert_in_ref("p:0", " [INLINE]", after="paragraph", track=False)

            # Then, insert a new paragraph after
            doc.insert_at_ref("p:0", "New paragraph.", position="after", track=False)

            # Verify inline edit - the [INLINE] should appear after "paragraph"
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert "[INLINE]" in text
            assert "paragraph" in text

            # Verify new paragraph
            new_para = doc.resolve_ref("p:1")
            new_text = get_text_content(new_para)
            assert new_text == "New paragraph."
        finally:
            docx_path.unlink()


# ============================================================================
# Test delete_in_ref() - untracked mode
# ============================================================================


class TestDeleteInRefUntracked:
    """Tests for Document.delete_in_ref() without tracking."""

    def test_delete_text_in_paragraph_untracked(self) -> None:
        """Test deleting text within a paragraph without tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.delete_in_ref("p:0", "First ", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert text == "paragraph content."
            assert "First" not in text
        finally:
            docx_path.unlink()

    def test_delete_text_at_end_untracked(self) -> None:
        """Test deleting text at the end of a paragraph."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.delete_in_ref("p:0", " content.", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert text == "First paragraph"
        finally:
            docx_path.unlink()

    def test_delete_text_in_middle_untracked(self) -> None:
        """Test deleting text in the middle of a paragraph."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.delete_in_ref("p:0", " paragraph", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert text == "First content."
        finally:
            docx_path.unlink()


# ============================================================================
# Test delete_in_ref() - tracked mode
# ============================================================================


class TestDeleteInRefTracked:
    """Tests for Document.delete_in_ref() with tracking."""

    def test_delete_text_in_paragraph_tracked(self) -> None:
        """Test deleting text within a paragraph with tracking."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.delete_in_ref("p:0", "First ", track=True, author="TestAgent")

            assert result.success
            assert doc.has_tracked_changes()

            # The paragraph should have deletion markers
            element = doc.resolve_ref("p:0")
            del_elem = element.find(f".//{{{WORD_NAMESPACE}}}del")
            assert del_elem is not None
            assert del_elem.get(f"{{{WORD_NAMESPACE}}}author") == "TestAgent"
        finally:
            docx_path.unlink()

    def test_delete_tracked_has_deltext(self) -> None:
        """Test that tracked deletion uses w:delText."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.delete_in_ref("p:0", "First ", track=True)

            element = doc.resolve_ref("p:0")
            del_text = element.find(f".//{{{WORD_NAMESPACE}}}delText")
            assert del_text is not None
            assert del_text.text == "First "
        finally:
            docx_path.unlink()

    def test_delete_tracked_preserves_remaining_text(self) -> None:
        """Test that tracked deletion preserves the remaining text."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.delete_in_ref("p:0", "paragraph ", track=True)

            element = doc.resolve_ref("p:0")
            # Should still have the non-deleted text
            text_parts = []
            for t_elem in element.iter(f"{{{WORD_NAMESPACE}}}t"):
                if t_elem.text:
                    text_parts.append(t_elem.text)
            remaining_text = "".join(text_parts)
            assert "First" in remaining_text or "content" in remaining_text
        finally:
            docx_path.unlink()


# ============================================================================
# Test delete_in_ref() - error handling
# ============================================================================


class TestDeleteInRefErrors:
    """Tests for error handling in delete_in_ref()."""

    def test_delete_invalid_ref_raises(self) -> None:
        """Test that invalid ref raises RefNotFoundError."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(RefNotFoundError):
                doc.delete_in_ref("p:999", "some text")
        finally:
            docx_path.unlink()

    def test_delete_text_not_found_raises(self) -> None:
        """Test that missing text raises TextNotFoundError."""
        from python_docx_redline.errors import TextNotFoundError

        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(TextNotFoundError):
                doc.delete_in_ref("p:0", "nonexistent text")
        finally:
            docx_path.unlink()

    def test_delete_ambiguous_text_raises(self) -> None:
        """Test that multiple occurrences raise AmbiguousTextError."""
        from python_docx_redline.errors import AmbiguousTextError

        # Create a document with duplicate text
        duplicate_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>word word word</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(duplicate_xml)
        try:
            doc = Document(docx_path)

            with pytest.raises(AmbiguousTextError):
                doc.delete_in_ref("p:0", "word")
        finally:
            docx_path.unlink()


# ============================================================================
# Test delete_in_ref() - table cell operations
# ============================================================================


class TestDeleteInRefWithTables:
    """Tests for delete_in_ref() with table elements."""

    def test_delete_text_in_table_cell(self) -> None:
        """Test deleting text within a table cell."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            result = doc.delete_in_ref("tbl:0/row:0/cell:0", "Cell ", track=False)

            assert result.success
            cell = doc.resolve_ref("tbl:0/row:0/cell:0")
            text = get_text_content(cell)
            assert text == "0,0 content"
        finally:
            docx_path.unlink()

    def test_delete_text_in_table_cell_tracked(self) -> None:
        """Test deleting text in a table cell with tracking."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            result = doc.delete_in_ref(
                "tbl:0/row:1/cell:1", " content", track=True, author="TableEditor"
            )

            assert result.success
            assert doc.has_tracked_changes()

            cell = doc.resolve_ref("tbl:0/row:1/cell:1")
            del_elem = cell.find(f".//{{{WORD_NAMESPACE}}}del")
            assert del_elem is not None
        finally:
            docx_path.unlink()


# ============================================================================
# Test delete_in_ref() - integration scenarios
# ============================================================================


class TestDeleteInRefIntegration:
    """Integration tests for delete_in_ref() workflows."""

    def test_find_and_delete_in_ref(self) -> None:
        """Test finding text, getting ref, then deleting within that ref."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # Find the paragraph containing "Second"
            matches = doc.find_all("Second paragraph")
            assert len(matches) > 0

            # Get ref for the paragraph
            para_element = matches[0].span.paragraph
            ref = doc.get_ref(para_element)

            # Delete "Second " from the paragraph
            result = doc.delete_in_ref(ref.path, "Second ", track=True)

            assert result.success
            element = doc.resolve_ref(ref.path)
            del_elem = element.find(f".//{{{WORD_NAMESPACE}}}del")
            assert del_elem is not None
        finally:
            docx_path.unlink()

    def test_delete_saves_correctly(self) -> None:
        """Test that delete_in_ref changes persist after save."""
        docx_path = create_test_docx()
        try:
            # Make deletion
            doc = Document(docx_path)
            doc.delete_in_ref("p:0", "First ", track=True)
            doc.save(docx_path)

            # Reload and verify
            doc2 = Document(docx_path)
            element = doc2.resolve_ref("p:0")
            del_elem = element.find(f".//{{{WORD_NAMESPACE}}}del")
            assert del_elem is not None
            assert doc2.has_tracked_changes()
        finally:
            docx_path.unlink()

    def test_delete_combined_with_insert(self) -> None:
        """Test delete_in_ref combined with insert_in_ref."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # First delete some text
            doc.delete_in_ref("p:0", "paragraph ", track=False)

            # Then insert new text
            doc.insert_in_ref("p:0", "modified ", after="First ", track=False)

            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert "First modified content" in text
            assert "paragraph" not in text
        finally:
            docx_path.unlink()


# ============================================================================
# Test replace_in_ref() - untracked mode
# ============================================================================


class TestReplaceInRefUntracked:
    """Tests for Document.replace_in_ref() without tracking."""

    def test_replace_in_ref_basic(self) -> None:
        """Test basic substring replacement in a paragraph."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.replace_in_ref("p:0", "First", "Updated", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert "Updated paragraph" in text
            assert "First paragraph" not in text
        finally:
            docx_path.unlink()

    def test_replace_in_ref_preserves_other_text(self) -> None:
        """Test that other text in the paragraph is preserved."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.replace_in_ref("p:0", "paragraph", "section", track=False)

            assert result.success
            element = doc.resolve_ref("p:0")
            text = get_text_content(element)
            assert "First section content." in text
        finally:
            docx_path.unlink()

    def test_replace_in_ref_in_table_cell(self) -> None:
        """Test replacement within a table cell ref."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)
        try:
            doc = Document(docx_path)

            result = doc.replace_in_ref("tbl:0/row:0/cell:0", "Cell 0,0", "Updated", track=False)

            assert result.success
            text = doc.get_text_at_ref("tbl:0/row:0/cell:0")
            assert "Updated" in text
            assert "Cell 0,0" not in text
        finally:
            docx_path.unlink()

    def test_replace_in_ref_occurrence_all(self) -> None:
        """Test replacing all occurrences of text in an element."""
        # Create a document with repeated text
        repeated_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Replace foo and foo and foo here.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(repeated_xml)
        try:
            doc = Document(docx_path)

            result = doc.replace_in_ref("p:0", "foo", "bar", occurrence="all", track=False)

            assert result.success
            text = doc.get_text_at_ref("p:0")
            assert "foo" not in text
            assert text.count("bar") == 3
        finally:
            docx_path.unlink()


# ============================================================================
# Test replace_in_ref() - tracked mode
# ============================================================================


class TestReplaceInRefTracked:
    """Tests for Document.replace_in_ref() with tracking."""

    def test_replace_in_ref_tracked(self) -> None:
        """Test tracked substring replacement."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            result = doc.replace_in_ref("p:0", "First", "Modified", track=True, author="TestAgent")

            assert result.success
            assert doc.has_tracked_changes()

            # Check that there's a deletion and insertion
            element = doc.resolve_ref("p:0")
            del_elem = element.find(f".//{{{WORD_NAMESPACE}}}del")
            ins_elem = element.find(f".//{{{WORD_NAMESPACE}}}ins")
            assert del_elem is not None
            assert ins_elem is not None
        finally:
            docx_path.unlink()

    def test_replace_in_ref_tracked_shows_original_as_deleted(self) -> None:
        """Test that tracked replace shows original text as deleted."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.replace_in_ref("p:0", "First", "Changed", track=True)

            element = doc.resolve_ref("p:0")
            del_text = element.find(f".//{{{WORD_NAMESPACE}}}delText")
            assert del_text is not None
            assert del_text.text == "First"
        finally:
            docx_path.unlink()

    def test_replace_in_ref_tracked_shows_new_as_inserted(self) -> None:
        """Test that tracked replace shows new text as inserted."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.replace_in_ref("p:0", "First", "Changed", track=True)

            element = doc.resolve_ref("p:0")
            ins_elem = element.find(f".//{{{WORD_NAMESPACE}}}ins")
            assert ins_elem is not None
            t_elem = ins_elem.find(f".//{{{WORD_NAMESPACE}}}t")
            assert t_elem is not None
            assert t_elem.text == "Changed"
        finally:
            docx_path.unlink()


# ============================================================================
# Test replace_in_ref() - error handling
# ============================================================================


class TestReplaceInRefErrors:
    """Tests for error handling in replace_in_ref()."""

    def test_replace_in_ref_invalid_ref(self) -> None:
        """Test that invalid ref raises RefNotFoundError."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(RefNotFoundError):
                doc.replace_in_ref("p:999", "text", "replacement")
        finally:
            docx_path.unlink()

    def test_replace_in_ref_text_not_found(self) -> None:
        """Test that missing text raises TextNotFoundError."""
        from python_docx_redline.errors import TextNotFoundError

        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            with pytest.raises(TextNotFoundError, match="Could not find"):
                doc.replace_in_ref("p:0", "nonexistent text", "replacement")
        finally:
            docx_path.unlink()

    def test_replace_in_ref_scoped_to_element(self) -> None:
        """Test that replacement is scoped to the specific element."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            # "paragraph" appears in multiple paragraphs
            # Replace only in p:0 should not affect p:1
            result = doc.replace_in_ref("p:0", "paragraph", "section", track=False)

            assert result.success
            # p:0 should be modified
            text_p0 = doc.get_text_at_ref("p:0")
            assert "section" in text_p0

            # p:1 should NOT be modified
            text_p1 = doc.get_text_at_ref("p:1")
            assert "paragraph" in text_p1
            assert "section" not in text_p1
        finally:
            docx_path.unlink()


# ============================================================================
# Test replace_in_ref() - integration
# ============================================================================


class TestReplaceInRefIntegration:
    """Integration tests for replace_in_ref()."""

    def test_replace_in_ref_save_and_reload(self) -> None:
        """Test that replace_in_ref edits persist after save and reload."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            doc.replace_in_ref("p:0", "First", "Updated", track=True)
            doc.save(docx_path)

            # Reload and verify
            doc2 = Document(docx_path)
            element = doc2.resolve_ref("p:0")
            text = get_text_content(element)
            assert "Updated" in text
            assert doc2.has_tracked_changes()
        finally:
            docx_path.unlink()

    def test_replace_in_ref_with_specific_occurrence(self) -> None:
        """Test replacing a specific occurrence within an element."""
        # Create a document with repeated text
        repeated_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First item, second item, third item.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(repeated_xml)
        try:
            doc = Document(docx_path)

            # Replace only the second occurrence of "item"
            result = doc.replace_in_ref("p:0", "item", "thing", occurrence=2, track=False)

            assert result.success
            text = doc.get_text_at_ref("p:0")
            # Should have: "First item, second thing, third item."
            assert text.count("item") == 2
            assert text.count("thing") == 1
            assert "second thing" in text
        finally:
            docx_path.unlink()
