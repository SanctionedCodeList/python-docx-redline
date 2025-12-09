"""Tests for document comparison functionality."""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_redline.document import Document

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def create_test_docx(paragraphs: list[str]) -> Path:
    """Create a test .docx file with specified paragraph texts.

    Args:
        paragraphs: List of paragraph text strings

    Returns:
        Path to the created .docx file
    """
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    # Build document XML with paragraphs
    para_xml = ""
    for text in paragraphs:
        para_xml += f"""
    <w:p>
      <w:r>
        <w:t>{text}</w:t>
      </w:r>
    </w:p>"""

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{para_xml}
  </w:body>
</w:document>"""

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
        docx.writestr("word/document.xml", document_xml)

    return docx_path


class TestCompareToBasic:
    """Tests for basic compare_to() functionality."""

    def test_identical_documents_no_changes(self):
        """Test comparing identical documents produces no changes."""
        original = Document(create_test_docx(["Line 1", "Line 2", "Line 3"]))
        modified = Document(create_test_docx(["Line 1", "Line 2", "Line 3"]))

        count = original.compare_to(modified)

        assert count == 0
        assert not original.has_tracked_changes()

    def test_simple_insertion(self):
        """Test detecting a simple paragraph insertion."""
        original = Document(create_test_docx(["Line 1", "Line 3"]))
        modified = Document(create_test_docx(["Line 1", "Line 2", "Line 3"]))

        count = original.compare_to(modified)

        assert count == 1
        assert original.has_tracked_changes()

        # Verify the insertion is tracked
        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert "w:ins" in xml_str
        assert "Line 2" in xml_str

    def test_simple_deletion(self):
        """Test detecting a simple paragraph deletion."""
        original = Document(create_test_docx(["Line 1", "Line 2", "Line 3"]))
        modified = Document(create_test_docx(["Line 1", "Line 3"]))

        count = original.compare_to(modified)

        assert count == 1
        assert original.has_tracked_changes()

        # Verify the deletion is tracked
        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert "w:del" in xml_str
        assert "w:delText" in xml_str

    def test_simple_replacement(self):
        """Test detecting a paragraph replacement (text changed)."""
        original = Document(create_test_docx(["Line 1", "Original text", "Line 3"]))
        modified = Document(create_test_docx(["Line 1", "Modified text", "Line 3"]))

        count = original.compare_to(modified)

        # Replace = delete old + insert new = 2 changes
        assert count == 2
        assert original.has_tracked_changes()

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert "w:del" in xml_str
        assert "w:ins" in xml_str

    def test_multiple_insertions(self):
        """Test detecting multiple paragraph insertions."""
        original = Document(create_test_docx(["Start", "End"]))
        modified = Document(create_test_docx(["Start", "New 1", "New 2", "End"]))

        count = original.compare_to(modified)

        assert count == 2
        assert original.has_tracked_changes()

    def test_multiple_deletions(self):
        """Test detecting multiple paragraph deletions."""
        original = Document(create_test_docx(["Start", "Delete 1", "Delete 2", "End"]))
        modified = Document(create_test_docx(["Start", "End"]))

        count = original.compare_to(modified)

        assert count == 2
        assert original.has_tracked_changes()


class TestCompareToEdgeCases:
    """Tests for edge cases in compare_to()."""

    def test_insert_at_beginning(self):
        """Test inserting paragraph at the very beginning."""
        original = Document(create_test_docx(["Line 1", "Line 2"]))
        modified = Document(create_test_docx(["New first line", "Line 1", "Line 2"]))

        count = original.compare_to(modified)

        assert count == 1
        assert original.has_tracked_changes()

    def test_insert_at_end(self):
        """Test inserting paragraph at the very end."""
        original = Document(create_test_docx(["Line 1", "Line 2"]))
        modified = Document(create_test_docx(["Line 1", "Line 2", "New last line"]))

        count = original.compare_to(modified)

        assert count == 1
        assert original.has_tracked_changes()

    def test_delete_first_paragraph(self):
        """Test deleting the first paragraph."""
        original = Document(create_test_docx(["First line", "Line 2", "Line 3"]))
        modified = Document(create_test_docx(["Line 2", "Line 3"]))

        count = original.compare_to(modified)

        assert count == 1
        assert original.has_tracked_changes()

    def test_delete_last_paragraph(self):
        """Test deleting the last paragraph."""
        original = Document(create_test_docx(["Line 1", "Line 2", "Last line"]))
        modified = Document(create_test_docx(["Line 1", "Line 2"]))

        count = original.compare_to(modified)

        assert count == 1
        assert original.has_tracked_changes()

    def test_empty_original(self):
        """Test comparing empty document to document with content."""
        original = Document(create_test_docx([]))
        modified = Document(create_test_docx(["New content"]))

        count = original.compare_to(modified)

        assert count == 1
        assert original.has_tracked_changes()

    def test_empty_modified(self):
        """Test comparing document with content to empty document."""
        original = Document(create_test_docx(["Content to delete"]))
        modified = Document(create_test_docx([]))

        count = original.compare_to(modified)

        assert count == 1
        assert original.has_tracked_changes()

    def test_both_empty(self):
        """Test comparing two empty documents."""
        original = Document(create_test_docx([]))
        modified = Document(create_test_docx([]))

        count = original.compare_to(modified)

        assert count == 0
        assert not original.has_tracked_changes()


class TestCompareToAuthor:
    """Tests for author handling in compare_to()."""

    def test_default_author(self):
        """Test that default author is used when not specified."""
        original = Document(create_test_docx(["Line 1"]), author="Default Author")
        modified = Document(create_test_docx(["Line 1", "New line"]))

        original.compare_to(modified)

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert 'w:author="Default Author"' in xml_str

    def test_custom_author(self):
        """Test that custom author is used when specified."""
        original = Document(create_test_docx(["Line 1"]))
        modified = Document(create_test_docx(["Line 1", "New line"]))

        original.compare_to(modified, author="Custom Author")

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert 'w:author="Custom Author"' in xml_str


class TestCompareToPersistence:
    """Tests for saving and reloading compared documents."""

    def test_changes_persist_after_save(self):
        """Test that comparison changes persist after save and reload."""
        original = Document(create_test_docx(["Line 1", "Line 3"]))
        modified = Document(create_test_docx(["Line 1", "Line 2", "Line 3"]))

        original.compare_to(modified)

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "compared.docx"
            original.save(output_path)

            reloaded = Document(output_path)
            assert reloaded.has_tracked_changes()

    def test_can_accept_comparison_changes(self):
        """Test that comparison changes can be accepted."""
        original = Document(create_test_docx(["Line 1", "Line 3"]))
        modified = Document(create_test_docx(["Line 1", "Line 2", "Line 3"]))

        original.compare_to(modified)
        original.accept_all_changes()

        assert not original.has_tracked_changes()


class TestCompareToComplex:
    """Tests for complex comparison scenarios."""

    def test_multiple_operations_same_region(self):
        """Test multiple changes in the same region."""
        original = Document(
            create_test_docx(["Start", "Old middle", "End"])
        )
        modified = Document(
            create_test_docx(["Start", "New middle 1", "New middle 2", "End"])
        )

        count = original.compare_to(modified)

        # Delete "Old middle" + Insert "New middle 1" + Insert "New middle 2" = 3
        assert count == 3
        assert original.has_tracked_changes()

    def test_completely_different_documents(self):
        """Test comparing completely different documents."""
        original = Document(create_test_docx(["A", "B", "C"]))
        modified = Document(create_test_docx(["X", "Y", "Z"]))

        count = original.compare_to(modified)

        # All paragraphs replaced = 3 deletes + 3 inserts = 6
        assert count == 6
        assert original.has_tracked_changes()

    def test_reorder_paragraphs(self):
        """Test detecting reordered paragraphs."""
        original = Document(create_test_docx(["First", "Second", "Third"]))
        modified = Document(create_test_docx(["Third", "Second", "First"]))

        count = original.compare_to(modified)

        # Reordering is detected as a complex change
        assert count > 0
        assert original.has_tracked_changes()

    def test_duplicate_paragraphs(self):
        """Test handling duplicate paragraph text."""
        original = Document(create_test_docx(["Same", "Same", "Different"]))
        modified = Document(create_test_docx(["Same", "Changed", "Different"]))

        count = original.compare_to(modified)

        # One paragraph changed
        assert count == 2  # Delete + Insert
        assert original.has_tracked_changes()

    def test_whitespace_only_changes(self):
        """Test that whitespace-only changes are detected."""
        original = Document(create_test_docx(["Text without spaces"]))
        modified = Document(create_test_docx(["Text with  extra  spaces"]))

        count = original.compare_to(modified)

        # Whitespace change is detected as a replacement
        assert count == 2  # Delete + Insert
        assert original.has_tracked_changes()


class TestCompareToSpecialContent:
    """Tests for special content handling."""

    def test_empty_paragraphs(self):
        """Test comparing documents with empty paragraphs."""
        original = Document(create_test_docx(["Line 1", "", "Line 3"]))
        modified = Document(create_test_docx(["Line 1", "New content", "Line 3"]))

        count = original.compare_to(modified)

        # Empty paragraph replaced with content
        assert count == 2  # Delete empty + Insert new
        assert original.has_tracked_changes()

    def test_special_characters(self):
        """Test comparing documents with special characters."""
        original = Document(
            create_test_docx(["Line with &amp; ampersand", "Line with &lt;brackets&gt;"])
        )
        modified = Document(
            create_test_docx(
                ["Line with &amp; ampersand", "Modified &lt;brackets&gt;"]
            )
        )

        count = original.compare_to(modified)

        assert count == 2  # Delete + Insert
        assert original.has_tracked_changes()

    def test_long_paragraphs(self):
        """Test comparing documents with long paragraphs."""
        long_text = "This is a very long paragraph. " * 100
        original = Document(create_test_docx(["Short", long_text]))
        modified = Document(create_test_docx(["Short", long_text + " Extra text."]))

        count = original.compare_to(modified)

        assert count == 2  # Delete + Insert for modified long paragraph
        assert original.has_tracked_changes()
