"""Tests for document comparison functionality."""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from python_docx_redline.document import Document

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
        original = Document(create_test_docx(["Start", "Old middle", "End"]))
        modified = Document(create_test_docx(["Start", "New middle 1", "New middle 2", "End"]))

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
            create_test_docx(["Line with &amp; ampersand", "Modified &lt;brackets&gt;"])
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


class TestCompareDocumentsFunction:
    """Tests for the standalone compare_documents() function."""

    def test_basic_comparison(self):
        """Test basic function usage with file paths."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1", "Line 2"])
        modified_path = create_test_docx(["Line 1", "New Line", "Line 2"])

        redline = compare_documents(original_path, modified_path)

        assert redline.has_tracked_changes()
        xml_str = etree.tostring(redline.xml_root, encoding="unicode")
        assert "w:ins" in xml_str
        assert "New Line" in xml_str

    def test_returns_document(self):
        """Test that function returns a Document object."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Test"])
        modified_path = create_test_docx(["Test"])

        result = compare_documents(original_path, modified_path)

        assert isinstance(result, Document)

    def test_identical_documents_no_changes(self):
        """Test comparing identical documents produces no changes."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1", "Line 2", "Line 3"])
        modified_path = create_test_docx(["Line 1", "Line 2", "Line 3"])

        redline = compare_documents(original_path, modified_path)

        assert not redline.has_tracked_changes()

    def test_author_parameter(self):
        """Test custom author is used."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1"])
        modified_path = create_test_docx(["Line 1", "New"])

        redline = compare_documents(original_path, modified_path, author="Custom Author")

        xml_str = etree.tostring(redline.xml_root, encoding="unicode")
        assert 'w:author="Custom Author"' in xml_str

    def test_default_author(self):
        """Test default author is 'Comparison'."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1"])
        modified_path = create_test_docx(["Line 1", "New"])

        redline = compare_documents(original_path, modified_path)

        xml_str = etree.tostring(redline.xml_root, encoding="unicode")
        assert 'w:author="Comparison"' in xml_str

    def test_minimal_edits_parameter(self):
        """Test minimal_edits works through the function."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["The quick brown fox"])
        modified_path = create_test_docx(["The slow brown fox"])

        redline = compare_documents(original_path, modified_path, minimal_edits=True)

        assert redline.has_tracked_changes()
        # Minimal edit should show word-level changes
        xml_str = etree.tostring(redline.xml_root, encoding="unicode")
        assert "quick" in xml_str
        assert "slow" in xml_str

    def test_with_bytes_input(self):
        """Test function works with bytes input."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Test"])
        modified_path = create_test_docx(["Test modified"])

        with open(original_path, "rb") as f:
            original_bytes = f.read()
        with open(modified_path, "rb") as f:
            modified_bytes = f.read()

        redline = compare_documents(original_bytes, modified_bytes)

        assert redline.has_tracked_changes()

    def test_deletion(self):
        """Test detecting paragraph deletion."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1", "Line 2", "Line 3"])
        modified_path = create_test_docx(["Line 1", "Line 3"])

        redline = compare_documents(original_path, modified_path)

        assert redline.has_tracked_changes()
        xml_str = etree.tostring(redline.xml_root, encoding="unicode")
        assert "w:del" in xml_str
        assert "w:delText" in xml_str

    def test_replacement(self):
        """Test detecting paragraph replacement."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1", "Original", "Line 3"])
        modified_path = create_test_docx(["Line 1", "Modified", "Line 3"])

        redline = compare_documents(original_path, modified_path)

        assert redline.has_tracked_changes()
        xml_str = etree.tostring(redline.xml_root, encoding="unicode")
        assert "w:del" in xml_str
        assert "w:ins" in xml_str

    def test_saveable(self):
        """Test that returned document can be saved."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1"])
        modified_path = create_test_docx(["Line 1", "New"])

        redline = compare_documents(original_path, modified_path)

        # Save to a temporary location
        temp_dir = Path(tempfile.mkdtemp())
        output_path = temp_dir / "output.docx"

        # Use validate=False since the test document is minimal
        redline.save(output_path, validate=False)

        # Verify it saved correctly
        assert output_path.exists()

        # Verify we can reload it
        reloaded = Document(output_path)
        assert reloaded.has_tracked_changes()


class TestComparisonStats:
    """Tests for the comparison_stats property and ComparisonStats class."""

    def test_stats_after_comparison(self):
        """Test that comparison_stats returns correct counts after comparison."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1", "Line 2", "Line 3"])
        modified_path = create_test_docx(["Line 1", "New Line", "Line 3"])

        redline = compare_documents(original_path, modified_path)
        stats = redline.comparison_stats

        # Replacing "Line 2" with "New Line" = 1 deletion + 1 insertion
        assert stats.deletions == 1
        assert stats.insertions == 1
        assert stats.total == 2
        assert stats.moves == 0
        assert stats.format_changes == 0

    def test_stats_insertion_only(self):
        """Test stats with only insertions."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1", "Line 3"])
        modified_path = create_test_docx(["Line 1", "Line 2", "Line 3"])

        redline = compare_documents(original_path, modified_path)
        stats = redline.comparison_stats

        assert stats.insertions == 1
        assert stats.deletions == 0
        assert stats.total == 1

    def test_stats_deletion_only(self):
        """Test stats with only deletions."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1", "Line 2", "Line 3"])
        modified_path = create_test_docx(["Line 1", "Line 3"])

        redline = compare_documents(original_path, modified_path)
        stats = redline.comparison_stats

        assert stats.insertions == 0
        assert stats.deletions == 1
        assert stats.total == 1

    def test_stats_no_changes(self):
        """Test stats when documents are identical."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1", "Line 2"])
        modified_path = create_test_docx(["Line 1", "Line 2"])

        redline = compare_documents(original_path, modified_path)
        stats = redline.comparison_stats

        assert stats.insertions == 0
        assert stats.deletions == 0
        assert stats.total == 0

    def test_stats_multiple_changes(self):
        """Test stats with multiple insertions and deletions."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["A", "B", "C", "D"])
        modified_path = create_test_docx(["A", "X", "Y", "D"])

        redline = compare_documents(original_path, modified_path)
        stats = redline.comparison_stats

        # B and C deleted, X and Y inserted
        assert stats.deletions == 2
        assert stats.insertions == 2
        assert stats.total == 4

    def test_stats_str_representation(self):
        """Test the string representation of ComparisonStats."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1", "Line 2"])
        modified_path = create_test_docx(["Line 1", "New", "Line 2"])

        redline = compare_documents(original_path, modified_path)
        stats = redline.comparison_stats

        # Should have readable string output
        stats_str = str(stats)
        assert "insertion" in stats_str

    def test_stats_str_no_changes(self):
        """Test string representation with no changes."""
        from python_docx_redline import compare_documents

        original_path = create_test_docx(["Line 1"])
        modified_path = create_test_docx(["Line 1"])

        redline = compare_documents(original_path, modified_path)
        stats = redline.comparison_stats

        assert str(stats) == "No changes"

    def test_stats_singular_plural(self):
        """Test correct singular/plural in string output."""
        from python_docx_redline.results import ComparisonStats

        # Single insertion
        stats1 = ComparisonStats(insertions=1, deletions=0)
        assert "1 insertion" in str(stats1)
        assert "insertions" not in str(stats1)

        # Multiple insertions
        stats2 = ComparisonStats(insertions=3, deletions=0)
        assert "3 insertions" in str(stats2)

    def test_stats_with_compare_to_method(self):
        """Test that comparison_stats works with compare_to() method too."""
        original_path = create_test_docx(["Line 1", "Old text"])
        modified_path = create_test_docx(["Line 1", "New text"])

        original = Document(original_path)
        modified = Document(modified_path)

        original.compare_to(modified)
        stats = original.comparison_stats

        # 1 deletion + 1 insertion for replacement
        assert stats.total == 2
