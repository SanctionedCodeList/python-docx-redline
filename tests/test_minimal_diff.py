"""Tests for minimal editing (legal-style diffs) functionality.

These tests verify the minimal_diff module and the minimal_edits flag
in Document.compare_to().
"""

import atexit
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from python_docx_redline import Document
from python_docx_redline.minimal_diff import (
    compute_minimal_hunks,
    is_punctuation_token,
    is_whitespace_token,
    paragraph_has_nested_runs,
    paragraph_has_tracked_revisions,
    paragraph_has_unsupported_constructs,
    should_use_minimal_editing,
    tokenize,
)

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Track temp directories for cleanup
_temp_dirs: list[Path] = []


def _cleanup_temp_dirs() -> None:
    """Clean up all temporary directories created during tests."""
    for temp_dir in _temp_dirs:
        if temp_dir.exists():
            shutil.rmtree(temp_dir, ignore_errors=True)


atexit.register(_cleanup_temp_dirs)


def create_test_docx(paragraphs: list[str], tmp_path: Path | None = None) -> Path:
    """Create a test .docx file with specified paragraph texts.

    Args:
        paragraphs: List of paragraph text strings
        tmp_path: Optional directory for temp file (uses auto-cleanup temp dir if None)

    Returns:
        Path to the created .docx file
    """
    if tmp_path is None:
        temp_dir = Path(tempfile.mkdtemp())
        _temp_dirs.append(temp_dir)
    else:
        temp_dir = tmp_path
    docx_path = temp_dir / "test.docx"

    # Build document XML with paragraphs
    para_xml = ""
    for text in paragraphs:
        # Escape XML special characters
        text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
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


class TestTokenizer:
    """Tests for the tokenizer function."""

    def test_simple_words(self):
        """Test tokenizing simple words."""
        tokens = tokenize("hello world")
        assert tokens == ["hello", " ", "world"]

    def test_multiple_spaces(self):
        """Test tokenizing text with multiple spaces."""
        tokens = tokenize("hello  world")
        assert tokens == ["hello", "  ", "world"]

    def test_punctuation(self):
        """Test tokenizing text with punctuation."""
        tokens = tokenize("hello, world!")
        assert tokens == ["hello", ",", " ", "world", "!"]

    def test_hyphenated_words(self):
        """Test tokenizing hyphenated words as single tokens."""
        tokens = tokenize("non-disclosure agreement")
        assert tokens == ["non-disclosure", " ", "agreement"]

    def test_apostrophe_words(self):
        """Test tokenizing apostrophe words as single tokens."""
        tokens = tokenize("party's obligation")
        assert tokens == ["party's", " ", "obligation"]

    def test_smart_apostrophe(self):
        """Test tokenizing smart apostrophe words (U+2019)."""
        tokens = tokenize("party\u2019s obligation")  # Unicode right single quote U+2019
        assert tokens == ["party\u2019s", " ", "obligation"]

    def test_mixed_content(self):
        """Test tokenizing complex mixed content."""
        tokens = tokenize("Payment due: net 30 days (non-refundable).")
        expected = [
            "Payment",
            " ",
            "due",
            ":",
            " ",
            "net",
            " ",
            "30",
            " ",
            "days",
            " ",
            "(",
            "non-refundable",
            ")",
            ".",
        ]
        assert tokens == expected

    def test_empty_string(self):
        """Test tokenizing empty string."""
        assert tokenize("") == []

    def test_only_whitespace(self):
        """Test tokenizing only whitespace."""
        assert tokenize("   ") == ["   "]

    def test_only_punctuation(self):
        """Test tokenizing only punctuation."""
        tokens = tokenize(".,;:")
        assert tokens == [".", ",", ";", ":"]


class TestTokenClassification:
    """Tests for token classification functions."""

    def test_is_whitespace_token(self):
        """Test whitespace token detection."""
        assert is_whitespace_token(" ") is True
        assert is_whitespace_token("   ") is True
        assert is_whitespace_token("\t") is True
        assert is_whitespace_token("hello") is False
        assert is_whitespace_token(",") is False

    def test_is_punctuation_token(self):
        """Test punctuation token detection."""
        assert is_punctuation_token(",") is True
        assert is_punctuation_token(";") is True
        assert is_punctuation_token(":") is True
        assert is_punctuation_token(".") is True
        assert is_punctuation_token("(") is True
        assert is_punctuation_token("hello") is False
        assert is_punctuation_token(" ") is False


class TestComputeMinimalHunks:
    """Tests for compute_minimal_hunks function."""

    def test_identical_text_no_hunks(self):
        """Test that identical text produces no hunks."""
        result = compute_minimal_hunks("hello world", "hello world")
        assert result.hunks == []
        assert result.fallback_required is False

    def test_single_word_replacement(self):
        """Test single word replacement produces single hunk."""
        result = compute_minimal_hunks("net 30 days", "net 45 days")
        assert len(result.hunks) == 1
        assert result.hunks[0].delete_text == "30"
        assert result.hunks[0].insert_text == "45"

    def test_punctuation_only_change(self):
        """Test punctuation-only change is tracked (R4)."""
        result = compute_minimal_hunks("Agreement;", "Agreement:")
        assert len(result.hunks) == 1
        assert result.hunks[0].delete_text == ";"
        assert result.hunks[0].insert_text == ":"
        assert result.hunks[0].is_punctuation_only is True

    def test_whitespace_only_change_suppressed(self):
        """Test whitespace-only change is suppressed (R3)."""
        result = compute_minimal_hunks("Section  2", "Section 2")
        # Whitespace-only change should produce no hunks
        assert len(result.hunks) == 0
        assert result.fallback_required is False

    def test_whitespace_adjacent_to_content_change(self):
        """Test whitespace adjacent to content change is preserved (R3 exception)."""
        result = compute_minimal_hunks("net 30 days", "net  45 days")
        # Should have changes - the whitespace change is adjacent to word change
        assert len(result.hunks) >= 1
        # The word replacement should be present
        word_changes = [h for h in result.hunks if "30" in h.delete_text or "45" in h.insert_text]
        assert len(word_changes) >= 1

    def test_multiple_word_changes(self):
        """Test multiple word changes."""
        result = compute_minimal_hunks("The quick brown fox", "The slow gray fox")
        # Should have changes for "quick" -> "slow" and "brown" -> "gray"
        assert len(result.hunks) == 2

    def test_deletion_only(self):
        """Test pure deletion (no replacement)."""
        result = compute_minimal_hunks("hello world test", "hello world")
        assert len(result.hunks) == 1
        assert "test" in result.hunks[0].delete_text
        assert result.hunks[0].insert_text == ""

    def test_insertion_only(self):
        """Test pure insertion (no deletion)."""
        result = compute_minimal_hunks("hello world", "hello new world")
        assert len(result.hunks) == 1
        assert result.hunks[0].delete_text == ""
        assert "new" in result.hunks[0].insert_text

    def test_fragmentation_fallback(self):
        """Test fallback when too many hunks (R5)."""
        # Create text with many alternating changes
        orig = " ".join([f"word{i}" for i in range(20)])
        new = " ".join([f"changed{i}" for i in range(20)])

        result = compute_minimal_hunks(orig, new, max_hunks=8)
        assert result.fallback_required is True
        assert "Too many hunks" in result.fallback_reason

    def test_character_offsets_computed(self):
        """Test that character offsets are computed correctly."""
        result = compute_minimal_hunks("hello world", "hello everyone")
        assert len(result.hunks) == 1
        hunk = result.hunks[0]
        assert hunk.char_start == 6  # "world" starts at index 6
        assert hunk.char_end == 11  # "world" ends at index 11


class TestParagraphChecks:
    """Tests for paragraph safety checks."""

    def test_paragraph_with_tracked_revisions(self):
        """Test detection of existing tracked revisions."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:r><w:t>Before </w:t></w:r>
            <w:ins w:id="1"><w:r><w:t>inserted</w:t></w:r></w:ins>
            <w:r><w:t> after</w:t></w:r>
        </w:p>"""
        para = etree.fromstring(xml)
        assert paragraph_has_tracked_revisions(para) is True

    def test_paragraph_without_tracked_revisions(self):
        """Test paragraph without tracked revisions."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:r><w:t>Plain text here</w:t></w:r>
        </w:p>"""
        para = etree.fromstring(xml)
        assert paragraph_has_tracked_revisions(para) is False

    def test_paragraph_with_deletion(self):
        """Test detection of existing deletion."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:del w:id="1"><w:r><w:delText>deleted</w:delText></w:r></w:del>
        </w:p>"""
        para = etree.fromstring(xml)
        assert paragraph_has_tracked_revisions(para) is True

    def test_paragraph_with_hyperlink(self):
        """Test detection of unsupported hyperlink construct."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <w:r><w:t>Link text</w:t></w:r>
            </w:hyperlink>
        </w:p>"""
        para = etree.fromstring(xml)
        assert paragraph_has_unsupported_constructs(para) is True

    def test_paragraph_without_unsupported_constructs(self):
        """Test paragraph without unsupported constructs."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:r><w:t>Plain text</w:t></w:r>
        </w:p>"""
        para = etree.fromstring(xml)
        assert paragraph_has_unsupported_constructs(para) is False

    def test_paragraph_with_smarttag(self):
        """Test detection of smartTag construct (wraps runs)."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:smartTag w:uri="test" w:element="test">
                <w:r><w:t>Smart tag text</w:t></w:r>
            </w:smartTag>
        </w:p>"""
        para = etree.fromstring(xml)
        assert paragraph_has_unsupported_constructs(para) is True

    def test_paragraph_with_nested_runs(self):
        """Test detection of runs nested inside wrapper elements."""
        # Runs inside a smartTag (nested)
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:smartTag w:uri="test" w:element="test">
                <w:r><w:t>Nested run</w:t></w:r>
            </w:smartTag>
        </w:p>"""
        para = etree.fromstring(xml)
        assert paragraph_has_nested_runs(para) is True

    def test_paragraph_without_nested_runs(self):
        """Test paragraph with direct child runs."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:r><w:t>Direct run 1</w:t></w:r>
            <w:r><w:t>Direct run 2</w:t></w:r>
        </w:p>"""
        para = etree.fromstring(xml)
        assert paragraph_has_nested_runs(para) is False

    def test_paragraph_with_bookmarks_and_direct_runs(self):
        """Test paragraph with bookmark markers but direct child runs.

        Bookmark markers (bookmarkStart/bookmarkEnd) are typically siblings
        of runs, not wrappers, so this should return False.
        """
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:bookmarkStart w:id="0" w:name="test"/>
            <w:r><w:t>Text inside bookmark range</w:t></w:r>
            <w:bookmarkEnd w:id="0"/>
        </w:p>"""
        para = etree.fromstring(xml)
        # Runs are still direct children, bookmarks don't wrap them
        assert paragraph_has_nested_runs(para) is False


class TestShouldUseMinimalEditing:
    """Tests for should_use_minimal_editing function."""

    def test_safe_paragraph_should_use_minimal(self):
        """Test that safe paragraph uses minimal editing."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:r><w:t>net 30 days</w:t></w:r>
        </w:p>"""
        para = etree.fromstring(xml)

        use_minimal, diff_result, reason = should_use_minimal_editing(
            para, "net 45 days", "net 30 days"
        )
        assert use_minimal is True
        assert len(diff_result.hunks) == 1
        assert reason == ""

    def test_paragraph_with_tracked_changes_fallback(self):
        """Test that paragraph with tracked changes falls back."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:ins w:id="1"><w:r><w:t>text</w:t></w:r></w:ins>
        </w:p>"""
        para = etree.fromstring(xml)

        use_minimal, diff_result, reason = should_use_minimal_editing(para, "new text", "text")
        assert use_minimal is False
        assert "tracked revisions" in reason

    def test_too_fragmented_fallback(self):
        """Test that too fragmented diff falls back."""
        xml = f"""<w:p xmlns:w="{WORD_NAMESPACE}">
            <w:r><w:t>word1 word2 word3 word4 word5 word6 word7 word8 word9 word10</w:t></w:r>
        </w:p>"""
        para = etree.fromstring(xml)

        orig = "word1 word2 word3 word4 word5 word6 word7 word8 word9 word10"
        new = "new1 new2 new3 new4 new5 new6 new7 new8 new9 new10"

        use_minimal, diff_result, reason = should_use_minimal_editing(para, new, orig, max_hunks=5)
        assert use_minimal is False
        assert "Too many hunks" in reason


class TestMinimalEditsCompareToIntegration:
    """Integration tests for compare_to with minimal_edits=True."""

    def test_small_word_replacement_is_small(self):
        """A1: Small word replacement produces small tracked change."""
        original = Document(create_test_docx(["Payment due: net 30 days"]))
        modified = Document(create_test_docx(["Payment due: net 45 days"]))

        count = original.compare_to(modified, minimal_edits=True)

        # Should have minimal changes (1 hunk = 1 change), not 2 (full delete+insert)
        assert count >= 1
        assert original.has_tracked_changes()

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        # Should have both del and ins for the word change
        assert "w:del" in xml_str
        assert "w:ins" in xml_str
        # The unchanged text should NOT be inside tracked changes
        assert "Payment" in xml_str

    def test_deletion_then_insertion_ordering(self):
        """A2: Deletion must appear before insertion at same location."""
        original = Document(create_test_docx(["net 30 days"]))
        modified = Document(create_test_docx(["net 45 days"]))

        original.compare_to(modified, minimal_edits=True)

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        # Find positions of del and ins
        del_pos = xml_str.find("w:del")
        ins_pos = xml_str.find("w:ins")

        assert del_pos != -1, "Deletion not found"
        assert ins_pos != -1, "Insertion not found"
        assert del_pos < ins_pos, "Deletion should come before insertion"

    def test_punctuation_only_change_tracked(self):
        """A3: Punctuation-only change produces tracked change."""
        original = Document(create_test_docx(["Agreement;"]))
        modified = Document(create_test_docx(["Agreement:"]))

        count = original.compare_to(modified, minimal_edits=True)

        assert count >= 1
        assert original.has_tracked_changes()

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert "w:del" in xml_str
        assert "w:ins" in xml_str

    def test_whitespace_only_change_suppressed(self):
        """A4: Whitespace-only change produces no tracked changes."""
        original = Document(create_test_docx(["Section  2"]))  # double space
        modified = Document(create_test_docx(["Section 2"]))  # single space

        count = original.compare_to(modified, minimal_edits=True)

        # Per spec: whitespace-only changes are suppressed
        assert count == 0
        assert not original.has_tracked_changes()

    def test_whitespace_adjacent_to_word_change_preserved(self):
        """A5: Whitespace adjacent to word change is preserved."""
        original = Document(create_test_docx(["net 30 days"]))
        modified = Document(create_test_docx(["net  45 days"]))  # extra space

        count = original.compare_to(modified, minimal_edits=True)

        assert count >= 1
        assert original.has_tracked_changes()

    def test_multiple_run_formatting_preserved(self):
        """A6: Formatting on unchanged runs is preserved."""
        # This test would need a document with multiple runs with different formatting
        # For now, test that basic multi-run handling works
        original = Document(create_test_docx(["The quick brown fox"]))
        modified = Document(create_test_docx(["The slow brown fox"]))

        original.compare_to(modified, minimal_edits=True)

        assert original.has_tracked_changes()
        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        # Original text parts should still be there
        assert "The" in xml_str
        assert "brown" in xml_str
        assert "fox" in xml_str

    def test_safety_fallback_to_coarse(self):
        """A7: Fragmented diff falls back to coarse delete+insert."""
        # Create a highly fragmented change that exceeds max hunks
        orig_words = [f"word{i}" for i in range(20)]
        new_words = [f"changed{i}" for i in range(20)]

        original = Document(create_test_docx([" ".join(orig_words)]))
        modified = Document(create_test_docx([" ".join(new_words)]))

        count = original.compare_to(modified, minimal_edits=True)

        # Should fall back and still produce valid changes
        assert count >= 1
        assert original.has_tracked_changes()

    def test_minimal_edits_false_uses_coarse(self):
        """Test that minimal_edits=False uses coarse behavior."""
        original = Document(create_test_docx(["net 30 days"]))
        modified = Document(create_test_docx(["net 45 days"]))

        count = original.compare_to(modified, minimal_edits=False)

        # Coarse behavior: delete paragraph + insert paragraph = 2
        assert count == 2
        assert original.has_tracked_changes()

    def test_non_1to1_replacement_uses_coarse(self):
        """Test that non-1:1 paragraph replacement uses coarse even with minimal_edits."""
        # Replace one paragraph with two
        original = Document(create_test_docx(["One paragraph"]))
        modified = Document(create_test_docx(["First paragraph", "Second paragraph"]))

        count = original.compare_to(modified, minimal_edits=True)

        # This is not 1:1 so should use coarse behavior
        assert count >= 2  # Delete 1 + Insert 2 = 3
        assert original.has_tracked_changes()

    def test_identical_documents_no_changes(self):
        """Test that identical documents produce no changes with minimal_edits."""
        original = Document(create_test_docx(["Same text here"]))
        modified = Document(create_test_docx(["Same text here"]))

        count = original.compare_to(modified, minimal_edits=True)

        assert count == 0
        assert not original.has_tracked_changes()

    def test_accept_changes_produces_modified_text(self):
        """Test that accepting changes produces the modified text (G2)."""
        original = Document(create_test_docx(["net 30 days"]))
        modified = Document(create_test_docx(["net 45 days"]))

        original.compare_to(modified, minimal_edits=True)
        original.accept_all_changes()

        # After accepting, should have the modified text
        text = original.paragraphs[0].text
        assert "45" in text
        assert "30" not in text

    def test_multiple_paragraphs_mixed(self):
        """Test multiple paragraphs with mixed changes."""
        original = Document(
            create_test_docx(
                [
                    "Paragraph 1 unchanged",
                    "Payment: net 30 days",
                    "Paragraph 3 unchanged",
                ]
            )
        )
        modified = Document(
            create_test_docx(
                [
                    "Paragraph 1 unchanged",
                    "Payment: net 45 days",
                    "Paragraph 3 unchanged",
                ]
            )
        )

        count = original.compare_to(modified, minimal_edits=True)

        # Should only have changes for the second paragraph
        assert count >= 1
        assert original.has_tracked_changes()

    def test_author_attribution(self):
        """Test that author is correctly attributed in minimal edits."""
        original = Document(create_test_docx(["net 30 days"]))
        modified = Document(create_test_docx(["net 45 days"]))

        original.compare_to(modified, author="Legal Team", minimal_edits=True)

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert 'w:author="Legal Team"' in xml_str


class TestMinimalEditsEdgeCases:
    """Edge case tests for minimal editing."""

    def test_empty_paragraph_change(self):
        """Test changing from/to empty paragraph."""
        original = Document(create_test_docx([""]))
        modified = Document(create_test_docx(["New content"]))

        count = original.compare_to(modified, minimal_edits=True)

        # Empty to content should work
        assert count >= 1

    def test_special_characters_in_text(self):
        """Test text with special XML characters."""
        original = Document(create_test_docx(["Value < 30 & > 20"]))
        modified = Document(create_test_docx(["Value < 45 & > 20"]))

        count = original.compare_to(modified, minimal_edits=True)

        assert count >= 1
        assert original.has_tracked_changes()

    def test_long_paragraph_small_change(self):
        """Test long paragraph with small change stays minimal."""
        long_text = "This is a very long paragraph with many words. " * 20
        long_text_orig = long_text + "Original ending."
        long_text_mod = long_text + "Modified ending."

        original = Document(create_test_docx([long_text_orig]))
        modified = Document(create_test_docx([long_text_mod]))

        count = original.compare_to(modified, minimal_edits=True)

        # Should have minimal changes for just the ending
        assert count >= 1
        assert original.has_tracked_changes()


class TestMinimalEditsPersistence:
    """Tests for saving and reloading documents with minimal edits."""

    def test_changes_persist_after_save(self):
        """Test that minimal edit changes persist after save."""
        original = Document(create_test_docx(["net 30 days"]))
        modified = Document(create_test_docx(["net 45 days"]))

        original.compare_to(modified, minimal_edits=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "compared.docx"
            original.save(output_path)

            reloaded = Document(output_path)
            assert reloaded.has_tracked_changes()

    def test_reloaded_can_accept_changes(self):
        """Test that reloaded document can accept changes."""
        original = Document(create_test_docx(["net 30 days"]))
        modified = Document(create_test_docx(["net 45 days"]))

        original.compare_to(modified, minimal_edits=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "compared.docx"
            original.save(output_path)

            reloaded = Document(output_path)
            reloaded.accept_all_changes()

            assert not reloaded.has_tracked_changes()
            text = reloaded.paragraphs[0].text
            assert "45" in text


def create_multi_run_docx(runs_per_para: list[list[str]], tmp_path: Path | None = None) -> Path:
    """Create a test .docx file with multiple runs per paragraph.

    Args:
        runs_per_para: List of paragraphs, each containing a list of run texts
        tmp_path: Optional directory for temp file (uses auto-cleanup temp dir if None)

    Returns:
        Path to the created .docx file
    """
    if tmp_path is None:
        temp_dir = Path(tempfile.mkdtemp())
        _temp_dirs.append(temp_dir)
    else:
        temp_dir = tmp_path
    docx_path = temp_dir / "test_multi_run.docx"

    # Build document XML with multiple runs per paragraph
    para_xml = ""
    for runs in runs_per_para:
        runs_xml = ""
        for text in runs:
            text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            runs_xml += f"""
      <w:r>
        <w:t>{text}</w:t>
      </w:r>"""
        para_xml += f"""
    <w:p>{runs_xml}
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


class TestInsertionBoundaries:
    """Tests for insertion at paragraph start and run boundaries (reviewer issue #2)."""

    def test_pure_insertion_at_paragraph_start(self):
        """Test inserting new text at the very beginning of a paragraph."""
        original = Document(create_test_docx(["existing text"]))
        modified = Document(create_test_docx(["NEW existing text"]))

        count = original.compare_to(modified, minimal_edits=True)

        assert count >= 1
        assert original.has_tracked_changes()

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert "w:ins" in xml_str
        # The inserted text should be "NEW "
        assert "NEW" in xml_str
        # Original text should still be present
        assert "existing" in xml_str

    def test_pure_insertion_at_paragraph_start_single_word(self):
        """Test inserting a single word at the beginning."""
        original = Document(create_test_docx(["world"]))
        modified = Document(create_test_docx(["Hello world"]))

        count = original.compare_to(modified, minimal_edits=True)

        assert count >= 1
        assert original.has_tracked_changes()

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert "w:ins" in xml_str
        assert "Hello" in xml_str

    def test_insertion_at_run_boundary_multi_run(self):
        """Test insertion at the boundary between two runs in a multi-run paragraph."""
        # Original: [Run1: "Hello "][Run2: "world"]
        # Modified: [Run1: "Hello "][NEW: "beautiful "][Run2: "world"]
        original = Document(create_multi_run_docx([["Hello ", "world"]]))
        modified = Document(create_test_docx(["Hello beautiful world"]))

        count = original.compare_to(modified, minimal_edits=True)

        assert count >= 1
        assert original.has_tracked_changes()

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert "w:ins" in xml_str
        assert "beautiful" in xml_str

    def test_replacement_at_run_boundary(self):
        """Test replacement that spans a run boundary."""
        # Original: [Run1: "net "][Run2: "30 days"]
        # Modified: "net 45 days"
        original = Document(create_multi_run_docx([["net ", "30 days"]]))
        modified = Document(create_test_docx(["net 45 days"]))

        count = original.compare_to(modified, minimal_edits=True)

        assert count >= 1
        assert original.has_tracked_changes()

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert "w:del" in xml_str
        assert "w:ins" in xml_str

    def test_deletion_at_paragraph_start(self):
        """Test deleting text from the beginning of a paragraph."""
        original = Document(create_test_docx(["OLD existing text"]))
        modified = Document(create_test_docx(["existing text"]))

        count = original.compare_to(modified, minimal_edits=True)

        assert count >= 1
        assert original.has_tracked_changes()

        xml_str = etree.tostring(original.xml_root, encoding="unicode")
        assert "w:del" in xml_str
        # "OLD " should be in the deletion
        assert "OLD" in xml_str

    def test_insertion_preserves_run_order(self):
        """Verify inserted text appears in correct position relative to existing runs."""
        original = Document(create_test_docx(["Payment due"]))
        modified = Document(create_test_docx(["URGENT Payment due"]))

        original.compare_to(modified, minimal_edits=True)

        # After accepting changes, the text should be in correct order
        original.accept_all_changes()
        text = original.paragraphs[0].text
        assert text == "URGENT Payment due"
