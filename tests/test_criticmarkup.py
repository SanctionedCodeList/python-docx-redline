"""Tests for CriticMarkup parser."""

import tempfile
import zipfile
from pathlib import Path

import pytest

from python_docx_redline.criticmarkup import (
    CriticOperation,
    OperationType,
    parse_criticmarkup,
    render_criticmarkup,
    strip_criticmarkup,
)


class TestParseInsertion:
    """Tests for parsing insertion operations."""

    def test_simple_insertion(self):
        """Parse a simple insertion."""
        ops = parse_criticmarkup("Hello {++world++}!")
        assert len(ops) == 1
        assert ops[0].type == OperationType.INSERTION
        assert ops[0].text == "world"
        assert ops[0].position == 6
        assert ops[0].end_position == 17  # Position after closing }

    def test_insertion_with_context(self):
        """Insertion should capture surrounding context."""
        ops = parse_criticmarkup("The quick brown {++lazy++} fox")
        assert len(ops) == 1
        assert ops[0].context_before == "The quick brown "
        assert ops[0].context_after == " fox"

    def test_multiple_insertions(self):
        """Parse multiple insertions in order."""
        ops = parse_criticmarkup("{++First++} and {++second++}")
        assert len(ops) == 2
        assert ops[0].text == "First"
        assert ops[1].text == "second"
        assert ops[0].position < ops[1].position

    def test_multiline_insertion(self):
        """Parse insertion spanning multiple lines."""
        text = "{++Line one\nLine two++}"
        ops = parse_criticmarkup(text)
        assert len(ops) == 1
        assert ops[0].text == "Line one\nLine two"


class TestParseDeletion:
    """Tests for parsing deletion operations."""

    def test_simple_deletion(self):
        """Parse a simple deletion."""
        ops = parse_criticmarkup("Hello {--world--}!")
        assert len(ops) == 1
        assert ops[0].type == OperationType.DELETION
        assert ops[0].text == "world"

    def test_deletion_with_context(self):
        """Deletion should capture surrounding context."""
        ops = parse_criticmarkup("Remove {--this text--} please")
        assert len(ops) == 1
        assert ops[0].context_before == "Remove "
        assert ops[0].context_after == " please"

    def test_multiple_deletions(self):
        """Parse multiple deletions in order."""
        ops = parse_criticmarkup("{--First--} and {--second--}")
        assert len(ops) == 2
        assert ops[0].text == "First"
        assert ops[1].text == "second"


class TestParseSubstitution:
    """Tests for parsing substitution operations."""

    def test_simple_substitution(self):
        """Parse a simple substitution."""
        ops = parse_criticmarkup("Payment in {~~30~>45~~} days")
        assert len(ops) == 1
        assert ops[0].type == OperationType.SUBSTITUTION
        assert ops[0].text == "30"
        assert ops[0].replacement == "45"

    def test_substitution_with_context(self):
        """Substitution should capture surrounding context."""
        ops = parse_criticmarkup("The {~~old~>new~~} value")
        assert len(ops) == 1
        assert ops[0].context_before == "The "
        assert ops[0].context_after == " value"

    def test_substitution_with_spaces(self):
        """Parse substitution with spaces in content."""
        ops = parse_criticmarkup("{~~old text~>new text~~}")
        assert len(ops) == 1
        assert ops[0].text == "old text"
        assert ops[0].replacement == "new text"


class TestParseComment:
    """Tests for parsing comment operations."""

    def test_simple_comment(self):
        """Parse a simple comment."""
        ops = parse_criticmarkup("Some text {>>review this<<}")
        assert len(ops) == 1
        assert ops[0].type == OperationType.COMMENT
        assert ops[0].comment == "review this"
        assert ops[0].text == ""  # Comments have no main text

    def test_comment_with_context(self):
        """Comment should capture surrounding context."""
        ops = parse_criticmarkup("Check this {>>important<<} part")
        assert len(ops) == 1
        assert ops[0].context_before == "Check this "
        assert ops[0].context_after == " part"


class TestParseHighlight:
    """Tests for parsing highlight operations."""

    def test_simple_highlight(self):
        """Parse a simple highlight."""
        ops = parse_criticmarkup("Please {==review this section==}")
        assert len(ops) == 1
        assert ops[0].type == OperationType.HIGHLIGHT
        assert ops[0].text == "review this section"
        assert ops[0].comment is None

    def test_highlight_with_comment(self):
        """Parse highlight with nested comment."""
        ops = parse_criticmarkup("Check {==this text=={>>needs review<<}}")
        assert len(ops) == 1
        assert ops[0].type == OperationType.HIGHLIGHT
        assert ops[0].text == "this text"
        assert ops[0].comment == "needs review"

    def test_highlight_comment_not_matched_twice(self):
        """Nested comment in highlight should not also match as standalone."""
        ops = parse_criticmarkup("{==marked=={>>comment<<}}")
        assert len(ops) == 1  # Should be ONE operation, not two
        assert ops[0].type == OperationType.HIGHLIGHT
        assert ops[0].text == "marked"
        assert ops[0].comment == "comment"


class TestParseMixed:
    """Tests for parsing mixed operation types."""

    def test_insertion_and_deletion(self):
        """Parse both insertion and deletion."""
        ops = parse_criticmarkup("The {--old--}{++new++} contract")
        assert len(ops) == 2
        assert ops[0].type == OperationType.DELETION
        assert ops[0].text == "old"
        assert ops[1].type == OperationType.INSERTION
        assert ops[1].text == "new"

    def test_all_operation_types(self):
        """Parse all five operation types in one text."""
        text = "{++inserted++} {--deleted--} {~~old~>new~~} {>>comment<<} {==highlighted==}"
        ops = parse_criticmarkup(text)
        assert len(ops) == 5

        types = {op.type for op in ops}
        assert types == {
            OperationType.INSERTION,
            OperationType.DELETION,
            OperationType.SUBSTITUTION,
            OperationType.COMMENT,
            OperationType.HIGHLIGHT,
        }

    def test_operations_sorted_by_position(self):
        """Operations should be returned sorted by position."""
        text = "End {++last++} start {++first++} middle"
        ops = parse_criticmarkup(text)
        assert ops[0].text == "last"  # Position 4
        assert ops[1].text == "first"  # Position 21


class TestParseContext:
    """Tests for context extraction."""

    def test_context_chars_parameter(self):
        """Test custom context character count."""
        text = "A" * 100 + "{++X++}" + "B" * 100
        ops = parse_criticmarkup(text, context_chars=10)
        assert len(ops[0].context_before) == 10
        assert len(ops[0].context_after) == 10

    def test_context_at_start(self):
        """Context before should handle start of text."""
        ops = parse_criticmarkup("{++start++} text")
        assert ops[0].context_before == ""

    def test_context_at_end(self):
        """Context after should handle end of text."""
        ops = parse_criticmarkup("text {++end++}")
        assert ops[0].context_after == ""

    def test_context_strips_criticmarkup(self):
        """Context should strip CriticMarkup from surrounding operations."""
        # If there's a deletion before our insertion, context should be clean
        text = "prefix {--removed--}{++added++} suffix"
        ops = parse_criticmarkup(text)

        # Find the insertion
        insertion = next(op for op in ops if op.type == OperationType.INSERTION)
        # Context before should not include the {--removed--} markup
        assert "{--" not in insertion.context_before


class TestParseEdgeCases:
    """Tests for edge cases and boundary conditions."""

    def test_empty_string(self):
        """Parse empty string."""
        ops = parse_criticmarkup("")
        assert len(ops) == 0

    def test_no_criticmarkup(self):
        """Parse text with no CriticMarkup."""
        ops = parse_criticmarkup("Just plain text here.")
        assert len(ops) == 0

    def test_escaped_braces(self):
        """Regular braces that look like CriticMarkup but aren't."""
        # These should NOT be parsed as operations
        ops = parse_criticmarkup("Use {curly braces} normally")
        assert len(ops) == 0

    def test_incomplete_markup(self):
        """Incomplete markup should not match."""
        ops = parse_criticmarkup("{++incomplete")
        assert len(ops) == 0

        ops = parse_criticmarkup("incomplete++}")
        assert len(ops) == 0

    def test_empty_insertion_not_matched(self):
        """Empty insertion content is not matched (requires at least one char)."""
        # Empty insertions like {++++} don't make semantic sense
        # and are not matched by the parser
        ops = parse_criticmarkup("{++++}")
        assert len(ops) == 0

    def test_special_characters_in_content(self):
        """CriticMarkup with special characters in content."""
        ops = parse_criticmarkup("{++text with $pecial ch@rs!++}")
        assert len(ops) == 1
        assert ops[0].text == "text with $pecial ch@rs!"

    def test_unicode_content(self):
        """CriticMarkup with unicode content."""
        ops = parse_criticmarkup("{++日本語テキスト++}")
        assert len(ops) == 1
        assert ops[0].text == "日本語テキスト"


class TestStripCriticmarkup:
    """Tests for strip_criticmarkup function."""

    def test_strip_insertion(self):
        """Strip insertion keeps inserted text."""
        assert strip_criticmarkup("Hello {++world++}!") == "Hello world!"

    def test_strip_deletion(self):
        """Strip deletion removes deleted text."""
        assert strip_criticmarkup("Say {--goodbye--}hello") == "Say hello"

    def test_strip_substitution(self):
        """Strip substitution keeps new text."""
        assert strip_criticmarkup("{~~old~>new~~}") == "new"

    def test_strip_comment(self):
        """Strip comment removes comment entirely."""
        assert strip_criticmarkup("Text{>>comment<<} here") == "Text here"

    def test_strip_highlight(self):
        """Strip highlight keeps highlighted text."""
        assert strip_criticmarkup("{==important==}") == "important"

    def test_strip_highlight_with_comment(self):
        """Strip highlight+comment keeps highlighted text only."""
        assert strip_criticmarkup("{==text=={>>note<<}}") == "text"

    def test_strip_mixed(self):
        """Strip multiple operations in one text."""
        text = "The {--old--}{++new++} {~~30~>45~~} day term"
        expected = "The new 45 day term"
        assert strip_criticmarkup(text) == expected

    def test_strip_no_markup(self):
        """Strip with no markup returns original text."""
        text = "Plain text without markup"
        assert strip_criticmarkup(text) == text


class TestRenderCriticmarkup:
    """Tests for render_criticmarkup function."""

    def test_render_insertion(self):
        """Render insertion operation."""
        ops = [
            CriticOperation(
                type=OperationType.INSERTION,
                text="world",
                position=6,
            )
        ]
        result = render_criticmarkup(ops, "Hello !")
        assert result == "Hello {++world++}!"

    def test_render_deletion(self):
        """Render deletion operation."""
        ops = [
            CriticOperation(
                type=OperationType.DELETION,
                text="world",
                position=6,
            )
        ]
        result = render_criticmarkup(ops, "Hello world!")
        assert result == "Hello {--world--}!"

    def test_render_substitution(self):
        """Render substitution operation."""
        ops = [
            CriticOperation(
                type=OperationType.SUBSTITUTION,
                text="old",
                replacement="new",
                position=0,
            )
        ]
        result = render_criticmarkup(ops, "old text")
        assert result == "{~~old~>new~~} text"

    def test_render_comment(self):
        """Render comment operation."""
        ops = [
            CriticOperation(
                type=OperationType.COMMENT,
                text="",
                comment="review",
                position=5,
            )
        ]
        result = render_criticmarkup(ops, "Hello world")
        assert result == "Hello{>>review<<} world"

    def test_render_highlight(self):
        """Render highlight operation."""
        ops = [
            CriticOperation(
                type=OperationType.HIGHLIGHT,
                text="important",
                position=0,
            )
        ]
        result = render_criticmarkup(ops, "important text")
        assert result == "{==important==} text"

    def test_render_highlight_with_comment(self):
        """Render highlight with comment operation."""
        ops = [
            CriticOperation(
                type=OperationType.HIGHLIGHT,
                text="check",
                comment="needs review",
                position=0,
            )
        ]
        result = render_criticmarkup(ops, "check this")
        assert result == "{==check=={>>needs review<<}} this"


class TestRoundTrip:
    """Tests for parse -> render round-trip consistency."""

    @pytest.mark.parametrize(
        "markup",
        [
            "Hello {++world++}!",
            "Say {--goodbye--}",
            "Change {~~old~>new~~} value",
            "Note {>>comment<<} here",
            "{==highlight==} this",
            "{==text=={>>note<<}}",
        ],
    )
    def test_roundtrip_single_operation(self, markup: str):
        """Single operations should round-trip correctly."""
        ops = parse_criticmarkup(markup)
        assert len(ops) == 1
        assert ops[0].type in OperationType


# =============================================================================
# DOCX to CriticMarkup Export Tests
# =============================================================================


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


class TestDocxToCriticmarkupExport:
    """Tests for DOCX to CriticMarkup export functionality."""

    def test_export_plain_text(self):
        """Export document with no tracked changes."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello world.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            result = doc.to_criticmarkup()
            assert result == "Hello world."
        finally:
            docx_path.unlink()

    def test_export_insertion(self):
        """Export document with tracked insertion."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello </w:t></w:r>
      <w:ins w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">
        <w:r><w:t>beautiful </w:t></w:r>
      </w:ins>
      <w:r><w:t>world.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            result = doc.to_criticmarkup()
            assert result == "Hello {++beautiful ++}world."
        finally:
            docx_path.unlink()

    def test_export_deletion(self):
        """Export document with tracked deletion."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello </w:t></w:r>
      <w:del w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">
        <w:r><w:delText>old </w:delText></w:r>
      </w:del>
      <w:r><w:t>world.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            result = doc.to_criticmarkup()
            assert result == "Hello {--old --}world."
        finally:
            docx_path.unlink()

    def test_export_insertion_and_deletion(self):
        """Export document with both insertion and deletion (substitution pattern)."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Payment in </w:t></w:r>
      <w:del w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">
        <w:r><w:delText>30</w:delText></w:r>
      </w:del>
      <w:ins w:id="2" w:author="Test" w:date="2024-01-01T00:00:00Z">
        <w:r><w:t>45</w:t></w:r>
      </w:ins>
      <w:r><w:t> days.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            result = doc.to_criticmarkup()
            assert result == "Payment in {--30--}{++45++} days."
        finally:
            docx_path.unlink()

    def test_export_multiple_paragraphs(self):
        """Export document with multiple paragraphs."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>First paragraph.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Second </w:t></w:r>
      <w:ins w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">
        <w:r><w:t>modified </w:t></w:r>
      </w:ins>
      <w:r><w:t>paragraph.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            result = doc.to_criticmarkup()
            assert "First paragraph." in result
            assert "Second {++modified ++}paragraph." in result
            # Paragraphs should be separated by double newlines
            assert "\n\n" in result
        finally:
            docx_path.unlink()

    def test_export_empty_document(self):
        """Export document with no content."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p></w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            result = doc.to_criticmarkup()
            assert result == ""
        finally:
            docx_path.unlink()

    def test_export_via_document_method(self):
        """Test that Document.to_criticmarkup() method works."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Test </w:t></w:r>
      <w:ins w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">
        <w:r><w:t>content</w:t></w:r>
      </w:ins>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            # Test the Document method directly
            result = doc.to_criticmarkup()
            assert "{++content++}" in result
        finally:
            docx_path.unlink()


# =============================================================================
# CriticMarkup to DOCX Import Tests
# =============================================================================


class TestApplyResult:
    """Tests for ApplyResult dataclass."""

    def test_apply_result_success_rate(self):
        """Test success rate calculation."""
        from python_docx_redline.criticmarkup import ApplyResult

        result = ApplyResult(total=10, successful=8, failed=2, errors=[])
        assert result.success_rate == 80.0

    def test_apply_result_empty(self):
        """Test success rate with no operations."""
        from python_docx_redline.criticmarkup import ApplyResult

        result = ApplyResult(total=0, successful=0, failed=0, errors=[])
        assert result.success_rate == 100.0

    def test_apply_result_repr(self):
        """Test string representation."""
        from python_docx_redline.criticmarkup import ApplyResult

        result = ApplyResult(total=5, successful=4, failed=1, errors=[])
        assert "4/5" in repr(result)
        assert "80.0%" in repr(result)


class TestCriticmarkupToDocxImport:
    """Tests for CriticMarkup to DOCX import functionality."""

    def test_import_deletion(self):
        """Import deletion from CriticMarkup."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello old world.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            # Apply deletion using CriticMarkup
            result = doc.apply_criticmarkup("Hello {--old --}world.", author="Test Author")

            assert result.successful == 1
            assert result.failed == 0

            # Verify the deletion was applied
            changes = doc.get_tracked_changes(change_type="deletion")
            assert len(changes) == 1
            assert changes[0].text == "old "
        finally:
            docx_path.unlink()

    def test_import_substitution(self):
        """Import substitution from CriticMarkup."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Payment in 30 days.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            # Apply substitution using CriticMarkup
            result = doc.apply_criticmarkup("Payment in {~~30~>45~~} days.", author="Test Author")

            assert result.successful == 1
            assert result.failed == 0

            # Verify the changes were applied
            insertions = doc.get_tracked_changes(change_type="insertion")
            deletions = doc.get_tracked_changes(change_type="deletion")

            # Substitution creates both a deletion and an insertion
            assert len(insertions) == 1
            assert len(deletions) == 1
            assert insertions[0].text == "45"
            assert deletions[0].text == "30"
        finally:
            docx_path.unlink()

    def test_import_insertion_with_context(self):
        """Import insertion using context to find location."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello world.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            # Apply insertion using CriticMarkup with context
            result = doc.apply_criticmarkup("Hello {++beautiful ++}world.", author="Test Author")

            assert result.successful == 1
            assert result.failed == 0

            # Verify the insertion was applied
            insertions = doc.get_tracked_changes(change_type="insertion")
            assert len(insertions) == 1
            assert insertions[0].text == "beautiful "
        finally:
            docx_path.unlink()

    def test_import_multiple_operations(self):
        """Import multiple operations from CriticMarkup."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>The old contract states payment in 30 days.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            # Apply multiple operations
            result = doc.apply_criticmarkup(
                "The {--old --}contract states payment in {~~30~>45~~} days.",
                author="Test Author",
            )

            # Should have 2 operations: 1 deletion + 1 substitution
            assert result.total == 2
            assert result.successful == 2
            assert result.failed == 0
        finally:
            docx_path.unlink()

    def test_import_via_document_method(self):
        """Test Document.apply_criticmarkup() method works."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Delete this text.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            # Test the Document method directly
            result = doc.apply_criticmarkup("Delete {--this --}text.", author="Method Test")

            assert result.successful == 1
            changes = doc.get_tracked_changes()
            assert len(changes) == 1
        finally:
            docx_path.unlink()

    def test_import_no_operations(self):
        """Import text with no CriticMarkup."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Plain text.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            result = doc.apply_criticmarkup("Plain text.")

            assert result.total == 0
            assert result.successful == 0
            assert result.failed == 0
        finally:
            docx_path.unlink()

    def test_import_stop_on_error(self):
        """Test stop_on_error parameter."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Some text here.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            # First operation will fail (text not found), second should not run
            result = doc.apply_criticmarkup("{--nonexistent--} {--text--}", stop_on_error=True)

            # Should stop after first failure
            assert result.failed >= 1
            assert len(result.errors) >= 1
        finally:
            docx_path.unlink()


# =============================================================================
# Integration Tests: Round-Trip Verification
# =============================================================================


class TestCriticmarkupRoundTrip:
    """Integration tests for CriticMarkup round-trip workflow.

    These tests verify the complete workflow:
    1. Start with a document
    2. Make changes (tracked insertions/deletions)
    3. Export to CriticMarkup
    4. Verify CriticMarkup contains the changes
    5. Apply CriticMarkup to a fresh document
    6. Verify the changes were applied correctly
    """

    def test_roundtrip_deletion(self):
        """Test round-trip: create deletion -> export -> import."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello beautiful world.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            # Step 1: Create a document and make a tracked deletion
            doc1 = Document(docx_path)
            doc1.delete_tracked("beautiful ", author="Test")

            # Step 2: Export to CriticMarkup
            markup = doc1.to_criticmarkup()

            # Step 3: Verify the export contains deletion markup
            assert "{--beautiful --}" in markup

            # Step 4: Create a fresh document and apply the markup
            doc2 = Document(docx_path)
            result = doc2.apply_criticmarkup(markup, author="Import")

            # Step 5: Verify the change was applied
            assert result.successful >= 1
            deletions = doc2.get_tracked_changes(change_type="deletion")
            assert len(deletions) >= 1
        finally:
            docx_path.unlink()

    def test_roundtrip_insertion(self):
        """Test round-trip: create insertion -> export -> verify markup."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello world.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            # Step 1: Create a document and make a tracked insertion
            doc = Document(docx_path)
            doc.insert_tracked("beautiful ", after="Hello ", author="Test")

            # Step 2: Export to CriticMarkup
            markup = doc.to_criticmarkup()

            # Step 3: Verify the export contains insertion markup
            assert "{++beautiful ++}" in markup
            assert "Hello" in markup
            assert "world" in markup
        finally:
            docx_path.unlink()

    def test_roundtrip_substitution_pattern(self):
        """Test round-trip with deletion+insertion pattern (like substitution)."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Payment due in 30 days.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            # Step 1: Create a substitution (delete + insert)
            doc1 = Document(docx_path)
            doc1.replace_tracked("30", "45", author="Test")

            # Step 2: Export to CriticMarkup
            markup = doc1.to_criticmarkup()

            # Step 3: Verify the export contains the changes
            assert "{--30--}" in markup
            assert "{++45++}" in markup

            # Step 4: Apply to fresh document
            doc2 = Document(docx_path)
            result = doc2.apply_criticmarkup("Payment due in {~~30~>45~~} days.", author="Import")

            # Step 5: Verify changes were applied
            assert result.successful == 1
            insertions = doc2.get_tracked_changes(change_type="insertion")
            deletions = doc2.get_tracked_changes(change_type="deletion")
            assert len(insertions) == 1
            assert len(deletions) == 1
        finally:
            docx_path.unlink()

    def test_roundtrip_multiple_paragraphs(self):
        """Test round-trip with multiple paragraphs."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>First paragraph content.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Second paragraph with text.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Third paragraph here.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            # Step 1: Make changes across multiple paragraphs
            doc = Document(docx_path)
            doc.delete_tracked("content", author="Test")
            doc.delete_tracked("with text", author="Test")

            # Step 2: Export to CriticMarkup
            markup = doc.to_criticmarkup()

            # Step 3: Verify all paragraphs are in the export
            assert "First paragraph" in markup
            assert "Second paragraph" in markup
            assert "Third paragraph" in markup

            # Step 4: Verify deletions are marked
            assert "{--content--}" in markup
            assert "{--with text--}" in markup

            # Step 5: Verify paragraphs are separated
            assert "\n\n" in markup
        finally:
            docx_path.unlink()

    def test_roundtrip_preserves_plain_text(self):
        """Test that plain text without changes survives round-trip."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>This is plain text without any changes.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            # Export without any changes
            markup = doc.to_criticmarkup()

            # Verify plain text is preserved without any markup
            assert markup == "This is plain text without any changes."
            assert "{++" not in markup
            assert "{--" not in markup
        finally:
            docx_path.unlink()

    def test_roundtrip_existing_tracked_changes(self):
        """Test that existing tracked changes in document are exported correctly."""
        from python_docx_redline import Document

        # Document that already has tracked changes
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>The </w:t></w:r>
      <w:del w:id="1" w:author="Original" w:date="2024-01-01T00:00:00Z">
        <w:r><w:delText>old </w:delText></w:r>
      </w:del>
      <w:ins w:id="2" w:author="Original" w:date="2024-01-01T00:00:00Z">
        <w:r><w:t>new </w:t></w:r>
      </w:ins>
      <w:r><w:t>contract.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            # Export should show existing tracked changes
            markup = doc.to_criticmarkup()

            assert "{--old --}" in markup
            assert "{++new ++}" in markup
            assert "The" in markup
            assert "contract" in markup
        finally:
            docx_path.unlink()

    def test_roundtrip_empty_document(self):
        """Test round-trip with empty/minimal document."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p></w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)

            # Export empty document
            markup = doc.to_criticmarkup()

            # Should be empty string
            assert markup == ""
        finally:
            docx_path.unlink()

    def test_workflow_edit_exported_markdown(self):
        """Test the complete workflow: export, edit markdown, import changes."""
        from python_docx_redline import Document

        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>The agreement requires 30 days notice.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        try:
            # Step 1: Export original document
            doc1 = Document(docx_path)
            original_markup = doc1.to_criticmarkup()
            assert original_markup == "The agreement requires 30 days notice."

            # Step 2: Simulate user editing the markdown with CriticMarkup
            edited_markup = "The {++revised ++}agreement requires {~~30~>45~~} days notice."

            # Step 3: Apply edits to a fresh document
            doc2 = Document(docx_path)
            result = doc2.apply_criticmarkup(edited_markup, author="Reviewer")

            # Step 4: Verify all edits were applied
            assert result.total == 2  # 1 insertion + 1 substitution
            assert result.successful == 2

            # Step 5: Verify tracked changes exist
            all_changes = doc2.get_tracked_changes()
            assert len(all_changes) >= 2
        finally:
            docx_path.unlink()
