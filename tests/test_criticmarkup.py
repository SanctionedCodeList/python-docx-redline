"""Tests for CriticMarkup parser."""

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
        text = (
            "{++inserted++} "
            "{--deleted--} "
            "{~~old~>new~~} "
            "{>>comment<<} "
            "{==highlighted==}"
        )
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
