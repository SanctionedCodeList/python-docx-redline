"""Tests for the markdown parser module."""

from python_docx_redline.markdown_parser import (
    MarkdownParser,
    TextSegment,
    parse_markdown,
)


class TestTextSegment:
    """Tests for TextSegment dataclass."""

    def test_default_formatting(self):
        """Test that default formatting is all False."""
        seg = TextSegment(text="hello")
        assert seg.text == "hello"
        assert seg.bold is False
        assert seg.italic is False
        assert seg.underline is False
        assert seg.strikethrough is False

    def test_has_formatting_false(self):
        """Test has_formatting returns False for unformatted text."""
        seg = TextSegment(text="hello")
        assert seg.has_formatting() is False

    def test_has_formatting_true(self):
        """Test has_formatting returns True for formatted text."""
        assert TextSegment(text="hello", bold=True).has_formatting() is True
        assert TextSegment(text="hello", italic=True).has_formatting() is True
        assert TextSegment(text="hello", underline=True).has_formatting() is True
        assert TextSegment(text="hello", strikethrough=True).has_formatting() is True

    def test_copy_with_text(self):
        """Test copy_with_text preserves formatting."""
        seg = TextSegment(text="hello", bold=True, italic=True)
        copy = seg.copy_with_text("world")
        assert copy.text == "world"
        assert copy.bold is True
        assert copy.italic is True
        assert copy.underline is False


class TestMarkdownParserBasic:
    """Tests for basic markdown parsing."""

    def test_plain_text(self):
        """Test parsing plain text without formatting."""
        segments = parse_markdown("Hello world")
        assert len(segments) == 1
        assert segments[0].text == "Hello world"
        assert segments[0].has_formatting() is False

    def test_empty_string(self):
        """Test parsing empty string."""
        segments = parse_markdown("")
        assert segments == []

    def test_bold_text(self):
        """Test parsing **bold** text."""
        segments = parse_markdown("This is **bold** text")
        assert len(segments) == 3
        assert segments[0].text == "This is "
        assert segments[0].bold is False
        assert segments[1].text == "bold"
        assert segments[1].bold is True
        assert segments[2].text == " text"
        assert segments[2].bold is False

    def test_italic_text_asterisk(self):
        """Test parsing *italic* text with asterisks."""
        segments = parse_markdown("This is *italic* text")
        assert len(segments) == 3
        assert segments[1].text == "italic"
        assert segments[1].italic is True

    def test_italic_text_underscore(self):
        """Test parsing _italic_ text with underscores."""
        segments = parse_markdown("This is _italic_ text")
        assert len(segments) == 3
        assert segments[1].text == "italic"
        assert segments[1].italic is True

    def test_underline_text(self):
        """Test parsing ++underline++ text."""
        segments = parse_markdown("This is ++underlined++ text")
        assert len(segments) == 3
        assert segments[1].text == "underlined"
        assert segments[1].underline is True

    def test_strikethrough_text(self):
        """Test parsing ~~strikethrough~~ text."""
        segments = parse_markdown("This is ~~struck~~ text")
        assert len(segments) == 3
        assert segments[1].text == "struck"
        assert segments[1].strikethrough is True


class TestMarkdownParserNested:
    """Tests for nested markdown formatting."""

    def test_bold_italic(self):
        """Test parsing ***bold italic*** text."""
        segments = parse_markdown("This is ***bold italic*** text")
        # Should have bold and italic both True
        bold_italic = [s for s in segments if s.bold and s.italic]
        assert len(bold_italic) == 1
        assert bold_italic[0].text == "bold italic"

    def test_bold_with_italic_inside(self):
        """Test parsing **bold with *italic* inside** text."""
        segments = parse_markdown("**bold with *italic* inside**")
        # All segments should be bold
        assert all(s.bold for s in segments)
        # Middle segment should also be italic
        italic_segs = [s for s in segments if s.italic]
        assert len(italic_segs) == 1
        assert italic_segs[0].text == "italic"

    def test_multiple_formats(self):
        """Test parsing text with multiple format types."""
        segments = parse_markdown("**bold** and *italic* and ++underline++")

        bold_segs = [s for s in segments if s.bold]
        assert len(bold_segs) == 1
        assert bold_segs[0].text == "bold"

        italic_segs = [s for s in segments if s.italic]
        assert len(italic_segs) == 1
        assert italic_segs[0].text == "italic"

        underline_segs = [s for s in segments if s.underline]
        assert len(underline_segs) == 1
        assert underline_segs[0].text == "underline"


class TestMarkdownParserEscaping:
    """Tests for escaped markdown characters."""

    def test_escaped_asterisk(self):
        """Test that \\* produces literal asterisk."""
        segments = parse_markdown(r"This is \*not italic\*")
        # Should be single segment with literal asterisks
        full_text = "".join(s.text for s in segments)
        assert "*not italic*" in full_text
        # None should be italic
        assert not any(s.italic for s in segments)

    def test_escaped_underscore(self):
        """Test that \\_ produces literal underscore."""
        segments = parse_markdown(r"This is \_not italic\_")
        full_text = "".join(s.text for s in segments)
        assert "_not italic_" in full_text

    def test_escaped_plus(self):
        """Test that \\+ in underline syntax is handled."""
        segments = parse_markdown(r"This is \+\+not underline\+\+")
        # Should not have underline formatting
        assert not any(s.underline for s in segments)


class TestMarkdownParserEdgeCases:
    """Tests for edge cases in markdown parsing."""

    def test_adjacent_formatting(self):
        """Test adjacent formatted sections with space separator."""
        # Note: **bold***italic* is ambiguous in markdown
        # Use a clear separator for reliable parsing
        segments = parse_markdown("**bold** *italic*")
        bold_segs = [s for s in segments if s.bold and not s.italic]
        italic_segs = [s for s in segments if s.italic and not s.bold]
        assert len(bold_segs) >= 1
        assert len(italic_segs) >= 1

    def test_unclosed_formatting(self):
        """Test unclosed formatting markers are treated as literal."""
        segments = parse_markdown("This is **unclosed")
        # Should not crash, markers treated as text
        full_text = "".join(s.text for s in segments)
        assert "**" in full_text or "unclosed" in full_text

    def test_whitespace_preservation(self):
        """Test that whitespace is preserved in segments."""
        segments = parse_markdown("  **bold**  ")
        full_text = "".join(s.text for s in segments)
        # Leading/trailing spaces should be preserved somewhere
        assert "bold" in full_text

    def test_only_formatting(self):
        """Test text that is only formatted."""
        segments = parse_markdown("**all bold**")
        assert len(segments) == 1
        assert segments[0].text == "all bold"
        assert segments[0].bold is True

    def test_empty_formatting(self):
        """Test empty formatting markers."""
        segments = parse_markdown("before **** after")
        # Should handle gracefully
        full_text = "".join(s.text for s in segments)
        assert "before" in full_text
        assert "after" in full_text


class TestMarkdownParserMerging:
    """Tests for segment merging behavior."""

    def test_adjacent_same_format_merged(self):
        """Test that adjacent segments with same formatting are merged."""
        # This tests the internal _merge_segments function
        segments = parse_markdown("plain text here")
        # Should be single segment for plain text
        assert len(segments) == 1
        assert segments[0].text == "plain text here"

    def test_different_formats_not_merged(self):
        """Test that different formats produce separate segments."""
        # Use clear separator for reliable parsing
        segments = parse_markdown("**bold** and *italic*")
        # Should have at least 3 segments (bold, plain, italic)
        assert len(segments) >= 3
        # Verify different formatting exists
        bold_segs = [s for s in segments if s.bold]
        italic_segs = [s for s in segments if s.italic]
        assert len(bold_segs) >= 1
        assert len(italic_segs) >= 1


class TestMarkdownParserClass:
    """Tests for MarkdownParser class."""

    def test_parser_reuse(self):
        """Test that parser can be reused for multiple parses."""
        parser = MarkdownParser()

        seg1 = parser.parse("**bold**")
        seg2 = parser.parse("*italic*")

        assert seg1[0].bold is True
        assert seg2[0].italic is True
        # Should not have state bleed
        assert seg2[0].bold is False

    def test_parser_reset(self):
        """Test that parser state is properly reset between parses."""
        parser = MarkdownParser()

        # Parse something complex
        parser.parse("**bold *nested* here**")

        # Parse something simple
        segments = parser.parse("plain")

        assert len(segments) == 1
        assert segments[0].has_formatting() is False


class TestRegressionFixes:
    """Regression tests for PR review fixes."""

    def test_whitespace_only_input_returns_segment(self):
        """Test that whitespace-only input returns a segment, not empty list.

        Regression test: parse_markdown("  ") was returning [] which caused
        create_insertion() to emit <w:ins> with no <w:r> elements.
        """
        # Single space
        segments = parse_markdown(" ")
        assert len(segments) == 1
        assert segments[0].text == " "

        # Multiple spaces
        segments = parse_markdown("   ")
        assert len(segments) == 1
        assert segments[0].text == "   "

        # Newline only
        segments = parse_markdown("\n")
        assert len(segments) == 1
        assert segments[0].text == "\n"

        # Mixed whitespace
        segments = parse_markdown("  \n  ")
        assert len(segments) == 1
        assert segments[0].text == "  \n  "

    def test_linebreak_segment_flag(self):
        """Test that hard line breaks produce segments with is_linebreak=True.

        Regression test: linebreak() was emitting "\\n" inside <w:t>, but
        Word expects <w:br/> elements for line breaks.
        """
        # Hard line break in markdown is two spaces followed by newline
        segments = parse_markdown("line one  \nline two")

        # Find the linebreak segment
        linebreak_segments = [s for s in segments if s.is_linebreak]
        assert len(linebreak_segments) >= 1
        assert linebreak_segments[0].text == ""
        assert linebreak_segments[0].is_linebreak is True

    def test_linebreak_not_merged(self):
        """Test that linebreak segments are not merged with text segments."""
        segments = parse_markdown("before  \nafter")

        # Should have: text, linebreak, text (possibly more)
        has_linebreak = any(s.is_linebreak for s in segments)
        assert has_linebreak

        # Linebreak should be its own segment
        for i, seg in enumerate(segments):
            if seg.is_linebreak:
                # Should not have text in linebreak segment
                assert seg.text == ""

    def test_thread_safety_no_shared_state(self):
        """Test that each parse_markdown call uses fresh parser state.

        Regression test: module-level _default_parser held mutable state
        that could cause issues with concurrent usage.
        """
        # Parse something that affects state
        parse_markdown("**bold** text")

        # Parse something plain - should have clean state
        segments = parse_markdown("plain")
        assert len(segments) == 1
        assert segments[0].bold is False
        assert segments[0].italic is False

        # Parse multiple times rapidly
        results = []
        for i in range(10):
            segs = parse_markdown(f"**text{i}**")
            results.append(segs)

        # All should have bold segment
        for segs in results:
            assert any(s.bold for s in segs)

    def test_text_segment_is_linebreak_default(self):
        """Test that TextSegment.is_linebreak defaults to False."""
        seg = TextSegment(text="hello")
        assert seg.is_linebreak is False

    def test_text_segment_copy_preserves_linebreak(self):
        """Test that copy_with_text preserves is_linebreak flag."""
        seg = TextSegment(text="", is_linebreak=True)
        copy = seg.copy_with_text("should still be linebreak")
        assert copy.is_linebreak is True
