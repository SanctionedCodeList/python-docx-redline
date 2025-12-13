"""
Markdown parser for Word document formatting.

This module provides a customized markdown parser that converts markdown-formatted
text into a list of text segments with formatting flags. The parser supports a
"Word-flavored" markdown syntax:

- *italic* or _italic_ -> italic text
- **bold** or __bold__ -> bold text
- ++underline++ -> underlined text (custom extension)
- ~~strikethrough~~ -> strikethrough text

The output is a list of TextSegment objects that can be converted to OOXML runs
with appropriate formatting properties.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from re import Match
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from mistune.core import InlineState
    from mistune.inline_parser import InlineParser as MistuneInlineParser
    from mistune.markdown import Markdown


@dataclass
class TextSegment:
    """A segment of text with formatting information.

    Attributes:
        text: The actual text content
        bold: Whether the text should be bold
        italic: Whether the text should be italic
        underline: Whether the text should be underlined
        strikethrough: Whether the text should have strikethrough
        is_linebreak: Whether this segment represents a line break (<w:br/>)
    """

    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    is_linebreak: bool = False

    def has_formatting(self) -> bool:
        """Check if this segment has any formatting applied."""
        return self.bold or self.italic or self.underline or self.strikethrough

    def copy_with_text(self, new_text: str) -> TextSegment:
        """Create a copy with different text but same formatting."""
        return TextSegment(
            text=new_text,
            bold=self.bold,
            italic=self.italic,
            underline=self.underline,
            strikethrough=self.strikethrough,
            is_linebreak=self.is_linebreak,
        )


# Underline pattern: ++text++ (similar to strikethrough pattern)
_UNDERLINE_END = re.compile(r"(?:[^\s+])\+\+(?!\+)")


def _parse_underline(inline: MistuneInlineParser, m: Match[str], state: InlineState) -> int | None:
    """Parse ++underline++ syntax."""
    pos = m.end()
    m1 = _UNDERLINE_END.search(state.src, pos)
    if not m1:
        return None
    end_pos = m1.end()
    text = state.src[pos : end_pos - 2]
    new_state = state.copy()
    new_state.src = text
    children = inline.render(new_state)
    state.append_token({"type": "underline", "children": children})
    return end_pos


def _underline_plugin(md: Markdown) -> None:
    """Register the ++underline++ syntax with mistune."""
    md.inline.register(
        "underline",
        r"\+\+(?=[^\s+])",
        _parse_underline,
        before="link",
    )


class SegmentRenderer:
    """Custom mistune renderer that outputs TextSegment objects.

    This renderer collects text segments with formatting information
    instead of producing HTML output.
    """

    NAME = "segment"

    def __init__(self) -> None:
        """Initialize the renderer."""
        self._segments: list[TextSegment] = []
        self._format_stack: list[dict[str, bool]] = [
            {"bold": False, "italic": False, "underline": False, "strikethrough": False}
        ]

    def reset(self) -> None:
        """Reset the renderer state for a new parse."""
        self._segments = []
        self._format_stack = [
            {"bold": False, "italic": False, "underline": False, "strikethrough": False}
        ]

    def get_segments(self) -> list[TextSegment]:
        """Get the accumulated segments and merge adjacent segments with same formatting."""
        return _merge_segments(self._segments)

    def _current_format(self) -> dict[str, bool]:
        """Get the current formatting state."""
        return self._format_stack[-1].copy()

    def _push_format(self, **kwargs: bool) -> None:
        """Push a new formatting state with modifications."""
        current = self._current_format()
        current.update(kwargs)
        self._format_stack.append(current)

    def _pop_format(self) -> None:
        """Pop the current formatting state."""
        if len(self._format_stack) > 1:
            self._format_stack.pop()

    def _add_text(self, text: str) -> None:
        """Add text with current formatting."""
        if text:
            fmt = self._current_format()
            self._segments.append(
                TextSegment(
                    text=text,
                    bold=fmt["bold"],
                    italic=fmt["italic"],
                    underline=fmt["underline"],
                    strikethrough=fmt["strikethrough"],
                )
            )

    # Required renderer methods for mistune
    def text(self, text: str) -> str:
        """Handle plain text."""
        self._add_text(text)
        return ""

    def emphasis(self, text: str) -> str:
        """Handle *italic* text - called after children are rendered."""
        return ""

    def strong(self, text: str) -> str:
        """Handle **bold** text - called after children are rendered."""
        return ""

    def strikethrough(self, text: str) -> str:
        """Handle ~~strikethrough~~ text - called after children are rendered."""
        return ""

    def underline(self, text: str) -> str:
        """Handle ++underline++ text - called after children are rendered."""
        return ""

    def codespan(self, text: str) -> str:
        """Handle `code` - treat as plain text."""
        self._add_text(text)
        return ""

    def linebreak(self) -> str:
        """Handle hard line breaks - emit a linebreak segment for <w:br/>."""
        self._segments.append(TextSegment(text="", is_linebreak=True))
        return ""

    def softbreak(self) -> str:
        """Handle soft breaks (single newline in source)."""
        self._add_text(" ")
        return ""

    def paragraph(self, text: str) -> str:
        """Handle paragraphs - just pass through."""
        return ""

    def link(self, text: str, url: str, title: str | None = None) -> str:
        """Handle links - just use the text."""
        return ""

    def image(self, text: str, url: str, title: str | None = None) -> str:
        """Handle images - ignore."""
        return ""

    def inline_html(self, html: str) -> str:
        """Handle inline HTML - treat as text."""
        self._add_text(html)
        return ""

    def _render_children(self, children: list[dict[str, Any]]) -> None:
        """Recursively render child tokens."""
        for token in children:
            tok_type = token["type"]

            if tok_type == "text":
                self._add_text(token.get("raw", ""))
            elif tok_type == "emphasis":
                self._push_format(italic=True)
                if "children" in token:
                    self._render_children(token["children"])
                self._pop_format()
            elif tok_type == "strong":
                self._push_format(bold=True)
                if "children" in token:
                    self._render_children(token["children"])
                self._pop_format()
            elif tok_type == "strikethrough":
                self._push_format(strikethrough=True)
                if "children" in token:
                    self._render_children(token["children"])
                self._pop_format()
            elif tok_type == "underline":
                self._push_format(underline=True)
                if "children" in token:
                    self._render_children(token["children"])
                self._pop_format()
            elif tok_type == "codespan":
                self._add_text(token.get("raw", ""))
            elif tok_type == "softbreak":
                self._add_text(" ")
            elif tok_type == "linebreak":
                self._segments.append(TextSegment(text="", is_linebreak=True))
            elif tok_type == "link":
                # For links, just render the children (link text)
                if "children" in token:
                    self._render_children(token["children"])
            elif "children" in token:
                # Generic handling for other token types with children
                self._render_children(token["children"])

    def __call__(self, tokens: list[dict[str, Any]], state: Any) -> str:
        """Render a list of tokens."""
        self._render_children(tokens)
        return ""


def _merge_segments(segments: list[TextSegment]) -> list[TextSegment]:
    """Merge adjacent segments with identical formatting.

    Args:
        segments: List of TextSegment objects

    Returns:
        Merged list where adjacent segments with same formatting are combined.
        Linebreak segments are never merged with other segments.
    """
    if not segments:
        return []

    merged: list[TextSegment] = []
    current = segments[0]

    for next_seg in segments[1:]:
        # Never merge linebreak segments
        if current.is_linebreak or next_seg.is_linebreak:
            if current.text or current.is_linebreak:
                merged.append(current)
            current = next_seg
        # Check if formatting matches
        elif (
            current.bold == next_seg.bold
            and current.italic == next_seg.italic
            and current.underline == next_seg.underline
            and current.strikethrough == next_seg.strikethrough
        ):
            # Merge text
            current = current.copy_with_text(current.text + next_seg.text)
        else:
            # Different formatting, save current and start new
            if current.text:  # Only add non-empty segments
                merged.append(current)
            current = next_seg

    # Add final segment
    if current.text or current.is_linebreak:
        merged.append(current)

    return merged


@dataclass
class MarkdownParser:
    """Parser for Word-flavored markdown to formatted text segments.

    This parser uses mistune for robust markdown parsing but outputs
    TextSegment objects instead of HTML.

    Supported syntax:
        - *italic* or _italic_ -> italic
        - **bold** or __bold__ -> bold
        - ++underline++ -> underline
        - ~~strikethrough~~ -> strikethrough
        - \\* -> escaped asterisk (literal *)
        - Combinations like ***bold italic*** work
    """

    _md: Any = field(init=False, repr=False)
    _renderer: SegmentRenderer = field(init=False, repr=False)

    def __post_init__(self) -> None:
        """Initialize the mistune-based parser."""
        import mistune

        self._renderer = SegmentRenderer()
        self._md = mistune.create_markdown(
            renderer=self._renderer,  # type: ignore[arg-type]
            plugins=["strikethrough", _underline_plugin],
        )

    def parse(self, text: str) -> list[TextSegment]:
        """Parse markdown text into formatted segments.

        Args:
            text: Markdown-formatted text

        Returns:
            List of TextSegment objects with formatting information
        """
        if not text:
            return []

        # Preserve leading/trailing whitespace that mistune would strip
        leading_ws = ""
        trailing_ws = ""
        stripped = text.lstrip()
        if len(stripped) < len(text):
            leading_ws = text[: len(text) - len(stripped)]
        stripped_end = text.rstrip()
        if len(stripped_end) < len(text):
            trailing_ws = text[len(stripped_end) :]

        # Reset renderer state
        self._renderer.reset()

        # Parse the text - this populates the renderer's segments
        self._md(text)

        segments = self._renderer.get_segments()

        # Handle whitespace-only input: if parsing returned nothing but we had text,
        # return a single plain segment with the original text (for xml:space="preserve")
        if not segments and text:
            return [TextSegment(text=text)]

        # Restore leading/trailing whitespace
        if segments and leading_ws:
            first = segments[0]
            segments[0] = first.copy_with_text(leading_ws + first.text)
        if segments and trailing_ws:
            last = segments[-1]
            segments[-1] = last.copy_with_text(last.text + trailing_ws)

        return segments


def parse_markdown(text: str) -> list[TextSegment]:
    """Convenience function to parse markdown text.

    This function creates a new parser instance each call for thread safety.
    The performance cost is minimal since parsing is typically done once per
    document operation.

    Args:
        text: Markdown-formatted text

    Returns:
        List of TextSegment objects with formatting information

    Example:
        >>> segments = parse_markdown("This is **bold** and *italic*")
        >>> for seg in segments:
        ...     print(f"{seg.text!r}: bold={seg.bold}, italic={seg.italic}")
        'This is ': bold=False, italic=False
        'bold': bold=True, italic=False
        ' and ': bold=False, italic=False
        'italic': bold=False, italic=True
    """
    parser = MarkdownParser()
    return parser.parse(text)
