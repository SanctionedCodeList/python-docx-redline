"""
Text and Markdown export from AccessibilityTree.

This module provides text and markdown export functionality that replaces
the pandoc dependency for document conversion workflows.
"""

from __future__ import annotations

import textwrap
from abc import ABC, abstractmethod
from dataclasses import dataclass
from io import StringIO
from typing import TYPE_CHECKING, Literal

from .types import AccessibilityNode, ElementType

if TYPE_CHECKING:
    from .tree import AccessibilityTree


@dataclass
class TextExportConfig:
    """Configuration for text/markdown export.

    Attributes:
        include_headers: Include document headers
        include_footers: Include document footers
        include_footnotes: Include footnotes
        include_endnotes: Include endnotes
        include_comments: Include comments as inline markers
        include_images: Include image placeholders
        tracked_changes: How to handle tracked changes
            - "accept": Show final state (insertions visible, deletions hidden)
            - "reject": Show original state (deletions visible, insertions hidden)
            - "all": Show both with CriticMarkup syntax {++ins++} {--del--}
        table_format: Table rendering format
            - "markdown": Standard markdown pipe tables
            - "simple": Simple pipe-separated format
            - "grid": Grid format with borders
        line_width: Line width for text wrapping (0 = no wrap)
        heading_style: Heading style for plain text
            - "underline": Use underline characters
            - "prefix": Use # prefix like markdown
    """

    include_headers: bool = True
    include_footers: bool = True
    include_footnotes: bool = True
    include_endnotes: bool = True
    include_comments: bool = False
    include_images: bool = True
    tracked_changes: Literal["accept", "reject", "all"] = "accept"
    table_format: Literal["markdown", "simple", "grid"] = "markdown"
    line_width: int = 0
    heading_style: Literal["underline", "prefix"] = "underline"


class TextExporter(ABC):
    """Base class for text export.

    This class provides common functionality for exporting AccessibilityTree
    to text-based formats. Subclasses implement format-specific rendering.
    """

    def __init__(self, tree: AccessibilityTree, config: TextExportConfig) -> None:
        """Initialize the exporter.

        Args:
            tree: The AccessibilityTree to export
            config: Export configuration
        """
        self.tree = tree
        self.config = config
        self.buffer = StringIO()
        self.footnotes: list[tuple[int, str]] = []
        self.endnotes: list[tuple[int, str]] = []
        self._footnote_counter = 0
        self._endnote_counter = 0

    def export(self) -> str:
        """Export the tree to text format.

        Returns:
            The exported text string
        """
        self.buffer = StringIO()
        self.footnotes = []
        self.endnotes = []
        self._footnote_counter = 0
        self._endnote_counter = 0

        # Render body content
        for child in self.tree.root.children:
            self._render_node(child)

        # Render document parts at end
        self._render_document_parts()

        return self.buffer.getvalue().rstrip() + "\n"

    def _render_node(self, node: AccessibilityNode) -> None:
        """Render a node to the buffer.

        Args:
            node: The node to render
        """
        if node.element_type == ElementType.PARAGRAPH:
            if node.level is not None:
                self._render_heading(node)
            else:
                self._render_paragraph(node)
        elif node.element_type == ElementType.TABLE:
            self._render_table(node)
        elif node.element_type == ElementType.TABLE_ROW:
            # Tables handle their own rows
            pass
        elif node.element_type == ElementType.TABLE_CELL:
            # Tables handle their own cells
            pass

    def _get_text_with_tracked_changes(self, node: AccessibilityNode) -> str:
        """Get text content with tracked changes applied.

        Args:
            node: The node to get text from

        Returns:
            Text with tracked changes applied according to config
        """
        # Check if node has changes
        changes = node.properties.get("_changes")
        if not changes or not isinstance(changes, list):
            return node.text

        # For accept mode (default), we want insertions but not deletions
        # For reject mode, we want deletions but not insertions
        # For all mode, we show both with CriticMarkup
        if self.config.tracked_changes == "accept":
            # The node.text already includes insertions by default
            # We need to remove deletion text from the combined text
            return self._apply_accept_mode(node.text, changes)
        elif self.config.tracked_changes == "reject":
            return self._apply_reject_mode(node.text, changes)
        else:  # "all"
            return self._apply_all_mode(node.text, changes)

    def _apply_accept_mode(self, text: str, changes: list[dict]) -> str:
        """Apply accept mode: show insertions, hide deletions.

        In accept mode, the base text already includes both insertion and deletion
        text. We need to remove the deletion text.
        """
        result = text
        for change in changes:
            if change.get("type") == "deletion":
                del_text = change.get("text", "")
                if del_text and del_text in result:
                    result = result.replace(del_text, "", 1)
        return result

    def _apply_reject_mode(self, text: str, changes: list[dict]) -> str:
        """Apply reject mode: show deletions, hide insertions.

        In reject mode, we want the original text (deletions visible, insertions hidden).
        """
        result = text
        for change in changes:
            if change.get("type") == "insertion":
                ins_text = change.get("text", "")
                if ins_text and ins_text in result:
                    result = result.replace(ins_text, "", 1)
        return result

    def _apply_all_mode(self, text: str, changes: list[dict]) -> str:
        """Apply all mode: show both with CriticMarkup syntax.

        CriticMarkup uses:
        - {++insertion++} for added text
        - {--deletion--} for removed text
        """
        result = text

        # Process changes in reverse order to maintain positions
        # First mark insertions, then deletions
        for change in changes:
            change_text = change.get("text", "")
            if not change_text:
                continue

            if change.get("type") == "insertion":
                if change_text in result:
                    result = result.replace(change_text, f"{{++{change_text}++}}", 1)
            elif change.get("type") == "deletion":
                if change_text in result:
                    result = result.replace(change_text, f"{{--{change_text}--}}", 1)

        return result

    def _render_images(self, node: AccessibilityNode) -> None:
        """Render image placeholders for a node.

        Args:
            node: Node potentially containing images
        """
        if not self.config.include_images or not node.has_images:
            return

        for img in node.images:
            self._render_image(img.name or img.alt_text or "unnamed image")

    def _render_comments(self, node: AccessibilityNode) -> None:
        """Render comments for a node.

        Args:
            node: Node potentially containing comments
        """
        if not self.config.include_comments or not node.has_comments:
            return

        for comment in node.comments:
            self._render_comment(comment.author, comment.text)

    def _render_document_parts(self) -> None:
        """Render headers, footers, footnotes, and endnotes at the end."""
        has_parts = False

        # Check if we have any document parts to render
        if self.config.include_footnotes and self.footnotes:
            has_parts = True
        if self.config.include_endnotes and self.endnotes:
            has_parts = True
        # Headers/footers would be added here if tree had them

        if not has_parts:
            return

        # Add separator
        self._write("\n---\n\n")

        # Render footnotes
        if self.config.include_footnotes and self.footnotes:
            self._render_footnotes_section()

        # Render endnotes
        if self.config.include_endnotes and self.endnotes:
            self._render_endnotes_section()

    @abstractmethod
    def _render_paragraph(self, node: AccessibilityNode) -> None:
        """Render a paragraph node.

        Args:
            node: The paragraph node
        """
        pass

    @abstractmethod
    def _render_heading(self, node: AccessibilityNode) -> None:
        """Render a heading node.

        Args:
            node: The heading node
        """
        pass

    @abstractmethod
    def _render_table(self, node: AccessibilityNode) -> None:
        """Render a table node.

        Args:
            node: The table node
        """
        pass

    @abstractmethod
    def _render_image(self, name: str) -> None:
        """Render an image placeholder.

        Args:
            name: The image name or alt text
        """
        pass

    @abstractmethod
    def _render_comment(self, author: str, text: str) -> None:
        """Render a comment.

        Args:
            author: Comment author
            text: Comment text
        """
        pass

    @abstractmethod
    def _render_footnotes_section(self) -> None:
        """Render the footnotes section."""
        pass

    @abstractmethod
    def _render_endnotes_section(self) -> None:
        """Render the endnotes section."""
        pass

    def _write(self, text: str) -> None:
        """Write text to the buffer.

        Args:
            text: Text to write
        """
        self.buffer.write(text)

    def _wrap_text(self, text: str) -> str:
        """Wrap text according to line_width config.

        Args:
            text: Text to wrap

        Returns:
            Wrapped text
        """
        if self.config.line_width <= 0:
            return text
        return textwrap.fill(text, width=self.config.line_width)


class PlainTextExporter(TextExporter):
    """Exporter for plain text format."""

    # Underline characters for different heading levels
    UNDERLINE_CHARS = {
        1: "=",
        2: "-",
        3: "~",
        4: ".",
        5: "'",
        6: "`",
    }

    def _render_paragraph(self, node: AccessibilityNode) -> None:
        """Render a paragraph as plain text."""
        text = self._get_text_with_tracked_changes(node)

        # Check for block quote style
        style = (node.style or "").lower()
        if "quote" in style or "blockquote" in style:
            self._render_block_quote(text)
        else:
            text = self._wrap_text(text)
            self._write(f"{text}\n\n")

        # Render images after paragraph
        self._render_images(node)

        # Render comments after paragraph
        self._render_comments(node)

    def _render_heading(self, node: AccessibilityNode) -> None:
        """Render a heading as plain text."""
        text = self._get_text_with_tracked_changes(node)
        level = node.level or 1

        if self.config.heading_style == "underline":
            underline_char = self.UNDERLINE_CHARS.get(level, "-")
            underline = underline_char * len(text)
            self._write(f"{text}\n{underline}\n\n")
        else:  # prefix style
            prefix = "#" * level
            self._write(f"{prefix} {text}\n\n")

    def _render_block_quote(self, text: str) -> None:
        """Render a block quote (indented).

        Args:
            text: The quote text
        """
        lines = text.split("\n")
        for line in lines:
            self._write(f"    {line}\n")
        self._write("\n")

    def _render_table(self, node: AccessibilityNode) -> None:
        """Render a table as plain text."""
        if self.config.table_format == "grid":
            self._render_table_grid(node)
        else:  # simple or markdown (same for plain text)
            self._render_table_simple(node)

    def _render_table_simple(self, node: AccessibilityNode) -> None:
        """Render a table in simple pipe format."""
        rows = node.children
        if not rows:
            return

        # Calculate column widths
        col_widths = self._calculate_column_widths(rows)

        # Render rows
        for i, row in enumerate(rows):
            cells = row.children
            cell_texts = [self._get_cell_text(cell) for cell in cells]

            # Pad cells to column widths
            padded = [
                text.ljust(col_widths[j]) if j < len(col_widths) else text
                for j, text in enumerate(cell_texts)
            ]

            self._write(" | ".join(padded).rstrip() + "\n")

            # Add separator after header row
            if i == 0 and row.properties.get("header") == "true":
                separator = "-+-".join("-" * w for w in col_widths)
                self._write(separator + "\n")

        self._write("\n")

    def _render_table_grid(self, node: AccessibilityNode) -> None:
        """Render a table in grid format with borders."""
        rows = node.children
        if not rows:
            return

        col_widths = self._calculate_column_widths(rows)

        def make_border(char: str = "-", sep: str = "+") -> str:
            parts = [char * (w + 2) for w in col_widths]
            return sep + sep.join(parts) + sep

        # Top border
        self._write(make_border("-", "+") + "\n")

        for i, row in enumerate(rows):
            cells = row.children
            cell_texts = [self._get_cell_text(cell) for cell in cells]

            # Pad cells
            padded = [
                f" {text.ljust(col_widths[j])} " if j < len(col_widths) else f" {text} "
                for j, text in enumerate(cell_texts)
            ]

            self._write("|" + "|".join(padded) + "|\n")

            # Header separator or row border
            if i == 0 and row.properties.get("header") == "true":
                self._write(make_border("=", "+") + "\n")
            else:
                self._write(make_border("-", "+") + "\n")

        self._write("\n")

    def _calculate_column_widths(self, rows: list[AccessibilityNode]) -> list[int]:
        """Calculate column widths for a table.

        Args:
            rows: Table rows

        Returns:
            List of column widths
        """
        if not rows:
            return []

        num_cols = max(len(row.children) for row in rows)
        widths = [0] * num_cols

        for row in rows:
            for i, cell in enumerate(row.children):
                if i < num_cols:
                    cell_text = self._get_cell_text(cell)
                    widths[i] = max(widths[i], len(cell_text))

        return widths

    def _get_cell_text(self, cell: AccessibilityNode) -> str:
        """Get text from a table cell.

        Args:
            cell: The cell node

        Returns:
            Cell text content
        """
        return self._get_text_with_tracked_changes(cell)

    def _render_image(self, name: str) -> None:
        """Render an image placeholder."""
        self._write(f"[image: {name}]\n\n")

    def _render_comment(self, author: str, text: str) -> None:
        """Render a comment."""
        self._write(f"[Comment by {author}: {text}]\n\n")

    def _render_footnotes_section(self) -> None:
        """Render the footnotes section."""
        self._write("Footnotes:\n")
        for num, text in self.footnotes:
            self._write(f"[{num}] {text}\n")
        self._write("\n")

    def _render_endnotes_section(self) -> None:
        """Render the endnotes section."""
        self._write("Endnotes:\n")
        for num, text in self.endnotes:
            self._write(f"[{num}] {text}\n")
        self._write("\n")


class MarkdownExporter(TextExporter):
    """Exporter for markdown format."""

    def _render_paragraph(self, node: AccessibilityNode) -> None:
        """Render a paragraph as markdown."""
        text = self._get_text_with_tracked_changes(node)

        # Check for block quote style
        style = (node.style or "").lower()
        if "quote" in style or "blockquote" in style:
            self._render_block_quote(text)
        else:
            text = self._wrap_text(text)
            self._write(f"{text}\n\n")

        # Render images after paragraph
        self._render_images(node)

        # Render comments after paragraph
        self._render_comments(node)

    def _render_heading(self, node: AccessibilityNode) -> None:
        """Render a heading as markdown."""
        text = self._get_text_with_tracked_changes(node)
        level = node.level or 1
        prefix = "#" * level
        self._write(f"{prefix} {text}\n\n")

    def _render_block_quote(self, text: str) -> None:
        """Render a block quote with > prefix.

        Args:
            text: The quote text
        """
        lines = text.split("\n")
        for line in lines:
            self._write(f"> {line}\n")
        self._write("\n")

    def _render_table(self, node: AccessibilityNode) -> None:
        """Render a table as markdown."""
        if self.config.table_format == "grid":
            self._render_table_grid(node)
        elif self.config.table_format == "simple":
            self._render_table_simple(node)
        else:  # markdown (default)
            self._render_table_markdown(node)

    def _render_table_markdown(self, node: AccessibilityNode) -> None:
        """Render a table in standard markdown format."""
        rows = node.children
        if not rows:
            return

        col_widths = self._calculate_column_widths(rows)

        for i, row in enumerate(rows):
            cells = row.children
            cell_texts = [self._get_cell_text(cell) for cell in cells]

            # Pad cells to column widths
            padded = [
                text.ljust(col_widths[j]) if j < len(col_widths) else text
                for j, text in enumerate(cell_texts)
            ]

            self._write("| " + " | ".join(padded) + " |\n")

            # Add markdown header separator after first row
            if i == 0:
                separator = "|" + "|".join("-" * (w + 2) for w in col_widths) + "|"
                self._write(separator + "\n")

        self._write("\n")

    def _render_table_simple(self, node: AccessibilityNode) -> None:
        """Render a table in simple pipe format."""
        rows = node.children
        if not rows:
            return

        col_widths = self._calculate_column_widths(rows)

        for i, row in enumerate(rows):
            cells = row.children
            cell_texts = [self._get_cell_text(cell) for cell in cells]

            padded = [
                text.ljust(col_widths[j]) if j < len(col_widths) else text
                for j, text in enumerate(cell_texts)
            ]

            self._write(" | ".join(padded).rstrip() + "\n")

            if i == 0:
                separator = "-+-".join("-" * w for w in col_widths)
                self._write(separator + "\n")

        self._write("\n")

    def _render_table_grid(self, node: AccessibilityNode) -> None:
        """Render a table in grid format."""
        rows = node.children
        if not rows:
            return

        col_widths = self._calculate_column_widths(rows)

        def make_border(char: str = "-", sep: str = "+") -> str:
            parts = [char * (w + 2) for w in col_widths]
            return sep + sep.join(parts) + sep

        self._write(make_border("-", "+") + "\n")

        for i, row in enumerate(rows):
            cells = row.children
            cell_texts = [self._get_cell_text(cell) for cell in cells]

            padded = [
                f" {text.ljust(col_widths[j])} " if j < len(col_widths) else f" {text} "
                for j, text in enumerate(cell_texts)
            ]

            self._write("|" + "|".join(padded) + "|\n")

            if i == 0 and row.properties.get("header") == "true":
                self._write(make_border("=", "+") + "\n")
            else:
                self._write(make_border("-", "+") + "\n")

        self._write("\n")

    def _calculate_column_widths(self, rows: list[AccessibilityNode]) -> list[int]:
        """Calculate column widths for a table."""
        if not rows:
            return []

        num_cols = max(len(row.children) for row in rows)
        widths = [0] * num_cols

        for row in rows:
            for i, cell in enumerate(row.children):
                if i < num_cols:
                    cell_text = self._get_cell_text(cell)
                    widths[i] = max(widths[i], len(cell_text))

        return widths

    def _get_cell_text(self, cell: AccessibilityNode) -> str:
        """Get text from a table cell."""
        return self._get_text_with_tracked_changes(cell)

    def _render_image(self, name: str) -> None:
        """Render an image placeholder in markdown."""
        # Use markdown image syntax but as placeholder
        self._write(f"[image: {name}]\n\n")

    def _render_comment(self, author: str, text: str) -> None:
        """Render a comment in markdown."""
        self._write(f"[Comment by {author}: {text}]\n\n")

    def _render_footnotes_section(self) -> None:
        """Render the footnotes section in markdown format."""
        for num, text in self.footnotes:
            self._write(f"[^{num}]: {text}\n")
        self._write("\n")

    def _render_endnotes_section(self) -> None:
        """Render the endnotes section."""
        self._write("**Endnotes:**\n\n")
        for num, text in self.endnotes:
            self._write(f"[^e{num}]: {text}\n")
        self._write("\n")


__all__ = [
    "TextExportConfig",
    "TextExporter",
    "PlainTextExporter",
    "MarkdownExporter",
]
