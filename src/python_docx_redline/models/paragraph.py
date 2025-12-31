"""
Paragraph wrapper class for convenient access to paragraph elements.
"""

import copy
from typing import TYPE_CHECKING

from lxml import etree

from python_docx_redline.constants import WORD_NAMESPACE
from python_docx_redline.markdown_parser import parse_markdown

if TYPE_CHECKING:
    from python_docx_redline.models.section import Section


class Paragraph:
    """Wrapper around a w:p (paragraph) element.

    Provides convenient Python API for working with paragraphs.
    """

    def __init__(self, element: etree._Element):
        """Initialize Paragraph wrapper.

        Args:
            element: The w:p XML element to wrap
        """
        if element.tag != f"{{{WORD_NAMESPACE}}}p":
            raise ValueError(f"Expected w:p element, got {element.tag}")
        self._element = element
        self._parent_section: Section | None = None

    @property
    def element(self) -> etree._Element:
        """Get the underlying XML element."""
        return self._element

    @property
    def text(self) -> str:
        """Get all text content from the paragraph.

        Extracts text from both w:t and w:delText elements, avoiding XML structural whitespace.
        This ensures continuous text isn't broken by XML formatting between runs.
        Includes text from tracked deletions (w:delText) as well as regular text (w:t).

        Returns:
            Combined text from all runs in the paragraph
        """
        # Extract text from both w:t and w:delText elements to avoid XML structural whitespace
        # Using .// finds all descendants within the paragraph
        text_elements = self._element.findall(f".//{{{WORD_NAMESPACE}}}t")
        deltext_elements = self._element.findall(f".//{{{WORD_NAMESPACE}}}delText")
        return "".join(elem.text or "" for elem in text_elements + deltext_elements)

    @text.setter
    def text(self, value: str) -> None:
        """Set the text content of the paragraph.

        This replaces all content with new runs while preserving paragraph
        properties (w:pPr). Supports markdown formatting:
        - **bold** -> bold text
        - *italic* -> italic text
        - ++underline++ -> underlined text
        - ~~strikethrough~~ -> strikethrough text

        Args:
            value: New text content (may include markdown formatting)
        """
        # Preserve paragraph properties (w:pPr)
        ppr = self._element.find(f"{{{WORD_NAMESPACE}}}pPr")
        preserved_ppr = copy.deepcopy(ppr) if ppr is not None else None

        # Remove all content elements EXCEPT pPr
        # This includes runs, hyperlinks, bookmarks, tracked changes, etc.
        elements_to_remove = []
        for child in self._element:
            if child.tag != f"{{{WORD_NAMESPACE}}}pPr":
                elements_to_remove.append(child)
        for elem in elements_to_remove:
            self._element.remove(elem)

        # Restore pPr if it was removed or ensure it's first
        if preserved_ppr is not None:
            # Check if pPr still exists (it shouldn't have been removed, but ensure position)
            existing_ppr = self._element.find(f"{{{WORD_NAMESPACE}}}pPr")
            if existing_ppr is None:
                self._element.insert(0, preserved_ppr)

        # Parse markdown and create runs
        segments = parse_markdown(value)

        # Handle empty text case
        if not segments:
            # Create a single empty run to maintain valid structure
            run = etree.SubElement(self._element, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = value
            return

        for segment in segments:
            run = etree.SubElement(self._element, f"{{{WORD_NAMESPACE}}}r")

            # Add run properties if any formatting is applied
            if segment.has_formatting():
                rpr = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}rPr")
                if segment.bold:
                    etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}b")
                if segment.italic:
                    etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}i")
                if segment.underline:
                    u_elem = etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}u")
                    u_elem.set(f"{{{WORD_NAMESPACE}}}val", "single")
                if segment.strikethrough:
                    etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}strike")

            # Handle linebreak segments
            if segment.is_linebreak:
                etree.SubElement(run, f"{{{WORD_NAMESPACE}}}br")
            else:
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                # Preserve whitespace if needed
                if segment.text and (segment.text[0].isspace() or segment.text[-1].isspace()):
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = segment.text

    @property
    def style(self) -> str | None:
        """Get the paragraph style.

        Returns:
            Style name (e.g., 'Heading1', 'Normal') or None if no style set
        """
        # Look for w:pPr/w:pStyle
        p_pr = self._element.find(f"{{{WORD_NAMESPACE}}}pPr")
        if p_pr is None:
            return None

        p_style = p_pr.find(f"{{{WORD_NAMESPACE}}}pStyle")
        if p_style is None:
            return None

        return p_style.get(f"{{{WORD_NAMESPACE}}}val")

    @style.setter
    def style(self, value: str | None) -> None:
        """Set the paragraph style.

        Args:
            value: Style name (e.g., 'Heading1', 'Normal') or None to remove style
        """
        # Get or create w:pPr
        p_pr = self._element.find(f"{{{WORD_NAMESPACE}}}pPr")
        if p_pr is None:
            p_pr = etree.Element(f"{{{WORD_NAMESPACE}}}pPr")
            # Insert at beginning of paragraph
            self._element.insert(0, p_pr)

        # Get or create w:pStyle
        p_style = p_pr.find(f"{{{WORD_NAMESPACE}}}pStyle")

        if value is None:
            # Remove style if it exists
            if p_style is not None:
                p_pr.remove(p_style)
        else:
            if p_style is None:
                p_style = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}pStyle")
            p_style.set(f"{{{WORD_NAMESPACE}}}val", value)

    @property
    def runs(self) -> list[etree._Element]:
        """Get all run (w:r) elements in this paragraph.

        Returns:
            List of w:r elements
        """
        return list(self._element.findall(f"{{{WORD_NAMESPACE}}}r"))

    def is_heading(self) -> bool:
        """Check if this paragraph is a heading.

        Returns:
            True if paragraph style starts with 'Heading', False otherwise
        """
        style = self.style
        if style is None:
            return False
        return style.startswith("Heading")

    def get_heading_level(self) -> int | None:
        """Get the heading level if this is a heading paragraph.

        Returns:
            Heading level (1-9) or None if not a heading
        """
        if not self.is_heading():
            return None

        style = self.style
        if style is None:
            return None

        # Extract number from style like "Heading1" or "Heading2"
        # Handle both "Heading1" and "heading 1" formats
        style_lower = style.lower()
        if "heading" in style_lower:
            # Try to extract the number
            import re

            match = re.search(r"(\d+)", style)
            if match:
                level = int(match.group(1))
                if 1 <= level <= 9:
                    return level

        return None

    def contains(self, text: str, case_sensitive: bool = True) -> bool:
        """Check if paragraph contains specific text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case sensitive

        Returns:
            True if text is found in paragraph
        """
        para_text = self.text
        if not case_sensitive:
            para_text = para_text.lower()
            text = text.lower()
        return text in para_text

    @property
    def parent_section(self) -> "Section | None":
        """Get the parent Section object if this paragraph belongs to one.

        Returns:
            Parent Section or None if not set
        """
        return self._parent_section

    def _set_parent_section(self, section: "Section | None") -> None:
        """Set the parent section (internal use only).

        Args:
            section: The parent Section object
        """
        self._parent_section = section

    def __repr__(self) -> str:
        """String representation of the paragraph."""
        text_preview = self.text[:50] + "..." if len(self.text) > 50 else self.text
        style_info = f" style={self.style}" if self.style else ""
        return f"<Paragraph{style_info}: {text_preview!r}>"
