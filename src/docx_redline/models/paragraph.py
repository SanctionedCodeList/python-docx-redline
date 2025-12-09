"""
Paragraph wrapper class for convenient access to paragraph elements.
"""

from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from docx_redline.models.section import Section

# Word namespace
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


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

        This replaces all runs with a single run containing the new text.

        Args:
            value: New text content
        """
        # Remove all existing runs
        for run in self._element.findall(f"{{{WORD_NAMESPACE}}}r"):
            self._element.remove(run)

        # Create new run with text
        run = etree.SubElement(self._element, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = value

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
