"""
Section wrapper class for document sections.

A Section represents a logical section of a document: a heading paragraph
followed by all paragraphs until the next heading.
"""

from typing import TYPE_CHECKING

from lxml import etree

from python_docx_redline.constants import WORD_NAMESPACE
from python_docx_redline.models.paragraph import Paragraph

if TYPE_CHECKING:
    pass


class Section:
    """Represents a logical section of a document.

    A section consists of:
    - A heading paragraph (or None for intro section)
    - All paragraphs until the next heading
    """

    def __init__(self, heading: Paragraph | None, paragraphs: list[Paragraph]):
        """Initialize a Section.

        Args:
            heading: The heading Paragraph, or None for intro section
            paragraphs: All paragraphs in this section (including heading)
        """
        self._heading = heading
        self._paragraphs = paragraphs

        # Set parent section for all paragraphs
        for para in paragraphs:
            para._set_parent_section(self)

    @property
    def heading(self) -> Paragraph | None:
        """Get the heading paragraph.

        Returns:
            The heading Paragraph or None for intro section
        """
        return self._heading

    @property
    def paragraphs(self) -> list[Paragraph]:
        """Get all paragraphs in this section.

        Returns:
            List of all paragraphs (including heading)
        """
        return self._paragraphs

    @property
    def heading_text(self) -> str | None:
        """Get the heading text.

        Returns:
            Heading text or None if no heading
        """
        if self._heading is None:
            return None
        return self._heading.text

    @property
    def heading_level(self) -> int | None:
        """Get the heading level.

        Returns:
            Heading level (1-9) or None if no heading
        """
        if self._heading is None:
            return None
        return self._heading.get_heading_level()

    def contains(self, text: str, case_sensitive: bool = True) -> bool:
        """Check if any paragraph in section contains text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case sensitive

        Returns:
            True if text found in any paragraph
        """
        return any(p.contains(text, case_sensitive) for p in self._paragraphs)

    def find_paragraph(self, text: str, case_sensitive: bool = True) -> Paragraph | None:
        """Find first paragraph containing text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case sensitive

        Returns:
            First matching Paragraph or None if not found
        """
        for para in self._paragraphs:
            if para.contains(text, case_sensitive):
                return para
        return None

    @classmethod
    def from_document(cls, xml_root: etree._Element) -> list["Section"]:
        """Parse document into sections.

        A section is defined as a heading paragraph + all following paragraphs
        until the next heading. Paragraphs before the first heading belong to
        an implicit intro section with no heading.

        Args:
            xml_root: The document root element

        Returns:
            List of Sections
        """
        # Get all paragraph elements
        all_p_elements = list(xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Wrap in Paragraph objects
        all_paragraphs = [Paragraph(p) for p in all_p_elements]

        # Group into sections
        sections: list[Section] = []
        current_heading: Paragraph | None = None
        current_paras: list[Paragraph] = []

        for para in all_paragraphs:
            if para.is_heading():
                # Start new section
                if current_heading is not None or current_paras:
                    # Save previous section
                    sections.append(cls(current_heading, current_paras))

                # Start new section with this heading
                current_heading = para
                current_paras = [para]
            else:
                # Add to current section
                current_paras.append(para)

        # Don't forget the last section
        if current_heading is not None or current_paras:
            sections.append(cls(current_heading, current_paras))

        return sections

    def __repr__(self) -> str:
        """String representation of the section."""
        if self._heading:
            heading_preview = self.heading_text or "(empty heading)"
            return f"<Section heading={heading_preview!r} paragraphs={len(self._paragraphs)}>"
        else:
            return f"<Section intro paragraphs={len(self._paragraphs)}>"

    def __len__(self) -> int:
        """Get number of paragraphs in section."""
        return len(self._paragraphs)
