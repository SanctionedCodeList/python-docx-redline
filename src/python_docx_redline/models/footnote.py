"""
Footnote and Endnote model classes for python_docx_redline.

These classes represent footnotes and endnotes in a Word document,
providing a high-level interface for accessing and manipulating them.
"""

from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from python_docx_redline.document import Document
    from python_docx_redline.models.paragraph import Paragraph

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class Footnote:
    """Represents a footnote in a Word document.

    Footnotes appear at the bottom of the page and are referenced by
    superscript numbers or symbols in the document text.

    Attributes:
        id: The footnote ID (unique identifier)
        element: The underlying lxml Element
        document: Reference to the parent Document
    """

    def __init__(self, element: etree._Element, document: "Document") -> None:
        """Initialize a Footnote from an XML element.

        Args:
            element: The w:footnote XML element
            document: Reference to the parent Document
        """
        self.element = element
        self.document = document
        self._id = element.get(f"{{{WORD_NAMESPACE}}}id")

    @property
    def id(self) -> str:
        """Get the footnote ID."""
        return self._id

    @property
    def paragraphs(self) -> list["Paragraph"]:
        """Get all paragraphs in the footnote.

        Returns:
            List of Paragraph objects
        """
        from python_docx_redline.models.paragraph import Paragraph

        para_elements = self.element.findall(f"{{{WORD_NAMESPACE}}}p")
        return [Paragraph(p) for p in para_elements]

    @property
    def text(self) -> str:
        """Get the text content of the footnote.

        Returns:
            All text from all paragraphs concatenated
        """
        return "\n".join(p.text for p in self.paragraphs)

    def contains(self, text: str, case_sensitive: bool = True) -> bool:
        """Check if the footnote contains specific text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case-sensitive

        Returns:
            True if text is found
        """
        footnote_text = self.text
        search_text = text

        if not case_sensitive:
            footnote_text = footnote_text.lower()
            search_text = search_text.lower()

        return search_text in footnote_text

    def __repr__(self) -> str:
        """Return string representation of the footnote."""
        preview = self.text[:50].replace("\n", " ")
        if len(self.text) > 50:
            preview += "..."
        return f'<Footnote id="{self.id}": "{preview}">'


class Endnote:
    """Represents an endnote in a Word document.

    Endnotes appear at the end of the document or section and are referenced
    by superscript numbers or symbols in the document text.

    Attributes:
        id: The endnote ID (unique identifier)
        element: The underlying lxml Element
        document: Reference to the parent Document
    """

    def __init__(self, element: etree._Element, document: "Document") -> None:
        """Initialize an Endnote from an XML element.

        Args:
            element: The w:endnote XML element
            document: Reference to the parent Document
        """
        self.element = element
        self.document = document
        self._id = element.get(f"{{{WORD_NAMESPACE}}}id")

    @property
    def id(self) -> str:
        """Get the endnote ID."""
        return self._id

    @property
    def paragraphs(self) -> list["Paragraph"]:
        """Get all paragraphs in the endnote.

        Returns:
            List of Paragraph objects
        """
        from python_docx_redline.models.paragraph import Paragraph

        para_elements = self.element.findall(f"{{{WORD_NAMESPACE}}}p")
        return [Paragraph(p) for p in para_elements]

    @property
    def text(self) -> str:
        """Get the text content of the endnote.

        Returns:
            All text from all paragraphs concatenated
        """
        return "\n".join(p.text for p in self.paragraphs)

    def contains(self, text: str, case_sensitive: bool = True) -> bool:
        """Check if the endnote contains specific text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case-sensitive

        Returns:
            True if text is found
        """
        endnote_text = self.text
        search_text = text

        if not case_sensitive:
            endnote_text = endnote_text.lower()
            search_text = search_text.lower()

        return search_text in endnote_text

    def __repr__(self) -> str:
        """Return string representation of the endnote."""
        preview = self.text[:50].replace("\n", " ")
        if len(self.text) > 50:
            preview += "..."
        return f'<Endnote id="{self.id}": "{preview}">'
