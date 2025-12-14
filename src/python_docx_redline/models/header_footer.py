"""
Header and Footer model classes for python_docx_redline.

These classes represent headers and footers in a Word document,
providing a high-level interface for accessing and manipulating them.

Headers and footers in OOXML are stored in separate XML files (header1.xml,
footer1.xml, etc.) and are linked via relationships in sectPr elements.
"""

from enum import Enum
from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from python_docx_redline.document import Document
    from python_docx_redline.models.paragraph import Paragraph

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class HeaderFooterType(Enum):
    """Types of headers and footers in Word documents.

    Word supports three types of headers/footers per section:
    - DEFAULT: Used on all pages except first (if first is different) and even pages
    - FIRST: Used on the first page of the section (if enabled)
    - EVEN: Used on even-numbered pages (if different from odd)
    """

    DEFAULT = "default"
    FIRST = "first"
    EVEN = "even"


class Header:
    """Represents a header in a Word document.

    Headers appear at the top of pages and can contain text, images,
    page numbers, and other content.

    Attributes:
        element: The underlying lxml Element (w:hdr root element)
        document: Reference to the parent Document
        header_type: The type of header (default, first, even)
        rel_id: The relationship ID linking this header
    """

    def __init__(
        self,
        element: etree._Element,
        document: "Document",
        header_type: HeaderFooterType,
        rel_id: str,
        file_path: str,
    ) -> None:
        """Initialize a Header from an XML element.

        Args:
            element: The w:hdr XML element (root of header XML file)
            document: Reference to the parent Document
            header_type: The type of header (default, first, even)
            rel_id: The relationship ID (e.g., "rId7")
            file_path: Path to the header XML file within the docx
        """
        self.element = element
        self.document = document
        self.header_type = header_type
        self.rel_id = rel_id
        self.file_path = file_path

    @property
    def type(self) -> str:
        """Get the header type as a string.

        Returns:
            'default', 'first', or 'even'
        """
        return self.header_type.value

    @property
    def paragraphs(self) -> list["Paragraph"]:
        """Get all paragraphs in the header.

        Returns:
            List of Paragraph objects
        """
        from python_docx_redline.models.paragraph import Paragraph

        para_elements = self.element.findall(f".//{{{WORD_NAMESPACE}}}p")
        return [Paragraph(p) for p in para_elements]

    @property
    def text(self) -> str:
        """Get the text content of the header.

        Returns:
            All text from all paragraphs concatenated with newlines
        """
        return "\n".join(p.text for p in self.paragraphs)

    def contains(self, text: str, case_sensitive: bool = True) -> bool:
        """Check if the header contains specific text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case-sensitive

        Returns:
            True if text is found
        """
        header_text = self.text
        search_text = text

        if not case_sensitive:
            header_text = header_text.lower()
            search_text = search_text.lower()

        return search_text in header_text

    def __repr__(self) -> str:
        """Return string representation of the header."""
        preview = self.text[:50].replace("\n", " ")
        if len(self.text) > 50:
            preview += "..."
        return f'<Header type="{self.type}": "{preview}">'


class Footer:
    """Represents a footer in a Word document.

    Footers appear at the bottom of pages and can contain text, images,
    page numbers, and other content.

    Attributes:
        element: The underlying lxml Element (w:ftr root element)
        document: Reference to the parent Document
        footer_type: The type of footer (default, first, even)
        rel_id: The relationship ID linking this footer
    """

    def __init__(
        self,
        element: etree._Element,
        document: "Document",
        footer_type: HeaderFooterType,
        rel_id: str,
        file_path: str,
    ) -> None:
        """Initialize a Footer from an XML element.

        Args:
            element: The w:ftr XML element (root of footer XML file)
            document: Reference to the parent Document
            footer_type: The type of footer (default, first, even)
            rel_id: The relationship ID (e.g., "rId8")
            file_path: Path to the footer XML file within the docx
        """
        self.element = element
        self.document = document
        self.footer_type = footer_type
        self.rel_id = rel_id
        self.file_path = file_path

    @property
    def type(self) -> str:
        """Get the footer type as a string.

        Returns:
            'default', 'first', or 'even'
        """
        return self.footer_type.value

    @property
    def paragraphs(self) -> list["Paragraph"]:
        """Get all paragraphs in the footer.

        Returns:
            List of Paragraph objects
        """
        from python_docx_redline.models.paragraph import Paragraph

        para_elements = self.element.findall(f".//{{{WORD_NAMESPACE}}}p")
        return [Paragraph(p) for p in para_elements]

    @property
    def text(self) -> str:
        """Get the text content of the footer.

        Returns:
            All text from all paragraphs concatenated with newlines
        """
        return "\n".join(p.text for p in self.paragraphs)

    def contains(self, text: str, case_sensitive: bool = True) -> bool:
        """Check if the footer contains specific text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case-sensitive

        Returns:
            True if text is found
        """
        footer_text = self.text
        search_text = text

        if not case_sensitive:
            footer_text = footer_text.lower()
            search_text = search_text.lower()

        return search_text in footer_text

    def __repr__(self) -> str:
        """Return string representation of the footer."""
        preview = self.text[:50].replace("\n", " ")
        if len(self.text) > 50:
            preview += "..."
        return f'<Footer type="{self.type}": "{preview}">'
