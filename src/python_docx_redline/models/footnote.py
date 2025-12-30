"""
Footnote and Endnote model classes for python_docx_redline.

These classes represent footnotes and endnotes in a Word document,
providing a high-level interface for accessing and manipulating them.
"""

from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from python_docx_redline.constants import WORD_NAMESPACE

if TYPE_CHECKING:
    from python_docx_redline.document import Document
    from python_docx_redline.models.paragraph import Paragraph


@dataclass
class FootnoteReference:
    """Represents the location where a footnote/endnote is referenced in the document.

    This dataclass captures the paragraph and run element where a note reference
    appears, along with the character position within the paragraph text.

    Attributes:
        paragraph: The Paragraph object containing the reference
        run_element: The w:r element containing the footnoteReference/endnoteReference
        position_in_paragraph: Character offset where the reference appears in paragraph text
    """

    paragraph: "Paragraph"
    run_element: etree._Element
    position_in_paragraph: int


@dataclass
class OrphanedFootnote:
    """Represents an orphaned footnote that has no reference in the document body.

    Orphaned footnotes occur when text containing footnote markers is deleted
    but the footnote content remains in footnotes.xml.

    Attributes:
        id: The footnote ID (unique identifier)
        text: The text content of the orphaned footnote
    """

    id: str
    text: str


@dataclass
class OrphanedEndnote:
    """Represents an orphaned endnote that has no reference in the document body.

    Orphaned endnotes occur when text containing endnote markers is deleted
    but the endnote content remains in endnotes.xml.

    Attributes:
        id: The endnote ID (unique identifier)
        text: The text content of the orphaned endnote
    """

    id: str
    text: str


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
            All text from all paragraphs concatenated, with the leading
            space (added by Word after footnoteRef) stripped.
        """
        full_text = "\n".join(p.text for p in self.paragraphs)
        # Strip the leading space that Word adds after footnoteRef
        return full_text.lstrip(" ") if full_text.startswith(" ") else full_text

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

    @property
    def formatted_text(self) -> list[dict]:
        """Get text with formatting information.

        Extracts text content along with formatting properties for each run
        in the footnote. Useful for programmatic access to rich text content.

        Returns:
            List of dicts, one per text run, with structure:
            {
                "text": str,        # The text content
                "bold": bool,       # True if bold
                "italic": bool,     # True if italic
                "underline": bool,  # True if underlined
                "strikethrough": bool,  # True if strikethrough
                "paragraph_index": int  # Which paragraph (0-indexed)
            }

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> for run in footnote.formatted_text:
            ...     if run["bold"]:
            ...         print(f"Bold: {run['text']}")
        """
        return _extract_formatted_text(self.element)

    @property
    def html(self) -> str:
        """Get footnote content as simple HTML.

        Converts the footnote content to HTML, preserving basic formatting
        like bold, italic, underline, and strikethrough.

        Returns:
            HTML string with paragraphs wrapped in <p> tags and formatting
            using <b>, <i>, <u>, and <s> tags.

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> print(footnote.html)
            '<p><b>Bold</b> and <i>italic</i> text</p>'
        """
        return _convert_to_html(self.element)

    @property
    def reference_location(self) -> FootnoteReference | None:
        """Get the location where this footnote is referenced in the document.

        Returns:
            FootnoteReference with paragraph, run element, and position,
            or None if reference not found in document.

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> ref = footnote.reference_location
            >>> if ref:
            ...     print(f"Referenced in: {ref.paragraph.text[:50]}")
        """
        if self.document is None:
            return None

        return self.document._note_ops.get_footnote_reference_location(self.id)

    def edit(
        self,
        new_text: str,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Edit the footnote text content.

        Args:
            new_text: The new text content for the footnote
            track: If True, track the edit as a change (Phase 3 feature)
            author: Author name for tracked changes (uses document author if None)

        Raises:
            ValueError: If document reference is not available

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> footnote.edit("Updated citation text")
        """
        if self.document is None:
            raise ValueError("Cannot edit footnote: no document reference")

        self.document.edit_footnote(self.id, new_text, track=track, author=author)

    def delete(self, renumber: bool = True) -> None:
        """Delete this footnote from the document.

        This removes both the footnote content and its reference in the document.

        Args:
            renumber: If True, renumber remaining footnotes sequentially (default)

        Raises:
            ValueError: If document reference is not available

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> footnote.delete()
        """
        if self.document is None:
            raise ValueError("Cannot delete footnote: no document reference")

        self.document.delete_footnote(self.id, renumber=renumber)

    def insert_tracked(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
    ) -> None:
        """Insert text with tracked changes inside this footnote.

        Searches for anchor text within the footnote and inserts new text
        as a tracked insertion (w:ins) either after or before it.

        Args:
            text: The text to insert (supports markdown formatting)
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If both or neither of after/before specified,
                       or if no document reference
            TextNotFoundError: If anchor text not found in footnote
            AmbiguousTextError: If anchor text found multiple times

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> footnote.insert_tracked(" [updated]", after="citation")
        """
        if self.document is None:
            raise ValueError("Cannot insert in footnote: no document reference")

        self.document.insert_tracked_in_footnote(
            self.id, text, after=after, before=before, author=author
        )

    def delete_tracked(self, text: str, author: str | None = None) -> None:
        """Delete text with tracked changes inside this footnote.

        Searches for text within the footnote and marks it as a tracked
        deletion (w:del).

        Args:
            text: The text to delete
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If no document reference
            TextNotFoundError: If text not found in footnote
            AmbiguousTextError: If text found multiple times

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> footnote.delete_tracked("outdated reference")
        """
        if self.document is None:
            raise ValueError("Cannot delete in footnote: no document reference")

        self.document.delete_tracked_in_footnote(self.id, text, author=author)

    def replace_tracked(self, find: str, replace: str, author: str | None = None) -> None:
        """Replace text with tracked changes inside this footnote.

        Searches for text within the footnote and replaces it, showing both
        the deletion of old text and insertion of new text as tracked changes.

        Args:
            find: The text to find and replace
            replace: The replacement text (supports markdown formatting)
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If no document reference
            TextNotFoundError: If find text not found in footnote
            AmbiguousTextError: If find text found multiple times

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> footnote.replace_tracked("2020", "2024")
        """
        if self.document is None:
            raise ValueError("Cannot replace in footnote: no document reference")

        self.document.replace_tracked_in_footnote(self.id, find, replace, author=author)

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
            All text from all paragraphs concatenated, with the leading
            space (added by Word after endnoteRef) stripped.
        """
        full_text = "\n".join(p.text for p in self.paragraphs)
        # Strip the leading space that Word adds after endnoteRef
        return full_text.lstrip(" ") if full_text.startswith(" ") else full_text

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

    @property
    def formatted_text(self) -> list[dict]:
        """Get text with formatting information.

        Extracts text content along with formatting properties for each run
        in the endnote. Useful for programmatic access to rich text content.

        Returns:
            List of dicts, one per text run, with structure:
            {
                "text": str,        # The text content
                "bold": bool,       # True if bold
                "italic": bool,     # True if italic
                "underline": bool,  # True if underlined
                "strikethrough": bool,  # True if strikethrough
                "paragraph_index": int  # Which paragraph (0-indexed)
            }

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> for run in endnote.formatted_text:
            ...     if run["bold"]:
            ...         print(f"Bold: {run['text']}")
        """
        return _extract_formatted_text(self.element)

    @property
    def html(self) -> str:
        """Get endnote content as simple HTML.

        Converts the endnote content to HTML, preserving basic formatting
        like bold, italic, underline, and strikethrough.

        Returns:
            HTML string with paragraphs wrapped in <p> tags and formatting
            using <b>, <i>, <u>, and <s> tags.

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> print(endnote.html)
            '<p><b>Bold</b> and <i>italic</i> text</p>'
        """
        return _convert_to_html(self.element)

    @property
    def reference_location(self) -> FootnoteReference | None:
        """Get the location where this endnote is referenced in the document.

        Returns:
            FootnoteReference with paragraph, run element, and position,
            or None if reference not found in document.

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> ref = endnote.reference_location
            >>> if ref:
            ...     print(f"Referenced in: {ref.paragraph.text[:50]}")
        """
        if self.document is None:
            return None

        return self.document._note_ops.get_endnote_reference_location(self.id)

    def edit(
        self,
        new_text: str,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Edit the endnote text content.

        Args:
            new_text: The new text content for the endnote
            track: If True, track the edit as a change (Phase 3 feature)
            author: Author name for tracked changes (uses document author if None)

        Raises:
            ValueError: If document reference is not available

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> endnote.edit("Updated citation text")
        """
        if self.document is None:
            raise ValueError("Cannot edit endnote: no document reference")

        self.document.edit_endnote(self.id, new_text, track=track, author=author)

    def delete(self, renumber: bool = True) -> None:
        """Delete this endnote from the document.

        This removes both the endnote content and its reference in the document.

        Args:
            renumber: If True, renumber remaining endnotes sequentially (default)

        Raises:
            ValueError: If document reference is not available

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> endnote.delete()
        """
        if self.document is None:
            raise ValueError("Cannot delete endnote: no document reference")

        self.document.delete_endnote(self.id, renumber=renumber)

    def insert_tracked(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
    ) -> None:
        """Insert text with tracked changes inside this endnote.

        Searches for anchor text within the endnote and inserts new text
        as a tracked insertion (w:ins) either after or before it.

        Args:
            text: The text to insert (supports markdown formatting)
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If both or neither of after/before specified,
                       or if no document reference
            TextNotFoundError: If anchor text not found in endnote
            AmbiguousTextError: If anchor text found multiple times

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> endnote.insert_tracked(" [see also]", after="reference")
        """
        if self.document is None:
            raise ValueError("Cannot insert in endnote: no document reference")

        self.document.insert_tracked_in_endnote(
            self.id, text, after=after, before=before, author=author
        )

    def delete_tracked(self, text: str, author: str | None = None) -> None:
        """Delete text with tracked changes inside this endnote.

        Searches for text within the endnote and marks it as a tracked
        deletion (w:del).

        Args:
            text: The text to delete
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If no document reference
            TextNotFoundError: If text not found in endnote
            AmbiguousTextError: If text found multiple times

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> endnote.delete_tracked("obsolete citation")
        """
        if self.document is None:
            raise ValueError("Cannot delete in endnote: no document reference")

        self.document.delete_tracked_in_endnote(self.id, text, author=author)

    def replace_tracked(self, find: str, replace: str, author: str | None = None) -> None:
        """Replace text with tracked changes inside this endnote.

        Searches for text within the endnote and replaces it, showing both
        the deletion of old text and insertion of new text as tracked changes.

        Args:
            find: The text to find and replace
            replace: The replacement text (supports markdown formatting)
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If no document reference
            TextNotFoundError: If find text not found in endnote
            AmbiguousTextError: If find text found multiple times

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> endnote.replace_tracked("ibid", "op. cit.")
        """
        if self.document is None:
            raise ValueError("Cannot replace in endnote: no document reference")

        self.document.replace_tracked_in_endnote(self.id, find, replace, author=author)

    def __repr__(self) -> str:
        """Return string representation of the endnote."""
        preview = self.text[:50].replace("\n", " ")
        if len(self.text) > 50:
            preview += "..."
        return f'<Endnote id="{self.id}": "{preview}">'


# Helper functions for formatting extraction


def _extract_formatted_text(note_element: etree._Element) -> list[dict]:
    """Extract formatted text from a footnote or endnote element.

    Args:
        note_element: The w:footnote or w:endnote XML element

    Returns:
        List of dicts with text and formatting info
    """
    result = []
    para_elements = note_element.findall(f"{{{WORD_NAMESPACE}}}p")

    for para_idx, para in enumerate(para_elements):
        runs = para.findall(f"{{{WORD_NAMESPACE}}}r")

        for run in runs:
            # Get text content
            text_parts = []
            for t_elem in run.findall(f"{{{WORD_NAMESPACE}}}t"):
                text_parts.append(t_elem.text or "")
            for t_elem in run.findall(f"{{{WORD_NAMESPACE}}}delText"):
                text_parts.append(t_elem.text or "")

            text = "".join(text_parts)
            # Skip empty runs and whitespace-only runs (like the space after footnoteRef)
            if not text or text.isspace():
                continue

            # Get run properties
            rpr = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            bold = False
            italic = False
            underline = False
            strikethrough = False

            if rpr is not None:
                bold = rpr.find(f"{{{WORD_NAMESPACE}}}b") is not None
                italic = rpr.find(f"{{{WORD_NAMESPACE}}}i") is not None
                underline = rpr.find(f"{{{WORD_NAMESPACE}}}u") is not None
                strikethrough = rpr.find(f"{{{WORD_NAMESPACE}}}strike") is not None

            result.append(
                {
                    "text": text,
                    "bold": bold,
                    "italic": italic,
                    "underline": underline,
                    "strikethrough": strikethrough,
                    "paragraph_index": para_idx,
                }
            )

    return result


def _convert_to_html(note_element: etree._Element) -> str:
    """Convert a footnote or endnote element to HTML.

    Args:
        note_element: The w:footnote or w:endnote XML element

    Returns:
        HTML string representation
    """
    import html

    formatted_text = _extract_formatted_text(note_element)

    if not formatted_text:
        return ""

    # Group runs by paragraph
    paragraphs: dict[int, list[dict]] = {}
    for run in formatted_text:
        para_idx = run["paragraph_index"]
        if para_idx not in paragraphs:
            paragraphs[para_idx] = []
        paragraphs[para_idx].append(run)

    # Build HTML
    html_parts = []
    for para_idx in sorted(paragraphs.keys()):
        para_runs = paragraphs[para_idx]
        para_html = []

        for run in para_runs:
            text = html.escape(run["text"])

            # Apply formatting tags (innermost to outermost)
            if run["strikethrough"]:
                text = f"<s>{text}</s>"
            if run["underline"]:
                text = f"<u>{text}</u>"
            if run["italic"]:
                text = f"<i>{text}</i>"
            if run["bold"]:
                text = f"<b>{text}</b>"

            para_html.append(text)

        html_parts.append(f"<p>{''.join(para_html)}</p>")

    return "".join(html_parts)
