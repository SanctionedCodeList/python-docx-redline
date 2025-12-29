"""
HyperlinkOperations class for handling hyperlinks in Word documents.

This module provides a dedicated class for all hyperlink operations,
including inserting, reading, editing, and removing hyperlinks in
document body, headers, footers, footnotes, and endnotes.

Supports both external hyperlinks (URLs) and internal hyperlinks
(bookmarks within the document).

OOXML Structure:
    External: <w:hyperlink r:id="rId5" w:tooltip="..."><w:r>...</w:r></w:hyperlink>
    Internal: <w:hyperlink w:anchor="BookmarkName"><w:r>...</w:r></w:hyperlink>
"""

from __future__ import annotations

import warnings
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from lxml import etree

from ..accessibility.bookmarks import BookmarkRegistry
from ..constants import OFFICE_RELATIONSHIPS_NAMESPACE, WORD_NAMESPACE
from ..errors import AmbiguousTextError, TextNotFoundError
from ..relationships import RelationshipManager, RelationshipTypes
from ..scope import ScopeEvaluator
from ..style_templates import ensure_standard_styles
from ..text_search import TextSpan

if TYPE_CHECKING:
    from ..document import Document


@dataclass
class HyperlinkInfo:
    """Information about a hyperlink in the document.

    Attributes:
        ref: Unique reference ID for the hyperlink (e.g., "lnk:5")
        text: The display text of the hyperlink
        target: The URL (external) or bookmark name (internal)
        is_external: True if this is an external URL, False if internal bookmark
        tooltip: Optional tooltip text shown on hover
        location: Where the hyperlink is located ("body", "header", "footer", etc.)
        r_id: Relationship ID for external links (e.g., "rId5"), None for internal
    """

    ref: str
    text: str
    target: str
    is_external: bool
    tooltip: str | None = None
    location: str = "body"
    r_id: str | None = None


class HyperlinkOperations:
    """Handles hyperlink operations for Word documents.

    This class encapsulates all hyperlink functionality, including:
    - Inserting external hyperlinks (URLs)
    - Inserting internal hyperlinks (to bookmarks)
    - Hyperlinks in headers, footers, footnotes, and endnotes
    - Reading and querying existing hyperlinks
    - Editing hyperlink URLs, text, and anchors
    - Removing hyperlinks

    The class takes a Document reference and operates on its XML structure.

    Example:
        >>> # Usually accessed through Document
        >>> doc = Document("contract.docx")
        >>> doc.insert_hyperlink(
        ...     url="https://example.com",
        ...     text="Click here",
        ...     after="For more information,"
        ... )
    """

    def __init__(self, document: Document) -> None:
        """Initialize HyperlinkOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    # ==================== Insert Hyperlinks ====================

    def insert_hyperlink(
        self,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
        after: str | None = None,
        before: str | None = None,
        scope: str | dict | Any | None = None,
        tooltip: str | None = None,
        track: bool = False,
        author: str | None = None,
    ) -> str | None:
        """Insert a hyperlink at a specific location in the document body.

        Supports both external hyperlinks (URLs) and internal hyperlinks
        (bookmarks). Use the `url` parameter for external links and the
        `anchor` parameter for internal links.

        Args:
            url: External URL to link to (mutually exclusive with anchor)
            anchor: Internal bookmark name to link to (mutually exclusive with url)
            text: The display text for the hyperlink
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            scope: Optional scope to limit search (paragraph ref, heading, etc.)
            tooltip: Optional tooltip text shown on hover
            track: If True, wrap insertion in tracked change markup
            author: Optional author override for tracked changes

        Returns:
            Relationship ID (rId) for external links, None for internal links

        Raises:
            ValueError: If both url and anchor specified, or neither specified
            ValueError: If both after and before specified, or neither specified
            TextNotFoundError: If anchor text not found
            AmbiguousTextError: If anchor text found multiple times

        Example:
            >>> # External hyperlink
            >>> doc.insert_hyperlink(
            ...     url="https://www.law.cornell.edu/uscode/text/28/1782",
            ...     text="28 U.S.C. section 1782",
            ...     after="discovery statute"
            ... )
            'rId15'

            >>> # Internal hyperlink to bookmark
            >>> doc.insert_hyperlink(
            ...     anchor="DefinitionsSection",
            ...     text="See Definitions",
            ...     after="as defined below"
            ... )
            None
        """
        # Validate url/anchor parameters
        if url is not None and anchor is not None:
            raise ValueError("Cannot specify both 'url' and 'anchor' parameters")
        if url is None and anchor is None:
            raise ValueError("Must specify either 'url' or 'anchor' parameter")

        # Validate after/before parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        # Ensure we have a valid package
        if not self._document._is_zip or not self._document._temp_dir:
            raise ValueError("Cannot add hyperlinks to non-ZIP documents")

        # Find location for hyperlink insertion
        anchor_text = after if after is not None else before
        insert_after = after is not None

        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(
            anchor_text,
            paragraphs,  # type: ignore[arg-type]
        )

        if not matches:
            scope_str = str(scope) if scope is not None and not isinstance(scope, str) else scope
            raise TextNotFoundError(anchor_text, scope_str)  # type: ignore[arg-type]

        if len(matches) > 1:
            raise AmbiguousTextError(anchor_text, matches)  # type: ignore[arg-type]

        match = matches[0]

        # Ensure Hyperlink style exists
        self._ensure_hyperlink_style()

        # Handle external links (url) vs internal links (anchor)
        r_id: str | None = None
        if url is not None:
            # External link: add hyperlink relationship
            package = self._document._package
            if not package:
                raise ValueError("Cannot add hyperlinks: package not available")

            rel_mgr = RelationshipManager(package, "word/document.xml")
            r_id = rel_mgr.add_unique_relationship(
                RelationshipTypes.HYPERLINK,
                url,
                target_mode="External",
            )
            rel_mgr.save()

            # Create hyperlink element with r:id
            hyperlink_elem = self._create_hyperlink_element(
                text=text,
                r_id=r_id,
                anchor=None,
                tooltip=tooltip,
            )
        else:
            # Internal link: no relationship needed, just w:anchor attribute
            # Note: anchor is not None here due to earlier validation

            # Validate that the bookmark exists (warn if not, but allow the link)
            bookmark_registry = BookmarkRegistry.from_xml(self._document.xml_root)
            if not bookmark_registry.get_bookmark(anchor):
                warnings.warn(
                    f"Bookmark '{anchor}' does not exist. Internal hyperlink will be broken.",
                    UserWarning,
                    stacklevel=2,
                )

            hyperlink_elem = self._create_hyperlink_element(
                text=text,
                r_id=None,
                anchor=anchor,
                tooltip=tooltip,
            )

        # Insert the hyperlink at the match location
        if insert_after:
            self._insert_after_match(match, hyperlink_elem)
        else:
            self._insert_before_match(match, hyperlink_elem)

        return r_id

    def insert_hyperlink_in_header(
        self,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
        after: str | None = None,
        before: str | None = None,
        header_type: str = "default",
        track: bool = False,
    ) -> str | None:
        """Insert a hyperlink in a header.

        Args:
            url: External URL (mutually exclusive with anchor)
            anchor: Internal bookmark name (mutually exclusive with url)
            text: Display text for the hyperlink
            after: Text to insert after
            before: Text to insert before
            header_type: "default", "first", or "even"
            track: If True, track the insertion

        Returns:
            Relationship ID for external links, None for internal

        Raises:
            ValueError: If both url and anchor specified, or neither specified
            ValueError: If invalid header_type specified
            TextNotFoundError: If anchor text not found in header
        """
        raise NotImplementedError("insert_hyperlink_in_header not yet implemented")

    def insert_hyperlink_in_footer(
        self,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
        after: str | None = None,
        before: str | None = None,
        footer_type: str = "default",
        track: bool = False,
    ) -> str | None:
        """Insert a hyperlink in a footer.

        Args:
            url: External URL (mutually exclusive with anchor)
            anchor: Internal bookmark name (mutually exclusive with url)
            text: Display text for the hyperlink
            after: Text to insert after
            before: Text to insert before
            footer_type: "default", "first", or "even"
            track: If True, track the insertion

        Returns:
            Relationship ID for external links, None for internal

        Raises:
            ValueError: If both url and anchor specified, or neither specified
            ValueError: If invalid footer_type specified
            TextNotFoundError: If anchor text not found in footer
        """
        raise NotImplementedError("insert_hyperlink_in_footer not yet implemented")

    def insert_hyperlink_in_footnote(
        self,
        note_id: str | int,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
        after: str | None = None,
        before: str | None = None,
        track: bool = False,
    ) -> str | None:
        """Insert a hyperlink inside an existing footnote.

        Args:
            note_id: The footnote ID to edit
            url: External URL (mutually exclusive with anchor)
            anchor: Internal bookmark name (mutually exclusive with url)
            text: Display text for the hyperlink
            after: Text to insert after within the footnote
            before: Text to insert before within the footnote
            track: If True, track the insertion

        Returns:
            Relationship ID for external links, None for internal

        Raises:
            NoteNotFoundError: If footnote not found
            ValueError: If both url and anchor specified, or neither specified
            TextNotFoundError: If anchor text not found in footnote
        """
        raise NotImplementedError("insert_hyperlink_in_footnote not yet implemented")

    def insert_hyperlink_in_endnote(
        self,
        note_id: str | int,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
        after: str | None = None,
        before: str | None = None,
        track: bool = False,
    ) -> str | None:
        """Insert a hyperlink inside an existing endnote.

        Args:
            note_id: The endnote ID to edit
            url: External URL (mutually exclusive with anchor)
            anchor: Internal bookmark name (mutually exclusive with url)
            text: Display text for the hyperlink
            after: Text to insert after within the endnote
            before: Text to insert before within the endnote
            track: If True, track the insertion

        Returns:
            Relationship ID for external links, None for internal

        Raises:
            NoteNotFoundError: If endnote not found
            ValueError: If both url and anchor specified, or neither specified
            TextNotFoundError: If anchor text not found in endnote
        """
        raise NotImplementedError("insert_hyperlink_in_endnote not yet implemented")

    # ==================== Read Hyperlinks ====================

    def get_all_hyperlinks(self) -> list[HyperlinkInfo]:
        """Get all hyperlinks in the document.

        Returns hyperlinks from all locations: body, headers, footers,
        footnotes, and endnotes.

        Returns:
            List of HyperlinkInfo objects with link details

        Example:
            >>> for link in doc.hyperlinks:
            ...     print(f"{link.text} -> {link.target}")
            ...     if link.is_external:
            ...         print(f"  External URL")
            ...     else:
            ...         print(f"  Internal bookmark")
        """
        raise NotImplementedError("get_all_hyperlinks not yet implemented")

    def get_hyperlink(self, ref: str) -> HyperlinkInfo | None:
        """Get a specific hyperlink by its ref.

        Args:
            ref: Hyperlink ref (e.g., "lnk:5")

        Returns:
            HyperlinkInfo if found, None otherwise

        Example:
            >>> link = doc.get_hyperlink("lnk:5")
            >>> if link:
            ...     print(f"Found: {link.text} -> {link.target}")
        """
        raise NotImplementedError("get_hyperlink not yet implemented")

    def get_hyperlinks_by_url(self, url_pattern: str) -> list[HyperlinkInfo]:
        """Find hyperlinks matching a URL pattern.

        Searches for hyperlinks whose target URL contains the given pattern.
        For internal links, searches the bookmark name instead.

        Args:
            url_pattern: String to search for in hyperlink targets

        Returns:
            List of matching HyperlinkInfo objects

        Example:
            >>> # Find all links to a specific domain
            >>> links = doc.get_hyperlinks_by_url("law.cornell.edu")
            >>> for link in links:
            ...     print(f"{link.text}: {link.target}")
        """
        raise NotImplementedError("get_hyperlinks_by_url not yet implemented")

    # ==================== Edit Hyperlinks ====================

    def edit_hyperlink_url(self, ref: str, new_url: str) -> None:
        """Change the URL of an external hyperlink.

        Updates the relationship target for the specified hyperlink.
        Only works for external hyperlinks (not internal bookmarks).

        Args:
            ref: Hyperlink ref (e.g., "lnk:5") or relationship ID (e.g., "rId5")
            new_url: The new URL to link to

        Raises:
            ValueError: If hyperlink not found or is an internal link
            ValueError: If new_url is empty

        Example:
            >>> doc.edit_hyperlink_url("lnk:5", "https://new-url.com")
        """
        # Validate new_url
        if not new_url or not new_url.strip():
            raise ValueError("new_url cannot be empty")

        # Ensure we have a valid package
        if not self._document._is_zip or not self._document._temp_dir:
            raise ValueError("Cannot edit hyperlinks in non-ZIP documents")

        package = self._document._package
        if not package:
            raise ValueError("Cannot edit hyperlinks: package not available")

        # Find the hyperlink element and its relationship ID
        hyperlink_elem, r_id = self._find_hyperlink_by_ref(ref)

        if hyperlink_elem is None:
            raise ValueError(f"Hyperlink not found: {ref}")

        if r_id is None:
            raise ValueError(
                f"Cannot edit URL of internal hyperlink '{ref}' - "
                "use edit_hyperlink_anchor() for internal links"
            )

        # Update the relationship target
        rel_mgr = RelationshipManager(package, "word/document.xml")
        updated = rel_mgr.update_relationship_target(r_id, new_url)

        if not updated:
            raise ValueError(f"Relationship '{r_id}' not found for hyperlink '{ref}'")

        rel_mgr.save()

    def edit_hyperlink_text(
        self,
        ref: str,
        new_text: str,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Change the display text of a hyperlink.

        Replaces the visible text of the hyperlink while keeping the
        same target URL or bookmark.

        Args:
            ref: Hyperlink ref (e.g., "lnk:5")
            new_text: The new display text
            track: If True, show text change as tracked change
            author: Optional author for tracked change

        Raises:
            ValueError: If hyperlink not found
            ValueError: If new_text is empty

        Example:
            >>> doc.edit_hyperlink_text("lnk:5", "Updated link text")
        """
        raise NotImplementedError("edit_hyperlink_text not yet implemented")

    def edit_hyperlink_anchor(self, ref: str, new_anchor: str) -> None:
        """Change the target bookmark of an internal hyperlink.

        Only works for internal hyperlinks (not external URLs).

        Args:
            ref: Hyperlink ref (e.g., "lnk:5")
            new_anchor: The new bookmark name to link to

        Raises:
            ValueError: If hyperlink not found or is an external link
            ValueError: If new_anchor is empty

        Example:
            >>> doc.edit_hyperlink_anchor("lnk:3", "NewBookmarkName")
        """
        raise NotImplementedError("edit_hyperlink_anchor not yet implemented")

    # ==================== Remove Hyperlinks ====================

    def remove_hyperlink(
        self,
        ref: str,
        keep_text: bool = True,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Remove a hyperlink from the document.

        Can either keep the display text (unlinking it) or remove both
        the link and the text entirely.

        Args:
            ref: Hyperlink ref (e.g., "lnk:5")
            keep_text: If True (default), keep the display text without the link.
                       If False, remove both the link and the text.
            track: If True and keep_text=False, show text removal as tracked deletion
            author: Optional author for tracked change

        Raises:
            ValueError: If hyperlink not found

        Example:
            >>> # Keep text, just remove the link
            >>> doc.remove_hyperlink("lnk:5")

            >>> # Remove link and text entirely
            >>> doc.remove_hyperlink("lnk:5", keep_text=False)

            >>> # Remove with tracking
            >>> doc.remove_hyperlink("lnk:5", keep_text=False, track=True)
        """
        raise NotImplementedError("remove_hyperlink not yet implemented")

    # ==================== Internal Helper Methods ====================

    def _create_hyperlink_element(
        self,
        text: str,
        r_id: str | None = None,
        anchor: str | None = None,
        tooltip: str | None = None,
    ) -> Any:
        """Create a w:hyperlink XML element.

        Creates the proper OOXML structure for a hyperlink with the
        Hyperlink character style applied.

        Args:
            text: Display text for the hyperlink
            r_id: Relationship ID for external links (mutually exclusive with anchor)
            anchor: Bookmark name for internal links (mutually exclusive with r_id)
            tooltip: Optional tooltip text

        Returns:
            lxml Element representing the w:hyperlink

        Raises:
            ValueError: If both r_id and anchor specified, or neither specified
        """
        # Validate r_id/anchor
        if r_id is not None and anchor is not None:
            raise ValueError("Cannot specify both 'r_id' and 'anchor' for hyperlink element")
        if r_id is None and anchor is None:
            raise ValueError("Must specify either 'r_id' or 'anchor' for hyperlink element")

        # Create w:hyperlink element with proper namespace
        nsmap = {
            "w": WORD_NAMESPACE,
            "r": OFFICE_RELATIONSHIPS_NAMESPACE,
        }
        hyperlink = etree.Element(f"{{{WORD_NAMESPACE}}}hyperlink", nsmap=nsmap)

        # Set r:id for external links or w:anchor for internal links
        if r_id is not None:
            hyperlink.set(f"{{{OFFICE_RELATIONSHIPS_NAMESPACE}}}id", r_id)
        else:
            hyperlink.set(f"{{{WORD_NAMESPACE}}}anchor", anchor)  # type: ignore[arg-type]

        # Set optional tooltip
        if tooltip:
            hyperlink.set(f"{{{WORD_NAMESPACE}}}tooltip", tooltip)

        # Create inner run with Hyperlink style
        run = etree.SubElement(hyperlink, f"{{{WORD_NAMESPACE}}}r")

        # Add run properties with Hyperlink character style
        rpr = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}rPr")
        rstyle = etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}rStyle")
        rstyle.set(f"{{{WORD_NAMESPACE}}}val", "Hyperlink")

        # Add display text
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        # Preserve whitespace if text has leading/trailing spaces
        if text and (text[0].isspace() or text[-1].isspace()):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text

        return hyperlink

    def _ensure_hyperlink_style(self) -> None:
        """Ensure the Hyperlink character style exists in the document.

        Creates the standard Hyperlink style if it doesn't already exist.
        This style provides the typical blue underlined text appearance
        (blue color #0563C1 with single underline).

        Also ensures the styles.xml relationship and content type exist.
        """
        ensure_standard_styles(self._document.styles, "Hyperlink")
        # Ensure styles.xml is properly registered if it was created
        self._ensure_styles_relationship()
        self._ensure_styles_content_type()
        self._document.styles.save()

    def _ensure_styles_relationship(self) -> None:
        """Ensure styles.xml relationship exists in document.xml.rels."""
        package = self._document._package
        if not package:
            return

        rel_mgr = RelationshipManager(package, "word/document.xml")
        rel_mgr.add_relationship(RelationshipTypes.STYLES, "styles.xml")
        rel_mgr.save()

    def _ensure_styles_content_type(self) -> None:
        """Ensure styles.xml content type exists in [Content_Types].xml."""
        from ..content_types import ContentTypeManager, ContentTypes

        package = self._document._package
        if not package:
            return

        ct_mgr = ContentTypeManager(package)
        ct_mgr.add_override("/word/styles.xml", ContentTypes.STYLES)
        ct_mgr.save()

    def _get_next_hyperlink_index(self) -> int:
        """Get the next available hyperlink index for generating refs.

        Returns:
            Integer index for generating unique hyperlink refs
        """
        raise NotImplementedError("_get_next_hyperlink_index not yet implemented")

    def _insert_after_match(self, match: TextSpan, element: Any) -> None:
        """Insert an XML element after a matched text span.

        The element is inserted as a sibling after the last run of the match,
        within the same paragraph.

        Args:
            match: TextSpan object representing where to insert
            element: The lxml Element to insert
        """
        paragraph = match.paragraph
        end_run = match.runs[match.end_run_index]

        # Find position of end_run in the paragraph
        run_index = list(paragraph).index(end_run)

        # Insert element after the end_run
        paragraph.insert(run_index + 1, element)

    def _insert_before_match(self, match: TextSpan, element: Any) -> None:
        """Insert an XML element before a matched text span.

        The element is inserted as a sibling before the first run of the match,
        within the same paragraph.

        Args:
            match: TextSpan object representing where to insert
            element: The lxml Element to insert
        """
        paragraph = match.paragraph
        start_run = match.runs[match.start_run_index]

        # Find position of start_run in the paragraph
        run_index = list(paragraph).index(start_run)

        # Insert element before the start_run
        paragraph.insert(run_index, element)
