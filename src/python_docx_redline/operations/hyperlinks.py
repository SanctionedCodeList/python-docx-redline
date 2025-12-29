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
from ..errors import AmbiguousTextError, NoteNotFoundError, TextNotFoundError
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

        # Validate header_type
        valid_types = {"default", "first", "even"}
        if header_type not in valid_types:
            raise ValueError(
                f"Invalid header_type '{header_type}'. Must be one of: {', '.join(valid_types)}"
            )

        # Ensure we have a valid package
        if not self._document._is_zip or not self._document._temp_dir:
            raise ValueError("Cannot add hyperlinks to non-ZIP documents")

        # Get the header
        header = self._document._header_footer_ops._get_header_by_type(header_type)
        if header is None:
            raise ValueError(f"No header of type '{header_type}' found in document")

        # Find location for hyperlink insertion
        anchor_text = after if after is not None else before
        insert_after = after is not None

        paragraphs = list(header.element.iter(f"{{{WORD_NAMESPACE}}}p"))

        matches = self._document._text_search.find_text(
            anchor_text,
            paragraphs,  # type: ignore[arg-type]
        )

        if not matches:
            raise TextNotFoundError(anchor_text)  # type: ignore[arg-type]

        if len(matches) > 1:
            raise AmbiguousTextError(anchor_text, matches)  # type: ignore[arg-type]

        match = matches[0]

        # Ensure Hyperlink style exists
        self._ensure_hyperlink_style()

        # Handle external links (url) vs internal links (anchor)
        r_id: str | None = None
        if url is not None:
            # External link: add hyperlink relationship to the header's .rels file
            package = self._document._package
            if not package:
                raise ValueError("Cannot add hyperlinks: package not available")

            # The part name for the header is "word/header1.xml" (or similar)
            # file_path is the relative path from word/ (e.g., "header1.xml")
            header_part_name = f"word/{header.file_path}"

            rel_mgr = RelationshipManager(package, header_part_name)
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
                tooltip=None,
            )
        else:
            # Internal link: no relationship needed, just w:anchor attribute
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
                tooltip=None,
            )

        # Insert the hyperlink at the match location
        if insert_after:
            self._insert_after_match(match, hyperlink_elem)
        else:
            self._insert_before_match(match, hyperlink_elem)

        # Save the modified header XML
        self._document._header_footer_ops._save_header_footer_xml(header.file_path, header.element)

        return r_id

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

        # Validate footer_type
        valid_types = {"default", "first", "even"}
        if footer_type not in valid_types:
            raise ValueError(
                f"Invalid footer_type '{footer_type}'. Must be one of: {', '.join(valid_types)}"
            )

        # Ensure we have a valid package
        if not self._document._is_zip or not self._document._temp_dir:
            raise ValueError("Cannot add hyperlinks to non-ZIP documents")

        # Get the footer
        footer = self._document._header_footer_ops._get_footer_by_type(footer_type)
        if footer is None:
            raise ValueError(f"No footer of type '{footer_type}' found in document")

        # Find location for hyperlink insertion
        anchor_text = after if after is not None else before
        insert_after = after is not None

        paragraphs = list(footer.element.iter(f"{{{WORD_NAMESPACE}}}p"))

        matches = self._document._text_search.find_text(
            anchor_text,
            paragraphs,  # type: ignore[arg-type]
        )

        if not matches:
            raise TextNotFoundError(anchor_text)  # type: ignore[arg-type]

        if len(matches) > 1:
            raise AmbiguousTextError(anchor_text, matches)  # type: ignore[arg-type]

        match = matches[0]

        # Ensure Hyperlink style exists
        self._ensure_hyperlink_style()

        # Handle external links (url) vs internal links (anchor)
        r_id: str | None = None
        if url is not None:
            # External link: add hyperlink relationship to the footer's .rels file
            package = self._document._package
            if not package:
                raise ValueError("Cannot add hyperlinks: package not available")

            # The part name for the footer is "word/footer1.xml" (or similar)
            # file_path is the relative path from word/ (e.g., "footer1.xml")
            footer_part_name = f"word/{footer.file_path}"

            rel_mgr = RelationshipManager(package, footer_part_name)
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
                tooltip=None,
            )
        else:
            # Internal link: no relationship needed, just w:anchor attribute
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
                tooltip=None,
            )

        # Insert the hyperlink at the match location
        if insert_after:
            self._insert_after_match(match, hyperlink_elem)
        else:
            self._insert_before_match(match, hyperlink_elem)

        # Save the modified footer XML
        self._document._header_footer_ops._save_header_footer_xml(footer.file_path, footer.element)

        return r_id

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

        Example:
            >>> doc.insert_hyperlink_in_footnote(
            ...     note_id=3,
            ...     url='https://law.cornell.edu/uscode/text/28/1782',
            ...     text='28 U.S.C. section 1782',
            ...     after='See'
            ... )
        """
        return self._insert_hyperlink_in_note(
            note_type="footnote",
            note_id=note_id,
            url=url,
            anchor=anchor,
            text=text,
            after=after,
            before=before,
            track=track,
        )

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

        Example:
            >>> doc.insert_hyperlink_in_endnote(
            ...     note_id=1,
            ...     url='https://example.com/reference',
            ...     text='online reference',
            ...     after='see the'
            ... )
        """
        return self._insert_hyperlink_in_note(
            note_type="endnote",
            note_id=note_id,
            url=url,
            anchor=anchor,
            text=text,
            after=after,
            before=before,
            track=track,
        )

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
        from ..accessibility.types import LinkType

        # Build relationships map for resolving external hyperlink URLs
        relationships = self._get_hyperlink_relationships()

        # Use BookmarkRegistry to extract hyperlinks from the document
        registry = BookmarkRegistry.from_xml(
            self._document.xml_root,
            relationships=relationships,
        )

        # Convert from accessibility HyperlinkInfo to operations HyperlinkInfo
        result: list[HyperlinkInfo] = []
        for link in registry.hyperlinks:
            is_external = link.link_type == LinkType.EXTERNAL
            result.append(
                HyperlinkInfo(
                    ref=link.ref,
                    text=link.text,
                    target=link.target,
                    is_external=is_external,
                    tooltip=None,  # Not tracked by BookmarkRegistry
                    location=link.from_location,
                    r_id=link.relationship_id,
                )
            )

        return result

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
        for link in self.get_all_hyperlinks():
            if link.ref == ref:
                return link
        return None

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
        return [link for link in self.get_all_hyperlinks() if url_pattern in link.target]

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
        # Validate new_text
        if not new_text:
            raise ValueError("new_text cannot be empty")

        # Find the hyperlink element
        hyperlink_elem, _ = self._find_hyperlink_by_ref(ref)
        if hyperlink_elem is None:
            raise ValueError(f"Hyperlink not found: {ref}")

        # Get the current text content
        old_text = self._get_hyperlink_text(hyperlink_elem)

        # Find all runs inside the hyperlink
        runs = list(hyperlink_elem.findall(f"{{{WORD_NAMESPACE}}}r"))

        if not runs:
            raise ValueError(f"Hyperlink has no runs to edit: {ref}")

        if track:
            # Create tracked deletion for old text and insertion for new text
            # Replace all runs with deletion + insertion
            deletion_xml = self._document._xml_generator.create_deletion(old_text, author)
            insertion_xml = self._document._xml_generator.create_insertion(new_text, author)

            # Parse both elements
            deletion_elements = self._parse_xml_elements(deletion_xml)
            insertion_elements = self._parse_xml_elements(insertion_xml)

            # Get first run position
            first_run_index = list(hyperlink_elem).index(runs[0])

            # Remove all existing runs
            for run in runs:
                hyperlink_elem.remove(run)

            # Insert deletion and insertion elements at the first run position
            insert_pos = first_run_index
            for elem in deletion_elements:
                hyperlink_elem.insert(insert_pos, elem)
                insert_pos += 1
            for elem in insertion_elements:
                hyperlink_elem.insert(insert_pos, elem)
                insert_pos += 1
        else:
            # Untracked edit: simply replace text content in runs
            # Strategy: clear all runs except the first, then set new text in the first
            # while preserving the Hyperlink style

            # Keep the first run and remove the rest
            first_run = runs[0]
            for run in runs[1:]:
                hyperlink_elem.remove(run)

            # Find or create the text element in the first run
            t_elem = first_run.find(f"{{{WORD_NAMESPACE}}}t")
            if t_elem is None:
                t_elem = etree.SubElement(first_run, f"{{{WORD_NAMESPACE}}}t")

            # Set the new text
            t_elem.text = new_text

            # Handle xml:space for leading/trailing whitespace
            if new_text and (new_text[0].isspace() or new_text[-1].isspace()):
                t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            else:
                # Remove the attribute if it exists and is not needed
                if "{http://www.w3.org/XML/1998/namespace}space" in t_elem.attrib:
                    del t_elem.attrib["{http://www.w3.org/XML/1998/namespace}space"]

            # Ensure the Hyperlink style is applied
            rpr = first_run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if rpr is None:
                rpr = etree.Element(f"{{{WORD_NAMESPACE}}}rPr")
                first_run.insert(0, rpr)

            rstyle = rpr.find(f"{{{WORD_NAMESPACE}}}rStyle")
            if rstyle is None:
                rstyle = etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}rStyle")
            rstyle.set(f"{{{WORD_NAMESPACE}}}val", "Hyperlink")

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
        # Validate new_anchor
        if not new_anchor or not new_anchor.strip():
            raise ValueError("new_anchor cannot be empty")

        # Find the hyperlink element
        hyperlink_elem, r_id = self._find_hyperlink_by_ref(ref)

        if hyperlink_elem is None:
            raise ValueError(f"Hyperlink not found: {ref}")

        # Check if this is an internal link (has w:anchor, no r:id)
        if r_id is not None:
            raise ValueError(
                f"Cannot edit anchor of external hyperlink '{ref}' - "
                "use edit_hyperlink_url() for external links"
            )

        # Verify the hyperlink has an existing anchor (is truly internal)
        ns_w = f"{{{WORD_NAMESPACE}}}"
        current_anchor = hyperlink_elem.get(f"{ns_w}anchor")
        if current_anchor is None:
            raise ValueError(
                f"Hyperlink '{ref}' has no anchor attribute - "
                "cannot determine if it is an internal link"
            )

        # Warn if the new bookmark doesn't exist
        bookmark_registry = BookmarkRegistry.from_xml(self._document.xml_root)
        if not bookmark_registry.get_bookmark(new_anchor):
            warnings.warn(
                f"Bookmark '{new_anchor}' does not exist. Internal hyperlink will be broken.",
                UserWarning,
                stacklevel=2,
            )

        # Update the anchor attribute
        hyperlink_elem.set(f"{ns_w}anchor", new_anchor)

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
        # Find the hyperlink element
        hyperlink_elem, _ = self._find_hyperlink_by_ref(ref)
        if hyperlink_elem is None:
            raise ValueError(f"Hyperlink not found: {ref}")

        parent = hyperlink_elem.getparent()
        if parent is None:
            raise ValueError(f"Hyperlink has no parent element: {ref}")

        # Get the position of the hyperlink in its parent
        hyperlink_index = list(parent).index(hyperlink_elem)

        if keep_text:
            # Extract runs from inside the hyperlink and insert them where the hyperlink was
            inner_runs = list(hyperlink_elem.findall(f"{{{WORD_NAMESPACE}}}r"))

            # Remove the hyperlink element
            parent.remove(hyperlink_elem)

            # Insert the inner runs at the hyperlink's former position
            for i, run in enumerate(inner_runs):
                # Remove the Hyperlink style from the run if present
                self._remove_hyperlink_style_from_run(run)
                parent.insert(hyperlink_index + i, run)
        else:
            # Remove the entire hyperlink including its text
            if track:
                # Extract the text content and create a tracked deletion
                link_text = self._get_hyperlink_text(hyperlink_elem)
                if link_text:
                    # Create tracked deletion element
                    deletion_xml = self._document._xml_generator.create_deletion(link_text, author)
                    # Parse the deletion XML
                    elements = self._parse_xml_elements(deletion_xml)
                    deletion_element = elements[0]

                    # Replace hyperlink with deletion element
                    parent.remove(hyperlink_elem)
                    parent.insert(hyperlink_index, deletion_element)
                else:
                    # No text, just remove the hyperlink
                    parent.remove(hyperlink_elem)
            else:
                # Untracked removal - just remove the hyperlink element
                parent.remove(hyperlink_elem)

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

    def _find_hyperlink_by_ref(self, ref: str) -> tuple[Any | None, str | None]:
        """Find a hyperlink element by its ref or relationship ID.

        Args:
            ref: Hyperlink ref (e.g., "lnk:5") or relationship ID (e.g., "rId5")

        Returns:
            Tuple of (hyperlink_element, relationship_id).
            relationship_id is None for internal links.
            Both are None if not found.

        Raises:
            ValueError: If ref format is invalid
        """
        ns_r = f"{{{OFFICE_RELATIONSHIPS_NAMESPACE}}}"

        # If ref is already a relationship ID (rId...), search directly
        if ref.startswith("rId"):
            for hyperlink in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}hyperlink"):
                if hyperlink.get(f"{ns_r}id") == ref:
                    return hyperlink, ref
            return None, None

        # Otherwise, parse the lnk:N format
        if not ref.startswith("lnk:"):
            raise ValueError(
                f"Invalid hyperlink ref format: '{ref}'. "
                "Expected 'lnk:N' (e.g., 'lnk:5') or 'rIdN' (e.g., 'rId5')"
            )

        try:
            target_index = int(ref[4:])
        except ValueError:
            raise ValueError(f"Invalid hyperlink ref: '{ref}'. Expected 'lnk:N' with integer N")

        # Iterate through hyperlinks in document order to find by index
        body = self._document.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return None, None

        link_index = 0
        for paragraph in body.findall(f"./{{{WORD_NAMESPACE}}}p"):
            for hyperlink in paragraph.findall(f".//{{{WORD_NAMESPACE}}}hyperlink"):
                if link_index == target_index:
                    # Found the hyperlink - get its relationship ID (if external)
                    r_id = hyperlink.get(f"{ns_r}id")
                    return hyperlink, r_id
                link_index += 1

        return None, None

    def _get_next_hyperlink_index(self) -> int:
        """Get the next available hyperlink index for generating refs.

        Returns:
            Integer index for generating unique hyperlink refs
        """
        # Count existing hyperlinks in the document body
        body = self._document.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return 0

        count = 0
        for _ in body.iter(f"{{{WORD_NAMESPACE}}}hyperlink"):
            count += 1

        return count

    def _get_hyperlink_relationships(self) -> dict[str, str]:
        """Build a dictionary mapping relationship IDs to target URLs.

        This reads the document.xml.rels file and extracts all hyperlink
        relationships, mapping rId values to their target URLs.

        Returns:
            Dictionary mapping relationship IDs (e.g., "rId5") to URLs
        """
        relationships: dict[str, str] = {}

        # Need a valid package to read relationships
        if not self._document._is_zip or not self._document._package:
            return relationships

        package = self._document._package
        rel_mgr = RelationshipManager(package, "word/document.xml")

        # Load the relationships file
        rel_mgr._ensure_loaded()
        if rel_mgr._root is None:
            return relationships

        # Extract all hyperlink relationships
        for rel in rel_mgr._root:
            rel_type = rel.get("Type", "")
            if rel_type == RelationshipTypes.HYPERLINK:
                rel_id = rel.get("Id", "")
                target = rel.get("Target", "")
                if rel_id and target:
                    relationships[rel_id] = target

        return relationships

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

    def _get_hyperlink_text(self, hyperlink_elem: etree._Element) -> str:
        """Extract the display text from a hyperlink element.

        Args:
            hyperlink_elem: The w:hyperlink element

        Returns:
            The text content of the hyperlink
        """
        text_parts = []
        for t_elem in hyperlink_elem.iter(f"{{{WORD_NAMESPACE}}}t"):
            if t_elem.text:
                text_parts.append(t_elem.text)
        return "".join(text_parts)

    def _remove_hyperlink_style_from_run(self, run: etree._Element) -> None:
        """Remove the Hyperlink character style from a run's properties.

        This removes the blue underline styling when unlinking a hyperlink
        while keeping the text.

        Args:
            run: The w:r element to modify
        """
        run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
        if run_props is not None:
            rstyle = run_props.find(f"{{{WORD_NAMESPACE}}}rStyle")
            if rstyle is not None:
                style_val = rstyle.get(f"{{{WORD_NAMESPACE}}}val")
                if style_val == "Hyperlink":
                    run_props.remove(rstyle)
                    # If rPr is now empty, remove it too
                    if len(run_props) == 0:
                        run.remove(run_props)

    def _parse_xml_elements(self, xml_content: str) -> list:
        """Parse XML content into lxml elements with proper namespaces.

        Args:
            xml_content: The XML string(s) to parse (can be multiple fragments)

        Returns:
            List of parsed lxml Elements
        """
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {xml_content}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        return list(root)

    def _insert_hyperlink_in_note(
        self,
        note_type: str,
        note_id: str | int,
        url: str | None,
        anchor: str | None,
        text: str,
        after: str | None,
        before: str | None,
        track: bool,
    ) -> str | None:
        """Internal implementation for inserting hyperlinks in footnotes/endnotes.

        This method handles the common logic for both footnote and endnote hyperlinks.
        Footnotes use word/footnotes.xml with rels in word/_rels/footnotes.xml.rels.
        Endnotes use word/endnotes.xml with rels in word/_rels/endnotes.xml.rels.

        Args:
            note_type: Either "footnote" or "endnote"
            note_id: The note ID to edit
            url: External URL (mutually exclusive with anchor)
            anchor: Internal bookmark name (mutually exclusive with url)
            text: Display text for the hyperlink
            after: Text to insert after within the note
            before: Text to insert before within the note
            track: If True, track the insertion (not yet implemented)

        Returns:
            Relationship ID for external links, None for internal

        Raises:
            NoteNotFoundError: If note not found
            ValueError: If parameters are invalid
            TextNotFoundError: If anchor text not found in note
            AmbiguousTextError: If anchor text found multiple times
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

        temp_dir = self._document._temp_dir
        package = self._document._package
        if not package:
            raise ValueError("Cannot add hyperlinks: package not available")

        # Determine file paths based on note type
        if note_type == "footnote":
            xml_path = temp_dir / "word" / "footnotes.xml"
            rels_part = "word/footnotes.xml"
            tag_name = "footnote"
        else:
            xml_path = temp_dir / "word" / "endnotes.xml"
            rels_part = "word/endnotes.xml"
            tag_name = "endnote"

        # Find the note element
        note_id_str = str(note_id)

        if not xml_path.exists():
            raise NoteNotFoundError(note_type, note_id_str, [])

        tree = etree.parse(str(xml_path))
        root = tree.getroot()

        # Find the note element by ID
        note_elem = None
        available_ids: list[str] = []
        for elem in root.findall(f"{{{WORD_NAMESPACE}}}{tag_name}"):
            elem_id = elem.get(f"{{{WORD_NAMESPACE}}}id")
            if elem.get(f"{{{WORD_NAMESPACE}}}type") is None:  # Skip separators
                if elem_id:
                    available_ids.append(elem_id)
                if elem_id == note_id_str:
                    note_elem = elem

        if note_elem is None:
            raise NoteNotFoundError(note_type, note_id_str, available_ids)

        # Get paragraphs from the note for text search
        paragraphs = list(note_elem.findall(f"{{{WORD_NAMESPACE}}}p"))
        if not paragraphs:
            anchor_text = after if after is not None else before
            raise TextNotFoundError(anchor_text)  # type: ignore[arg-type]

        # Find location for hyperlink insertion
        anchor_text = after if after is not None else before
        insert_after = after is not None

        matches = self._document._text_search.find_text(
            anchor_text,
            paragraphs,  # type: ignore[arg-type]
        )

        if not matches:
            raise TextNotFoundError(anchor_text)  # type: ignore[arg-type]

        if len(matches) > 1:
            raise AmbiguousTextError(anchor_text, matches)  # type: ignore[arg-type]

        match = matches[0]

        # Ensure Hyperlink style exists
        self._ensure_hyperlink_style()

        # Handle external links (url) vs internal links (anchor)
        r_id: str | None = None
        if url is not None:
            # External link: add hyperlink relationship to the notes rels file
            rel_mgr = RelationshipManager(package, rels_part)
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
                tooltip=None,
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
                    stacklevel=3,
                )

            hyperlink_elem = self._create_hyperlink_element(
                text=text,
                r_id=None,
                anchor=anchor,
                tooltip=None,
            )

        # Insert the hyperlink at the match location
        if insert_after:
            self._insert_after_match(match, hyperlink_elem)
        else:
            self._insert_before_match(match, hyperlink_elem)

        # Save the modified XML
        tree.write(
            str(xml_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

        return r_id
