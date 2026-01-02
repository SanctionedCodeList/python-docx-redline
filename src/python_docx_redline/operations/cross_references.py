"""
CrossReferenceOperations class for handling cross-reference manipulation.

This module provides a dedicated class for cross-reference operations, including:
- Inserting cross-references to bookmarks, headings, figures, tables, and notes
- Field code generation for REF, PAGEREF, and NOTEREF fields
- Switch mapping from display options to field switches
- Managing dirty flags for Word field calculation
- Creating and managing bookmarks

Cross-references in Word are implemented as field codes that Word calculates
when the document is opened. This module generates the correct OOXML structure
and marks fields dirty so Word updates them.
"""

from __future__ import annotations

import logging
import random
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from lxml import etree

from ..constants import WORD_NAMESPACE, XML_NAMESPACE, w
from ..errors import (
    AmbiguousTextError,
    BookmarkAlreadyExistsError,
    InvalidBookmarkNameError,
    TextNotFoundError,
)
from ..scope import ScopeEvaluator

if TYPE_CHECKING:
    from ..document import Document

logger = logging.getLogger(__name__)


# =============================================================================
# Data Models
# =============================================================================


@dataclass
class CrossReference:
    """Information about a cross-reference in the document.

    Represents a REF, PAGEREF, or NOTEREF field that references
    another location in the document.

    Attributes:
        ref: Unique reference ID (e.g., "xref:5")
        field_type: "REF", "PAGEREF", or "NOTEREF"
        target_bookmark: The bookmark being referenced
        switches: Raw field switches (e.g., "\\h \\r")
        display_value: Current cached display value (may be stale)
        is_dirty: Whether field is marked for update
        is_hyperlink: Has \\h switch
        position: Location in document (e.g., "p:15")
        show_position: Has \\p switch
        number_format: "full" (\\w), "relative" (\\r), "no_context" (\\n), or None
        suppress_non_numeric: Has \\d switch
    """

    ref: str
    field_type: str
    target_bookmark: str
    switches: str
    display_value: str | None
    is_dirty: bool
    is_hyperlink: bool
    position: str
    show_position: bool = False
    number_format: str | None = None
    suppress_non_numeric: bool = False


@dataclass
class CrossReferenceTarget:
    """A potential target for a cross-reference.

    Represents something that can be referenced: a bookmark, heading,
    caption, or note.

    Attributes:
        type: "bookmark", "heading", "figure", "table", "footnote", or "endnote"
        bookmark_name: The bookmark name (may be auto-generated for headings/captions)
        display_name: Human-readable name
        text_preview: First ~100 chars of target content
        position: Location in document
        is_hidden: Is this a hidden _Ref bookmark?
        number: "1", "2.1", "Figure 3", etc. (for numbered items)
        level: Heading level 1-9 (for headings)
        sequence_id: SEQ field identifier "Figure", "Table" (for captions)
    """

    type: str
    bookmark_name: str
    display_name: str
    text_preview: str
    position: str
    is_hidden: bool
    number: str | None = None
    level: int | None = None
    sequence_id: str | None = None


@dataclass
class BookmarkInfo:
    """Information about a bookmark in the document.

    Bookmarks are named locations in a document that can be referenced
    by hyperlinks and cross-references. They support bidirectional
    reference tracking.

    Attributes:
        name: Bookmark name identifier
        bookmark_id: Internal OOXML bookmark ID
        location: Paragraph index where the bookmark is located (e.g., "p:5")
        text_preview: Preview of the bookmarked text (first 100 chars)
        is_hidden: Whether this is a hidden bookmark (starts with _)
        span_end_location: Location where bookmark ends if it spans multiple paragraphs
        referenced_by: List of cross-reference refs that reference this bookmark
    """

    name: str
    bookmark_id: str
    location: str
    text_preview: str = ""
    is_hidden: bool = False
    span_end_location: str | None = None
    referenced_by: list[str] = field(default_factory=list)


# =============================================================================
# Switch Mapping
# =============================================================================


# Maps display option to (field_type, list of switches)
# If field_type is None, use "REF" as default
DISPLAY_SWITCH_MAP: dict[str, tuple[str | None, list[str]]] = {
    # Basic display options
    "text": (None, []),  # REF with no special switches - shows bookmarked text
    "page": ("PAGEREF", []),  # PAGEREF shows page number
    # Numbered heading display options
    "number": (None, ["\\n"]),  # Paragraph number without trailing periods
    "full_number": (None, ["\\w"]),  # Full paragraph number (e.g., "1.2.3")
    "relative_number": (None, ["\\r"]),  # Relative paragraph number
    # Position display
    "above_below": (None, ["\\p"]),  # Insert "above" or "below"
    # Caption display options
    "label_number": (None, []),  # Full "Figure 1" - depends on bookmark placement
    "number_only": (None, ["\\n"]),  # Just the number from caption
    "label_only": (None, []),  # Just "Figure" or "Table" - requires text parsing
    "caption_text": (None, []),  # Caption text only - requires bookmark adjustment
}


class CrossReferenceOperations:
    """Operations for cross-reference manipulation.

    This class encapsulates cross-reference functionality, including:
    - Generating field codes for REF, PAGEREF, and NOTEREF fields
    - Mapping display options to field switches
    - Managing dirty flags for Word field calculation

    Cross-references rely on Word's field calculation engine. When the
    document is opened in Word, it will populate the field results with
    actual values based on the document structure.

    Example:
        >>> doc = Document("report.docx")
        >>> # Create field code elements for a cross-reference
        >>> ops = CrossReferenceOperations(doc)
        >>> field_runs = ops._create_field_code(
        ...     field_type="REF",
        ...     bookmark_name="_Ref123456",
        ...     switches=["\\h"],
        ...     placeholder_text="Section 2.1"
        ... )
    """

    def __init__(self, document: Document) -> None:
        """Initialize CrossReferenceOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    def _create_field_code(
        self,
        field_type: str,
        bookmark_name: str,
        switches: list[str],
        placeholder_text: str = "",
    ) -> list[etree._Element]:
        """Create XML elements for a field code (runs, not paragraph).

        This generates the complex field pattern used by Word for
        cross-references: begin -> instruction -> separate -> result -> end.

        The field begin element has the dirty flag set, which tells Word
        to recalculate the field value when the document is opened.

        Args:
            field_type: "REF", "PAGEREF", or "NOTEREF"
            bookmark_name: The bookmark being referenced
            switches: List of field switches (e.g., ["\\h", "\\r"])
            placeholder_text: Text to display until Word calculates the value

        Returns:
            List of w:r (run) elements that form the complete field structure.
            The caller is responsible for inserting these into a paragraph.

        Example:
            >>> runs = ops._create_field_code("REF", "_Ref123", ["\\h"], "see above")
            >>> # Insert runs into a paragraph
            >>> for run in runs:
            ...     paragraph.append(run)
        """
        runs: list[etree._Element] = []
        nsmap = {None: WORD_NAMESPACE}

        # Field begin with dirty flag
        run_begin = etree.Element(w("r"), nsmap=nsmap)
        fld_char_begin = etree.SubElement(run_begin, w("fldChar"))
        fld_char_begin.set(w("fldCharType"), "begin")
        fld_char_begin.set(w("dirty"), "true")
        runs.append(run_begin)

        # Field instruction
        run_instr = etree.Element(w("r"), nsmap=nsmap)
        instr_text = etree.SubElement(run_instr, w("instrText"))
        instr_text.set(f"{{{XML_NAMESPACE}}}space", "preserve")

        # Build instruction string: " FIELD_TYPE bookmark switches "
        switch_str = " ".join(switches) if switches else ""
        if switch_str:
            instr_text.text = f" {field_type} {bookmark_name} {switch_str} "
        else:
            instr_text.text = f" {field_type} {bookmark_name} "
        runs.append(run_instr)

        # Field separator
        run_sep = etree.Element(w("r"), nsmap=nsmap)
        fld_char_sep = etree.SubElement(run_sep, w("fldChar"))
        fld_char_sep.set(w("fldCharType"), "separate")
        runs.append(run_sep)

        # Placeholder result
        run_result = etree.Element(w("r"), nsmap=nsmap)
        text_result = etree.SubElement(run_result, w("t"))
        text_result.text = placeholder_text or "[Update field]"
        runs.append(run_result)

        # Field end
        run_end = etree.Element(w("r"), nsmap=nsmap)
        fld_char_end = etree.SubElement(run_end, w("fldChar"))
        fld_char_end.set(w("fldCharType"), "end")
        runs.append(run_end)

        return runs

    def _get_switches_for_display(
        self,
        display: str,
        hyperlink: bool = True,
    ) -> tuple[str, list[str]]:
        """Map a display option to field type and switches.

        Converts a human-readable display option to the corresponding
        Word field type and switches.

        Args:
            display: Display option from the API:
                - "text": Show bookmarked text content
                - "page": Show page number (uses PAGEREF)
                - "number": Paragraph number without trailing periods (\\n)
                - "full_number": Full paragraph number like "1.2.3" (\\w)
                - "relative_number": Relative paragraph number (\\r)
                - "above_below": "above" or "below" based on position (\\p)
                - "label_number": Full "Figure 1" label
                - "number_only": Just the number from caption (\\n)
                - "label_only": Just "Figure" or "Table"
                - "caption_text": Caption text only
            hyperlink: If True, add \\h switch to make reference clickable

        Returns:
            Tuple of (field_type, switches_list) where:
            - field_type is "REF", "PAGEREF", or "NOTEREF"
            - switches_list is a list of switch strings

        Raises:
            ValueError: If display option is not recognized

        Example:
            >>> ops._get_switches_for_display("page", hyperlink=True)
            ("PAGEREF", ["\\h"])
            >>> ops._get_switches_for_display("number", hyperlink=False)
            ("REF", ["\\n"])
        """
        if display not in DISPLAY_SWITCH_MAP:
            valid_options = ", ".join(sorted(DISPLAY_SWITCH_MAP.keys()))
            raise ValueError(f"Unknown display option '{display}'. Valid options: {valid_options}")

        field_type_override, switches = DISPLAY_SWITCH_MAP[display]

        # Determine field type
        field_type = field_type_override or "REF"

        # Copy switches list to avoid modifying the original
        result_switches = list(switches)

        # Add hyperlink switch if requested
        if hyperlink:
            result_switches.append("\\h")

        return field_type, result_switches

    # =========================================================================
    # Bookmark Management (Phase 2)
    # =========================================================================

    def _validate_bookmark_name(self, name: str) -> None:
        """Validate a bookmark name according to Word's rules.

        Bookmark names must:
        - Start with a letter (a-z, A-Z)
        - Contain only alphanumeric characters and underscores
        - Be at most 40 characters long
        - Not contain spaces

        Args:
            name: The bookmark name to validate

        Raises:
            InvalidBookmarkNameError: If the name is invalid
        """
        if not name:
            raise InvalidBookmarkNameError(name, "bookmark name cannot be empty")

        if len(name) > 40:
            raise InvalidBookmarkNameError(
                name, f"bookmark name must be at most 40 characters (got {len(name)})"
            )

        if not name[0].isalpha():
            raise InvalidBookmarkNameError(name, "bookmark name must start with a letter")

        # Check for valid characters (alphanumeric and underscore only)
        if not re.match(r"^[A-Za-z][A-Za-z0-9_]*$", name):
            raise InvalidBookmarkNameError(
                name,
                "bookmark name can only contain letters, numbers, and underscores",
            )

    def _generate_ref_bookmark_name(self) -> str:
        """Generate a unique hidden bookmark name in the _Ref format.

        Word uses _Ref followed by 9 digits for auto-generated bookmarks.
        We generate a random 9-digit number and check for uniqueness.

        Returns:
            A unique bookmark name like "_Ref123456789"
        """
        existing_names = {bk.name for bk in self.list_bookmarks(include_hidden=True)}

        # Try up to 100 times to generate a unique name
        for _ in range(100):
            # Generate a 9-digit random number
            random_digits = random.randint(100000000, 999999999)
            bookmark_name = f"_Ref{random_digits}"

            if bookmark_name not in existing_names:
                return bookmark_name

        # Fallback: use incrementing counter based on existing _Ref bookmarks
        max_num = 0
        for name in existing_names:
            if name.startswith("_Ref"):
                try:
                    num = int(name[4:])
                    max_num = max(max_num, num)
                except ValueError:
                    pass

        return f"_Ref{max_num + 1:09d}"

    def create_bookmark(
        self,
        name: str,
        at: str,
        scope: str | dict | Any | None = None,
    ) -> str:
        """Create a named bookmark at the specified text.

        The bookmark will span the matched text, allowing cross-references
        to display that text content.

        Args:
            name: Bookmark name (must be unique, no spaces, max 40 chars,
                  must start with a letter, can contain letters, numbers, underscores)
            at: The text to bookmark
            scope: Optional scope to limit text search (paragraph ref, heading, etc.)

        Returns:
            The bookmark name

        Raises:
            InvalidBookmarkNameError: If bookmark name is invalid
            BookmarkAlreadyExistsError: If bookmark name already exists
            TextNotFoundError: If text not found
            AmbiguousTextError: If text found multiple times

        Example:
            >>> doc.cross_references.create_bookmark("ImportantClause", at="Force Majeure")
            'ImportantClause'
        """
        # Validate the bookmark name
        self._validate_bookmark_name(name)

        # Check if bookmark already exists
        existing = self.get_bookmark(name)
        if existing is not None:
            raise BookmarkAlreadyExistsError(name)

        # Find the text in the document
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(at, paragraphs)

        if not matches:
            scope_str = str(scope) if scope is not None and not isinstance(scope, str) else scope
            raise TextNotFoundError(at, scope_str)

        if len(matches) > 1:
            raise AmbiguousTextError(at, matches)

        match = matches[0]

        # Find the next available bookmark ID
        bookmark_id = self._get_next_bookmark_id()

        # Create bookmark elements
        bookmark_start = etree.Element(w("bookmarkStart"))
        bookmark_start.set(w("id"), str(bookmark_id))
        bookmark_start.set(w("name"), name)

        bookmark_end = etree.Element(w("bookmarkEnd"))
        bookmark_end.set(w("id"), str(bookmark_id))

        # Insert bookmark elements around the matched text
        # The bookmark start goes before the first run
        # The bookmark end goes after the last run
        paragraph = match.paragraph
        start_run = match.runs[match.start_run_index]
        end_run = match.runs[match.end_run_index]

        # Find position of start_run in the paragraph
        para_children = list(paragraph)
        start_index = para_children.index(start_run)
        end_index = para_children.index(end_run)

        # Insert bookmark start before the first run
        paragraph.insert(start_index, bookmark_start)

        # After inserting bookmark_start, end_index shifts by 1
        paragraph.insert(end_index + 2, bookmark_end)

        logger.info(f"Created bookmark '{name}' at text '{at[:50]}...'")

        return name

    def _get_next_bookmark_id(self) -> int:
        """Get the next available bookmark ID.

        Scans all bookmarkStart elements to find the maximum ID and returns
        the next available one.

        Returns:
            The next available bookmark ID (integer)
        """
        max_id = -1

        for bookmark_start in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}bookmarkStart"):
            try:
                bk_id = int(bookmark_start.get(w("id"), "0"))
                max_id = max(max_id, bk_id)
            except ValueError:
                pass

        return max_id + 1

    def list_bookmarks(
        self,
        include_hidden: bool = False,
    ) -> list[BookmarkInfo]:
        """List all bookmarks in the document.

        Args:
            include_hidden: If True, include hidden bookmarks that start with
                           underscore (like _Ref, _Toc, _GoBack). Default False
                           filters these out for a cleaner user experience.

        Returns:
            List of BookmarkInfo objects with bookmark details

        Example:
            >>> # List user-visible bookmarks only
            >>> for bk in doc.cross_references.list_bookmarks():
            ...     print(f"{bk.name}: {bk.text_preview[:50]}")

            >>> # Include hidden bookmarks
            >>> all_bookmarks = doc.cross_references.list_bookmarks(include_hidden=True)
        """
        bookmarks: list[BookmarkInfo] = []

        # Build a mapping of paragraph elements to their index
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        para_to_index = {id(p): i for i, p in enumerate(all_paragraphs)}

        # Find all bookmarkStart elements
        for bookmark_start in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}bookmarkStart"):
            name = bookmark_start.get(w("name"), "")
            bookmark_id = bookmark_start.get(w("id"), "")

            if not name:
                continue

            # Check if this is a hidden bookmark
            is_hidden = name.startswith("_")

            # Skip hidden bookmarks unless requested
            if is_hidden and not include_hidden:
                continue

            # Find the parent paragraph
            parent = bookmark_start.getparent()
            while parent is not None and parent.tag != f"{{{WORD_NAMESPACE}}}p":
                parent = parent.getparent()

            if parent is None:
                continue

            # Get paragraph index
            para_index = para_to_index.get(id(parent), 0)
            location = f"p:{para_index}"

            # Get text preview from the paragraph
            text_parts = []
            for t_elem in parent.iter(f"{{{WORD_NAMESPACE}}}t"):
                if t_elem.text:
                    text_parts.append(t_elem.text)
            text_preview = "".join(text_parts)[:100]

            # Find span end location (if bookmark spans multiple paragraphs)
            span_end_location = None
            for bookmark_end in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}bookmarkEnd"):
                if bookmark_end.get(w("id")) == bookmark_id:
                    end_parent = bookmark_end.getparent()
                    while end_parent is not None and end_parent.tag != f"{{{WORD_NAMESPACE}}}p":
                        end_parent = end_parent.getparent()

                    if end_parent is not None:
                        end_para_index = para_to_index.get(id(end_parent), para_index)
                        if end_para_index != para_index:
                            span_end_location = f"p:{end_para_index}"
                    break

            bookmarks.append(
                BookmarkInfo(
                    name=name,
                    bookmark_id=bookmark_id,
                    location=location,
                    text_preview=text_preview,
                    is_hidden=is_hidden,
                    span_end_location=span_end_location,
                    referenced_by=[],
                )
            )

        return bookmarks

    def get_bookmark(self, name: str) -> BookmarkInfo | None:
        """Get information about a specific bookmark.

        Args:
            name: The bookmark name to look up

        Returns:
            BookmarkInfo if found, None otherwise

        Example:
            >>> bk = doc.cross_references.get_bookmark("DefinitionsSection")
            >>> if bk:
            ...     print(f"Found at {bk.location}: {bk.text_preview}")
        """
        # Include hidden in the search since we're looking for a specific name
        for bookmark in self.list_bookmarks(include_hidden=True):
            if bookmark.name == name:
                return bookmark
        return None

    # =========================================================================
    # Phase 3: Basic Cross-Reference Insertion
    # =========================================================================

    def insert_cross_reference(
        self,
        target: str,
        display: str = "text",
        after: str | None = None,
        before: str | None = None,
        scope: str | dict | Any | None = None,
        hyperlink: bool = True,
        track: bool = False,
        author: str | None = None,
    ) -> str:
        """Insert a cross-reference to a target bookmark.

        Cross-references are dynamic fields that Word calculates when the
        document is opened. The reference will display content based on the
        display option (bookmark text, page number, "above"/"below", etc.).

        Args:
            target: The bookmark name to reference. The bookmark must exist
                in the document.
            display: How to display the reference. Options:
                - "text": Show the bookmarked text content (default)
                - "page": Show the page number where the bookmark is located
                - "above_below": Show "above" or "below" based on position
                - "number": Show paragraph number without trailing periods
                - "full_number": Show full paragraph number (e.g., "1.2.3")
                - "relative_number": Show relative paragraph number
                - "label_number": For captions, show "Figure 1" style
                - "number_only": For captions, show just the number
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            scope: Optional scope to limit text search (paragraph ref, heading, etc.)
            hyperlink: If True (default), make the reference a clickable link
            track: If True, wrap insertion in tracked change markup (not yet implemented)
            author: Optional author override for tracked changes (not yet implemented)

        Returns:
            The bookmark name that was referenced

        Raises:
            ValueError: If both after and before specified, or neither specified
            ValueError: If display option is invalid
            CrossReferenceTargetNotFoundError: If target bookmark doesn't exist
            TextNotFoundError: If anchor text not found
            AmbiguousTextError: If anchor text found multiple times

        Example:
            >>> # Reference a bookmark by its bookmarked text
            >>> doc.cross_references.insert_cross_reference(
            ...     target="DefinitionsSection",
            ...     display="text",
            ...     after="as defined in "
            ... )
            'DefinitionsSection'

            >>> # Insert a page reference
            >>> doc.cross_references.insert_cross_reference(
            ...     target="AppendixA",
            ...     display="page",
            ...     after="(see page "
            ... )
            'AppendixA'
        """
        # Validate after/before parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        # Resolve the target to a bookmark name
        bookmark_name, was_created = self._resolve_target(target)

        # Get field type and switches for the display option
        field_type, switches = self._get_switches_for_display(display, hyperlink)

        # Create the field code runs
        # Use bookmark text as placeholder for "text" display, otherwise use generic
        if display == "text":
            placeholder = self._get_bookmark_text_preview(bookmark_name)
        elif display == "page":
            placeholder = "[#]"
        elif display == "above_below":
            placeholder = "[above/below]"
        else:
            placeholder = "[Update field]"

        field_runs = self._create_field_code(
            field_type=field_type,
            bookmark_name=bookmark_name,
            switches=switches,
            placeholder_text=placeholder,
        )

        # Find the insertion position
        anchor_text = after if after is not None else before
        insert_after = after is not None

        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(anchor_text, paragraphs)

        if not matches:
            scope_str = str(scope) if scope is not None and not isinstance(scope, str) else scope
            raise TextNotFoundError(anchor_text, scope_str)

        if len(matches) > 1:
            raise AmbiguousTextError(anchor_text, matches)

        match = matches[0]

        # Insert the field at the position
        self._insert_field_at_position(field_runs, match, insert_after)

        logger.info(
            f"Inserted cross-reference to '{bookmark_name}' "
            f"with display='{display}' {'after' if insert_after else 'before'} "
            f"'{anchor_text[:30]}...'"
        )

        return bookmark_name

    def _resolve_target(self, target: str) -> tuple[str, bool]:
        """Resolve a target specification to a bookmark name.

        Handles multiple target formats:
        - Direct bookmark name: "MyBookmark"
        - Heading reference: "heading:Introduction" or "heading:Chapter 1"
        - Figure reference: "figure:1" or "figure:Architecture Diagram"
        - Table reference: "table:2" or "table:Revenue Data"

        Args:
            target: Target specification. Can be a bookmark name, heading reference,
                   or caption reference (figure/table).

        Returns:
            Tuple of (bookmark_name, was_created) where:
            - bookmark_name: The resolved bookmark name
            - was_created: True if a new bookmark was created for the target

        Raises:
            CrossReferenceTargetNotFoundError: If the target doesn't exist
        """
        from ..errors import CrossReferenceTargetNotFoundError

        # Check if this is a heading reference (Phase 4)
        if target.startswith("heading:"):
            heading_text = target[8:]  # Remove "heading:" prefix
            bookmark_name, error_msg = self._resolve_heading_target(heading_text)
            if error_msg is not None:
                raise CrossReferenceTargetNotFoundError(target, [])
            # Determine if bookmark was created (it exists if no error)
            # We return True if we created it, but _resolve_heading_target
            # handles creation transparently
            return bookmark_name, True

        # Check if this is a figure reference (Phase 5)
        if target.startswith("figure:"):
            identifier = target[7:]  # Remove "figure:" prefix
            bookmark_name, error_msg = self._resolve_caption_target("Figure", identifier)
            if error_msg is not None:
                raise CrossReferenceTargetNotFoundError(target, [])
            return bookmark_name, True

        # Check if this is a table reference (Phase 5)
        if target.startswith("table:"):
            identifier = target[6:]  # Remove "table:" prefix
            bookmark_name, error_msg = self._resolve_caption_target("Table", identifier)
            if error_msg is not None:
                raise CrossReferenceTargetNotFoundError(target, [])
            return bookmark_name, True

        # Check if this is a footnote reference (Phase 6)
        if target.startswith("footnote:"):
            note_id = target[9:]  # Remove "footnote:" prefix
            bookmark_name, error_msg = self._resolve_note_target("footnote", note_id)
            if error_msg is not None:
                raise CrossReferenceTargetNotFoundError(target, [])
            return bookmark_name, True

        # Check if this is an endnote reference (Phase 6)
        if target.startswith("endnote:"):
            note_id = target[8:]  # Remove "endnote:" prefix
            bookmark_name, error_msg = self._resolve_note_target("endnote", note_id)
            if error_msg is not None:
                raise CrossReferenceTargetNotFoundError(target, [])
            return bookmark_name, True

        # Check if this is a direct bookmark name
        bookmark = self.get_bookmark(target)

        if bookmark is None:
            # Get list of available bookmarks for the error message
            all_bookmarks = self.list_bookmarks(include_hidden=True)
            available_names = [bk.name for bk in all_bookmarks]
            raise CrossReferenceTargetNotFoundError(target, available_names)

        return target, False

    def _insert_field_at_position(
        self,
        field_runs: list[etree._Element],
        match: Any,
        insert_after: bool,
    ) -> None:
        """Insert field code runs at a text match position.

        The field runs are inserted into the same paragraph as the match,
        either after or before the matched text.

        Args:
            field_runs: List of w:r elements that form the field code
            match: TextSpan object representing where to insert
            insert_after: If True, insert after the match; if False, before
        """
        paragraph = match.paragraph

        if insert_after:
            # Insert after the last run of the match
            end_run = match.runs[match.end_run_index]
            run_index = list(paragraph).index(end_run)

            # Insert all field runs after the end_run
            for i, run in enumerate(field_runs):
                paragraph.insert(run_index + 1 + i, run)
        else:
            # Insert before the first run of the match
            start_run = match.runs[match.start_run_index]
            run_index = list(paragraph).index(start_run)

            # Insert all field runs before the start_run
            for i, run in enumerate(field_runs):
                paragraph.insert(run_index + i, run)

    def _get_bookmark_text_preview(self, bookmark_name: str) -> str:
        """Get a text preview for a bookmark to use as placeholder.

        Args:
            bookmark_name: The bookmark name to look up

        Returns:
            The first 50 characters of the bookmarked text, or "[Update field]"
            if the bookmark cannot be found or has no text.
        """
        bookmark = self.get_bookmark(bookmark_name)
        if bookmark and bookmark.text_preview:
            # Truncate to reasonable length for placeholder
            preview = bookmark.text_preview[:50]
            if len(bookmark.text_preview) > 50:
                preview += "..."
            return preview
        return "[Update field]"

    # =========================================================================
    # Phase 4: Heading References
    # =========================================================================

    def _resolve_heading_target(self, heading_text: str) -> tuple[str, str | None]:
        """Resolve a heading text to a bookmark name.

        Finds a heading by its text content (partial match supported) and
        returns or creates a bookmark at that heading.

        Args:
            heading_text: Text to search for in headings. Supports partial match.

        Returns:
            Tuple of (bookmark_name, error_message) where:
            - bookmark_name: The resolved or created bookmark name
            - error_message: None on success, error description on failure

        Example:
            >>> bookmark_name, err = ops._resolve_heading_target("Introduction")
            >>> if err is None:
            ...     print(f"Found/created bookmark: {bookmark_name}")
        """
        # Find the heading paragraph
        heading_para = self._find_heading_paragraph(heading_text)
        if heading_para is None:
            return "", f"No heading found containing '{heading_text}'"

        # Check if the paragraph already has a _Ref bookmark
        existing_bookmark = self._find_existing_ref_bookmark(heading_para)
        if existing_bookmark is not None:
            logger.debug(f"Reusing existing bookmark '{existing_bookmark}' for heading")
            return existing_bookmark, None

        # Create a new _Ref bookmark at the paragraph
        bookmark_name = self._generate_ref_bookmark_name()
        self._create_bookmark_at_paragraph(bookmark_name, heading_para)
        logger.info(f"Created bookmark '{bookmark_name}' for heading '{heading_text[:50]}...'")

        return bookmark_name, None

    def _find_heading_paragraph(self, heading_text: str) -> etree._Element | None:
        """Find a heading paragraph by its text content.

        Searches for paragraphs with heading styles (Heading 1-9) that
        contain the specified text. Supports partial matching.

        Args:
            heading_text: Text to search for in heading paragraphs.
                         Case-insensitive partial match.

        Returns:
            The first matching heading paragraph element, or None if not found.

        Note:
            If multiple headings match, returns the first one found in document order.
        """
        body = self._document.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return None

        heading_text_lower = heading_text.lower()

        # Search all paragraphs for heading styles
        for para in body.iter(f"{{{WORD_NAMESPACE}}}p"):
            # Check if this paragraph has a heading style
            style_name = self._get_paragraph_style(para)
            if style_name is None:
                continue

            # Check if it's a heading style (Heading1, Heading 1, etc.)
            heading_level = self._get_heading_level_from_style(style_name)
            if heading_level is None:
                continue

            # Extract paragraph text
            para_text = self._extract_paragraph_text(para)
            if para_text.lower().find(heading_text_lower) != -1:
                return para

        return None

    def _get_paragraph_style(self, para: etree._Element) -> str | None:
        """Get the style name of a paragraph.

        Args:
            para: Paragraph element

        Returns:
            Style name string or None if no style is applied
        """
        para_props = para.find(f"./{{{WORD_NAMESPACE}}}pPr")
        if para_props is not None:
            para_style = para_props.find(f"./{{{WORD_NAMESPACE}}}pStyle")
            if para_style is not None:
                return para_style.get(w("val"))
        return None

    def _get_heading_level_from_style(self, style: str) -> int | None:
        """Determine heading level from style name.

        Handles various heading style naming conventions:
        - "Heading1", "Heading2", etc.
        - "Heading 1", "Heading 2", etc.
        - "Title" (treated as level 1)

        Args:
            style: Style name to analyze

        Returns:
            Heading level (1-9) or None if not a heading style
        """
        if not style:
            return None

        style_lower = style.lower()

        # Handle "HeadingN" and "Heading N" patterns
        if style_lower.startswith("heading"):
            suffix = style_lower[7:].strip()
            if suffix.isdigit():
                level = int(suffix)
                if 1 <= level <= 9:
                    return level

        # Handle Title as level 1
        if style_lower == "title":
            return 1

        return None

    def _extract_paragraph_text(self, para: etree._Element) -> str:
        """Extract all text content from a paragraph.

        Args:
            para: Paragraph element

        Returns:
            Concatenated text from all text runs
        """
        text_parts = []
        for t_elem in para.iter(f"{{{WORD_NAMESPACE}}}t"):
            if t_elem.text:
                text_parts.append(t_elem.text)
        return "".join(text_parts)

    def _find_existing_ref_bookmark(self, para: etree._Element) -> str | None:
        """Check if a paragraph already has a _Ref bookmark.

        Looks for bookmarkStart elements with names starting with "_Ref"
        that are direct children of the paragraph or its immediate ancestors.

        Args:
            para: Paragraph element to check

        Returns:
            The bookmark name if found, None otherwise
        """
        # Check for bookmarkStart in this paragraph with _Ref prefix
        for bookmark_start in para.iter(f"{{{WORD_NAMESPACE}}}bookmarkStart"):
            name = bookmark_start.get(w("name"), "")
            if name.startswith("_Ref"):
                return name

        return None

    def _create_bookmark_at_paragraph(self, name: str, para: etree._Element) -> None:
        """Create a hidden _Ref bookmark at a paragraph.

        The bookmark wraps the entire paragraph content (all runs).

        Args:
            name: Bookmark name (should start with "_Ref")
            para: Paragraph element to bookmark
        """
        # Find the next available bookmark ID
        bookmark_id = self._get_next_bookmark_id()

        # Create bookmark elements
        bookmark_start = etree.Element(w("bookmarkStart"))
        bookmark_start.set(w("id"), str(bookmark_id))
        bookmark_start.set(w("name"), name)

        bookmark_end = etree.Element(w("bookmarkEnd"))
        bookmark_end.set(w("id"), str(bookmark_id))

        # Find all runs in the paragraph
        runs = list(para.findall(f"./{{{WORD_NAMESPACE}}}r"))

        if runs:
            # Insert bookmark start before first run
            first_run = runs[0]
            first_run_index = list(para).index(first_run)
            para.insert(first_run_index, bookmark_start)

            # Insert bookmark end after last run
            # (index shifted by 1 due to bookmark_start insertion)
            last_run = runs[-1]
            last_run_index = list(para).index(last_run)
            para.insert(last_run_index + 1, bookmark_end)
        else:
            # No runs - insert both at beginning of paragraph
            # (may happen with empty paragraphs)
            para_props = para.find(f"./{{{WORD_NAMESPACE}}}pPr")
            if para_props is not None:
                insert_index = list(para).index(para_props) + 1
            else:
                insert_index = 0
            para.insert(insert_index, bookmark_start)
            para.insert(insert_index + 1, bookmark_end)

        logger.debug(f"Created bookmark '{name}' with id={bookmark_id} at paragraph")

    def create_heading_bookmark(
        self,
        heading_text: str,
        bookmark_name: str | None = None,
    ) -> str:
        """Create a bookmark at a heading for later cross-referencing.

        This is a convenience method for pre-creating bookmarks at headings
        before inserting cross-references. Useful when you want to control
        the bookmark name or create bookmarks in advance.

        Args:
            heading_text: Text to search for in headings (partial match).
            bookmark_name: Optional custom bookmark name. If None, a unique
                          _Ref name is auto-generated.

        Returns:
            The bookmark name that was created or found.

        Raises:
            CrossReferenceTargetNotFoundError: If no matching heading is found.
            InvalidBookmarkNameError: If custom bookmark_name is invalid.
            BookmarkAlreadyExistsError: If custom bookmark_name already exists.

        Example:
            >>> # Auto-generate bookmark name
            >>> bk = doc.cross_references.create_heading_bookmark("Introduction")
            >>> print(f"Created bookmark: {bk}")  # e.g., "_Ref123456789"
            >>>
            >>> # Use custom bookmark name
            >>> bk = doc.cross_references.create_heading_bookmark(
            ...     "Chapter 1",
            ...     bookmark_name="ChapterOneRef"
            ... )
        """
        from ..errors import CrossReferenceTargetNotFoundError

        # Find the heading paragraph
        heading_para = self._find_heading_paragraph(heading_text)
        if heading_para is None:
            raise CrossReferenceTargetNotFoundError(f"heading:{heading_text}", [])

        # Check if heading already has a _Ref bookmark
        existing_bookmark = self._find_existing_ref_bookmark(heading_para)
        if existing_bookmark is not None:
            # If user requested a specific name but heading already has one,
            # we could either raise an error or return the existing one.
            # For simplicity, return the existing one with a log message.
            if bookmark_name is not None and bookmark_name != existing_bookmark:
                logger.warning(
                    f"Heading already has bookmark '{existing_bookmark}', "
                    f"ignoring requested name '{bookmark_name}'"
                )
            return existing_bookmark

        # Determine bookmark name
        if bookmark_name is not None:
            # Validate custom bookmark name
            self._validate_bookmark_name(bookmark_name)
            # Check if it already exists
            if self.get_bookmark(bookmark_name) is not None:
                raise BookmarkAlreadyExistsError(bookmark_name)
            name = bookmark_name
        else:
            # Generate a unique _Ref name
            name = self._generate_ref_bookmark_name()

        # Create the bookmark at the heading
        self._create_bookmark_at_paragraph(name, heading_para)
        logger.info(f"Created heading bookmark '{name}' at '{heading_text[:50]}...'")

        return name

    # =========================================================================
    # Phase 5: Caption References (Figures and Tables)
    # =========================================================================

    def _resolve_caption_target(
        self,
        seq_id: str,
        identifier: str,
    ) -> tuple[str, str | None]:
        """Resolve a caption to a bookmark name.

        Finds a caption paragraph (e.g., "Figure 1: Architecture Diagram") by
        its SEQ field identifier and either number or caption text.

        Args:
            seq_id: The SEQ field identifier, either "Figure" or "Table"
            identifier: Either a number ("1", "2") or caption text
                       (e.g., "Architecture Diagram")

        Returns:
            Tuple of (bookmark_name, error_message) where:
            - bookmark_name: The resolved or created bookmark name
            - error_message: None on success, error description on failure

        Example:
            >>> bookmark_name, err = ops._resolve_caption_target("Figure", "1")
            >>> if err is None:
            ...     print(f"Found/created bookmark: {bookmark_name}")
        """
        # Find the caption paragraph
        caption_para = self._find_caption_paragraph(seq_id, identifier)
        if caption_para is None:
            return "", f"No {seq_id.lower()} found matching '{identifier}'"

        # Check if the paragraph already has a _Ref bookmark
        existing_bookmark = self._find_existing_ref_bookmark(caption_para)
        if existing_bookmark is not None:
            logger.debug(f"Reusing existing bookmark '{existing_bookmark}' for caption")
            return existing_bookmark, None

        # Create a new _Ref bookmark at the paragraph
        bookmark_name = self._generate_ref_bookmark_name()
        self._create_bookmark_at_paragraph(bookmark_name, caption_para)
        logger.info(f"Created bookmark '{bookmark_name}' for {seq_id.lower()} '{identifier}'")

        return bookmark_name, None

    def _find_caption_paragraph(
        self,
        seq_id: str,
        identifier: str,
    ) -> etree._Element | None:
        """Find a caption paragraph by its SEQ field identifier and number/text.

        Searches for paragraphs with Caption style that contain a SEQ field
        matching the specified identifier. Matches by either caption number
        or caption text.

        Args:
            seq_id: The SEQ field identifier, either "Figure" or "Table"
            identifier: Either a number string ("1", "2") or caption text
                       (case-insensitive partial match)

        Returns:
            The first matching caption paragraph element, or None if not found.

        Note:
            The method searches both simple fields (fldSimple) and complex
            fields (fldChar begin/separate/end pattern).
        """
        body = self._document.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return None

        # Determine if identifier is a number or text
        is_number = identifier.isdigit()
        identifier_lower = identifier.lower()

        # Search all paragraphs for caption style with matching SEQ field
        for para in body.iter(f"{{{WORD_NAMESPACE}}}p"):
            # Check if this paragraph has Caption style
            style_name = self._get_paragraph_style(para)
            if style_name is None or style_name.lower() != "caption":
                continue

            # Check for SEQ field with matching identifier
            caption_number = self._parse_caption_number(para, seq_id)
            if caption_number is None:
                continue

            # Match by number or by caption text
            if is_number:
                if caption_number == identifier:
                    return para
            else:
                # Match by caption text (partial match)
                caption_text = self._get_caption_text(para, seq_id)
                if caption_text.lower().find(identifier_lower) != -1:
                    return para

        return None

    def _parse_caption_number(
        self,
        paragraph: etree._Element,
        seq_id: str,
    ) -> str | None:
        """Extract caption number from SEQ field in a paragraph.

        Handles both simple fields (fldSimple) and complex fields
        (fldChar begin/separate/end pattern).

        Args:
            paragraph: Paragraph element containing the caption
            seq_id: The SEQ field identifier to look for (e.g., "Figure", "Table")

        Returns:
            The caption number as a string (e.g., "1", "2"), or None if no
            matching SEQ field is found.
        """
        # First, try to find simple field (fldSimple)
        for fld_simple in paragraph.iter(f"{{{WORD_NAMESPACE}}}fldSimple"):
            instr = fld_simple.get(w("instr"), "")
            if self._is_matching_seq_field(instr, seq_id):
                # Get the field result text
                for t_elem in fld_simple.iter(f"{{{WORD_NAMESPACE}}}t"):
                    if t_elem.text and t_elem.text.strip():
                        return t_elem.text.strip()

        # Try to find complex field pattern (fldChar begin/separate/end)
        in_field = False
        current_instr = ""
        found_sep = False

        for elem in paragraph.iter():
            if elem.tag == f"{{{WORD_NAMESPACE}}}fldChar":
                fld_type = elem.get(w("fldCharType"))
                if fld_type == "begin":
                    in_field = True
                    current_instr = ""
                    found_sep = False
                elif fld_type == "separate":
                    if in_field and self._is_matching_seq_field(current_instr, seq_id):
                        found_sep = True
                elif fld_type == "end":
                    in_field = False
                    found_sep = False
            elif elem.tag == f"{{{WORD_NAMESPACE}}}instrText":
                if in_field:
                    current_instr += elem.text or ""
            elif elem.tag == f"{{{WORD_NAMESPACE}}}t":
                if found_sep and elem.text and elem.text.strip():
                    return elem.text.strip()

        return None

    def _is_matching_seq_field(self, instr: str, seq_id: str) -> bool:
        """Check if an instruction string is a SEQ field for the given identifier.

        Args:
            instr: Field instruction text (e.g., " SEQ Figure \\* ARABIC ")
            seq_id: The SEQ identifier to match (e.g., "Figure", "Table")

        Returns:
            True if the instruction is a SEQ field matching the identifier
        """
        # Normalize and check for SEQ field
        instr_upper = instr.upper().strip()
        seq_id_upper = seq_id.upper()

        # Pattern: "SEQ identifier" possibly with switches
        if instr_upper.startswith("SEQ "):
            # Extract the identifier after "SEQ "
            parts = instr_upper[4:].split()
            if parts and parts[0] == seq_id_upper:
                return True

        return False

    def _get_caption_text(
        self,
        paragraph: etree._Element,
        seq_id: str,
    ) -> str:
        """Extract caption text after the number from a caption paragraph.

        Returns the text that follows the SEQ field, typically the caption
        description (e.g., ": Architecture Diagram" -> "Architecture Diagram").

        Args:
            paragraph: Caption paragraph element
            seq_id: The SEQ field identifier (used to locate the field)

        Returns:
            The caption text (description) without the label and number.
            Returns empty string if no text is found after the SEQ field.
        """
        # Extract all text from the paragraph
        full_text = self._extract_paragraph_text(paragraph)

        # Find the position after the number
        # Caption format is typically: "Figure 1: Description" or "Table 2 - Description"
        caption_number = self._parse_caption_number(paragraph, seq_id)
        if caption_number is None:
            return full_text

        # Find the number in the full text and get everything after it
        try:
            # Look for pattern like "Figure 1" or "Table 2"
            label_pattern = f"{seq_id} {caption_number}"
            idx = full_text.find(label_pattern)
            if idx != -1:
                # Skip past the label and number
                after_idx = idx + len(label_pattern)
                text_after = full_text[after_idx:].strip()
                # Remove leading separator characters (: - etc.)
                text_after = text_after.lstrip(":- ").strip()
                return text_after
        except (ValueError, IndexError):
            pass

        # Fallback: return everything after the first colon or dash
        for sep in [":", "-", "."]:
            if sep in full_text:
                idx = full_text.index(sep)
                return full_text[idx + 1 :].strip()

        return ""

    # =========================================================================
    # Phase 6: Note References and Convenience Methods
    # =========================================================================

    def _resolve_note_target(
        self,
        note_type: str,
        note_id: str,
    ) -> tuple[str, str | None]:
        """Resolve a footnote/endnote to a bookmark name.

        Finds the note reference in the document body and returns or creates
        a bookmark at that location.

        Args:
            note_type: Either "footnote" or "endnote"
            note_id: The note number/ID as a string

        Returns:
            Tuple of (bookmark_name, error_message) where:
            - bookmark_name: The resolved or created bookmark name
            - error_message: None on success, error description on failure

        Example:
            >>> bookmark_name, err = ops._resolve_note_target("footnote", "1")
            >>> if err is None:
            ...     print(f"Found/created bookmark: {bookmark_name}")
        """
        # Validate note_id
        if not note_id.isdigit():
            return "", f"Invalid {note_type} ID: '{note_id}' (must be numeric)"

        # Check if the note exists using NoteOperations
        note_ops = self._document._note_ops
        try:
            if note_type == "footnote":
                note_ops.get_footnote(note_id)
            else:
                note_ops.get_endnote(note_id)
        except Exception as e:
            return "", f"{note_type.capitalize()} {note_id} not found: {e}"

        # Check if there's already a _Ref bookmark for this note
        existing_bookmark = self._find_note_bookmark(note_type, note_id)
        if existing_bookmark is not None:
            logger.debug(f"Reusing existing bookmark '{existing_bookmark}' for {note_type}")
            return existing_bookmark, None

        # Create a new _Ref bookmark at the note reference
        try:
            bookmark_name = self._create_note_bookmark(note_type, note_id)
            logger.info(f"Created bookmark '{bookmark_name}' for {note_type} {note_id}")
            return bookmark_name, None
        except Exception as e:
            return "", f"Failed to create bookmark for {note_type} {note_id}: {e}"

    def _find_note_bookmark(self, note_type: str, note_id: str) -> str | None:
        """Find an existing _Ref bookmark for a note reference.

        Looks for bookmarkStart elements with _Ref prefix that encompass the
        note reference run.

        Args:
            note_type: Either "footnote" or "endnote"
            note_id: The note number/ID as a string

        Returns:
            The bookmark name if found, None otherwise
        """
        # Find the note reference location
        note_ops = self._document._note_ops
        if note_type == "footnote":
            ref_info = note_ops._find_footnote_reference(note_id)
        else:
            ref_info = note_ops._find_endnote_reference(note_id)

        if ref_info is None:
            return None

        run, para = ref_info

        # Get all direct children as a list to check positions
        children = list(para)

        # Find the position of the run in the paragraph
        try:
            run_pos = children.index(run)
        except ValueError:
            return None

        # Look for bookmarkStart elements that wrap this run
        for bookmark_start in para.iter(f"{{{WORD_NAMESPACE}}}bookmarkStart"):
            name = bookmark_start.get(w("name"), "")
            if not name.startswith("_Ref"):
                continue

            bookmark_id = bookmark_start.get(w("id"))

            # Check if bookmark_start is a direct child of para
            try:
                start_pos = children.index(bookmark_start)
            except ValueError:
                continue

            # Find the corresponding bookmark end
            for bookmark_end in para.iter(f"{{{WORD_NAMESPACE}}}bookmarkEnd"):
                if bookmark_end.get(w("id")) == bookmark_id:
                    try:
                        end_pos = children.index(bookmark_end)
                    except ValueError:
                        continue

                    # Check if the bookmark wraps around the run
                    if start_pos < run_pos < end_pos:
                        return name
                    break

        return None

    def _create_note_bookmark(self, note_type: str, note_id: str) -> str:
        """Create a _Ref bookmark at a note reference location.

        Creates a hidden bookmark that wraps the footnote or endnote reference
        marker in the document body.

        Args:
            note_type: Either "footnote" or "endnote"
            note_id: The note number/ID as a string

        Returns:
            The created bookmark name

        Raises:
            ValueError: If the note reference cannot be found
        """
        # Find the note reference location
        note_ops = self._document._note_ops
        if note_type == "footnote":
            ref_info = note_ops._find_footnote_reference(note_id)
        else:
            ref_info = note_ops._find_endnote_reference(note_id)

        if ref_info is None:
            raise ValueError(f"Cannot find {note_type} reference for ID {note_id}")

        run, para = ref_info

        # Generate a unique bookmark name
        bookmark_name = self._generate_ref_bookmark_name()

        # Get next bookmark ID
        bookmark_id = self._get_next_bookmark_id()

        # Create bookmark elements
        bookmark_start = etree.Element(w("bookmarkStart"))
        bookmark_start.set(w("id"), str(bookmark_id))
        bookmark_start.set(w("name"), bookmark_name)

        bookmark_end = etree.Element(w("bookmarkEnd"))
        bookmark_end.set(w("id"), str(bookmark_id))

        # Insert bookmark around the run containing the note reference
        run_index = list(para).index(run)

        # Insert bookmark start before the run
        para.insert(run_index, bookmark_start)

        # Insert bookmark end after the run (shifted by 1 due to bookmark_start)
        para.insert(run_index + 2, bookmark_end)

        logger.debug(f"Created note bookmark '{bookmark_name}' for {note_type} {note_id}")

        return bookmark_name

    def insert_page_reference(
        self,
        target: str,
        after: str | None = None,
        before: str | None = None,
        scope: str | dict | Any | None = None,
        show_position: bool = False,
        hyperlink: bool = True,
        track: bool = False,
        author: str | None = None,
    ) -> str:
        """Insert a page number reference to a target.

        This is a convenience method for inserting cross-references that
        display the page number where the target is located.

        Args:
            target: The target to reference. Can be:
                - Bookmark name: "MyBookmark"
                - Heading reference: "heading:Chapter 1"
                - Figure reference: "figure:1" or "figure:Architecture Diagram"
                - Table reference: "table:2" or "table:Revenue Data"
                - Note reference: "footnote:1" or "endnote:2"
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            scope: Optional scope to limit text search
            show_position: If True, add "above" or "below" indicator (\\p switch)
            hyperlink: If True (default), make the reference a clickable link
            track: If True, wrap insertion in tracked change markup
            author: Optional author override for tracked changes

        Returns:
            The bookmark name that was referenced

        Raises:
            ValueError: If both after and before specified, or neither
            CrossReferenceTargetNotFoundError: If target doesn't exist
            TextNotFoundError: If anchor text not found

        Example:
            >>> doc.cross_references.insert_page_reference(
            ...     target="AppendixA",
            ...     after="(see page "
            ... )
            >>> # Creates: "(see page [5])" where [5] is a PAGEREF field
        """
        # Validate after/before parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        # Resolve the target to a bookmark name
        bookmark_name, was_created = self._resolve_target(target)

        # Build switches for PAGEREF
        switches: list[str] = []
        if show_position:
            switches.append("\\p")
        if hyperlink:
            switches.append("\\h")

        # Create the field code runs
        field_runs = self._create_field_code(
            field_type="PAGEREF",
            bookmark_name=bookmark_name,
            switches=switches,
            placeholder_text="[#]",
        )

        # Find the insertion position
        anchor_text = after if after is not None else before
        insert_after = after is not None

        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(anchor_text, paragraphs)

        if not matches:
            from ..errors import TextNotFoundError

            scope_str = str(scope) if scope is not None and not isinstance(scope, str) else scope
            raise TextNotFoundError(anchor_text, scope_str)

        if len(matches) > 1:
            from ..errors import AmbiguousTextError

            raise AmbiguousTextError(anchor_text, matches)

        match = matches[0]

        # Insert the field at the position
        self._insert_field_at_position(field_runs, match, insert_after)

        logger.info(
            f"Inserted page reference to '{bookmark_name}' "
            f"{'after' if insert_after else 'before'} '{anchor_text[:30]}...'"
        )

        return bookmark_name

    def insert_note_reference(
        self,
        note_type: str,
        note_id: int | str,
        after: str | None = None,
        before: str | None = None,
        scope: str | dict | Any | None = None,
        show_position: bool = False,
        use_note_style: bool = True,
        hyperlink: bool = True,
        track: bool = False,
        author: str | None = None,
    ) -> str:
        """Insert a reference to a footnote or endnote.

        Creates a NOTEREF field that displays the note number with optional
        formatting to match the document's footnote/endnote style.

        Args:
            note_type: Either "footnote" or "endnote"
            note_id: The note number/ID to reference
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            scope: Optional scope to limit text search
            show_position: If True, add "above" or "below" indicator (\\p switch)
            use_note_style: If True (default), use footnote/endnote reference
                           formatting style (\\f switch)
            hyperlink: If True (default), make the reference a clickable link
            track: If True, wrap insertion in tracked change markup
            author: Optional author override for tracked changes

        Returns:
            The bookmark name that was referenced

        Raises:
            ValueError: If note_type is invalid, or after/before issues
            CrossReferenceTargetNotFoundError: If note doesn't exist
            TextNotFoundError: If anchor text not found

        Example:
            >>> # Reference footnote 1
            >>> doc.cross_references.insert_note_reference(
            ...     note_type="footnote",
            ...     note_id=1,
            ...     after="(see also footnote "
            ... )
            >>> # Creates: "(see also footnote [1])" where [1] is a NOTEREF field
        """
        # Validate note_type
        if note_type not in ("footnote", "endnote"):
            raise ValueError(f"note_type must be 'footnote' or 'endnote', got '{note_type}'")

        # Validate after/before parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        # Resolve the note to a bookmark
        note_id_str = str(note_id)
        bookmark_name, error_msg = self._resolve_note_target(note_type, note_id_str)
        if error_msg is not None:
            from ..errors import CrossReferenceTargetNotFoundError

            raise CrossReferenceTargetNotFoundError(f"{note_type}:{note_id}", [])

        # Build switches for NOTEREF
        switches: list[str] = []
        if use_note_style:
            switches.append("\\f")
        if show_position:
            switches.append("\\p")
        if hyperlink:
            switches.append("\\h")

        # Create the field code runs
        field_runs = self._create_field_code(
            field_type="NOTEREF",
            bookmark_name=bookmark_name,
            switches=switches,
            placeholder_text=note_id_str,
        )

        # Find the insertion position
        anchor_text = after if after is not None else before
        insert_after = after is not None

        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(anchor_text, paragraphs)

        if not matches:
            from ..errors import TextNotFoundError

            scope_str = str(scope) if scope is not None and not isinstance(scope, str) else scope
            raise TextNotFoundError(anchor_text, scope_str)

        if len(matches) > 1:
            from ..errors import AmbiguousTextError

            raise AmbiguousTextError(anchor_text, matches)

        match = matches[0]

        # Insert the field at the position
        self._insert_field_at_position(field_runs, match, insert_after)

        logger.info(
            f"Inserted {note_type} reference to note {note_id_str} "
            f"{'after' if insert_after else 'before'} '{anchor_text[:30]}...'"
        )

        return bookmark_name

    # =========================================================================
    # Phase 7: Inspection and Field Management
    # =========================================================================

    def get_cross_references(self) -> list[CrossReference]:
        """Get all cross-references in the document.

        Scans the document for REF, PAGEREF, and NOTEREF fields and returns
        information about each one.

        Returns:
            List of CrossReference objects containing:
            - ref: Unique reference ID (e.g., "xref:0", "xref:1")
            - field_type: "REF", "PAGEREF", or "NOTEREF"
            - target_bookmark: The bookmark name being referenced
            - switches: Raw field switches string
            - display_value: Current cached display value (may be stale)
            - is_dirty: Whether field is marked for update
            - is_hyperlink: Whether \\h switch is present
            - position: Location in document (e.g., "p:5")
            - show_position: Whether \\p switch is present
            - number_format: "full", "relative", or "no_context" (from \\w, \\r, \\n)
            - suppress_non_numeric: Whether \\d switch is present

        Example:
            >>> doc = Document("report.docx")
            >>> for xref in doc.cross_references.get_cross_references():
            ...     print(f"{xref.ref}: {xref.field_type} -> {xref.target_bookmark}")
            xref:0: REF -> _Ref123456789
            xref:1: PAGEREF -> AppendixA
            xref:2: NOTEREF -> _Ref987654321
        """
        cross_refs: list[CrossReference] = []

        # Extract all cross-reference fields
        fields = self._extract_fields_from_body(["REF", "PAGEREF", "NOTEREF"])

        for idx, (field_type, field_elem, position) in enumerate(fields):
            xref = self._build_cross_reference_from_field(
                ref_id=f"xref:{idx}",
                field_type=field_type,
                field_elem=field_elem,
                position=position,
            )
            if xref is not None:
                cross_refs.append(xref)

        return cross_refs

    def get_cross_reference_targets(self) -> list[CrossReferenceTarget]:
        """Get all potential cross-reference targets in the document.

        Returns a list of all elements that can be referenced:
        - Bookmarks (both user-defined and hidden _Ref bookmarks)
        - Headings (paragraphs with Heading 1-9 styles)
        - Figures (paragraphs with Caption style containing SEQ Figure)
        - Tables (paragraphs with Caption style containing SEQ Table)
        - Footnotes (footnote reference marks in the document)
        - Endnotes (endnote reference marks in the document)

        Returns:
            List of CrossReferenceTarget objects with target details.

        Example:
            >>> doc = Document("report.docx")
            >>> targets = doc.cross_references.get_cross_reference_targets()
            >>> for target in targets:
            ...     print(f"{target.type}: {target.display_name}")
            bookmark: DefinitionsSection
            heading: 1. Introduction
            heading: 2. Methodology
            figure: Figure 1: Architecture Diagram
            table: Table 1: Results Summary
            footnote: Footnote 1
        """
        targets: list[CrossReferenceTarget] = []

        # Get all bookmarks
        targets.extend(self._get_bookmark_targets())

        # Get all headings
        targets.extend(self._get_heading_targets())

        # Get all figure captions
        targets.extend(self._get_caption_targets("Figure"))

        # Get all table captions
        targets.extend(self._get_caption_targets("Table"))

        # Get all footnotes
        targets.extend(self._get_note_targets("footnote"))

        # Get all endnotes
        targets.extend(self._get_note_targets("endnote"))

        return targets

    def mark_cross_references_dirty(self) -> int:
        """Mark all cross-reference fields for update.

        Sets the w:dirty="true" attribute on all REF, PAGEREF, and NOTEREF
        field begin elements. This tells Word to recalculate the field values
        when the document is opened.

        Returns:
            The number of fields that were marked dirty.

        Example:
            >>> doc = Document("report.docx")
            >>> count = doc.cross_references.mark_cross_references_dirty()
            >>> print(f"Marked {count} cross-references for update")
            Marked 15 cross-references for update
            >>> doc.save("report_updated.docx")
        """
        count = 0

        body = self._document.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return count

        # Find all fldChar elements with fldCharType="begin"
        for fld_char in body.iter(f"{{{WORD_NAMESPACE}}}fldChar"):
            fld_type = fld_char.get(w("fldCharType"))
            if fld_type != "begin":
                continue

            # Check if this is a cross-reference field by looking at the instruction
            # The instruction follows the begin fldChar
            instr_text = self._get_field_instruction_from_begin(fld_char)
            if instr_text is None:
                continue

            # Parse the instruction to check if it's a cross-reference field
            parsed = self._parse_field_instruction(instr_text)
            if parsed["field_type"] in ("REF", "PAGEREF", "NOTEREF"):
                # Set dirty flag
                fld_char.set(w("dirty"), "true")
                count += 1
                logger.debug(
                    f"Marked {parsed['field_type']} field to '{parsed['bookmark']}' as dirty"
                )

        logger.info(f"Marked {count} cross-reference fields as dirty")
        return count

    def _parse_field_instruction(self, instruction: str) -> dict:
        """Parse a field instruction string into its components.

        Handles instructions like:
        - " REF _Ref123456 \\h "
        - " PAGEREF MyBookmark \\h \\p "
        - " NOTEREF _Ref789 \\f \\h "

        Args:
            instruction: The field instruction text

        Returns:
            Dictionary with keys:
            - field_type: "REF", "PAGEREF", "NOTEREF", or "UNKNOWN"
            - bookmark: The bookmark name being referenced
            - switches: Raw switches string
            - is_hyperlink: True if \\h switch present
            - show_position: True if \\p switch present
            - number_format: "full" (\\w), "relative" (\\r), "no_context" (\\n), or None
            - suppress_non_numeric: True if \\d switch present
            - use_note_style: True if \\f switch present (NOTEREF only)
        """
        result = {
            "field_type": "UNKNOWN",
            "bookmark": "",
            "switches": "",
            "is_hyperlink": False,
            "show_position": False,
            "number_format": None,
            "suppress_non_numeric": False,
            "use_note_style": False,
        }

        # Normalize whitespace
        instruction = instruction.strip()
        if not instruction:
            return result

        # Split into tokens
        # Handle quoted strings and switches
        tokens = []
        current = ""
        in_quotes = False

        for char in instruction:
            if char == '"':
                in_quotes = not in_quotes
                current += char
            elif char.isspace() and not in_quotes:
                if current:
                    tokens.append(current)
                    current = ""
            else:
                current += char

        if current:
            tokens.append(current)

        if not tokens:
            return result

        # First token should be the field type
        field_type = tokens[0].upper()
        if field_type in ("REF", "PAGEREF", "NOTEREF"):
            result["field_type"] = field_type
        else:
            return result

        # Second token should be the bookmark name
        if len(tokens) > 1:
            result["bookmark"] = tokens[1]

        # Remaining tokens are switches
        switch_tokens = tokens[2:] if len(tokens) > 2 else []
        result["switches"] = " ".join(switch_tokens)

        # Parse individual switches
        switches_str = result["switches"]
        result["is_hyperlink"] = "\\h" in switches_str
        result["show_position"] = "\\p" in switches_str
        result["suppress_non_numeric"] = "\\d" in switches_str
        result["use_note_style"] = "\\f" in switches_str

        # Number format switches
        if "\\w" in switches_str:
            result["number_format"] = "full"
        elif "\\r" in switches_str:
            result["number_format"] = "relative"
        elif "\\n" in switches_str:
            result["number_format"] = "no_context"

        return result

    def _extract_fields_from_body(
        self,
        field_types: list[str],
    ) -> list[tuple[str, etree._Element, str]]:
        """Extract all fields of specified types from the document body.

        Finds complex field codes (begin/instrText/separate/result/end pattern)
        and returns information about each matching field.

        Args:
            field_types: List of field type names to find (e.g., ["REF", "PAGEREF"])

        Returns:
            List of tuples containing:
            - field_type: The matched field type string
            - field_elem: The fldChar begin element
            - position: Location string (e.g., "p:5")
        """
        results: list[tuple[str, etree._Element, str]] = []

        body = self._document.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return results

        # Build a paragraph index map
        all_paragraphs = list(body.iter(f"{{{WORD_NAMESPACE}}}p"))
        para_to_index = {id(p): i for i, p in enumerate(all_paragraphs)}

        # Find all fldChar elements with fldCharType="begin"
        for fld_char in body.iter(f"{{{WORD_NAMESPACE}}}fldChar"):
            fld_type = fld_char.get(w("fldCharType"))
            if fld_type != "begin":
                continue

            # Get the field instruction
            instr_text = self._get_field_instruction_from_begin(fld_char)
            if instr_text is None:
                continue

            # Parse to get field type
            parsed = self._parse_field_instruction(instr_text)
            if parsed["field_type"] in field_types:
                # Find parent paragraph
                parent = fld_char.getparent()
                while parent is not None and parent.tag != f"{{{WORD_NAMESPACE}}}p":
                    parent = parent.getparent()

                if parent is not None:
                    para_idx = para_to_index.get(id(parent), 0)
                    position = f"p:{para_idx}"
                else:
                    position = "unknown"

                results.append((parsed["field_type"], fld_char, position))

        return results

    def _get_field_instruction_from_begin(
        self,
        begin_fld_char: etree._Element,
    ) -> str | None:
        """Get the field instruction text following a begin fldChar.

        Traverses sibling runs to collect all instrText elements until
        the separate or end fldChar is found.

        Args:
            begin_fld_char: The fldChar element with fldCharType="begin"

        Returns:
            The concatenated instruction text, or None if not found
        """
        instruction_parts = []

        # Get the parent run of the begin fldChar
        begin_run = begin_fld_char.getparent()
        if begin_run is None or begin_run.tag != f"{{{WORD_NAMESPACE}}}r":
            return None

        # Get the parent paragraph
        para = begin_run.getparent()
        if para is None:
            return None

        # Get all children of the paragraph
        children = list(para)

        # Find the index of the begin run
        try:
            begin_idx = children.index(begin_run)
        except ValueError:
            return None

        # Iterate through siblings looking for instrText and end/separate
        for child in children[begin_idx:]:
            if child.tag == f"{{{WORD_NAMESPACE}}}r":
                # Check for fldChar (separate or end)
                fld_char = child.find(f"{{{WORD_NAMESPACE}}}fldChar")
                if fld_char is not None and child != begin_run:
                    fld_type = fld_char.get(w("fldCharType"))
                    if fld_type in ("separate", "end"):
                        break

                # Collect instrText
                instr_text = child.find(f"{{{WORD_NAMESPACE}}}instrText")
                if instr_text is not None and instr_text.text:
                    instruction_parts.append(instr_text.text)

        if instruction_parts:
            return "".join(instruction_parts)
        return None

    def _build_cross_reference_from_field(
        self,
        ref_id: str,
        field_type: str,
        field_elem: etree._Element,
        position: str,
    ) -> CrossReference | None:
        """Build a CrossReference object from a field begin element.

        Args:
            ref_id: Unique reference ID to assign
            field_type: The field type ("REF", "PAGEREF", "NOTEREF")
            field_elem: The fldChar begin element
            position: Location string (e.g., "p:5")

        Returns:
            CrossReference object, or None if the field cannot be parsed
        """
        # Get the full instruction text
        instr_text = self._get_field_instruction_from_begin(field_elem)
        if instr_text is None:
            return None

        # Parse the instruction
        parsed = self._parse_field_instruction(instr_text)

        # Check if dirty
        is_dirty = field_elem.get(w("dirty")) == "true"

        # Get the display value from the field result
        display_value = self._get_field_result_text(field_elem)

        return CrossReference(
            ref=ref_id,
            field_type=field_type,
            target_bookmark=parsed["bookmark"],
            switches=parsed["switches"],
            display_value=display_value,
            is_dirty=is_dirty,
            is_hyperlink=parsed["is_hyperlink"],
            position=position,
            show_position=parsed["show_position"],
            number_format=parsed["number_format"],
            suppress_non_numeric=parsed["suppress_non_numeric"],
        )

    def _get_field_result_text(self, begin_fld_char: etree._Element) -> str | None:
        """Get the display text from a field's result section.

        The result is the text between the separate and end fldChar elements.

        Args:
            begin_fld_char: The fldChar element with fldCharType="begin"

        Returns:
            The result text, or None if not found
        """
        # Get the parent run
        begin_run = begin_fld_char.getparent()
        if begin_run is None or begin_run.tag != f"{{{WORD_NAMESPACE}}}r":
            return None

        # Get the parent paragraph
        para = begin_run.getparent()
        if para is None:
            return None

        # Get all children
        children = list(para)

        # Find the begin run index
        try:
            begin_idx = children.index(begin_run)
        except ValueError:
            return None

        # Track state as we iterate
        found_separate = False
        result_parts = []

        for child in children[begin_idx:]:
            if child.tag == f"{{{WORD_NAMESPACE}}}r":
                fld_char = child.find(f"{{{WORD_NAMESPACE}}}fldChar")
                if fld_char is not None:
                    fld_type = fld_char.get(w("fldCharType"))
                    if fld_type == "separate":
                        found_separate = True
                        continue
                    elif fld_type == "end":
                        break

                # Collect text after separate
                if found_separate:
                    t_elem = child.find(f"{{{WORD_NAMESPACE}}}t")
                    if t_elem is not None and t_elem.text:
                        result_parts.append(t_elem.text)

        if result_parts:
            return "".join(result_parts)
        return None

    def _get_bookmark_targets(self) -> list[CrossReferenceTarget]:
        """Get all bookmarks as potential cross-reference targets."""
        targets: list[CrossReferenceTarget] = []

        for bookmark in self.list_bookmarks(include_hidden=True):
            targets.append(
                CrossReferenceTarget(
                    type="bookmark",
                    bookmark_name=bookmark.name,
                    display_name=bookmark.name,
                    text_preview=bookmark.text_preview,
                    position=bookmark.location,
                    is_hidden=bookmark.is_hidden,
                    number=None,
                    level=None,
                    sequence_id=None,
                )
            )

        return targets

    def _get_heading_targets(self) -> list[CrossReferenceTarget]:
        """Get all headings as potential cross-reference targets."""
        targets: list[CrossReferenceTarget] = []

        body = self._document.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return targets

        # Build paragraph index map
        all_paragraphs = list(body.iter(f"{{{WORD_NAMESPACE}}}p"))
        para_to_index = {id(p): i for i, p in enumerate(all_paragraphs)}

        for para in all_paragraphs:
            style_name = self._get_paragraph_style(para)
            if style_name is None:
                continue

            level = self._get_heading_level_from_style(style_name)
            if level is None:
                continue

            # Extract paragraph text
            para_text = self._extract_paragraph_text(para)
            if not para_text.strip():
                continue

            # Check for existing _Ref bookmark
            existing_bookmark = self._find_existing_ref_bookmark(para)

            # Get position
            para_idx = para_to_index.get(id(para), 0)
            position = f"p:{para_idx}"

            targets.append(
                CrossReferenceTarget(
                    type="heading",
                    bookmark_name=existing_bookmark or "",
                    display_name=para_text[:100],
                    text_preview=para_text[:100],
                    position=position,
                    is_hidden=existing_bookmark is not None and existing_bookmark.startswith("_"),
                    number=None,  # Would need outline numbering to determine
                    level=level,
                    sequence_id=None,
                )
            )

        return targets

    def _get_caption_targets(self, seq_id: str) -> list[CrossReferenceTarget]:
        """Get all captions of specified type as potential cross-reference targets.

        Args:
            seq_id: "Figure" or "Table"
        """
        targets: list[CrossReferenceTarget] = []

        body = self._document.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return targets

        # Build paragraph index map
        all_paragraphs = list(body.iter(f"{{{WORD_NAMESPACE}}}p"))
        para_to_index = {id(p): i for i, p in enumerate(all_paragraphs)}

        for para in all_paragraphs:
            # Check if this paragraph has Caption style
            style_name = self._get_paragraph_style(para)
            if style_name is None or style_name.lower() != "caption":
                continue

            # Check for SEQ field with matching identifier
            caption_number = self._parse_caption_number(para, seq_id)
            if caption_number is None:
                continue

            # Get caption text
            para_text = self._extract_paragraph_text(para)
            caption_text = self._get_caption_text(para, seq_id)

            # Check for existing _Ref bookmark
            existing_bookmark = self._find_existing_ref_bookmark(para)

            # Get position
            para_idx = para_to_index.get(id(para), 0)
            position = f"p:{para_idx}"

            # Build display name
            display_name = f"{seq_id} {caption_number}"
            if caption_text:
                display_name += f": {caption_text[:50]}"

            targets.append(
                CrossReferenceTarget(
                    type=seq_id.lower(),
                    bookmark_name=existing_bookmark or "",
                    display_name=display_name,
                    text_preview=para_text[:100],
                    position=position,
                    is_hidden=existing_bookmark is not None and existing_bookmark.startswith("_"),
                    number=caption_number,
                    level=None,
                    sequence_id=seq_id,
                )
            )

        return targets

    def _get_note_targets(self, note_type: str) -> list[CrossReferenceTarget]:
        """Get all notes of specified type as potential cross-reference targets.

        Args:
            note_type: "footnote" or "endnote"
        """
        targets: list[CrossReferenceTarget] = []

        # Get notes from NoteOperations
        note_ops = self._document._note_ops

        try:
            if note_type == "footnote":
                notes = note_ops.footnotes
            else:
                notes = note_ops.endnotes
        except Exception:
            # If we can't get notes, return empty list
            return targets

        for note in notes:
            # Get note ID - Footnote/Endnote objects use 'id' attribute
            note_id = note.id if hasattr(note, "id") else ""

            # Check for existing bookmark
            existing_bookmark = self._find_note_bookmark(note_type, str(note_id))

            # Build display name
            display_name = f"{note_type.capitalize()} {note_id}"

            # Get note text content - Footnote/Endnote objects use 'text' property
            note_text = ""
            if hasattr(note, "text"):
                note_text = note.text[:100] if note.text else ""

            targets.append(
                CrossReferenceTarget(
                    type=note_type,
                    bookmark_name=existing_bookmark or "",
                    display_name=display_name,
                    text_preview=note_text,
                    position="",  # Position would require finding the reference in body
                    is_hidden=existing_bookmark is not None and existing_bookmark.startswith("_"),
                    number=str(note_id),
                    level=None,
                    sequence_id=None,
                )
            )

        return targets
