"""
Core types for the DocTree accessibility layer.

This module defines the fundamental data structures used to represent
Word document elements in a semantic, accessible way.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum, auto
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from lxml import etree


class ElementType(Enum):
    """Types of document elements that can be referenced.

    These map to OOXML element types and are used in ref prefixes.
    """

    # Document structure
    DOCUMENT = auto()
    BODY = auto()

    # Block-level elements
    PARAGRAPH = auto()
    RUN = auto()
    TABLE = auto()
    TABLE_ROW = auto()
    TABLE_CELL = auto()

    # Document parts
    HEADER = auto()
    FOOTER = auto()
    FOOTNOTE = auto()
    ENDNOTE = auto()

    # Tracked changes
    INSERTION = auto()
    DELETION = auto()

    # Annotations
    COMMENT = auto()
    BOOKMARK = auto()

    # Drawing elements
    IMAGE = auto()
    CHART = auto()
    SHAPE = auto()

    # Section (logical grouping)
    SECTION = auto()


# Mapping from ElementType to ref prefix
ELEMENT_TYPE_TO_PREFIX: dict[ElementType, str] = {
    ElementType.PARAGRAPH: "p",
    ElementType.RUN: "r",
    ElementType.TABLE: "tbl",
    ElementType.TABLE_ROW: "row",
    ElementType.TABLE_CELL: "cell",
    ElementType.HEADER: "hdr",
    ElementType.FOOTER: "ftr",
    ElementType.FOOTNOTE: "fn",
    ElementType.ENDNOTE: "en",
    ElementType.INSERTION: "ins",
    ElementType.DELETION: "del",
    ElementType.COMMENT: "cmt",
    ElementType.BOOKMARK: "bk",
    ElementType.IMAGE: "img",
    ElementType.CHART: "chart",
    ElementType.SHAPE: "shape",
    ElementType.SECTION: "sec",
}

# Reverse mapping from prefix to ElementType
PREFIX_TO_ELEMENT_TYPE: dict[str, ElementType] = {v: k for k, v in ELEMENT_TYPE_TO_PREFIX.items()}


class ChangeType(Enum):
    """Types of tracked changes."""

    INSERTION = auto()
    DELETION = auto()
    FORMAT = auto()
    MOVE_FROM = auto()
    MOVE_TO = auto()


@dataclass
class ChangeInfo:
    """Information about a tracked change.

    Attributes:
        change_type: Type of change (insertion, deletion, etc.)
        author: Author who made the change
        date: When the change was made
        change_id: OOXML change ID
        text: The changed text content
    """

    change_type: ChangeType
    author: str
    date: datetime | None = None
    change_id: str | None = None
    text: str = ""


@dataclass
class CommentInfo:
    """Information about a comment.

    Attributes:
        comment_id: Unique comment identifier
        author: Author of the comment
        date: When the comment was made
        text: Comment text content
        resolved: Whether the comment is resolved
        replies: List of reply comments
    """

    comment_id: str
    author: str
    text: str
    date: datetime | None = None
    resolved: bool = False
    replies: list[CommentInfo] = field(default_factory=list)


@dataclass
class Ref:
    """A reference to a document element.

    Refs provide stable, unambiguous identifiers for document elements.
    They can be ordinal-based (p:5) or fingerprint-based (p:~xK4mNp2q).

    Attributes:
        path: The full ref path (e.g., "p:5", "tbl:0/row:2/cell:1")

    Examples:
        >>> ref = Ref.parse("p:5")
        >>> ref.path
        'p:5'
        >>> ref.element_type
        <ElementType.PARAGRAPH: 3>
        >>> ref.ordinal
        5

        >>> ref = Ref.parse("tbl:0/row:2/cell:1/p:0")
        >>> ref.path
        'tbl:0/row:2/cell:1/p:0'
        >>> ref.segments
        [('tbl', '0'), ('row', '2'), ('cell', '1'), ('p', '0')]
    """

    path: str

    # Regex pattern for parsing a single ref segment
    _SEGMENT_PATTERN = re.compile(r"^([a-z]+):(.+)$")

    # Regex for fingerprint identifiers (start with ~)
    _FINGERPRINT_PATTERN = re.compile(r"^~[A-Za-z0-9]+$")

    @classmethod
    def parse(cls, ref_string: str) -> Ref:
        """Parse a ref string into a Ref object.

        Args:
            ref_string: The ref string (e.g., "p:5", "tbl:0/row:2/cell:1")

        Returns:
            Parsed Ref object

        Raises:
            ValueError: If the ref string is invalid
        """
        if not ref_string or not ref_string.strip():
            raise ValueError("Ref string cannot be empty")

        # Validate each segment
        segments = ref_string.split("/")
        for segment in segments:
            match = cls._SEGMENT_PATTERN.match(segment)
            if not match:
                raise ValueError(f"Invalid ref segment format: '{segment}'")

            prefix = match.group(1)
            identifier = match.group(2)

            if prefix not in PREFIX_TO_ELEMENT_TYPE:
                raise ValueError(f"Unknown element type prefix: '{prefix}'")

            # Validate identifier (must be integer ordinal or fingerprint)
            if not cls._FINGERPRINT_PATTERN.match(identifier):
                try:
                    int(identifier)
                except ValueError:
                    msg = (
                        f"Invalid identifier '{identifier}': must be integer or fingerprint (~xxx)"
                    )
                    raise ValueError(msg)

        return cls(path=ref_string)

    @property
    def segments(self) -> list[tuple[str, str]]:
        """Parse the path into (prefix, identifier) tuples.

        Returns:
            List of (prefix, identifier) tuples
        """
        result = []
        for segment in self.path.split("/"):
            match = self._SEGMENT_PATTERN.match(segment)
            if match:
                result.append((match.group(1), match.group(2)))
        return result

    @property
    def element_type(self) -> ElementType:
        """Get the element type of the final segment.

        Returns:
            ElementType of the referenced element
        """
        segments = self.segments
        if not segments:
            raise ValueError(f"Cannot determine element type from path: {self.path}")
        prefix = segments[-1][0]
        return PREFIX_TO_ELEMENT_TYPE[prefix]

    @property
    def ordinal(self) -> int | None:
        """Get the ordinal index of the final segment, if applicable.

        Returns:
            Integer ordinal or None if this is a fingerprint ref
        """
        segments = self.segments
        if not segments:
            return None
        identifier = segments[-1][1]
        if self._FINGERPRINT_PATTERN.match(identifier):
            return None
        return int(identifier)

    @property
    def fingerprint(self) -> str | None:
        """Get the fingerprint of the final segment, if applicable.

        Returns:
            Fingerprint string (without ~) or None if this is an ordinal ref
        """
        segments = self.segments
        if not segments:
            return None
        identifier = segments[-1][1]
        if self._FINGERPRINT_PATTERN.match(identifier):
            return identifier[1:]  # Remove leading ~
        return None

    @property
    def is_fingerprint(self) -> bool:
        """Check if this is a fingerprint-based ref.

        Returns:
            True if this ref uses a fingerprint identifier
        """
        return self.fingerprint is not None

    @property
    def parent_path(self) -> str | None:
        """Get the parent ref path, if any.

        Returns:
            Parent path string or None if this is a top-level ref

        Example:
            >>> Ref.parse("tbl:0/row:2/cell:1").parent_path
            'tbl:0/row:2'
        """
        parts = self.path.rsplit("/", 1)
        if len(parts) == 1:
            return None
        return parts[0]

    def with_child(self, element_type: ElementType, identifier: int | str) -> Ref:
        """Create a child ref under this ref.

        Args:
            element_type: Type of child element
            identifier: Ordinal index or fingerprint

        Returns:
            New Ref for the child element

        Example:
            >>> parent = Ref.parse("tbl:0/row:2")
            >>> child = parent.with_child(ElementType.TABLE_CELL, 1)
            >>> child.path
            'tbl:0/row:2/cell:1'
        """
        prefix = ELEMENT_TYPE_TO_PREFIX[element_type]
        child_segment = f"{prefix}:{identifier}"
        return Ref(path=f"{self.path}/{child_segment}")

    def __str__(self) -> str:
        return self.path

    def __hash__(self) -> int:
        return hash(self.path)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, Ref):
            return self.path == other.path
        if isinstance(other, str):
            return self.path == other
        return False


@dataclass
class ViewMode:
    """Configuration for what to include in an accessibility tree.

    Attributes:
        include_body: Include main document body
        include_headers: Include headers
        include_footers: Include footers
        include_footnotes: Include footnotes
        include_endnotes: Include endnotes
        include_comments: Include comment annotations
        include_tracked_changes: Include tracked change markers
        include_formatting: Include run-level formatting details
        verbosity: Output verbosity level ("minimal", "standard", "full")
    """

    include_body: bool = True
    include_headers: bool = False
    include_footers: bool = False
    include_footnotes: bool = False
    include_endnotes: bool = False
    include_comments: bool = False
    include_tracked_changes: bool = True
    include_formatting: bool = False
    verbosity: str = "standard"

    def __post_init__(self) -> None:
        """Validate verbosity level."""
        valid_levels = ("minimal", "standard", "full")
        if self.verbosity not in valid_levels:
            raise ValueError(f"verbosity must be one of {valid_levels}, got '{self.verbosity}'")


@dataclass
class AccessibilityNode:
    """A node in the accessibility tree representing a document element.

    This is the core data structure for the DocTree accessibility layer,
    providing a semantic view of document elements with stable references.

    Attributes:
        ref: Stable reference identifier for this element
        element_type: Type of document element
        text: Text content (may be empty for structural elements)
        children: Child nodes (for hierarchical elements)
        style: Applied style name (e.g., "Heading1", "Normal")
        level: Heading level (1-9) for heading elements
        change: Tracked change information if this element is part of a change
        comments: Comments attached to this element
        properties: Additional element-specific properties
        _element: Reference to underlying lxml element (for internal use)
    """

    ref: Ref
    element_type: ElementType
    text: str = ""
    children: list[AccessibilityNode] = field(default_factory=list)
    style: str | None = None
    level: int | None = None
    change: ChangeInfo | None = None
    comments: list[CommentInfo] = field(default_factory=list)
    properties: dict[str, str] = field(default_factory=dict)
    _element: etree._Element | None = field(default=None, repr=False, compare=False)

    @property
    def has_children(self) -> bool:
        """Check if this node has children."""
        return len(self.children) > 0

    @property
    def has_changes(self) -> bool:
        """Check if this node has tracked changes."""
        return self.change is not None

    @property
    def has_comments(self) -> bool:
        """Check if this node has comments."""
        return len(self.comments) > 0

    def find_by_ref(self, ref: str | Ref) -> AccessibilityNode | None:
        """Find a descendant node by ref.

        Args:
            ref: Ref to search for

        Returns:
            Matching node or None
        """
        ref_str = ref if isinstance(ref, str) else ref.path

        if self.ref.path == ref_str:
            return self

        for child in self.children:
            result = child.find_by_ref(ref_str)
            if result:
                return result

        return None

    def find_all_by_type(self, element_type: ElementType) -> list[AccessibilityNode]:
        """Find all descendant nodes of a given type.

        Args:
            element_type: Type to search for

        Returns:
            List of matching nodes
        """
        results = []

        if self.element_type == element_type:
            results.append(self)

        for child in self.children:
            results.extend(child.find_all_by_type(element_type))

        return results
