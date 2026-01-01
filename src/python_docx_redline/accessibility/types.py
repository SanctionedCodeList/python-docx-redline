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
    HYPERLINK = auto()
    CROSS_REFERENCE = auto()  # REF, PAGEREF, NOTEREF fields

    # Drawing elements
    IMAGE = auto()
    CHART = auto()
    SHAPE = auto()
    DIAGRAM = auto()  # SmartArt
    OLE_OBJECT = auto()  # OLE embedded objects
    VML = auto()  # Legacy VML graphics

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
    ElementType.HYPERLINK: "lnk",
    ElementType.CROSS_REFERENCE: "xref",
    ElementType.IMAGE: "img",
    ElementType.CHART: "chart",
    ElementType.SHAPE: "shape",
    ElementType.DIAGRAM: "diagram",
    ElementType.OLE_OBJECT: "obj",
    ElementType.VML: "vml",
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


class LinkType(Enum):
    """Types of hyperlinks in a document."""

    INTERNAL = auto()  # Link to bookmark within the document
    EXTERNAL = auto()  # Link to external URL


@dataclass
class BookmarkInfo:
    """Information about a bookmark in the document.

    Bookmarks are named locations in a document that can be referenced
    by hyperlinks and cross-references. They support bidirectional
    reference tracking.

    Attributes:
        ref: Reference to this bookmark (e.g., "bk:DefinitionsSection")
        name: Bookmark name identifier
        location: Ref of the element containing the bookmark start
        text_preview: Preview of the bookmarked text (first 100 chars)
        bookmark_id: Internal OOXML bookmark ID
        span_end_location: Optional ref where bookmark ends (for range bookmarks)
        referenced_by: List of refs that reference this bookmark (links, xrefs)
    """

    ref: str
    name: str
    location: str
    text_preview: str = ""
    bookmark_id: str | None = None
    span_end_location: str | None = None
    referenced_by: list[str] = field(default_factory=list)


@dataclass
class HyperlinkInfo:
    """Information about a hyperlink in the document.

    Hyperlinks can be internal (pointing to a bookmark) or external
    (pointing to a URL). Internal links support bidirectional tracking.

    Attributes:
        ref: Reference to this hyperlink (e.g., "lnk:0")
        link_type: Whether the link is internal or external
        from_location: Ref of the paragraph containing this link
        text: Display text of the hyperlink
        target: Target bookmark name (for internal) or URL (for external)
        target_location: For internal links, the resolved location ref
        anchor: Optional bookmark anchor within the target
        relationship_id: OOXML relationship ID (rId)
        is_broken: Whether this link points to a missing target
        error: Error message if link is broken
    """

    ref: str
    link_type: LinkType
    from_location: str
    text: str = ""
    target: str = ""
    target_location: str | None = None
    anchor: str | None = None
    relationship_id: str | None = None
    is_broken: bool = False
    error: str | None = None


class FieldType(Enum):
    """Types of cross-reference fields in Word."""

    REF = auto()  # Reference to bookmark text
    PAGEREF = auto()  # Reference to page number
    NOTEREF = auto()  # Reference to footnote/endnote number


@dataclass
class CrossReferenceInfo:
    """Information about a cross-reference field in the document.

    Cross-references are field codes (REF, PAGEREF, NOTEREF) that reference
    bookmarks and display dynamic content like text, page numbers, or
    footnote numbers.

    Attributes:
        ref: Reference ID for this cross-reference (e.g., "xref:5")
        field_type: Type of field (REF, PAGEREF, NOTEREF)
        target_bookmark: Name of the bookmark being referenced
        from_location: Ref of the paragraph containing this cross-reference
        display_value: Current displayed text (may be stale if is_dirty)
        is_dirty: Whether the field is marked for update by Word
        is_hyperlink: Whether the cross-reference is a clickable hyperlink (\\h switch)
        show_position: Whether to show "above"/"below" position (\\p switch)
        number_format: Number format: "full", "relative", "no_context", or None
        switches: Raw field switches string
        target_location: Resolved location of the target bookmark
        is_broken: Whether this references a non-existent bookmark
        error: Error message if the cross-reference is broken
    """

    ref: str
    field_type: FieldType
    target_bookmark: str
    from_location: str
    display_value: str = ""
    is_dirty: bool = False
    is_hyperlink: bool = False
    show_position: bool = False
    number_format: str | None = None
    switches: str = ""
    target_location: str | None = None
    is_broken: bool = False
    error: str | None = None


@dataclass
class ReferenceValidationResult:
    """Result of validating document references.

    This captures broken links, orphan bookmarks, cross-references, and other
    reference issues.

    Attributes:
        is_valid: Whether all references are valid
        broken_links: List of hyperlinks with invalid targets
        broken_cross_references: List of cross-references with invalid targets
        orphan_bookmarks: Bookmarks not referenced by any link or cross-reference
        missing_bookmarks: Bookmark names that are referenced but don't exist
        warnings: List of warning messages
    """

    is_valid: bool = True
    broken_links: list[HyperlinkInfo] = field(default_factory=list)
    broken_cross_references: list[CrossReferenceInfo] = field(default_factory=list)
    orphan_bookmarks: list[BookmarkInfo] = field(default_factory=list)
    missing_bookmarks: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


class ImageType(Enum):
    """Types of embedded images and graphics.

    These distinguish between different DrawingML element types.
    """

    IMAGE = auto()  # Standard image (w:drawing with pic:pic)
    CHART = auto()  # Chart (c:chart)
    DIAGRAM = auto()  # SmartArt/diagram (dgm:relIds)
    SHAPE = auto()  # AutoShape (wps:wsp)
    VML = auto()  # Legacy VML (w:pict)
    OLE_OBJECT = auto()  # OLE embedded object (w:object)


class ImagePositionType(Enum):
    """Position type for images."""

    INLINE = auto()  # Positioned inline with text (wp:inline)
    FLOATING = auto()  # Floating, anchored to page/paragraph (wp:anchor)


@dataclass
class ImageSize:
    """Image dimensions.

    Attributes:
        width_emu: Width in EMUs (English Metric Units)
        height_emu: Height in EMUs
    """

    width_emu: int
    height_emu: int

    # EMU per inch constant
    EMU_PER_INCH = 914400
    EMU_PER_CM = 360000

    @property
    def width_inches(self) -> float:
        """Get width in inches."""
        return self.width_emu / self.EMU_PER_INCH

    @property
    def height_inches(self) -> float:
        """Get height in inches."""
        return self.height_emu / self.EMU_PER_INCH

    @property
    def width_cm(self) -> float:
        """Get width in centimeters."""
        return self.width_emu / self.EMU_PER_CM

    @property
    def height_cm(self) -> float:
        """Get height in centimeters."""
        return self.height_emu / self.EMU_PER_CM

    def to_display_string(self) -> str:
        """Return human-readable size string."""
        return f"{self.width_inches:.1f}in x {self.height_inches:.1f}in"


@dataclass
class ImagePosition:
    """Position information for floating images.

    Attributes:
        horizontal: Horizontal position ("left", "center", "right", or offset)
        vertical: Vertical position ("top", "center", "bottom", or offset)
        relative_to: What the position is relative to ("page", "column", "paragraph", etc.)
        wrap_type: Text wrapping type ("none", "square", "tight", "topAndBottom", etc.)
    """

    horizontal: str = ""
    vertical: str = ""
    relative_to: str = ""
    wrap_type: str = ""


@dataclass
class ImageInfo:
    """Information about an embedded image or graphic.

    Attributes:
        ref: Reference path for this image (e.g., "img:5/0", "chart:12/0")
        image_type: Type of image (IMAGE, CHART, DIAGRAM, etc.)
        position_type: Whether inline or floating
        name: Display name for the image
        alt_text: Accessibility description
        size: Image dimensions
        format: Image format (png, jpeg, etc.) if known
        relationship_id: The r:embed or r:id value linking to the media file
        paragraph_ref: Ref of the paragraph containing this image
        position: Position info for floating images
        _element: Internal reference to the XML element
    """

    ref: str
    image_type: ImageType
    position_type: ImagePositionType
    name: str = ""
    alt_text: str = ""
    size: ImageSize | None = None
    format: str = ""
    relationship_id: str = ""
    paragraph_ref: str = ""
    position: ImagePosition | None = None
    _element: etree._Element | None = field(default=None, repr=False, compare=False)

    @property
    def is_inline(self) -> bool:
        """Check if image is inline."""
        return self.position_type == ImagePositionType.INLINE

    @property
    def is_floating(self) -> bool:
        """Check if image is floating."""
        return self.position_type == ImagePositionType.FLOATING

    def to_yaml_dict(self, mode: str = "content") -> dict:
        """Convert to YAML-friendly dictionary.

        Args:
            mode: "content" for basic info, "styling" for full details

        Returns:
            Dictionary suitable for YAML serialization
        """
        result: dict = {
            "ref": self.ref,
            "type": self.image_type.name.lower(),
            "position_type": "inline" if self.is_inline else "floating",
        }

        if self.name:
            result["name"] = self.name

        if self.alt_text:
            result["alt_text"] = self.alt_text

        if self.size:
            if mode == "styling":
                result["size"] = {
                    "width_emu": self.size.width_emu,
                    "height_emu": self.size.height_emu,
                }
            else:
                result["size"] = self.size.to_display_string()

        if mode == "styling":
            if self.format:
                result["format"] = self.format
            if self.relationship_id:
                result["relationship_id"] = self.relationship_id
            if self.position and self.is_floating:
                result["position"] = {
                    "horizontal": self.position.horizontal,
                    "vertical": self.position.vertical,
                    "relative_to": self.position.relative_to,
                    "wrap": self.position.wrap_type,
                }

        return result


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
        images: Images and embedded objects in this element
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
    images: list[ImageInfo] = field(default_factory=list)
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

    @property
    def has_images(self) -> bool:
        """Check if this node has images."""
        return len(self.images) > 0

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
