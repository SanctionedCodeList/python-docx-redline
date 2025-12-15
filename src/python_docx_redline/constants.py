"""
Centralized constants for OOXML namespaces and other magic values.

This module consolidates all namespace URLs, namespace maps, and other constants
that were previously duplicated across multiple files. Import from here to ensure
consistency and make updates easier.
"""

# =============================================================================
# Word Processing Namespaces
# =============================================================================

# Main WordprocessingML namespace (Word 2007+)
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Word version-specific namespaces
W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"  # Word 2010
W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml"  # Word 2012
W16DU_NAMESPACE = "http://schemas.microsoft.com/office/word/2023/wordml"  # Word 2023


# =============================================================================
# DrawingML Namespaces
# =============================================================================

# DrawingML main namespace
A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Drawing picture namespace
PIC_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture"

# Word Processing Drawing namespace (inline/anchor positioning)
WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


# =============================================================================
# Package and Relationship Namespaces
# =============================================================================

# Open Packaging Convention namespaces
PACKAGE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"

# Office Document relationships
OFFICE_RELATIONSHIPS_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
RELATIONSHIP_NAMESPACE = OFFICE_RELATIONSHIPS_NAMESPACE  # Alias for compatibility

# Markup Compatibility namespace
MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# XML namespace
XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"


# =============================================================================
# Relationship Types
# =============================================================================

# Document part relationships
REL_TYPE_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_TYPE_COMMENTS_EXTENDED = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
REL_TYPE_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
REL_TYPE_COMMENTS_EXTENSIBLE = (
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
)
REL_TYPE_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
REL_TYPE_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"


# =============================================================================
# Namespace Maps
# =============================================================================

# Basic namespace map with just the main Word namespace
NSMAP = {"w": WORD_NAMESPACE}

# Extended namespace map including Word version namespaces
NSMAP_EXTENDED = {
    "w": WORD_NAMESPACE,
    "w14": W14_NAMESPACE,
    "w15": W15_NAMESPACE,
    "w16du": W16DU_NAMESPACE,
}

# Full namespace map for creating elements with all common namespaces
NSMAP_FULL = {
    "w": WORD_NAMESPACE,
    "w14": W14_NAMESPACE,
    "w15": W15_NAMESPACE,
    "w16du": W16DU_NAMESPACE,
    "r": OFFICE_RELATIONSHIPS_NAMESPACE,
    "mc": MC_NAMESPACE,
}

# DrawingML namespace map
NSMAP_DRAWING = {
    "a": A_NAMESPACE,
    "pic": PIC_NAMESPACE,
    "wp": WP_NAMESPACE,
    "r": OFFICE_RELATIONSHIPS_NAMESPACE,
}


# =============================================================================
# Default/Magic Numbers
# =============================================================================

# Text search context - characters to show around match
CONTEXT_CHARS_DEFAULT = 40

# Maximum iterations for batch operations to prevent infinite loops
MAX_BATCH_ITERATIONS = 100

# Maximum tracked hunks allowed per paragraph before fallback
MAX_TRACKED_HUNKS_PER_PARAGRAPH = 8


# =============================================================================
# Helper Functions
# =============================================================================


def w(tag: str) -> str:
    """Create a fully qualified Word namespace tag.

    Args:
        tag: Tag name without namespace prefix (e.g., "p", "r", "t")

    Returns:
        Fully qualified tag (e.g., "{http://...wordprocessingml/2006/main}p")

    Example:
        >>> w("p")
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    """
    return f"{{{WORD_NAMESPACE}}}{tag}"


def w14(tag: str) -> str:
    """Create a fully qualified Word 2010 namespace tag.

    Args:
        tag: Tag name without namespace prefix

    Returns:
        Fully qualified tag with w14 namespace
    """
    return f"{{{W14_NAMESPACE}}}{tag}"


def w15(tag: str) -> str:
    """Create a fully qualified Word 2012 namespace tag.

    Args:
        tag: Tag name without namespace prefix

    Returns:
        Fully qualified tag with w15 namespace
    """
    return f"{{{W15_NAMESPACE}}}{tag}"


def w16du(tag: str) -> str:
    """Create a fully qualified Word 2023 namespace tag.

    Args:
        tag: Tag name without namespace prefix

    Returns:
        Fully qualified tag with w16du namespace
    """
    return f"{{{W16DU_NAMESPACE}}}{tag}"


def a(tag: str) -> str:
    """Create a fully qualified DrawingML main namespace tag.

    Args:
        tag: Tag name without namespace prefix (e.g., "blip", "ext")

    Returns:
        Fully qualified tag with DrawingML main namespace
    """
    return f"{{{A_NAMESPACE}}}{tag}"


def pic(tag: str) -> str:
    """Create a fully qualified DrawingML picture namespace tag.

    Args:
        tag: Tag name without namespace prefix (e.g., "pic", "blipFill")

    Returns:
        Fully qualified tag with DrawingML picture namespace
    """
    return f"{{{PIC_NAMESPACE}}}{tag}"


def wp(tag: str) -> str:
    """Create a fully qualified Word Processing Drawing namespace tag.

    Args:
        tag: Tag name without namespace prefix (e.g., "inline", "extent")

    Returns:
        Fully qualified tag with WP drawing namespace
    """
    return f"{{{WP_NAMESPACE}}}{tag}"


def r(tag: str) -> str:
    """Create a fully qualified Office Relationships namespace tag.

    Args:
        tag: Tag name without namespace prefix (e.g., "embed")

    Returns:
        Fully qualified tag with relationship namespace
    """
    return f"{{{OFFICE_RELATIONSHIPS_NAMESPACE}}}{tag}"
