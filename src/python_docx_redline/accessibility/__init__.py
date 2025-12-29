"""
Accessibility layer for Word documents.

This module provides a semantic accessibility layer inspired by browser ARIA trees,
enabling structured navigation and ref-based editing of Word documents.
"""

from .bookmarks import BookmarkRegistry, add_bookmark, rename_bookmark
from .images import ImageExtractor, get_images_from_document
from .outline import (
    DocumentSizeInfo,
    OutlineTree,
    RefTree,
    SearchResult,
    SearchResults,
    SectionInfo,
    SectionTree,
    TableTree,
    estimate_tokens,
    truncate_to_token_budget,
)
from .outline import (
    SectionDetectionConfig as OutlineSectionDetectionConfig,
)
from .registry import RefRegistry
from .sections import (
    DetectedSection,
    DetectionConfidence,
    DetectionMetadata,
    DetectionMethod,
    HeuristicConfig,
    SectionDetectionConfig,
    SectionDetector,
    create_section_nodes,
    detect_sections,
)
from .tree import AccessibilityTree, DocumentStats
from .types import (
    AccessibilityNode,
    BookmarkInfo,
    ChangeInfo,
    ChangeType,
    CommentInfo,
    ElementType,
    HyperlinkInfo,
    ImageInfo,
    ImagePosition,
    ImagePositionType,
    ImageSize,
    ImageType,
    LinkType,
    Ref,
    ReferenceValidationResult,
    ViewMode,
)

__all__ = [
    "AccessibilityNode",
    "AccessibilityTree",
    "BookmarkInfo",
    "BookmarkRegistry",
    "ChangeInfo",
    "ChangeType",
    "CommentInfo",
    "DetectedSection",
    "DetectionConfidence",
    "DetectionMetadata",
    "DetectionMethod",
    "DocumentSizeInfo",
    "DocumentStats",
    "ElementType",
    "HeuristicConfig",
    "HyperlinkInfo",
    "ImageExtractor",
    "ImageInfo",
    "ImagePosition",
    "ImagePositionType",
    "ImageSize",
    "ImageType",
    "LinkType",
    "OutlineSectionDetectionConfig",
    "OutlineTree",
    "Ref",
    "ReferenceValidationResult",
    "RefRegistry",
    "RefTree",
    "SearchResult",
    "SearchResults",
    "SectionDetectionConfig",
    "SectionDetector",
    "SectionInfo",
    "SectionTree",
    "TableTree",
    "ViewMode",
    "add_bookmark",
    "create_section_nodes",
    "detect_sections",
    "estimate_tokens",
    "get_images_from_document",
    "rename_bookmark",
    "truncate_to_token_budget",
]
