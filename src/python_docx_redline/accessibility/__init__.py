"""
Accessibility layer for Word documents.

This module provides a semantic accessibility layer inspired by browser ARIA trees,
enabling structured navigation and ref-based editing of Word documents.
"""

from .images import ImageExtractor, get_images_from_document
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
    "ChangeInfo",
    "ChangeType",
    "CommentInfo",
    "DetectedSection",
    "DetectionConfidence",
    "DetectionMetadata",
    "DetectionMethod",
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
    "Ref",
    "ReferenceValidationResult",
    "RefRegistry",
    "SectionDetectionConfig",
    "SectionDetector",
    "ViewMode",
    "create_section_nodes",
    "detect_sections",
    "get_images_from_document",
]
