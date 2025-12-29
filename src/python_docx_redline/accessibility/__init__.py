"""
Accessibility layer for Word documents.

This module provides a semantic accessibility layer inspired by browser ARIA trees,
enabling structured navigation and ref-based editing of Word documents.
"""

from .registry import RefRegistry
from .types import (
    AccessibilityNode,
    ChangeInfo,
    ChangeType,
    CommentInfo,
    ElementType,
    Ref,
    ViewMode,
)

__all__ = [
    "AccessibilityNode",
    "ChangeInfo",
    "ChangeType",
    "CommentInfo",
    "ElementType",
    "Ref",
    "RefRegistry",
    "ViewMode",
]
