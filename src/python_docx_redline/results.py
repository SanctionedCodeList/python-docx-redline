"""
Result classes for document operations.

This module provides result types that track the success/failure of
document editing operations.
"""

from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .text_search import TextSpan


@dataclass
class EditResult:
    """Result of applying a single edit operation.

    Attributes:
        success: Whether the edit was applied successfully
        edit_type: Type of edit (e.g., "insert_tracked", "replace_tracked")
        message: Human-readable message about the result
        span: Optional TextSpan indicating where the edit was applied
        error: Optional exception that occurred during the edit
    """

    success: bool
    edit_type: str
    message: str
    span: "TextSpan | None" = None
    error: Exception | None = None

    def __str__(self) -> str:
        """Get string representation of the result."""
        status = "✓" if self.success else "✗"
        return f"{status} {self.edit_type}: {self.message}"


@dataclass
class AcceptResult:
    """Result of accepting tracked changes.

    Attributes:
        insertions: Number of insertions accepted
        deletions: Number of deletions accepted
    """

    insertions: int
    deletions: int

    def __str__(self) -> str:
        """Get string representation of the result."""
        return f"Accepted {self.insertions} insertions, {self.deletions} deletions"


@dataclass
class RejectResult:
    """Result of rejecting tracked changes.

    Attributes:
        insertions: Number of insertions rejected
        deletions: Number of deletions rejected
    """

    insertions: int
    deletions: int

    def __str__(self) -> str:
        """Get string representation of the result."""
        return f"Rejected {self.insertions} insertions, {self.deletions} deletions"
