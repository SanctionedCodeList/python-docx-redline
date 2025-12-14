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


@dataclass
class ComparisonStats:
    """Statistics from a document comparison operation.

    Provides counts of different types of tracked changes found in a document,
    typically after using compare_documents() or compare_to().

    Attributes:
        insertions: Number of text insertions (w:ins elements)
        deletions: Number of text deletions (w:del elements)
        moves: Number of move operations (w:moveFrom/w:moveTo pairs)
        format_changes: Number of formatting changes (w:rPrChange/w:pPrChange)
        total: Total number of all tracked changes

    Example:
        >>> redline = compare_documents("v1.docx", "v2.docx")
        >>> stats = redline.comparison_stats
        >>> print(f"Found {stats.insertions} insertions and {stats.deletions} deletions")
        >>> print(f"Total changes: {stats.total}")
    """

    insertions: int
    deletions: int
    moves: int = 0
    format_changes: int = 0

    @property
    def total(self) -> int:
        """Total number of tracked changes."""
        return self.insertions + self.deletions + self.moves + self.format_changes

    def __str__(self) -> str:
        """Get string representation of the statistics."""
        parts = []
        if self.insertions:
            parts.append(f"{self.insertions} insertion{'s' if self.insertions != 1 else ''}")
        if self.deletions:
            parts.append(f"{self.deletions} deletion{'s' if self.deletions != 1 else ''}")
        if self.moves:
            parts.append(f"{self.moves} move{'s' if self.moves != 1 else ''}")
        if self.format_changes:
            parts.append(
                f"{self.format_changes} format change{'s' if self.format_changes != 1 else ''}"
            )
        if not parts:
            return "No changes"
        return ", ".join(parts)


@dataclass
class FormatResult:
    """Result of a format operation.

    Attributes:
        success: Whether the operation completed without error
        changed: Whether any formatting changes were actually applied
        text_matched: The text that was formatted
        paragraph_index: Index of the affected paragraph (or -1 for multi-para)
        changes_applied: Dictionary of formatting changes applied
        previous_formatting: List of dicts, one per affected run, with previous values
        change_id: The w:id assigned to this tracked change (last one if multiple)
        runs_affected: Number of runs that were modified
    """

    success: bool
    changed: bool
    text_matched: str
    paragraph_index: int
    changes_applied: dict[str, object]
    previous_formatting: list[dict[str, object]]
    change_id: int
    runs_affected: int = 1

    def __str__(self) -> str:
        """Get string representation of the result."""
        if not self.success:
            return f"✗ Failed to format '{self.text_matched}'"
        if self.changed:
            changes = ", ".join(f"{k}={v}" for k, v in self.changes_applied.items())
            return f"✓ Formatted '{self.text_matched}': {changes}"
        return f"○ No change to '{self.text_matched}' (already formatted)"
