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
