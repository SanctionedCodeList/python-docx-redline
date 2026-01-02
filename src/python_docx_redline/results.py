"""
Result classes for document operations.

This module provides result types that track the success/failure of
document editing operations.
"""

from dataclasses import dataclass, field
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
        index: Index of this edit in the batch (0-indexed)
        old_text: The old/find text for the edit (if applicable)
        new_text: The new/replace text for the edit (if applicable)
        suggestions: List of suggested similar text matches (for failed edits)
    """

    success: bool
    edit_type: str
    message: str
    span: "TextSpan | None" = None
    error: Exception | None = None
    index: int = -1
    old_text: str | None = None
    new_text: str | None = None
    suggestions: list[str] = field(default_factory=list)

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


@dataclass
class BatchResult:
    """Result of applying multiple edits in batch mode.

    Provides organized access to succeeded and failed edits, along with
    summary statistics and pretty-print output.

    Attributes:
        succeeded: List of EditResult objects for successful edits
        failed: List of EditResult objects for failed edits
        dry_run: Whether this was a dry run (no actual changes made)

    Example:
        >>> results = doc.apply_edits(edits, continue_on_error=True)
        >>> print(results)
        # 3 edits applied
        # 1 edit failed:
        #   [2] "nonexistent phrase" -> TextNotFoundError
        #       Suggestions: "nonexistent phrases" (1 match)
        >>> print(results.summary)
        "3/4 edits applied successfully"
    """

    succeeded: list[EditResult] = field(default_factory=list)
    failed: list[EditResult] = field(default_factory=list)
    dry_run: bool = False

    @property
    def total(self) -> int:
        """Total number of edits processed."""
        return len(self.succeeded) + len(self.failed)

    @property
    def success_count(self) -> int:
        """Number of successful edits."""
        return len(self.succeeded)

    @property
    def failure_count(self) -> int:
        """Number of failed edits."""
        return len(self.failed)

    @property
    def all_succeeded(self) -> bool:
        """Whether all edits succeeded."""
        return len(self.failed) == 0 and len(self.succeeded) > 0

    @property
    def all_results(self) -> list[EditResult]:
        """All results in original order by index."""
        return sorted(self.succeeded + self.failed, key=lambda r: r.index)

    @property
    def summary(self) -> str:
        """Get a one-line summary of the batch result."""
        if self.dry_run:
            prefix = "(dry run) "
        else:
            prefix = ""
        if self.total == 0:
            return f"{prefix}No edits to apply"
        return f"{prefix}{self.success_count}/{self.total} edits applied successfully"

    def __str__(self) -> str:
        """Get pretty-printed representation of the result."""
        lines = []

        # Header
        if self.dry_run:
            lines.append("=== Batch Edit Results (DRY RUN) ===")
        else:
            lines.append("=== Batch Edit Results ===")

        # Success summary
        if self.succeeded:
            lines.append(f"+ {len(self.succeeded)} edit(s) applied:")
            for result in self.succeeded:
                old_text = f'"{result.old_text}"' if result.old_text else ""
                new_text = f' -> "{result.new_text}"' if result.new_text else ""
                lines.append(f"  [{result.index}] {old_text}{new_text}")

        # Failure summary
        if self.failed:
            lines.append(f"x {len(self.failed)} edit(s) failed:")
            for result in self.failed:
                old_text = f'"{result.old_text}"' if result.old_text else ""
                error_type = type(result.error).__name__ if result.error else "Error"
                lines.append(f"  [{result.index}] {old_text} -> {error_type}")
                if result.suggestions:
                    sug_str = ", ".join(f'"{s}"' for s in result.suggestions[:3])
                    lines.append(f"      Suggestions: {sug_str}")

        # Summary line
        lines.append("")
        lines.append(self.summary)

        return "\n".join(lines)

    def __repr__(self) -> str:
        """Get detailed representation for debugging."""
        return (
            f"BatchResult(succeeded={len(self.succeeded)}, "
            f"failed={len(self.failed)}, dry_run={self.dry_run})"
        )

    def __bool__(self) -> bool:
        """Return True if all edits succeeded."""
        return self.all_succeeded

    def __len__(self) -> int:
        """Return total number of edits processed."""
        return self.total

    def __iter__(self):
        """Iterate over all results in original order."""
        return iter(self.all_results)
