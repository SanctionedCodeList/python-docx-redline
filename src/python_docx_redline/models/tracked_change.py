"""
TrackedChange model class for representing tracked changes in Word documents.

Provides a Pythonic API for accessing tracked change metadata including
the change type, author, date, text content, and the underlying XML element.
"""

from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from typing import TYPE_CHECKING

from lxml import etree

from python_docx_redline.constants import WORD_NAMESPACE

if TYPE_CHECKING:
    from python_docx_redline.document import Document


class ChangeType(Enum):
    """Types of tracked changes in Word documents.

    Attributes:
        INSERTION: Text that was added (w:ins)
        DELETION: Text that was removed (w:del)
        MOVE_FROM: Source location of moved text (w:moveFrom)
        MOVE_TO: Destination location of moved text (w:moveTo)
        FORMAT_RUN: Run property change (w:rPrChange)
        FORMAT_PARAGRAPH: Paragraph property change (w:pPrChange)
    """

    INSERTION = "insertion"
    DELETION = "deletion"
    MOVE_FROM = "move_from"
    MOVE_TO = "move_to"
    FORMAT_RUN = "format_run"
    FORMAT_PARAGRAPH = "format_paragraph"


@dataclass
class TrackedChange:
    """Represents a single tracked change in a Word document.

    This class provides convenient access to tracked change metadata and
    supports operations like accepting or rejecting individual changes.

    Attributes:
        id: The change ID (w:id attribute value)
        change_type: Type of change (insertion, deletion, formatting, etc.)
        author: Author who made the change
        date: Timestamp when the change was made
        text: The text content of the change
        element: Reference to the underlying XML element

    Example:
        >>> changes = doc.get_tracked_changes()
        >>> for change in changes:
        ...     print(f"{change.id}: {change.change_type.value} by {change.author}")
        ...     print(f"  Text: {change.text[:50]}...")
    """

    id: str
    change_type: ChangeType
    author: str
    date: datetime | None
    text: str
    element: etree._Element
    _document: "Document | None" = None

    def __post_init__(self) -> None:
        """Post-initialization to ensure proper field types."""
        # Ensure id is always a string
        self.id = str(self.id)

    @classmethod
    def from_element(
        cls,
        element: etree._Element,
        change_type: ChangeType,
        document: "Document | None" = None,
    ) -> "TrackedChange":
        """Create a TrackedChange from an XML element.

        Args:
            element: The XML element (w:ins, w:del, w:rPrChange, etc.)
            change_type: The type of change
            document: Optional reference to parent Document

        Returns:
            TrackedChange instance with extracted metadata
        """
        # Extract attributes
        change_id = element.get(f"{{{WORD_NAMESPACE}}}id", "")
        author = element.get(f"{{{WORD_NAMESPACE}}}author", "")

        # Parse date
        date_str = element.get(f"{{{WORD_NAMESPACE}}}date")
        date: datetime | None = None
        if date_str:
            try:
                date = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
            except ValueError:
                pass

        # Extract text content based on change type
        text = cls._extract_text(element, change_type)

        return cls(
            id=change_id,
            change_type=change_type,
            author=author,
            date=date,
            text=text,
            element=element,
            _document=document,
        )

    @staticmethod
    def _extract_text(element: etree._Element, change_type: ChangeType) -> str:
        """Extract text content from a tracked change element.

        Args:
            element: The XML element
            change_type: The type of change

        Returns:
            The text content of the change
        """
        if change_type in (ChangeType.INSERTION, ChangeType.MOVE_TO):
            # Get text from w:t elements
            text_elements = element.findall(f".//{{{WORD_NAMESPACE}}}t")
            return "".join(elem.text or "" for elem in text_elements)

        elif change_type in (ChangeType.DELETION, ChangeType.MOVE_FROM):
            # Get text from w:delText elements
            text_elements = element.findall(f".//{{{WORD_NAMESPACE}}}delText")
            return "".join(elem.text or "" for elem in text_elements)

        elif change_type in (ChangeType.FORMAT_RUN, ChangeType.FORMAT_PARAGRAPH):
            # Format changes don't have text content themselves
            # We can try to get the text from the parent run/paragraph
            return ""

        return ""

    def accept(self) -> None:
        """Accept this tracked change.

        For insertions: keeps the inserted text, removes tracking markup.
        For deletions: removes the deleted text entirely.
        For formatting: keeps new formatting, removes change tracking.

        Raises:
            ValueError: If document reference is not available
        """
        if self._document is None:
            raise ValueError("Cannot accept change: no document reference")

        self._document.accept_change(self.id)

    def reject(self) -> None:
        """Reject this tracked change.

        For insertions: removes the inserted text entirely.
        For deletions: restores the deleted text.
        For formatting: reverts to previous formatting.

        Raises:
            ValueError: If document reference is not available
        """
        if self._document is None:
            raise ValueError("Cannot reject change: no document reference")

        self._document.reject_change(self.id)

    @property
    def is_insertion(self) -> bool:
        """Check if this is an insertion change."""
        return self.change_type == ChangeType.INSERTION

    @property
    def is_deletion(self) -> bool:
        """Check if this is a deletion change."""
        return self.change_type == ChangeType.DELETION

    @property
    def is_move(self) -> bool:
        """Check if this is a move change (from or to)."""
        return self.change_type in (ChangeType.MOVE_FROM, ChangeType.MOVE_TO)

    @property
    def is_format_change(self) -> bool:
        """Check if this is a formatting change."""
        return self.change_type in (ChangeType.FORMAT_RUN, ChangeType.FORMAT_PARAGRAPH)

    def __repr__(self) -> str:
        """String representation of the tracked change."""
        text_preview = self.text[:30] + "..." if len(self.text) > 30 else self.text
        return (
            f"<TrackedChange id={self.id} type={self.change_type.value} "
            f"author={self.author!r}: {text_preview!r}>"
        )

    def __eq__(self, other: object) -> bool:
        """Check equality based on change ID."""
        if not isinstance(other, TrackedChange):
            return NotImplemented
        return self.id == other.id

    def __hash__(self) -> int:
        """Hash based on change ID."""
        return hash(self.id)
