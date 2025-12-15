"""
Match class for representing text search results.

This module provides the Match dataclass returned by Document.find_all(),
which represents a single text match with location metadata and context.
"""

from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .text_search import TextSpan


@dataclass
class Match:
    """Represents a single text match with location metadata.

    This class is returned by Document.find_all() to provide rich information
    about where text was found in the document, including paragraph location,
    surrounding context, and the ability to access the underlying TextSpan.

    Attributes:
        index: Zero-based index of this match in the search results
        text: The matched text
        context: Surrounding text for disambiguation (up to 40 chars before/after)
        paragraph_index: Zero-based index of the paragraph containing this match
        paragraph_text: Full text of the paragraph containing this match
        location: Human-readable location string (e.g., "body" or "table:0:row:2:cell:1")
        span: The underlying TextSpan object for advanced operations

    Example:
        >>> matches = doc.find_all("production products")
        >>> for match in matches:
        ...     print(f"[{match.index}] {match.location}: {match.context}")
        [0] body: ...Therefore, production products utilizing...
        [1] table:0:row:45:cell:1: ...Therefore, production products utilizing Adeia's...
    """

    index: int
    text: str
    context: str
    paragraph_index: int
    paragraph_text: str
    location: str
    span: "TextSpan"

    def __repr__(self) -> str:
        """Return a detailed string representation."""
        return (
            f"Match(index={self.index}, text={self.text!r}, "
            f"location={self.location!r}, paragraph_index={self.paragraph_index})"
        )

    def __str__(self) -> str:
        """Return a user-friendly string representation."""
        # Truncate context for display if needed
        display_context = self.context
        if len(display_context) > 80:
            display_context = display_context[:77] + "..."

        return f"[{self.index}] {self.location}: {display_context}"
