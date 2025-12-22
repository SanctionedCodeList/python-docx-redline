"""CriticMarkup parser for python-docx-redline.

This module parses CriticMarkup syntax into structured operations that can be
applied to Word documents as tracked changes.

CriticMarkup Syntax Reference:
    - Insertion: {++inserted text++}
    - Deletion: {--deleted text--}
    - Substitution: {~~old~>new~~}
    - Comment: {>>comment text<<}
    - Highlight: {==marked text==}
    - Highlight + Comment: {==text=={>>comment<<}}

See: http://criticmarkup.com/
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum


class OperationType(Enum):
    """Types of CriticMarkup operations."""

    INSERTION = "insertion"
    DELETION = "deletion"
    SUBSTITUTION = "substitution"
    COMMENT = "comment"
    HIGHLIGHT = "highlight"


@dataclass
class CriticOperation:
    """A single CriticMarkup operation.

    Attributes:
        type: The type of operation (insertion, deletion, etc.)
        text: Main text content (inserted text, deleted text, or highlighted text)
        replacement: For substitutions, the new text replacing the old
        comment: For comments and highlighted comments, the comment text
        position: Character position in the source text where this operation starts
        end_position: Character position where this operation ends
        context_before: Text appearing before this operation (for locating in document)
        context_after: Text appearing after this operation (for locating in document)
    """

    type: OperationType
    text: str
    replacement: str | None = None
    comment: str | None = None
    position: int = 0
    end_position: int = 0
    context_before: str = ""
    context_after: str = ""


# Regex patterns for CriticMarkup operations
# Order matters: more specific patterns should come first

# Highlight with nested comment: {==text=={>>comment<<}}
_HIGHLIGHT_COMMENT_PATTERN = re.compile(r"\{==(.+?)==\{>>(.+?)<<\}\}", re.DOTALL)

# Simple patterns for each operation type
_INSERTION_PATTERN = re.compile(r"\{\+\+(.+?)\+\+\}", re.DOTALL)
_DELETION_PATTERN = re.compile(r"\{--(.+?)--\}", re.DOTALL)
_SUBSTITUTION_PATTERN = re.compile(r"\{~~(.+?)~>(.+?)~~\}", re.DOTALL)
_COMMENT_PATTERN = re.compile(r"\{>>(.+?)<<\}", re.DOTALL)
_HIGHLIGHT_PATTERN = re.compile(r"\{==(.+?)==\}", re.DOTALL)


def parse_criticmarkup(text: str, context_chars: int = 50) -> list[CriticOperation]:
    """Parse CriticMarkup syntax into operations.

    Args:
        text: Markdown text with CriticMarkup annotations
        context_chars: Number of characters of context to extract before/after
            each operation (default: 50)

    Returns:
        List of CriticOperation objects in document order (sorted by position)

    Example:
        >>> ops = parse_criticmarkup("Hello {++world++}!")
        >>> len(ops)
        1
        >>> ops[0].type
        <OperationType.INSERTION: 'insertion'>
        >>> ops[0].text
        'world'
    """
    operations: list[CriticOperation] = []

    # Track which ranges we've already processed to avoid double-matching
    # (e.g., a comment inside a highlight shouldn't also match as standalone)
    processed_ranges: set[tuple[int, int]] = set()

    # 1. First, find highlight+comment combinations (most specific)
    for match in _HIGHLIGHT_COMMENT_PATTERN.finditer(text):
        start, end = match.start(), match.end()
        highlighted_text = match.group(1)
        comment_text = match.group(2)

        operations.append(
            CriticOperation(
                type=OperationType.HIGHLIGHT,
                text=highlighted_text,
                comment=comment_text,
                position=start,
                end_position=end,
                context_before=_extract_context_before(text, start, context_chars),
                context_after=_extract_context_after(text, end, context_chars),
            )
        )
        processed_ranges.add((start, end))

    # 2. Substitutions: {~~old~>new~~}
    for match in _SUBSTITUTION_PATTERN.finditer(text):
        start, end = match.start(), match.end()
        if _overlaps_processed(start, end, processed_ranges):
            continue

        old_text = match.group(1)
        new_text = match.group(2)

        operations.append(
            CriticOperation(
                type=OperationType.SUBSTITUTION,
                text=old_text,
                replacement=new_text,
                position=start,
                end_position=end,
                context_before=_extract_context_before(text, start, context_chars),
                context_after=_extract_context_after(text, end, context_chars),
            )
        )
        processed_ranges.add((start, end))

    # 3. Insertions: {++text++}
    for match in _INSERTION_PATTERN.finditer(text):
        start, end = match.start(), match.end()
        if _overlaps_processed(start, end, processed_ranges):
            continue

        operations.append(
            CriticOperation(
                type=OperationType.INSERTION,
                text=match.group(1),
                position=start,
                end_position=end,
                context_before=_extract_context_before(text, start, context_chars),
                context_after=_extract_context_after(text, end, context_chars),
            )
        )
        processed_ranges.add((start, end))

    # 4. Deletions: {--text--}
    for match in _DELETION_PATTERN.finditer(text):
        start, end = match.start(), match.end()
        if _overlaps_processed(start, end, processed_ranges):
            continue

        operations.append(
            CriticOperation(
                type=OperationType.DELETION,
                text=match.group(1),
                position=start,
                end_position=end,
                context_before=_extract_context_before(text, start, context_chars),
                context_after=_extract_context_after(text, end, context_chars),
            )
        )
        processed_ranges.add((start, end))

    # 5. Standalone comments: {>>text<<}
    for match in _COMMENT_PATTERN.finditer(text):
        start, end = match.start(), match.end()
        if _overlaps_processed(start, end, processed_ranges):
            continue

        operations.append(
            CriticOperation(
                type=OperationType.COMMENT,
                text="",  # Comments don't have main text, just comment content
                comment=match.group(1),
                position=start,
                end_position=end,
                context_before=_extract_context_before(text, start, context_chars),
                context_after=_extract_context_after(text, end, context_chars),
            )
        )
        processed_ranges.add((start, end))

    # 6. Standalone highlights: {==text==} (without nested comment)
    for match in _HIGHLIGHT_PATTERN.finditer(text):
        start, end = match.start(), match.end()
        if _overlaps_processed(start, end, processed_ranges):
            continue

        operations.append(
            CriticOperation(
                type=OperationType.HIGHLIGHT,
                text=match.group(1),
                position=start,
                end_position=end,
                context_before=_extract_context_before(text, start, context_chars),
                context_after=_extract_context_after(text, end, context_chars),
            )
        )
        processed_ranges.add((start, end))

    # Sort by position to maintain document order
    return sorted(operations, key=lambda op: op.position)


def _overlaps_processed(start: int, end: int, processed: set[tuple[int, int]]) -> bool:
    """Check if a range overlaps with any already-processed range."""
    for proc_start, proc_end in processed:
        # Check for any overlap
        if start < proc_end and end > proc_start:
            return True
    return False


def _extract_context_before(text: str, position: int, chars: int) -> str:
    """Extract context text before a position.

    Extracts up to `chars` characters before `position`, stopping at
    the beginning of the text. Strips CriticMarkup syntax from the context.
    """
    start = max(0, position - chars)
    context = text[start:position]
    # Strip any CriticMarkup from context to get clean text
    return strip_criticmarkup(context)


def _extract_context_after(text: str, position: int, chars: int) -> str:
    """Extract context text after a position.

    Extracts up to `chars` characters after `position`, stopping at
    the end of the text. Strips CriticMarkup syntax from the context.
    """
    end = min(len(text), position + chars)
    context = text[position:end]
    # Strip any CriticMarkup from context to get clean text
    return strip_criticmarkup(context)


def strip_criticmarkup(text: str) -> str:
    """Remove CriticMarkup syntax, keeping the resulting text.

    For insertions, keeps the inserted text.
    For deletions, removes the deleted text.
    For substitutions, keeps the new text.
    For comments, removes them entirely.
    For highlights, keeps the highlighted text.

    Args:
        text: Text potentially containing CriticMarkup

    Returns:
        Clean text with CriticMarkup resolved

    Example:
        >>> strip_criticmarkup("Hello {++world++}!")
        'Hello world!'
        >>> strip_criticmarkup("Say {--goodbye--}hello")
        'Say hello'
        >>> strip_criticmarkup("{~~old~>new~~}")
        'new'
    """
    result = text

    # Handle highlight+comment first (keep highlighted text, remove comment)
    result = _HIGHLIGHT_COMMENT_PATTERN.sub(r"\1", result)

    # Handle substitutions (keep new text)
    result = _SUBSTITUTION_PATTERN.sub(r"\2", result)

    # Handle insertions (keep inserted text)
    result = _INSERTION_PATTERN.sub(r"\1", result)

    # Handle deletions (remove deleted text)
    result = _DELETION_PATTERN.sub("", result)

    # Handle standalone comments (remove entirely)
    result = _COMMENT_PATTERN.sub("", result)

    # Handle standalone highlights (keep highlighted text)
    result = _HIGHLIGHT_PATTERN.sub(r"\1", result)

    return result


def render_criticmarkup(operations: list[CriticOperation], base_text: str) -> str:
    """Render operations back to CriticMarkup syntax.

    This is useful for round-tripping: parse a document, modify operations,
    then render back to CriticMarkup format.

    Args:
        operations: List of CriticOperation objects
        base_text: The base text to insert operations into

    Returns:
        Text with CriticMarkup syntax applied

    Note:
        This function assumes operations don't overlap and are positioned
        relative to the base_text. For complex scenarios, use the Document
        class methods instead.
    """
    # Sort operations by position in reverse order so we can insert
    # from end to beginning without invalidating positions
    sorted_ops = sorted(operations, key=lambda op: op.position, reverse=True)

    result = base_text

    for op in sorted_ops:
        markup = _operation_to_markup(op)

        if op.type == OperationType.INSERTION:
            # Insert at position
            result = result[: op.position] + markup + result[op.position :]
        elif op.type == OperationType.DELETION:
            # Replace the deleted text with markup
            # Need to find the text to replace
            text_end = op.position + len(op.text)
            result = result[: op.position] + markup + result[text_end:]
        elif op.type == OperationType.SUBSTITUTION:
            # Replace old text with markup
            text_end = op.position + len(op.text)
            result = result[: op.position] + markup + result[text_end:]
        elif op.type in (OperationType.COMMENT, OperationType.HIGHLIGHT):
            # For highlights, wrap the text; for comments, insert at position
            if op.text:
                text_end = op.position + len(op.text)
                result = result[: op.position] + markup + result[text_end:]
            else:
                result = result[: op.position] + markup + result[op.position :]

    return result


def _operation_to_markup(op: CriticOperation) -> str:
    """Convert a single operation to its CriticMarkup syntax."""
    if op.type == OperationType.INSERTION:
        return f"{{++{op.text}++}}"
    elif op.type == OperationType.DELETION:
        return f"{{--{op.text}--}}"
    elif op.type == OperationType.SUBSTITUTION:
        return f"{{~~{op.text}~>{op.replacement}~~}}"
    elif op.type == OperationType.COMMENT:
        return f"{{>>{op.comment}<<}}"
    elif op.type == OperationType.HIGHLIGHT:
        if op.comment:
            return f"{{=={op.text}=={{>>{op.comment}<<}}}}"
        else:
            return f"{{=={op.text}==}}"
    else:
        raise ValueError(f"Unknown operation type: {op.type}")
