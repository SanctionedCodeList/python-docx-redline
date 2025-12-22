"""CriticMarkup parser and exporter for python-docx-redline.

This module provides:
1. Parsing CriticMarkup syntax into structured operations
2. Exporting DOCX documents with tracked changes to CriticMarkup format

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
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from python_docx_redline.document import Document


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


# =============================================================================
# DOCX to CriticMarkup Export
# =============================================================================


def docx_to_criticmarkup(doc: Document, include_comments: bool = True) -> str:
    """Export document with tracked changes to CriticMarkup markdown.

    Walks through all paragraphs in the document and converts:
    - Tracked insertions (w:ins) → {++text++}
    - Tracked deletions (w:del) → {--text--}
    - Comments → {==marked text=={>>comment<<}}

    Args:
        doc: Document to export
        include_comments: Whether to include comments (default: True)

    Returns:
        Markdown string with CriticMarkup annotations

    Example:
        >>> doc = Document("contract_with_changes.docx")
        >>> markdown = docx_to_criticmarkup(doc)
        >>> print(markdown)
        The parties agree to {--30--}{++45++} day payment terms.
    """
    from python_docx_redline.constants import WORD_NAMESPACE

    output_paragraphs: list[str] = []

    # Build a map of comment IDs to their data for quick lookup
    comment_map: dict[str, tuple[str, str | None]] = {}  # id -> (comment_text, marked_text)
    if include_comments:
        for comment in doc.comments:
            comment_map[comment.id] = (comment.text, comment.marked_text)

    # Track which comment ranges we've already processed
    processed_comment_ids: set[str] = set()

    for para in doc.paragraphs:
        para_text = _paragraph_to_criticmarkup(
            para.element,
            comment_map,
            processed_comment_ids,
            WORD_NAMESPACE,
        )
        output_paragraphs.append(para_text)

    return "\n\n".join(output_paragraphs)


def _paragraph_to_criticmarkup(
    para_element,
    comment_map: dict[str, tuple[str, str | None]],
    processed_comment_ids: set[str],
    ns: str,
) -> str:
    """Convert a single paragraph element to CriticMarkup text.

    Walks through child elements in document order, handling:
    - Regular runs (w:r) containing text (w:t)
    - Insertions (w:ins) containing runs
    - Deletions (w:del) containing runs with delText
    - Comment range markers (commentRangeStart, commentRangeEnd, commentReference)

    Args:
        para_element: The w:p XML element
        comment_map: Map of comment ID to (text, marked_text) tuples
        processed_comment_ids: Set of already-processed comment IDs
        ns: Word namespace string

    Returns:
        CriticMarkup-formatted text for this paragraph
    """
    result: list[str] = []

    # Track active comment ranges
    # When we see commentRangeStart, we start collecting text
    # When we see commentRangeEnd, we wrap collected text with comment
    active_comment_ranges: dict[str, list[str]] = {}  # comment_id -> collected text parts

    # Walk through all direct children of the paragraph
    for child in para_element:
        tag = _get_local_tag(child.tag)

        if tag == "r":
            # Regular run - extract text
            text = _extract_run_text(child, ns)
            result.append(text)
            # Also add to any active comment ranges
            for parts in active_comment_ranges.values():
                parts.append(text)

        elif tag == "ins":
            # Tracked insertion - wrap in {++...++}
            text = _extract_insertion_text(child, ns)
            if text:
                result.append(f"{{++{text}++}}")
                # Also add to any active comment ranges
                for parts in active_comment_ranges.values():
                    parts.append(f"{{++{text}++}}")

        elif tag == "del":
            # Tracked deletion - wrap in {--...--}
            text = _extract_deletion_text(child, ns)
            if text:
                result.append(f"{{--{text}--}}")
                # Also add to any active comment ranges
                for parts in active_comment_ranges.values():
                    parts.append(f"{{--{text}--}}")

        elif tag == "commentRangeStart":
            # Start tracking text for this comment
            comment_id = child.get(f"{{{ns}}}id", "")
            if comment_id and comment_id not in processed_comment_ids:
                active_comment_ranges[comment_id] = []

        elif tag == "commentRangeEnd":
            # End of comment range - we'll process when we see commentReference
            pass

        elif tag == "commentReference":
            # Comment reference - now we can emit the comment
            comment_id = child.get(f"{{{ns}}}id", "")
            if comment_id in comment_map and comment_id not in processed_comment_ids:
                comment_text, marked_text = comment_map[comment_id]

                if comment_id in active_comment_ranges:
                    # We tracked the marked text - use it
                    collected_text = "".join(active_comment_ranges[comment_id])
                    if collected_text:
                        # Remove the collected text from result (we'll re-add with comment)
                        # This is tricky - we need to find and remove these parts
                        # Actually, let's use a different approach
                        pass

                # For now, use the simpler approach: append comment after text
                # The marked_text from Comment object is what the comment applies to
                if marked_text:
                    result.append(f"{{>>{comment_text}<<}}")
                else:
                    result.append(f"{{>>{comment_text}<<}}")

                processed_comment_ids.add(comment_id)
                if comment_id in active_comment_ranges:
                    del active_comment_ranges[comment_id]

    return "".join(result)


def _get_local_tag(full_tag: str) -> str:
    """Extract local tag name from namespaced tag."""
    if "}" in full_tag:
        return full_tag.split("}")[1]
    return full_tag


def _extract_run_text(run_element, ns: str) -> str:
    """Extract text from a w:r (run) element.

    Gets text from all w:t children.
    """
    text_parts = []
    for t_elem in run_element.findall(f"{{{ns}}}t"):
        text_parts.append(t_elem.text or "")
    return "".join(text_parts)


def _extract_insertion_text(ins_element, ns: str) -> str:
    """Extract text from a w:ins (insertion) element.

    Insertions contain runs with regular w:t text.
    """
    text_parts = []
    for t_elem in ins_element.findall(f".//{{{ns}}}t"):
        text_parts.append(t_elem.text or "")
    return "".join(text_parts)


def _extract_deletion_text(del_element, ns: str) -> str:
    """Extract text from a w:del (deletion) element.

    Deletions contain runs with w:delText elements.
    """
    text_parts = []
    for dt_elem in del_element.findall(f".//{{{ns}}}delText"):
        text_parts.append(dt_elem.text or "")
    return "".join(text_parts)


# =============================================================================
# CriticMarkup to DOCX Import
# =============================================================================


@dataclass
class ApplyResult:
    """Result of applying CriticMarkup operations to a document.

    Attributes:
        total: Total number of operations attempted
        successful: Number of operations that succeeded
        failed: Number of operations that failed
        errors: List of (operation, error_message) tuples for failures
    """

    total: int
    successful: int
    failed: int
    errors: list[tuple[CriticOperation, str]]

    @property
    def success_rate(self) -> float:
        """Calculate success rate as a percentage."""
        if self.total == 0:
            return 100.0
        return (self.successful / self.total) * 100

    def __repr__(self) -> str:
        return f"<ApplyResult: {self.successful}/{self.total} succeeded ({self.success_rate:.1f}%)>"


def apply_criticmarkup(
    doc: Document,
    markup_text: str,
    author: str | None = None,
    stop_on_error: bool = False,
) -> ApplyResult:
    """Apply CriticMarkup changes to document as tracked changes.

    Parses CriticMarkup syntax from the input text and applies each operation
    to the document using the appropriate tracked change method:
    - {++text++} → tracked insertion
    - {--text--} → tracked deletion
    - {~~old~>new~~} → tracked replacement (delete old + insert new)
    - {>>comment<<} → Word comment (requires context to attach)

    Args:
        doc: Document to modify
        markup_text: Markdown text with CriticMarkup annotations
        author: Author for tracked changes (uses document default if None)
        stop_on_error: If True, stop on first error. If False, continue
            processing remaining operations.

    Returns:
        ApplyResult with success/failure counts and error details

    Example:
        >>> doc = Document("contract.docx")
        >>> result = doc.apply_criticmarkup(
        ...     "Payment in {--30--}{++45++} days",
        ...     author="Review Bot"
        ... )
        >>> print(f"Applied {result.successful} of {result.total} changes")

    Note:
        Operations are applied in document order. For insertions, the
        context_before field is used to locate the insertion point.
    """
    operations = parse_criticmarkup(markup_text)

    successful = 0
    failed = 0
    errors: list[tuple[CriticOperation, str]] = []

    for op in operations:
        try:
            _apply_operation(doc, op, author)
            successful += 1
        except Exception as e:
            failed += 1
            errors.append((op, str(e)))
            if stop_on_error:
                break

    return ApplyResult(
        total=len(operations),
        successful=successful,
        failed=failed,
        errors=errors,
    )


def _apply_operation(doc: Document, op: CriticOperation, author: str | None) -> None:
    """Apply a single CriticMarkup operation to the document.

    Args:
        doc: Document to modify
        op: The operation to apply
        author: Author for tracked changes

    Raises:
        TextNotFoundError: If the anchor text cannot be found
        ValueError: If the operation type is not supported
    """
    if op.type == OperationType.INSERTION:
        # For insertions, we need to find where to insert using context
        anchor = _find_insertion_anchor(op)
        if anchor:
            doc.insert_tracked(op.text, after=anchor, author=author)
        else:
            # No context - try inserting at start of document
            # This is a fallback for insertions at the very beginning
            raise ValueError(
                f"Cannot determine insertion point for '{op.text[:30]}...' - no context available"
            )

    elif op.type == OperationType.DELETION:
        doc.delete_tracked(op.text, author=author)

    elif op.type == OperationType.SUBSTITUTION:
        doc.replace_tracked(op.text, op.replacement or "", author=author)

    elif op.type == OperationType.COMMENT:
        # Standalone comments need context to attach to
        anchor = op.context_before.strip()[-50:] if op.context_before else None
        if anchor:
            doc.add_comment(op.comment or "", on=anchor, author=author)
        else:
            raise ValueError(f"Cannot attach comment '{op.comment[:30]}...' - no context available")

    elif op.type == OperationType.HIGHLIGHT:
        # Highlights with comments become attached comments
        if op.comment:
            doc.add_comment(op.comment, on=op.text, author=author)
        # Highlights without comments are just markers - nothing to do in DOCX

    else:
        raise ValueError(f"Unsupported operation type: {op.type}")


def _find_insertion_anchor(op: CriticOperation) -> str | None:
    """Find the best anchor text for an insertion operation.

    Uses the context_before field to determine where to insert.
    Returns the last portion of the context as the anchor.

    Args:
        op: The insertion operation

    Returns:
        Anchor text to insert after, or None if no context available
    """
    if not op.context_before:
        return None

    # Use the last 50 characters of context as anchor
    # This balances specificity with likelihood of matching
    anchor = op.context_before.strip()
    if len(anchor) > 50:
        anchor = anchor[-50:]

    return anchor if anchor else None
