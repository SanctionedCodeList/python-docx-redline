"""
Minimal editing (legal-style diffs) for tracked changes.

This module implements word-level diffing for tracked changes generation,
producing small, human-looking redlines instead of "delete entire old text +
insert entire new text" patterns.

The key components:
1. Tokenizer - splits text into words, whitespace, and punctuation tokens
2. EditHunk - represents a single tracked change operation
3. compute_minimal_hunks() - produces legal-style diff hunks from two texts
4. apply_minimal_edits() - applies hunks to OOXML paragraph structure
"""

import logging
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import TYPE_CHECKING, Any

from lxml import etree

from .constants import MAX_TRACKED_HUNKS_PER_PARAGRAPH, WORD_NAMESPACE

if TYPE_CHECKING:
    from .text_search import TextSpan

logger = logging.getLogger(__name__)

# Tokenizer pattern:
# - \s+ : whitespace runs
# - [\w]+(?:[''\\-][\w]+)* : word tokens, including hyphenated/apostrophe words
#   Apostrophe characters supported: ' (ASCII), ' (U+2019), ' (U+2018)
# - [^\w\s] : single punctuation characters
TOKENIZER_PATTERN = re.compile(r"(\s+|[\w]+(?:['\u2019\u2018\-][\w]+)*|[^\w\s])")


def tokenize(text: str) -> list[str]:
    """Tokenize text into words, whitespace, and punctuation tokens.

    Handles:
    - Whitespace runs preserved as single tokens
    - Hyphenated words as single tokens (e.g., "non-disclosure")
    - Apostrophe words as single tokens (e.g., "party's")
    - Punctuation as individual tokens

    Args:
        text: The text to tokenize

    Returns:
        List of tokens preserving the original text when joined
    """
    tokens = TOKENIZER_PATTERN.findall(text)
    return tokens


def is_whitespace_token(token: str) -> bool:
    """Check if a token is whitespace-only."""
    return token.isspace()


def is_punctuation_token(token: str) -> bool:
    """Check if a token is punctuation-only."""
    return bool(re.match(r"^[^\w\s]+$", token))


@dataclass
class EditHunk:
    """Represents a single edit operation (delete, insert, or replace).

    Attributes:
        orig_token_start: Starting token index in original text
        orig_token_end: Ending token index in original text (exclusive)
        delete_text: Text to delete (empty for pure insertions)
        insert_text: Text to insert (empty for pure deletions)
        is_whitespace_only: True if both delete and insert are whitespace-only
        is_punctuation_only: True if both delete and insert are punctuation-only
        adjacent_to_content_change: True if adjacent to a non-whitespace change
    """

    orig_token_start: int
    orig_token_end: int
    delete_text: str = ""
    insert_text: str = ""
    is_whitespace_only: bool = False
    is_punctuation_only: bool = False
    adjacent_to_content_change: bool = False

    # Character offsets (computed later when needed for OOXML application)
    char_start: int = -1
    char_end: int = -1


@dataclass
class MinimalDiffResult:
    """Result of computing minimal diff hunks.

    Attributes:
        hunks: List of edit hunks to apply
        fallback_required: True if coarse diff should be used instead
        fallback_reason: Explanation if fallback is required
    """

    hunks: list[EditHunk] = field(default_factory=list)
    fallback_required: bool = False
    fallback_reason: str = ""


def _classify_token_span(tokens: list[str]) -> tuple[bool, bool]:
    """Classify a span of tokens.

    Returns:
        Tuple of (is_whitespace_only, is_punctuation_only)
    """
    if not tokens:
        return True, True  # Empty spans are both

    all_whitespace = all(is_whitespace_token(t) for t in tokens)
    all_punct_or_ws = all(is_whitespace_token(t) or is_punctuation_token(t) for t in tokens)
    has_punctuation = any(is_punctuation_token(t) for t in tokens)

    is_ws_only = all_whitespace
    is_punct_only = all_punct_or_ws and has_punctuation and not all_whitespace

    return is_ws_only, is_punct_only


def compute_minimal_hunks(
    orig_text: str,
    new_text: str,
    max_hunks: int = MAX_TRACKED_HUNKS_PER_PARAGRAPH,
) -> MinimalDiffResult:
    """Compute minimal edit hunks between two texts.

    Implements legal-style diff rules:
    - R1: Word-level hunks (not character-level)
    - R2: Deletion then insertion ordering (handled by caller)
    - R3: Whitespace-only changes suppressed unless adjacent to content change
    - R4: Punctuation-only changes are allowed standalone
    - R5: Fallback if too many hunks

    Args:
        orig_text: Original paragraph text
        new_text: Modified paragraph text
        max_hunks: Maximum allowed hunks before fallback

    Returns:
        MinimalDiffResult with hunks or fallback indication
    """
    result = MinimalDiffResult()

    # Tokenize both texts
    orig_tokens = tokenize(orig_text)
    new_tokens = tokenize(new_text)

    # Compute token-level diff
    matcher = SequenceMatcher(None, orig_tokens, new_tokens, autojunk=False)
    opcodes = matcher.get_opcodes()

    # Build raw hunks from opcodes
    raw_hunks: list[EditHunk] = []

    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            continue

        delete_tokens = orig_tokens[i1:i2]
        insert_tokens = new_tokens[j1:j2]

        delete_text = "".join(delete_tokens)
        insert_text = "".join(insert_tokens)

        # Classify this hunk
        del_ws_only, del_punct_only = _classify_token_span(delete_tokens)
        ins_ws_only, ins_punct_only = _classify_token_span(insert_tokens)

        # Both sides must be whitespace-only for the hunk to be whitespace-only
        is_ws_only = del_ws_only and ins_ws_only
        # Punctuation-only if both sides are punctuation/ws and at least one has punctuation
        is_punct_only = (del_punct_only or del_ws_only) and (ins_punct_only or ins_ws_only)
        is_punct_only = is_punct_only and (del_punct_only or ins_punct_only)
        # But not if it's whitespace only
        is_punct_only = is_punct_only and not is_ws_only

        hunk = EditHunk(
            orig_token_start=i1,
            orig_token_end=i2,
            delete_text=delete_text,
            insert_text=insert_text,
            is_whitespace_only=is_ws_only,
            is_punctuation_only=is_punct_only,
        )
        raw_hunks.append(hunk)

    # R3: Mark whitespace-only hunks as adjacent if next to non-whitespace changes
    # A whitespace hunk is "adjacent" if it immediately precedes or follows
    # a non-whitespace hunk in the opcode stream
    for i, hunk in enumerate(raw_hunks):
        if hunk.is_whitespace_only:
            # Check previous hunk
            if i > 0 and not raw_hunks[i - 1].is_whitespace_only:
                hunk.adjacent_to_content_change = True
            # Check next hunk
            if i < len(raw_hunks) - 1 and not raw_hunks[i + 1].is_whitespace_only:
                hunk.adjacent_to_content_change = True

    # Filter hunks: suppress whitespace-only unless adjacent
    filtered_hunks = [
        h for h in raw_hunks if not h.is_whitespace_only or h.adjacent_to_content_change
    ]

    # R5: Check fragmentation limit
    if len(filtered_hunks) > max_hunks:
        result.fallback_required = True
        result.fallback_reason = f"Too many hunks ({len(filtered_hunks)} > {max_hunks})"
        return result

    # If only whitespace differs and all were suppressed, no changes needed
    if not filtered_hunks:
        # This is intentional per spec - whitespace-only paragraph differences
        # produce no tracked changes
        return result

    # Compute character offsets for each hunk
    _compute_char_offsets(filtered_hunks, orig_tokens)

    result.hunks = filtered_hunks
    return result


def _compute_char_offsets(hunks: list[EditHunk], orig_tokens: list[str]) -> None:
    """Compute character offsets for each hunk based on token positions.

    Args:
        hunks: List of hunks to update with character offsets
        orig_tokens: Original tokens for offset computation
    """
    # Build token start positions
    token_char_starts: list[int] = []
    pos = 0
    for token in orig_tokens:
        token_char_starts.append(pos)
        pos += len(token)
    # Append total length for end-of-text boundary
    token_char_starts.append(pos)

    for hunk in hunks:
        hunk.char_start = token_char_starts[hunk.orig_token_start]
        hunk.char_end = token_char_starts[hunk.orig_token_end]


def paragraph_has_tracked_revisions(paragraph: Any) -> bool:
    """Check if a paragraph already contains tracked revisions.

    Tracked revisions that make minimal editing unsafe:
    - w:ins (insertions)
    - w:del (deletions)
    - w:moveFrom (move source)
    - w:moveTo (move destination)

    Args:
        paragraph: The w:p XML element to check

    Returns:
        True if paragraph contains tracked revisions
    """
    tracked_tags = [
        f"{{{WORD_NAMESPACE}}}ins",
        f"{{{WORD_NAMESPACE}}}del",
        f"{{{WORD_NAMESPACE}}}moveFrom",
        f"{{{WORD_NAMESPACE}}}moveTo",
    ]

    for tag in tracked_tags:
        if paragraph.find(f".//{tag}") is not None:
            return True

    return False


def paragraph_has_unsupported_constructs(paragraph: Any) -> bool:
    """Check if a paragraph has constructs that make minimal editing unsafe.

    Unsupported for MVP:
    - w:fldChar (field characters)
    - w:hyperlink (hyperlinks)
    - w:sdt (structured document tags / content controls)
    - w:customXml (custom XML)
    - w:smartTag (smart tags that wrap runs)

    Args:
        paragraph: The w:p XML element to check

    Returns:
        True if paragraph has unsupported constructs
    """
    unsupported_tags = [
        f"{{{WORD_NAMESPACE}}}fldChar",
        f"{{{WORD_NAMESPACE}}}hyperlink",
        f"{{{WORD_NAMESPACE}}}sdt",
        f"{{{WORD_NAMESPACE}}}customXml",
        f"{{{WORD_NAMESPACE}}}smartTag",
    ]

    for tag in unsupported_tags:
        if paragraph.find(f".//{tag}") is not None:
            return True

    return False


def paragraph_has_nested_runs(paragraph: Any) -> bool:
    """Check if paragraph has runs that are not direct children.

    Our OOXML manipulation code uses list(paragraph).index(run) which
    assumes runs are direct children of the paragraph. If runs are nested
    inside wrapper elements (e.g., bookmarkStart/End ranges, comment ranges),
    this will fail or corrupt structure.

    Args:
        paragraph: The w:p XML element to check

    Returns:
        True if any runs are nested (not direct children)
    """
    # Get all descendant runs
    all_runs = paragraph.findall(f".//{{{WORD_NAMESPACE}}}r")

    # Get only direct child runs
    direct_runs = paragraph.findall(f"./{{{WORD_NAMESPACE}}}r")

    # If counts differ, some runs are nested
    return len(all_runs) != len(direct_runs)


def should_use_minimal_editing(
    orig_paragraph: Any,
    new_text: str,
    orig_text: str,
    max_hunks: int = MAX_TRACKED_HUNKS_PER_PARAGRAPH,
) -> tuple[bool, MinimalDiffResult, str]:
    """Determine if minimal editing should be used for a paragraph replacement.

    Checks safety conditions and computes diff to determine viability.

    Args:
        orig_paragraph: The original w:p XML element
        new_text: The new paragraph text
        orig_text: The original paragraph text
        max_hunks: Maximum hunks before fallback

    Returns:
        Tuple of (should_use_minimal, diff_result, reason_if_not)
    """
    # Check for existing tracked revisions
    if paragraph_has_tracked_revisions(orig_paragraph):
        return False, MinimalDiffResult(), "Paragraph has existing tracked revisions"

    # Check for unsupported constructs
    if paragraph_has_unsupported_constructs(orig_paragraph):
        return False, MinimalDiffResult(), "Paragraph has unsupported constructs"

    # Check for nested runs (runs inside wrapper elements)
    if paragraph_has_nested_runs(orig_paragraph):
        return False, MinimalDiffResult(), "Paragraph has nested runs"

    # Compute the diff
    diff_result = compute_minimal_hunks(orig_text, new_text, max_hunks)

    if diff_result.fallback_required:
        return False, diff_result, diff_result.fallback_reason

    return True, diff_result, ""


# --- OOXML Application Functions ---


@dataclass
class RunSpan:
    """Maps character positions to runs in a paragraph.

    Attributes:
        runs: List of w:r elements in the paragraph
        char_to_run: Maps character index to (run_index, offset_in_run)
        full_text: Concatenated text from all runs
    """

    runs: list[Any]
    char_to_run: list[tuple[int, int]]
    full_text: str


def build_paragraph_char_map(paragraph: Any) -> RunSpan:
    """Build a character-to-run map for a paragraph.

    Similar to TextSearch but for the full paragraph, mapping each
    character position to its containing run and offset.

    Args:
        paragraph: The w:p XML element

    Returns:
        RunSpan with character mapping
    """
    runs = list(paragraph.findall(f".//{{{WORD_NAMESPACE}}}r"))

    char_to_run: list[tuple[int, int]] = []
    full_text_chars: list[str] = []

    for run_idx, run in enumerate(runs):
        # Get text from w:t elements only
        text_elements = run.findall(f".//{{{WORD_NAMESPACE}}}t")
        run_text = "".join(elem.text or "" for elem in text_elements)

        for char_idx, char in enumerate(run_text):
            char_to_run.append((run_idx, char_idx))
            full_text_chars.append(char)

    return RunSpan(
        runs=runs,
        char_to_run=char_to_run,
        full_text="".join(full_text_chars),
    )


def _get_run_rpr(run: Any) -> Any | None:
    """Get the w:rPr (run properties) element from a run.

    Args:
        run: The w:r element

    Returns:
        The w:rPr element or None if not present
    """
    return run.find(f"{{{WORD_NAMESPACE}}}rPr")


def _clone_rpr(rpr: Any | None) -> Any | None:
    """Clone a w:rPr element.

    Args:
        rpr: The w:rPr element to clone (or None)

    Returns:
        A deep copy of the element or None
    """
    if rpr is None:
        return None
    from copy import deepcopy

    return deepcopy(rpr)


def _create_run_with_text(
    text: str,
    rpr: Any | None = None,
    is_deletion: bool = False,
    rsid_attr: str | None = None,
) -> Any:
    """Create a w:r element with text content.

    Args:
        text: The text content
        rpr: Optional w:rPr to include (will be cloned)
        is_deletion: If True, uses w:delText instead of w:t
        rsid_attr: Optional rsid attribute value

    Returns:
        A new w:r element
    """
    run = etree.Element(f"{{{WORD_NAMESPACE}}}r")

    if rsid_attr:
        if is_deletion:
            run.set(f"{{{WORD_NAMESPACE}}}rsidDel", rsid_attr)
        else:
            run.set(f"{{{WORD_NAMESPACE}}}rsidR", rsid_attr)

    if rpr is not None:
        cloned = _clone_rpr(rpr)
        if cloned is not None:
            run.append(cloned)

    text_tag = f"{{{WORD_NAMESPACE}}}delText" if is_deletion else f"{{{WORD_NAMESPACE}}}t"
    text_elem = etree.SubElement(run, text_tag)
    text_elem.text = text

    # Preserve whitespace if needed
    if text and (text[0].isspace() or text[-1].isspace()):
        text_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    return run


def _find_runs_for_span(
    char_start: int,
    char_end: int,
    char_to_run: list[tuple[int, int]],
) -> tuple[int, int, int, int]:
    """Find the run indices and offsets for a character span.

    Args:
        char_start: Starting character index
        char_end: Ending character index (exclusive)
        char_to_run: Character to run mapping

    Returns:
        Tuple of (start_run_idx, start_offset, end_run_idx, end_offset)
    """
    if char_start >= len(char_to_run):
        # Span is at/past end of text
        if char_to_run:
            last_run_idx, last_offset = char_to_run[-1]
            return last_run_idx, last_offset + 1, last_run_idx, last_offset + 1
        return 0, 0, 0, 0

    start_run_idx, start_offset = char_to_run[char_start]

    if char_end <= 0 or char_end > len(char_to_run):
        end_idx = len(char_to_run) - 1
    else:
        end_idx = char_end - 1

    end_run_idx, end_offset = char_to_run[end_idx]
    # Make end_offset exclusive
    end_offset += 1

    return start_run_idx, start_offset, end_run_idx, end_offset


def apply_minimal_edits_to_paragraph(
    paragraph: Any,
    hunks: list[EditHunk],
    xml_generator: Any,
    author: str | None = None,
) -> None:
    """Apply minimal edit hunks to a paragraph's OOXML structure.

    This is the core function that transforms the paragraph by:
    1. Splitting runs at hunk boundaries
    2. Wrapping deleted content in w:del
    3. Inserting w:ins for new content
    4. Preserving formatting from neighboring runs

    The key rule (R2) is deletion-then-insertion ordering.

    Args:
        paragraph: The w:p XML element to modify
        hunks: List of EditHunk objects to apply
        xml_generator: TrackedXMLGenerator for creating ins/del elements
        author: Author for tracked changes
    """
    if not hunks:
        return

    # Build character map
    run_span = build_paragraph_char_map(paragraph)

    if not run_span.runs:
        # No runs to edit - this shouldn't happen but handle gracefully
        return

    # Process hunks in reverse order to avoid position shifting
    for hunk in reversed(hunks):
        _apply_single_hunk(
            paragraph,
            hunk,
            run_span,
            xml_generator,
            author,
        )

        # Rebuild the character map after each hunk application
        # because run structure has changed
        run_span = build_paragraph_char_map(paragraph)


def _apply_single_hunk(
    paragraph: Any,
    hunk: EditHunk,
    run_span: RunSpan,
    xml_generator: Any,
    author: str | None,
) -> None:
    """Apply a single hunk to the paragraph.

    Args:
        paragraph: The w:p element
        hunk: The hunk to apply
        run_span: Current character-to-run mapping
        xml_generator: For creating ins/del XML
        author: Author for tracked changes
    """
    # Handle pure insertion (no deletion)
    if not hunk.delete_text and hunk.insert_text:
        _apply_insertion_only(paragraph, hunk, run_span, xml_generator, author)
        return

    # Handle deletion (with or without insertion)
    _apply_deletion_with_optional_insertion(paragraph, hunk, run_span, xml_generator, author)


def _apply_insertion_only(
    paragraph: Any,
    hunk: EditHunk,
    run_span: RunSpan,
    xml_generator: Any,
    author: str | None,
) -> None:
    """Apply a pure insertion (no deletion).

    The insertion is placed at the hunk's character position.
    """
    # Find insertion point
    if hunk.char_start >= len(run_span.char_to_run):
        # Insert at end of paragraph
        insert_after_run_idx = len(run_span.runs) - 1 if run_span.runs else -1
        insert_at_offset = None
    else:
        insert_after_run_idx, insert_at_offset = run_span.char_to_run[hunk.char_start]

    # Get formatting from nearest run
    rpr = None
    if run_span.runs:
        if insert_after_run_idx >= 0:
            rpr = _get_run_rpr(run_span.runs[insert_after_run_idx])
        elif insert_after_run_idx + 1 < len(run_span.runs):
            rpr = _get_run_rpr(run_span.runs[insert_after_run_idx + 1])

    # Create insertion element
    ins_xml = xml_generator.create_insertion(hunk.insert_text, author)
    ins_elem = _parse_tracked_xml(ins_xml)

    # Apply formatting to the run inside the insertion
    if rpr is not None and ins_elem is not None:
        inner_run = ins_elem.find(f".//{{{WORD_NAMESPACE}}}r")
        if inner_run is not None:
            # Insert cloned rPr at the beginning
            cloned_rpr = _clone_rpr(rpr)
            if cloned_rpr is not None:
                inner_run.insert(0, cloned_rpr)

    if ins_elem is None:
        return

    # Insert into paragraph
    if insert_at_offset is not None and insert_at_offset > 0:
        # Need to split the run - insertion point is in the middle of a run
        _split_run_and_insert(
            paragraph,
            run_span.runs,
            insert_after_run_idx,
            insert_at_offset,
            [ins_elem],
        )
    elif insert_at_offset == 0 and insert_after_run_idx >= 0:
        # Insert at the very beginning of a run - insert BEFORE the run
        ref_run = run_span.runs[insert_after_run_idx]
        ref_idx = list(paragraph).index(ref_run)
        paragraph.insert(ref_idx, ins_elem)
    elif insert_after_run_idx >= 0:
        # Insert after the run (offset is None, meaning end of paragraph)
        ref_run = run_span.runs[insert_after_run_idx]
        ref_idx = list(paragraph).index(ref_run)
        paragraph.insert(ref_idx + 1, ins_elem)
    else:
        # insert_after_run_idx == -1: Insert at beginning of paragraph, after w:pPr if present
        ppr = paragraph.find(f"{{{WORD_NAMESPACE}}}pPr")
        if ppr is not None:
            idx = list(paragraph).index(ppr) + 1
            paragraph.insert(idx, ins_elem)
        else:
            paragraph.insert(0, ins_elem)


def _apply_deletion_with_optional_insertion(
    paragraph: Any,
    hunk: EditHunk,
    run_span: RunSpan,
    xml_generator: Any,
    author: str | None,
) -> None:
    """Apply a deletion, optionally followed by an insertion.

    Per R2, deletion comes before insertion at the same location.
    """
    # Find the runs affected by the deletion
    start_run_idx, start_offset, end_run_idx, end_offset = _find_runs_for_span(
        hunk.char_start, hunk.char_end, run_span.char_to_run
    )

    # Get formatting from the first affected run
    rpr = None
    if start_run_idx < len(run_span.runs):
        rpr = _get_run_rpr(run_span.runs[start_run_idx])

    # Create deletion element
    del_xml = xml_generator.create_deletion(hunk.delete_text, author)
    del_elem = _parse_tracked_xml(del_xml)

    # Apply formatting to the run inside the deletion
    if rpr is not None and del_elem is not None:
        inner_run = del_elem.find(f".//{{{WORD_NAMESPACE}}}r")
        if inner_run is not None:
            cloned_rpr = _clone_rpr(rpr)
            if cloned_rpr is not None:
                inner_run.insert(0, cloned_rpr)

    # Create insertion element if needed
    ins_elem = None
    if hunk.insert_text:
        ins_xml = xml_generator.create_insertion(hunk.insert_text, author)
        ins_elem = _parse_tracked_xml(ins_xml)

        if rpr is not None and ins_elem is not None:
            inner_run = ins_elem.find(f".//{{{WORD_NAMESPACE}}}r")
            if inner_run is not None:
                cloned_rpr = _clone_rpr(rpr)
                if cloned_rpr is not None:
                    inner_run.insert(0, cloned_rpr)

    if del_elem is None:
        return

    # Build list of elements to insert (deletion first, then insertion per R2)
    elements_to_insert = [del_elem]
    if ins_elem is not None:
        elements_to_insert.append(ins_elem)

    # Apply the change to the paragraph
    _replace_span_in_paragraph(
        paragraph,
        run_span.runs,
        start_run_idx,
        start_offset,
        end_run_idx,
        end_offset,
        elements_to_insert,
    )


def _parse_tracked_xml(xml_str: str) -> Any | None:
    """Parse tracked change XML string to element.

    Args:
        xml_str: XML string (w:ins or w:del)

    Returns:
        Parsed lxml element or None on error
    """
    try:
        wrapper = (
            f'<root xmlns:w="{WORD_NAMESPACE}" '
            f'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" '
            f'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
            f"{xml_str}</root>"
        )
        root = etree.fromstring(wrapper.encode())
        # Return the first child (w:ins or w:del)
        return root[0] if len(root) > 0 else None
    except etree.XMLSyntaxError:
        return None


def _split_run_and_insert(
    paragraph: Any,
    runs: list[Any],
    run_idx: int,
    split_offset: int,
    elements: list[Any],
) -> None:
    """Split a run at a character offset and insert elements.

    Args:
        paragraph: The w:p element
        runs: List of runs
        run_idx: Index of run to split
        split_offset: Character offset within run to split at
        elements: Elements to insert at the split point
    """
    if run_idx >= len(runs):
        return

    run = runs[run_idx]
    rpr = _get_run_rpr(run)

    # Get the text from this run
    text_elements = run.findall(f".//{{{WORD_NAMESPACE}}}t")
    run_text = "".join(elem.text or "" for elem in text_elements)

    # Split the text
    before_text = run_text[:split_offset]
    after_text = run_text[split_offset:]

    # Find run position in paragraph
    run_pos = list(paragraph).index(run)

    # Remove original run
    paragraph.remove(run)

    # Create after-run first (we'll insert in reverse order)
    insert_pos = run_pos
    if after_text:
        after_run = _create_run_with_text(after_text, rpr)
        paragraph.insert(insert_pos, after_run)

    # Insert tracked change elements
    for elem in reversed(elements):
        paragraph.insert(insert_pos, elem)

    # Create before-run
    if before_text:
        before_run = _create_run_with_text(before_text, rpr)
        paragraph.insert(insert_pos, before_run)


def _replace_span_in_paragraph(
    paragraph: Any,
    runs: list[Any],
    start_run_idx: int,
    start_offset: int,
    end_run_idx: int,
    end_offset: int,
    replacement_elements: list[Any],
) -> None:
    """Replace a character span in a paragraph with new elements.

    Handles three cases:
    1. Span within single run (entire run) - replace the run
    2. Span within single run (partial) - split and replace
    3. Span across multiple runs - remove middle runs, split ends

    Args:
        paragraph: The w:p element
        runs: List of runs
        start_run_idx: Starting run index
        start_offset: Offset in starting run
        end_run_idx: Ending run index
        end_offset: Offset in ending run (exclusive)
        replacement_elements: Elements to insert (w:del, w:ins, etc.)
    """
    if start_run_idx >= len(runs):
        # Nothing to replace - just append elements
        for elem in replacement_elements:
            paragraph.append(elem)
        return

    # Clamp indices
    end_run_idx = min(end_run_idx, len(runs) - 1)

    # Get formatting from first run for potential new runs
    rpr = _get_run_rpr(runs[start_run_idx])

    # Case 1: Single run replacement
    if start_run_idx == end_run_idx:
        _replace_in_single_run(
            paragraph,
            runs,
            start_run_idx,
            start_offset,
            end_offset,
            replacement_elements,
            rpr,
        )
        return

    # Case 2: Multi-run replacement
    _replace_across_runs(
        paragraph,
        runs,
        start_run_idx,
        start_offset,
        end_run_idx,
        end_offset,
        replacement_elements,
        rpr,
    )


def _replace_in_single_run(
    paragraph: Any,
    runs: list[Any],
    run_idx: int,
    start_offset: int,
    end_offset: int,
    replacement_elements: list[Any],
    rpr: Any | None,
) -> None:
    """Replace a span within a single run."""
    run = runs[run_idx]

    # Get run text
    text_elements = run.findall(f".//{{{WORD_NAMESPACE}}}t")
    run_text = "".join(elem.text or "" for elem in text_elements)

    before_text = run_text[:start_offset]
    after_text = run_text[end_offset:]

    # Find run position
    run_pos = list(paragraph).index(run)

    # Remove original run
    paragraph.remove(run)

    # Build new content at this position
    insert_pos = run_pos

    # Insert after-text run if needed
    if after_text:
        after_run = _create_run_with_text(after_text, rpr)
        paragraph.insert(insert_pos, after_run)

    # Insert replacement elements
    for elem in reversed(replacement_elements):
        paragraph.insert(insert_pos, elem)

    # Insert before-text run if needed
    if before_text:
        before_run = _create_run_with_text(before_text, rpr)
        paragraph.insert(insert_pos, before_run)


def _replace_across_runs(
    paragraph: Any,
    runs: list[Any],
    start_run_idx: int,
    start_offset: int,
    end_run_idx: int,
    end_offset: int,
    replacement_elements: list[Any],
    rpr: Any | None,
) -> None:
    """Replace a span that crosses multiple runs."""
    # Get text from end run for after-text
    end_run = runs[end_run_idx]
    end_text_elements = end_run.findall(f".//{{{WORD_NAMESPACE}}}t")
    end_run_text = "".join(elem.text or "" for elem in end_text_elements)
    after_text = end_run_text[end_offset:]

    # Get text from start run for before-text
    start_run = runs[start_run_idx]
    start_text_elements = start_run.findall(f".//{{{WORD_NAMESPACE}}}t")
    start_run_text = "".join(elem.text or "" for elem in start_text_elements)
    before_text = start_run_text[:start_offset]

    # Find position of start run
    start_pos = list(paragraph).index(start_run)

    # Remove all runs in the span (in reverse order to maintain indices)
    for idx in range(end_run_idx, start_run_idx - 1, -1):
        if idx < len(runs):
            try:
                paragraph.remove(runs[idx])
            except ValueError:
                pass  # Run not in paragraph (shouldn't happen)

    # Insert new content at start position
    insert_pos = start_pos

    # Insert after-text run if needed
    if after_text:
        end_rpr = _get_run_rpr(end_run) or rpr
        after_run = _create_run_with_text(after_text, end_rpr)
        paragraph.insert(insert_pos, after_run)

    # Insert replacement elements
    for elem in reversed(replacement_elements):
        paragraph.insert(insert_pos, elem)

    # Insert before-text run if needed
    if before_text:
        before_run = _create_run_with_text(before_text, rpr)
        paragraph.insert(insert_pos, before_run)


# --- TextSpan-Aware Minimal Editing ---


def apply_minimal_edits_to_textspan(
    match: "TextSpan",
    replacement_text: str,
    xml_generator: Any,
    author: str | None = None,
) -> tuple[bool, str]:
    """Apply word-level minimal edits to a TextSpan replacement.

    This function adapts minimal diffing for TextSpan-based replace operations,
    which operate on a portion of a paragraph rather than the whole paragraph.

    The function computes word-level hunks between the matched text and
    replacement text, then applies those changes using delete-then-insert
    ordering for a human-looking redline.

    Args:
        match: The TextSpan being replaced (contains runs, offsets, paragraph)
        replacement_text: The new text to insert
        xml_generator: TrackedXMLGenerator for creating tracked change elements
        author: Author for tracked changes (optional)

    Returns:
        Tuple of (success: bool, reason: str).
        - If success is True, minimal edits were applied and reason is empty.
        - If success is False, reason explains why fallback is needed.

    Example:
        >>> success, reason = apply_minimal_edits_to_textspan(
        ...     match, "45 days", doc._xml_generator, "Editor"
        ... )
        >>> if not success:
        ...     logger.info("Falling back to coarse edit: %s", reason)
        ...     # Apply coarse delete + insert instead
    """
    paragraph = match.paragraph
    matched_text = match.text

    # Safety check 1: Existing tracked revisions
    if paragraph_has_tracked_revisions(paragraph):
        return False, "Paragraph has existing tracked revisions"

    # Safety check 2: Unsupported constructs
    if paragraph_has_unsupported_constructs(paragraph):
        return False, "Paragraph has unsupported constructs"

    # Safety check 3: Nested runs
    if paragraph_has_nested_runs(paragraph):
        return False, "Paragraph has nested runs"

    # Compute word-level hunks between matched text and replacement
    diff_result = compute_minimal_hunks(matched_text, replacement_text)

    if diff_result.fallback_required:
        return False, diff_result.fallback_reason

    if not diff_result.hunks:
        # No changes needed (whitespace-only diff was suppressed)
        # This is a successful minimal edit - just nothing to do
        return True, ""

    # Calculate the starting character position of the TextSpan within the paragraph
    # We need this to translate hunk positions (relative to matched_text)
    # to positions within the paragraph's full text
    span_start_char = _calculate_textspan_char_offset(match)

    # Adjust hunk character positions to be relative to paragraph
    for hunk in diff_result.hunks:
        hunk.char_start += span_start_char
        hunk.char_end += span_start_char

    # Apply the hunks to the paragraph
    apply_minimal_edits_to_paragraph(
        paragraph,
        diff_result.hunks,
        xml_generator,
        author,
    )

    return True, ""


def _calculate_textspan_char_offset(match: "TextSpan") -> int:
    """Calculate the character offset where a TextSpan starts within its paragraph.

    Args:
        match: The TextSpan to calculate offset for

    Returns:
        Character offset from the start of the paragraph to the start of the TextSpan
    """
    char_offset = 0

    # Count characters in runs before the start run
    for idx in range(match.start_run_index):
        run = match.runs[idx]
        text_elements = run.findall(f".//{{{WORD_NAMESPACE}}}t")
        run_text = "".join(elem.text or "" for elem in text_elements)
        char_offset += len(run_text)

    # Add the offset within the start run
    char_offset += match.start_offset

    return char_offset
