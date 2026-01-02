"""
TrackedChangeOperations class for handling insert, delete, replace, move with tracking.

This module provides a dedicated class for all tracked change operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

from lxml import etree

from ..constants import WORD_NAMESPACE
from ..errors import AmbiguousTextError, TextNotFoundError
from ..scope import ScopeEvaluator
from ..suggestions import SuggestionGenerator

if TYPE_CHECKING:
    from ..document import Document
    from ..text_search import TextSpan

logger = logging.getLogger(__name__)


class TrackedChangeOperations:
    """Handles insert, delete, replace, and move operations with tracking.

    This class encapsulates all tracked change functionality, including:
    - Inserting text with tracking (after/before anchor)
    - Deleting text with tracking
    - Replacing text with tracking (delete old + insert new)
    - Moving text with tracking (linked source/destination markers)

    The class takes a Document reference and operates on its XML structure.

    Example:
        >>> # Usually accessed through Document
        >>> doc = Document("contract.docx")
        >>> doc.insert_tracked("new text", after="Section 2.1")
        >>> doc.delete_tracked("old clause")
        >>> doc.replace_tracked("30 days", "45 days")
        >>> doc.move_tracked("Section A", after="Table of Contents")
    """

    def __init__(self, document: Document) -> None:
        """Initialize TrackedChangeOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    def _record_change_ids(self, start_id: int, end_id: int) -> None:
        """Record change IDs with the edit group registry if a group is active.

        Args:
            start_id: The first change ID created (inclusive)
            end_id: The next_change_id after the operation (exclusive)
        """
        # Check if the document has an edit groups registry with an active group
        if hasattr(self._document, "_edit_groups_instance"):
            edit_groups = self._document._edit_groups_instance
            if edit_groups.active_group is not None:
                for change_id in range(start_id, end_id):
                    edit_groups.add_change_id(change_id)

    def _select_matches(
        self, matches: list[TextSpan], occurrence: int | list[int] | str, text: str
    ) -> list[TextSpan]:
        """Select target matches based on occurrence parameter.

        Args:
            matches: List of all matches found
            occurrence: Which occurrence(s) to select - int (1-indexed), list of ints, or string
            text: Original search text (for error messages)

        Returns:
            List of selected TextSpan matches

        Raises:
            AmbiguousTextError: If multiple matches and occurrence not specified
            ValueError: If occurrence is out of range
        """
        if occurrence == "first" or occurrence == 1:
            return [matches[0]]
        elif occurrence == "last":
            return [matches[-1]]
        elif occurrence == "all":
            return matches
        elif isinstance(occurrence, list):
            # Handle list of indices (1-indexed)
            selected = []
            for idx in occurrence:
                if not isinstance(idx, int):
                    raise ValueError(f"List elements must be integers, got {type(idx)}")
                if not (1 <= idx <= len(matches)):
                    raise ValueError(f"Occurrence {idx} out of range (1-{len(matches)})")
                selected.append(matches[idx - 1])
            return selected
        elif isinstance(occurrence, int) and 1 <= occurrence <= len(matches):
            return [matches[occurrence - 1]]
        elif isinstance(occurrence, int):
            raise ValueError(f"Occurrence {occurrence} out of range (1-{len(matches)})")
        elif len(matches) > 1:
            raise AmbiguousTextError(text, matches)
        else:
            return matches

    def _find_unique_match(
        self,
        text: str,
        scope: str | dict | Any | None,
        regex: bool,
        normalize_special_chars: bool,
    ) -> TextSpan:
        """Find a unique text match in the document.

        Args:
            text: The text or regex pattern to find
            scope: Limit search scope
            regex: Whether to treat text as regex
            normalize_special_chars: Whether to normalize quotes

        Returns:
            The single TextSpan match

        Raises:
            TextNotFoundError: If text is not found
            AmbiguousTextError: If multiple matches found
        """
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(
            text,
            paragraphs,
            regex=regex,
            normalize_special_chars=normalize_special_chars and not regex,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(text, paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(text, matches)

        return matches[0]

    def _parse_xml_elements(self, xml_content: str) -> list:
        """Parse XML content into lxml elements with proper namespaces.

        Args:
            xml_content: The XML string(s) to parse (can be multiple fragments)

        Returns:
            List of parsed lxml Elements
        """
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {xml_content}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        return list(root)

    def insert(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        normalize_special_chars: bool = True,
        fuzzy: float | dict[str, Any] | None = None,
        track: bool = False,
    ) -> None:
        """Insert text after or before a specific location.

        This method searches for the anchor text in the document and inserts
        the new text either immediately after it or immediately before it.

        Args:
            text: The text to insert (supports markdown formatting: **bold**, *italic*,
                ++underline++, ~~strikethrough~~)
            after: The text or regex pattern to insert after (optional)
            before: The text or regex pattern to insert before (optional)
            author: Optional author override (uses document author if None). Ignored
                when track=False.
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            occurrence: Which occurrence(s) to insert at: 1 (first), 2 (second), "first",
                "last", "all", or list of indices [1, 3, 5] (default: "first")
            regex: Whether to treat anchor as a regex pattern (default: False)
            normalize_special_chars: Auto-convert straight quotes to smart quotes for
                matching (default: True)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'
            track: If True, insert as tracked change (w:ins wrapper). If False, insert
                as plain text without tracking (default: False).

        Raises:
            ValueError: If both 'after' and 'before' are specified, or if neither is specified
            TextNotFoundError: If the anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found and
                occurrence not specified
            re.error: If regex=True and the pattern is invalid
            ImportError: If fuzzy matching requested but rapidfuzz not installed
        """
        # Validate parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        # Determine anchor text and insertion mode
        anchor: str = after if after is not None else before  # type: ignore[assignment]
        insert_after = after is not None

        # Find all matches
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Parse fuzzy configuration if provided
        from ..fuzzy import parse_fuzzy_config

        fuzzy_config = parse_fuzzy_config(fuzzy)

        matches = self._document._text_search.find_text(
            anchor,
            paragraphs,
            regex=regex,
            normalize_special_chars=normalize_special_chars and not regex and not fuzzy_config,
            fuzzy=fuzzy_config,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(anchor, paragraphs)
            raise TextNotFoundError(anchor, suggestions=suggestions)

        # Select target matches based on occurrence
        target_matches = self._select_matches(matches, occurrence, anchor)

        # Insert at each target match (process in reverse to preserve indices)
        for match in reversed(target_matches):
            if track:
                # Capture change ID before operation for edit group tracking
                start_id = self._document._xml_generator.next_change_id

                # Tracked insertion: wrap in <w:ins>
                insertion_xml = self._document._xml_generator.create_insertion(text, author)
                elements = self._parse_xml_elements(insertion_xml)
                insertion_element = elements[0]

                # Record change IDs with edit group registry
                self._record_change_ids(start_id, self._document._xml_generator.next_change_id)
            else:
                # Untracked insertion: plain runs
                # Get source run for formatting if available
                source_run = match.runs[0] if match.runs else None
                plain_runs = self._document._xml_generator.create_plain_runs(
                    text, source_run=source_run
                )
                # Use list of runs (could be multiple for markdown formatting)
                insertion_element = plain_runs

            if insert_after:
                self._insert_after_match(match, insertion_element)
            else:
                self._insert_before_match(match, insertion_element)

    def delete(
        self,
        text: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        normalize_special_chars: bool = True,
        fuzzy: float | dict[str, Any] | None = None,
        track: bool = False,
    ) -> None:
        """Delete text from the document.

        This method searches for the specified text in the document and removes it.

        Args:
            text: The text or regex pattern to delete
            author: Optional author override (uses document author if None). Ignored
                when track=False.
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            occurrence: Which occurrence(s) to delete: 1 (first), 2 (second), "first", "last",
                "all", or list of indices [1, 3, 5] (default: "first")
            regex: Whether to treat 'text' as a regex pattern (default: False)
            normalize_special_chars: Auto-convert straight quotes to smart quotes for
                matching (default: True)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'
            track: If True, delete as tracked change (w:del wrapper). If False, remove
                text without tracking (default: False).

        Raises:
            TextNotFoundError: If the text is not found
            AmbiguousTextError: If multiple occurrences of text are found and
                occurrence not specified
            re.error: If regex=True and the pattern is invalid
            ImportError: If fuzzy matching requested but rapidfuzz not installed
        """
        # Find all matches
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Parse fuzzy configuration if provided
        from ..fuzzy import parse_fuzzy_config

        fuzzy_config = parse_fuzzy_config(fuzzy)

        matches = self._document._text_search.find_text(
            text,
            paragraphs,
            regex=regex,
            normalize_special_chars=normalize_special_chars and not regex and not fuzzy_config,
            fuzzy=fuzzy_config,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(text, paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        # Select target matches based on occurrence
        target_matches = self._select_matches(matches, occurrence, text)

        # Delete each target match (process in reverse to preserve indices)
        for match in reversed(target_matches):
            if track:
                # Capture change ID before operation for edit group tracking
                start_id = self._document._xml_generator.next_change_id

                # Tracked deletion: wrap in <w:del>
                deletion_xml = self._document._xml_generator.create_deletion(match.text, author)
                elements = self._parse_xml_elements(deletion_xml)
                deletion_element = elements[0]
                # Replace the matched text with deletion
                self._replace_match_with_element(match, deletion_element)

                # Record change IDs with edit group registry
                self._record_change_ids(start_id, self._document._xml_generator.next_change_id)
            else:
                # Untracked deletion: simply remove the matched runs
                self._remove_match(match)

    def replace(
        self,
        find: str,
        replace: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        normalize_special_chars: bool = True,
        show_context: bool = False,
        check_continuity: bool = False,
        context_chars: int = 50,
        fuzzy: float | dict[str, Any] | None = None,
        track: bool = False,
        minimal: bool | None = None,
    ) -> None:
        """Find and replace text in the document.

        This method searches for text and replaces it with new text. When track=True,
        the operation shows both the deletion of the old text and insertion of the
        new text as tracked changes.

        When regex=True, the replacement string can use capture groups:
        - \\1, \\2, etc. for numbered groups
        - \\g<name> for named groups

        Args:
            find: Text or regex pattern to find
            replace: Replacement text (can include capture group references if regex=True).
                Supports markdown formatting: **bold**, *italic*, ++underline++,
                ~~strikethrough~~
            author: Optional author override (uses document author if None). Ignored
                when track=False.
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            occurrence: Which occurrence(s) to replace: 1 (first), 2 (second), "first", "last",
                "all", or list of indices [1, 3, 5] (default: "first")
            regex: Whether to treat 'find' as a regex pattern (default: False)
            normalize_special_chars: Auto-convert straight quotes to smart quotes for
                matching (default: True)
            show_context: Show text before/after the match for preview (default: False)
            check_continuity: Check if replacement may create sentence fragments (default: False)
            context_chars: Number of characters to show before/after when show_context=True
                (default: 50)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'
            track: If True, show as tracked change (w:del + w:ins). If False, replace
                text without tracking (default: False).
            minimal: If True, use word-level diffing for human-looking tracked changes.
                If False, use coarse delete-all + insert-all. If None (default),
                uses the document's minimal_edits setting. Only applies when track=True.

        Raises:
            TextNotFoundError: If the 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found and
                occurrence not specified
            re.error: If regex=True and the pattern is invalid
            ImportError: If fuzzy matching requested but rapidfuzz not installed

        Warnings:
            ContinuityWarning: If check_continuity=True and potential sentence fragment detected
        """
        # Find all matches
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Parse fuzzy configuration if provided
        from ..fuzzy import parse_fuzzy_config

        fuzzy_config = parse_fuzzy_config(fuzzy)

        matches = self._document._text_search.find_text(
            find,
            paragraphs,
            regex=regex,
            normalize_special_chars=normalize_special_chars and not regex and not fuzzy_config,
            fuzzy=fuzzy_config,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(find, paragraphs)
            raise TextNotFoundError(find, suggestions=suggestions)

        # Select target matches based on occurrence
        target_matches = self._select_matches(matches, occurrence, find)

        # Replace each target match (process in reverse to preserve indices)
        for match in reversed(target_matches):
            matched_text = match.text

            # Show context preview if requested
            if show_context:
                self._log_context_preview(match, replace, context_chars)

            # Handle capture group expansion for regex replacements
            replacement_text = self._expand_replacement(match, replace, regex)

            # Check continuity if requested
            if check_continuity:
                self._check_and_warn_continuity(match, replacement_text, context_chars)

            if track:
                # Capture change ID before operation for edit group tracking
                start_id = self._document._xml_generator.next_change_id

                # Determine effective minimal setting
                use_minimal = minimal if minimal is not None else self._document._minimal_edits

                if use_minimal:
                    # Attempt word-level minimal edit
                    from ..minimal_diff import apply_minimal_edits_to_textspan

                    success, reason = apply_minimal_edits_to_textspan(
                        match,
                        replacement_text,
                        self._document._xml_generator,
                        author,
                    )
                    if success:
                        # Record change IDs with edit group registry
                        self._record_change_ids(
                            start_id, self._document._xml_generator.next_change_id
                        )
                        continue  # Minimal edit applied successfully
                    else:
                        # Log fallback at INFO level
                        logger.info(
                            "Falling back to coarse tracked change for '%s' -> '%s': %s",
                            matched_text[:50],
                            replacement_text[:50],
                            reason,
                        )

                # Coarse tracked replace: deletion + insertion XML
                deletion_xml = self._document._xml_generator.create_deletion(matched_text, author)
                insertion_xml = self._document._xml_generator.create_insertion(
                    replacement_text, author
                )
                elements = self._parse_xml_elements(f"{deletion_xml}\n    {insertion_xml}")
                # Replace the matched text with deletion + insertion
                self._replace_match_with_elements(match, elements)

                # Record change IDs with edit group registry
                self._record_change_ids(start_id, self._document._xml_generator.next_change_id)
            else:
                # Untracked replace: just replace with plain runs
                # Get source run for formatting
                source_run = match.runs[0] if match.runs else None
                new_runs = self._document._xml_generator.create_plain_runs(
                    replacement_text, source_run=source_run
                )
                if len(new_runs) == 1:
                    self._replace_match_with_element(match, new_runs[0])
                else:
                    self._replace_match_with_elements(match, new_runs)

    def _log_context_preview(self, match: TextSpan, replacement: str, context_chars: int) -> None:
        """Log context preview for debugging."""
        before, matched, after = self._get_detailed_context(match, context_chars)
        logger.debug(
            "Context preview:\n"
            "BEFORE (%d chars): %r\n"
            "MATCH (%d chars): %r\n"
            "AFTER (%d chars): %r\n"
            "REPLACEMENT (%d chars): %r",
            len(before),
            before,
            len(matched),
            matched,
            len(after),
            after,
            len(replacement),
            replacement,
        )

    def _expand_replacement(self, match: TextSpan, replace: str, regex: bool) -> str:
        """Expand capture group references in replacement text if regex mode."""
        if regex and match.match_obj:
            return match.match_obj.expand(replace)
        return replace

    def _check_and_warn_continuity(
        self, match: TextSpan, replacement_text: str, context_chars: int
    ) -> None:
        """Check for continuity issues and emit warnings."""
        import warnings

        from ..errors import ContinuityWarning

        _, _, after_text = self._get_detailed_context(match, context_chars)
        continuity_warnings = self._check_continuity(replacement_text, after_text)

        for warning_msg in continuity_warnings:
            suggestions = [
                "Include more context in your replacement text",
                "Adjust the 'find' text to include the connecting phrase",
                "Review the following text to ensure grammatical correctness",
            ]
            warnings.warn(
                ContinuityWarning(warning_msg, after_text, suggestions),
                stacklevel=3,
            )

    def _get_detailed_context(
        self, match: TextSpan, context_chars: int = 50
    ) -> tuple[str, str, str]:
        """Extract detailed context around a match for preview.

        Args:
            match: TextSpan object representing the matched text
            context_chars: Number of characters to extract before/after (default: 50)

        Returns:
            Tuple of (before_text, matched_text, after_text)
        """
        # Extract text from the paragraph
        text_elements = match.paragraph.findall(f".//{{{WORD_NAMESPACE}}}t")
        para_text = "".join(elem.text or "" for elem in text_elements)
        matched = match.text

        # Find the match position in the full paragraph text
        match_pos = para_text.find(matched)
        if match_pos == -1:
            # Fallback: couldn't find match in paragraph
            return ("", matched, "")

        # Extract context
        before_start = max(0, match_pos - context_chars)
        after_end = min(len(para_text), match_pos + len(matched) + context_chars)

        before_text = para_text[before_start:match_pos]
        after_text = para_text[match_pos + len(matched) : after_end]

        # Add ellipsis if truncated
        if before_start > 0:
            before_text = "..." + before_text
        if after_end < len(para_text):
            after_text = after_text + "..."

        return (before_text, matched, after_text)

    def _check_continuity(self, replacement: str, next_text: str) -> list[str]:
        """Check if replacement may create a sentence fragment.

        Analyzes the text immediately following the replacement to detect
        potential grammatical issues like sentence fragments or disconnected clauses.

        Args:
            replacement: The replacement text
            next_text: Text immediately following where replacement will be inserted

        Returns:
            List of warning messages (empty if no issues detected)
        """

        warnings: list[str] = []

        # Skip check if no following text or it's just whitespace
        if not next_text or not next_text.strip():
            return warnings

        # Get the first ~30 chars of following text for analysis
        next_preview = next_text.strip()[:30]

        # Heuristic 1: Starts with lowercase letter (excluding special cases)
        # Skip 'i' for Roman numerals
        if next_preview and next_preview[0].islower() and next_preview[0] != "i":
            warnings.append("Next text starts with lowercase letter - may be a sentence fragment")

        # Heuristic 2: Starts with connecting phrase
        connecting_phrases = [
            "in question",
            "of which",
            "that is",
            "to which",
            "which is",
            "who is",
            "whose",
            "wherein",
            "whereby",
        ]

        next_lower = next_preview.lower()
        for phrase in connecting_phrases:
            if next_lower.startswith(phrase):
                warnings.append(
                    f"Next text starts with connecting phrase '{phrase}' - "
                    f"may require preceding context"
                )
                break

        # Heuristic 3: Starts with continuation punctuation
        if next_preview and next_preview[0] in [",", ";", ":", "—", "–"]:
            warnings.append("Next text starts with continuation punctuation - likely a fragment")

        return warnings

    def move(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
        source_scope: str | dict | Any | None = None,
        dest_scope: str | dict | Any | None = None,
        regex: bool = False,
        normalize_special_chars: bool = True,
        track: bool = False,
    ) -> None:
        """Move text to a new location.

        When track=True, creates linked move markers that show the text was
        relocated rather than deleted and re-added. This provides better context
        for document reviewers in Word.

        In Word's track changes view (track=True):
        - Source location shows text with strikethrough and "Moved" annotation
        - Destination shows text with underline and "Moved" annotation
        - Both locations are linked with matching move markers

        When track=False, simply deletes from source and inserts at destination
        without any tracking markers.

        Args:
            text: The text to move (or regex pattern if regex=True)
            after: Text to insert the moved content after (at destination)
            before: Text to insert the moved content before (at destination)
            author: Optional author override (uses document author if None). Ignored
                when track=False.
            source_scope: Limit source text search scope
            dest_scope: Limit destination anchor search scope
            regex: Whether to treat 'text' and anchor as regex patterns (default: False)
            normalize_special_chars: Auto-convert straight quotes to smart quotes for
                matching (default: True)
            track: If True, show as tracked move (linked moveFrom/moveTo markers).
                If False, move text without tracking (default: False).

        Raises:
            ValueError: If both 'after' and 'before' are specified, or if neither is specified
            TextNotFoundError: If the source text or destination anchor is not found
            AmbiguousTextError: If multiple occurrences of source text or anchor are found
            re.error: If regex=True and a pattern is invalid
        """
        # Validate parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        # Determine destination anchor and insertion mode
        dest_anchor: str = after if after is not None else before  # type: ignore[assignment]
        insert_after = after is not None

        # Find source text and destination anchor
        source_match = self._find_unique_match(text, source_scope, regex, normalize_special_chars)
        dest_match = self._find_unique_match(
            dest_anchor, dest_scope, regex, normalize_special_chars
        )

        source_text = source_match.text

        if track:
            # Capture change ID before operation for edit group tracking
            start_id = self._document._xml_generator.next_change_id

            # Tracked move: create linked moveFrom/moveTo markers
            move_name = self._generate_move_name()
            move_to_elements, move_from_elements = self._create_move_elements(
                source_text, move_name, author
            )

            # Insert moveTo at destination first (so we don't mess up source position)
            if insert_after:
                self._insert_after_match(dest_match, move_to_elements)
            else:
                self._insert_before_match(dest_match, move_to_elements)

            # Replace source text with moveFrom markers
            self._replace_match_with_elements(source_match, move_from_elements)

            # Record change IDs with edit group registry
            self._record_change_ids(start_id, self._document._xml_generator.next_change_id)
        else:
            # Untracked move: delete from source and insert at destination
            # Get source run for formatting
            source_run = source_match.runs[0] if source_match.runs else None

            # Create plain runs for destination (using source formatting)
            plain_runs = self._document._xml_generator.create_plain_runs(
                source_text, source_run=source_run
            )

            # Insert at destination first (before deleting source)
            if insert_after:
                self._insert_after_match(dest_match, plain_runs)
            else:
                self._insert_before_match(dest_match, plain_runs)

            # Remove source text
            self._remove_match(source_match)

    def _create_move_elements(
        self, source_text: str, move_name: str, author: str | None
    ) -> tuple[list, list]:
        """Create move-to and move-from XML elements.

        Args:
            source_text: The text being moved
            move_name: Unique identifier linking source and destination
            author: Author for the tracked change

        Returns:
            Tuple of (move_to_elements, move_from_elements)
        """
        move_from_xml, _, _ = self._document._xml_generator.create_move_from(
            source_text, move_name, author
        )
        move_to_xml, _, _ = self._document._xml_generator.create_move_to(
            source_text, move_name, author
        )

        move_to_elements = self._parse_xml_elements(move_to_xml)
        move_from_elements = self._parse_xml_elements(move_from_xml)

        return move_to_elements, move_from_elements

    def _generate_move_name(self) -> str:
        """Generate a unique move name for linking source and destination.

        Scans the document for existing move names and returns the next one.

        Returns:
            Unique move name (e.g., "move1", "move2", etc.)
        """
        # Find all existing move names
        existing_names: set[str] = set()

        # Check moveFromRangeStart elements
        for elem in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}moveFromRangeStart"):
            name = elem.get(f"{{{WORD_NAMESPACE}}}name")
            if name:
                existing_names.add(name)

        # Check moveToRangeStart elements
        for elem in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}moveToRangeStart"):
            name = elem.get(f"{{{WORD_NAMESPACE}}}name")
            if name:
                existing_names.add(name)

        # Find next available number
        counter = 1
        while f"move{counter}" in existing_names:
            counter += 1

        return f"move{counter}"

    def _insert_after_match(self, match: TextSpan, insertion_element: Any) -> None:
        """Insert XML element(s) after a matched text span.

        Args:
            match: TextSpan object representing where to insert
            insertion_element: The lxml Element or list of Elements to insert
        """
        # Get the paragraph containing the match
        paragraph = match.paragraph

        # Find the run where the match ends
        end_run = match.runs[match.end_run_index]

        # Find the position of the end run in the paragraph
        run_index = list(paragraph).index(end_run)

        # Handle single element or list
        if isinstance(insertion_element, list):
            # Insert elements in order after the end run
            for i, elem in enumerate(insertion_element):
                paragraph.insert(run_index + 1 + i, elem)
        else:
            # Insert the new element after the end run
            paragraph.insert(run_index + 1, insertion_element)

    def _insert_before_match(self, match: TextSpan, insertion_element: Any) -> None:
        """Insert XML element(s) before a matched text span.

        Args:
            match: TextSpan object representing where to insert
            insertion_element: The lxml Element or list of Elements to insert
        """
        # Get the paragraph containing the match
        paragraph = match.paragraph

        # Find the run where the match starts
        start_run = match.runs[match.start_run_index]

        # Find the position of the start run in the paragraph
        run_index = list(paragraph).index(start_run)

        # Handle single element or list
        if isinstance(insertion_element, list):
            # Insert elements in order before the start run
            for i, elem in enumerate(insertion_element):
                paragraph.insert(run_index + i, elem)
        else:
            # Insert the new element before the start run
            paragraph.insert(run_index, insertion_element)

    def _replace_match_with_element(self, match: TextSpan, replacement_element: Any) -> None:
        """Replace matched text with a single XML element.

        This handles the complexity of text potentially spanning multiple runs.
        The matched runs are removed and replaced with the new element.

        Handles runs that may be inside tracked change wrappers (w:ins, w:del).
        When deleting text inside a wrapper, the replacement element (w:del)
        must be placed at paragraph level, wrapping the original wrapper content.

        Args:
            match: TextSpan object representing the text to replace
            replacement_element: The lxml Element to insert in place of matched text
        """
        paragraph = match.paragraph

        # If the match is within a single run
        if match.start_run_index == match.end_run_index:
            run = match.runs[match.start_run_index]
            run_text = "".join(run.itertext())

            # Get the actual parent of the run (may be paragraph or a wrapper)
            actual_parent = run.getparent()
            if actual_parent is None:
                actual_parent = paragraph

            # If the match is the entire run, replace the run
            if match.start_offset == 0 and match.end_offset == len(run_text):
                # Check if the parent is a tracked change wrapper
                if self._is_tracked_change_wrapper(actual_parent):
                    # The replacement element (e.g., w:del) must be placed at paragraph level
                    # wrapping the original wrapper content
                    self._replace_run_in_wrapper_with_element(
                        paragraph, run, actual_parent, replacement_element
                    )
                else:
                    try:
                        run_index = list(actual_parent).index(run)
                    except ValueError:
                        # Fallback to paragraph
                        run_index = list(paragraph).index(run)
                        actual_parent = paragraph
                    actual_parent.remove(run)
                    actual_parent.insert(run_index, replacement_element)
            else:
                # Match is partial - need to split the run
                self._split_and_replace_in_run(
                    paragraph, run, match.start_offset, match.end_offset, replacement_element
                )
        else:
            # Match spans multiple runs - need to preserve text before/after match
            start_run = match.runs[match.start_run_index]
            end_run = match.runs[match.end_run_index]

            # Get the actual parent of the start run
            actual_parent = start_run.getparent()
            if actual_parent is None:
                actual_parent = paragraph

            try:
                start_run_index = list(actual_parent).index(start_run)
            except ValueError:
                # Fallback to paragraph
                start_run_index = list(paragraph).index(start_run)
                actual_parent = paragraph

            # Get text before match in the first run (only from w:t elements)
            first_run_text = self._get_run_text_content(start_run)
            before_text = first_run_text[: match.start_offset]

            # Get text after match in the last run (only from w:t elements)
            last_run_text = self._get_run_text_content(end_run)
            after_text = last_run_text[match.end_offset :]

            # Remove all runs in the match
            for i in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[i]
                run_parent = run.getparent()
                if run_parent is not None and run in run_parent:
                    run_parent.remove(run)

            # Build replacement elements: [before_run] + replacement + [after_run]
            new_elements = self._build_split_elements(
                start_run, before_text, after_text, [replacement_element]
            )

            # Insert all elements at the position of the first removed run
            for i, elem in enumerate(new_elements):
                actual_parent.insert(start_run_index + i, elem)

    def _replace_run_in_wrapper_with_element(
        self,
        paragraph: Any,
        run: Any,
        wrapper: Any,
        replacement_element: Any,
    ) -> None:
        """Replace a run inside a tracked change wrapper with a new element.

        When a run inside a w:ins or w:del needs to be replaced with another
        tracked change element (e.g., deleting text inside an insertion), the
        replacement must be placed at paragraph level wrapping the original content.

        For example, deleting "text" inside <w:ins><w:r><w:t>text</w:t></w:r></w:ins>
        should produce:
            <w:del>
                <w:ins>
                    <w:r><w:delText>text</w:delText></w:r>
                </w:ins>
            </w:del>

        Args:
            paragraph: The containing paragraph
            run: The run being replaced
            wrapper: The tracked change wrapper containing the run
            replacement_element: The new element (e.g., w:del) to wrap the content
        """
        wrapper_children = list(wrapper)
        run_index_in_wrapper = wrapper_children.index(run)

        # Create an empty wrapper element with same attributes as replacement
        # (we don't want the pre-generated content, just the wrapper attributes)
        empty_replacement = self._create_empty_wrapper_like(replacement_element)

        # Check if the run is the only child of the wrapper
        if len(wrapper_children) == 1:
            # Simple case: entire wrapper content is being replaced
            # Find wrapper position in paragraph
            wrapper_index = list(paragraph).index(wrapper)

            # Move the wrapper (with its content) inside the empty replacement element
            paragraph.remove(wrapper)

            # The run inside the wrapper needs to have its content converted
            # if the replacement is a w:del, convert w:t to w:delText
            if self._is_deletion_wrapper(replacement_element):
                self._convert_run_text_to_deltext(run)

            empty_replacement.append(wrapper)
            paragraph.insert(wrapper_index, empty_replacement)
        else:
            # Complex case: only part of the wrapper content is being replaced
            # Need to split the wrapper
            wrapper_index = list(paragraph).index(wrapper)

            # Create elements for before, replacement, and after
            elements_to_insert: list[Any] = []

            # Content before the run stays in a cloned wrapper
            if run_index_in_wrapper > 0:
                before_wrapper = self._clone_wrapper(wrapper)
                for i in range(run_index_in_wrapper):
                    child_copy = etree.fromstring(etree.tostring(wrapper_children[i]))
                    before_wrapper.append(child_copy)
                elements_to_insert.append(before_wrapper)

            # The run being replaced goes in empty_replacement wrapped in cloned wrapper
            run_wrapper = self._clone_wrapper(wrapper)
            run_copy = etree.fromstring(etree.tostring(run))
            if self._is_deletion_wrapper(replacement_element):
                self._convert_run_text_to_deltext(run_copy)
            run_wrapper.append(run_copy)
            empty_replacement.append(run_wrapper)
            elements_to_insert.append(empty_replacement)

            # Content after the run stays in a cloned wrapper
            if run_index_in_wrapper < len(wrapper_children) - 1:
                after_wrapper = self._clone_wrapper(wrapper)
                for i in range(run_index_in_wrapper + 1, len(wrapper_children)):
                    child_copy = etree.fromstring(etree.tostring(wrapper_children[i]))
                    after_wrapper.append(child_copy)
                elements_to_insert.append(after_wrapper)

            # Remove original wrapper and insert new elements
            paragraph.remove(wrapper)
            for i, elem in enumerate(elements_to_insert):
                paragraph.insert(wrapper_index + i, elem)

    def _replace_run_in_wrapper_with_elements(
        self,
        paragraph: Any,
        run: Any,
        wrapper: Any,
        replacement_elements: list[Any],
    ) -> None:
        """Replace a run inside a tracked change wrapper with multiple elements.

        When a run inside a w:ins or w:del needs to be replaced with multiple
        tracked change elements (e.g., replacing text inside an insertion with
        a deletion + insertion pair), the tracked change elements must be placed
        at paragraph level.

        For replace_tracked on text inside <w:ins>:
        - The w:del goes at paragraph level wrapping a copy of the w:ins with delText
        - The w:ins for the new text goes at paragraph level

        Args:
            paragraph: The containing paragraph
            run: The run being replaced
            wrapper: The tracked change wrapper containing the run
            replacement_elements: The new elements (e.g., [w:del, w:ins]) to insert
        """
        wrapper_children = list(wrapper)
        run_index_in_wrapper = wrapper_children.index(run)
        wrapper_index = list(paragraph).index(wrapper)

        # Build the elements to insert at paragraph level
        elements_to_insert: list[Any] = []

        # Content before the run stays in a cloned wrapper
        if run_index_in_wrapper > 0:
            before_wrapper = self._clone_wrapper(wrapper)
            for i in range(run_index_in_wrapper):
                child_copy = etree.fromstring(etree.tostring(wrapper_children[i]))
                before_wrapper.append(child_copy)
            elements_to_insert.append(before_wrapper)

        # Process each replacement element
        for repl_elem in replacement_elements:
            if self._is_tracked_change_wrapper(repl_elem):
                # Tracked change elements need special handling:
                # - w:del needs to wrap a copy of the original wrapper with the run
                # - w:ins can be inserted directly at paragraph level
                if self._is_deletion_wrapper(repl_elem):
                    # Create empty del wrapper with same attributes
                    empty_del = self._create_empty_wrapper_like(repl_elem)
                    # Clone the original wrapper and put the run (with delText) inside
                    run_wrapper = self._clone_wrapper(wrapper)
                    run_copy = etree.fromstring(etree.tostring(run))
                    self._convert_run_text_to_deltext(run_copy)
                    run_wrapper.append(run_copy)
                    empty_del.append(run_wrapper)
                    elements_to_insert.append(empty_del)
                else:
                    # For w:ins elements, insert directly at paragraph level
                    elements_to_insert.append(repl_elem)
            else:
                # Non-tracked-change elements can be inserted directly
                elements_to_insert.append(repl_elem)

        # Content after the run stays in a cloned wrapper
        if run_index_in_wrapper < len(wrapper_children) - 1:
            after_wrapper = self._clone_wrapper(wrapper)
            for i in range(run_index_in_wrapper + 1, len(wrapper_children)):
                child_copy = etree.fromstring(etree.tostring(wrapper_children[i]))
                after_wrapper.append(child_copy)
            elements_to_insert.append(after_wrapper)

        # Remove original wrapper and insert new elements
        paragraph.remove(wrapper)
        for i, elem in enumerate(elements_to_insert):
            paragraph.insert(wrapper_index + i, elem)

    def _create_empty_wrapper_like(self, wrapper: Any) -> Any:
        """Create an empty wrapper element with the same tag and attributes.

        This is used when we need the wrapper structure (w:del, w:ins, etc.)
        but not its pre-generated content.

        Args:
            wrapper: The wrapper element to copy tag and attributes from

        Returns:
            A new empty element with the same tag and attributes
        """
        new_wrapper = etree.Element(wrapper.tag)
        for attr_name, attr_value in wrapper.attrib.items():
            new_wrapper.set(attr_name, attr_value)
        return new_wrapper

    def _convert_run_text_to_deltext(self, run: Any) -> None:
        """Convert w:t elements in a run to w:delText elements.

        This is needed when text inside a w:ins is being deleted - the text
        must be marked with w:delText to indicate it's deleted content.

        Args:
            run: The run element containing w:t elements to convert
        """
        for t_elem in run.findall(f".//{{{WORD_NAMESPACE}}}t"):
            t_elem.tag = f"{{{WORD_NAMESPACE}}}delText"

    def _replace_match_with_elements(
        self, match: TextSpan, replacement_elements: list[Any]
    ) -> None:
        """Replace matched text with multiple XML elements.

        Used for replace_tracked which needs both deletion and insertion elements.

        Handles runs that may be inside tracked change wrappers (w:ins, w:del).
        When the run is inside a wrapper and replacement_elements contain tracked
        change elements (w:del, w:ins), those elements must be placed at
        paragraph level, not inside the wrapper.

        Args:
            match: TextSpan object representing the text to replace
            replacement_elements: List of lxml Elements to insert in place of matched text
        """
        paragraph = match.paragraph

        # Similar to _replace_match_with_element but inserts multiple elements
        if match.start_run_index == match.end_run_index:
            run = match.runs[match.start_run_index]
            run_text = "".join(run.itertext())

            # Get the actual parent of the run (may be paragraph or a wrapper)
            actual_parent = run.getparent()
            if actual_parent is None:
                actual_parent = paragraph

            # If the match is the entire run, replace the run
            if match.start_offset == 0 and match.end_offset == len(run_text):
                # Check if the parent is a tracked change wrapper
                if self._is_tracked_change_wrapper(actual_parent):
                    # Replacement elements must be placed at paragraph level
                    self._replace_run_in_wrapper_with_elements(
                        paragraph, run, actual_parent, replacement_elements
                    )
                else:
                    try:
                        run_index = list(actual_parent).index(run)
                    except ValueError:
                        # Fallback to paragraph
                        run_index = list(paragraph).index(run)
                        actual_parent = paragraph
                    actual_parent.remove(run)
                    # Insert elements in order
                    for i, elem in enumerate(replacement_elements):
                        actual_parent.insert(run_index + i, elem)
            else:
                # Match is partial - need to split the run
                self._split_and_replace_in_run_multiple(
                    paragraph,
                    run,
                    match.start_offset,
                    match.end_offset,
                    replacement_elements,
                )
        else:
            # Match spans multiple runs - need to preserve text before/after match
            # and handle cases where runs are in different parents (wrappers vs paragraph)
            self._replace_multirun_match_with_elements(match, paragraph, replacement_elements)

    def _get_run_text_content(self, run: Any) -> str:
        """Extract text content from a run, avoiding XML structural whitespace.

        Only extracts text from w:t and w:delText elements, not from XML
        formatting whitespace between tags.

        Args:
            run: A w:r (run) Element

        Returns:
            Text content of the run
        """
        text_elements = run.findall(f".//{{{WORD_NAMESPACE}}}t")
        deltext_elements = run.findall(f".//{{{WORD_NAMESPACE}}}delText")
        all_text_elements = text_elements + deltext_elements
        return "".join(elem.text or "" for elem in all_text_elements)

    def _create_text_run(self, text: str, source_run: Any) -> Any:
        """Create a new text run with optional properties from source run.

        Args:
            text: The text content for the new run
            source_run: The run to copy properties from (if any)

        Returns:
            A new w:r element with the text content
        """
        new_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
        run_props = source_run.find(f"{{{WORD_NAMESPACE}}}rPr")
        if run_props is not None:
            new_run.append(etree.fromstring(etree.tostring(run_props)))
        text_elem = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}t")
        if text and (text[0].isspace() or text[-1].isspace()):
            text_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_elem.text = text
        return new_run

    def _get_run_text_info(self, run: Any) -> tuple[str, int]:
        """Extract text content and position info from a run.

        Args:
            run: The run element to extract from

        Returns:
            Tuple of (run_text, number_of_text_elements)
        """
        text_elements = list(run.iter(f"{{{WORD_NAMESPACE}}}t"))
        if not text_elements:
            return "", 0
        return text_elements[0].text or "", len(text_elements)

    def _is_tracked_change_wrapper(self, elem: Any) -> bool:
        """Check if an element is a tracked change wrapper.

        Tracked change wrappers include:
        - w:ins (tracked insertion)
        - w:del (tracked deletion)
        - w:moveFrom (tracked move source)
        - w:moveTo (tracked move destination)

        Args:
            elem: The element to check

        Returns:
            True if the element is a tracked change wrapper
        """
        if elem is None:
            return False
        tag = elem.tag
        return tag in (
            f"{{{WORD_NAMESPACE}}}ins",
            f"{{{WORD_NAMESPACE}}}del",
            f"{{{WORD_NAMESPACE}}}moveFrom",
            f"{{{WORD_NAMESPACE}}}moveTo",
        )

    def _is_insertion_wrapper(self, elem: Any) -> bool:
        """Check if an element is a tracked insertion wrapper (w:ins).

        Args:
            elem: The element to check

        Returns:
            True if the element is a w:ins element
        """
        if elem is None:
            return False
        return elem.tag == f"{{{WORD_NAMESPACE}}}ins"

    def _is_deletion_wrapper(self, elem: Any) -> bool:
        """Check if an element is a tracked deletion wrapper (w:del).

        Args:
            elem: The element to check

        Returns:
            True if the element is a w:del element
        """
        if elem is None:
            return False
        return elem.tag == f"{{{WORD_NAMESPACE}}}del"

    def _get_wrapper_author(self, wrapper: Any) -> str | None:
        """Get the author attribute from a tracked change wrapper.

        Args:
            wrapper: A w:ins or w:del element

        Returns:
            The author name, or None if not found
        """
        return wrapper.get(f"{{{WORD_NAMESPACE}}}author")

    def _is_same_author(self, wrapper: Any, author: str | None) -> bool:
        """Check if the wrapper's author matches the given author.

        Args:
            wrapper: A w:ins or w:del element
            author: The author to compare against (uses document author if None)

        Returns:
            True if authors match
        """
        current_author = author if author is not None else self._document._xml_generator.author
        wrapper_author = self._get_wrapper_author(wrapper)
        return wrapper_author == current_author

    def _get_run_parent_info(self, run: Any) -> tuple[Any, Any, int]:
        """Get parent information for a run, handling tracked change wrappers.

        A run might be a direct child of a paragraph, or nested inside a
        tracked change wrapper (w:ins or w:del).

        Args:
            run: The run element

        Returns:
            Tuple of (paragraph, immediate_parent, index_in_parent)
            - paragraph: The containing paragraph element
            - immediate_parent: The direct parent (paragraph or wrapper)
            - index_in_parent: The run's index in its immediate parent
        """
        immediate_parent = run.getparent()
        if immediate_parent is None:
            raise ValueError("Run has no parent element")

        # Find the containing paragraph
        paragraph = immediate_parent
        while paragraph is not None and paragraph.tag != f"{{{WORD_NAMESPACE}}}p":
            paragraph = paragraph.getparent()

        if paragraph is None:
            raise ValueError("Run is not contained in a paragraph")

        run_index = list(immediate_parent).index(run)
        return paragraph, immediate_parent, run_index

    def _clone_wrapper(self, wrapper: Any, new_id: int | None = None) -> Any:
        """Create a copy of a tracked change wrapper (w:ins or w:del) with a new ID.

        Args:
            wrapper: The wrapper element to clone
            new_id: Optional new ID to use. If None, generates a new unique ID.

        Returns:
            A new wrapper element with the same attributes but a new ID
        """
        if new_id is None:
            new_id = self._document._xml_generator.next_change_id

        # Create new element with same tag
        new_wrapper = etree.Element(wrapper.tag)

        # Copy all attributes except id
        for attr_name, attr_value in wrapper.attrib.items():
            if attr_name == f"{{{WORD_NAMESPACE}}}id":
                new_wrapper.set(attr_name, str(new_id))
            else:
                new_wrapper.set(attr_name, attr_value)

        # If no id was present, add one
        id_attr = f"{{{WORD_NAMESPACE}}}id"
        if id_attr not in new_wrapper.attrib:
            new_wrapper.set(id_attr, str(new_id))

        return new_wrapper

    def _get_wrapper_position_in_paragraph(self, wrapper: Any, paragraph: Any) -> int:
        """Get the position of a wrapper element within its paragraph.

        Args:
            wrapper: The wrapper element (w:ins or w:del)
            paragraph: The containing paragraph

        Returns:
            The index of the wrapper in the paragraph's children

        Raises:
            ValueError: If wrapper is not a direct child of paragraph
        """
        return list(paragraph).index(wrapper)

    def _extract_remaining_content_from_wrapper(
        self,
        wrapper: Any,
        paragraph: Any,
        start_run_index: int | None = None,
        end_run_index: int | None = None,
    ) -> tuple[Any | None, Any | None]:
        """Extract content before and/or after specified runs from a wrapper.

        This creates new wrapper elements containing the content that should be
        preserved (not part of the match). The original wrapper should be handled
        separately (removed or modified).

        Args:
            wrapper: The wrapper element containing runs
            paragraph: The containing paragraph
            start_run_index: Index of first run in match (content before this is extracted)
            end_run_index: Index of last run in match (content after this is extracted)

        Returns:
            Tuple of (before_wrapper, after_wrapper) - either may be None if no content
        """
        runs_in_wrapper = list(wrapper)
        before_wrapper = None
        after_wrapper = None

        # Extract content before the match
        if start_run_index is not None and start_run_index > 0:
            before_wrapper = self._clone_wrapper(wrapper)
            for i in range(start_run_index):
                run_copy = etree.fromstring(etree.tostring(runs_in_wrapper[i]))
                before_wrapper.append(run_copy)

        # Extract content after the match
        if end_run_index is not None and end_run_index < len(runs_in_wrapper) - 1:
            after_wrapper = self._clone_wrapper(wrapper)
            for i in range(end_run_index + 1, len(runs_in_wrapper)):
                run_copy = etree.fromstring(etree.tostring(runs_in_wrapper[i]))
                after_wrapper.append(run_copy)

        return before_wrapper, after_wrapper

    def _wrap_runs_in_cloned_wrapper(self, runs: list[Any], original_wrapper: Any) -> Any:
        """Wrap runs in a cloned wrapper with same attributes but new ID.

        Args:
            runs: List of run elements to wrap
            original_wrapper: The wrapper to clone attributes from

        Returns:
            New wrapper element containing the runs
        """
        new_wrapper = self._clone_wrapper(original_wrapper)
        for run in runs:
            new_wrapper.append(run)
        return new_wrapper

    def _build_split_elements(
        self, run: Any, before_text: str, after_text: str, replacement_elements: list[Any]
    ) -> list[Any]:
        """Build the list of elements for a run split operation.

        Args:
            run: The original run being split
            before_text: Text before the replacement (may be empty)
            after_text: Text after the replacement (may be empty)
            replacement_elements: Elements to insert in the middle

        Returns:
            List of elements to insert
        """
        new_elements = []
        if before_text:
            new_elements.append(self._create_text_run(before_text, run))
        new_elements.extend(replacement_elements)
        if after_text:
            new_elements.append(self._create_text_run(after_text, run))
        return new_elements

    def _replace_run_with_elements(self, paragraph: Any, run: Any, new_elements: list[Any]) -> None:
        """Replace a run with a list of new elements.

        Handles runs that may be inside tracked change wrappers (w:ins, w:del).
        Uses getparent() to find the actual parent rather than assuming the
        paragraph is the direct parent.

        Args:
            paragraph: The paragraph containing the run (used as fallback)
            run: The run to replace
            new_elements: Elements to insert in place of the run
        """
        # Get the actual parent of the run (may be paragraph or a wrapper)
        actual_parent = run.getparent()
        if actual_parent is None:
            # Fallback to paragraph if no parent found
            actual_parent = paragraph

        try:
            run_index = list(actual_parent).index(run)
        except ValueError:
            # Run not in expected parent, try paragraph as fallback
            run_index = list(paragraph).index(run)
            actual_parent = paragraph

        actual_parent.remove(run)
        for i, elem in enumerate(new_elements):
            actual_parent.insert(run_index + i, elem)

    def _split_and_replace_in_run(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_element: Any,
    ) -> None:
        """Split a run and replace a portion with a new element.

        When the run is inside a tracked change wrapper and the replacement is
        also a tracked change element, the replacement must be placed at
        paragraph level with the matched portion wrapped appropriately.

        Args:
            paragraph: The paragraph containing the run
            run: The run to split
            start_offset: Character offset where match starts
            end_offset: Character offset where match ends (exclusive)
            replacement_element: Element to insert in place of matched text
        """
        run_text, num_elements = self._get_run_text_info(run)
        if num_elements == 0:
            return

        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]
        matched_text = run_text[start_offset:end_offset]

        # Check if run is inside a tracked change wrapper
        actual_parent = run.getparent()
        if (
            actual_parent is not None
            and self._is_tracked_change_wrapper(actual_parent)
            and self._is_tracked_change_wrapper(replacement_element)
        ):
            # Special handling: replacement goes at paragraph level
            self._split_run_in_wrapper_with_element(
                paragraph,
                run,
                actual_parent,
                before_text,
                matched_text,
                after_text,
                replacement_element,
            )
        else:
            # Standard handling: insert elements at run level
            new_elements = self._build_split_elements(
                run, before_text, after_text, [replacement_element]
            )
            self._replace_run_with_elements(paragraph, run, new_elements)

    def _split_run_in_wrapper_with_element(
        self,
        paragraph: Any,
        run: Any,
        wrapper: Any,
        before_text: str,
        matched_text: str,
        after_text: str,
        replacement_element: Any,
    ) -> None:
        """Split a run inside a wrapper and place replacement at paragraph level.

        This handles partial text matches within a run that's inside a tracked
        change wrapper. The wrapper must be split to maintain proper structure:
        - Before text stays in original wrapper (or cloned if needed)
        - Matched text goes in replacement_element wrapping a cloned wrapper
        - After text goes in a new cloned wrapper

        Args:
            paragraph: The containing paragraph
            run: The run being split
            wrapper: The tracked change wrapper containing the run
            before_text: Text before the match
            matched_text: The matched text being replaced
            after_text: Text after the match
            replacement_element: The new element (e.g., w:del)
        """
        wrapper_children = list(wrapper)
        run_index_in_wrapper = wrapper_children.index(run)
        wrapper_index = list(paragraph).index(wrapper)

        # Create an empty wrapper element with same attributes as replacement
        empty_replacement = self._create_empty_wrapper_like(replacement_element)

        # Build the elements to insert at paragraph level
        elements_to_insert: list[Any] = []

        # Content before the run in wrapper stays (if any)
        if run_index_in_wrapper > 0:
            before_runs_wrapper = self._clone_wrapper(wrapper)
            for i in range(run_index_in_wrapper):
                child_copy = etree.fromstring(etree.tostring(wrapper_children[i]))
                before_runs_wrapper.append(child_copy)
            elements_to_insert.append(before_runs_wrapper)

        # Before text from the split run (if any) - stays in wrapper
        if before_text:
            before_text_wrapper = self._clone_wrapper(wrapper)
            before_run = self._create_text_run(before_text, run)
            before_text_wrapper.append(before_run)
            elements_to_insert.append(before_text_wrapper)

        # Matched text goes in replacement wrapped in cloned wrapper
        matched_wrapper = self._clone_wrapper(wrapper)
        matched_run = self._create_text_run(matched_text, run)
        if self._is_deletion_wrapper(replacement_element):
            self._convert_run_text_to_deltext(matched_run)
        matched_wrapper.append(matched_run)
        empty_replacement.append(matched_wrapper)
        elements_to_insert.append(empty_replacement)

        # After text from the split run (if any) - in new wrapper
        if after_text:
            after_text_wrapper = self._clone_wrapper(wrapper)
            after_run = self._create_text_run(after_text, run)
            after_text_wrapper.append(after_run)
            elements_to_insert.append(after_text_wrapper)

        # Content after the run in wrapper (if any)
        if run_index_in_wrapper < len(wrapper_children) - 1:
            after_runs_wrapper = self._clone_wrapper(wrapper)
            for i in range(run_index_in_wrapper + 1, len(wrapper_children)):
                child_copy = etree.fromstring(etree.tostring(wrapper_children[i]))
                after_runs_wrapper.append(child_copy)
            elements_to_insert.append(after_runs_wrapper)

        # Remove original wrapper and insert new elements
        paragraph.remove(wrapper)
        for i, elem in enumerate(elements_to_insert):
            paragraph.insert(wrapper_index + i, elem)

    def _split_run_in_wrapper_with_elements(
        self,
        paragraph: Any,
        run: Any,
        wrapper: Any,
        before_text: str,
        after_text: str,
        replacement_elements: list[Any],
    ) -> None:
        """Split a run inside a wrapper and place multiple replacements at paragraph level.

        Similar to _split_run_in_wrapper_with_element but handles multiple replacement
        elements (e.g., for replace_tracked which produces both w:del and w:ins).

        The structure produced is:
        - Before text stays in original wrapper type
        - Each tracked change replacement is placed at paragraph level
        - w:del wraps a copy of the original wrapper with delText
        - w:ins is inserted directly at paragraph level
        - After text stays in original wrapper type

        Args:
            paragraph: The containing paragraph
            run: The run being split
            wrapper: The tracked change wrapper containing the run
            before_text: Text before the match
            after_text: Text after the match
            replacement_elements: The new elements (e.g., [w:del, w:ins])
        """
        wrapper_children = list(wrapper)
        run_index_in_wrapper = wrapper_children.index(run)
        wrapper_index = list(paragraph).index(wrapper)

        matched_text = self._get_run_text_content(run)[
            len(before_text) : len(self._get_run_text_content(run)) - len(after_text)
        ]

        # Build the elements to insert at paragraph level
        elements_to_insert: list[Any] = []

        # Content before the run in wrapper stays (if any)
        if run_index_in_wrapper > 0:
            before_runs_wrapper = self._clone_wrapper(wrapper)
            for i in range(run_index_in_wrapper):
                child_copy = etree.fromstring(etree.tostring(wrapper_children[i]))
                before_runs_wrapper.append(child_copy)
            elements_to_insert.append(before_runs_wrapper)

        # Before text from the split run (if any) - stays in wrapper
        if before_text:
            before_text_wrapper = self._clone_wrapper(wrapper)
            before_run = self._create_text_run(before_text, run)
            before_text_wrapper.append(before_run)
            elements_to_insert.append(before_text_wrapper)

        # Process each replacement element
        for repl_elem in replacement_elements:
            if self._is_tracked_change_wrapper(repl_elem):
                if self._is_deletion_wrapper(repl_elem):
                    # Create empty del wrapper with same attributes
                    empty_del = self._create_empty_wrapper_like(repl_elem)
                    # Clone the original wrapper and put the matched text (with delText) inside
                    matched_wrapper = self._clone_wrapper(wrapper)
                    matched_run = self._create_text_run(matched_text, run)
                    self._convert_run_text_to_deltext(matched_run)
                    matched_wrapper.append(matched_run)
                    empty_del.append(matched_wrapper)
                    elements_to_insert.append(empty_del)
                else:
                    # For w:ins elements, insert directly at paragraph level
                    elements_to_insert.append(repl_elem)
            else:
                # Non-tracked-change elements can be inserted directly
                elements_to_insert.append(repl_elem)

        # After text from the split run (if any) - in new wrapper
        if after_text:
            after_text_wrapper = self._clone_wrapper(wrapper)
            after_run = self._create_text_run(after_text, run)
            after_text_wrapper.append(after_run)
            elements_to_insert.append(after_text_wrapper)

        # Content after the run in wrapper (if any)
        if run_index_in_wrapper < len(wrapper_children) - 1:
            after_runs_wrapper = self._clone_wrapper(wrapper)
            for i in range(run_index_in_wrapper + 1, len(wrapper_children)):
                child_copy = etree.fromstring(etree.tostring(wrapper_children[i]))
                after_runs_wrapper.append(child_copy)
            elements_to_insert.append(after_runs_wrapper)

        # Remove original wrapper and insert new elements
        paragraph.remove(wrapper)
        for i, elem in enumerate(elements_to_insert):
            paragraph.insert(wrapper_index + i, elem)

    def _split_and_replace_in_run_multiple(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_elements: list[Any],
    ) -> None:
        """Split a run and replace a portion with multiple new elements.

        When the run is inside a tracked change wrapper and replacement_elements
        contain tracked change elements, those elements must be placed at
        paragraph level with proper nesting.

        Args:
            paragraph: The paragraph containing the run
            run: The run to split
            start_offset: Character offset where match starts
            end_offset: Character offset where match ends (exclusive)
            replacement_elements: Elements to insert in place of matched text
        """
        run_text, num_elements = self._get_run_text_info(run)
        if num_elements == 0:
            return

        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        # Check if run is inside a tracked change wrapper and any replacement
        # elements are also tracked change wrappers
        actual_parent = run.getparent()
        has_tracked_replacement = any(
            self._is_tracked_change_wrapper(elem) for elem in replacement_elements
        )
        if (
            actual_parent is not None
            and self._is_tracked_change_wrapper(actual_parent)
            and has_tracked_replacement
        ):
            # Special handling: tracked change elements go at paragraph level
            self._split_run_in_wrapper_with_elements(
                paragraph,
                run,
                actual_parent,
                before_text,
                after_text,
                replacement_elements,
            )
        else:
            # Standard handling: insert elements at run level
            new_elements = self._build_split_elements(
                run, before_text, after_text, replacement_elements
            )
            self._replace_run_with_elements(paragraph, run, new_elements)

    def _replace_multirun_match_with_elements(
        self,
        match: TextSpan,
        paragraph: Any,
        replacement_elements: list[Any],
    ) -> None:
        """Replace a match spanning multiple runs with new elements.

        This handles the complex case where runs may have different parents:
        - Some runs may be direct children of the paragraph
        - Some runs may be inside tracked change wrappers (w:ins, w:del)

        When runs are removed from wrappers, any remaining content in those
        wrappers is preserved by creating new wrapper elements.

        Args:
            match: TextSpan object representing the text to replace
            paragraph: The paragraph containing the runs
            replacement_elements: Elements to insert in place of matched text
        """
        start_run = match.runs[match.start_run_index]
        end_run = match.runs[match.end_run_index]

        # Save parent information BEFORE removing runs (getparent() won't work after removal)
        start_run_parent = start_run.getparent()
        end_run_parent = end_run.getparent()

        # Get text before match in the first run
        first_run_text = self._get_run_text_content(start_run)
        before_text = first_run_text[: match.start_offset]

        # Get text after match in the last run
        last_run_text = self._get_run_text_content(end_run)
        after_text = last_run_text[match.end_offset :]

        # Track wrappers that need to have remaining content preserved
        # Key: wrapper element, Value: (runs_to_keep_before, runs_to_keep_after)
        wrapper_remaining_content: dict[Any, tuple[list[Any], list[Any]]] = {}

        # Analyze each run's parent and track wrapper content
        for i in range(match.start_run_index, match.end_run_index + 1):
            run = match.runs[i]
            run_parent = run.getparent()

            if run_parent is not None and self._is_tracked_change_wrapper(run_parent):
                wrapper = run_parent
                wrapper_children = list(wrapper)
                run_idx_in_wrapper = wrapper_children.index(run)

                if wrapper not in wrapper_remaining_content:
                    # First time seeing this wrapper - initialize
                    wrapper_remaining_content[wrapper] = ([], [])

                # For first run in wrapper that's part of match, save content before
                before_runs, after_runs = wrapper_remaining_content[wrapper]
                if not before_runs and run_idx_in_wrapper > 0:
                    # Only save before content if this is the first matched run in wrapper
                    is_first_matched_in_wrapper = True
                    for j in range(match.start_run_index, i):
                        other_run = match.runs[j]
                        if other_run.getparent() == wrapper:
                            is_first_matched_in_wrapper = False
                            break
                    if is_first_matched_in_wrapper:
                        for k in range(run_idx_in_wrapper):
                            before_runs.append(wrapper_children[k])

                # For last run in wrapper that's part of match, save content after
                is_last_matched_in_wrapper = True
                for j in range(i + 1, match.end_run_index + 1):
                    other_run = match.runs[j]
                    if other_run.getparent() == wrapper:
                        is_last_matched_in_wrapper = False
                        break
                if is_last_matched_in_wrapper and run_idx_in_wrapper < len(wrapper_children) - 1:
                    for k in range(run_idx_in_wrapper + 1, len(wrapper_children)):
                        after_runs.append(wrapper_children[k])

                wrapper_remaining_content[wrapper] = (before_runs, after_runs)

        # Find the insertion point in the paragraph
        # We need to insert at paragraph level, finding the right position
        insertion_index = self._find_paragraph_insertion_index(match, paragraph, start_run)

        # Remove all runs in the match from their parents
        for i in range(match.start_run_index, match.end_run_index + 1):
            run = match.runs[i]
            run_parent = run.getparent()
            if run_parent is not None and run in run_parent:
                run_parent.remove(run)

        # Remove empty wrappers and handle wrapper remaining content
        preserved_before: list[Any] = []
        preserved_after: list[Any] = []

        for wrapper, (before_runs, after_runs) in wrapper_remaining_content.items():
            # Check if wrapper is empty or only has matched runs
            wrapper_still_has_content = len(list(wrapper)) > 0

            if not wrapper_still_has_content:
                # Remove empty wrapper
                if wrapper.getparent() is not None:
                    wrapper.getparent().remove(wrapper)

            # Create new wrappers for remaining content
            if before_runs:
                new_before_wrapper = self._clone_wrapper(wrapper)
                for run in before_runs:
                    # Clone the run since it's still attached to original wrapper
                    run_copy = etree.fromstring(etree.tostring(run))
                    new_before_wrapper.append(run_copy)
                preserved_before.append(new_before_wrapper)

            if after_runs:
                new_after_wrapper = self._clone_wrapper(wrapper)
                for run in after_runs:
                    # Clone the run since it's still attached to original wrapper
                    run_copy = etree.fromstring(etree.tostring(run))
                    new_after_wrapper.append(run_copy)
                preserved_after.append(new_after_wrapper)

        # Build the elements to insert
        elements_to_insert: list[Any] = []

        # Add preserved content from wrappers (before)
        elements_to_insert.extend(preserved_before)

        # Add before_text run if needed
        # If the start_run was inside a wrapper, wrap the before_text in a cloned wrapper
        if before_text:
            before_run = self._create_text_run(before_text, start_run)
            if start_run_parent is not None and self._is_tracked_change_wrapper(start_run_parent):
                # Wrap in cloned wrapper to preserve original attribution
                before_wrapper = self._clone_wrapper(start_run_parent)
                before_wrapper.append(before_run)
                elements_to_insert.append(before_wrapper)
            else:
                elements_to_insert.append(before_run)

        # Add replacement elements
        elements_to_insert.extend(replacement_elements)

        # Add after_text run if needed
        # If the end_run was inside a wrapper, wrap the after_text in a cloned wrapper
        if after_text:
            after_run = self._create_text_run(after_text, end_run)
            if end_run_parent is not None and self._is_tracked_change_wrapper(end_run_parent):
                # Wrap in cloned wrapper to preserve original attribution
                after_wrapper = self._clone_wrapper(end_run_parent)
                after_wrapper.append(after_run)
                elements_to_insert.append(after_wrapper)
            else:
                elements_to_insert.append(after_run)

        # Add preserved content from wrappers (after)
        elements_to_insert.extend(preserved_after)

        # Insert all elements at the insertion point
        for i, elem in enumerate(elements_to_insert):
            paragraph.insert(insertion_index + i, elem)

    def _find_paragraph_insertion_index(
        self, match: TextSpan, paragraph: Any, start_run: Any
    ) -> int:
        """Find the insertion index in the paragraph for replacement elements.

        This handles the case where the start_run might be inside a wrapper,
        finding the appropriate position at the paragraph level.

        Args:
            match: The TextSpan being replaced
            paragraph: The paragraph element
            start_run: The first run in the match

        Returns:
            The index in the paragraph where elements should be inserted
        """
        start_parent = start_run.getparent()

        if start_parent == paragraph:
            # Start run is direct child of paragraph
            return list(paragraph).index(start_run)
        elif self._is_tracked_change_wrapper(start_parent):
            # Start run is inside a wrapper - find wrapper's position
            try:
                return list(paragraph).index(start_parent)
            except ValueError:
                # Wrapper not found - scan for first match run's position
                pass

        # Fallback: find the first paragraph-level element that contains
        # or precedes any of the matched runs
        for i, child in enumerate(paragraph):
            if child == start_run:
                return i
            if self._is_tracked_change_wrapper(child):
                for run_idx in range(match.start_run_index, match.end_run_index + 1):
                    if match.runs[run_idx] in child:
                        return i

        # Last resort: append at end
        return len(list(paragraph))

    def _remove_match(self, match: TextSpan) -> None:
        """Remove matched text without creating tracked change markers.

        This is used for untracked deletion - the text is simply removed
        from the document without creating a <w:del> wrapper.

        Handles:
        - Single run matches (remove run entirely, or split if partial)
        - Multi-run matches (remove all matched runs, preserve text before/after)
        - Runs inside tracked change wrappers (preserves wrapper structure)

        Args:
            match: TextSpan object representing the text to remove
        """
        paragraph = match.paragraph

        # If the match is within a single run
        if match.start_run_index == match.end_run_index:
            run = match.runs[match.start_run_index]
            run_text = self._get_run_text_content(run)

            # Get the actual parent of the run (may be paragraph or a wrapper)
            actual_parent = run.getparent()
            if actual_parent is None:
                actual_parent = paragraph

            # If the match is the entire run, just remove the run
            if match.start_offset == 0 and match.end_offset == len(run_text):
                if run in actual_parent:
                    actual_parent.remove(run)
                    # Clean up empty wrappers
                    if self._is_tracked_change_wrapper(actual_parent) and len(actual_parent) == 0:
                        wrapper_parent = actual_parent.getparent()
                        if wrapper_parent is not None:
                            wrapper_parent.remove(actual_parent)
            else:
                # Match is partial - need to split the run and remove middle portion
                self._split_and_remove_in_run(paragraph, run, match.start_offset, match.end_offset)
        else:
            # Match spans multiple runs - need to preserve text before/after match
            self._remove_multirun_match(match, paragraph)

    def _split_and_remove_in_run(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
    ) -> None:
        """Split a run and remove a portion without replacement.

        This is similar to _split_and_replace_in_run but doesn't insert
        any replacement elements - just preserves the before/after text.

        Args:
            paragraph: The paragraph containing the run
            run: The run to split
            start_offset: Character offset where match starts
            end_offset: Character offset where match ends (exclusive)
        """
        run_text, num_elements = self._get_run_text_info(run)
        if num_elements == 0:
            return

        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        # Build replacement elements (before and after runs only)
        new_elements = []
        if before_text:
            new_elements.append(self._create_text_run(before_text, run))
        if after_text:
            new_elements.append(self._create_text_run(after_text, run))

        # Replace the run with the new elements (or nothing if both empty)
        self._replace_run_with_elements(paragraph, run, new_elements)

    def _remove_multirun_match(self, match: TextSpan, paragraph: Any) -> None:
        """Remove a match spanning multiple runs without replacement.

        This is similar to _replace_multirun_match_with_elements but doesn't
        insert any replacement elements - just preserves surrounding text.

        Args:
            match: TextSpan object representing the text to remove
            paragraph: The paragraph containing the runs
        """
        start_run = match.runs[match.start_run_index]
        end_run = match.runs[match.end_run_index]

        # Save parent information BEFORE removing runs
        start_run_parent = start_run.getparent()
        end_run_parent = end_run.getparent()

        # Get text before match in the first run
        first_run_text = self._get_run_text_content(start_run)
        before_text = first_run_text[: match.start_offset]

        # Get text after match in the last run
        last_run_text = self._get_run_text_content(end_run)
        after_text = last_run_text[match.end_offset :]

        # Track wrappers that need remaining content preserved
        wrapper_remaining_content: dict[Any, tuple[list[Any], list[Any]]] = {}

        # Analyze each run's parent and track wrapper content
        for i in range(match.start_run_index, match.end_run_index + 1):
            run = match.runs[i]
            run_parent = run.getparent()

            if run_parent is not None and self._is_tracked_change_wrapper(run_parent):
                wrapper = run_parent
                wrapper_children = list(wrapper)
                run_idx_in_wrapper = wrapper_children.index(run)

                if wrapper not in wrapper_remaining_content:
                    wrapper_remaining_content[wrapper] = ([], [])

                before_runs, after_runs = wrapper_remaining_content[wrapper]

                # Save content before first matched run in wrapper
                if not before_runs and run_idx_in_wrapper > 0:
                    is_first_matched_in_wrapper = True
                    for j in range(match.start_run_index, i):
                        if match.runs[j].getparent() == wrapper:
                            is_first_matched_in_wrapper = False
                            break
                    if is_first_matched_in_wrapper:
                        for k in range(run_idx_in_wrapper):
                            before_runs.append(wrapper_children[k])

                # Save content after last matched run in wrapper
                is_last_matched_in_wrapper = True
                for j in range(i + 1, match.end_run_index + 1):
                    if match.runs[j].getparent() == wrapper:
                        is_last_matched_in_wrapper = False
                        break
                if is_last_matched_in_wrapper and run_idx_in_wrapper < len(wrapper_children) - 1:
                    for k in range(run_idx_in_wrapper + 1, len(wrapper_children)):
                        after_runs.append(wrapper_children[k])

                wrapper_remaining_content[wrapper] = (before_runs, after_runs)

        # Find insertion point in paragraph
        insertion_index = self._find_paragraph_insertion_index(match, paragraph, start_run)

        # Remove all runs in the match from their parents
        for i in range(match.start_run_index, match.end_run_index + 1):
            run = match.runs[i]
            run_parent = run.getparent()
            if run_parent is not None and run in run_parent:
                run_parent.remove(run)

        # Handle empty wrappers and preserved content
        preserved_before: list[Any] = []
        preserved_after: list[Any] = []

        for wrapper, (before_runs, after_runs) in wrapper_remaining_content.items():
            wrapper_still_has_content = len(list(wrapper)) > 0

            if not wrapper_still_has_content:
                if wrapper.getparent() is not None:
                    wrapper.getparent().remove(wrapper)

            if before_runs:
                new_before_wrapper = self._clone_wrapper(wrapper)
                for run in before_runs:
                    run_copy = etree.fromstring(etree.tostring(run))
                    new_before_wrapper.append(run_copy)
                preserved_before.append(new_before_wrapper)

            if after_runs:
                new_after_wrapper = self._clone_wrapper(wrapper)
                for run in after_runs:
                    run_copy = etree.fromstring(etree.tostring(run))
                    new_after_wrapper.append(run_copy)
                preserved_after.append(new_after_wrapper)

        # Build elements to insert (just before/after text, no replacement)
        elements_to_insert: list[Any] = []

        # Add preserved content from wrappers (before)
        elements_to_insert.extend(preserved_before)

        # Add before_text run if needed
        if before_text:
            before_run = self._create_text_run(before_text, start_run)
            if start_run_parent is not None and self._is_tracked_change_wrapper(start_run_parent):
                before_wrapper = self._clone_wrapper(start_run_parent)
                before_wrapper.append(before_run)
                elements_to_insert.append(before_wrapper)
            else:
                elements_to_insert.append(before_run)

        # Add after_text run if needed
        if after_text:
            after_run = self._create_text_run(after_text, end_run)
            if end_run_parent is not None and self._is_tracked_change_wrapper(end_run_parent):
                after_wrapper = self._clone_wrapper(end_run_parent)
                after_wrapper.append(after_run)
                elements_to_insert.append(after_wrapper)
            else:
                elements_to_insert.append(after_run)

        # Add preserved content from wrappers (after)
        elements_to_insert.extend(preserved_after)

        # Insert all elements at the insertion point
        for i, elem in enumerate(elements_to_insert):
            paragraph.insert(insertion_index + i, elem)
