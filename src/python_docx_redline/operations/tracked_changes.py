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
        enable_quote_normalization: bool,
    ) -> TextSpan:
        """Find a unique text match in the document.

        Args:
            text: The text or regex pattern to find
            scope: Limit search scope
            regex: Whether to treat text as regex
            enable_quote_normalization: Whether to normalize quotes

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
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
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
        enable_quote_normalization: bool = True,
        fuzzy: float | dict[str, Any] | None = None,
    ) -> None:
        """Insert text with tracked changes after or before a specific location.

        This method searches for the anchor text in the document and inserts
        the new text either immediately after it or immediately before it as
        a tracked insertion.

        Args:
            text: The text to insert
            after: The text or regex pattern to insert after (optional)
            before: The text or regex pattern to insert before (optional)
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            occurrence: Which occurrence(s) to insert at: 1 (first), 2 (second), "first",
                "last", "all", or list of indices [1, 3, 5] (default: "first")
            regex: Whether to treat anchor as a regex pattern (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'

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
            normalize_quotes_for_matching=enable_quote_normalization
            and not regex
            and not fuzzy_config,
            fuzzy=fuzzy_config,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(anchor, paragraphs)
            raise TextNotFoundError(anchor, suggestions=suggestions)

        # Select target matches based on occurrence
        target_matches = self._select_matches(matches, occurrence, anchor)

        # Generate insertion XML
        insertion_xml = self._document._xml_generator.create_insertion(text, author)

        # Insert at each target match (process in reverse to preserve indices)
        for match in reversed(target_matches):
            elements = self._parse_xml_elements(insertion_xml)
            insertion_element = elements[0]

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
        enable_quote_normalization: bool = True,
        fuzzy: float | dict[str, Any] | None = None,
    ) -> None:
        """Delete text with tracked changes.

        This method searches for the specified text in the document and marks
        it as a tracked deletion.

        Args:
            text: The text or regex pattern to delete
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            occurrence: Which occurrence(s) to delete: 1 (first), 2 (second), "first", "last",
                "all", or list of indices [1, 3, 5] (default: "first")
            regex: Whether to treat 'text' as a regex pattern (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'

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
            normalize_quotes_for_matching=enable_quote_normalization
            and not regex
            and not fuzzy_config,
            fuzzy=fuzzy_config,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(text, paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        # Select target matches based on occurrence
        target_matches = self._select_matches(matches, occurrence, text)

        # Delete each target match (process in reverse to preserve indices)
        for match in reversed(target_matches):
            # Generate and parse the deletion XML
            deletion_xml = self._document._xml_generator.create_deletion(match.text, author)
            elements = self._parse_xml_elements(deletion_xml)
            deletion_element = elements[0]

            # Replace the matched text with deletion
            self._replace_match_with_element(match, deletion_element)

    def replace(
        self,
        find: str,
        replace: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        enable_quote_normalization: bool = True,
        show_context: bool = False,
        check_continuity: bool = False,
        context_chars: int = 50,
        fuzzy: float | dict[str, Any] | None = None,
    ) -> None:
        """Find and replace text with tracked changes.

        This method searches for text and replaces it with new text, showing
        both the deletion of the old text and insertion of the new text as
        tracked changes.

        When regex=True, the replacement string can use capture groups:
        - \\1, \\2, etc. for numbered groups
        - \\g<name> for named groups

        Args:
            find: Text or regex pattern to find
            replace: Replacement text (can include capture group references if regex=True)
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            occurrence: Which occurrence(s) to replace: 1 (first), 2 (second), "first", "last",
                "all", or list of indices [1, 3, 5] (default: "first")
            regex: Whether to treat 'find' as a regex pattern (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)
            show_context: Show text before/after the match for preview (default: False)
            check_continuity: Check if replacement may create sentence fragments (default: False)
            context_chars: Number of characters to show before/after when show_context=True
                (default: 50)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'

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
            normalize_quotes_for_matching=enable_quote_normalization
            and not regex
            and not fuzzy_config,
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

            # Generate and parse deletion + insertion XMLs
            deletion_xml = self._document._xml_generator.create_deletion(matched_text, author)
            insertion_xml = self._document._xml_generator.create_insertion(replacement_text, author)
            elements = self._parse_xml_elements(f"{deletion_xml}\n    {insertion_xml}")

            # Replace the matched text with deletion + insertion
            self._replace_match_with_elements(match, elements)

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
        enable_quote_normalization: bool = True,
    ) -> None:
        """Move text to a new location with proper move tracking.

        Unlike delete + insert, move tracking creates linked markers that show
        the text was relocated rather than deleted and re-added. This provides
        better context for document reviewers in Word.

        In Word's track changes view:
        - Source location shows text with strikethrough and "Moved" annotation
        - Destination shows text with underline and "Moved" annotation
        - Both locations are linked with matching move markers

        Args:
            text: The text to move (or regex pattern if regex=True)
            after: Text to insert the moved content after (at destination)
            before: Text to insert the moved content before (at destination)
            author: Optional author override (uses document author if None)
            source_scope: Limit source text search scope
            dest_scope: Limit destination anchor search scope
            regex: Whether to treat 'text' and anchor as regex patterns (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)

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
        source_match = self._find_unique_match(
            text, source_scope, regex, enable_quote_normalization
        )
        dest_match = self._find_unique_match(
            dest_anchor, dest_scope, regex, enable_quote_normalization
        )

        source_text = source_match.text

        # Generate XML elements for move operation
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

        Args:
            match: TextSpan object representing the text to replace
            replacement_element: The lxml Element to insert in place of matched text
        """
        paragraph = match.paragraph

        # If the match is within a single run
        if match.start_run_index == match.end_run_index:
            run = match.runs[match.start_run_index]
            run_text = "".join(run.itertext())

            # If the match is the entire run, replace the run
            if match.start_offset == 0 and match.end_offset == len(run_text):
                run_index = list(paragraph).index(run)
                paragraph.remove(run)
                paragraph.insert(run_index, replacement_element)
            else:
                # Match is partial - need to split the run
                self._split_and_replace_in_run(
                    paragraph, run, match.start_offset, match.end_offset, replacement_element
                )
        else:
            # Match spans multiple runs - remove all matched runs and insert replacement
            start_run = match.runs[match.start_run_index]
            start_run_index = list(paragraph).index(start_run)

            # Remove all runs in the match
            for i in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[i]
                if run in paragraph:
                    paragraph.remove(run)

            # Insert replacement at the position of the first removed run
            paragraph.insert(start_run_index, replacement_element)

    def _replace_match_with_elements(
        self, match: TextSpan, replacement_elements: list[Any]
    ) -> None:
        """Replace matched text with multiple XML elements.

        Used for replace_tracked which needs both deletion and insertion elements.

        Args:
            match: TextSpan object representing the text to replace
            replacement_elements: List of lxml Elements to insert in place of matched text
        """
        paragraph = match.paragraph

        # Similar to _replace_match_with_element but inserts multiple elements
        if match.start_run_index == match.end_run_index:
            run = match.runs[match.start_run_index]
            run_text = "".join(run.itertext())

            # If the match is the entire run, replace the run
            if match.start_offset == 0 and match.end_offset == len(run_text):
                run_index = list(paragraph).index(run)
                paragraph.remove(run)
                # Insert elements in order
                for i, elem in enumerate(replacement_elements):
                    paragraph.insert(run_index + i, elem)
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
            # Match spans multiple runs
            start_run = match.runs[match.start_run_index]
            start_run_index = list(paragraph).index(start_run)

            # Remove all runs in the match
            for i in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[i]
                if run in paragraph:
                    paragraph.remove(run)

            # Insert all replacement elements at the position of the first removed run
            for i, elem in enumerate(replacement_elements):
                paragraph.insert(start_run_index + i, elem)

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

        Args:
            paragraph: The paragraph containing the run
            run: The run to replace
            new_elements: Elements to insert in place of the run
        """
        run_index = list(paragraph).index(run)
        paragraph.remove(run)
        for i, elem in enumerate(new_elements):
            paragraph.insert(run_index + i, elem)

    def _split_and_replace_in_run(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_element: Any,
    ) -> None:
        """Split a run and replace a portion with a new element.

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
        new_elements = self._build_split_elements(
            run, before_text, after_text, [replacement_element]
        )
        self._replace_run_with_elements(paragraph, run, new_elements)

    def _split_and_replace_in_run_multiple(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_elements: list[Any],
    ) -> None:
        """Split a run and replace a portion with multiple new elements.

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
        new_elements = self._build_split_elements(
            run, before_text, after_text, replacement_elements
        )
        self._replace_run_with_elements(paragraph, run, new_elements)
