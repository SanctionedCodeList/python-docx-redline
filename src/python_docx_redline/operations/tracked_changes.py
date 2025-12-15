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

    def insert(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
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
            regex: Whether to treat anchor as a regex pattern (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)

        Raises:
            ValueError: If both 'after' and 'before' are specified, or if neither is specified
            TextNotFoundError: If the anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
            re.error: If regex=True and the pattern is invalid
        """
        # Validate parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        # Determine anchor text and insertion mode
        anchor: str = after if after is not None else before  # type: ignore[assignment]
        insert_after = after is not None

        # Get all paragraphs in the document
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search with optional quote normalization
        matches = self._document._text_search.find_text(
            anchor,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not matches:
            # Generate smart suggestions
            suggestions = SuggestionGenerator.generate_suggestions(anchor, paragraphs)
            raise TextNotFoundError(anchor, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(anchor, matches)

        # We have exactly one match
        match = matches[0]

        # Generate the insertion XML
        insertion_xml = self._document._xml_generator.create_insertion(text, author)

        # Parse the insertion XML with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {insertion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        insertion_element = root[0]  # Get the first child (the actual insertion)

        # Insert at the appropriate position
        if insert_after:
            self._insert_after_match(match, insertion_element)
        else:
            self._insert_before_match(match, insertion_element)

    def delete(
        self,
        text: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Delete text with tracked changes.

        This method searches for the specified text in the document and marks
        it as a tracked deletion.

        Args:
            text: The text or regex pattern to delete
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            regex: Whether to treat 'text' as a regex pattern (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)

        Raises:
            TextNotFoundError: If the text is not found
            AmbiguousTextError: If multiple occurrences of text are found
            re.error: If regex=True and the pattern is invalid
        """
        # Get all paragraphs in the document
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search with optional quote normalization
        matches = self._document._text_search.find_text(
            text,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not matches:
            # Generate smart suggestions
            suggestions = SuggestionGenerator.generate_suggestions(text, paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(text, matches)

        # We have exactly one match
        match = matches[0]

        # Generate the deletion XML
        deletion_xml = self._document._xml_generator.create_deletion(text, author)

        # Parse the deletion XML with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        deletion_element = root[0]  # Get the first child (the actual deletion)

        # Replace the matched text with deletion
        self._replace_match_with_element(match, deletion_element)

    def replace(
        self,
        find: str,
        replace: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
        show_context: bool = False,
        check_continuity: bool = False,
        context_chars: int = 50,
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
            regex: Whether to treat 'find' as a regex pattern (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)
            show_context: Show text before/after the match for preview (default: False)
            check_continuity: Check if replacement may create sentence fragments (default: False)
            context_chars: Number of characters to show before/after when show_context=True
                (default: 50)

        Raises:
            TextNotFoundError: If the 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
            re.error: If regex=True and the pattern is invalid

        Warnings:
            ContinuityWarning: If check_continuity=True and potential sentence fragment detected
        """
        # Get all paragraphs in the document
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search with optional quote normalization
        matches = self._document._text_search.find_text(
            find,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not matches:
            # Generate smart suggestions
            suggestions = SuggestionGenerator.generate_suggestions(find, paragraphs)
            raise TextNotFoundError(find, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(find, matches)

        # We have exactly one match
        match = matches[0]

        # Get the actual matched text for deletion
        matched_text = match.text

        # Show context preview if requested
        if show_context:
            before, matched, after = self._document._get_detailed_context(match, context_chars)
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
                len(replace),
                replace,
            )

        # Handle capture group expansion for regex replacements
        if regex and match.match_obj:
            # Use expand() to handle capture group references like \1, \2, etc.
            replacement_text = match.match_obj.expand(replace)
        else:
            replacement_text = replace

        # Check continuity if requested
        if check_continuity:
            _, _, after_text = self._document._get_detailed_context(match, context_chars)
            continuity_warnings = self._document._check_continuity(replacement_text, after_text)

            if continuity_warnings:
                import warnings

                from ..errors import ContinuityWarning

                for warning_msg in continuity_warnings:
                    suggestions = [
                        "Include more context in your replacement text",
                        "Adjust the 'find' text to include the connecting phrase",
                        "Review the following text to ensure grammatical correctness",
                    ]
                    warnings.warn(
                        ContinuityWarning(warning_msg, after_text, suggestions),
                        stacklevel=2,
                    )

        # Generate deletion XML for the old text (use actual matched text)
        deletion_xml = self._document._xml_generator.create_deletion(matched_text, author)

        # Generate insertion XML for the new text (with expanded capture groups if regex)
        insertion_xml = self._document._xml_generator.create_insertion(replacement_text, author)

        # Parse both XMLs with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
    {insertion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        deletion_element = root[0]  # First child (deletion)
        insertion_element = root[1]  # Second child (insertion)

        # Replace the matched text with deletion + insertion
        self._replace_match_with_elements(match, [deletion_element, insertion_element])

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

        # Get all paragraphs
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # 1. Find the source text to move
        source_paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, source_scope)
        source_matches = self._document._text_search.find_text(
            text,
            source_paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not source_matches:
            suggestions = SuggestionGenerator.generate_suggestions(text, source_paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        if len(source_matches) > 1:
            raise AmbiguousTextError(text, source_matches)

        source_match = source_matches[0]
        source_text = source_match.text

        # 2. Find the destination anchor
        dest_paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, dest_scope)
        dest_matches = self._document._text_search.find_text(
            dest_anchor,
            dest_paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not dest_matches:
            suggestions = SuggestionGenerator.generate_suggestions(dest_anchor, dest_paragraphs)
            raise TextNotFoundError(dest_anchor, suggestions=suggestions)

        if len(dest_matches) > 1:
            raise AmbiguousTextError(dest_anchor, dest_matches)

        dest_match = dest_matches[0]

        # 3. Generate a unique move name to link source and destination
        move_name = self._generate_move_name()

        # 4. Generate moveFrom XML (for source location)
        move_from_xml, _, _ = self._document._xml_generator.create_move_from(
            source_text, move_name, author
        )

        # 5. Generate moveTo XML (for destination location)
        move_to_xml, _, _ = self._document._xml_generator.create_move_to(
            source_text, move_name, author
        )

        # 6. Parse both XMLs with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {move_to_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        # Get all three elements (moveToRangeStart, moveTo, moveToRangeEnd)
        move_to_elements = list(root)

        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {move_from_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        # Get all three elements (moveFromRangeStart, moveFrom, moveFromRangeEnd)
        move_from_elements = list(root)

        # 7. First, insert the moveTo at the destination
        # (do this first so we don't mess up source position)
        if insert_after:
            self._insert_after_match(dest_match, move_to_elements)
        else:
            self._insert_before_match(dest_match, move_to_elements)

        # 8. Replace the source text with moveFrom markers
        self._replace_match_with_elements(source_match, move_from_elements)

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
        # Get the full text of the run
        text_elements = list(run.iter(f"{{{WORD_NAMESPACE}}}t"))
        if not text_elements:
            return

        # For simplicity, we'll work with the first text element
        # (Word typically has one w:t per run)
        text_elem = text_elements[0]
        run_text = text_elem.text or ""

        # Split into before, match, after
        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        run_index = list(paragraph).index(run)

        # Build new elements
        new_elements = []

        # Add before text if it exists
        if before_text:
            before_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            # Copy run properties if they exist
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                before_run.append(etree.fromstring(etree.tostring(run_props)))
            before_t = etree.SubElement(before_run, f"{{{WORD_NAMESPACE}}}t")
            if before_text and (before_text[0].isspace() or before_text[-1].isspace()):
                before_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            before_t.text = before_text
            new_elements.append(before_run)

        # Add replacement element
        new_elements.append(replacement_element)

        # Add after text if it exists
        if after_text:
            after_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            # Copy run properties if they exist
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                after_run.append(etree.fromstring(etree.tostring(run_props)))
            after_t = etree.SubElement(after_run, f"{{{WORD_NAMESPACE}}}t")
            if after_text and (after_text[0].isspace() or after_text[-1].isspace()):
                after_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            after_t.text = after_text
            new_elements.append(after_run)

        # Remove original run
        paragraph.remove(run)

        # Insert new elements
        for i, elem in enumerate(new_elements):
            paragraph.insert(run_index + i, elem)

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
        # Get the full text of the run
        text_elements = list(run.iter(f"{{{WORD_NAMESPACE}}}t"))
        if not text_elements:
            return

        text_elem = text_elements[0]
        run_text = text_elem.text or ""

        # Split into before, match, after
        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        run_index = list(paragraph).index(run)

        # Build new elements
        new_elements = []

        # Add before text if it exists
        if before_text:
            before_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                before_run.append(etree.fromstring(etree.tostring(run_props)))
            before_t = etree.SubElement(before_run, f"{{{WORD_NAMESPACE}}}t")
            if before_text and (before_text[0].isspace() or before_text[-1].isspace()):
                before_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            before_t.text = before_text
            new_elements.append(before_run)

        # Add all replacement elements
        new_elements.extend(replacement_elements)

        # Add after text if it exists
        if after_text:
            after_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                after_run.append(etree.fromstring(etree.tostring(run_props)))
            after_t = etree.SubElement(after_run, f"{{{WORD_NAMESPACE}}}t")
            if after_text and (after_text[0].isspace() or after_text[-1].isspace()):
                after_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            after_t.text = after_text
            new_elements.append(after_run)

        # Remove original run
        paragraph.remove(run)

        # Insert new elements
        for i, elem in enumerate(new_elements):
            paragraph.insert(run_index + i, elem)
