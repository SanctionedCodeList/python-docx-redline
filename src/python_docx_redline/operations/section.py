"""
SectionOperations class for handling section and paragraph operations.

This module provides a dedicated class for section-related operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import TYPE_CHECKING, Any

from lxml import etree

from ..constants import WORD_NAMESPACE
from ..errors import AmbiguousTextError, TextNotFoundError
from ..scope import ScopeEvaluator
from ..suggestions import SuggestionGenerator

if TYPE_CHECKING:
    from ..document import Document
    from ..models.paragraph import Paragraph
    from ..models.section import Section


class SectionOperations:
    """Handles section and paragraph operations.

    This class encapsulates all section/paragraph functionality, including:
    - Inserting paragraphs with tracked changes
    - Deleting sections with tracked changes
    - Normalizing currency and date formats
    - Updating section references

    The class takes a Document reference and operates on its XML structure.

    Example:
        >>> # Usually accessed through Document
        >>> doc = Document("contract.docx")
        >>> doc.insert_paragraph("New clause text", after="Section 2.1")
        >>> doc.delete_section("Outdated Section")
    """

    def __init__(self, document: Document) -> None:
        """Initialize SectionOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    def insert_paragraph(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        style: str | None = None,
        track: bool = True,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> Paragraph:
        """Insert a complete new paragraph with tracked changes.

        Args:
            text: Text content for the new paragraph
            after: Text to search for as insertion point (insert after this)
            before: Text to search for as insertion point (insert before this)
            style: Paragraph style (e.g., 'Normal', 'Heading1')
            track: Whether to track this insertion (default True)
            author: Optional author override (uses document author if None)
            scope: Limit search scope for anchor text

        Returns:
            The created Paragraph object

        Raises:
            ValueError: If neither 'after' nor 'before' is specified, or both are
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
        """
        from ..models.paragraph import Paragraph

        # Validate arguments
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before'")
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before'")

        anchor_text = after if after is not None else before
        insert_after = after is not None

        # After validation, anchor_text is guaranteed to be a string
        assert anchor_text is not None

        # Get all paragraphs
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find the anchor paragraph
        matches = self._document._text_search.find_text(anchor_text, paragraphs)

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(anchor_text, paragraphs)
            raise TextNotFoundError(anchor_text, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(anchor_text, matches)

        match = matches[0]
        anchor_paragraph = match.paragraph

        # Create new paragraph element
        new_p = etree.Element(f"{{{WORD_NAMESPACE}}}p")

        # Add style if specified
        if style:
            p_pr = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}pPr")
            p_style = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}pStyle")
            p_style.set(f"{{{WORD_NAMESPACE}}}val", style)

        # If tracked, wrap the runs in w:ins
        if track:
            from datetime import timezone

            author_name = author if author is not None else self._document.author
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._document._xml_generator.next_change_id
            self._document._xml_generator.next_change_id += 1

            # Create w:ins element to wrap the run
            ins = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}ins")
            ins.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
            ins.set(f"{{{WORD_NAMESPACE}}}author", author_name)
            ins.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
            ins.set(
                "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc",
                timestamp,
            )

            # Add text content inside the w:ins element
            run = etree.SubElement(ins, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = text
        else:
            # Add text content directly to paragraph
            run = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = text

        element_to_insert = new_p

        # Insert the paragraph in the document
        parent = anchor_paragraph.getparent()
        if parent is None:
            raise ValueError("Anchor paragraph has no parent")

        anchor_index = list(parent).index(anchor_paragraph)

        if insert_after:
            # Insert after anchor
            parent.insert(anchor_index + 1, element_to_insert)
        else:
            # Insert before anchor
            parent.insert(anchor_index, element_to_insert)

        # Return Paragraph wrapper
        # new_p is always the actual paragraph element (whether tracked or not)
        return Paragraph(new_p)

    def insert_paragraphs(
        self,
        texts: list[str],
        after: str | None = None,
        before: str | None = None,
        style: str | None = None,
        track: bool = True,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> list[Paragraph]:
        """Insert multiple paragraphs with tracked changes.

        This is more efficient than calling insert_paragraph() multiple times
        as it maintains proper ordering and positioning.

        Args:
            texts: List of text content for new paragraphs
            after: Text to search for as insertion point (insert after this)
            before: Text to search for as insertion point (insert before this)
            style: Paragraph style for all paragraphs (e.g., 'Normal', 'Heading1')
            track: Whether to track these insertions (default True)
            author: Optional author override (uses document author if None)
            scope: Limit search scope for anchor text

        Returns:
            List of created Paragraph objects

        Raises:
            ValueError: If neither 'after' nor 'before' is specified, or both are
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
        """
        from ..models.paragraph import Paragraph as ParagraphClass

        if not texts:
            return []

        # Insert the first paragraph to find the anchor position
        first_para = self.insert_paragraph(
            texts[0],
            after=after,
            before=before,
            style=style,
            track=track,
            author=author,
            scope=scope,
        )

        created_paragraphs = [first_para]

        # Get the parent of the first paragraph
        parent = first_para.element.getparent()
        if parent is None:
            raise ValueError("First paragraph has no parent")
        insertion_index = list(parent).index(first_para.element)

        # Insert remaining paragraphs after the first one
        for i, para_text in enumerate(texts[1:], start=1):
            # Create new paragraph element
            new_p = etree.Element(f"{{{WORD_NAMESPACE}}}p")

            # Add style if specified
            if style:
                p_pr = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}pPr")
                p_style = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}pStyle")
                p_style.set(f"{{{WORD_NAMESPACE}}}val", style)

            # If tracked, wrap the runs in w:ins
            if track:
                from datetime import timezone

                author_name = author if author is not None else self._document.author
                timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                change_id = self._document._xml_generator.next_change_id
                self._document._xml_generator.next_change_id += 1

                # Create w:ins element to wrap the run
                ins = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}ins")
                ins.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                ins.set(f"{{{WORD_NAMESPACE}}}author", author_name)
                ins.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
                ins.set(
                    "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc",
                    timestamp,
                )

                # Add text content inside the w:ins element
                run = etree.SubElement(ins, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = para_text
            else:
                # Add text content directly to paragraph
                run = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = para_text

            # Insert after previous paragraph
            parent.insert(insertion_index + i, new_p)
            created_paragraphs.append(ParagraphClass(new_p))

        return created_paragraphs

    def delete_section(
        self,
        heading: str,
        track: bool = True,
        update_toc: bool = False,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> Section:
        """Delete an entire section by heading text.

        Args:
            heading: Heading text of section to delete
            track: Delete as tracked change (default True)
            update_toc: No-op, kept for API compatibility. TOC updates require
                opening the document in Word.
            author: Author name for tracked changes
            scope: Limit search scope

        Returns:
            Section object representing the deleted section

        Raises:
            TextNotFoundError: If heading not found
            AmbiguousTextError: If multiple sections match

        Examples:
            >>> doc.delete_section("Methods", track=True)
            >>> doc.delete_section("Outdated Section", track=False)
        """
        from ..models.section import Section

        all_sections = Section.from_document(self._document.xml_root)
        all_sections = self._filter_sections_by_scope(all_sections, scope)
        section = self._find_single_section_match(all_sections, heading)

        if track:
            self._delete_section_tracked(section, author)
        else:
            self._delete_section_untracked(section)

        return section

    def _filter_sections_by_scope(
        self, sections: list[Section], scope: str | dict | Any | None
    ) -> list[Section]:
        """Filter sections by scope, keeping those with paragraphs in scope."""
        if scope is None:
            return sections
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs_in_scope = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)
        scope_para_set = set(paragraphs_in_scope)
        return [s for s in sections if any(p.element in scope_para_set for p in s.paragraphs)]

    def _find_single_section_match(self, sections: list[Section], heading: str) -> Section:
        """Find exactly one section matching the heading, raising errors otherwise."""
        matches = [
            s
            for s in sections
            if s.heading is not None and s.contains(heading, case_sensitive=False)
        ]

        if not matches:
            heading_paragraphs = [s.heading.element for s in sections if s.heading is not None]
            suggestions = SuggestionGenerator.generate_suggestions(heading, heading_paragraphs)
            raise TextNotFoundError(heading, suggestions=suggestions)

        if len(matches) > 1:
            self._raise_ambiguous_section_error(matches, heading)

        return matches[0]

    def _raise_ambiguous_section_error(self, matches: list[Section], heading: str) -> None:
        """Raise AmbiguousTextError with TextSpan representations of matching sections."""
        from ..text_search import TextSpan

        match_spans = []
        for section in matches:
            if section.heading:
                runs = list(section.heading.element.iter(f"{{{WORD_NAMESPACE}}}r"))
                if runs:
                    heading_text = section.heading_text or ""
                    span = TextSpan(
                        runs=runs,
                        start_run_index=0,
                        end_run_index=len(runs) - 1,
                        start_offset=0,
                        end_offset=len(heading_text.strip()),
                        paragraph=section.heading.element,
                    )
                    match_spans.append(span)

        raise AmbiguousTextError(heading, match_spans)

    def _delete_section_tracked(self, section: Section, author: str | None) -> None:
        """Delete section paragraphs with tracked changes."""
        from datetime import timezone

        author_name = author if author is not None else self._document.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        for para in section.paragraphs:
            runs = list(para.element.iter(f"{{{WORD_NAMESPACE}}}r"))
            if not runs:
                continue
            del_elem = self._create_deletion_element(author_name, timestamp)
            self._wrap_runs_in_deletion(para.element, runs, del_elem)

    def _create_deletion_element(self, author: str, timestamp: str) -> Any:
        """Create a w:del element for tracked deletion."""
        change_id = self._document._xml_generator.next_change_id
        self._document._xml_generator.next_change_id += 1

        del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
        del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
        del_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
        del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
        del_elem.set(
            "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc",
            timestamp,
        )
        return del_elem

    def _wrap_runs_in_deletion(self, para_element: Any, runs: list[Any], del_elem: Any) -> None:
        """Wrap runs in a deletion element, converting w:t to w:delText."""
        for run in runs:
            run_parent = run.getparent()
            if run_parent is not None:
                run_parent.remove(run)
            self._convert_text_to_deltext(run)
            del_elem.append(run)

        p_pr = para_element.find(f"{{{WORD_NAMESPACE}}}pPr")
        if p_pr is not None:
            p_pr_index = list(para_element).index(p_pr)
            para_element.insert(p_pr_index + 1, del_elem)
        else:
            para_element.insert(0, del_elem)

    def _convert_text_to_deltext(self, run: Any) -> None:
        """Convert w:t elements in a run to w:delText."""
        for t_elem in run.iter(f"{{{WORD_NAMESPACE}}}t"):
            deltext = etree.Element(f"{{{WORD_NAMESPACE}}}delText")
            deltext.text = t_elem.text
            xml_space = t_elem.get("{http://www.w3.org/XML/1998/namespace}space")
            if xml_space:
                deltext.set("{http://www.w3.org/XML/1998/namespace}space", xml_space)
            t_parent = t_elem.getparent()
            t_index = list(t_parent).index(t_elem)
            t_parent.remove(t_elem)
            t_parent.insert(t_index, deltext)

    def _delete_section_untracked(self, section: Section) -> None:
        """Delete section paragraphs without tracking changes."""
        for para in section.paragraphs:
            parent = para.element.getparent()
            if parent is not None:
                parent.remove(para.element)

    def normalize_currency(
        self,
        currency_symbol: str = "$",
        decimal_places: int = 2,
        thousands_separator: bool = True,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Normalize currency amounts to a consistent format with tracked changes.

        Finds various currency formats and normalizes them to a standard format.
        This reduces manual regex work and prevents formatting inconsistencies.

        Detected formats:
            - $100, $100.0 → $100.00
            - $1000 → $1,000.00 (if thousands_separator=True)
            - $1,000 → $1,000.00

        Args:
            currency_symbol: The currency symbol to use (default: "$")
            decimal_places: Number of decimal places (default: 2)
            thousands_separator: Whether to include thousands separators (default: True)
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})

        Returns:
            Number of currency amounts normalized

        Example:
            >>> # Normalize all $ amounts to $X,XXX.XX format
            >>> count = doc.normalize_currency()
            >>>
            >>> # Normalize to £X.XX without thousands separator
            >>> count = doc.normalize_currency("£", thousands_separator=False)
        """
        # Build regex pattern for currency amounts
        # Matches: $100, $100.00, $1,000, $1,000.50, etc.
        pattern = rf"{re.escape(currency_symbol)}\d{{1,3}}(?:,?\d{{3}})*(?:\.\d+)?"

        # Get all paragraphs
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find all currency matches
        self._document._text_search.find_text(
            pattern,
            paragraphs,
            regex=True,
            normalize_quotes_for_matching=False,
        )

        # Helper to format amount
        def format_amount(amount_str: str) -> str:
            amount = float(amount_str.replace(",", ""))
            formatted = f"{amount:.{decimal_places}f}"
            if thousands_separator and "." in formatted:
                integer_part, decimal_part = formatted.split(".")
                integer_with_commas = f"{int(integer_part):,}"
                return f"{integer_with_commas}.{decimal_part}"
            elif thousands_separator:
                formatted_int = f"{int(float(formatted)):,}"
                if decimal_places > 0:
                    return formatted_int + "." + "0" * decimal_places
                return formatted_int
            return formatted

        # Process one match at a time to avoid XML reference issues
        replacement_count = 0
        max_iterations = 100  # Prevent infinite loop

        for _ in range(max_iterations):
            # Get fresh paragraphs and matches each iteration
            all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
            paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)
            matches = self._document._text_search.find_text(
                pattern,
                paragraphs,
                regex=True,
                normalize_quotes_for_matching=False,
            )

            if not matches:
                break  # No more matches

            # Process only the first match
            match = matches[0]
            matched_text = match.text
            amount_str = matched_text[len(currency_symbol) :]

            try:
                replacement_text = f"{currency_symbol}{format_amount(amount_str)}"
            except ValueError:
                break  # Can't parse, stop

            # Skip if already correct
            if matched_text == replacement_text:
                break

            # Use existing replace logic which handles single match
            try:
                # Create exact pattern for this specific match to avoid ambiguity
                exact_pattern = re.escape(matched_text)
                self._document.replace_tracked(
                    find=exact_pattern,
                    replace=replacement_text,
                    author=author,
                    scope=scope,
                    regex=True,
                    enable_quote_normalization=False,
                )
                replacement_count += 1
            except (TextNotFoundError, AmbiguousTextError):
                break  # Can't replace, stop

        return replacement_count

    def normalize_dates(
        self,
        to_format: str = "%B %d, %Y",
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Normalize dates to a consistent format with tracked changes.

        Automatically detects common date formats and converts them to the target format.
        This prevents manual regex work and ensures date consistency.

        Detected formats:
            - MM/DD/YYYY (e.g., 12/08/2025)
            - M/D/YYYY (e.g., 1/8/2025)
            - YYYY-MM-DD (e.g., 2025-12-08)
            - Month DD, YYYY (e.g., December 08, 2025 or Dec 08, 2025)
            - DD Month YYYY (e.g., 08 December 2025)

        Args:
            to_format: Python datetime format string for output (default: "%B %d, %Y")
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})

        Returns:
            Number of dates normalized

        Example:
            >>> # Convert all dates to "December 08, 2025" format
            >>> count = doc.normalize_dates()
            >>>
            >>> # Convert all dates to ISO format
            >>> count = doc.normalize_dates("%Y-%m-%d")
        """
        # Resolve author
        author_name = author if author is not None else self._document.author

        # Common date patterns with their corresponding datetime format strings
        months_long = (
            "January|February|March|April|May|June|July|August|September|October|November|December"
        )
        months_short = "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"

        date_patterns = [
            # MM/DD/YYYY or M/D/YYYY
            (r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", "%m/%d/%Y"),
            # YYYY-MM-DD
            (r"\b(\d{4})-(\d{2})-(\d{2})\b", "%Y-%m-%d"),
            # Month DD, YYYY (e.g., December 08, 2025)
            (
                rf"\b({months_long}) (\d{{1,2}}), (\d{{4}})\b",
                "%B %d, %Y",
            ),
            # Mon DD, YYYY (e.g., Dec 08, 2025)
            (
                rf"\b({months_short}) (\d{{1,2}}), (\d{{4}})\b",
                "%b %d, %Y",
            ),
            # DD Month YYYY (e.g., 08 December 2025)
            (
                rf"\b(\d{{1,2}}) ({months_long}) (\d{{4}})\b",
                "%d %B %Y",
            ),
        ]

        # Get all paragraphs
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        all_matches = []
        for pattern, date_format in date_patterns:
            matches = self._document._text_search.find_text(
                pattern,
                paragraphs,
                regex=True,
                normalize_quotes_for_matching=False,
            )
            # Store matches with their format
            for match in matches:
                all_matches.append((match, date_format))

        if not all_matches:
            return 0

        # Sort by position (reverse) to process from end to beginning
        # This prevents position invalidation issues
        all_matches.sort(
            key=lambda x: (
                list(all_paragraphs).index(x[0].paragraph),
                x[0].start_run_index,
                x[0].start_offset,
            ),
            reverse=True,
        )

        # Process each match
        replacement_count = 0
        for match, date_format in all_matches:
            matched_text = match.text

            # Parse the date using the detected format
            try:
                parsed_date = datetime.strptime(matched_text, date_format)
            except ValueError:
                continue  # Skip if parsing fails

            # Format to target format
            replacement_text = parsed_date.strftime(to_format)

            # Skip if already in correct format
            if matched_text == replacement_text:
                continue

            # Generate tracked change XML
            deletion_xml = self._document._xml_generator.create_deletion(matched_text, author_name)
            insertion_xml = self._document._xml_generator.create_insertion(
                replacement_text, author_name
            )

            # Parse XMLs with namespace context
            wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
    {insertion_xml}
</root>"""
            root = etree.fromstring(wrapped_xml.encode("utf-8"))
            deletion_element = root[0]
            insertion_element = root[1]

            # Replace the matched text with deletion + insertion
            self._document._replace_match_with_elements(
                match, [deletion_element, insertion_element]
            )
            replacement_count += 1

        return replacement_count

    def update_section_references(
        self,
        old_number: str,
        new_number: str,
        section_word: str = "Section",
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Update section/article references with tracked changes.

        Finds references like "Section 2.1" and updates them to "Section 3.1".
        Prevents manual regex errors when renumbering document sections.

        Args:
            old_number: Old section number (e.g., "2.1")
            new_number: New section number (e.g., "3.1")
            section_word: Word used for sections (default: "Section",
                could be "Article", "Clause", etc.)
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text",
                dict={"contains": "text"})

        Returns:
            Number of references updated

        Example:
            >>> # Update all "Section 2.1" references to "Section 3.1"
            >>> count = doc.update_section_references("2.1", "3.1")
            >>>
            >>> # Update article references
            >>> count = doc.update_section_references("5", "6", section_word="Article")
        """
        # Escape special regex characters in the numbers
        old_escaped = re.escape(old_number)
        new_number_text = new_number

        # Build pattern: "Section 2.1" with optional trailing punctuation
        pattern = rf"\b{re.escape(section_word)}\s+{old_escaped}\b"

        # Use replace_tracked with regex
        try:
            self._document.replace_tracked(
                find=pattern,
                replace=f"{section_word} {new_number_text}",
                author=author,
                scope=scope,
                regex=True,
                enable_quote_normalization=False,
            )
            return 1
        except TextNotFoundError:
            return 0
        except AmbiguousTextError:
            # Multiple occurrences - need to replace all of them
            # Fall back to manual batch replacement
            author_name = author if author is not None else self._document.author

            all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
            paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

            matches = self._document._text_search.find_text(
                pattern,
                paragraphs,
                regex=True,
                normalize_quotes_for_matching=False,
            )

            if not matches:
                return 0

            # Process in reverse order
            replacement_count = 0
            for match in reversed(matches):
                matched_text = match.text
                replacement_text = f"{section_word} {new_number_text}"

                # Generate tracked change XML
                deletion_xml = self._document._xml_generator.create_deletion(
                    matched_text, author_name
                )
                insertion_xml = self._document._xml_generator.create_insertion(
                    replacement_text, author_name
                )

                # Parse XMLs with namespace context
                wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
    {insertion_xml}
</root>"""
                root = etree.fromstring(wrapped_xml.encode("utf-8"))
                deletion_element = root[0]
                insertion_element = root[1]

                # Replace the matched text with deletion + insertion
                self._document._replace_match_with_elements(
                    match, [deletion_element, insertion_element]
                )
                replacement_count += 1

            return replacement_count
