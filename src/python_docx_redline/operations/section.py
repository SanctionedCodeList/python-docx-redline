"""
SectionOperations class for handling section and paragraph operations.

This module provides a dedicated class for section-related operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

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
