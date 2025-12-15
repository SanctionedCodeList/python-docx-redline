"""Formatting operations for text and paragraph formatting with tracked changes.

This module provides the FormatOperations class for applying formatting
changes to Word documents with proper change tracking.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from lxml import etree

if TYPE_CHECKING:
    from ..document import Document

from ..constants import WORD_NAMESPACE
from ..errors import AmbiguousTextError, TextNotFoundError
from ..format_builder import (
    ParagraphPropertyBuilder,
    RunPropertyBuilder,
    get_run_text,
    split_run_at_offset,
)
from ..results import FormatResult
from ..scope import ScopeEvaluator
from ..suggestions import SuggestionGenerator


class FormatOperations:
    """Handles formatting operations with tracked changes.

    This class extracts formatting operations from the Document class
    to improve separation of concerns and maintainability.

    Attributes:
        _document: Reference to the parent Document instance
    """

    def __init__(self, document: Document) -> None:
        """Initialize the FormatOperations.

        Args:
            document: The parent Document instance
        """
        self._document = document

    @property
    def xml_root(self) -> Any:
        """Get the XML root element from the document."""
        return self._document.xml_root

    def apply_style(
        self,
        find: str,
        style: str,
        scope: str | dict | Any | None = None,
        regex: bool = False,
    ) -> int:
        """Apply a paragraph style to paragraphs containing specific text.

        Changes the style of paragraphs that contain the search text.
        This is useful for programmatically formatting document sections.

        Args:
            find: Text to search for (or regex pattern if regex=True)
            style: Paragraph style name (e.g., 'Heading1', 'Normal', 'Quote')
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            regex: Whether to treat 'find' as a regex pattern (default: False)

        Returns:
            Number of paragraphs whose style was changed

        Example:
            >>> # Make all paragraphs containing "Section" into headings
            >>> count = doc.apply_style("Section", "Heading1")
            >>>
            >>> # Apply quote style to paragraphs with specific text
            >>> count = doc.apply_style("As stated in", "Quote")
        """
        from ..models.paragraph import Paragraph as ParagraphClass

        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find paragraphs containing the text
        matches = self._document._text_search.find_text(
            find,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=not regex,
        )

        if not matches:
            return 0

        # Get unique paragraphs (a paragraph might have multiple matches)
        unique_paragraphs = {match.paragraph for match in matches}

        # Apply style to each paragraph
        count = 0
        for para_element in unique_paragraphs:
            para = ParagraphClass(para_element)
            if para.style != style:
                para.style = style
                count += 1

        return count

    def format_text(
        self,
        find: str,
        bold: bool | None = None,
        italic: bool | None = None,
        color: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
    ) -> int:
        """Apply text formatting (bold, italic, color) to specific text.

        Finds text and applies formatting to the runs containing it.
        This allows surgical formatting changes without affecting surrounding text.

        Args:
            find: Text to search for (or regex pattern if regex=True)
            bold: Set bold formatting (True/False/None to leave unchanged)
            italic: Set italic formatting (True/False/None to leave unchanged)
            color: Set text color as hex (e.g., "FF0000" for red, None to leave unchanged)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            regex: Whether to treat 'find' as a regex pattern (default: False)

        Returns:
            Number of text occurrences formatted

        Example:
            >>> # Make all occurrences of "IMPORTANT" bold and red
            >>> count = doc.format_text("IMPORTANT", bold=True, color="FF0000")
            >>>
            >>> # Make section references italic
            >>> count = doc.format_text(r"Section \\d+\\.\\d+", italic=True, regex=True)
        """
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(
            find, paragraphs, regex=regex, normalize_quotes_for_matching=not regex
        )
        if not matches:
            return 0

        count = 0
        for match in matches:
            for run_idx in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[run_idx]
                r_pr = self._get_or_create_run_properties(run)
                self._apply_simple_formatting(r_pr, bold, italic, color)
            count += 1
        return count

    def _get_or_create_run_properties(self, run: Any) -> Any:
        """Get or create run properties element for a run."""
        r_pr = run.find(f"{{{WORD_NAMESPACE}}}rPr")
        if r_pr is None:
            r_pr = etree.Element(f"{{{WORD_NAMESPACE}}}rPr")
            run.insert(0, r_pr)
        return r_pr

    def _apply_simple_formatting(
        self,
        r_pr: Any,
        bold: bool | None,
        italic: bool | None,
        color: str | None,
    ) -> None:
        """Apply simple formatting (bold, italic, color) to run properties."""
        self._set_toggle_property(r_pr, "b", bold)
        self._set_toggle_property(r_pr, "i", italic)
        if color is not None:
            self._set_color_property(r_pr, color)

    def _set_toggle_property(self, r_pr: Any, prop_name: str, value: bool | None) -> None:
        """Set a toggle property (bold, italic) on run properties."""
        if value is None:
            return
        elem = r_pr.find(f"{{{WORD_NAMESPACE}}}{prop_name}")
        if value:
            if elem is None:
                etree.SubElement(r_pr, f"{{{WORD_NAMESPACE}}}{prop_name}")
        else:
            if elem is not None:
                r_pr.remove(elem)

    def _set_color_property(self, r_pr: Any, color: str) -> None:
        """Set color property on run properties."""
        color_elem = r_pr.find(f"{{{WORD_NAMESPACE}}}color")
        if color_elem is None:
            color_elem = etree.SubElement(r_pr, f"{{{WORD_NAMESPACE}}}color")
        color_elem.set(f"{{{WORD_NAMESPACE}}}val", color)

    def format_tracked(
        self,
        text: str,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | str | None = None,
        strikethrough: bool | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        color: str | None = None,
        highlight: str | None = None,
        superscript: bool | None = None,
        subscript: bool | None = None,
        small_caps: bool | None = None,
        all_caps: bool | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | str = "first",
        author: str | None = None,
        enable_quote_normalization: bool = True,
    ) -> FormatResult:
        """Apply character formatting to text with tracked changes.

        This method finds text and applies formatting changes that are tracked
        as revisions in Word. The previous formatting state is preserved in
        <w:rPrChange> elements, allowing users to accept or reject the
        formatting changes in Word.

        Args:
            text: The text to format (found via text search)
            bold: Set bold on (True), off (False), or leave unchanged (None)
            italic: Set italic on/off/unchanged
            underline: Set underline on/off/unchanged, or underline style name
            strikethrough: Set strikethrough on/off/unchanged
            font_name: Set font family name
            font_size: Set font size in points
            color: Set text color as hex "#RRGGBB" or "auto"
            highlight: Set highlight color name (e.g., "yellow", "green")
            superscript: Set superscript on/off/unchanged
            subscript: Set subscript on/off/unchanged
            small_caps: Set small caps on/off/unchanged
            all_caps: Set all caps on/off/unchanged
            scope: Limit search to specific paragraphs/sections
            occurrence: Which occurrence to format: 1, 2, "first", "last", or "all"
            author: Override default author for this change
            enable_quote_normalization: Auto-convert straight quotes to smart quotes
                for matching (default: True)

        Returns:
            FormatResult with details of the formatting applied

        Raises:
            TextNotFoundError: If text is not found
            AmbiguousTextError: If multiple matches and occurrence not specified

        Example:
            >>> doc.format_tracked("IMPORTANT", bold=True, color="#FF0000")
            >>> doc.format_tracked("Section 2.1", italic=True, scope="section:Introduction")
            >>> doc.format_tracked("Note:", underline=True, font_size=14)
        """
        # Build format updates dict
        format_updates = self._build_format_updates(
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            font_name=font_name,
            font_size=font_size,
            color=color,
            highlight=highlight,
            superscript=superscript,
            subscript=subscript,
            small_caps=small_caps,
            all_caps=all_caps,
        )

        if not format_updates:
            raise ValueError("At least one formatting property must be specified")

        # Find matches
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(
            text,
            paragraphs,
            regex=False,
            normalize_quotes_for_matching=enable_quote_normalization,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(text, paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        # Select target matches based on occurrence
        target_matches = self._select_matches(matches, occurrence, text)

        # Apply formatting to each match
        return self._apply_run_formatting(target_matches, format_updates, all_paragraphs, author)

    def _build_format_updates(
        self,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | str | None = None,
        strikethrough: bool | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        color: str | None = None,
        highlight: str | None = None,
        superscript: bool | None = None,
        subscript: bool | None = None,
        small_caps: bool | None = None,
        all_caps: bool | None = None,
    ) -> dict[str, Any]:
        """Build format updates dict from non-None values."""
        format_updates: dict[str, Any] = {}
        if bold is not None:
            format_updates["bold"] = bold
        if italic is not None:
            format_updates["italic"] = italic
        if underline is not None:
            format_updates["underline"] = underline
        if strikethrough is not None:
            format_updates["strikethrough"] = strikethrough
        if font_name is not None:
            format_updates["font_name"] = font_name
        if font_size is not None:
            format_updates["font_size"] = font_size
        if color is not None:
            format_updates["color"] = color
        if highlight is not None:
            format_updates["highlight"] = highlight
        if superscript is not None:
            format_updates["superscript"] = superscript
        if subscript is not None:
            format_updates["subscript"] = subscript
        if small_caps is not None:
            format_updates["small_caps"] = small_caps
        if all_caps is not None:
            format_updates["all_caps"] = all_caps
        return format_updates

    def _select_matches(self, matches: list, occurrence: int | str, text: str) -> list:
        """Select target matches based on occurrence parameter."""
        if occurrence == "first" or occurrence == 1:
            return [matches[0]]
        elif occurrence == "last":
            return [matches[-1]]
        elif occurrence == "all":
            return matches
        elif isinstance(occurrence, int) and 1 <= occurrence <= len(matches):
            return [matches[occurrence - 1]]
        elif isinstance(occurrence, int):
            raise ValueError(f"Occurrence {occurrence} out of range (1-{len(matches)})")
        elif len(matches) > 1:
            raise AmbiguousTextError(text, matches)
        else:
            return matches

    def _apply_run_formatting(
        self,
        target_matches: list,
        format_updates: dict[str, Any],
        all_paragraphs: list,
        author: str | None,
    ) -> FormatResult:
        """Apply formatting to runs in target matches."""
        from copy import deepcopy

        runs_affected = 0
        last_change_id = 0
        all_previous_formatting: list[dict[str, object]] = []
        para_index = -1

        for match in target_matches:
            para = match.paragraph
            para_index = all_paragraphs.index(para) if para in all_paragraphs else -1

            # Build list of runs to format, handling mid-run splits
            runs_to_format = self._prepare_runs_for_formatting(match)

            # Apply formatting to the runs
            for run in runs_to_format:
                existing_rpr = run.find(f"{{{WORD_NAMESPACE}}}rPr")
                previous_rpr = deepcopy(existing_rpr) if existing_rpr is not None else None

                # Extract previous formatting
                prev_formatting = RunPropertyBuilder.extract(previous_rpr)
                all_previous_formatting.append(prev_formatting)

                # Create new rPr with merged formatting
                new_rpr = RunPropertyBuilder.merge(existing_rpr, format_updates)

                # Check if there are actual changes
                if not RunPropertyBuilder.has_changes(previous_rpr, new_rpr):
                    continue

                # Create the tracked change element
                rpr_change, last_change_id = (
                    self._document._xml_generator.create_run_property_change(previous_rpr, author)
                )

                new_rpr.append(rpr_change)

                # Replace or insert the rPr in the run
                if existing_rpr is not None:
                    run.remove(existing_rpr)
                run.insert(0, new_rpr)

                runs_affected += 1

        return FormatResult(
            success=True,
            changed=runs_affected > 0,
            text_matched=target_matches[0].text if target_matches else "",
            paragraph_index=para_index if len(target_matches) == 1 else -1,
            changes_applied=format_updates,
            previous_formatting=all_previous_formatting,
            change_id=last_change_id,
            runs_affected=runs_affected,
        )

    def _prepare_runs_for_formatting(self, match) -> list:
        """Prepare runs for formatting, splitting at match boundaries if needed."""
        runs_to_format = []

        for run_idx in range(match.start_run_index, match.end_run_index + 1):
            run = match.runs[run_idx]
            run_text = get_run_text(run)

            is_start = run_idx == match.start_run_index
            is_end = run_idx == match.end_run_index
            is_single = is_start and is_end

            if is_single and (match.start_offset > 0 or match.end_offset < len(run_text)):
                # Match is within a single run - split at both ends
                runs_to_format.extend(self._split_single_run(run, match, run_text))
            elif is_start and match.start_offset > 0:
                # Split start run
                runs_to_format.append(self._split_start_run(run, match.start_offset))
            elif is_end and match.end_offset < len(run_text):
                # Split end run
                runs_to_format.append(self._split_end_run(run, match.end_offset))
            else:
                # Whole run is within match
                runs_to_format.append(run)

        return runs_to_format

    def _split_single_run(self, run, match, run_text: str) -> list:
        """Split a single run when match is in the middle."""
        parent = run.getparent()
        idx = list(parent).index(run)

        if match.start_offset > 0:
            before_run, remainder = split_run_at_offset(run, match.start_offset)
            parent.insert(idx, before_run)
            adjusted_end = match.end_offset - match.start_offset
            if adjusted_end < len(run_text) - match.start_offset:
                middle_run, after_run = split_run_at_offset(remainder, adjusted_end)
                parent.remove(run)
                parent.insert(idx + 1, middle_run)
                parent.insert(idx + 2, after_run)
                return [middle_run]
            else:
                parent.remove(run)
                parent.insert(idx + 1, remainder)
                return [remainder]
        else:
            # Only split at end
            middle_run, after_run = split_run_at_offset(run, match.end_offset)
            parent.remove(run)
            parent.insert(idx, middle_run)
            parent.insert(idx + 1, after_run)
            return [middle_run]

    def _split_start_run(self, run, start_offset: int):
        """Split start run and return the part to format."""
        before_run, after_run = split_run_at_offset(run, start_offset)
        parent = run.getparent()
        idx = list(parent).index(run)
        parent.remove(run)
        parent.insert(idx, before_run)
        parent.insert(idx + 1, after_run)
        return after_run

    def _split_end_run(self, run, end_offset: int):
        """Split end run and return the part to format."""
        before_run, after_run = split_run_at_offset(run, end_offset)
        parent = run.getparent()
        idx = list(parent).index(run)
        parent.remove(run)
        parent.insert(idx, before_run)
        parent.insert(idx + 1, after_run)
        return before_run

    def format_paragraph_tracked(
        self,
        *,
        containing: str | None = None,
        starting_with: str | None = None,
        ending_with: str | None = None,
        index: int | None = None,
        alignment: str | None = None,
        spacing_before: float | None = None,
        spacing_after: float | None = None,
        line_spacing: float | None = None,
        indent_left: float | None = None,
        indent_right: float | None = None,
        indent_first_line: float | None = None,
        indent_hanging: float | None = None,
        scope: str | dict | Any | None = None,
        author: str | None = None,
    ) -> FormatResult:
        """Apply paragraph formatting with tracked changes.

        This method finds a paragraph and applies formatting changes that are
        tracked as revisions in Word. The previous formatting state is preserved
        in <w:pPrChange> elements.

        Args:
            containing: Find paragraph containing this text
            starting_with: Find paragraph starting with this text
            ending_with: Find paragraph ending with this text
            index: Target paragraph by index (0-based)
            alignment: Set paragraph alignment ("left", "center", "right", "justify")
            spacing_before: Set space before paragraph (points)
            spacing_after: Set space after paragraph (points)
            line_spacing: Set line spacing multiplier (e.g., 1.0, 1.5, 2.0)
            indent_left: Set left indent (inches)
            indent_right: Set right indent (inches)
            indent_first_line: Set first line indent (inches)
            indent_hanging: Set hanging indent (inches)
            scope: Limit search to specific sections
            author: Override default author for this change

        Returns:
            FormatResult with details of the formatting applied

        Raises:
            TextNotFoundError: If no matching paragraph found
            ValueError: If no targeting parameter or formatting parameter specified

        Example:
            >>> doc.format_paragraph_tracked(containing="WHEREAS", alignment="center")
            >>> doc.format_paragraph_tracked(index=0, spacing_after=12)
        """
        # Validate at least one targeting parameter
        if containing is None and starting_with is None and ending_with is None and index is None:
            raise ValueError(
                "At least one targeting parameter required: "
                "containing, starting_with, ending_with, or index"
            )

        # Build format updates dict
        format_updates = self._build_paragraph_format_updates(
            alignment=alignment,
            spacing_before=spacing_before,
            spacing_after=spacing_after,
            line_spacing=line_spacing,
            indent_left=indent_left,
            indent_right=indent_right,
            indent_first_line=indent_first_line,
            indent_hanging=indent_hanging,
        )

        if not format_updates:
            raise ValueError("At least one formatting property must be specified")

        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find target paragraph
        target_para, para_index = self._find_target_paragraph(
            paragraphs,
            all_paragraphs,
            containing=containing,
            starting_with=starting_with,
            ending_with=ending_with,
            index=index,
        )

        # Apply formatting
        return self._apply_paragraph_formatting(target_para, para_index, format_updates, author)

    def _build_paragraph_format_updates(
        self,
        alignment: str | None = None,
        spacing_before: float | None = None,
        spacing_after: float | None = None,
        line_spacing: float | None = None,
        indent_left: float | None = None,
        indent_right: float | None = None,
        indent_first_line: float | None = None,
        indent_hanging: float | None = None,
    ) -> dict[str, Any]:
        """Build paragraph format updates dict."""
        format_updates: dict[str, Any] = {}
        if alignment is not None:
            format_updates["alignment"] = alignment
        if spacing_before is not None:
            format_updates["spacing_before"] = spacing_before
        if spacing_after is not None:
            format_updates["spacing_after"] = spacing_after
        if line_spacing is not None:
            format_updates["line_spacing"] = line_spacing
        if indent_left is not None:
            format_updates["indent_left"] = indent_left
        if indent_right is not None:
            format_updates["indent_right"] = indent_right
        if indent_first_line is not None:
            format_updates["indent_first_line"] = indent_first_line
        if indent_hanging is not None:
            format_updates["indent_hanging"] = indent_hanging
        return format_updates

    def _find_target_paragraph(
        self,
        paragraphs: list,
        all_paragraphs: list,
        containing: str | None = None,
        starting_with: str | None = None,
        ending_with: str | None = None,
        index: int | None = None,
    ) -> tuple[Any, int]:
        """Find the target paragraph based on search criteria."""
        if index is not None:
            if 0 <= index < len(paragraphs):
                target_para = paragraphs[index]
                para_index = (
                    all_paragraphs.index(target_para) if target_para in all_paragraphs else index
                )
                return target_para, para_index
            else:
                raise ValueError(f"Paragraph index {index} out of range (0-{len(paragraphs) - 1})")

        # Search for paragraph by text content
        for i, para in enumerate(paragraphs):
            para_text = self._document._get_paragraph_text(para)

            if containing is not None and containing not in para_text:
                continue
            if starting_with is not None and not para_text.startswith(starting_with):
                continue
            if ending_with is not None and not para_text.endswith(ending_with):
                continue

            para_index = all_paragraphs.index(para) if para in all_paragraphs else i
            return para, para_index

        search_text = containing or starting_with or ending_with or ""
        raise TextNotFoundError(
            search_text,
            suggestions=["Check paragraph content", "Try a different search term"],
        )

    def _apply_paragraph_formatting(
        self,
        target_para: Any,
        para_index: int,
        format_updates: dict[str, Any],
        author: str | None,
    ) -> FormatResult:
        """Apply paragraph formatting with tracking."""
        from copy import deepcopy

        existing_ppr = target_para.find(f"{{{WORD_NAMESPACE}}}pPr")
        previous_ppr = deepcopy(existing_ppr) if existing_ppr is not None else None

        # Extract previous formatting
        prev_formatting = ParagraphPropertyBuilder.extract(previous_ppr)

        # Create new pPr with merged formatting
        new_ppr = ParagraphPropertyBuilder.merge(existing_ppr, format_updates)

        # Check if there are actual changes
        if not ParagraphPropertyBuilder.has_changes(previous_ppr, new_ppr):
            return FormatResult(
                success=True,
                changed=False,
                text_matched=self._document._get_paragraph_text(target_para)[:50],
                paragraph_index=para_index,
                changes_applied={},
                previous_formatting=[prev_formatting],
                change_id=0,
                runs_affected=0,
            )

        # Create the tracked change element
        ppr_change, change_id = self._document._xml_generator.create_paragraph_property_change(
            previous_ppr, author
        )

        new_ppr.append(ppr_change)

        # Replace or insert the pPr in the paragraph
        if existing_ppr is not None:
            target_para.remove(existing_ppr)
        target_para.insert(0, new_ppr)

        return FormatResult(
            success=True,
            changed=True,
            text_matched=self._document._get_paragraph_text(target_para)[:50],
            paragraph_index=para_index,
            changes_applied=format_updates,
            previous_formatting=[prev_formatting],
            change_id=change_id,
            runs_affected=1,
        )

    def copy_format(
        self,
        from_text: str,
        to_text: str,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Copy formatting from one text to another.

        Finds the source text, extracts its formatting (bold, italic, color, etc.),
        and applies the same formatting to the target text.

        Args:
            from_text: Source text to copy formatting from
            to_text: Target text to apply formatting to
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})

        Returns:
            Number of target occurrences formatted

        Example:
            >>> # Copy formatting from headers to make matching text look the same
            >>> count = doc.copy_format("Chapter 1", "Chapter 2")
        """
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        source_matches = self._document._text_search.find_text(
            from_text, paragraphs, regex=False, normalize_quotes_for_matching=True
        )
        if not source_matches:
            raise TextNotFoundError(from_text)

        source_r_pr = self._get_source_run_properties(source_matches[0])
        if source_r_pr is None:
            return 0

        bold, italic, color = self._extract_simple_formatting(source_r_pr)
        return self.format_text(to_text, bold=bold, italic=italic, color=color, scope=scope)

    def _get_source_run_properties(self, match: Any) -> Any | None:
        """Get run properties from the first run of a match."""
        source_run = match.runs[match.start_run_index]
        return source_run.find(f"{{{WORD_NAMESPACE}}}rPr")

    def _extract_simple_formatting(self, r_pr: Any) -> tuple[bool, bool, str | None]:
        """Extract simple formatting (bold, italic, color) from run properties."""
        bold = r_pr.find(f"{{{WORD_NAMESPACE}}}b") is not None
        italic = r_pr.find(f"{{{WORD_NAMESPACE}}}i") is not None
        color_elem = r_pr.find(f"{{{WORD_NAMESPACE}}}color")
        color = color_elem.get(f"{{{WORD_NAMESPACE}}}val") if color_elem is not None else None
        return bold, italic, color
