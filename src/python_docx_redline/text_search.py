"""
Text search functionality for finding text in Word documents.

This module handles the core algorithm for finding text that may be fragmented
across multiple <w:r> (run) elements in the OOXML structure.

Algorithm Note:
    This implementation uses a character map approach for efficient read-only
    text searching. For an alternative approach using single-character run
    normalization (better for complex replacements), see Eric White's algorithm
    documented in docs/ERIC_WHITE_ALGORITHM.md.
"""

import re
from dataclasses import dataclass
from typing import Any

from .constants import WORD_NAMESPACE


def _parse_tag(tag: str) -> str:
    """Parse a tag name into a fully qualified namespace tag.

    Args:
        tag: Tag name like "w:r" or "w:t"

    Returns:
        Fully qualified tag like "{namespace}r"
    """
    if ":" in tag:
        prefix, local = tag.split(":", 1)
        if prefix == "w":
            return f"{{{WORD_NAMESPACE}}}{local}"
    return tag


def _get_run_text(run: Any) -> str:
    """Extract text from a run, avoiding XML structural whitespace.

    Extracts text from both w:t and w:delText elements to support searching
    in deleted text (tracked changes). This allows adding comments on text
    that has been marked for deletion.

    Args:
        run: A w:r (run) Element

    Returns:
        Text content of the run (includes both normal and deleted text)
    """
    # Find all w:t elements within this run
    text_elements = run.findall(f".//{{{WORD_NAMESPACE}}}t")
    # Also find w:delText elements (deleted text in tracked changes)
    deltext_elements = run.findall(f".//{{{WORD_NAMESPACE}}}delText")

    # Combine both types of text elements
    all_text_elements = text_elements + deltext_elements
    return "".join(elem.text or "" for elem in all_text_elements)


@dataclass
class TextSpan:
    """Represents found text across potentially multiple runs.

    Attributes:
        runs: List of lxml Element objects representing the runs
        start_run_index: Index of the run where the text starts
        end_run_index: Index of the run where the text ends
        start_offset: Character offset within the start run
        end_offset: Character offset within the end run (exclusive)
        paragraph: The parent paragraph Element
        match_obj: Optional regex Match object for capture group support
        text: The actual text content (computed on demand)
        context: Surrounding context for disambiguation (computed on demand)
    """

    runs: list[Any]  # lxml Elements
    start_run_index: int
    end_run_index: int
    start_offset: int
    end_offset: int
    paragraph: Any  # lxml Element
    match_obj: Any = None  # Optional re.Match object for regex matches

    @property
    def text(self) -> str:
        """Get the matched text."""
        result = []

        for idx in range(self.start_run_index, self.end_run_index + 1):
            run = self.runs[idx]
            run_text = _get_run_text(run)

            if idx == self.start_run_index and idx == self.end_run_index:
                # Text is within a single run
                result.append(run_text[self.start_offset : self.end_offset])
            elif idx == self.start_run_index:
                # First run
                result.append(run_text[self.start_offset :])
            elif idx == self.end_run_index:
                # Last run
                result.append(run_text[: self.end_offset])
            else:
                # Middle run
                result.append(run_text)

        return "".join(result)

    @property
    def context(self) -> str:
        """Get surrounding context for disambiguation.

        Returns up to 40 characters before and after the matched text.
        """
        # Extract text only from w:t elements
        text_elements = self.paragraph.findall(f".//{{{WORD_NAMESPACE}}}t")
        para_text = "".join(elem.text or "" for elem in text_elements)
        matched = self.text

        # Find the match position in the full paragraph text
        match_pos = para_text.find(matched)
        if match_pos == -1:
            return matched

        # Get context window
        context_before = 40
        context_after = 40

        start = max(0, match_pos - context_before)
        end = min(len(para_text), match_pos + len(matched) + context_after)

        context = para_text[start:end]

        # Add ellipsis if needed
        if start > 0:
            context = "..." + context
        if end < len(para_text):
            context = context + "..."

        return context


class TextSearch:
    """Handles searching for text in Word documents with fragmentation support.

    The core challenge is that text in Word documents can be split across
    multiple <w:r> (run) elements, making simple text search unreliable.
    This class builds a character map to handle fragmentation correctly.
    """

    def find_text(
        self,
        text: str,
        paragraphs: list[Any],
        case_sensitive: bool = True,
        regex: bool = False,
        normalize_quotes_for_matching: bool = False,
        fuzzy: dict[str, Any] | None = None,
    ) -> list[TextSpan]:
        """Find all occurrences of text in the given paragraphs.

        This is the core algorithm that handles text fragmentation:
        1. Build a character map that tracks which run each character belongs to
        2. Concatenate all text from all runs
        3. Search in the concatenated text (literal, regex, or fuzzy)
        4. Map the results back to the original runs

        Args:
            text: The text or regex pattern to search for
            paragraphs: List of paragraph Elements to search in
            case_sensitive: Whether to perform case-sensitive search (default: True)
            regex: Whether to treat text as a regex pattern (default: False)
            normalize_quotes_for_matching: Normalize quotes to straight quotes for matching
                (default: False)
            fuzzy: Fuzzy matching configuration dict with keys:
                - threshold: Similarity threshold (0.0 to 1.0)
                - algorithm: Matching algorithm (ratio, partial_ratio, etc.)
                - normalize_whitespace: Whether to normalize whitespace
                (default: None for exact matching)

        Returns:
            List of TextSpan objects representing each match

        Raises:
            re.error: If regex=True and the pattern is invalid
            ImportError: If fuzzy matching requested but rapidfuzz not installed
        """
        from .quote_normalization import normalize_quotes as normalize_quotes_func

        results = []

        # Fuzzy and regex are mutually exclusive
        if fuzzy and regex:
            raise ValueError("Cannot use both fuzzy matching and regex")

        # Prepare search pattern
        if fuzzy:
            # Fuzzy matching mode - import fuzzy functions
            from .fuzzy import fuzzy_find_all

            search_text = text  # Keep original for fuzzy matching
            pattern = None
        elif regex:
            # Compile regex pattern with case sensitivity flag
            flags = 0 if case_sensitive else re.IGNORECASE
            try:
                pattern = re.compile(text, flags)
            except re.error as e:
                raise re.error(f"Invalid regex pattern '{text}': {e}") from e
            search_text = None  # Not used for regex
        else:
            # Prepare literal search text based on case sensitivity
            search_text = text if case_sensitive else text.lower()
            # Apply quote normalization if requested
            if normalize_quotes_for_matching:
                search_text = normalize_quotes_func(search_text)
            pattern = None  # Not used for literal search

        for para in paragraphs:
            # Get all runs in this paragraph
            runs = list(para.iter(_parse_tag("w:r")))

            if not runs:
                continue

            # Build character map: char_index -> (run_index, offset_in_run)
            char_map = []
            full_text_chars = []

            for run_idx, run in enumerate(runs):
                run_text = _get_run_text(run)
                for char_idx, char in enumerate(run_text):
                    char_map.append((run_idx, char_idx))
                    full_text_chars.append(char)

            # Join into full paragraph text
            full_text = "".join(full_text_chars)

            # Normalize document text for matching if requested
            search_full_text = full_text
            if normalize_quotes_for_matching and not regex and not fuzzy:
                search_full_text = normalize_quotes_func(search_full_text)
            if not case_sensitive and not regex and not fuzzy:
                search_full_text = search_full_text.lower()

            # Find all occurrences
            if fuzzy:
                # Use fuzzy search
                assert search_text is not None  # Type guard: search_text is set when fuzzy=True
                fuzzy_matches = fuzzy_find_all(
                    full_text,
                    search_text,
                    threshold=fuzzy["threshold"],
                    algorithm=fuzzy["algorithm"],
                    normalize_ws=fuzzy["normalize_whitespace"],
                )

                for start_pos, end_pos, similarity in fuzzy_matches:
                    # Map the found position back to runs
                    start_run_idx, start_offset = char_map[start_pos]
                    end_run_idx, end_offset = char_map[end_pos - 1]

                    # Create TextSpan for this match
                    span = TextSpan(
                        runs=runs,
                        start_run_index=start_run_idx,
                        end_run_index=end_run_idx,
                        start_offset=start_offset,
                        end_offset=end_offset + 1,  # Make end_offset exclusive
                        paragraph=para,
                    )
                    results.append(span)
            elif regex:
                # Use regex search
                assert pattern is not None  # Type guard: pattern is set when regex=True
                for match in pattern.finditer(full_text):
                    pos = match.start()
                    match_len = match.end() - match.start()

                    # Map the found position back to runs
                    start_run_idx, start_offset = char_map[pos]
                    end_run_idx, end_offset = char_map[pos + match_len - 1]

                    # Create TextSpan for this match with regex Match object
                    span = TextSpan(
                        runs=runs,
                        start_run_index=start_run_idx,
                        end_run_index=end_run_idx,
                        start_offset=start_offset,
                        end_offset=end_offset + 1,  # Make end_offset exclusive
                        paragraph=para,
                        match_obj=match,  # Store match for capture group support
                    )
                    results.append(span)
            else:
                # Use literal search
                assert search_text is not None  # Type guard: search_text is set when regex=False
                # Use the already-normalized and case-adjusted text
                search_in = search_full_text

                start = 0
                while True:
                    pos = search_in.find(search_text, start)
                    if pos == -1:
                        break

                    # Map the found position back to runs
                    start_run_idx, start_offset = char_map[pos]
                    end_run_idx, end_offset = char_map[pos + len(search_text) - 1]

                    # Create TextSpan for this match
                    span = TextSpan(
                        runs=runs,
                        start_run_index=start_run_idx,
                        end_run_index=end_run_idx,
                        start_offset=start_offset,
                        end_offset=end_offset + 1,  # Make end_offset exclusive
                        paragraph=para,
                    )
                    results.append(span)

                    # Move past this match for the next search
                    start = pos + 1

        return results
