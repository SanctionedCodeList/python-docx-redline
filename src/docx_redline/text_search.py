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

from dataclasses import dataclass
from typing import Any

# Word namespace
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


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
        text: The actual text content (computed on demand)
        context: Surrounding context for disambiguation (computed on demand)
    """

    runs: list[Any]  # lxml Elements
    start_run_index: int
    end_run_index: int
    start_offset: int
    end_offset: int
    paragraph: Any  # lxml Element

    @property
    def text(self) -> str:
        """Get the matched text."""
        result = []

        for idx in range(self.start_run_index, self.end_run_index + 1):
            run = self.runs[idx]
            run_text = "".join(run.itertext())

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
        para_text = "".join(self.paragraph.itertext())
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
        self, text: str, paragraphs: list[Any], case_sensitive: bool = True
    ) -> list[TextSpan]:
        """Find all occurrences of text in the given paragraphs.

        This is the core algorithm that handles text fragmentation:
        1. Build a character map that tracks which run each character belongs to
        2. Concatenate all text from all runs
        3. Search in the concatenated text
        4. Map the results back to the original runs

        Args:
            text: The text to search for
            paragraphs: List of paragraph Elements to search in
            case_sensitive: Whether to perform case-sensitive search (default: True)

        Returns:
            List of TextSpan objects representing each match
        """
        results = []

        # Prepare search text based on case sensitivity
        search_text = text if case_sensitive else text.lower()

        for para in paragraphs:
            # Get all runs in this paragraph
            runs = list(para.iter(_parse_tag("w:r")))

            if not runs:
                continue

            # Build character map: char_index -> (run_index, offset_in_run)
            char_map = []
            full_text_chars = []

            for run_idx, run in enumerate(runs):
                run_text = "".join(run.itertext())
                for char_idx, char in enumerate(run_text):
                    char_map.append((run_idx, char_idx))
                    full_text_chars.append(char)

            # Join into full paragraph text
            full_text = "".join(full_text_chars)

            # Apply case insensitivity if needed
            search_in = full_text if case_sensitive else full_text.lower()

            # Find all occurrences of the search text
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
