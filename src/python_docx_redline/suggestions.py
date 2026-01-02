"""
Suggestion generation for helpful error messages.

This module provides smart suggestions when text cannot be found in documents,
helping users quickly identify and fix common issues.
"""

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    pass


class SuggestionGenerator:
    """Generates helpful suggestions when text cannot be found.

    Analyzes the search text and document to provide actionable suggestions
    for resolving TextNotFoundError issues.
    """

    @staticmethod
    def generate_suggestions(text: str, paragraphs: list[Any]) -> list[str]:
        """Generate helpful suggestions when text is not found.

        Args:
            text: The text that was searched for
            paragraphs: List of paragraph Elements that were searched

        Returns:
            List of suggestion strings to help user find the text
        """
        suggestions = []

        # Check for curly quotes (common issue when copying from Word/PDF)
        if '"' in text or "'" in text:
            # Curly double quotes: " and "
            curly_double_quotes = "\u201c\u201d"
            has_curly_double = any(
                c in "".join(p.itertext()) for p in paragraphs for c in curly_double_quotes
            )
            # Curly single quotes/apostrophes: ' ' '
            curly_single_quotes = "\u2018\u2019\u2032"
            has_curly_single = any(
                c in "".join(p.itertext()) for p in paragraphs for c in curly_single_quotes
            )

            if has_curly_double and '"' in text:
                suggestions.append(
                    "Document contains curly quotes (\u201c\u201d). "
                    "Try replacing straight quotes with curly quotes in search text"
                )
            if has_curly_single and "'" in text:
                suggestions.append(
                    "Document contains curly apostrophes (\u2018\u2019). "
                    "Try replacing straight apostrophes with curly ones in search text"
                )

        # Check for double spaces
        if "  " in text:
            suggestions.append(
                "Search text contains double spaces. "
                "Document may have single spaces - try removing extra spaces"
            )

        # Check for leading/trailing whitespace
        if text != text.strip():
            stripped = text.strip()
            suggestions.append(f'Search text has leading/trailing whitespace. Try: "{stripped}"')

        # Check if the text might be in the document but with different formatting
        text_lower = text.lower()
        doc_text_lower = "".join("".join(p.itertext()) for p in paragraphs).lower()

        if text_lower in doc_text_lower:
            suggestions.append(
                "Text found with case-insensitive search. Check capitalization in your search text"
            )

        # Suggest checking for typos if no other suggestions
        if not suggestions:
            suggestions.extend(
                [
                    "Check for typos in the search text",
                    "Try searching for a shorter or more unique phrase",
                    "Verify the text exists in the document",
                ]
            )

        return suggestions

    @staticmethod
    def check_common_issues(text: str) -> list[str]:
        """Check for common issues in search text.

        Args:
            text: The text to check

        Returns:
            List of potential issues found
        """
        issues = []

        # Check for curly quotes
        if "\u201c" in text or "\u201d" in text:
            issues.append("Text contains curly double quotes")
        if "\u2018" in text or "\u2019" in text:
            issues.append("Text contains curly apostrophes")

        # Check for special characters that might cause issues
        if "\u00a0" in text:  # Non-breaking space
            issues.append("Text contains non-breaking spaces (\\u00a0)")

        if "\u200b" in text:  # Zero-width space
            issues.append("Text contains zero-width spaces (\\u200b)")

        # Check for unusual whitespace
        if "\t" in text:
            issues.append("Text contains tab characters")

        if "\r" in text or "\n" in text:
            issues.append("Text contains line breaks")

        return issues

    @staticmethod
    def find_similar_text(
        search_text: str,
        paragraphs: list[Any],
        max_suggestions: int = 3,
        min_similarity: float = 0.6,
    ) -> list[str]:
        """Find similar text in the document for suggestion purposes.

        Uses fuzzy matching to find text in the document that is similar to
        the search text. This helps users identify typos or near-matches.

        Args:
            search_text: The text that was searched for
            paragraphs: List of paragraph Elements to search in
            max_suggestions: Maximum number of suggestions to return
            min_similarity: Minimum similarity threshold (0.0 to 1.0)

        Returns:
            List of similar text strings found in the document

        Example:
            >>> # Document contains "production products"
            >>> similar = SuggestionGenerator.find_similar_text(
            ...     "producton products", paragraphs
            ... )
            >>> print(similar)
            ['production products']
        """
        try:
            from rapidfuzz import fuzz
        except ImportError:
            # If rapidfuzz is not available, return empty list
            return []

        if not search_text or not paragraphs:
            return []

        # Extract all text from paragraphs
        doc_text = " ".join("".join(p.itertext()) for p in paragraphs)

        # Skip if document is empty
        if not doc_text.strip():
            return []

        # Use sliding window approach to find similar substrings
        search_len = len(search_text)
        min_len = max(1, int(search_len * 0.7))
        max_len = int(search_len * 1.5)

        candidates: list[tuple[str, float]] = []
        seen: set[str] = set()

        # Check windows of varying sizes
        for window_len in range(min_len, min(max_len + 1, len(doc_text) + 1)):
            for start in range(len(doc_text) - window_len + 1):
                window = doc_text[start : start + window_len]

                # Skip if already seen (normalized)
                normalized = window.strip().lower()
                if normalized in seen or not normalized:
                    continue
                seen.add(normalized)

                # Calculate similarity
                similarity = fuzz.ratio(search_text.lower(), normalized) / 100.0

                if similarity >= min_similarity:
                    # Store original text (trimmed)
                    original = window.strip()
                    candidates.append((original, similarity))

        # Sort by similarity (descending) and take top suggestions
        candidates.sort(key=lambda x: x[1], reverse=True)

        # Remove duplicates while preserving order
        unique_suggestions: list[str] = []
        seen_lower: set[str] = set()
        for text, _score in candidates:
            text_lower = text.lower()
            if text_lower not in seen_lower:
                unique_suggestions.append(text)
                seen_lower.add(text_lower)
                if len(unique_suggestions) >= max_suggestions:
                    break

        return unique_suggestions
