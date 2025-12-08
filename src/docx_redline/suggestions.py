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
