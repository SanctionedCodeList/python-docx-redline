"""
Quote normalization for flexible text matching.

This module provides functionality to normalize quote characters, allowing
users to search with straight quotes (keyboard defaults) while matching
smart/curly quotes that are common in Word documents.

The approach is to normalize BOTH the search text and document text to straight
quotes for comparison, rather than trying to replicate Word's context-aware
smart quote logic.
"""


def normalize_quotes(text: str) -> str:
    """Normalize all quotes to straight quotes for matching.

    Word documents often contain smart/curly quotes that differ from the
    straight quotes produced by keyboard input. This function converts all
    quote variants to straight quotes so users can search with straight quotes
    and match any quote style.

    This is simpler and more reliable than trying to convert straight quotes
    to smart quotes, which requires context-awareness.

    Conversions (all to straight quotes):
        - ' (U+2018) → ' (U+0027) - Left single quotation mark
        - ' (U+2019) → ' (U+0027) - Right single quotation mark (apostrophes)
        - " (U+201C) → " (U+0022) - Left double quotation mark
        - " (U+201D) → " (U+0022) - Right double quotation mark

    Args:
        text: Text containing any quote style

    Returns:
        Text with all quotes converted to straight quotes

    Examples:
        >>> normalize_quotes("plaintiff's claim")
        "plaintiff's claim"  # ' becomes '

        >>> normalize_quotes('"free trial" offer')
        '"free trial" offer'  # " and " become "
    """
    # Convert all smart quote variants to straight quotes
    text = text.replace("\u2018", "'")  # Left single quotation mark → '
    text = text.replace("\u2019", "'")  # Right single quotation mark → '
    text = text.replace("\u201c", '"')  # Left double quotation mark → "
    text = text.replace("\u201d", '"')  # Right double quotation mark → "

    return text


def denormalize_quotes(text: str) -> str:
    """Convert smart/curly quotes back to straight quotes.

    This is the inverse of normalize_quotes(), converting smart quotes
    back to their straight equivalents.

    Conversions:
        - ' (U+2018) → ' (U+0027) - Left single quotation mark
        - ' (U+2019) → ' (U+0027) - Right single quotation mark
        - " (U+201C) → " (U+0022) - Left double quotation mark
        - " (U+201D) → " (U+0022) - Right double quotation mark

    Args:
        text: Text containing smart quotes to denormalize

    Returns:
        Text with smart quotes converted to straight quotes

    Examples:
        >>> denormalize_quotes("plaintiff's claim")
        "plaintiff's claim"  # ' becomes '

        >>> denormalize_quotes('"free trial" offer')
        '"free trial" offer'  # " and " become "
    """
    # Convert both left and right single quotation marks to straight apostrophe
    text = text.replace("\u2018", "'")  # Left single quotation mark
    text = text.replace("\u2019", "'")  # Right single quotation mark

    # Convert both left and right double quotation marks to straight quotes
    text = text.replace("\u201c", '"')  # Left double quotation mark
    text = text.replace("\u201d", '"')  # Right double quotation mark

    return text


def has_smart_quotes(text: str) -> bool:
    """Check if text contains smart/curly quotes.

    Args:
        text: Text to check

    Returns:
        True if text contains any smart quote characters
    """
    smart_quote_chars = "\u2018\u2019\u201c\u201d"
    return any(c in text for c in smart_quote_chars)


def has_straight_quotes(text: str) -> bool:
    """Check if text contains straight quotes.

    Args:
        text: Text to check

    Returns:
        True if text contains straight quote or apostrophe characters
    """
    return "'" in text or '"' in text
