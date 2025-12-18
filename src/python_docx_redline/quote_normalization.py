"""
Special character normalization for flexible text matching.

This module provides functionality to normalize special characters (quotes, bullets,
dashes) allowing users to search with standard keyboard characters while matching
the various Unicode variants common in Word documents.

The approach is to normalize BOTH the search text and document text to standard
ASCII equivalents for comparison.
"""


def normalize_special_chars(text: str) -> str:
    """Normalize special characters (quotes, bullets, dashes) for matching.

    Word documents often contain typographic characters that differ from standard
    keyboard input. This function converts common variants to their ASCII
    equivalents so users can search with standard characters.

    Normalizations:
        Quotes:
        - ' (U+2018) → ' (U+0027) - Left single quotation mark
        - ' (U+2019) → ' (U+0027) - Right single quotation mark (apostrophes)
        - " (U+201C) → " (U+0022) - Left double quotation mark
        - " (U+201D) → " (U+0022) - Right double quotation mark

        Bullets (all → •):
        - · (U+00B7) → • - Middle dot
        - ◦ (U+25E6) → • - White bullet
        - ▪ (U+25AA) → • - Black small square
        - ▫ (U+25AB) → • - White small square
        - ‣ (U+2023) → • - Triangular bullet
        - ⁃ (U+2043) → • - Hyphen bullet
        - ● (U+25CF) → • - Black circle
        - ○ (U+25CB) → • - White circle

        Dashes:
        - – (U+2013) → - - En dash
        - — (U+2014) → - - Em dash
        - ‐ (U+2010) → - - Hyphen
        - ‑ (U+2011) → - - Non-breaking hyphen
        - − (U+2212) → - - Minus sign

    Args:
        text: Text containing special characters

    Returns:
        Text with special characters normalized to ASCII equivalents

    Examples:
        >>> normalize_special_chars("plaintiff's claim")
        "plaintiff's claim"  # ' becomes '

        >>> normalize_special_chars("• First item")
        "• First item"  # Various bullets become •

        >>> normalize_special_chars("2020–2024")
        "2020-2024"  # En dash becomes hyphen
    """
    # Quotes - convert smart quotes to straight quotes
    text = text.replace("\u2018", "'")  # Left single quotation mark → '
    text = text.replace("\u2019", "'")  # Right single quotation mark → '
    text = text.replace("\u201c", '"')  # Left double quotation mark → "
    text = text.replace("\u201d", '"')  # Right double quotation mark → "

    # Bullets - normalize to standard bullet •
    text = text.replace("\u00b7", "\u2022")  # Middle dot → bullet
    text = text.replace("\u25e6", "\u2022")  # White bullet → bullet
    text = text.replace("\u25aa", "\u2022")  # Black small square → bullet
    text = text.replace("\u25ab", "\u2022")  # White small square → bullet
    text = text.replace("\u2023", "\u2022")  # Triangular bullet → bullet
    text = text.replace("\u2043", "\u2022")  # Hyphen bullet → bullet
    text = text.replace("\u25cf", "\u2022")  # Black circle → bullet
    text = text.replace("\u25cb", "\u2022")  # White circle → bullet

    # Dashes - normalize to standard hyphen-minus
    text = text.replace("\u2013", "-")  # En dash → hyphen
    text = text.replace("\u2014", "-")  # Em dash → hyphen
    text = text.replace("\u2010", "-")  # Hyphen → hyphen-minus
    text = text.replace("\u2011", "-")  # Non-breaking hyphen → hyphen
    text = text.replace("\u2212", "-")  # Minus sign → hyphen

    return text


def normalize_quotes(text: str) -> str:
    """Normalize all quotes to straight quotes for matching.

    .. deprecated::
        Use :func:`normalize_special_chars` instead, which handles quotes
        plus bullets and dashes.

    Word documents often contain smart/curly quotes that differ from the
    straight quotes produced by keyboard input. This function converts all
    quote variants to straight quotes so users can search with straight quotes
    and match any quote style.

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
    # For backwards compatibility, just do quote normalization
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


def has_special_chars(text: str) -> bool:
    """Check if text contains any normalizable special characters.

    Args:
        text: Text to check

    Returns:
        True if text contains smart quotes, special bullets, or special dashes
    """
    special_chars = (
        "\u2018\u2019\u201c\u201d"  # Smart quotes
        "\u00b7\u25e6\u25aa\u25ab\u2023\u2043\u25cf\u25cb"  # Bullets
        "\u2013\u2014\u2010\u2011\u2212"  # Dashes
    )
    return any(c in text for c in special_chars)
