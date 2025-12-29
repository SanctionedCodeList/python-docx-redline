"""
Predefined standard style templates for Word documents.

This module provides factory functions that return Style objects for common
Word document styles such as footnote references, footnote text, endnote
references, endnote text, and hyperlinks.

These templates match Word's built-in style definitions and can be used with
StyleManager.ensure_style() or the ensure_standard_styles() helper function.

Example:
    >>> from python_docx_redline.style_templates import (
    ...     ensure_standard_styles,
    ...     get_footnote_reference_style,
    ... )
    >>> # Use factory function directly
    >>> style = get_footnote_reference_style()
    >>> styles.add(style)
    >>> # Or use the helper to ensure multiple styles at once
    >>> ensure_standard_styles(
    ...     styles,
    ...     "FootnoteReference",
    ...     "FootnoteText",
    ...     "FootnoteTextChar",
    ... )
"""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from .models.style import (
    ParagraphFormatting,
    RunFormatting,
    Style,
    StyleType,
)

if TYPE_CHECKING:
    from .styles import StyleManager


def get_footnote_reference_style() -> Style:
    """Get the standard FootnoteReference style definition.

    This is a character style that formats footnote reference marks with
    superscript formatting. It's based on DefaultParagraphFont and is
    typically hidden from the UI until used.

    Returns:
        A Style object configured as a footnote reference character style.

    Example:
        >>> style = get_footnote_reference_style()
        >>> style.style_id
        'FootnoteReference'
        >>> style.run_formatting.superscript
        True
    """
    return Style(
        style_id="FootnoteReference",
        name="footnote reference",
        style_type=StyleType.CHARACTER,
        based_on="DefaultParagraphFont",
        run_formatting=RunFormatting(superscript=True),
        ui_priority=99,
        semi_hidden=True,
        unhide_when_used=True,
    )


def get_footnote_text_style() -> Style:
    """Get the standard FootnoteText style definition.

    This is a paragraph style for footnote text content. It uses smaller
    font size (10pt) and single line spacing with no spacing after paragraphs.
    It's linked to the FootnoteTextChar character style.

    Returns:
        A Style object configured as a footnote text paragraph style.

    Example:
        >>> style = get_footnote_text_style()
        >>> style.style_id
        'FootnoteText'
        >>> style.run_formatting.font_size
        10.0
    """
    return Style(
        style_id="FootnoteText",
        name="footnote text",
        style_type=StyleType.PARAGRAPH,
        based_on="Normal",
        linked_style="FootnoteTextChar",
        paragraph_formatting=ParagraphFormatting(
            spacing_after=0,
            line_spacing=1.0,
        ),
        run_formatting=RunFormatting(font_size=10),
        ui_priority=99,
        semi_hidden=True,
        unhide_when_used=True,
    )


def get_footnote_text_char_style() -> Style:
    """Get the FootnoteTextChar linked character style definition.

    This is a character style linked to FootnoteText. It applies the same
    character formatting (10pt font) when used inline within other paragraphs.

    Returns:
        A Style object configured as a footnote text character style.

    Example:
        >>> style = get_footnote_text_char_style()
        >>> style.style_id
        'FootnoteTextChar'
        >>> style.linked_style
        'FootnoteText'
    """
    return Style(
        style_id="FootnoteTextChar",
        name="Footnote Text Char",
        style_type=StyleType.CHARACTER,
        based_on="DefaultParagraphFont",
        linked_style="FootnoteText",
        run_formatting=RunFormatting(font_size=10),
        ui_priority=99,
        semi_hidden=True,
        unhide_when_used=True,
    )


def get_endnote_reference_style() -> Style:
    """Get the standard EndnoteReference style definition.

    This is a character style that formats endnote reference marks with
    superscript formatting. It's similar to FootnoteReference and is based
    on DefaultParagraphFont.

    Returns:
        A Style object configured as an endnote reference character style.

    Example:
        >>> style = get_endnote_reference_style()
        >>> style.style_id
        'EndnoteReference'
        >>> style.run_formatting.superscript
        True
    """
    return Style(
        style_id="EndnoteReference",
        name="endnote reference",
        style_type=StyleType.CHARACTER,
        based_on="DefaultParagraphFont",
        run_formatting=RunFormatting(superscript=True),
        ui_priority=99,
        semi_hidden=True,
        unhide_when_used=True,
    )


def get_endnote_text_style() -> Style:
    """Get the standard EndnoteText style definition.

    This is a paragraph style for endnote text content. Like FootnoteText,
    it uses smaller font size (10pt) and single line spacing. It's linked
    to the EndnoteTextChar character style.

    Returns:
        A Style object configured as an endnote text paragraph style.

    Example:
        >>> style = get_endnote_text_style()
        >>> style.style_id
        'EndnoteText'
        >>> style.paragraph_formatting.line_spacing
        1.0
    """
    return Style(
        style_id="EndnoteText",
        name="endnote text",
        style_type=StyleType.PARAGRAPH,
        based_on="Normal",
        linked_style="EndnoteTextChar",
        paragraph_formatting=ParagraphFormatting(
            spacing_after=0,
            line_spacing=1.0,
        ),
        run_formatting=RunFormatting(font_size=10),
        ui_priority=99,
        semi_hidden=True,
        unhide_when_used=True,
    )


def get_endnote_text_char_style() -> Style:
    """Get the EndnoteTextChar linked character style definition.

    This is a character style linked to EndnoteText. It applies the same
    character formatting (10pt font) when used inline within other paragraphs.

    Returns:
        A Style object configured as an endnote text character style.

    Example:
        >>> style = get_endnote_text_char_style()
        >>> style.style_id
        'EndnoteTextChar'
        >>> style.linked_style
        'EndnoteText'
    """
    return Style(
        style_id="EndnoteTextChar",
        name="Endnote Text Char",
        style_type=StyleType.CHARACTER,
        based_on="DefaultParagraphFont",
        linked_style="EndnoteText",
        run_formatting=RunFormatting(font_size=10),
        ui_priority=99,
        semi_hidden=True,
        unhide_when_used=True,
    )


def get_hyperlink_style() -> Style:
    """Get the standard Hyperlink style definition.

    This is a character style for hyperlinks, using the standard blue color
    (#0563C1) with single underline formatting.

    Returns:
        A Style object configured as a hyperlink character style.

    Example:
        >>> style = get_hyperlink_style()
        >>> style.style_id
        'Hyperlink'
        >>> style.run_formatting.color
        '#0563C1'
        >>> style.run_formatting.underline
        'single'
    """
    return Style(
        style_id="Hyperlink",
        name="Hyperlink",
        style_type=StyleType.CHARACTER,
        based_on="DefaultParagraphFont",
        run_formatting=RunFormatting(
            color="#0563C1",
            underline="single",
        ),
        ui_priority=99,
        unhide_when_used=True,
    )


# Dictionary mapping style IDs to factory functions
STANDARD_STYLES: dict[str, Callable[[], Style]] = {
    "FootnoteReference": get_footnote_reference_style,
    "FootnoteText": get_footnote_text_style,
    "FootnoteTextChar": get_footnote_text_char_style,
    "EndnoteReference": get_endnote_reference_style,
    "EndnoteText": get_endnote_text_style,
    "EndnoteTextChar": get_endnote_text_char_style,
    "Hyperlink": get_hyperlink_style,
}


def ensure_standard_styles(style_manager: StyleManager, *style_ids: str) -> None:
    """Ensure multiple standard styles exist in the document.

    This is a convenience function for ensuring that one or more standard
    styles exist in the document. For each style_id provided, if the style
    doesn't already exist, it will be created using the predefined template.

    Args:
        style_manager: The StyleManager to add styles to
        *style_ids: Style IDs to ensure (e.g., "FootnoteReference", "FootnoteText")

    Raises:
        ValueError: If an unknown style ID is provided

    Example:
        >>> ensure_standard_styles(
        ...     doc.styles,
        ...     "FootnoteReference", "FootnoteText", "FootnoteTextChar"
        ... )
        >>> # All three styles now exist in the document
        >>> "FootnoteReference" in doc.styles
        True
    """
    for style_id in style_ids:
        if style_id not in STANDARD_STYLES:
            available = ", ".join(sorted(STANDARD_STYLES.keys()))
            raise ValueError(
                f"Unknown standard style ID: '{style_id}'. Available styles: {available}"
            )

        # Only add if the style doesn't already exist
        if style_id not in style_manager:
            style = STANDARD_STYLES[style_id]()
            style_manager.add(style)
