"""
Style model classes for Word document style management.

Provides data classes for representing styles, run formatting, and paragraph
formatting in a Pythonic way, hiding the underlying OOXML complexity.

These models are used by StyleManager to read and write styles to/from
word/styles.xml.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class StyleType(Enum):
    """Types of styles in Word documents.

    Word documents support four types of styles, each controlling
    different aspects of document formatting.

    Attributes:
        PARAGRAPH: Applied to whole paragraphs (includes both paragraph
            and character formatting)
        CHARACTER: Applied to runs of text within paragraphs
        TABLE: Applied to tables
        NUMBERING: Applied to numbered/bulleted lists
    """

    PARAGRAPH = "paragraph"
    CHARACTER = "character"
    TABLE = "table"
    NUMBERING = "numbering"


@dataclass
class RunFormatting:
    """Character-level formatting properties.

    Represents formatting that can be applied to runs (spans) of text
    within a paragraph. All properties are optional (None means inherit
    from parent style or document defaults).

    Attributes:
        bold: Whether text is bold
        italic: Whether text is italic
        underline: Underline style (True for single underline, or a
            style name like "double", "wave", "dotted", etc.)
        strikethrough: Whether text has strikethrough
        font_name: Font family name (e.g., "Arial", "Times New Roman")
        font_size: Font size in points (e.g., 12.0)
        color: Text color as hex string "#RRGGBB" or "auto"
        highlight: Highlight color name (e.g., "yellow", "green")
        superscript: Whether text is superscript
        subscript: Whether text is subscript
        small_caps: Whether text is in small capitals
        all_caps: Whether text is in all capitals

    Example:
        >>> fmt = RunFormatting(bold=True, font_size=14.0, color="#FF0000")
        >>> fmt.bold
        True
    """

    bold: bool | None = None
    italic: bool | None = None
    underline: bool | str | None = None
    strikethrough: bool | None = None
    font_name: str | None = None
    font_size: float | None = None
    color: str | None = None
    highlight: str | None = None
    superscript: bool | None = None
    subscript: bool | None = None
    small_caps: bool | None = None
    all_caps: bool | None = None


@dataclass
class ParagraphFormatting:
    """Paragraph-level formatting properties.

    Represents formatting that applies to entire paragraphs. All properties
    are optional (None means inherit from parent style or document defaults).

    Attributes:
        alignment: Text alignment ("left", "center", "right", "justify")
        spacing_before: Space before paragraph in points
        spacing_after: Space after paragraph in points
        line_spacing: Line spacing multiplier (1.0 = single, 1.5 = 1.5 lines,
            2.0 = double)
        indent_left: Left indent in inches
        indent_right: Right indent in inches
        indent_first_line: First line indent in inches (positive = indent,
            use indent_hanging for negative)
        indent_hanging: Hanging indent in inches (positive = hanging)
        keep_next: Keep paragraph with next paragraph on same page
        keep_lines: Keep all lines of paragraph on same page
        outline_level: Heading outline level (0-8, where 0 = Heading 1)

    Example:
        >>> fmt = ParagraphFormatting(
        ...     alignment="justify",
        ...     spacing_after=12.0,
        ...     line_spacing=1.5
        ... )
    """

    alignment: str | None = None
    spacing_before: float | None = None
    spacing_after: float | None = None
    line_spacing: float | None = None
    indent_left: float | None = None
    indent_right: float | None = None
    indent_first_line: float | None = None
    indent_hanging: float | None = None
    keep_next: bool | None = None
    keep_lines: bool | None = None
    outline_level: int | None = None


@dataclass
class Style:
    """Represents a Word document style.

    A style combines an identifier, metadata, and formatting properties
    that can be applied to paragraphs or character runs. Styles can
    inherit from other styles via the based_on property.

    Attributes:
        style_id: Internal style identifier used in document references
            (e.g., "Heading1", "FootnoteReference")
        name: Display name shown in Word's UI (e.g., "Heading 1",
            "footnote reference")
        style_type: Type of style (paragraph, character, table, numbering)
        based_on: style_id of parent style to inherit from
        next_style: style_id of style to apply after pressing Enter
            (paragraph styles only)
        linked_style: style_id of linked style (character style linked
            to paragraph style or vice versa)
        run_formatting: Character formatting properties
        paragraph_formatting: Paragraph formatting properties (only used
            for paragraph styles)
        ui_priority: Sort order in Word's style gallery (lower = higher
            priority)
        quick_format: Whether style appears in Quick Style gallery
        semi_hidden: Whether style is hidden from UI
        unhide_when_used: Whether to unhide style when first used

    Example:
        >>> style = Style(
        ...     style_id="FootnoteReference",
        ...     name="footnote reference",
        ...     style_type=StyleType.CHARACTER,
        ...     based_on="DefaultParagraphFont",
        ...     run_formatting=RunFormatting(superscript=True),
        ...     ui_priority=99,
        ...     unhide_when_used=True,
        ... )
    """

    style_id: str
    name: str
    style_type: StyleType
    based_on: str | None = None
    next_style: str | None = None
    linked_style: str | None = None
    run_formatting: RunFormatting = field(default_factory=RunFormatting)
    paragraph_formatting: ParagraphFormatting = field(default_factory=ParagraphFormatting)
    ui_priority: int | None = None
    quick_format: bool = False
    semi_hidden: bool = False
    unhide_when_used: bool = False
    _element: Any = field(default=None, repr=False, compare=False)

    def __repr__(self) -> str:
        """String representation of the style."""
        return f"<Style style_id={self.style_id!r} name={self.name!r} type={self.style_type.value}>"
