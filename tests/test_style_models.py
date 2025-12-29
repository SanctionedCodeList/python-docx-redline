"""
Tests for the style data models.

These tests verify the core style types used for Word document styling:
- StyleType enum
- RunFormatting dataclass
- ParagraphFormatting dataclass
- Style dataclass
"""

import pytest

from python_docx_redline.models.style import (
    ParagraphFormatting,
    RunFormatting,
    Style,
    StyleType,
)


class TestStyleType:
    """Tests for StyleType enum."""

    def test_paragraph_type_exists(self) -> None:
        """Test that PARAGRAPH type exists with correct value."""
        assert StyleType.PARAGRAPH.value == "paragraph"

    def test_character_type_exists(self) -> None:
        """Test that CHARACTER type exists with correct value."""
        assert StyleType.CHARACTER.value == "character"

    def test_table_type_exists(self) -> None:
        """Test that TABLE type exists with correct value."""
        assert StyleType.TABLE.value == "table"

    def test_numbering_type_exists(self) -> None:
        """Test that NUMBERING type exists with correct value."""
        assert StyleType.NUMBERING.value == "numbering"

    def test_all_four_types_present(self) -> None:
        """Test that exactly four style types are defined."""
        all_types = list(StyleType)
        assert len(all_types) == 4
        assert set(all_types) == {
            StyleType.PARAGRAPH,
            StyleType.CHARACTER,
            StyleType.TABLE,
            StyleType.NUMBERING,
        }

    def test_enum_from_value(self) -> None:
        """Test creating enum from string value."""
        assert StyleType("paragraph") == StyleType.PARAGRAPH
        assert StyleType("character") == StyleType.CHARACTER
        assert StyleType("table") == StyleType.TABLE
        assert StyleType("numbering") == StyleType.NUMBERING

    def test_invalid_value_raises(self) -> None:
        """Test that invalid value raises ValueError."""
        with pytest.raises(ValueError):
            StyleType("invalid")


class TestRunFormatting:
    """Tests for RunFormatting dataclass."""

    def test_default_values_all_none(self) -> None:
        """Test that all default values are None."""
        fmt = RunFormatting()

        assert fmt.bold is None
        assert fmt.italic is None
        assert fmt.underline is None
        assert fmt.strikethrough is None
        assert fmt.font_name is None
        assert fmt.font_size is None
        assert fmt.color is None
        assert fmt.highlight is None
        assert fmt.superscript is None
        assert fmt.subscript is None
        assert fmt.small_caps is None
        assert fmt.all_caps is None

    def test_create_with_bold(self) -> None:
        """Test creating with bold formatting."""
        fmt = RunFormatting(bold=True)

        assert fmt.bold is True
        assert fmt.italic is None

    def test_create_with_italic(self) -> None:
        """Test creating with italic formatting."""
        fmt = RunFormatting(italic=True)

        assert fmt.italic is True

    def test_create_with_underline_bool(self) -> None:
        """Test creating with boolean underline."""
        fmt = RunFormatting(underline=True)

        assert fmt.underline is True

    def test_create_with_underline_string(self) -> None:
        """Test creating with underline style string."""
        fmt = RunFormatting(underline="double")

        assert fmt.underline == "double"

    def test_create_with_font_name(self) -> None:
        """Test creating with font name."""
        fmt = RunFormatting(font_name="Arial")

        assert fmt.font_name == "Arial"

    def test_create_with_font_size(self) -> None:
        """Test creating with font size."""
        fmt = RunFormatting(font_size=14.0)

        assert fmt.font_size == 14.0

    def test_create_with_color(self) -> None:
        """Test creating with text color."""
        fmt = RunFormatting(color="#FF0000")

        assert fmt.color == "#FF0000"

    def test_create_with_highlight(self) -> None:
        """Test creating with highlight color."""
        fmt = RunFormatting(highlight="yellow")

        assert fmt.highlight == "yellow"

    def test_create_with_superscript(self) -> None:
        """Test creating with superscript."""
        fmt = RunFormatting(superscript=True)

        assert fmt.superscript is True
        assert fmt.subscript is None

    def test_create_with_subscript(self) -> None:
        """Test creating with subscript."""
        fmt = RunFormatting(subscript=True)

        assert fmt.subscript is True

    def test_create_with_small_caps(self) -> None:
        """Test creating with small caps."""
        fmt = RunFormatting(small_caps=True)

        assert fmt.small_caps is True

    def test_create_with_all_caps(self) -> None:
        """Test creating with all caps."""
        fmt = RunFormatting(all_caps=True)

        assert fmt.all_caps is True

    def test_create_with_strikethrough(self) -> None:
        """Test creating with strikethrough."""
        fmt = RunFormatting(strikethrough=True)

        assert fmt.strikethrough is True

    def test_create_with_multiple_properties(self) -> None:
        """Test creating with multiple formatting properties."""
        fmt = RunFormatting(
            bold=True,
            italic=True,
            font_size=12.0,
            color="#0000FF",
        )

        assert fmt.bold is True
        assert fmt.italic is True
        assert fmt.font_size == 12.0
        assert fmt.color == "#0000FF"
        assert fmt.underline is None

    def test_all_properties_accessible(self) -> None:
        """Test that all 12 properties are accessible."""
        fmt = RunFormatting()
        properties = [
            "bold",
            "italic",
            "underline",
            "strikethrough",
            "font_name",
            "font_size",
            "color",
            "highlight",
            "superscript",
            "subscript",
            "small_caps",
            "all_caps",
        ]

        for prop in properties:
            assert hasattr(fmt, prop)


class TestParagraphFormatting:
    """Tests for ParagraphFormatting dataclass."""

    def test_default_values_all_none(self) -> None:
        """Test that all default values are None."""
        fmt = ParagraphFormatting()

        assert fmt.alignment is None
        assert fmt.spacing_before is None
        assert fmt.spacing_after is None
        assert fmt.line_spacing is None
        assert fmt.indent_left is None
        assert fmt.indent_right is None
        assert fmt.indent_first_line is None
        assert fmt.indent_hanging is None
        assert fmt.keep_next is None
        assert fmt.keep_lines is None
        assert fmt.outline_level is None

    def test_create_with_alignment(self) -> None:
        """Test creating with text alignment."""
        fmt = ParagraphFormatting(alignment="center")

        assert fmt.alignment == "center"

    def test_create_with_left_alignment(self) -> None:
        """Test creating with left alignment."""
        fmt = ParagraphFormatting(alignment="left")

        assert fmt.alignment == "left"

    def test_create_with_right_alignment(self) -> None:
        """Test creating with right alignment."""
        fmt = ParagraphFormatting(alignment="right")

        assert fmt.alignment == "right"

    def test_create_with_justify_alignment(self) -> None:
        """Test creating with justify alignment."""
        fmt = ParagraphFormatting(alignment="justify")

        assert fmt.alignment == "justify"

    def test_create_with_spacing_before(self) -> None:
        """Test creating with spacing before."""
        fmt = ParagraphFormatting(spacing_before=12.0)

        assert fmt.spacing_before == 12.0

    def test_create_with_spacing_after(self) -> None:
        """Test creating with spacing after."""
        fmt = ParagraphFormatting(spacing_after=6.0)

        assert fmt.spacing_after == 6.0

    def test_create_with_line_spacing(self) -> None:
        """Test creating with line spacing."""
        fmt = ParagraphFormatting(line_spacing=1.5)

        assert fmt.line_spacing == 1.5

    def test_create_with_double_line_spacing(self) -> None:
        """Test creating with double line spacing."""
        fmt = ParagraphFormatting(line_spacing=2.0)

        assert fmt.line_spacing == 2.0

    def test_create_with_indent_left(self) -> None:
        """Test creating with left indent."""
        fmt = ParagraphFormatting(indent_left=0.5)

        assert fmt.indent_left == 0.5

    def test_create_with_indent_right(self) -> None:
        """Test creating with right indent."""
        fmt = ParagraphFormatting(indent_right=0.25)

        assert fmt.indent_right == 0.25

    def test_create_with_first_line_indent(self) -> None:
        """Test creating with first line indent."""
        fmt = ParagraphFormatting(indent_first_line=0.5)

        assert fmt.indent_first_line == 0.5

    def test_create_with_hanging_indent(self) -> None:
        """Test creating with hanging indent."""
        fmt = ParagraphFormatting(indent_hanging=0.25)

        assert fmt.indent_hanging == 0.25

    def test_create_with_keep_next(self) -> None:
        """Test creating with keep next."""
        fmt = ParagraphFormatting(keep_next=True)

        assert fmt.keep_next is True

    def test_create_with_keep_lines(self) -> None:
        """Test creating with keep lines."""
        fmt = ParagraphFormatting(keep_lines=True)

        assert fmt.keep_lines is True

    def test_create_with_outline_level(self) -> None:
        """Test creating with outline level."""
        fmt = ParagraphFormatting(outline_level=0)

        assert fmt.outline_level == 0

    def test_create_with_outline_level_heading2(self) -> None:
        """Test creating with outline level for Heading 2."""
        fmt = ParagraphFormatting(outline_level=1)

        assert fmt.outline_level == 1

    def test_create_with_multiple_properties(self) -> None:
        """Test creating with multiple formatting properties."""
        fmt = ParagraphFormatting(
            alignment="justify",
            spacing_before=12.0,
            spacing_after=12.0,
            line_spacing=1.5,
            indent_left=0.5,
        )

        assert fmt.alignment == "justify"
        assert fmt.spacing_before == 12.0
        assert fmt.spacing_after == 12.0
        assert fmt.line_spacing == 1.5
        assert fmt.indent_left == 0.5

    def test_all_properties_accessible(self) -> None:
        """Test that all 11 properties are accessible."""
        fmt = ParagraphFormatting()
        properties = [
            "alignment",
            "spacing_before",
            "spacing_after",
            "line_spacing",
            "indent_left",
            "indent_right",
            "indent_first_line",
            "indent_hanging",
            "keep_next",
            "keep_lines",
            "outline_level",
        ]

        for prop in properties:
            assert hasattr(fmt, prop)


class TestStyle:
    """Tests for Style dataclass."""

    def test_create_minimal_style(self) -> None:
        """Test creating a style with only required fields."""
        style = Style(
            style_id="Normal",
            name="Normal",
            style_type=StyleType.PARAGRAPH,
        )

        assert style.style_id == "Normal"
        assert style.name == "Normal"
        assert style.style_type == StyleType.PARAGRAPH

    def test_minimal_style_default_values(self) -> None:
        """Test that minimal style has correct default values."""
        style = Style(
            style_id="Normal",
            name="Normal",
            style_type=StyleType.PARAGRAPH,
        )

        assert style.based_on is None
        assert style.next_style is None
        assert style.linked_style is None
        assert style.ui_priority is None
        assert style.quick_format is False
        assert style.semi_hidden is False
        assert style.unhide_when_used is False

    def test_default_factory_run_formatting(self) -> None:
        """Test that run_formatting has default factory creating empty RunFormatting."""
        style = Style(
            style_id="Test",
            name="Test",
            style_type=StyleType.CHARACTER,
        )

        assert isinstance(style.run_formatting, RunFormatting)
        assert style.run_formatting.bold is None

    def test_default_factory_paragraph_formatting(self) -> None:
        """Test that paragraph_formatting has default factory creating empty ParagraphFormatting."""
        style = Style(
            style_id="Test",
            name="Test",
            style_type=StyleType.PARAGRAPH,
        )

        assert isinstance(style.paragraph_formatting, ParagraphFormatting)
        assert style.paragraph_formatting.alignment is None

    def test_create_full_style(self) -> None:
        """Test creating a style with all properties."""
        run_fmt = RunFormatting(bold=True, font_size=24.0)
        para_fmt = ParagraphFormatting(
            alignment="center",
            spacing_after=12.0,
            outline_level=0,
        )

        style = Style(
            style_id="Heading1",
            name="Heading 1",
            style_type=StyleType.PARAGRAPH,
            based_on="Normal",
            next_style="Normal",
            linked_style="Heading1Char",
            run_formatting=run_fmt,
            paragraph_formatting=para_fmt,
            ui_priority=9,
            quick_format=True,
            semi_hidden=False,
            unhide_when_used=False,
        )

        assert style.style_id == "Heading1"
        assert style.name == "Heading 1"
        assert style.style_type == StyleType.PARAGRAPH
        assert style.based_on == "Normal"
        assert style.next_style == "Normal"
        assert style.linked_style == "Heading1Char"
        assert style.run_formatting.bold is True
        assert style.run_formatting.font_size == 24.0
        assert style.paragraph_formatting.alignment == "center"
        assert style.paragraph_formatting.outline_level == 0
        assert style.ui_priority == 9
        assert style.quick_format is True

    def test_element_excluded_from_repr(self) -> None:
        """Test that _element is excluded from repr."""
        style = Style(
            style_id="Test",
            name="Test",
            style_type=StyleType.CHARACTER,
            _element=object(),  # Some mock element
        )

        repr_str = repr(style)
        assert "_element" not in repr_str
        assert "Test" in repr_str

    def test_element_excluded_from_compare(self) -> None:
        """Test that _element is excluded from comparison."""
        style1 = Style(
            style_id="Test",
            name="Test",
            style_type=StyleType.CHARACTER,
            _element=object(),
        )
        style2 = Style(
            style_id="Test",
            name="Test",
            style_type=StyleType.CHARACTER,
            _element=object(),  # Different object
        )

        # Styles should be equal even with different _element values
        assert style1 == style2

    def test_custom_repr(self) -> None:
        """Test the custom __repr__ method."""
        style = Style(
            style_id="Heading1",
            name="Heading 1",
            style_type=StyleType.PARAGRAPH,
        )

        repr_str = repr(style)
        assert "<Style" in repr_str
        assert "style_id='Heading1'" in repr_str
        assert "name='Heading 1'" in repr_str
        assert "type=paragraph" in repr_str

    def test_character_style_repr(self) -> None:
        """Test repr for character style."""
        style = Style(
            style_id="Strong",
            name="Strong",
            style_type=StyleType.CHARACTER,
        )

        repr_str = repr(style)
        assert "type=character" in repr_str


class TestRealWorldStyleScenarios:
    """Tests for real-world style scenarios."""

    def test_footnote_reference_style(self) -> None:
        """Test creating a FootnoteReference-like character style."""
        style = Style(
            style_id="FootnoteReference",
            name="footnote reference",
            style_type=StyleType.CHARACTER,
            based_on="DefaultParagraphFont",
            run_formatting=RunFormatting(superscript=True),
            ui_priority=99,
            unhide_when_used=True,
            semi_hidden=True,
        )

        assert style.style_id == "FootnoteReference"
        assert style.style_type == StyleType.CHARACTER
        assert style.run_formatting.superscript is True
        assert style.ui_priority == 99
        assert style.unhide_when_used is True
        assert style.semi_hidden is True

    def test_heading1_style(self) -> None:
        """Test creating a Heading1-like paragraph style."""
        style = Style(
            style_id="Heading1",
            name="Heading 1",
            style_type=StyleType.PARAGRAPH,
            based_on="Normal",
            next_style="Normal",
            linked_style="Heading1Char",
            run_formatting=RunFormatting(
                bold=True,
                font_size=16.0,
                color="#2F5496",
            ),
            paragraph_formatting=ParagraphFormatting(
                spacing_before=12.0,
                spacing_after=0.0,
                keep_next=True,
                keep_lines=True,
                outline_level=0,
            ),
            ui_priority=9,
            quick_format=True,
        )

        assert style.style_id == "Heading1"
        assert style.style_type == StyleType.PARAGRAPH
        assert style.run_formatting.bold is True
        assert style.run_formatting.font_size == 16.0
        assert style.paragraph_formatting.outline_level == 0
        assert style.paragraph_formatting.keep_next is True
        assert style.quick_format is True

    def test_normal_style(self) -> None:
        """Test creating a Normal-like paragraph style."""
        style = Style(
            style_id="Normal",
            name="Normal",
            style_type=StyleType.PARAGRAPH,
            run_formatting=RunFormatting(
                font_name="Calibri",
                font_size=11.0,
            ),
            paragraph_formatting=ParagraphFormatting(
                spacing_after=8.0,
                line_spacing=1.15,
            ),
            quick_format=True,
        )

        assert style.style_id == "Normal"
        assert style.style_type == StyleType.PARAGRAPH
        assert style.based_on is None  # Normal typically has no parent
        assert style.run_formatting.font_name == "Calibri"
        assert style.run_formatting.font_size == 11.0
        assert style.paragraph_formatting.spacing_after == 8.0
        assert style.paragraph_formatting.line_spacing == 1.15

    def test_strong_character_style(self) -> None:
        """Test creating a Strong (bold) character style."""
        style = Style(
            style_id="Strong",
            name="Strong",
            style_type=StyleType.CHARACTER,
            based_on="DefaultParagraphFont",
            run_formatting=RunFormatting(bold=True),
            ui_priority=22,
            quick_format=True,
        )

        assert style.style_id == "Strong"
        assert style.style_type == StyleType.CHARACTER
        assert style.run_formatting.bold is True
        assert style.quick_format is True

    def test_emphasis_character_style(self) -> None:
        """Test creating an Emphasis (italic) character style."""
        style = Style(
            style_id="Emphasis",
            name="Emphasis",
            style_type=StyleType.CHARACTER,
            based_on="DefaultParagraphFont",
            run_formatting=RunFormatting(italic=True),
            ui_priority=20,
            quick_format=True,
        )

        assert style.style_id == "Emphasis"
        assert style.style_type == StyleType.CHARACTER
        assert style.run_formatting.italic is True

    def test_quote_paragraph_style(self) -> None:
        """Test creating a Quote paragraph style."""
        style = Style(
            style_id="Quote",
            name="Quote",
            style_type=StyleType.PARAGRAPH,
            based_on="Normal",
            next_style="Normal",
            run_formatting=RunFormatting(
                italic=True,
                color="#404040",
            ),
            paragraph_formatting=ParagraphFormatting(
                indent_left=0.5,
                indent_right=0.5,
                spacing_before=12.0,
                spacing_after=12.0,
            ),
            ui_priority=29,
            quick_format=True,
        )

        assert style.style_id == "Quote"
        assert style.run_formatting.italic is True
        assert style.paragraph_formatting.indent_left == 0.5
        assert style.paragraph_formatting.indent_right == 0.5

    def test_heading2_style_with_outline_level(self) -> None:
        """Test creating a Heading2-like style with outline level 1."""
        style = Style(
            style_id="Heading2",
            name="Heading 2",
            style_type=StyleType.PARAGRAPH,
            based_on="Normal",
            next_style="Normal",
            run_formatting=RunFormatting(
                bold=True,
                font_size=13.0,
                color="#2F5496",
            ),
            paragraph_formatting=ParagraphFormatting(
                spacing_before=12.0,
                spacing_after=0.0,
                keep_next=True,
                outline_level=1,  # Heading 2 = outline level 1
            ),
            ui_priority=9,
            quick_format=True,
        )

        assert style.paragraph_formatting.outline_level == 1

    def test_table_style(self) -> None:
        """Test creating a table style."""
        style = Style(
            style_id="TableGrid",
            name="Table Grid",
            style_type=StyleType.TABLE,
            ui_priority=59,
        )

        assert style.style_id == "TableGrid"
        assert style.style_type == StyleType.TABLE
