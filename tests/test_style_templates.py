"""
Tests for the style_templates module.

These tests verify the predefined standard style templates and the
ensure_standard_styles helper function.
"""

from unittest.mock import MagicMock

import pytest

from python_docx_redline.models.style import StyleType
from python_docx_redline.style_templates import (
    STANDARD_STYLES,
    ensure_standard_styles,
    ensure_toc_styles,
    get_endnote_reference_style,
    get_endnote_text_char_style,
    get_endnote_text_style,
    get_footnote_reference_style,
    get_footnote_text_char_style,
    get_footnote_text_style,
    get_hyperlink_style,
    get_toc_heading_style,
    get_toc_level_style,
)


class TestFootnoteReferenceStyle:
    """Tests for get_footnote_reference_style factory function."""

    def test_returns_style_object(self) -> None:
        """Test that factory returns a Style object."""
        style = get_footnote_reference_style()
        assert style is not None

    def test_style_id(self) -> None:
        """Test that style_id is FootnoteReference."""
        style = get_footnote_reference_style()
        assert style.style_id == "FootnoteReference"

    def test_name(self) -> None:
        """Test that name is 'footnote reference'."""
        style = get_footnote_reference_style()
        assert style.name == "footnote reference"

    def test_style_type_is_character(self) -> None:
        """Test that style type is CHARACTER."""
        style = get_footnote_reference_style()
        assert style.style_type == StyleType.CHARACTER

    def test_based_on_default_paragraph_font(self) -> None:
        """Test that style is based on DefaultParagraphFont."""
        style = get_footnote_reference_style()
        assert style.based_on == "DefaultParagraphFont"

    def test_superscript_formatting(self) -> None:
        """Test that run formatting includes superscript."""
        style = get_footnote_reference_style()
        assert style.run_formatting.superscript is True

    def test_ui_priority(self) -> None:
        """Test that ui_priority is 99."""
        style = get_footnote_reference_style()
        assert style.ui_priority == 99

    def test_semi_hidden(self) -> None:
        """Test that style is semi-hidden."""
        style = get_footnote_reference_style()
        assert style.semi_hidden is True

    def test_unhide_when_used(self) -> None:
        """Test that unhide_when_used is True."""
        style = get_footnote_reference_style()
        assert style.unhide_when_used is True


class TestFootnoteTextStyle:
    """Tests for get_footnote_text_style factory function."""

    def test_returns_style_object(self) -> None:
        """Test that factory returns a Style object."""
        style = get_footnote_text_style()
        assert style is not None

    def test_style_id(self) -> None:
        """Test that style_id is FootnoteText."""
        style = get_footnote_text_style()
        assert style.style_id == "FootnoteText"

    def test_name(self) -> None:
        """Test that name is 'footnote text'."""
        style = get_footnote_text_style()
        assert style.name == "footnote text"

    def test_style_type_is_paragraph(self) -> None:
        """Test that style type is PARAGRAPH."""
        style = get_footnote_text_style()
        assert style.style_type == StyleType.PARAGRAPH

    def test_based_on_normal(self) -> None:
        """Test that style is based on Normal."""
        style = get_footnote_text_style()
        assert style.based_on == "Normal"

    def test_linked_to_footnote_text_char(self) -> None:
        """Test that style is linked to FootnoteTextChar."""
        style = get_footnote_text_style()
        assert style.linked_style == "FootnoteTextChar"

    def test_spacing_after_zero(self) -> None:
        """Test that spacing_after is 0."""
        style = get_footnote_text_style()
        assert style.paragraph_formatting.spacing_after == 0

    def test_line_spacing_single(self) -> None:
        """Test that line spacing is 1.0 (single)."""
        style = get_footnote_text_style()
        assert style.paragraph_formatting.line_spacing == 1.0

    def test_font_size_10pt(self) -> None:
        """Test that font size is 10 points."""
        style = get_footnote_text_style()
        assert style.run_formatting.font_size == 10

    def test_ui_priority(self) -> None:
        """Test that ui_priority is 99."""
        style = get_footnote_text_style()
        assert style.ui_priority == 99

    def test_semi_hidden(self) -> None:
        """Test that style is semi-hidden."""
        style = get_footnote_text_style()
        assert style.semi_hidden is True


class TestFootnoteTextCharStyle:
    """Tests for get_footnote_text_char_style factory function."""

    def test_returns_style_object(self) -> None:
        """Test that factory returns a Style object."""
        style = get_footnote_text_char_style()
        assert style is not None

    def test_style_id(self) -> None:
        """Test that style_id is FootnoteTextChar."""
        style = get_footnote_text_char_style()
        assert style.style_id == "FootnoteTextChar"

    def test_name(self) -> None:
        """Test that name is 'Footnote Text Char'."""
        style = get_footnote_text_char_style()
        assert style.name == "Footnote Text Char"

    def test_style_type_is_character(self) -> None:
        """Test that style type is CHARACTER."""
        style = get_footnote_text_char_style()
        assert style.style_type == StyleType.CHARACTER

    def test_based_on_default_paragraph_font(self) -> None:
        """Test that style is based on DefaultParagraphFont."""
        style = get_footnote_text_char_style()
        assert style.based_on == "DefaultParagraphFont"

    def test_linked_to_footnote_text(self) -> None:
        """Test that style is linked to FootnoteText."""
        style = get_footnote_text_char_style()
        assert style.linked_style == "FootnoteText"

    def test_font_size_10pt(self) -> None:
        """Test that font size is 10 points."""
        style = get_footnote_text_char_style()
        assert style.run_formatting.font_size == 10


class TestEndnoteReferenceStyle:
    """Tests for get_endnote_reference_style factory function."""

    def test_returns_style_object(self) -> None:
        """Test that factory returns a Style object."""
        style = get_endnote_reference_style()
        assert style is not None

    def test_style_id(self) -> None:
        """Test that style_id is EndnoteReference."""
        style = get_endnote_reference_style()
        assert style.style_id == "EndnoteReference"

    def test_name(self) -> None:
        """Test that name is 'endnote reference'."""
        style = get_endnote_reference_style()
        assert style.name == "endnote reference"

    def test_style_type_is_character(self) -> None:
        """Test that style type is CHARACTER."""
        style = get_endnote_reference_style()
        assert style.style_type == StyleType.CHARACTER

    def test_based_on_default_paragraph_font(self) -> None:
        """Test that style is based on DefaultParagraphFont."""
        style = get_endnote_reference_style()
        assert style.based_on == "DefaultParagraphFont"

    def test_superscript_formatting(self) -> None:
        """Test that run formatting includes superscript."""
        style = get_endnote_reference_style()
        assert style.run_formatting.superscript is True

    def test_ui_priority(self) -> None:
        """Test that ui_priority is 99."""
        style = get_endnote_reference_style()
        assert style.ui_priority == 99


class TestEndnoteTextStyle:
    """Tests for get_endnote_text_style factory function."""

    def test_returns_style_object(self) -> None:
        """Test that factory returns a Style object."""
        style = get_endnote_text_style()
        assert style is not None

    def test_style_id(self) -> None:
        """Test that style_id is EndnoteText."""
        style = get_endnote_text_style()
        assert style.style_id == "EndnoteText"

    def test_name(self) -> None:
        """Test that name is 'endnote text'."""
        style = get_endnote_text_style()
        assert style.name == "endnote text"

    def test_style_type_is_paragraph(self) -> None:
        """Test that style type is PARAGRAPH."""
        style = get_endnote_text_style()
        assert style.style_type == StyleType.PARAGRAPH

    def test_based_on_normal(self) -> None:
        """Test that style is based on Normal."""
        style = get_endnote_text_style()
        assert style.based_on == "Normal"

    def test_linked_to_endnote_text_char(self) -> None:
        """Test that style is linked to EndnoteTextChar."""
        style = get_endnote_text_style()
        assert style.linked_style == "EndnoteTextChar"

    def test_spacing_after_zero(self) -> None:
        """Test that spacing_after is 0."""
        style = get_endnote_text_style()
        assert style.paragraph_formatting.spacing_after == 0

    def test_line_spacing_single(self) -> None:
        """Test that line spacing is 1.0 (single)."""
        style = get_endnote_text_style()
        assert style.paragraph_formatting.line_spacing == 1.0

    def test_font_size_10pt(self) -> None:
        """Test that font size is 10 points."""
        style = get_endnote_text_style()
        assert style.run_formatting.font_size == 10


class TestEndnoteTextCharStyle:
    """Tests for get_endnote_text_char_style factory function."""

    def test_returns_style_object(self) -> None:
        """Test that factory returns a Style object."""
        style = get_endnote_text_char_style()
        assert style is not None

    def test_style_id(self) -> None:
        """Test that style_id is EndnoteTextChar."""
        style = get_endnote_text_char_style()
        assert style.style_id == "EndnoteTextChar"

    def test_name(self) -> None:
        """Test that name is 'Endnote Text Char'."""
        style = get_endnote_text_char_style()
        assert style.name == "Endnote Text Char"

    def test_style_type_is_character(self) -> None:
        """Test that style type is CHARACTER."""
        style = get_endnote_text_char_style()
        assert style.style_type == StyleType.CHARACTER

    def test_linked_to_endnote_text(self) -> None:
        """Test that style is linked to EndnoteText."""
        style = get_endnote_text_char_style()
        assert style.linked_style == "EndnoteText"


class TestHyperlinkStyle:
    """Tests for get_hyperlink_style factory function."""

    def test_returns_style_object(self) -> None:
        """Test that factory returns a Style object."""
        style = get_hyperlink_style()
        assert style is not None

    def test_style_id(self) -> None:
        """Test that style_id is Hyperlink."""
        style = get_hyperlink_style()
        assert style.style_id == "Hyperlink"

    def test_name(self) -> None:
        """Test that name is 'Hyperlink'."""
        style = get_hyperlink_style()
        assert style.name == "Hyperlink"

    def test_style_type_is_character(self) -> None:
        """Test that style type is CHARACTER."""
        style = get_hyperlink_style()
        assert style.style_type == StyleType.CHARACTER

    def test_based_on_default_paragraph_font(self) -> None:
        """Test that style is based on DefaultParagraphFont."""
        style = get_hyperlink_style()
        assert style.based_on == "DefaultParagraphFont"

    def test_blue_color(self) -> None:
        """Test that color is standard hyperlink blue (#0563C1)."""
        style = get_hyperlink_style()
        assert style.run_formatting.color == "#0563C1"

    def test_underline_single(self) -> None:
        """Test that underline style is 'single'."""
        style = get_hyperlink_style()
        assert style.run_formatting.underline == "single"

    def test_ui_priority(self) -> None:
        """Test that ui_priority is 99."""
        style = get_hyperlink_style()
        assert style.ui_priority == 99

    def test_unhide_when_used(self) -> None:
        """Test that unhide_when_used is True."""
        style = get_hyperlink_style()
        assert style.unhide_when_used is True


class TestStandardStylesDict:
    """Tests for the STANDARD_STYLES dictionary."""

    def test_contains_all_standard_styles(self) -> None:
        """Test that dictionary contains all 17 standard styles."""
        expected_styles = {
            "FootnoteReference",
            "FootnoteText",
            "FootnoteTextChar",
            "EndnoteReference",
            "EndnoteText",
            "EndnoteTextChar",
            "Hyperlink",
            # TOC styles
            "TOCHeading",
            "TOC1",
            "TOC2",
            "TOC3",
            "TOC4",
            "TOC5",
            "TOC6",
            "TOC7",
            "TOC8",
            "TOC9",
        }
        assert set(STANDARD_STYLES.keys()) == expected_styles

    def test_values_are_callable(self) -> None:
        """Test that all values are callable (factory functions)."""
        for style_id, factory in STANDARD_STYLES.items():
            assert callable(factory), f"{style_id} value is not callable"

    def test_factories_return_styles(self) -> None:
        """Test that all factories return Style objects."""
        for style_id, factory in STANDARD_STYLES.items():
            style = factory()
            assert style.style_id == style_id, (
                f"Factory for {style_id} returned style with id {style.style_id}"
            )

    def test_footnote_reference_mapping(self) -> None:
        """Test FootnoteReference maps to correct factory."""
        assert STANDARD_STYLES["FootnoteReference"] is get_footnote_reference_style

    def test_footnote_text_mapping(self) -> None:
        """Test FootnoteText maps to correct factory."""
        assert STANDARD_STYLES["FootnoteText"] is get_footnote_text_style

    def test_footnote_text_char_mapping(self) -> None:
        """Test FootnoteTextChar maps to correct factory."""
        assert STANDARD_STYLES["FootnoteTextChar"] is get_footnote_text_char_style

    def test_endnote_reference_mapping(self) -> None:
        """Test EndnoteReference maps to correct factory."""
        assert STANDARD_STYLES["EndnoteReference"] is get_endnote_reference_style

    def test_endnote_text_mapping(self) -> None:
        """Test EndnoteText maps to correct factory."""
        assert STANDARD_STYLES["EndnoteText"] is get_endnote_text_style

    def test_endnote_text_char_mapping(self) -> None:
        """Test EndnoteTextChar maps to correct factory."""
        assert STANDARD_STYLES["EndnoteTextChar"] is get_endnote_text_char_style

    def test_hyperlink_mapping(self) -> None:
        """Test Hyperlink maps to correct factory."""
        assert STANDARD_STYLES["Hyperlink"] is get_hyperlink_style


class TestEnsureStandardStyles:
    """Tests for the ensure_standard_styles helper function."""

    def test_raises_for_unknown_style_id(self) -> None:
        """Test that ValueError is raised for unknown style ID."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        with pytest.raises(ValueError, match="Unknown standard style ID"):
            ensure_standard_styles(mock_manager, "NonExistentStyle")

    def test_error_message_lists_available_styles(self) -> None:
        """Test that error message includes available style IDs."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        with pytest.raises(ValueError) as exc_info:
            ensure_standard_styles(mock_manager, "BadStyle")

        error_msg = str(exc_info.value)
        assert "FootnoteReference" in error_msg
        assert "FootnoteText" in error_msg
        assert "Hyperlink" in error_msg

    def test_adds_style_when_not_present(self) -> None:
        """Test that style is added when not already in manager."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        ensure_standard_styles(mock_manager, "FootnoteReference")

        mock_manager.add.assert_called_once()
        added_style = mock_manager.add.call_args[0][0]
        assert added_style.style_id == "FootnoteReference"

    def test_skips_style_when_already_present(self) -> None:
        """Test that style is not added when already in manager."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=True)

        ensure_standard_styles(mock_manager, "FootnoteReference")

        mock_manager.add.assert_not_called()

    def test_handles_multiple_styles(self) -> None:
        """Test that multiple styles can be ensured at once."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        ensure_standard_styles(
            mock_manager,
            "FootnoteReference",
            "FootnoteText",
            "FootnoteTextChar",
        )

        assert mock_manager.add.call_count == 3

    def test_handles_mixed_existing_and_new_styles(self) -> None:
        """Test handling of mix of existing and new styles."""
        mock_manager = MagicMock()

        # FootnoteReference exists, FootnoteText does not
        def contains_check(style_id: str) -> bool:
            return style_id == "FootnoteReference"

        mock_manager.__contains__ = MagicMock(side_effect=contains_check)

        ensure_standard_styles(
            mock_manager,
            "FootnoteReference",
            "FootnoteText",
        )

        # Only FootnoteText should be added
        assert mock_manager.add.call_count == 1
        added_style = mock_manager.add.call_args[0][0]
        assert added_style.style_id == "FootnoteText"

    def test_handles_empty_style_list(self) -> None:
        """Test that no error occurs with empty style list."""
        mock_manager = MagicMock()

        # Should not raise
        ensure_standard_styles(mock_manager)

        mock_manager.add.assert_not_called()

    def test_stops_on_first_invalid_style(self) -> None:
        """Test that processing stops at first unknown style."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        with pytest.raises(ValueError):
            ensure_standard_styles(
                mock_manager,
                "FootnoteReference",
                "InvalidStyle",
                "FootnoteText",
            )

        # Only FootnoteReference should have been added before the error
        assert mock_manager.add.call_count == 1

    def test_all_standard_styles_can_be_ensured(self) -> None:
        """Test that all standard styles can be ensured without error."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        # Should not raise
        ensure_standard_styles(
            mock_manager,
            "FootnoteReference",
            "FootnoteText",
            "FootnoteTextChar",
            "EndnoteReference",
            "EndnoteText",
            "EndnoteTextChar",
            "Hyperlink",
        )

        assert mock_manager.add.call_count == 7


class TestStyleConsistency:
    """Tests to verify consistency between related styles."""

    def test_footnote_text_and_char_linked(self) -> None:
        """Test that FootnoteText and FootnoteTextChar are properly linked."""
        text_style = get_footnote_text_style()
        char_style = get_footnote_text_char_style()

        assert text_style.linked_style == char_style.style_id
        assert char_style.linked_style == text_style.style_id

    def test_endnote_text_and_char_linked(self) -> None:
        """Test that EndnoteText and EndnoteTextChar are properly linked."""
        text_style = get_endnote_text_style()
        char_style = get_endnote_text_char_style()

        assert text_style.linked_style == char_style.style_id
        assert char_style.linked_style == text_style.style_id

    def test_footnote_and_endnote_reference_similar(self) -> None:
        """Test that footnote and endnote reference styles have similar properties."""
        fn_ref = get_footnote_reference_style()
        en_ref = get_endnote_reference_style()

        assert fn_ref.style_type == en_ref.style_type
        assert fn_ref.based_on == en_ref.based_on
        assert fn_ref.run_formatting.superscript == en_ref.run_formatting.superscript
        assert fn_ref.ui_priority == en_ref.ui_priority

    def test_footnote_and_endnote_text_similar(self) -> None:
        """Test that footnote and endnote text styles have similar properties."""
        fn_text = get_footnote_text_style()
        en_text = get_endnote_text_style()

        assert fn_text.style_type == en_text.style_type
        assert fn_text.based_on == en_text.based_on
        assert (
            fn_text.paragraph_formatting.spacing_after == en_text.paragraph_formatting.spacing_after
        )
        assert (
            fn_text.paragraph_formatting.line_spacing == en_text.paragraph_formatting.line_spacing
        )
        assert fn_text.run_formatting.font_size == en_text.run_formatting.font_size

    def test_all_reference_styles_are_character_type(self) -> None:
        """Test that all reference styles are character type."""
        fn_ref = get_footnote_reference_style()
        en_ref = get_endnote_reference_style()

        assert fn_ref.style_type == StyleType.CHARACTER
        assert en_ref.style_type == StyleType.CHARACTER

    def test_all_text_styles_are_paragraph_type(self) -> None:
        """Test that all text styles are paragraph type."""
        fn_text = get_footnote_text_style()
        en_text = get_endnote_text_style()

        assert fn_text.style_type == StyleType.PARAGRAPH
        assert en_text.style_type == StyleType.PARAGRAPH

    def test_all_char_styles_are_character_type(self) -> None:
        """Test that all char styles are character type."""
        fn_char = get_footnote_text_char_style()
        en_char = get_endnote_text_char_style()
        hyperlink = get_hyperlink_style()

        assert fn_char.style_type == StyleType.CHARACTER
        assert en_char.style_type == StyleType.CHARACTER
        assert hyperlink.style_type == StyleType.CHARACTER


class TestTocHeadingStyle:
    """Tests for get_toc_heading_style factory function."""

    def test_returns_style_object(self) -> None:
        """Test that factory returns a Style object."""
        style = get_toc_heading_style()
        assert style is not None

    def test_style_id(self) -> None:
        """Test that style_id is TOCHeading."""
        style = get_toc_heading_style()
        assert style.style_id == "TOCHeading"

    def test_name(self) -> None:
        """Test that name is 'TOC Heading'."""
        style = get_toc_heading_style()
        assert style.name == "TOC Heading"

    def test_style_type_is_paragraph(self) -> None:
        """Test that style type is PARAGRAPH."""
        style = get_toc_heading_style()
        assert style.style_type == StyleType.PARAGRAPH

    def test_based_on_heading1(self) -> None:
        """Test that style is based on Heading1."""
        style = get_toc_heading_style()
        assert style.based_on == "Heading1"

    def test_next_style_is_normal(self) -> None:
        """Test that next style is Normal."""
        style = get_toc_heading_style()
        assert style.next_style == "Normal"

    def test_outline_level_excludes_from_toc(self) -> None:
        """Test that outline_level is 9 to exclude from TOC."""
        style = get_toc_heading_style()
        assert style.paragraph_formatting.outline_level == 9

    def test_spacing(self) -> None:
        """Test that spacing before and after are set."""
        style = get_toc_heading_style()
        assert style.paragraph_formatting.spacing_before == 24.0
        assert style.paragraph_formatting.spacing_after == 12.0

    def test_bold_formatting(self) -> None:
        """Test that run formatting includes bold."""
        style = get_toc_heading_style()
        assert style.run_formatting.bold is True

    def test_font_size(self) -> None:
        """Test that font size is 14 points."""
        style = get_toc_heading_style()
        assert style.run_formatting.font_size == 14.0

    def test_ui_priority(self) -> None:
        """Test that ui_priority is 39."""
        style = get_toc_heading_style()
        assert style.ui_priority == 39

    def test_semi_hidden(self) -> None:
        """Test that style is semi-hidden."""
        style = get_toc_heading_style()
        assert style.semi_hidden is True

    def test_unhide_when_used(self) -> None:
        """Test that unhide_when_used is True."""
        style = get_toc_heading_style()
        assert style.unhide_when_used is True


class TestTocLevelStyle:
    """Tests for get_toc_level_style factory function."""

    def test_returns_style_object(self) -> None:
        """Test that factory returns a Style object."""
        style = get_toc_level_style(1)
        assert style is not None

    def test_style_id_toc1(self) -> None:
        """Test that style_id is TOC1 for level 1."""
        style = get_toc_level_style(1)
        assert style.style_id == "TOC1"

    def test_style_id_toc2(self) -> None:
        """Test that style_id is TOC2 for level 2."""
        style = get_toc_level_style(2)
        assert style.style_id == "TOC2"

    def test_style_id_toc3(self) -> None:
        """Test that style_id is TOC3 for level 3."""
        style = get_toc_level_style(3)
        assert style.style_id == "TOC3"

    def test_name_toc1(self) -> None:
        """Test that name is 'toc 1' for level 1."""
        style = get_toc_level_style(1)
        assert style.name == "toc 1"

    def test_style_type_is_paragraph(self) -> None:
        """Test that style type is PARAGRAPH."""
        style = get_toc_level_style(1)
        assert style.style_type == StyleType.PARAGRAPH

    def test_based_on_normal(self) -> None:
        """Test that style is based on Normal."""
        style = get_toc_level_style(1)
        assert style.based_on == "Normal"

    def test_next_style_is_normal(self) -> None:
        """Test that next style is Normal."""
        style = get_toc_level_style(1)
        assert style.next_style == "Normal"

    def test_indent_level1_zero(self) -> None:
        """Test that indent is 0 for level 1."""
        style = get_toc_level_style(1)
        assert style.paragraph_formatting.indent_left == 0.0

    def test_indent_level2(self) -> None:
        """Test that indent is 0.25 inches for level 2."""
        style = get_toc_level_style(2)
        assert style.paragraph_formatting.indent_left == 0.25

    def test_indent_level3(self) -> None:
        """Test that indent is 0.5 inches for level 3."""
        style = get_toc_level_style(3)
        assert style.paragraph_formatting.indent_left == 0.5

    def test_indent_level9(self) -> None:
        """Test that indent is 2.0 inches for level 9."""
        style = get_toc_level_style(9)
        assert style.paragraph_formatting.indent_left == 2.0

    def test_spacing_after(self) -> None:
        """Test that spacing_after is 5.0."""
        style = get_toc_level_style(1)
        assert style.paragraph_formatting.spacing_after == 5.0

    def test_tab_stops_defined(self) -> None:
        """Test that tab stops are defined."""
        style = get_toc_level_style(1)
        assert style.paragraph_formatting.tab_stops is not None
        assert len(style.paragraph_formatting.tab_stops) == 1

    def test_tab_stop_position(self) -> None:
        """Test that tab stop is at 6.5 inches."""
        style = get_toc_level_style(1)
        tab_stop = style.paragraph_formatting.tab_stops[0]
        assert tab_stop.position == 6.5

    def test_tab_stop_alignment(self) -> None:
        """Test that tab stop is right-aligned."""
        style = get_toc_level_style(1)
        tab_stop = style.paragraph_formatting.tab_stops[0]
        assert tab_stop.alignment == "right"

    def test_tab_stop_leader(self) -> None:
        """Test that tab stop has dot leader."""
        style = get_toc_level_style(1)
        tab_stop = style.paragraph_formatting.tab_stops[0]
        assert tab_stop.leader == "dot"

    def test_level1_bold(self) -> None:
        """Test that level 1 is bold."""
        style = get_toc_level_style(1)
        assert style.run_formatting.bold is True

    def test_level2_not_bold(self) -> None:
        """Test that level 2 is not bold."""
        style = get_toc_level_style(2)
        assert style.run_formatting.bold is False

    def test_level3_not_bold(self) -> None:
        """Test that level 3 is not bold."""
        style = get_toc_level_style(3)
        assert style.run_formatting.bold is False

    def test_ui_priority(self) -> None:
        """Test that ui_priority is 39."""
        style = get_toc_level_style(1)
        assert style.ui_priority == 39

    def test_semi_hidden(self) -> None:
        """Test that style is semi-hidden."""
        style = get_toc_level_style(1)
        assert style.semi_hidden is True

    def test_invalid_level_below_1(self) -> None:
        """Test that level below 1 raises ValueError."""
        with pytest.raises(ValueError, match="must be between 1 and 9"):
            get_toc_level_style(0)

    def test_invalid_level_above_9(self) -> None:
        """Test that level above 9 raises ValueError."""
        with pytest.raises(ValueError, match="must be between 1 and 9"):
            get_toc_level_style(10)


class TestStandardStylesDictToc:
    """Tests for TOC styles in STANDARD_STYLES dictionary."""

    def test_contains_toc_heading(self) -> None:
        """Test that TOCHeading is in STANDARD_STYLES."""
        assert "TOCHeading" in STANDARD_STYLES

    def test_contains_toc1_through_toc9(self) -> None:
        """Test that TOC1 through TOC9 are in STANDARD_STYLES."""
        for level in range(1, 10):
            assert f"TOC{level}" in STANDARD_STYLES

    def test_toc_heading_factory_returns_correct_style(self) -> None:
        """Test that TOCHeading factory returns correct style."""
        style = STANDARD_STYLES["TOCHeading"]()
        assert style.style_id == "TOCHeading"

    def test_toc1_factory_returns_correct_style(self) -> None:
        """Test that TOC1 factory returns correct style."""
        style = STANDARD_STYLES["TOC1"]()
        assert style.style_id == "TOC1"

    def test_toc2_factory_returns_correct_style(self) -> None:
        """Test that TOC2 factory returns correct style."""
        style = STANDARD_STYLES["TOC2"]()
        assert style.style_id == "TOC2"

    def test_toc3_factory_returns_correct_style(self) -> None:
        """Test that TOC3 factory returns correct style."""
        style = STANDARD_STYLES["TOC3"]()
        assert style.style_id == "TOC3"


class TestEnsureTocStyles:
    """Tests for the ensure_toc_styles helper function."""

    def test_default_creates_toc_heading_and_toc1_to_3(self) -> None:
        """Test that default creates TOCHeading and TOC1-3."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        ensure_toc_styles(mock_manager)

        # Should add TOCHeading, TOC1, TOC2, TOC3 = 4 styles
        assert mock_manager.add.call_count == 4

        # Verify the style IDs that were added
        added_style_ids = [call[0][0].style_id for call in mock_manager.add.call_args_list]
        assert "TOCHeading" in added_style_ids
        assert "TOC1" in added_style_ids
        assert "TOC2" in added_style_ids
        assert "TOC3" in added_style_ids

    def test_levels_1_creates_toc_heading_and_toc1(self) -> None:
        """Test that levels=1 creates TOCHeading and TOC1 only."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        ensure_toc_styles(mock_manager, levels=1)

        # Should add TOCHeading, TOC1 = 2 styles
        assert mock_manager.add.call_count == 2

    def test_levels_5_creates_toc_heading_and_toc1_to_5(self) -> None:
        """Test that levels=5 creates TOCHeading and TOC1-5."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        ensure_toc_styles(mock_manager, levels=5)

        # Should add TOCHeading, TOC1-5 = 6 styles
        assert mock_manager.add.call_count == 6

    def test_levels_9_creates_all_toc_styles(self) -> None:
        """Test that levels=9 creates TOCHeading and TOC1-9."""
        mock_manager = MagicMock()
        mock_manager.__contains__ = MagicMock(return_value=False)

        ensure_toc_styles(mock_manager, levels=9)

        # Should add TOCHeading, TOC1-9 = 10 styles
        assert mock_manager.add.call_count == 10

    def test_skips_existing_styles(self) -> None:
        """Test that existing styles are not added again."""
        mock_manager = MagicMock()

        # Pretend TOCHeading and TOC1 already exist
        def contains_check(style_id: str) -> bool:
            return style_id in ("TOCHeading", "TOC1")

        mock_manager.__contains__ = MagicMock(side_effect=contains_check)

        ensure_toc_styles(mock_manager, levels=3)

        # Should only add TOC2, TOC3 = 2 styles
        assert mock_manager.add.call_count == 2

    def test_invalid_levels_below_1(self) -> None:
        """Test that levels below 1 raises ValueError."""
        mock_manager = MagicMock()

        with pytest.raises(ValueError, match="must be between 1 and 9"):
            ensure_toc_styles(mock_manager, levels=0)

    def test_invalid_levels_above_9(self) -> None:
        """Test that levels above 9 raises ValueError."""
        mock_manager = MagicMock()

        with pytest.raises(ValueError, match="must be between 1 and 9"):
            ensure_toc_styles(mock_manager, levels=10)
