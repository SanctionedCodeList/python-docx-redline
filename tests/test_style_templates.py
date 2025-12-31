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
    get_endnote_reference_style,
    get_endnote_text_char_style,
    get_endnote_text_style,
    get_footnote_reference_style,
    get_footnote_text_char_style,
    get_footnote_text_style,
    get_hyperlink_style,
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

    def test_contains_all_seven_styles(self) -> None:
        """Test that dictionary contains all 7 standard styles."""
        expected_styles = {
            "FootnoteReference",
            "FootnoteText",
            "FootnoteTextChar",
            "EndnoteReference",
            "EndnoteText",
            "EndnoteTextChar",
            "Hyperlink",
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
