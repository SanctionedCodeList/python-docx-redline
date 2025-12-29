"""Tests for StyleManager class."""

from pathlib import Path

from lxml import etree

from python_docx_redline.models.style import (
    StyleType,
)
from python_docx_redline.package import OOXMLPackage
from python_docx_redline.styles import STYLES_PATH, StyleManager

FIXTURES_DIR = Path(__file__).parent / "fixtures"


class TestStyleManagerLoad:
    """Tests for StyleManager loading from existing documents."""

    def test_load_styles_from_docx(self):
        """Test that StyleManager can load styles from a real document."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Should have loaded some styles
            assert len(style_mgr._styles) > 0

            # Normal style should exist (it's required)
            assert "Normal" in style_mgr._styles
            normal = style_mgr._styles["Normal"]
            assert normal.name == "Normal"
            assert normal.style_type == StyleType.PARAGRAPH

    def test_load_heading_styles(self):
        """Test that heading styles are properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading1 style should exist
            assert "Heading1" in style_mgr._styles
            heading1 = style_mgr._styles["Heading1"]
            assert heading1.name == "heading 1"
            assert heading1.style_type == StyleType.PARAGRAPH
            assert heading1.based_on == "Normal"
            assert heading1.next_style == "Normal"

            # Should have outline level
            assert heading1.paragraph_formatting.outline_level == 0

    def test_load_character_style(self):
        """Test that character styles are properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading1Char should exist as linked character style
            assert "Heading1Char" in style_mgr._styles
            h1_char = style_mgr._styles["Heading1Char"]
            assert h1_char.style_type == StyleType.CHARACTER
            assert h1_char.linked_style == "Heading1"

    def test_load_table_style(self):
        """Test that table styles are properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # TableNormal should exist
            assert "TableNormal" in style_mgr._styles
            table_normal = style_mgr._styles["TableNormal"]
            assert table_normal.style_type == StyleType.TABLE

    def test_load_numbering_style(self):
        """Test that numbering styles are properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # NoList should exist as default numbering style
            assert "NoList" in style_mgr._styles
            no_list = style_mgr._styles["NoList"]
            assert no_list.style_type == StyleType.NUMBERING


class TestStyleManagerParseRunFormatting:
    """Tests for parsing run (character) formatting."""

    def test_parse_bold_style(self):
        """Test that bold formatting is properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading1 style has bold
            heading1 = style_mgr._styles["Heading1"]
            assert heading1.run_formatting.bold is True

    def test_parse_italic_style(self):
        """Test that italic formatting is properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading4 or similar has italic
            heading4 = style_mgr._styles.get("Heading4")
            if heading4:
                assert heading4.run_formatting.italic is True

    def test_parse_font_size(self):
        """Test that font size is properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading1 has font size 28 half-points = 14 points
            heading1 = style_mgr._styles["Heading1"]
            assert heading1.run_formatting.font_size == 14.0

    def test_parse_color(self):
        """Test that text color is properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading1 has a color defined
            heading1 = style_mgr._styles["Heading1"]
            # Color should be formatted as #RRGGBB
            if heading1.run_formatting.color:
                assert heading1.run_formatting.color.startswith("#")


class TestStyleManagerParseParagraphFormatting:
    """Tests for parsing paragraph formatting."""

    def test_parse_alignment(self):
        """Test that alignment is properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Check styles that might have alignment set
            # (The simple_document might not have explicit alignment)
            for style in style_mgr._styles.values():
                if style.paragraph_formatting.alignment:
                    assert style.paragraph_formatting.alignment in [
                        "left",
                        "center",
                        "right",
                        "justify",
                    ]

    def test_parse_keep_properties(self):
        """Test that keep_next and keep_lines are properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading styles typically have keepNext and keepLines
            heading1 = style_mgr._styles["Heading1"]
            assert heading1.paragraph_formatting.keep_next is True
            assert heading1.paragraph_formatting.keep_lines is True

    def test_parse_spacing_before(self):
        """Test that spacing before is properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading1 has spacing before of 480 twips = 24 points
            heading1 = style_mgr._styles["Heading1"]
            assert heading1.paragraph_formatting.spacing_before == 24.0

    def test_parse_outline_level(self):
        """Test that outline level is properly parsed for headings."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading1 = outline level 0, Heading2 = 1, etc.
            heading1 = style_mgr._styles["Heading1"]
            assert heading1.paragraph_formatting.outline_level == 0

            heading2 = style_mgr._styles.get("Heading2")
            if heading2:
                assert heading2.paragraph_formatting.outline_level == 1


class TestStyleManagerUIProperties:
    """Tests for parsing UI properties."""

    def test_parse_quick_format(self):
        """Test that qFormat property is properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Normal style has qFormat
            normal = style_mgr._styles["Normal"]
            assert normal.quick_format is True

    def test_parse_semi_hidden(self):
        """Test that semiHidden property is properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # DefaultParagraphFont is typically semiHidden
            dpf = style_mgr._styles.get("DefaultParagraphFont")
            if dpf:
                assert dpf.semi_hidden is True

    def test_parse_ui_priority(self):
        """Test that uiPriority is properly parsed."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Check that styles have ui_priority values
            heading1 = style_mgr._styles["Heading1"]
            assert heading1.ui_priority == 9


class TestStyleManagerMinimalStyles:
    """Tests for creating minimal styles when none exist."""

    def test_create_minimal_styles(self):
        """Test that minimal styles are created when styles.xml doesn't exist."""

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            # Remove the styles.xml file
            styles_path = package.temp_dir / STYLES_PATH
            if styles_path.exists():
                styles_path.unlink()

            # Create StyleManager - should create minimal styles
            style_mgr = StyleManager(package)

            # Should have Normal style
            assert "Normal" in style_mgr._styles
            normal = style_mgr._styles["Normal"]
            assert normal.style_type == StyleType.PARAGRAPH
            assert normal.quick_format is True

            # Should have DefaultParagraphFont
            assert "DefaultParagraphFont" in style_mgr._styles
            dpf = style_mgr._styles["DefaultParagraphFont"]
            assert dpf.style_type == StyleType.CHARACTER

            # Should be marked as modified
            assert style_mgr.is_modified is True


class TestStyleManagerElementReference:
    """Tests for element reference storage."""

    def test_element_reference_stored(self):
        """Test that Style objects store reference to original XML element."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Each style should have _element reference
            normal = style_mgr._styles["Normal"]
            assert normal._element is not None
            assert isinstance(normal._element, etree._Element)


class TestStyleManagerSave:
    """Tests for saving styles."""

    def test_save_modified_styles(self):
        """Test that modified styles can be saved."""

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            # Remove styles.xml to trigger minimal creation
            styles_path = package.temp_dir / STYLES_PATH
            if styles_path.exists():
                styles_path.unlink()

            style_mgr = StyleManager(package)
            assert style_mgr.is_modified is True

            # Save
            style_mgr.save()

            # Should no longer be modified
            assert style_mgr.is_modified is False

            # File should exist
            assert styles_path.exists()

    def test_save_preserves_existing_unchanged(self):
        """Test that saving when unmodified does not write file."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Should not be marked as modified
            assert style_mgr.is_modified is False

            styles_path = package.temp_dir / STYLES_PATH

            # Save (but nothing modified)
            style_mgr.save()

            # File should still exist and no error raised
            assert styles_path.exists()


class TestStyleManagerGet:
    """Tests for StyleManager.get() method."""

    def test_get_existing_style_by_id(self):
        """Test that get() returns the correct style for an existing style ID."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Get Normal style by ID
            normal = style_mgr.get("Normal")
            assert normal is not None
            assert normal.style_id == "Normal"
            assert normal.name == "Normal"
            assert normal.style_type == StyleType.PARAGRAPH

    def test_get_nonexistent_style_returns_none(self):
        """Test that get() returns None for a non-existent style ID."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            result = style_mgr.get("NonExistentStyle123")
            assert result is None

    def test_get_style_properties_correctly_populated(self):
        """Test that get() returns a Style with correctly populated properties."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Get Heading1 which has many properties set
            heading1 = style_mgr.get("Heading1")
            assert heading1 is not None
            assert heading1.style_id == "Heading1"
            assert heading1.name == "heading 1"
            assert heading1.style_type == StyleType.PARAGRAPH
            assert heading1.based_on == "Normal"
            assert heading1.next_style == "Normal"
            # Verify some formatting properties
            assert heading1.run_formatting.bold is True
            assert heading1.paragraph_formatting.outline_level == 0

    def test_get_character_style(self):
        """Test that get() works for character styles."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Get Heading1Char character style
            h1_char = style_mgr.get("Heading1Char")
            assert h1_char is not None
            assert h1_char.style_type == StyleType.CHARACTER
            assert h1_char.linked_style == "Heading1"


class TestStyleManagerGetByName:
    """Tests for StyleManager.get_by_name() method."""

    def test_get_by_exact_name(self):
        """Test that get_by_name() returns the correct style for an exact name match."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Get by exact name
            normal = style_mgr.get_by_name("Normal")
            assert normal is not None
            assert normal.style_id == "Normal"
            assert normal.name == "Normal"

    def test_get_by_name_case_insensitive(self):
        """Test that get_by_name() is case-insensitive."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Get by different case variations
            normal_lower = style_mgr.get_by_name("normal")
            normal_upper = style_mgr.get_by_name("NORMAL")
            normal_mixed = style_mgr.get_by_name("NoRmAl")

            assert normal_lower is not None
            assert normal_upper is not None
            assert normal_mixed is not None

            # All should return the same style
            assert normal_lower.style_id == "Normal"
            assert normal_upper.style_id == "Normal"
            assert normal_mixed.style_id == "Normal"

    def test_get_by_name_heading_1_case_insensitive(self):
        """Test case-insensitivity with heading style names."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading1 has name "heading 1"
            heading = style_mgr.get_by_name("Heading 1")
            assert heading is not None
            assert heading.style_id == "Heading1"

            heading_upper = style_mgr.get_by_name("HEADING 1")
            assert heading_upper is not None
            assert heading_upper.style_id == "Heading1"

    def test_get_by_name_nonexistent_returns_none(self):
        """Test that get_by_name() returns None for a non-existent name."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            result = style_mgr.get_by_name("Non Existent Style Name")
            assert result is None


class TestStyleManagerList:
    """Tests for StyleManager.list() method."""

    def test_list_all_styles(self):
        """Test that list() returns all non-hidden styles."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            all_styles = style_mgr.list()
            assert len(all_styles) > 0

            # All returned styles should be Style objects
            for style in all_styles:
                assert hasattr(style, "style_id")
                assert hasattr(style, "name")
                assert hasattr(style, "style_type")

    def test_list_filter_by_paragraph_type(self):
        """Test that list() can filter by StyleType.PARAGRAPH."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            para_styles = style_mgr.list(style_type=StyleType.PARAGRAPH)

            assert len(para_styles) > 0
            for style in para_styles:
                assert style.style_type == StyleType.PARAGRAPH

            # Normal should be in the list
            style_ids = [s.style_id for s in para_styles]
            assert "Normal" in style_ids

    def test_list_filter_by_character_type(self):
        """Test that list() can filter by StyleType.CHARACTER."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            char_styles = style_mgr.list(style_type=StyleType.CHARACTER)

            # All returned styles should be character type
            for style in char_styles:
                assert style.style_type == StyleType.CHARACTER

    def test_list_include_hidden_false_excludes_semi_hidden(self):
        """Test that include_hidden=False excludes semi_hidden styles."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Get visible styles (default behavior)
            visible_styles = style_mgr.list(include_hidden=False)

            # None should be semi_hidden
            for style in visible_styles:
                assert style.semi_hidden is False

    def test_list_include_hidden_true_includes_all_styles(self):
        """Test that include_hidden=True includes semi_hidden styles."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Get all styles including hidden
            all_styles = style_mgr.list(include_hidden=True)

            # Should include some semi_hidden styles
            semi_hidden_count = sum(1 for s in all_styles if s.semi_hidden)

            # DefaultParagraphFont is typically semi_hidden
            assert semi_hidden_count > 0

            # All styles should have more than visible-only
            visible_styles = style_mgr.list(include_hidden=False)
            assert len(all_styles) > len(visible_styles)

    def test_list_combined_filters(self):
        """Test that list() works with both type filter and include_hidden."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Get all character styles including hidden
            all_char_styles = style_mgr.list(style_type=StyleType.CHARACTER, include_hidden=True)

            # Get only visible character styles
            visible_char_styles = style_mgr.list(
                style_type=StyleType.CHARACTER, include_hidden=False
            )

            # All should be CHARACTER type
            for style in all_char_styles:
                assert style.style_type == StyleType.CHARACTER

            # Hidden list should be >= visible list
            assert len(all_char_styles) >= len(visible_char_styles)


class TestStyleManagerContains:
    """Tests for StyleManager.__contains__() method."""

    def test_contains_existing_style_returns_true(self):
        """Test that 'in' returns True for existing style ID."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            assert "Normal" in style_mgr
            assert "Heading1" in style_mgr

    def test_contains_nonexistent_style_returns_false(self):
        """Test that 'in' returns False for non-existent style ID."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            assert "NonExistentStyle123" not in style_mgr
            assert "FakeStyleName" not in style_mgr

    def test_contains_case_sensitive(self):
        """Test that __contains__ is case-sensitive (matches style ID exactly)."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Style ID is case-sensitive
            assert "Normal" in style_mgr
            assert "normal" not in style_mgr
            assert "NORMAL" not in style_mgr


class TestStyleManagerIter:
    """Tests for StyleManager.__iter__() method."""

    def test_iter_yields_all_styles(self):
        """Test that iteration yields all styles."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            styles_via_iter = list(style_mgr)
            expected_count = len(style_mgr._styles)

            assert len(styles_via_iter) == expected_count

    def test_iter_used_in_for_loop(self):
        """Test that iteration works in a for loop."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style_ids = []
            for style in style_mgr:
                assert hasattr(style, "style_id")
                style_ids.append(style.style_id)

            assert "Normal" in style_ids
            assert len(style_ids) > 0

    def test_iter_yields_style_objects(self):
        """Test that iteration yields Style objects with proper attributes."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            for style in style_mgr:
                # Each yielded item should be a Style object
                assert hasattr(style, "style_id")
                assert hasattr(style, "name")
                assert hasattr(style, "style_type")
                assert hasattr(style, "run_formatting")
                assert hasattr(style, "paragraph_formatting")


class TestStyleManagerLen:
    """Tests for StyleManager.__len__() method."""

    def test_len_returns_correct_count(self):
        """Test that len() returns the correct count of styles."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            count = len(style_mgr)
            assert count > 0
            assert count == len(style_mgr._styles)

    def test_len_matches_iteration_count(self):
        """Test that len() matches the count from iteration."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            len_count = len(style_mgr)
            iter_count = sum(1 for _ in style_mgr)

            assert len_count == iter_count

    def test_len_minimal_styles(self):
        """Test len() with minimal styles (when styles.xml doesn't exist)."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            # Remove styles.xml to trigger minimal creation
            styles_path = package.temp_dir / STYLES_PATH
            if styles_path.exists():
                styles_path.unlink()

            style_mgr = StyleManager(package)

            # Should have at least Normal and DefaultParagraphFont
            assert len(style_mgr) >= 2


class TestStyleManagerWithRealDocument:
    """Integration tests with real document to verify expected styles."""

    def test_common_styles_exist(self):
        """Test that common styles like Normal exist in real document."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Normal style should exist
            normal = style_mgr.get("Normal")
            assert normal is not None
            assert normal.name == "Normal"
            assert normal.style_type == StyleType.PARAGRAPH

    def test_heading_styles_hierarchy(self):
        """Test that heading styles have correct hierarchy (based_on Normal)."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            heading1 = style_mgr.get("Heading1")
            if heading1:
                assert heading1.based_on == "Normal"
                assert heading1.next_style == "Normal"
                assert heading1.paragraph_formatting.outline_level == 0

    def test_style_properties_match_expected_values(self):
        """Test that style properties match expected values from document."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Verify Normal style
            normal = style_mgr.get("Normal")
            assert normal is not None
            assert normal.quick_format is True

            # Verify Heading1 style
            heading1 = style_mgr.get("Heading1")
            assert heading1 is not None
            assert heading1.ui_priority == 9
            assert heading1.run_formatting.bold is True

    def test_linked_styles_are_consistent(self):
        """Test that linked styles reference each other correctly."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Heading1 should have a linked character style
            heading1 = style_mgr.get("Heading1")
            if heading1 and heading1.linked_style:
                linked_char = style_mgr.get(heading1.linked_style)
                assert linked_char is not None
                assert linked_char.style_type == StyleType.CHARACTER
                # The character style should link back
                assert linked_char.linked_style == "Heading1"

    def test_table_and_numbering_styles_exist(self):
        """Test that table and numbering styles exist in real document."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # TableNormal should exist
            table_normal = style_mgr.get("TableNormal")
            assert table_normal is not None
            assert table_normal.style_type == StyleType.TABLE

            # NoList should exist
            no_list = style_mgr.get("NoList")
            assert no_list is not None
            assert no_list.style_type == StyleType.NUMBERING

    def test_list_filters_work_with_real_document(self):
        """Test that list() filters work correctly with real document."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Get counts by type
            para_count = len(style_mgr.list(style_type=StyleType.PARAGRAPH))
            char_count = len(style_mgr.list(style_type=StyleType.CHARACTER))
            table_count = len(style_mgr.list(style_type=StyleType.TABLE, include_hidden=True))
            num_count = len(style_mgr.list(style_type=StyleType.NUMBERING, include_hidden=True))

            # Should have at least some of each type
            assert para_count > 0
            assert char_count >= 0  # May be 0 if all are hidden
            assert table_count > 0
            assert num_count > 0

            # Total from filters should equal total (with include_hidden=True)
            all_styles = style_mgr.list(include_hidden=True)
            sum_by_type = sum(
                len(style_mgr.list(style_type=t, include_hidden=True)) for t in StyleType
            )
            assert sum_by_type == len(all_styles)


class TestStyleManagerAdd:
    """Tests for adding new styles."""

    def test_add_character_style(self):
        """Test adding a simple character style."""
        from python_docx_redline.models.style import RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Create a custom style
            custom = Style(
                style_id="MyHighlight",
                name="My Highlight",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(bold=True, color="#FF0000"),
            )

            style_mgr.add(custom)

            # Style should now exist in the manager
            assert "MyHighlight" in style_mgr
            retrieved = style_mgr.get("MyHighlight")
            assert retrieved is not None
            assert retrieved.name == "My Highlight"
            assert retrieved.style_type == StyleType.CHARACTER

            # Manager should be marked as modified
            assert style_mgr.is_modified is True

    def test_add_paragraph_style(self):
        """Test adding a paragraph style with formatting."""
        from python_docx_redline.models.style import ParagraphFormatting, RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            custom = Style(
                style_id="CustomParagraph",
                name="Custom Paragraph",
                style_type=StyleType.PARAGRAPH,
                based_on="Normal",
                paragraph_formatting=ParagraphFormatting(
                    alignment="justify",
                    spacing_after=12.0,
                    line_spacing=1.5,
                ),
                run_formatting=RunFormatting(font_name="Arial", font_size=11.0),
            )

            style_mgr.add(custom)

            assert "CustomParagraph" in style_mgr
            retrieved = style_mgr.get("CustomParagraph")
            assert retrieved.based_on == "Normal"

    def test_add_duplicate_raises_error(self):
        """Test that adding a duplicate style_id raises ValueError."""
        import pytest

        from python_docx_redline.models.style import Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # "Normal" already exists
            duplicate = Style(
                style_id="Normal",
                name="Duplicate Normal",
                style_type=StyleType.PARAGRAPH,
            )

            with pytest.raises(ValueError, match="already exists"):
                style_mgr.add(duplicate)

    def test_add_style_with_all_properties(self):
        """Test adding a style with all properties set."""
        from python_docx_redline.models.style import ParagraphFormatting, RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            custom = Style(
                style_id="FullFeatured",
                name="Full Featured Style",
                style_type=StyleType.PARAGRAPH,
                based_on="Normal",
                next_style="Normal",
                linked_style="FullFeaturedChar",
                ui_priority=50,
                quick_format=True,
                semi_hidden=False,
                unhide_when_used=True,
                run_formatting=RunFormatting(
                    bold=True,
                    italic=True,
                    underline="double",
                    font_name="Times New Roman",
                    font_size=14.0,
                    color="#0000FF",
                    superscript=True,
                ),
                paragraph_formatting=ParagraphFormatting(
                    alignment="center",
                    spacing_before=6.0,
                    spacing_after=12.0,
                    line_spacing=2.0,
                    indent_left=0.5,
                    keep_next=True,
                    keep_lines=True,
                    outline_level=1,
                ),
            )

            style_mgr.add(custom)

            assert "FullFeatured" in style_mgr


class TestStyleManagerEnsureStyle:
    """Tests for the ensure_style method."""

    def test_ensure_style_returns_existing(self):
        """Test that ensure_style returns existing style if it exists."""
        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Normal already exists
            style = style_mgr.ensure_style(
                style_id="Normal",
                name="should be ignored",
                style_type=StyleType.CHARACTER,  # Wrong type, but should be ignored
            )

            # Should return the existing style
            assert style.style_id == "Normal"
            assert style.style_type == StyleType.PARAGRAPH  # Original type

            # Should NOT be modified since we just returned existing
            assert style_mgr.is_modified is False

    def test_ensure_style_creates_if_missing(self):
        """Test that ensure_style creates a new style if missing."""
        from python_docx_redline.models.style import RunFormatting

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # FootnoteReference typically doesn't exist in simple_document
            style = style_mgr.ensure_style(
                style_id="FootnoteReference",
                name="footnote reference",
                style_type=StyleType.CHARACTER,
                based_on="DefaultParagraphFont",
                run_formatting=RunFormatting(superscript=True),
                ui_priority=99,
                unhide_when_used=True,
            )

            # Should be created
            assert style.style_id == "FootnoteReference"
            assert style.name == "footnote reference"
            assert style.style_type == StyleType.CHARACTER
            assert style.based_on == "DefaultParagraphFont"
            assert style.run_formatting.superscript is True
            assert style.ui_priority == 99
            assert style.unhide_when_used is True

            # Should be in manager
            assert "FootnoteReference" in style_mgr
            assert style_mgr.is_modified is True

    def test_ensure_style_idempotent(self):
        """Test that calling ensure_style twice returns same style."""
        from python_docx_redline.models.style import RunFormatting

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # First call - creates it
            style1 = style_mgr.ensure_style(
                style_id="NewStyle",
                name="New Style",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(bold=True),
            )

            # Second call - returns existing
            style2 = style_mgr.ensure_style(
                style_id="NewStyle",
                name="Different Name",  # Should be ignored
                style_type=StyleType.PARAGRAPH,  # Should be ignored
            )

            # Should be the same style
            assert style1.style_id == style2.style_id
            assert style2.name == "New Style"  # Original name
            assert style2.style_type == StyleType.CHARACTER  # Original type


class TestStyleManagerStyleToElement:
    """Tests for converting Style to XML element."""

    def test_style_to_element_basic(self):
        """Test converting a basic style to XML."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="TestStyle",
                name="Test Style",
                style_type=StyleType.CHARACTER,
            )

            elem = style_mgr._style_to_element(style)

            # Check element attributes
            assert elem.tag == w("style")
            assert elem.get(w("type")) == "character"
            assert elem.get(w("styleId")) == "TestStyle"

            # Check name child
            name_elem = elem.find(w("name"))
            assert name_elem is not None
            assert name_elem.get(w("val")) == "Test Style"

    def test_style_to_element_with_formatting(self):
        """Test converting a style with run formatting to XML."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="FormattedStyle",
                name="Formatted Style",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(
                    bold=True,
                    italic=True,
                    font_size=14.0,
                    superscript=True,
                ),
            )

            elem = style_mgr._style_to_element(style)

            # Check rPr exists
            rpr = elem.find(w("rPr"))
            assert rpr is not None

            # Check bold
            bold = rpr.find(w("b"))
            assert bold is not None
            assert bold.get(w("val")) is None  # True is implicit

            # Check italic
            italic = rpr.find(w("i"))
            assert italic is not None

            # Check font size (14pt = 28 half-points)
            sz = rpr.find(w("sz"))
            assert sz is not None
            assert sz.get(w("val")) == "28"

            # Check superscript
            vert_align = rpr.find(w("vertAlign"))
            assert vert_align is not None
            assert vert_align.get(w("val")) == "superscript"


class TestStyleManagerStyleToElementDetailed:
    """Detailed tests for _style_to_element() XML generation."""

    def test_style_to_element_with_based_on(self):
        """Test that basedOn element is created correctly."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="CustomStyle",
                name="Custom Style",
                style_type=StyleType.CHARACTER,
                based_on="DefaultParagraphFont",
            )

            elem = style_mgr._style_to_element(style)

            # Check basedOn element exists with correct value
            based_on_elem = elem.find(w("basedOn"))
            assert based_on_elem is not None
            assert based_on_elem.get(w("val")) == "DefaultParagraphFont"

    def test_style_to_element_superscript_creates_vert_align(self):
        """Test that superscript creates w:vertAlign element."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="SuperscriptStyle",
                name="Superscript Style",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(superscript=True),
            )

            elem = style_mgr._style_to_element(style)

            # Check w:rPr/w:vertAlign exists with superscript value
            rpr = elem.find(w("rPr"))
            assert rpr is not None
            vert_align = rpr.find(w("vertAlign"))
            assert vert_align is not None
            assert vert_align.get(w("val")) == "superscript"

    def test_style_to_element_subscript_creates_vert_align(self):
        """Test that subscript creates w:vertAlign element."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="SubscriptStyle",
                name="Subscript Style",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(subscript=True),
            )

            elem = style_mgr._style_to_element(style)

            rpr = elem.find(w("rPr"))
            assert rpr is not None
            vert_align = rpr.find(w("vertAlign"))
            assert vert_align is not None
            assert vert_align.get(w("val")) == "subscript"

    def test_style_to_element_alignment_creates_jc(self):
        """Test that alignment creates w:pPr/w:jc element."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import ParagraphFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="CenteredStyle",
                name="Centered Style",
                style_type=StyleType.PARAGRAPH,
                paragraph_formatting=ParagraphFormatting(alignment="center"),
            )

            elem = style_mgr._style_to_element(style)

            # Check w:pPr/w:jc exists with center value
            ppr = elem.find(w("pPr"))
            assert ppr is not None
            jc = ppr.find(w("jc"))
            assert jc is not None
            assert jc.get(w("val")) == "center"

    def test_style_to_element_justify_alignment_maps_to_both(self):
        """Test that 'justify' alignment maps to OOXML 'both' value."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import ParagraphFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="JustifiedStyle",
                name="Justified Style",
                style_type=StyleType.PARAGRAPH,
                paragraph_formatting=ParagraphFormatting(alignment="justify"),
            )

            elem = style_mgr._style_to_element(style)

            ppr = elem.find(w("pPr"))
            jc = ppr.find(w("jc"))
            assert jc is not None
            # "justify" should map to "both" in OOXML
            assert jc.get(w("val")) == "both"

    def test_style_to_element_ui_properties(self):
        """Test that UI properties are correctly generated."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="UITestStyle",
                name="UI Test Style",
                style_type=StyleType.CHARACTER,
                ui_priority=99,
                quick_format=True,
                semi_hidden=True,
                unhide_when_used=True,
            )

            elem = style_mgr._style_to_element(style)

            # Check uiPriority
            ui_priority = elem.find(w("uiPriority"))
            assert ui_priority is not None
            assert ui_priority.get(w("val")) == "99"

            # Check qFormat (presence indicates True)
            assert elem.find(w("qFormat")) is not None

            # Check semiHidden
            assert elem.find(w("semiHidden")) is not None

            # Check unhideWhenUsed
            assert elem.find(w("unhideWhenUsed")) is not None

    def test_style_to_element_next_and_linked_styles(self):
        """Test that next_style and linked_style are generated correctly."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="HeadingTest",
                name="Heading Test",
                style_type=StyleType.PARAGRAPH,
                next_style="Normal",
                linked_style="HeadingTestChar",
            )

            elem = style_mgr._style_to_element(style)

            # Check next element
            next_elem = elem.find(w("next"))
            assert next_elem is not None
            assert next_elem.get(w("val")) == "Normal"

            # Check link element
            link_elem = elem.find(w("link"))
            assert link_elem is not None
            assert link_elem.get(w("val")) == "HeadingTestChar"

    def test_style_to_element_run_formatting_comprehensive(self):
        """Test comprehensive run formatting XML generation."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="ComprehensiveRun",
                name="Comprehensive Run Style",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(
                    bold=True,
                    italic=True,
                    underline="double",
                    strikethrough=True,
                    font_name="Arial",
                    font_size=12.0,
                    color="#FF0000",
                    highlight="yellow",
                    small_caps=True,
                ),
            )

            elem = style_mgr._style_to_element(style)
            rpr = elem.find(w("rPr"))
            assert rpr is not None

            # Check bold (no val attribute means True)
            assert rpr.find(w("b")) is not None

            # Check italic
            assert rpr.find(w("i")) is not None

            # Check underline
            u = rpr.find(w("u"))
            assert u is not None
            assert u.get(w("val")) == "double"

            # Check strikethrough
            assert rpr.find(w("strike")) is not None

            # Check font
            fonts = rpr.find(w("rFonts"))
            assert fonts is not None
            assert fonts.get(w("ascii")) == "Arial"

            # Check font size (12pt = 24 half-points)
            sz = rpr.find(w("sz"))
            assert sz is not None
            assert sz.get(w("val")) == "24"

            # Check color (without # prefix)
            color = rpr.find(w("color"))
            assert color is not None
            assert color.get(w("val")) == "FF0000"

            # Check highlight
            highlight = rpr.find(w("highlight"))
            assert highlight is not None
            assert highlight.get(w("val")) == "yellow"

            # Check small caps
            assert rpr.find(w("smallCaps")) is not None

    def test_style_to_element_paragraph_formatting_comprehensive(self):
        """Test comprehensive paragraph formatting XML generation."""
        from python_docx_redline.constants import w
        from python_docx_redline.models.style import ParagraphFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = Style(
                style_id="ComprehensivePara",
                name="Comprehensive Para Style",
                style_type=StyleType.PARAGRAPH,
                paragraph_formatting=ParagraphFormatting(
                    alignment="right",
                    spacing_before=12.0,
                    spacing_after=6.0,
                    line_spacing=1.5,
                    indent_left=0.5,
                    indent_right=0.25,
                    indent_first_line=0.5,
                    keep_next=True,
                    keep_lines=True,
                    outline_level=2,
                ),
            )

            elem = style_mgr._style_to_element(style)
            ppr = elem.find(w("pPr"))
            assert ppr is not None

            # Check alignment
            jc = ppr.find(w("jc"))
            assert jc is not None
            assert jc.get(w("val")) == "right"

            # Check spacing (12pt before = 240 twips, 6pt after = 120 twips)
            spacing = ppr.find(w("spacing"))
            assert spacing is not None
            assert spacing.get(w("before")) == "240"
            assert spacing.get(w("after")) == "120"
            # Line spacing 1.5 = 360 (in 240ths)
            assert spacing.get(w("line")) == "360"
            assert spacing.get(w("lineRule")) == "auto"

            # Check indentation (0.5 inch = 720 twips)
            ind = ppr.find(w("ind"))
            assert ind is not None
            assert ind.get(w("left")) == "720"
            assert ind.get(w("right")) == "360"  # 0.25 inch = 360 twips
            assert ind.get(w("firstLine")) == "720"

            # Check keep properties
            assert ppr.find(w("keepNext")) is not None
            assert ppr.find(w("keepLines")) is not None

            # Check outline level
            outline = ppr.find(w("outlineLvl"))
            assert outline is not None
            assert outline.get(w("val")) == "2"


class TestStyleManagerAddDetailed:
    """Detailed tests for add() method."""

    def test_add_sets_modified_flag(self):
        """Test that add() sets _modified to True."""
        from python_docx_redline.models.style import Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Initially should not be modified (loaded from existing file)
            assert style_mgr.is_modified is False

            custom = Style(
                style_id="ModifiedTest",
                name="Modified Test",
                style_type=StyleType.CHARACTER,
            )
            style_mgr.add(custom)

            # Should now be modified
            assert style_mgr.is_modified is True

    def test_add_style_accessible_via_get(self):
        """Test that added style is accessible via get() immediately."""
        from python_docx_redline.models.style import RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            custom = Style(
                style_id="GetTestStyle",
                name="Get Test Style",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(italic=True),
            )
            style_mgr.add(custom)

            # Should be retrievable immediately
            retrieved = style_mgr.get("GetTestStyle")
            assert retrieved is not None
            assert retrieved.name == "Get Test Style"
            assert retrieved.run_formatting.italic is True

    def test_add_style_stores_element_reference(self):
        """Test that added style has _element reference set."""
        from python_docx_redline.models.style import Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            custom = Style(
                style_id="ElementRefTest",
                name="Element Ref Test",
                style_type=StyleType.CHARACTER,
            )
            style_mgr.add(custom)

            # The style should have _element set
            assert custom._element is not None
            assert isinstance(custom._element, etree._Element)


class TestStyleManagerSaveDetailed:
    """Detailed tests for save() method."""

    def test_save_sets_modified_to_false(self):
        """Test that save() sets _modified to False."""
        from python_docx_redline.models.style import Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Add a style to mark as modified
            custom = Style(
                style_id="SaveModifiedTest",
                name="Save Modified Test",
                style_type=StyleType.CHARACTER,
            )
            style_mgr.add(custom)
            assert style_mgr.is_modified is True

            # Save
            style_mgr.save()

            # Should no longer be modified
            assert style_mgr.is_modified is False

    def test_save_writes_changes_to_file(self):
        """Test that save() writes changes to the styles.xml file."""
        from python_docx_redline.models.style import RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            custom = Style(
                style_id="SaveWriteTest",
                name="Save Write Test",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(bold=True),
            )
            style_mgr.add(custom)
            style_mgr.save()

            # Read the saved file directly
            styles_path = package.temp_dir / STYLES_PATH
            assert styles_path.exists()

            # Parse and verify the style was written
            tree = etree.parse(str(styles_path))
            root = tree.getroot()

            from python_docx_redline.constants import w

            # Find our style
            found = False
            for style_elem in root.findall(w("style")):
                if style_elem.get(w("styleId")) == "SaveWriteTest":
                    found = True
                    # Verify it has the expected content
                    name_elem = style_elem.find(w("name"))
                    assert name_elem is not None
                    assert name_elem.get(w("val")) == "Save Write Test"
                    break
            assert found, "Style 'SaveWriteTest' not found in saved file"

    def test_saved_styles_persist_after_reload(self):
        """Test that saved styles persist after creating new StyleManager."""
        from python_docx_redline.models.style import RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            custom = Style(
                style_id="PersistTest",
                name="Persist Test",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(italic=True, color="#00FF00"),
            )
            style_mgr.add(custom)
            style_mgr.save()

            # Create new StyleManager to reload
            style_mgr2 = StyleManager(package)

            # Style should be present
            assert "PersistTest" in style_mgr2
            reloaded = style_mgr2.get("PersistTest")
            assert reloaded is not None
            assert reloaded.name == "Persist Test"
            assert reloaded.style_type == StyleType.CHARACTER
            assert reloaded.run_formatting.italic is True
            assert reloaded.run_formatting.color == "#00FF00"


class TestStyleManagerEnsureStyleDetailed:
    """Detailed tests for ensure_style() method."""

    def test_ensure_style_with_run_formatting(self):
        """Test ensure_style with run_formatting parameter."""
        from python_docx_redline.models.style import RunFormatting

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = style_mgr.ensure_style(
                style_id="RunFormattedStyle",
                name="Run Formatted Style",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(bold=True, font_size=16.0),
            )

            assert style.run_formatting.bold is True
            assert style.run_formatting.font_size == 16.0

    def test_ensure_style_with_paragraph_formatting(self):
        """Test ensure_style with paragraph_formatting parameter."""
        from python_docx_redline.models.style import ParagraphFormatting

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            style = style_mgr.ensure_style(
                style_id="ParaFormattedStyle",
                name="Para Formatted Style",
                style_type=StyleType.PARAGRAPH,
                paragraph_formatting=ParagraphFormatting(
                    alignment="center",
                    spacing_after=10.0,
                ),
            )

            assert style.paragraph_formatting.alignment == "center"
            assert style.paragraph_formatting.spacing_after == 10.0


class TestStyleManagerRoundTrip:
    """Tests for round-trip style preservation."""

    def test_add_save_reload_preserves_style(self):
        """Test that adding, saving, and reloading preserves the style."""
        from python_docx_redline.models.style import RunFormatting, Style

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Add a custom style
            custom = Style(
                style_id="RoundTripTest",
                name="Round Trip Test",
                style_type=StyleType.CHARACTER,
                based_on="DefaultParagraphFont",
                run_formatting=RunFormatting(bold=True, color="#FF5500"),
                ui_priority=75,
                semi_hidden=True,
                unhide_when_used=True,
            )
            style_mgr.add(custom)

            # Save
            style_mgr.save()

            # Create new StyleManager to reload
            style_mgr2 = StyleManager(package)

            # Check the style was preserved
            assert "RoundTripTest" in style_mgr2
            reloaded = style_mgr2.get("RoundTripTest")

            assert reloaded.name == "Round Trip Test"
            assert reloaded.style_type == StyleType.CHARACTER
            assert reloaded.based_on == "DefaultParagraphFont"
            assert reloaded.run_formatting.bold is True
            assert reloaded.run_formatting.color == "#FF5500"
            assert reloaded.ui_priority == 75
            assert reloaded.semi_hidden is True
            assert reloaded.unhide_when_used is True

    def test_footnote_reference_style_round_trip(self):
        """Test creating a FootnoteReference style as used in notes.py."""
        from python_docx_redline.models.style import RunFormatting

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # This mirrors _ensure_footnote_styles in notes.py
            _style = style_mgr.ensure_style(
                style_id="FootnoteReference",
                name="footnote reference",
                style_type=StyleType.CHARACTER,
                based_on="DefaultParagraphFont",
                run_formatting=RunFormatting(superscript=True),
                ui_priority=99,
                unhide_when_used=True,
            )

            # Save and reload
            style_mgr.save()
            style_mgr2 = StyleManager(package)

            # Verify
            reloaded = style_mgr2.get("FootnoteReference")
            assert reloaded is not None
            assert reloaded.run_formatting.superscript is True
            assert reloaded.ui_priority == 99

    def test_footnote_text_paragraph_style_round_trip(self):
        """Test creating a FootnoteText-like paragraph style round trip."""
        from python_docx_redline.models.style import ParagraphFormatting, RunFormatting

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Create a FootnoteText-like paragraph style
            _style = style_mgr.ensure_style(
                style_id="FootnoteText",
                name="footnote text",
                style_type=StyleType.PARAGRAPH,
                based_on="Normal",
                next_style="FootnoteText",
                run_formatting=RunFormatting(font_size=10.0),
                paragraph_formatting=ParagraphFormatting(
                    spacing_after=0,
                    line_spacing=1.0,
                ),
                ui_priority=99,
                semi_hidden=True,
                unhide_when_used=True,
            )

            # Save and reload
            style_mgr.save()
            style_mgr2 = StyleManager(package)

            # Verify
            reloaded = style_mgr2.get("FootnoteText")
            assert reloaded is not None
            assert reloaded.name == "footnote text"
            assert reloaded.style_type == StyleType.PARAGRAPH
            assert reloaded.based_on == "Normal"
            assert reloaded.next_style == "FootnoteText"
            assert reloaded.run_formatting.font_size == 10.0
            assert reloaded.paragraph_formatting.spacing_after == 0
            assert reloaded.paragraph_formatting.line_spacing == 1.0
            assert reloaded.ui_priority == 99
            assert reloaded.semi_hidden is True
            assert reloaded.unhide_when_used is True

    def test_multiple_styles_round_trip(self):
        """Test creating multiple styles and saving them together."""
        from python_docx_redline.models.style import (
            ParagraphFormatting,
            RunFormatting,
            Style,
        )

        docx_path = FIXTURES_DIR / "simple_document.docx"
        with OOXMLPackage.open(docx_path) as package:
            style_mgr = StyleManager(package)

            # Add multiple styles
            style1 = Style(
                style_id="CustomTitle",
                name="Custom Title",
                style_type=StyleType.PARAGRAPH,
                run_formatting=RunFormatting(bold=True, font_size=24.0),
                paragraph_formatting=ParagraphFormatting(alignment="center"),
            )
            style_mgr.add(style1)

            style2 = Style(
                style_id="CustomEmphasis",
                name="Custom Emphasis",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(italic=True, color="#0000FF"),
            )
            style_mgr.add(style2)

            style3 = Style(
                style_id="CustomCode",
                name="Custom Code",
                style_type=StyleType.CHARACTER,
                run_formatting=RunFormatting(font_name="Courier New", font_size=10.0),
            )
            style_mgr.add(style3)

            # Save
            style_mgr.save()

            # Reload
            style_mgr2 = StyleManager(package)

            # Verify all styles exist and have correct properties
            title = style_mgr2.get("CustomTitle")
            assert title is not None
            assert title.run_formatting.bold is True
            assert title.run_formatting.font_size == 24.0
            assert title.paragraph_formatting.alignment == "center"

            emphasis = style_mgr2.get("CustomEmphasis")
            assert emphasis is not None
            assert emphasis.run_formatting.italic is True
            assert emphasis.run_formatting.color == "#0000FF"

            code = style_mgr2.get("CustomCode")
            assert code is not None
            assert code.run_formatting.font_name == "Courier New"
            assert code.run_formatting.font_size == 10.0
