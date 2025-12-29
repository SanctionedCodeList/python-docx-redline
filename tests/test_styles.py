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
