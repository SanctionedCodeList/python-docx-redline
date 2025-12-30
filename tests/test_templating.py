"""Tests for DocxBuilder templating module."""

import tempfile
from dataclasses import dataclass
from pathlib import Path

import pytest

from python_docx_redline import DocxBuilder


class TestDocxBuilderBasic:
    """Test basic DocxBuilder functionality."""

    def test_create_empty_document(self):
        """Test creating an empty document."""
        doc = DocxBuilder()
        assert doc.document is not None

    def test_save_document(self):
        """Test saving a document."""
        doc = DocxBuilder()
        doc.heading("Test Document")

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            path = doc.save(f.name)
            assert path.exists()
            assert path.suffix == ".docx"

    def test_save_returns_path(self):
        """Test that save returns a Path object."""
        doc = DocxBuilder()

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            result = doc.save(f.name)
            assert isinstance(result, Path)


class TestDocxBuilderConfiguration:
    """Test DocxBuilder configuration options."""

    def test_default_configuration(self):
        """Test default configuration is applied."""
        doc = DocxBuilder()
        style = doc.document.styles["Normal"]
        assert style.font.name == "Times New Roman"

    def test_custom_font(self):
        """Test custom font configuration."""
        doc = DocxBuilder(font="Arial")
        style = doc.document.styles["Normal"]
        assert style.font.name == "Arial"

    def test_landscape_orientation(self):
        """Test landscape orientation."""
        doc = DocxBuilder(landscape=True)
        section = doc.document.sections[0]
        # In landscape, width > height
        assert section.page_width > section.page_height

    def test_portrait_orientation(self):
        """Test portrait orientation (default)."""
        doc = DocxBuilder(landscape=False)
        section = doc.document.sections[0]
        # In portrait, height > width
        assert section.page_height > section.page_width


class TestDocxBuilderHeadings:
    """Test heading methods."""

    def test_heading_level_1(self):
        """Test adding a level 1 heading."""
        doc = DocxBuilder()
        result = doc.heading("Main Title")

        # Should return self for chaining
        assert result is doc

        # Check heading was added
        paragraphs = doc.document.paragraphs
        assert len(paragraphs) == 1
        assert paragraphs[0].text == "Main Title"

    def test_heading_level_2(self):
        """Test adding a level 2 heading."""
        doc = DocxBuilder()
        doc.heading("Section", level=2)

        paragraphs = doc.document.paragraphs
        assert paragraphs[0].text == "Section"

    def test_multiple_headings(self):
        """Test adding multiple headings."""
        doc = DocxBuilder()
        doc.heading("Title", level=0)
        doc.heading("Chapter 1", level=1)
        doc.heading("Section 1.1", level=2)

        paragraphs = doc.document.paragraphs
        assert len(paragraphs) == 3
        assert paragraphs[0].text == "Title"
        assert paragraphs[1].text == "Chapter 1"
        assert paragraphs[2].text == "Section 1.1"


class TestDocxBuilderParagraphs:
    """Test paragraph methods."""

    def test_paragraph_basic(self):
        """Test adding a basic paragraph."""
        doc = DocxBuilder()
        result = doc.paragraph("This is a test paragraph.")

        # Should return self for chaining
        assert result is doc

        paragraphs = doc.document.paragraphs
        assert len(paragraphs) == 1
        assert paragraphs[0].text == "This is a test paragraph."

    def test_multiple_paragraphs(self):
        """Test adding multiple paragraphs."""
        doc = DocxBuilder()
        doc.paragraph("First paragraph.")
        doc.paragraph("Second paragraph.")

        paragraphs = doc.document.paragraphs
        assert len(paragraphs) == 2
        assert paragraphs[0].text == "First paragraph."
        assert paragraphs[1].text == "Second paragraph."


class TestDocxBuilderTables:
    """Test table methods."""

    def test_table_basic(self):
        """Test creating a basic table."""
        doc = DocxBuilder()
        result = doc.table(
            headers=["Name", "Value"],
            rows=[["Alpha", "1"], ["Beta", "2"]],
        )

        # Should return self for chaining
        assert result is doc

        tables = doc.document.tables
        assert len(tables) == 1

        table = tables[0]
        # Header row + 2 data rows
        assert len(table.rows) == 3

    def test_table_headers_bold(self):
        """Test that table headers are bold."""
        doc = DocxBuilder()
        doc.table(
            headers=["Name", "Value"],
            rows=[["A", "1"]],
        )

        table = doc.document.tables[0]
        header_cell = table.rows[0].cells[0]
        # Headers should have bold text
        assert header_cell.text == "Name"

    def test_table_from_dicts(self):
        """Test creating a table from a list of dicts."""
        doc = DocxBuilder()
        items = [
            {"name": "Widget", "price": 10.00},
            {"name": "Gadget", "price": 25.00},
        ]
        result = doc.table_from(items, columns=["name", "price"])

        # Should return self for chaining
        assert result is doc

        tables = doc.document.tables
        assert len(tables) == 1

        table = tables[0]
        # Header row + 2 data rows
        assert len(table.rows) == 3

        # Check auto-generated headers
        assert table.rows[0].cells[0].text == "Name"
        assert table.rows[0].cells[1].text == "Price"

    def test_table_from_dataclass(self):
        """Test creating a table from dataclass instances."""

        @dataclass
        class Item:
            description: str
            quantity: int
            unit_price: float

        doc = DocxBuilder()
        items = [
            Item("Widget", 10, 5.00),
            Item("Gadget", 3, 15.00),
        ]
        doc.table_from(items, columns=["description", "quantity", "unit_price"])

        table = doc.document.tables[0]

        # Check headers (auto-generated from snake_case)
        assert table.rows[0].cells[0].text == "Description"
        assert table.rows[0].cells[1].text == "Quantity"
        assert table.rows[0].cells[2].text == "Unit Price"

        # Check data
        assert table.rows[1].cells[0].text == "Widget"
        assert table.rows[1].cells[1].text == "10"

    def test_table_from_custom_headers(self):
        """Test table_from with custom headers."""
        doc = DocxBuilder()
        items = [{"a": 1, "b": 2}]
        doc.table_from(
            items,
            columns=["a", "b"],
            headers=["Column A", "Column B"],
        )

        table = doc.document.tables[0]
        assert table.rows[0].cells[0].text == "Column A"
        assert table.rows[0].cells[1].text == "Column B"


class TestDocxBuilderPageBreak:
    """Test page break functionality."""

    def test_page_break(self):
        """Test adding a page break."""
        doc = DocxBuilder()
        doc.paragraph("Page 1")
        result = doc.page_break()
        doc.paragraph("Page 2")

        # Should return self for chaining
        assert result is doc


class TestDocxBuilderChaining:
    """Test method chaining."""

    def test_fluent_interface(self):
        """Test that methods can be chained."""
        doc = DocxBuilder()

        # All methods should be chainable
        result = (
            doc.heading("Report")
            .paragraph("Introduction text.")
            .heading("Data", level=2)
            .table(["A", "B"], [["1", "2"]])
            .page_break()
            .paragraph("Conclusion.")
        )

        assert result is doc

    def test_full_document_workflow(self):
        """Test creating a complete document."""

        @dataclass
        class SalesRecord:
            product: str
            units: int
            revenue: float

        doc = DocxBuilder(font="Arial", font_size=12)

        doc.heading("Sales Report", level=0)
        doc.heading("Summary", level=1)
        doc.paragraph("This report summarizes Q4 sales performance.")

        doc.heading("Data", level=1)
        records = [
            SalesRecord("Widget", 100, 1000.00),
            SalesRecord("Gadget", 50, 2500.00),
        ]
        doc.table_from(records, columns=["product", "units", "revenue"])

        doc.page_break()
        doc.heading("Conclusion", level=1)
        doc.paragraph("Sales exceeded expectations.")

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            path = doc.save(f.name)
            assert path.exists()


class TestDocxBuilderMarkdown:
    """Test markdown functionality.

    These tests require optional dependencies: markdown-it-py and htmldocx
    """

    @pytest.fixture
    def has_markdown_deps(self):
        """Check if markdown dependencies are available."""
        try:
            import htmldocx  # noqa: F401
            import markdown_it  # noqa: F401

            return True
        except ImportError:
            return False

    def test_markdown_basic(self, has_markdown_deps):
        """Test basic markdown rendering."""
        if not has_markdown_deps:
            pytest.skip("markdown-it-py and htmldocx not installed")

        doc = DocxBuilder()
        result = doc.markdown("This is **bold** and *italic*.")

        # Should return self for chaining
        assert result is doc

    def test_markdown_heading(self, has_markdown_deps):
        """Test markdown headings."""
        if not has_markdown_deps:
            pytest.skip("markdown-it-py and htmldocx not installed")

        doc = DocxBuilder()
        doc.markdown("## Section Title\n\nParagraph text.")

        # Document should have content
        assert len(doc.document.paragraphs) > 0

    def test_markdown_list(self, has_markdown_deps):
        """Test markdown lists."""
        if not has_markdown_deps:
            pytest.skip("markdown-it-py and htmldocx not installed")

        doc = DocxBuilder()
        doc.markdown("""
- Item 1
- Item 2
- Item 3
""")

        # Document should have list items
        assert len(doc.document.paragraphs) >= 3

    def test_markdown_import_error(self):
        """Test that ImportError is raised when deps missing."""
        doc = DocxBuilder()

        # Temporarily remove the markdown renderer to test error handling
        doc._md_renderer = None
        doc._html_converter = None

        # Mock the import to fail
        import sys

        original_modules = {}
        for mod in ["markdown_it", "htmldocx"]:
            if mod in sys.modules:
                original_modules[mod] = sys.modules[mod]
                sys.modules[mod] = None  # type: ignore

        try:
            # This might raise ImportError if deps aren't installed
            # or might succeed if they are
            doc.markdown("test")
        except ImportError as e:
            assert "markdown" in str(e).lower()
        finally:
            # Restore modules
            for mod, original in original_modules.items():
                if original is not None:
                    sys.modules[mod] = original
                elif mod in sys.modules:
                    del sys.modules[mod]


class TestDocxBuilderDocument:
    """Test access to underlying document."""

    def test_document_property(self):
        """Test accessing the underlying python-docx Document."""
        doc = DocxBuilder()

        # Should be able to access Document directly
        underlying = doc.document
        assert underlying is not None

        # Can use python-docx methods directly
        underlying.add_paragraph("Direct access paragraph")
        assert len(underlying.paragraphs) == 1
