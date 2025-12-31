"""
Tests for text and markdown export from AccessibilityTree.

These tests verify:
- Plain text export (to_text)
- Markdown export (to_markdown)
- Tracked changes handling (accept/reject/all modes)
- Table rendering (markdown/simple/grid formats)
- Heading styles (underline/prefix)
- Configuration options
"""

from lxml import etree

from python_docx_redline.accessibility import (
    AccessibilityTree,
    TextExportConfig,
    ViewMode,
)

# ============================================================================
# Test XML Documents
# ============================================================================

MINIMAL_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_HEADINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Title"/>
      </w:pPr>
      <w:r>
        <w:t>Document Title</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>First Section</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading2"/>
      </w:pPr>
      <w:r>
        <w:t>Subsection</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More content.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_TRACKED_CHANGES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Original text </w:t>
      </w:r>
      <w:ins w:id="1" w:author="Alice" w:date="2024-01-15T10:30:00Z">
        <w:r>
          <w:t>with insertion</w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t> and </w:t>
      </w:r>
      <w:del w:id="2" w:author="Bob" w:date="2024-01-16T14:00:00Z">
        <w:r>
          <w:delText>deleted text</w:delText>
        </w:r>
      </w:del>
      <w:r>
        <w:t> here.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_TABLE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Intro paragraph.</w:t>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Header 1</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Header 2</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Header 3</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 1</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 2</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 3</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:t>Final paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


# ============================================================================
# Helper Functions
# ============================================================================


def create_tree_from_xml(xml_content: str, view_mode: ViewMode | None = None) -> AccessibilityTree:
    """Create an AccessibilityTree from raw XML content."""
    root = etree.fromstring(xml_content.encode("utf-8"))
    return AccessibilityTree.from_xml(root, view_mode=view_mode)


# ============================================================================
# Plain Text Export Tests
# ============================================================================


class TestPlainTextBasic:
    """Tests for basic plain text export."""

    def test_basic_paragraphs(self) -> None:
        """Test exporting simple paragraphs to plain text."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)
        text = tree.to_text()

        assert "First paragraph." in text
        assert "Second paragraph." in text

    def test_paragraphs_separated_by_blank_lines(self) -> None:
        """Test that paragraphs are separated by blank lines."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)
        text = tree.to_text()

        # Should have blank line between paragraphs
        assert "First paragraph.\n\nSecond paragraph." in text


class TestPlainTextHeadings:
    """Tests for heading rendering in plain text."""

    def test_heading_underline_style(self) -> None:
        """Test heading with underline style (default)."""
        tree = create_tree_from_xml(DOCUMENT_WITH_HEADINGS)
        config = TextExportConfig(heading_style="underline")
        text = tree.to_text(config)

        # Title should have = underline (length matches "Document Title" = 14 chars)
        assert "Document Title\n==============" in text

        # Heading1 should have = underline (Title and Heading1 are both level 1)
        assert "First Section\n=============" in text

        # Heading2 should have - underline
        assert "Subsection\n----------" in text

    def test_heading_prefix_style(self) -> None:
        """Test heading with # prefix style."""
        tree = create_tree_from_xml(DOCUMENT_WITH_HEADINGS)
        config = TextExportConfig(heading_style="prefix")
        text = tree.to_text(config)

        # Title and Heading1 should have # prefix
        assert "# Document Title" in text
        assert "# First Section" in text

        # Heading2 should have ## prefix
        assert "## Subsection" in text


class TestPlainTextTables:
    """Tests for table rendering in plain text."""

    def test_table_simple_format(self) -> None:
        """Test table in simple pipe format."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)
        config = TextExportConfig(table_format="simple")
        text = tree.to_text(config)

        # Should have header row
        assert "Header 1" in text
        assert "Header 2" in text
        assert "Header 3" in text

        # Should have separator
        assert "---" in text

        # Should have data row
        assert "Cell 1" in text
        assert "Cell 2" in text

    def test_table_grid_format(self) -> None:
        """Test table in grid format."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)
        config = TextExportConfig(table_format="grid")
        text = tree.to_text(config)

        # Should have grid characters
        assert "+" in text
        assert "|" in text

        # Should have double line for header separator
        assert "=" in text


# ============================================================================
# Markdown Export Tests
# ============================================================================


class TestMarkdownBasic:
    """Tests for basic markdown export."""

    def test_basic_paragraphs(self) -> None:
        """Test exporting simple paragraphs to markdown."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)
        md = tree.to_markdown()

        assert "First paragraph." in md
        assert "Second paragraph." in md


class TestMarkdownHeadings:
    """Tests for heading rendering in markdown."""

    def test_heading_hash_syntax(self) -> None:
        """Test that headings use # syntax in markdown."""
        tree = create_tree_from_xml(DOCUMENT_WITH_HEADINGS)
        md = tree.to_markdown()

        # Title and Heading1 should have # prefix (both are level 1)
        assert "# Document Title" in md
        assert "# First Section" in md

        # Heading2 should have ## prefix
        assert "## Subsection" in md


class TestMarkdownTables:
    """Tests for table rendering in markdown."""

    def test_table_markdown_format(self) -> None:
        """Test table in standard markdown format."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)
        config = TextExportConfig(table_format="markdown")
        md = tree.to_markdown(config)

        # Should have pipe characters
        assert "| Header 1" in md
        assert "| Header 2" in md

        # Should have markdown separator
        assert "|---" in md or "| ---" in md or "|--" in md

    def test_table_simple_format(self) -> None:
        """Test table in simple format in markdown."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)
        config = TextExportConfig(table_format="simple")
        md = tree.to_markdown(config)

        assert "Header 1" in md
        assert "-+-" in md

    def test_table_grid_format(self) -> None:
        """Test table in grid format in markdown."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)
        config = TextExportConfig(table_format="grid")
        md = tree.to_markdown(config)

        assert "+" in md
        assert "=" in md  # Header separator


# ============================================================================
# Tracked Changes Tests
# ============================================================================


class TestTrackedChangesAccept:
    """Tests for accept mode (default) - show insertions, hide deletions."""

    def test_accept_mode_shows_insertions(self) -> None:
        """Test that accept mode shows insertions."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="accept")
        text = tree.to_text(config)

        # Insertion should be visible
        assert "with insertion" in text

    def test_accept_mode_hides_deletions(self) -> None:
        """Test that accept mode hides deletions."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="accept")
        text = tree.to_text(config)

        # Deletion should be hidden
        assert "deleted text" not in text

    def test_accept_mode_is_default(self) -> None:
        """Test that accept mode is the default."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        text = tree.to_text()  # No config, should use defaults

        assert "with insertion" in text
        assert "deleted text" not in text


class TestTrackedChangesReject:
    """Tests for reject mode - show deletions, hide insertions."""

    def test_reject_mode_shows_deletions(self) -> None:
        """Test that reject mode shows deletions."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="reject")
        text = tree.to_text(config)

        # Deletion should be visible
        assert "deleted text" in text

    def test_reject_mode_hides_insertions(self) -> None:
        """Test that reject mode hides insertions."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="reject")
        text = tree.to_text(config)

        # Insertion should be hidden
        assert "with insertion" not in text


class TestTrackedChangesAll:
    """Tests for all mode - show both with CriticMarkup syntax."""

    def test_all_mode_shows_both(self) -> None:
        """Test that all mode shows both insertions and deletions."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="all")
        text = tree.to_text(config)

        # Both should be present
        assert "with insertion" in text
        assert "deleted text" in text

    def test_all_mode_uses_criticmarkup_insertion(self) -> None:
        """Test that all mode uses CriticMarkup syntax for insertions."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="all")
        text = tree.to_text(config)

        # Insertion should have {++...++} markers
        assert "{++with insertion++}" in text

    def test_all_mode_uses_criticmarkup_deletion(self) -> None:
        """Test that all mode uses CriticMarkup syntax for deletions."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="all")
        text = tree.to_text(config)

        # Deletion should have {--...--} markers
        assert "{--deleted text--}" in text


class TestTrackedChangesInMarkdown:
    """Tests for tracked changes in markdown export."""

    def test_markdown_accept_mode(self) -> None:
        """Test accept mode in markdown."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="accept")
        md = tree.to_markdown(config)

        assert "with insertion" in md
        assert "deleted text" not in md

    def test_markdown_reject_mode(self) -> None:
        """Test reject mode in markdown."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="reject")
        md = tree.to_markdown(config)

        assert "deleted text" in md
        assert "with insertion" not in md

    def test_markdown_all_mode(self) -> None:
        """Test all mode in markdown."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)
        config = TextExportConfig(tracked_changes="all")
        md = tree.to_markdown(config)

        assert "{++with insertion++}" in md
        assert "{--deleted text--}" in md


# ============================================================================
# Configuration Tests
# ============================================================================


class TestTextExportConfig:
    """Tests for TextExportConfig dataclass."""

    def test_default_values(self) -> None:
        """Test default configuration values."""
        config = TextExportConfig()

        assert config.include_headers is True
        assert config.include_footers is True
        assert config.include_footnotes is True
        assert config.include_endnotes is True
        assert config.include_comments is False
        assert config.include_images is True
        assert config.tracked_changes == "accept"
        assert config.table_format == "markdown"
        assert config.line_width == 0
        assert config.heading_style == "underline"

    def test_custom_values(self) -> None:
        """Test custom configuration values."""
        config = TextExportConfig(
            include_comments=True,
            tracked_changes="all",
            table_format="grid",
            line_width=80,
            heading_style="prefix",
        )

        assert config.include_comments is True
        assert config.tracked_changes == "all"
        assert config.table_format == "grid"
        assert config.line_width == 80
        assert config.heading_style == "prefix"


class TestLineWrapping:
    """Tests for line wrapping functionality."""

    def test_no_wrapping_by_default(self) -> None:
        """Test that lines are not wrapped by default."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)
        text = tree.to_text()

        # Lines should not be wrapped
        assert "First paragraph." in text

    def test_wrapping_with_config(self) -> None:
        """Test that lines are wrapped when configured."""
        # Create a document with a long paragraph
        long_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a very long paragraph that should be wrapped when the line width is set to a small value like 40 characters.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

        tree = create_tree_from_xml(long_xml)
        config = TextExportConfig(line_width=40)
        text = tree.to_text(config)

        # Lines should be wrapped
        lines = text.strip().split("\n")
        # With wrapping at 40 chars, the paragraph should span multiple lines
        # (empty lines don't count)
        non_empty_lines = [line for line in lines if line.strip()]
        assert len(non_empty_lines) > 1


# ============================================================================
# Integration Tests
# ============================================================================


class TestExportIntegration:
    """Integration tests combining multiple features."""

    def test_document_with_headings_and_content(self) -> None:
        """Test exporting a document with headings and regular content."""
        tree = create_tree_from_xml(DOCUMENT_WITH_HEADINGS)

        text = tree.to_text()
        md = tree.to_markdown()

        # Both should have all content
        for output in [text, md]:
            assert "Document Title" in output
            assert "First Section" in output
            assert "Content paragraph." in output
            assert "Subsection" in output
            assert "More content." in output

    def test_document_with_table_and_text(self) -> None:
        """Test exporting a document with tables and regular text."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        text = tree.to_text()
        md = tree.to_markdown()

        # Both should have all content
        for output in [text, md]:
            assert "Intro paragraph." in output
            assert "Header 1" in output
            assert "Cell 1" in output
            assert "Final paragraph." in output

    def test_export_consistency(self) -> None:
        """Test that multiple exports produce consistent results."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        text1 = tree.to_text()
        text2 = tree.to_text()
        md1 = tree.to_markdown()
        md2 = tree.to_markdown()

        assert text1 == text2
        assert md1 == md2


class TestEdgeCases:
    """Tests for edge cases and special scenarios."""

    def test_empty_document(self) -> None:
        """Test exporting an empty document."""
        empty_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
  </w:body>
</w:document>"""

        tree = create_tree_from_xml(empty_xml)
        text = tree.to_text()
        md = tree.to_markdown()

        # Should not crash, should produce minimal output
        assert text.strip() == ""
        assert md.strip() == ""

    def test_empty_paragraph(self) -> None:
        """Test exporting a document with empty paragraphs."""
        empty_para_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p></w:p>
    <w:p>
      <w:r>
        <w:t>Content</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

        tree = create_tree_from_xml(empty_para_xml)
        text = tree.to_text()

        assert "Content" in text

    def test_table_with_varying_column_counts(self) -> None:
        """Test table where rows have different column counts."""
        uneven_table_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>A</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>C</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"""

        tree = create_tree_from_xml(uneven_table_xml)

        # Should not crash
        text = tree.to_text()
        markdown = tree.to_markdown()

        assert "A" in text
        assert "B" in text
        assert "C" in text
        # Verify markdown also works
        assert "A" in markdown


class TestFromAccessibilityModule:
    """Test that exports can be accessed via the accessibility module."""

    def test_import_text_export_config(self) -> None:
        """Test that TextExportConfig can be imported from accessibility."""
        from python_docx_redline.accessibility import TextExportConfig

        config = TextExportConfig()
        assert config is not None

    def test_tree_has_to_text_method(self) -> None:
        """Test that AccessibilityTree has to_text method."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        assert hasattr(tree, "to_text")
        assert callable(tree.to_text)

    def test_tree_has_to_markdown_method(self) -> None:
        """Test that AccessibilityTree has to_markdown method."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        assert hasattr(tree, "to_markdown")
        assert callable(tree.to_markdown)
