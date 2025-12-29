"""
Tests for the AccessibilityTree class.

These tests verify:
- Tree building from document XML
- Three verbosity levels (minimal, standard, full)
- Tracked changes extraction
- YAML serialization format
- Node finding and iteration
"""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from python_docx_redline import Document
from python_docx_redline.accessibility import (
    AccessibilityTree,
    DocumentStats,
    ElementType,
    ViewMode,
)

# Test XML documents

MINIMAL_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Second paragraph heading.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph.</w:t>
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
    <w:p>
      <w:r>
        <w:t>Second paragraph with no changes.</w:t>
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
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:t>Final paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_RUNS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Normal text </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>bold text</w:t>
      </w:r>
      <w:r>
        <w:t> and </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:i/>
        </w:rPr>
        <w:t>italic text</w:t>
      </w:r>
      <w:r>
        <w:t>.</w:t>
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


def create_tree_from_xml(xml_content: str, view_mode: ViewMode | None = None) -> AccessibilityTree:
    """Create an AccessibilityTree from raw XML content."""
    root = etree.fromstring(xml_content.encode("utf-8"))
    return AccessibilityTree.from_xml(root, view_mode=view_mode)


def create_test_docx(content: str) -> Path:
    """Create a minimal but valid OOXML test .docx file."""
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", content)

    return docx_path


class TestAccessibilityTreeBuilding:
    """Tests for tree building."""

    def test_build_tree_from_xml(self) -> None:
        """Test building tree from XML."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        assert tree is not None
        assert tree.root is not None
        assert tree.root.element_type == ElementType.DOCUMENT

    def test_tree_has_correct_stats(self) -> None:
        """Test that tree stats are correct."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        assert tree.stats.paragraphs == 3
        assert tree.stats.tables == 0
        assert tree.stats.tracked_changes == 0

    def test_tree_with_table_stats(self) -> None:
        """Test stats for document with table."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        assert tree.stats.paragraphs == 2  # Only body paragraphs, not table cells
        assert tree.stats.tables == 1

    def test_paragraphs_have_refs(self) -> None:
        """Test that paragraphs get correct refs."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        paragraphs = tree.find_all(element_type=ElementType.PARAGRAPH)
        assert len(paragraphs) == 3
        assert paragraphs[0].ref.path == "p:0"
        assert paragraphs[1].ref.path == "p:1"
        assert paragraphs[2].ref.path == "p:2"

    def test_paragraphs_have_text(self) -> None:
        """Test that paragraphs have correct text."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        paragraphs = tree.find_all(element_type=ElementType.PARAGRAPH)
        assert paragraphs[0].text == "First paragraph."
        assert paragraphs[1].text == "Second paragraph heading."
        assert paragraphs[2].text == "Third paragraph."

    def test_paragraph_style_extracted(self) -> None:
        """Test that paragraph styles are extracted."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        paragraphs = tree.find_all(element_type=ElementType.PARAGRAPH)
        assert paragraphs[0].style is None
        assert paragraphs[1].style == "Heading1"
        assert paragraphs[2].style is None

    def test_heading_level_detected(self) -> None:
        """Test that heading levels are detected."""
        tree = create_tree_from_xml(DOCUMENT_WITH_HEADINGS)

        paragraphs = tree.find_all(element_type=ElementType.PARAGRAPH)

        # Title is level 1
        assert paragraphs[0].level == 1
        assert paragraphs[0].text == "Document Title"

        # Heading1 is level 1
        assert paragraphs[1].level == 1

        # Normal paragraph has no level
        assert paragraphs[2].level is None

        # Heading2 is level 2
        assert paragraphs[3].level == 2


class TestAccessibilityTreeTrackedChanges:
    """Tests for tracked changes extraction."""

    def test_tracked_changes_detected(self) -> None:
        """Test that tracked changes are detected."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)

        assert tree.stats.tracked_changes == 2

    def test_paragraph_has_changes_property(self) -> None:
        """Test that paragraphs with changes have has_changes property."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)

        paragraphs = tree.find_all(element_type=ElementType.PARAGRAPH)

        # First paragraph has changes
        assert paragraphs[0].properties.get("has_changes") == "true"
        assert paragraphs[0].properties.get("change_count") == "2"

        # Second paragraph has no changes
        assert paragraphs[1].properties.get("has_changes") is None

    def test_change_details_extracted(self) -> None:
        """Test that change details are correctly extracted."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)

        paragraphs = tree.find_all(element_type=ElementType.PARAGRAPH)
        changes = paragraphs[0].properties.get("_changes")

        assert changes is not None
        assert len(changes) == 2

        # Check insertion
        insertion = changes[0]
        assert insertion["type"] == "insertion"
        assert insertion["text"] == "with insertion"
        assert insertion["author"] == "Alice"

        # Check deletion
        deletion = changes[1]
        assert deletion["type"] == "deletion"
        assert deletion["text"] == "deleted text"
        assert deletion["author"] == "Bob"


class TestAccessibilityTreeTables:
    """Tests for table building."""

    def test_table_has_ref(self) -> None:
        """Test that table gets correct ref."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        tables = tree.find_all(element_type=ElementType.TABLE)
        assert len(tables) == 1
        assert tables[0].ref.path == "tbl:0"

    def test_table_has_dimensions(self) -> None:
        """Test that table has row and column count."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        tables = tree.find_all(element_type=ElementType.TABLE)
        assert tables[0].properties["rows"] == "2"
        assert tables[0].properties["cols"] == "2"

    def test_table_has_rows(self) -> None:
        """Test that table has row children."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        tables = tree.find_all(element_type=ElementType.TABLE)
        rows = tables[0].children

        assert len(rows) == 2
        assert rows[0].element_type == ElementType.TABLE_ROW
        assert rows[1].element_type == ElementType.TABLE_ROW

    def test_row_has_header_property(self) -> None:
        """Test that first row has header property."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        tables = tree.find_all(element_type=ElementType.TABLE)
        rows = tables[0].children

        assert rows[0].properties.get("header") == "true"
        assert rows[1].properties.get("header") is None

    def test_cells_have_text(self) -> None:
        """Test that cells have correct text."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        tables = tree.find_all(element_type=ElementType.TABLE)
        row1 = tables[0].children[0]
        cells = row1.children

        assert cells[0].text == "Header 1"
        assert cells[1].text == "Header 2"


class TestAccessibilityTreeFinding:
    """Tests for finding nodes."""

    def test_find_by_ref(self) -> None:
        """Test finding node by ref."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        node = tree.find_by_ref("p:1")

        assert node is not None
        assert node.text == "Second paragraph heading."

    def test_find_by_ref_not_found(self) -> None:
        """Test finding non-existent ref returns None."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        node = tree.find_by_ref("p:99")

        assert node is None

    def test_find_all_by_type(self) -> None:
        """Test finding all nodes by type."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        paragraphs = tree.find_all(element_type=ElementType.PARAGRAPH)
        tables = tree.find_all(element_type=ElementType.TABLE)

        assert len(paragraphs) == 2
        assert len(tables) == 1

    def test_find_all_by_heading_level(self) -> None:
        """Test finding all nodes by heading level."""
        tree = create_tree_from_xml(DOCUMENT_WITH_HEADINGS)

        h1 = tree.find_all(heading_level=1)
        h2 = tree.find_all(heading_level=2)

        assert len(h1) == 2  # Title and Heading1
        assert len(h2) == 1  # Heading2

    def test_find_all_by_text_contains(self) -> None:
        """Test finding nodes by text content."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        matches = tree.find_all(text_contains="paragraph")

        # All three paragraphs contain "paragraph"
        assert len(matches) == 3

    def test_find_all_with_changes(self) -> None:
        """Test finding nodes with tracked changes."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)

        # Note: has_changes filter checks node.change, not properties
        # For now, we use properties check internally
        paragraphs = tree.find_all(element_type=ElementType.PARAGRAPH)
        with_changes = [p for p in paragraphs if p.properties.get("has_changes") == "true"]

        assert len(with_changes) == 1

    def test_iter_nodes(self) -> None:
        """Test iterating over all nodes."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        nodes = list(tree.iter_nodes())

        # Root + 3 paragraphs
        assert len(nodes) == 4


class TestYamlMinimalVerbosity:
    """Tests for minimal verbosity YAML output."""

    def test_minimal_yaml_format(self) -> None:
        """Test minimal YAML format."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML, view_mode=ViewMode(verbosity="minimal"))

        yaml = tree.to_yaml()

        # Check header
        assert "document:" in yaml
        assert "verbosity: minimal" in yaml

        # Check minimal format for paragraphs
        assert '- p "First paragraph." [ref=p:0]' in yaml
        assert '- h1 "Second paragraph heading." [ref=p:1]' in yaml
        assert '- p "Third paragraph." [ref=p:2]' in yaml

    def test_minimal_table_format(self) -> None:
        """Test minimal table format."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE, view_mode=ViewMode(verbosity="minimal"))

        yaml = tree.to_yaml()

        # Table shows dimensions only
        assert "- table [ref=tbl:0] [2x2]" in yaml


class TestYamlStandardVerbosity:
    """Tests for standard verbosity YAML output."""

    def test_standard_yaml_format(self) -> None:
        """Test standard YAML format."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        yaml = tree.to_yaml()

        # Check header
        assert "document:" in yaml
        assert "verbosity: standard" in yaml

        # Check stats
        assert "paragraphs: 3" in yaml
        assert "tables: 0" in yaml

        # Check paragraph format
        assert "- paragraph [ref=p:0]:" in yaml
        assert 'text: "First paragraph."' in yaml

    def test_standard_heading_format(self) -> None:
        """Test standard heading format."""
        tree = create_tree_from_xml(DOCUMENT_WITH_HEADINGS)

        yaml = tree.to_yaml()

        assert "- heading [ref=p:0] [level=1]:" in yaml
        assert 'text: "Document Title"' in yaml
        assert "style: Title" in yaml

    def test_standard_table_format(self) -> None:
        """Test standard table format."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        yaml = tree.to_yaml()

        # Table with dimensions
        assert "- table [ref=tbl:0] [rows=2] [cols=2]:" in yaml

        # Header row
        assert "- row [ref=tbl:0/row:0] [header]:" in yaml

        # Cells
        assert '- cell: "Header 1"' in yaml

    def test_standard_tracked_changes(self) -> None:
        """Test standard format with tracked changes."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)

        yaml = tree.to_yaml()

        assert "has_changes: true" in yaml
        assert "changes:" in yaml
        assert "- type: insertion" in yaml
        assert 'text: "with insertion"' in yaml
        assert "author: Alice" in yaml


class TestYamlFullVerbosity:
    """Tests for full verbosity YAML output."""

    def test_full_yaml_shows_runs(self) -> None:
        """Test full verbosity shows runs."""
        tree = create_tree_from_xml(
            DOCUMENT_WITH_RUNS,
            view_mode=ViewMode(verbosity="full", include_formatting=True),
        )

        yaml = tree.to_yaml()

        # Should show runs
        assert "runs:" in yaml
        assert "[bold]" in yaml
        assert "[italic]" in yaml


class TestYamlTrackedChangesSummary:
    """Tests for tracked changes summary in YAML."""

    def test_tracked_changes_summary_at_end(self) -> None:
        """Test that tracked changes are summarized at the end."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TRACKED_CHANGES)

        yaml = tree.to_yaml()

        # Should have tracked_changes section at end
        assert "\ntracked_changes:" in yaml
        assert "- ref: change:0" in yaml
        assert "- ref: change:1" in yaml
        assert "location: p:0" in yaml


class TestAccessibilityTreeIntegration:
    """Integration tests using full Document class."""

    def test_tree_from_document(self) -> None:
        """Test building tree from Document object."""
        docx_path = create_test_docx(MINIMAL_DOCUMENT_XML)

        try:
            doc = Document(docx_path)
            tree = AccessibilityTree.from_document(doc)

            assert tree is not None
            assert tree.stats.paragraphs == 3
            assert tree.document_path == docx_path

        finally:
            docx_path.unlink()

    def test_tree_yaml_includes_path(self) -> None:
        """Test that YAML includes document path."""
        docx_path = create_test_docx(MINIMAL_DOCUMENT_XML)

        try:
            doc = Document(docx_path)
            tree = AccessibilityTree.from_document(doc)

            yaml = tree.to_yaml()
            assert f'path: "{docx_path}"' in yaml

        finally:
            docx_path.unlink()


class TestViewModeConfiguration:
    """Tests for ViewMode configuration."""

    def test_verbosity_override_in_to_yaml(self) -> None:
        """Test that verbosity can be overridden in to_yaml()."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML, view_mode=ViewMode(verbosity="minimal"))

        # Default should be minimal
        yaml_minimal = tree.to_yaml()
        assert "verbosity: minimal" in yaml_minimal

        # Override to standard
        yaml_standard = tree.to_yaml(verbosity="standard")
        assert "verbosity: standard" in yaml_standard

    def test_include_tracked_changes_disabled(self) -> None:
        """Test disabling tracked changes in output."""
        tree = create_tree_from_xml(
            DOCUMENT_WITH_TRACKED_CHANGES,
            view_mode=ViewMode(include_tracked_changes=False),
        )

        yaml = tree.to_yaml()

        # Should not include tracked changes section
        assert "\ntracked_changes:" not in yaml


class TestDocumentStats:
    """Tests for DocumentStats dataclass."""

    def test_default_stats(self) -> None:
        """Test default DocumentStats values."""
        stats = DocumentStats()

        assert stats.paragraphs == 0
        assert stats.tables == 0
        assert stats.tracked_changes == 0
        assert stats.comments == 0

    def test_stats_with_values(self) -> None:
        """Test DocumentStats with values."""
        stats = DocumentStats(
            paragraphs=10,
            tables=2,
            tracked_changes=5,
            comments=3,
        )

        assert stats.paragraphs == 10
        assert stats.tables == 2
        assert stats.tracked_changes == 5
        assert stats.comments == 3


class TestAccessibilityTreeRepr:
    """Tests for tree string representation."""

    def test_repr(self) -> None:
        """Test __repr__ output."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        repr_str = repr(tree)

        assert "AccessibilityTree" in repr_str
        assert "paragraphs=2" in repr_str
        assert "tables=1" in repr_str
        assert "tracked_changes=0" in repr_str
