"""
Tests for large document handling and outline mode.

These tests verify:
- OutlineTree for section-level document overview
- expand_section() for progressive section loading
- expand_refs() for targeted element loading
- get_table() for paginated table access
- search() for document-wide text search
- Token budgeting and estimation
- Automatic degradation based on document size
"""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from python_docx_redline.accessibility import (
    AccessibilityTree,
    DocumentSizeInfo,
    ElementType,
    OutlineTree,
    RefTree,
    SearchResults,
    SectionInfo,
    SectionTree,
    TableTree,
    estimate_tokens,
    truncate_to_token_budget,
)
from python_docx_redline.accessibility.outline import SectionDetectionConfig

# ============================================================================
# Test XML Documents
# ============================================================================


DOCUMENT_WITH_SECTIONS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Introduction</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This is the introduction section.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More intro text here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Methods</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This describes the methods used.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Results</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>The results are presented here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Additional result details.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_TABLE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Table follows:</w:t>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A3</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B3</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C3</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A4</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B4</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C4</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A5</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B5</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C5</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"""


DOCUMENT_WITH_SEARCHABLE_TEXT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>The indemnification clause applies to all parties.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This agreement is binding.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>The indemnification includes legal fees.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>All terms and conditions apply.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>No additional indemnification is required.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


LARGE_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {paragraphs}
  </w:body>
</w:document>"""


DOCUMENT_NO_HEADINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph without heading.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph also without heading.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


# ============================================================================
# Helper Functions
# ============================================================================


def create_tree_from_xml(xml_content: str) -> AccessibilityTree:
    """Create an AccessibilityTree from raw XML content."""
    root = etree.fromstring(xml_content.encode("utf-8"))
    return AccessibilityTree.from_xml(root)


def create_outline_from_xml(xml_content: str) -> OutlineTree:
    """Create an OutlineTree from raw XML content."""
    root = etree.fromstring(xml_content.encode("utf-8"))
    return OutlineTree.from_xml(root)


def generate_large_document(num_paragraphs: int) -> str:
    """Generate a document with many paragraphs for testing."""
    paragraphs = []
    for i in range(num_paragraphs):
        paragraphs.append(f"""
    <w:p>
      <w:r>
        <w:t>Paragraph {i + 1}: Lorem ipsum dolor sit amet.</w:t>
      </w:r>
    </w:p>""")
    return LARGE_DOCUMENT_XML.format(paragraphs="".join(paragraphs))


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


# ============================================================================
# OutlineTree Tests
# ============================================================================


class TestOutlineTreeBuilding:
    """Tests for OutlineTree construction."""

    def test_build_outline_from_xml(self) -> None:
        """Test building outline from XML."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        assert outline is not None
        assert len(outline.sections) == 3

    def test_section_headings_extracted(self) -> None:
        """Test that section headings are extracted correctly."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        assert outline.sections[0].heading == "Introduction"
        assert outline.sections[1].heading == "Methods"
        assert outline.sections[2].heading == "Results"

    def test_section_refs_generated(self) -> None:
        """Test that section refs are generated correctly."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        assert outline.sections[0].ref.path == "sec:0"
        assert outline.sections[1].ref.path == "sec:1"
        assert outline.sections[2].ref.path == "sec:2"

    def test_section_paragraph_counts(self) -> None:
        """Test that paragraph counts are correct."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        # Introduction: 2 content paragraphs (not counting heading)
        assert outline.sections[0].paragraph_count == 2

        # Methods: 1 content paragraph
        assert outline.sections[1].paragraph_count == 1

        # Results: 2 content paragraphs
        assert outline.sections[2].paragraph_count == 2

    def test_section_detection_method(self) -> None:
        """Test that detection method is recorded."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        assert outline.sections[0].detection_method == "heading_style"
        assert outline.sections[0].confidence == "high"

    def test_outline_yaml_output(self) -> None:
        """Test YAML output of outline."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        yaml = outline.to_yaml()

        assert "mode: outline" in yaml
        assert "outline:" in yaml
        assert "section [ref=sec:0]" in yaml
        assert 'heading: "Introduction"' in yaml

    def test_document_no_headings(self) -> None:
        """Test document with no headings creates implicit section."""
        outline = create_outline_from_xml(DOCUMENT_NO_HEADINGS)

        assert len(outline.sections) == 1
        assert outline.sections[0].heading == "(Document Content)"
        assert outline.sections[0].detection_method == "implicit"


class TestOutlineTreeSizeInfo:
    """Tests for DocumentSizeInfo in outline mode."""

    def test_size_info_calculated(self) -> None:
        """Test that size info is calculated."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        assert outline.size_info.paragraph_count == 8  # 3 headings + 5 content
        assert outline.size_info.table_count == 0

    def test_degradation_tier_small_document(self) -> None:
        """Test degradation tier for small documents."""
        doc_xml = generate_large_document(50)
        root = etree.fromstring(doc_xml.encode("utf-8"))
        size_info = DocumentSizeInfo.from_xml(root)

        assert size_info.degradation_tier == 1
        assert size_info.recommended_mode == "content"

    def test_degradation_tier_medium_document(self) -> None:
        """Test degradation tier for medium documents."""
        doc_xml = generate_large_document(200)
        root = etree.fromstring(doc_xml.encode("utf-8"))
        size_info = DocumentSizeInfo.from_xml(root)

        assert size_info.degradation_tier == 2
        assert size_info.recommended_mode == "content"

    def test_degradation_tier_large_document(self) -> None:
        """Test degradation tier for large documents."""
        doc_xml = generate_large_document(400)
        root = etree.fromstring(doc_xml.encode("utf-8"))
        size_info = DocumentSizeInfo.from_xml(root)

        assert size_info.degradation_tier == 3
        assert size_info.recommended_mode == "outline"

    def test_degradation_tier_very_large_document(self) -> None:
        """Test degradation tier for very large documents."""
        doc_xml = generate_large_document(600)
        root = etree.fromstring(doc_xml.encode("utf-8"))
        size_info = DocumentSizeInfo.from_xml(root)

        assert size_info.degradation_tier == 4
        assert size_info.recommended_mode == "outline"


# ============================================================================
# expand_section Tests
# ============================================================================


class TestExpandSection:
    """Tests for expand_section() method."""

    def test_expand_section_basic(self) -> None:
        """Test basic section expansion."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)

        section = tree.expand_section("sec:0")

        assert isinstance(section, SectionTree)
        assert section.heading == "Introduction"
        # Section includes heading paragraph + 2 content paragraphs = 3
        assert section.paragraph_count == 3

    def test_expand_section_content_nodes(self) -> None:
        """Test that section content nodes are built."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)

        section = tree.expand_section("sec:0")

        # Should include the heading + 2 content paragraphs = 3 total
        assert len(section.content) == 3
        assert section.content[0].text == "Introduction"
        assert section.content[1].text == "This is the introduction section."

    def test_expand_section_styling_mode(self) -> None:
        """Test section expansion in styling mode."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)

        section = tree.expand_section("sec:0", mode="styling")

        assert isinstance(section, SectionTree)
        # In styling mode, run-level details should be included
        # (children of paragraph nodes)

    def test_expand_section_invalid_ref(self) -> None:
        """Test that invalid section ref raises error."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)

        try:
            tree.expand_section("p:0")  # Not a section ref
            assert False, "Should have raised RefNotFoundError"
        except Exception as e:
            assert "Expected section ref" in str(e)

    def test_expand_section_out_of_bounds(self) -> None:
        """Test that out of bounds section ref raises error."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)

        try:
            tree.expand_section("sec:99")
            assert False, "Should have raised RefNotFoundError"
        except Exception as e:
            assert "out of bounds" in str(e)


# ============================================================================
# expand_refs Tests
# ============================================================================


class TestExpandRefs:
    """Tests for expand_refs() method."""

    def test_expand_refs_single(self) -> None:
        """Test expanding a single ref."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)

        result = tree.expand_refs(["p:0"])

        assert isinstance(result, RefTree)
        assert "p:0" in result.nodes
        assert result.nodes["p:0"].text == "Introduction"

    def test_expand_refs_multiple(self) -> None:
        """Test expanding multiple refs."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)

        result = tree.expand_refs(["p:0", "p:1", "p:2"])

        assert len(result.nodes) == 3
        assert "p:0" in result.nodes
        assert "p:1" in result.nodes
        assert "p:2" in result.nodes

    def test_expand_refs_not_found(self) -> None:
        """Test handling of refs that cannot be found."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)

        result = tree.expand_refs(["p:0", "p:999"])

        assert "p:0" in result.nodes
        assert "p:999" in result.not_found

    def test_expand_refs_styling_mode(self) -> None:
        """Test ref expansion in styling mode."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)

        result = tree.expand_refs(["p:1"], mode="styling")

        assert "p:1" in result.nodes


# ============================================================================
# get_table Tests
# ============================================================================


class TestGetTable:
    """Tests for get_table() method."""

    def test_get_table_all_rows(self) -> None:
        """Test getting all rows from a table."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        table = tree.get_table("tbl:0")

        assert isinstance(table, TableTree)
        assert table.total_rows == 5
        assert table.total_cols == 3
        assert len(table.rows) == 5
        assert table.has_more is False

    def test_get_table_pagination(self) -> None:
        """Test table pagination."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        # Get first 2 rows
        table = tree.get_table("tbl:0", max_rows=2)

        assert len(table.rows) == 2
        assert table.has_more is True
        assert table.page == 1

    def test_get_table_second_page(self) -> None:
        """Test getting second page of table."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        table = tree.get_table("tbl:0", max_rows=2, page=2)

        assert len(table.rows) == 2
        assert table.page == 2
        # Third page would have 1 row

    def test_get_table_last_page(self) -> None:
        """Test getting last page of table."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        table = tree.get_table("tbl:0", max_rows=2, page=3)

        assert len(table.rows) == 1  # Only 1 row left
        assert table.has_more is False

    def test_get_table_yaml_output(self) -> None:
        """Test table YAML output with pagination info."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        table = tree.get_table("tbl:0", max_rows=2)
        yaml = table.to_yaml()

        assert "total_rows: 5" in yaml
        assert "page: 1" in yaml
        assert "has_more: true" in yaml

    def test_get_table_not_found(self) -> None:
        """Test error when table ref not found."""
        tree = create_tree_from_xml(DOCUMENT_WITH_TABLE)

        try:
            tree.get_table("tbl:99")
            assert False, "Should have raised RefNotFoundError"
        except Exception:
            pass


# ============================================================================
# search Tests
# ============================================================================


class TestSearch:
    """Tests for search() method."""

    def test_search_basic(self) -> None:
        """Test basic text search."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SEARCHABLE_TEXT)

        results = tree.search("indemnification")

        assert isinstance(results, SearchResults)
        assert results.total_matches == 3
        assert len(results.results) == 3

    def test_search_case_insensitive(self) -> None:
        """Test case-insensitive search."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SEARCHABLE_TEXT)

        results = tree.search("INDEMNIFICATION")

        assert results.total_matches == 3

    def test_search_case_sensitive(self) -> None:
        """Test case-sensitive search."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SEARCHABLE_TEXT)

        results = tree.search("INDEMNIFICATION", case_sensitive=True)

        assert results.total_matches == 0

    def test_search_max_results(self) -> None:
        """Test limiting search results."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SEARCHABLE_TEXT)

        # Search for a term that has exactly 3 occurrences, limit to 2
        results = tree.search("indemnification", max_results=2)

        assert len(results.results) == 2
        # When we stop at 2, we've counted all matches up to that point
        # The actual count depends on when we stopped
        assert results.total_matches >= 2

    def test_search_context(self) -> None:
        """Test that search results include context."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SEARCHABLE_TEXT)

        results = tree.search("indemnification")

        assert len(results.results) > 0
        assert results.results[0].context != ""
        assert "indemnification" in results.results[0].context.lower()

    def test_search_refs(self) -> None:
        """Test that search results include refs."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SEARCHABLE_TEXT)

        results = tree.search("agreement")

        assert results.results[0].ref.path == "p:1"

    def test_search_no_results(self) -> None:
        """Test search with no results."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SEARCHABLE_TEXT)

        results = tree.search("xyz123nonexistent")

        assert results.total_matches == 0
        assert len(results.results) == 0
        assert results.truncated is False

    def test_search_yaml_output(self) -> None:
        """Test search results YAML output."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SEARCHABLE_TEXT)

        results = tree.search("indemnification")
        yaml = results.to_yaml()

        assert "search_results:" in yaml
        assert 'query: "indemnification"' in yaml
        assert "total_matches: 3" in yaml


# ============================================================================
# Token Budgeting Tests
# ============================================================================


class TestTokenBudgeting:
    """Tests for token estimation and budgeting."""

    def test_estimate_tokens_short_text(self) -> None:
        """Test token estimation for short text."""
        tokens = estimate_tokens("Hello, world!")

        # ~13 characters / 4 = ~3 tokens
        assert tokens == 3

    def test_estimate_tokens_longer_text(self) -> None:
        """Test token estimation for longer text."""
        text = "This is a longer piece of text that should result in more tokens."
        tokens = estimate_tokens(text)

        assert tokens > 10  # Should be roughly len(text) / 4

    def test_truncate_empty_list(self) -> None:
        """Test truncation of empty list."""
        result = truncate_to_token_budget([], max_tokens=100)
        assert result == []

    def test_truncate_under_budget(self) -> None:
        """Test that nodes under budget are not truncated."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)
        nodes = tree.find_all(element_type=ElementType.PARAGRAPH)[:3]

        result = truncate_to_token_budget(nodes, max_tokens=10000)

        assert len(result) == 3

    def test_truncate_unlimited(self) -> None:
        """Test with max_tokens=0 (unlimited)."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)
        nodes = tree.find_all(element_type=ElementType.PARAGRAPH)

        result = truncate_to_token_budget(nodes, max_tokens=0)

        assert len(result) == len(nodes)


# ============================================================================
# Section Detection Config Tests
# ============================================================================


class TestSectionDetectionConfig:
    """Tests for section detection configuration."""

    def test_default_config(self) -> None:
        """Test default configuration values."""
        config = SectionDetectionConfig()

        assert config.detect_heading_styles is True
        assert config.detect_outline_level is True
        assert config.detect_bold_first_line is True
        assert config.max_heading_length == 100

    def test_custom_config(self) -> None:
        """Test custom configuration."""
        config = SectionDetectionConfig(
            detect_heading_styles=True,
            detect_bold_first_line=False,
            max_heading_length=50,
        )

        assert config.detect_bold_first_line is False
        assert config.max_heading_length == 50

    def test_numbering_patterns(self) -> None:
        """Test that numbering patterns are compiled."""
        config = SectionDetectionConfig()

        assert len(config.numbering_patterns) >= 4


# ============================================================================
# SectionInfo Tests
# ============================================================================


class TestSectionInfo:
    """Tests for SectionInfo dataclass."""

    def test_section_info_creation(self) -> None:
        """Test creating SectionInfo."""
        from python_docx_redline.accessibility import Ref

        section = SectionInfo(
            ref=Ref(path="sec:0"),
            heading="Test Section",
            heading_ref=Ref(path="p:0"),
            heading_level=1,
            paragraph_count=5,
            table_count=1,
        )

        assert section.ref.path == "sec:0"
        assert section.heading == "Test Section"
        assert section.paragraph_count == 5

    def test_section_info_defaults(self) -> None:
        """Test SectionInfo default values."""
        from python_docx_redline.accessibility import Ref

        section = SectionInfo(
            ref=Ref(path="sec:0"),
            heading="Test",
            heading_ref=Ref(path="p:0"),
        )

        assert section.heading_level == 1
        assert section.paragraph_count == 0
        assert section.table_count == 0
        assert section.children == []


# ============================================================================
# Integration Tests
# ============================================================================


class TestLargeDocumentIntegration:
    """Integration tests for large document handling."""

    def test_outline_get_section_lookup(self) -> None:
        """Test looking up sections in outline."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        section = outline.get_section("sec:1")

        assert section is not None
        assert section.heading == "Methods"

    def test_outline_get_section_not_found(self) -> None:
        """Test section lookup for non-existent section."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        section = outline.get_section("sec:99")

        assert section is None

    def test_workflow_outline_to_section(self) -> None:
        """Test complete workflow from outline to section expansion."""
        # Build outline
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)
        assert len(outline.sections) == 3

        # Find section of interest
        target_section = None
        for section in outline.sections:
            if section.heading == "Results":
                target_section = section
                break

        assert target_section is not None

        # Expand the section
        tree = create_tree_from_xml(DOCUMENT_WITH_SECTIONS)
        expanded = tree.expand_section(target_section.ref.path)

        assert expanded.heading == "Results"
        assert len(expanded.content) > 0

    def test_search_then_expand(self) -> None:
        """Test workflow of searching then expanding found refs."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SEARCHABLE_TEXT)

        # Search for a term
        results = tree.search("indemnification", max_results=2)

        # Expand the found refs
        refs = [r.ref.path for r in results.results]
        expanded = tree.expand_refs(refs)

        assert len(expanded.nodes) == 2
        for ref in refs:
            assert ref in expanded.nodes


class TestOutlineTreeRepr:
    """Tests for OutlineTree string representation."""

    def test_repr(self) -> None:
        """Test __repr__ output."""
        outline = create_outline_from_xml(DOCUMENT_WITH_SECTIONS)

        repr_str = repr(outline)

        assert "OutlineTree" in repr_str
        assert "sections=3" in repr_str
