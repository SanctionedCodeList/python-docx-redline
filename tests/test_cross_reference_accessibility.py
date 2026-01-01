"""
Tests for cross-reference extraction and accessibility tree support.

These tests verify:
- Cross-reference extraction from document XML (REF, PAGEREF, NOTEREF fields)
- Both simple and complex field formats
- Broken cross-reference detection
- Integration with AccessibilityTree
- YAML output for cross-references
"""

from lxml import etree

from python_docx_redline.accessibility import (
    AccessibilityTree,
    ViewMode,
)
from python_docx_redline.accessibility.bookmarks import (
    BookmarkRegistry,
    CrossReferenceRegistry,
)
from python_docx_redline.accessibility.types import FieldType

# Test XML documents

DOCUMENT_WITH_SIMPLE_REF = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="Section1"/>
      <w:r>
        <w:t>Section 1 Content</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:r>
        <w:t>See also </w:t>
      </w:r>
      <w:fldSimple w:instr=" REF Section1 \\h ">
        <w:r>
          <w:t>Section 1 Content</w:t>
        </w:r>
      </w:fldSimple>
      <w:r>
        <w:t> for more details.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_COMPLEX_REF = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="TargetBookmark"/>
      <w:r>
        <w:t>Target Content</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:instrText> REF TargetBookmark \\h </w:instrText>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <w:t>Target Content</w:t>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_PAGEREF = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="Chapter1"/>
      <w:r>
        <w:t>Chapter 1</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:r>
        <w:t>See page </w:t>
      </w:r>
      <w:fldSimple w:instr=" PAGEREF Chapter1 \\h ">
        <w:r>
          <w:t>5</w:t>
        </w:r>
      </w:fldSimple>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_NOTEREF = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="_Ref123456"/>
      <w:r>
        <w:t>Footnote reference</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:r>
        <w:t>See footnote </w:t>
      </w:r>
      <w:fldSimple w:instr=" NOTEREF _Ref123456 \\h ">
        <w:r>
          <w:t>1</w:t>
        </w:r>
      </w:fldSimple>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_BROKEN_XREF = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="ValidBookmark"/>
      <w:r>
        <w:t>Valid content.</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:fldSimple w:instr=" REF NonExistentBookmark \\h ">
        <w:r>
          <w:t>Error! Reference source not found.</w:t>
        </w:r>
      </w:fldSimple>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_MULTIPLE_XREFS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="Section1"/>
      <w:r>
        <w:t>First Section</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="Section2"/>
      <w:r>
        <w:t>Second Section</w:t>
      </w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
    <w:p>
      <w:r>
        <w:t>References: </w:t>
      </w:r>
      <w:fldSimple w:instr=" REF Section1 \\h ">
        <w:r>
          <w:t>First Section</w:t>
        </w:r>
      </w:fldSimple>
      <w:r>
        <w:t> and </w:t>
      </w:r>
      <w:fldSimple w:instr=" REF Section2 \\h ">
        <w:r>
          <w:t>Second Section</w:t>
        </w:r>
      </w:fldSimple>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Page of Section 1: </w:t>
      </w:r>
      <w:fldSimple w:instr=" PAGEREF Section1 \\h ">
        <w:r>
          <w:t>1</w:t>
        </w:r>
      </w:fldSimple>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_DIRTY_XREF = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="Target"/>
      <w:r>
        <w:t>Target content</w:t>
      </w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:fldSimple w:instr=" REF Target \\h " w:dirty="true">
        <w:r>
          <w:t>Outdated value</w:t>
        </w:r>
      </w:fldSimple>
    </w:p>
  </w:body>
</w:document>"""


MINIMAL_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Simple paragraph with no cross-references.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


def create_registry_from_xml(
    xml_content: str,
    bookmark_registry: BookmarkRegistry | None = None,
) -> CrossReferenceRegistry:
    """Create a CrossReferenceRegistry from raw XML content."""
    root = etree.fromstring(xml_content.encode("utf-8"))
    if bookmark_registry is None:
        bookmark_registry = BookmarkRegistry.from_xml(root)
    return CrossReferenceRegistry.from_xml(root, bookmark_registry)


def create_tree_from_xml(
    xml_content: str,
    view_mode: ViewMode | None = None,
) -> AccessibilityTree:
    """Create an AccessibilityTree from raw XML content."""
    root = etree.fromstring(xml_content.encode("utf-8"))
    return AccessibilityTree.from_xml(root, view_mode=view_mode)


class TestCrossReferenceExtraction:
    """Tests for cross-reference extraction."""

    def test_extract_simple_ref(self) -> None:
        """Test extraction of simple REF field."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        assert len(registry.cross_references) == 1

        xref = registry.cross_references[0]
        assert xref.field_type == FieldType.REF
        assert xref.target_bookmark == "Section1"

    def test_extract_complex_ref(self) -> None:
        """Test extraction of complex REF field."""
        registry = create_registry_from_xml(DOCUMENT_WITH_COMPLEX_REF)

        assert len(registry.cross_references) == 1

        xref = registry.cross_references[0]
        assert xref.field_type == FieldType.REF
        assert xref.target_bookmark == "TargetBookmark"

    def test_extract_pageref(self) -> None:
        """Test extraction of PAGEREF field."""
        registry = create_registry_from_xml(DOCUMENT_WITH_PAGEREF)

        assert len(registry.cross_references) == 1

        xref = registry.cross_references[0]
        assert xref.field_type == FieldType.PAGEREF
        assert xref.target_bookmark == "Chapter1"
        assert xref.display_value == "5"

    def test_extract_noteref(self) -> None:
        """Test extraction of NOTEREF field."""
        registry = create_registry_from_xml(DOCUMENT_WITH_NOTEREF)

        assert len(registry.cross_references) == 1

        xref = registry.cross_references[0]
        assert xref.field_type == FieldType.NOTEREF
        assert xref.target_bookmark == "_Ref123456"
        assert xref.display_value == "1"

    def test_xref_has_ref(self) -> None:
        """Test that cross-references have sequential refs."""
        registry = create_registry_from_xml(DOCUMENT_WITH_MULTIPLE_XREFS)

        assert len(registry.cross_references) == 3
        assert registry.cross_references[0].ref == "xref:0"
        assert registry.cross_references[1].ref == "xref:1"
        assert registry.cross_references[2].ref == "xref:2"

    def test_xref_has_from_location(self) -> None:
        """Test that cross-references have from_location."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        xref = registry.cross_references[0]
        assert xref.from_location == "p:1"

    def test_xref_has_display_value(self) -> None:
        """Test that cross-references have display value."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        xref = registry.cross_references[0]
        assert xref.display_value == "Section 1 Content"

    def test_xref_hyperlink_switch(self) -> None:
        """Test that \\h switch is detected."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        xref = registry.cross_references[0]
        assert xref.is_hyperlink is True

    def test_xref_target_location_resolved(self) -> None:
        """Test that target location is resolved from bookmark."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        xref = registry.cross_references[0]
        assert xref.target_location == "p:0"

    def test_no_xrefs_in_minimal_doc(self) -> None:
        """Test that document without cross-refs returns empty list."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        assert len(registry.cross_references) == 0

    def test_dirty_xref_detected(self) -> None:
        """Test that dirty field flag is detected."""
        registry = create_registry_from_xml(DOCUMENT_WITH_DIRTY_XREF)

        assert len(registry.cross_references) == 1
        xref = registry.cross_references[0]
        assert xref.is_dirty is True


class TestBrokenCrossReferences:
    """Tests for broken cross-reference detection."""

    def test_detect_broken_xref(self) -> None:
        """Test detection of cross-reference to missing bookmark."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BROKEN_XREF)

        assert len(registry.cross_references) == 1

        xref = registry.cross_references[0]
        assert xref.is_broken is True
        assert xref.target_bookmark == "NonExistentBookmark"
        assert xref.error is not None

    def test_get_broken_xrefs(self) -> None:
        """Test get_broken() method."""
        registry = create_registry_from_xml(DOCUMENT_WITH_BROKEN_XREF)

        broken = registry.get_broken()

        assert len(broken) == 1
        assert broken[0].target_bookmark == "NonExistentBookmark"

    def test_valid_xref_not_broken(self) -> None:
        """Test that valid cross-references are not marked as broken."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        xref = registry.cross_references[0]
        assert xref.is_broken is False
        assert xref.error is None


class TestCrossReferenceQueries:
    """Tests for cross-reference query methods."""

    def test_get_all(self) -> None:
        """Test get_all() method."""
        registry = create_registry_from_xml(DOCUMENT_WITH_MULTIPLE_XREFS)

        all_xrefs = registry.get_all()

        assert len(all_xrefs) == 3

    def test_get_by_target(self) -> None:
        """Test get_by_target() method."""
        registry = create_registry_from_xml(DOCUMENT_WITH_MULTIPLE_XREFS)

        xrefs_to_section1 = registry.get_by_target("Section1")

        # Should have 2 references to Section1 (REF and PAGEREF)
        assert len(xrefs_to_section1) == 2

    def test_get_by_target_none_found(self) -> None:
        """Test get_by_target() when no matches."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        xrefs = registry.get_by_target("NonExistent")

        assert len(xrefs) == 0

    def test_get_dirty(self) -> None:
        """Test get_dirty() method."""
        registry = create_registry_from_xml(DOCUMENT_WITH_DIRTY_XREF)

        dirty = registry.get_dirty()

        assert len(dirty) == 1


class TestAccessibilityTreeCrossReferences:
    """Tests for cross-references in AccessibilityTree."""

    def test_tree_has_xref_stats(self) -> None:
        """Test that tree stats include cross-reference count."""
        tree = create_tree_from_xml(DOCUMENT_WITH_MULTIPLE_XREFS)

        assert tree.stats.cross_references == 3

    def test_tree_no_xref_stats_when_empty(self) -> None:
        """Test that tree stats are zero when no cross-refs."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        assert tree.stats.cross_references == 0

    def test_tree_cross_references_property(self) -> None:
        """Test accessing cross_references via tree property."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        xrefs = tree.cross_references

        assert len(xrefs) == 1
        assert xrefs[0].target_bookmark == "Section1"

    def test_tree_get_cross_references_to(self) -> None:
        """Test get_cross_references_to() method."""
        tree = create_tree_from_xml(DOCUMENT_WITH_MULTIPLE_XREFS)

        xrefs = tree.get_cross_references_to("Section1")

        assert len(xrefs) == 2

    def test_tree_validate_references_includes_xrefs(self) -> None:
        """Test that validate_references includes broken cross-refs."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BROKEN_XREF)

        result = tree.validate_references()

        assert not result.is_valid
        assert len(result.broken_cross_references) == 1
        assert "NonExistentBookmark" in result.missing_bookmarks


class TestYamlCrossReferenceOutput:
    """Tests for YAML output of cross-references."""

    def test_yaml_includes_xref_stats(self) -> None:
        """Test that YAML output includes cross_references stats."""
        tree = create_tree_from_xml(DOCUMENT_WITH_MULTIPLE_XREFS)

        yaml = tree.to_yaml()

        assert "cross_references: 3" in yaml

    def test_yaml_no_xref_stats_when_empty(self) -> None:
        """Test that cross_references stat is omitted when zero."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        yaml = tree.to_yaml()

        # Should not have cross_references in stats
        assert "cross_references: 0" not in yaml

    def test_yaml_includes_xref_section(self) -> None:
        """Test that YAML includes cross_references section."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        yaml = tree.to_yaml()

        assert "\ncross_references:" in yaml

    def test_yaml_xref_has_ref(self) -> None:
        """Test that YAML cross-references have refs."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        yaml = tree.to_yaml()

        assert "- ref: xref:0" in yaml

    def test_yaml_xref_has_target(self) -> None:
        """Test that YAML cross-references have target bookmark."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        yaml = tree.to_yaml()

        assert "target: Section1" in yaml

    def test_yaml_xref_has_from(self) -> None:
        """Test that YAML cross-references have from location."""
        tree = create_tree_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        yaml = tree.to_yaml()

        assert "from: p:1" in yaml

    def test_yaml_xref_groups_by_type(self) -> None:
        """Test that YAML groups cross-references by type."""
        tree = create_tree_from_xml(DOCUMENT_WITH_MULTIPLE_XREFS)

        yaml = tree.to_yaml()

        assert "text_references:" in yaml
        assert "page_references:" in yaml

    def test_yaml_broken_xref_section(self) -> None:
        """Test that broken cross-refs appear in broken section."""
        tree = create_tree_from_xml(DOCUMENT_WITH_BROKEN_XREF)

        yaml = tree.to_yaml()

        assert "broken:" in yaml
        assert "NonExistentBookmark" in yaml

    def test_yaml_no_xref_section_when_empty(self) -> None:
        """Test that cross_references section is omitted when no xrefs."""
        tree = create_tree_from_xml(MINIMAL_DOCUMENT_XML)

        yaml = tree.to_yaml()

        assert "\ncross_references:" not in yaml


class TestCrossReferenceRegistryYamlDict:
    """Tests for CrossReferenceRegistry.to_yaml_dict()."""

    def test_yaml_dict_structure(self) -> None:
        """Test the structure of YAML dict output."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        result = registry.to_yaml_dict()

        assert "cross_references" in result
        assert len(result["cross_references"]) == 1

    def test_yaml_dict_xref_fields(self) -> None:
        """Test that YAML dict contains all xref fields."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        result = registry.to_yaml_dict()
        xref = result["cross_references"][0]

        assert "ref" in xref
        assert "type" in xref
        assert "target" in xref
        assert "from" in xref

    def test_yaml_dict_empty_when_no_xrefs(self) -> None:
        """Test that empty document produces empty dict."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        result = registry.to_yaml_dict()

        assert result == {} or "cross_references" not in result


class TestCrossReferenceFieldTypes:
    """Tests for different field types."""

    def test_ref_field_type(self) -> None:
        """Test REF field type detection."""
        registry = create_registry_from_xml(DOCUMENT_WITH_SIMPLE_REF)

        xref = registry.cross_references[0]
        assert xref.field_type == FieldType.REF

    def test_pageref_field_type(self) -> None:
        """Test PAGEREF field type detection."""
        registry = create_registry_from_xml(DOCUMENT_WITH_PAGEREF)

        xref = registry.cross_references[0]
        assert xref.field_type == FieldType.PAGEREF

    def test_noteref_field_type(self) -> None:
        """Test NOTEREF field type detection."""
        registry = create_registry_from_xml(DOCUMENT_WITH_NOTEREF)

        xref = registry.cross_references[0]
        assert xref.field_type == FieldType.NOTEREF
