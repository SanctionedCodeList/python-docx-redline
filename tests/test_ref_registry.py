"""
Tests for the RefRegistry class.

These tests verify ref resolution and generation:
- Ordinal-based resolution (p:5)
- Fingerprint-based resolution (p:~xK4mNp2q)
- Nested ref resolution (tbl:0/row:2/cell:1)
- Cache invalidation
"""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from python_docx_redline import Document
from python_docx_redline.accessibility.registry import RefRegistry
from python_docx_redline.accessibility.types import ElementType
from python_docx_redline.constants import WORD_NAMESPACE
from python_docx_redline.errors import RefNotFoundError

# Minimal Word document XML structure
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


DOCUMENT_WITH_TABLE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
              <w:t>Cell 0,0</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 0,1</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 1,0</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell 1,1</w:t>
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


def create_test_docx(content: str = MINIMAL_DOCUMENT_XML) -> Path:
    """Create a minimal but valid OOXML test .docx file.

    Args:
        content: The document.xml content

    Returns:
        Path to the created .docx file
    """
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


def create_registry_from_xml(xml_content: str) -> RefRegistry:
    """Create a RefRegistry from raw XML content.

    Args:
        xml_content: Document XML content

    Returns:
        RefRegistry initialized with the parsed XML
    """
    root = etree.fromstring(xml_content.encode("utf-8"))
    return RefRegistry(root)


class TestRefRegistryOrdinalResolution:
    """Tests for ordinal-based ref resolution."""

    def test_resolve_first_paragraph(self) -> None:
        """Test resolving the first paragraph."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        element = registry.resolve_ref("p:0")

        assert element is not None
        assert element.tag == f"{{{WORD_NAMESPACE}}}p"
        # Check it contains "First paragraph"
        text = registry._get_text_content(element)
        assert "First paragraph" in text

    def test_resolve_second_paragraph(self) -> None:
        """Test resolving the second paragraph."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        element = registry.resolve_ref("p:1")

        text = registry._get_text_content(element)
        assert "Second paragraph" in text

    def test_resolve_last_paragraph(self) -> None:
        """Test resolving the last paragraph."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        element = registry.resolve_ref("p:2")

        text = registry._get_text_content(element)
        assert "Third paragraph" in text

    def test_resolve_out_of_bounds_raises(self) -> None:
        """Test that resolving out of bounds ref raises error."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        with pytest.raises(RefNotFoundError, match="out of bounds"):
            registry.resolve_ref("p:99")

    def test_resolve_table(self) -> None:
        """Test resolving a table."""
        registry = create_registry_from_xml(DOCUMENT_WITH_TABLE_XML)

        element = registry.resolve_ref("tbl:0")

        assert element is not None
        assert element.tag == f"{{{WORD_NAMESPACE}}}tbl"

    def test_resolve_table_row(self) -> None:
        """Test resolving a table row."""
        registry = create_registry_from_xml(DOCUMENT_WITH_TABLE_XML)

        element = registry.resolve_ref("tbl:0/row:1")

        assert element is not None
        assert element.tag == f"{{{WORD_NAMESPACE}}}tr"

    def test_resolve_table_cell(self) -> None:
        """Test resolving a table cell."""
        registry = create_registry_from_xml(DOCUMENT_WITH_TABLE_XML)

        element = registry.resolve_ref("tbl:0/row:0/cell:1")

        assert element is not None
        assert element.tag == f"{{{WORD_NAMESPACE}}}tc"
        text = registry._get_text_content(element)
        assert "Cell 0,1" in text


class TestRefRegistryGetRef:
    """Tests for getting refs from elements."""

    def test_get_ref_for_paragraph(self) -> None:
        """Test getting ref for a paragraph."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        # Get the second paragraph
        element = registry.resolve_ref("p:1")
        ref = registry.get_ref(element)

        assert ref.path == "p:1"

    def test_get_ref_roundtrip(self) -> None:
        """Test that get_ref and resolve_ref are inverses."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        # Resolve a ref
        element = registry.resolve_ref("p:2")

        # Get the ref back
        ref = registry.get_ref(element)

        # Resolve again
        element2 = registry.resolve_ref(ref)

        assert element is element2


class TestRefRegistryFingerprint:
    """Tests for fingerprint-based refs."""

    def test_compute_fingerprint_deterministic(self) -> None:
        """Test that fingerprint computation is deterministic."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        element = registry.resolve_ref("p:0")
        fp1 = registry._compute_fingerprint(element)
        fp2 = registry._compute_fingerprint(element)

        assert fp1 == fp2

    def test_fingerprints_differ_for_different_content(self) -> None:
        """Test that different content produces different fingerprints."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        elem1 = registry.resolve_ref("p:0")
        elem2 = registry.resolve_ref("p:1")

        fp1 = registry._compute_fingerprint(elem1)
        fp2 = registry._compute_fingerprint(elem2)

        assert fp1 != fp2

    def test_get_ref_with_fingerprint(self) -> None:
        """Test getting a fingerprint-based ref."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        element = registry.resolve_ref("p:1")
        ref = registry.get_ref(element, use_fingerprint=True)

        assert ref.is_fingerprint
        assert ref.path.startswith("p:~")

    def test_resolve_fingerprint_ref(self) -> None:
        """Test resolving a fingerprint-based ref."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        # Get a fingerprint ref
        element = registry.resolve_ref("p:1")
        ref = registry.get_ref(element, use_fingerprint=True)

        # Resolve it
        resolved = registry.resolve_ref(ref)

        assert resolved is element

    def test_fingerprint_ref_not_found_raises(self) -> None:
        """Test that non-existent fingerprint raises error."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        with pytest.raises(RefNotFoundError, match="No element found with fingerprint"):
            registry.resolve_ref("p:~nonexistent123")


class TestRefRegistryEnumeration:
    """Tests for enumerating refs."""

    def test_get_all_refs_paragraphs(self) -> None:
        """Test getting all paragraph refs."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        refs = registry.get_all_refs(ElementType.PARAGRAPH)

        assert len(refs) == 3
        assert refs[0].path == "p:0"
        assert refs[1].path == "p:1"
        assert refs[2].path == "p:2"

    def test_count_elements(self) -> None:
        """Test counting elements."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        count = registry.count_elements(ElementType.PARAGRAPH)

        assert count == 3

    def test_count_tables(self) -> None:
        """Test counting tables."""
        registry = create_registry_from_xml(DOCUMENT_WITH_TABLE_XML)

        count = registry.count_elements(ElementType.TABLE)

        assert count == 1


class TestRefRegistryCacheInvalidation:
    """Tests for cache invalidation."""

    def test_invalidate_clears_caches(self) -> None:
        """Test that invalidate clears all caches."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        # Access to populate caches
        registry.resolve_ref("p:0")
        registry.get_all_refs(ElementType.PARAGRAPH)

        # Invalidate
        registry.invalidate()

        # Caches should be cleared
        assert registry._ordinal_cache == {}
        assert registry._fingerprint_cache == {}


class TestRefRegistryIntegration:
    """Integration tests using the full Document class."""

    def test_registry_with_document(self) -> None:
        """Test using RefRegistry with a Document object."""
        docx_path = create_test_docx(MINIMAL_DOCUMENT_XML)

        try:
            doc = Document(docx_path)
            registry = RefRegistry(doc.xml_root)

            # Resolve refs
            p0 = registry.resolve_ref("p:0")
            p1 = registry.resolve_ref("p:1")
            p2 = registry.resolve_ref("p:2")

            # Verify text content
            assert "First" in registry._get_text_content(p0)
            assert "Second" in registry._get_text_content(p1)
            assert "Third" in registry._get_text_content(p2)

        finally:
            docx_path.unlink()

    def test_registry_with_table_document(self) -> None:
        """Test RefRegistry with a document containing tables."""
        docx_path = create_test_docx(DOCUMENT_WITH_TABLE_XML)

        try:
            doc = Document(docx_path)
            registry = RefRegistry(doc.xml_root)

            # Count elements
            assert registry.count_elements(ElementType.PARAGRAPH) == 2  # Body paragraphs only
            assert registry.count_elements(ElementType.TABLE) == 1

            # Resolve table and cell
            table = registry.resolve_ref("tbl:0")
            assert table is not None

            cell = registry.resolve_ref("tbl:0/row:1/cell:0")
            text = registry._get_text_content(cell)
            assert "Cell 1,0" in text

        finally:
            docx_path.unlink()


class TestRefNotFoundError:
    """Tests for RefNotFoundError."""

    def test_error_message_includes_ref(self) -> None:
        """Test that error message includes the ref path."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        with pytest.raises(RefNotFoundError) as exc_info:
            registry.resolve_ref("p:99")

        assert "p:99" in str(exc_info.value)

    def test_error_message_includes_reason(self) -> None:
        """Test that error message includes the reason."""
        registry = create_registry_from_xml(MINIMAL_DOCUMENT_XML)

        with pytest.raises(RefNotFoundError) as exc_info:
            registry.resolve_ref("p:99")

        assert "out of bounds" in str(exc_info.value)

    def test_unsupported_element_type_error(self) -> None:
        """Test error for unsupported element types."""
        # Directly test with an unsupported type
        # Note: We need to bypass parse validation which checks prefixes
        # Instead we test the error class directly
        error = RefNotFoundError("xyz:5", "Unknown element type")
        assert "xyz:5" in str(error)
        assert "Unknown element type" in str(error)


class TestStaleRefError:
    """Tests for StaleRefError."""

    def test_stale_ref_error_message(self) -> None:
        """Test StaleRefError message format."""
        from python_docx_redline.errors import StaleRefError

        error = StaleRefError("p:~abc123", "Element content has changed")

        msg = str(error)
        assert "p:~abc123" in msg
        assert "Element content has changed" in msg
        assert "Regenerate the accessibility tree" in msg
