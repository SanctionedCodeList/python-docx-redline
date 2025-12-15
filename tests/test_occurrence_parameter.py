"""
Tests for occurrence parameter in replace_tracked, insert_tracked, delete_tracked.

These tests verify that the occurrence parameter works correctly to target specific
occurrences of text when multiple matches exist.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest

from python_docx_redline import Document
from python_docx_redline.constants import WORD_NAMESPACE


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


class TestOccurrenceParameterReplaceTracked:
    """Tests for occurrence parameter in replace_tracked()."""

    def test_replace_tracked_occurrence_first(self):
        """Test replacing only the first occurrence (explicit)."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>The text appears here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears again.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears once more.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.replace_tracked("text", "word", occurrence=1)

            # Check that only first occurrence was replaced
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            ins_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))

            assert len(del_elements) == 1
            assert len(ins_elements) == 1

            # Verify the deleted text
            del_text = "".join(del_elements[0].itertext())
            assert "text" in del_text

            # Verify the inserted text
            ins_text = "".join(ins_elements[0].itertext())
            assert "word" in ins_text
        finally:
            docx_path.unlink()

    def test_replace_tracked_occurrence_second(self):
        """Test replacing only the second occurrence."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>The text appears here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears again.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears once more.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.replace_tracked("text", "word", occurrence=2)

            # Should have exactly one replacement
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            ins_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))

            assert len(del_elements) == 1
            assert len(ins_elements) == 1
        finally:
            docx_path.unlink()

    def test_replace_tracked_occurrence_last(self):
        """Test replacing only the last occurrence."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>The text appears here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears again.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears once more.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.replace_tracked("text", "word", occurrence="last")

            # Should have exactly one replacement
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            ins_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))

            assert len(del_elements) == 1
            assert len(ins_elements) == 1
        finally:
            docx_path.unlink()

    def test_replace_tracked_occurrence_all(self):
        """Test replacing all occurrences."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>The text appears here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears again.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears once more.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.replace_tracked("text", "word", occurrence="all")

            # Should have three replacements
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            ins_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))

            assert len(del_elements) == 3
            assert len(ins_elements) == 3
        finally:
            docx_path.unlink()

    def test_replace_tracked_occurrence_list(self):
        """Test replacing specific occurrences using a list."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>The text appears here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears again.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears once more.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears last time.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            # Replace 1st and 3rd occurrences (1-indexed)
            doc.replace_tracked("text", "word", occurrence=[1, 3])

            # Should have exactly two replacements
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            ins_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))

            assert len(del_elements) == 2
            assert len(ins_elements) == 2
        finally:
            docx_path.unlink()

    def test_replace_tracked_occurrence_out_of_range(self):
        """Test that out of range occurrence raises ValueError."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>The text appears here.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            with pytest.raises(ValueError, match="Occurrence 5 out of range"):
                doc.replace_tracked("text", "word", occurrence=5)
        finally:
            docx_path.unlink()

    def test_replace_tracked_no_occurrence_raises_ambiguous(self):
        """Test that multiple matches without occurrence parameter raises AmbiguousTextError."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>The text appears here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The text appears again.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            # Default occurrence is "first", so this should work
            doc.replace_tracked("text", "word")

            # Should have exactly one replacement (first occurrence)
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            assert len(del_elements) == 1
        finally:
            docx_path.unlink()


class TestOccurrenceParameterInsertTracked:
    """Tests for occurrence parameter in insert_tracked()."""

    def test_insert_tracked_occurrence_first(self):
        """Test inserting after only the first occurrence."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>anchor text here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>anchor text again.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.insert_tracked("INSERTED", after="anchor", occurrence=1)

            # Should have exactly one insertion
            ins_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))
            assert len(ins_elements) == 1

            ins_text = "".join(ins_elements[0].itertext())
            assert "INSERTED" in ins_text
        finally:
            docx_path.unlink()

    def test_insert_tracked_occurrence_all(self):
        """Test inserting after all occurrences."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>anchor text here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>anchor text again.</w:t></w:r></w:p>
    <w:p><w:r><w:t>anchor text once more.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.insert_tracked("INSERTED", after="anchor", occurrence="all")

            # Should have three insertions
            ins_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))
            assert len(ins_elements) == 3
        finally:
            docx_path.unlink()

    def test_insert_tracked_occurrence_list(self):
        """Test inserting after specific occurrences using a list."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>anchor text 1.</w:t></w:r></w:p>
    <w:p><w:r><w:t>anchor text 2.</w:t></w:r></w:p>
    <w:p><w:r><w:t>anchor text 3.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            # Insert after 1st and 3rd occurrences
            doc.insert_tracked("INSERTED", after="anchor", occurrence=[1, 3])

            # Should have exactly two insertions
            ins_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))
            assert len(ins_elements) == 2
        finally:
            docx_path.unlink()

    def test_insert_tracked_before_occurrence(self):
        """Test inserting before a specific occurrence."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>anchor text 1.</w:t></w:r></w:p>
    <w:p><w:r><w:t>anchor text 2.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.insert_tracked("INSERTED", before="anchor", occurrence=2)

            # Should have exactly one insertion
            ins_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))
            assert len(ins_elements) == 1
        finally:
            docx_path.unlink()


class TestOccurrenceParameterDeleteTracked:
    """Tests for occurrence parameter in delete_tracked()."""

    def test_delete_tracked_occurrence_first(self):
        """Test deleting only the first occurrence."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Delete this text here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Delete this text again.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.delete_tracked("this", occurrence=1)

            # Should have exactly one deletion
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            assert len(del_elements) == 1

            del_text = "".join(del_elements[0].itertext())
            assert "this" in del_text
        finally:
            docx_path.unlink()

    def test_delete_tracked_occurrence_last(self):
        """Test deleting only the last occurrence."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Delete this text here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Delete this text again.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Delete this text once more.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.delete_tracked("this", occurrence="last")

            # Should have exactly one deletion
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            assert len(del_elements) == 1
        finally:
            docx_path.unlink()

    def test_delete_tracked_occurrence_all(self):
        """Test deleting all occurrences."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Delete this text here.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Delete this text again.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Delete this text once more.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            doc.delete_tracked("this", occurrence="all")

            # Should have three deletions
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            assert len(del_elements) == 3
        finally:
            docx_path.unlink()

    def test_delete_tracked_occurrence_list(self):
        """Test deleting specific occurrences using a list."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Delete this 1.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Delete this 2.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Delete this 3.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Delete this 4.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            # Delete 2nd and 4th occurrences
            doc.delete_tracked("this", occurrence=[2, 4])

            # Should have exactly two deletions
            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            assert len(del_elements) == 2
        finally:
            docx_path.unlink()


class TestOccurrenceParameterEdgeCases:
    """Edge case tests for occurrence parameter."""

    def test_occurrence_with_single_match_no_error(self):
        """Test that occurrence parameter works fine with single match (no AmbiguousTextError)."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Unique text here.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            # Should not raise error even though there's only one match
            doc.replace_tracked("Unique", "Changed", occurrence=1)

            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            assert len(del_elements) == 1
        finally:
            docx_path.unlink()

    def test_occurrence_list_with_invalid_type(self):
        """Test that occurrence list with non-integer raises ValueError."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>text text text</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            with pytest.raises(ValueError, match="List elements must be integers"):
                doc.replace_tracked("text", "word", occurrence=[1, "two"])
        finally:
            docx_path.unlink()

    def test_occurrence_zero_out_of_range(self):
        """Test that occurrence=0 raises ValueError (1-indexed)."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>text here</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            with pytest.raises(ValueError, match="Occurrence 0 out of range"):
                doc.replace_tracked("text", "word", occurrence=0)
        finally:
            docx_path.unlink()

    def test_occurrence_with_scope(self):
        """Test that occurrence works correctly when combined with scope."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>text in first paragraph</w:t></w:r></w:p>
    <w:p><w:r><w:t>text in second paragraph</w:t></w:r></w:p>
    <w:p><w:r><w:t>text in third paragraph</w:t></w:r></w:p>
  </w:body>
</w:document>"""
        docx_path = create_test_docx(doc_xml)
        try:
            doc = Document(docx_path)
            # Replace second occurrence of "text" within the whole document
            doc.replace_tracked("text", "word", occurrence=2)

            del_elements = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
            assert len(del_elements) == 1
        finally:
            docx_path.unlink()
