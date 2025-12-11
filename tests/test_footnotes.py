"""Tests for footnote and endnote functionality."""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from python_docx_redline import Document

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def create_test_docx(content: str | None = None) -> Path:
    """Create a minimal but valid OOXML test .docx file."""
    if content is None:
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a test document with some text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Another paragraph for testing.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

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


class TestFootnoteBasic:
    """Basic footnote functionality tests."""

    def test_insert_footnote_basic(self):
        """Test basic footnote insertion."""
        doc = Document(create_test_docx())

        # Insert a footnote
        footnote_id = doc.insert_footnote("This is a footnote", at="test document")

        assert footnote_id == 1

        # Verify footnote exists in properties
        footnotes = doc.footnotes
        assert len(footnotes) == 1
        assert footnotes[0].text == "This is a footnote"
        assert footnotes[0].id == "1"

    def test_insert_multiple_footnotes(self):
        """Test inserting multiple footnotes."""
        doc = Document(create_test_docx())

        id1 = doc.insert_footnote("First footnote", at="test document")
        id2 = doc.insert_footnote("Second footnote", at="Another paragraph")

        assert id1 == 1
        assert id2 == 2

        footnotes = doc.footnotes
        assert len(footnotes) == 2
        assert footnotes[0].text == "First footnote"
        assert footnotes[1].text == "Second footnote"

    def test_footnote_reference_in_document(self):
        """Test that footnote reference is inserted in document."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Test footnote", at="test document")

        # Check for footnoteReference element in XML
        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "footnoteReference" in xml_str

    def test_footnotes_property_empty(self):
        """Test footnotes property when no footnotes exist."""
        doc = Document(create_test_docx())

        footnotes = doc.footnotes
        assert len(footnotes) == 0

    def test_footnote_with_author(self):
        """Test footnote insertion with custom author."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote text", at="test document", author="John Doe")

        # Footnote should be created successfully
        footnotes = doc.footnotes
        assert len(footnotes) == 1


class TestEndnoteBasic:
    """Basic endnote functionality tests."""

    def test_insert_endnote_basic(self):
        """Test basic endnote insertion."""
        doc = Document(create_test_docx())

        # Insert an endnote
        endnote_id = doc.insert_endnote("This is an endnote", at="test document")

        assert endnote_id == 1

        # Verify endnote exists in properties
        endnotes = doc.endnotes
        assert len(endnotes) == 1
        assert endnotes[0].text == "This is an endnote"
        assert endnotes[0].id == "1"

    def test_insert_multiple_endnotes(self):
        """Test inserting multiple endnotes."""
        doc = Document(create_test_docx())

        id1 = doc.insert_endnote("First endnote", at="test document")
        id2 = doc.insert_endnote("Second endnote", at="Another paragraph")

        assert id1 == 1
        assert id2 == 2

        endnotes = doc.endnotes
        assert len(endnotes) == 2
        assert endnotes[0].text == "First endnote"
        assert endnotes[1].text == "Second endnote"

    def test_endnote_reference_in_document(self):
        """Test that endnote reference is inserted in document."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Test endnote", at="test document")

        # Check for endnoteReference element in XML
        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "endnoteReference" in xml_str

    def test_endnotes_property_empty(self):
        """Test endnotes property when no endnotes exist."""
        doc = Document(create_test_docx())

        endnotes = doc.endnotes
        assert len(endnotes) == 0


class TestFootnoteModel:
    """Tests for Footnote model class."""

    def test_footnote_text_property(self):
        """Test footnote text property."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote content here", at="test document")

        footnote = doc.footnotes[0]
        assert footnote.text == "Footnote content here"

    def test_footnote_paragraphs_property(self):
        """Test footnote paragraphs property."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote text", at="test document")

        footnote = doc.footnotes[0]
        assert len(footnote.paragraphs) == 1
        assert footnote.paragraphs[0].text == "Footnote text"

    def test_footnote_contains(self):
        """Test footnote contains method."""
        doc = Document(create_test_docx())

        doc.insert_footnote("This is a test footnote", at="test document")

        footnote = doc.footnotes[0]
        assert footnote.contains("test")
        assert not footnote.contains("missing")

    def test_footnote_contains_case_insensitive(self):
        """Test footnote contains with case insensitive search."""
        doc = Document(create_test_docx())

        doc.insert_footnote("This is a TEST footnote", at="test document")

        footnote = doc.footnotes[0]
        assert footnote.contains("test", case_sensitive=False)
        assert not footnote.contains("test", case_sensitive=True)

    def test_footnote_repr(self):
        """Test footnote string representation."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Short note", at="test document")

        footnote = doc.footnotes[0]
        repr_str = repr(footnote)
        assert "Footnote" in repr_str
        assert "Short note" in repr_str


class TestEndnoteModel:
    """Tests for Endnote model class."""

    def test_endnote_text_property(self):
        """Test endnote text property."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Endnote content here", at="test document")

        endnote = doc.endnotes[0]
        assert endnote.text == "Endnote content here"

    def test_endnote_paragraphs_property(self):
        """Test endnote paragraphs property."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Endnote text", at="test document")

        endnote = doc.endnotes[0]
        assert len(endnote.paragraphs) == 1
        assert endnote.paragraphs[0].text == "Endnote text"

    def test_endnote_contains(self):
        """Test endnote contains method."""
        doc = Document(create_test_docx())

        doc.insert_endnote("This is a test endnote", at="test document")

        endnote = doc.endnotes[0]
        assert endnote.contains("test")
        assert not endnote.contains("missing")

    def test_endnote_repr(self):
        """Test endnote string representation."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Short note", at="test document")

        endnote = doc.endnotes[0]
        repr_str = repr(endnote)
        assert "Endnote" in repr_str
        assert "Short note" in repr_str


class TestFootnoteEndnoteMixed:
    """Tests for documents with both footnotes and endnotes."""

    def test_mixed_footnotes_and_endnotes(self):
        """Test document with both footnotes and endnotes."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote 1", at="test document")
        doc.insert_endnote("Endnote 1", at="test document")
        doc.insert_footnote("Footnote 2", at="Another paragraph")
        doc.insert_endnote("Endnote 2", at="Another paragraph")

        assert len(doc.footnotes) == 2
        assert len(doc.endnotes) == 2

    def test_footnote_endnote_ids_independent(self):
        """Test that footnote and endnote IDs are independent."""
        doc = Document(create_test_docx())

        footnote_id = doc.insert_footnote("Footnote", at="test document")
        endnote_id = doc.insert_endnote("Endnote", at="test document")

        # Both should start at 1
        assert footnote_id == 1
        assert endnote_id == 1


class TestFootnoteEndnotePersistence:
    """Tests for footnote/endnote persistence across save/load."""

    def test_footnote_persists_after_save(self):
        """Test that footnotes persist after save and reload."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Persistent footnote", at="test document")

        # Save and reload
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "footnote_test.docx"
            doc.save(output_path)

            reloaded_doc = Document(output_path)

            # Verify footnote persisted
            footnotes = reloaded_doc.footnotes
            assert len(footnotes) == 1
            assert footnotes[0].text == "Persistent footnote"

    def test_endnote_persists_after_save(self):
        """Test that endnotes persist after save and reload."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Persistent endnote", at="test document")

        # Save and reload
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "endnote_test.docx"
            doc.save(output_path)

            reloaded_doc = Document(output_path)

            # Verify endnote persisted
            endnotes = reloaded_doc.endnotes
            assert len(endnotes) == 1
            assert endnotes[0].text == "Persistent endnote"


class TestFootnoteEndnoteErrors:
    """Tests for error handling in footnote/endnote operations."""

    def test_insert_footnote_text_not_found(self):
        """Test footnote insertion with non-existent anchor text."""
        from python_docx_redline import TextNotFoundError

        doc = Document(create_test_docx())

        with pytest.raises(TextNotFoundError):
            doc.insert_footnote("Footnote", at="nonexistent text")

    def test_insert_endnote_text_not_found(self):
        """Test endnote insertion with non-existent anchor text."""
        from python_docx_redline import TextNotFoundError

        doc = Document(create_test_docx())

        with pytest.raises(TextNotFoundError):
            doc.insert_endnote("Endnote", at="nonexistent text")

    def test_insert_footnote_ambiguous_text(self):
        """Test footnote insertion with ambiguous anchor text."""
        from python_docx_redline import AmbiguousTextError

        # Create doc with duplicate text
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Same text</w:t></w:r></w:p>
    <w:p><w:r><w:t>Same text</w:t></w:r></w:p>
  </w:body>
</w:document>"""

        doc = Document(create_test_docx(content))

        with pytest.raises(AmbiguousTextError):
            doc.insert_footnote("Footnote", at="Same text")
