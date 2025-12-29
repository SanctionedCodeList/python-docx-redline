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


class TestGetFootnoteEndnote:
    """Tests for get_footnote and get_endnote methods."""

    def test_get_footnote_by_id(self):
        """Test retrieving a footnote by its ID."""
        doc = Document(create_test_docx())

        doc.insert_footnote("First footnote", at="test document")
        doc.insert_footnote("Second footnote", at="Another paragraph")

        # Get by ID
        footnote = doc.get_footnote(1)
        assert footnote.text == "First footnote"
        assert footnote.id == "1"

        footnote2 = doc.get_footnote(2)
        assert footnote2.text == "Second footnote"
        assert footnote2.id == "2"

    def test_get_footnote_by_string_id(self):
        """Test retrieving a footnote by string ID."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Test footnote", at="test document")

        # Get by string ID
        footnote = doc.get_footnote("1")
        assert footnote.text == "Test footnote"

    def test_get_footnote_not_found(self):
        """Test NoteNotFoundError when footnote doesn't exist."""
        from python_docx_redline import NoteNotFoundError

        doc = Document(create_test_docx())
        doc.insert_footnote("Only footnote", at="test document")

        with pytest.raises(NoteNotFoundError) as exc_info:
            doc.get_footnote(99)

        assert "99" in str(exc_info.value)
        assert "footnote" in str(exc_info.value).lower()

    def test_get_footnote_empty_document(self):
        """Test NoteNotFoundError on document with no footnotes."""
        from python_docx_redline import NoteNotFoundError

        doc = Document(create_test_docx())

        with pytest.raises(NoteNotFoundError) as exc_info:
            doc.get_footnote(1)

        assert "No footnotes exist" in str(exc_info.value)

    def test_get_endnote_by_id(self):
        """Test retrieving an endnote by its ID."""
        doc = Document(create_test_docx())

        doc.insert_endnote("First endnote", at="test document")
        doc.insert_endnote("Second endnote", at="Another paragraph")

        endnote = doc.get_endnote(1)
        assert endnote.text == "First endnote"
        assert endnote.id == "1"

        endnote2 = doc.get_endnote(2)
        assert endnote2.text == "Second endnote"

    def test_get_endnote_not_found(self):
        """Test NoteNotFoundError when endnote doesn't exist."""
        from python_docx_redline import NoteNotFoundError

        doc = Document(create_test_docx())
        doc.insert_endnote("Only endnote", at="test document")

        with pytest.raises(NoteNotFoundError) as exc_info:
            doc.get_endnote(99)

        assert "99" in str(exc_info.value)
        assert "endnote" in str(exc_info.value).lower()


class TestDeleteFootnoteEndnote:
    """Tests for delete_footnote and delete_endnote methods."""

    def test_delete_footnote_basic(self):
        """Test basic footnote deletion."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote to delete", at="test document")
        assert len(doc.footnotes) == 1

        doc.delete_footnote(1)
        assert len(doc.footnotes) == 0

    def test_delete_footnote_removes_reference(self):
        """Test that footnote reference is removed from document."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Test footnote", at="test document")

        # Verify reference exists
        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "footnoteReference" in xml_str

        doc.delete_footnote(1)

        # Verify reference is removed
        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "footnoteReference" not in xml_str

    def test_delete_footnote_with_renumbering(self):
        """Test that footnotes are renumbered after deletion."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote 1", at="test document")
        doc.insert_footnote("Footnote 2", at="Another paragraph")

        # Delete first footnote
        doc.delete_footnote(1)

        # Remaining footnote should be renumbered to 1
        footnotes = doc.footnotes
        assert len(footnotes) == 1
        assert footnotes[0].id == "1"
        assert footnotes[0].text == "Footnote 2"

    def test_delete_footnote_without_renumbering(self):
        """Test deletion without renumbering."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote 1", at="test document")
        doc.insert_footnote("Footnote 2", at="Another paragraph")

        # Delete first footnote without renumbering
        doc.delete_footnote(1, renumber=False)

        # Remaining footnote should keep its original ID
        footnotes = doc.footnotes
        assert len(footnotes) == 1
        assert footnotes[0].id == "2"

    def test_delete_footnote_not_found(self):
        """Test NoteNotFoundError when deleting non-existent footnote."""
        from python_docx_redline import NoteNotFoundError

        doc = Document(create_test_docx())
        doc.insert_footnote("Only footnote", at="test document")

        with pytest.raises(NoteNotFoundError):
            doc.delete_footnote(99)

    def test_delete_middle_footnote_renumbers_correctly(self):
        """Test renumbering when middle footnote is deleted."""
        doc = Document(create_test_docx())

        # Create document with 3 unique paragraphs for 3 footnotes
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Third paragraph.</w:t></w:r></w:p>
  </w:body>
</w:document>"""

        doc = Document(create_test_docx(content))

        doc.insert_footnote("Footnote A", at="First paragraph")
        doc.insert_footnote("Footnote B", at="Second paragraph")
        doc.insert_footnote("Footnote C", at="Third paragraph")

        # Delete middle footnote
        doc.delete_footnote(2)

        # Verify renumbering
        footnotes = doc.footnotes
        assert len(footnotes) == 2
        assert footnotes[0].id == "1"
        assert footnotes[0].text == "Footnote A"
        assert footnotes[1].id == "2"
        assert footnotes[1].text == "Footnote C"

    def test_delete_endnote_basic(self):
        """Test basic endnote deletion."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Endnote to delete", at="test document")
        assert len(doc.endnotes) == 1

        doc.delete_endnote(1)
        assert len(doc.endnotes) == 0

    def test_delete_endnote_removes_reference(self):
        """Test that endnote reference is removed from document."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Test endnote", at="test document")

        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "endnoteReference" in xml_str

        doc.delete_endnote(1)

        xml_str = etree.tostring(doc.xml_root, encoding="unicode")
        assert "endnoteReference" not in xml_str

    def test_delete_endnote_with_renumbering(self):
        """Test that endnotes are renumbered after deletion."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Endnote 1", at="test document")
        doc.insert_endnote("Endnote 2", at="Another paragraph")

        doc.delete_endnote(1)

        endnotes = doc.endnotes
        assert len(endnotes) == 1
        assert endnotes[0].id == "1"
        assert endnotes[0].text == "Endnote 2"

    def test_delete_endnote_not_found(self):
        """Test NoteNotFoundError when deleting non-existent endnote."""
        from python_docx_redline import NoteNotFoundError

        doc = Document(create_test_docx())
        doc.insert_endnote("Only endnote", at="test document")

        with pytest.raises(NoteNotFoundError):
            doc.delete_endnote(99)


class TestEditFootnoteEndnote:
    """Tests for edit_footnote and edit_endnote methods."""

    def test_edit_footnote_basic(self):
        """Test basic footnote text editing."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Original text", at="test document")
        assert doc.footnotes[0].text == "Original text"

        doc.edit_footnote(1, "Updated text")

        # Need to re-fetch the footnote after edit
        footnotes = doc.footnotes
        assert len(footnotes) == 1
        assert footnotes[0].text == "Updated text"

    def test_edit_footnote_preserves_id(self):
        """Test that editing preserves the footnote ID."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Original", at="test document")
        doc.edit_footnote(1, "Edited")

        footnote = doc.get_footnote(1)
        assert footnote.id == "1"
        assert footnote.text == "Edited"

    def test_edit_footnote_not_found(self):
        """Test NoteNotFoundError when editing non-existent footnote."""
        from python_docx_redline import NoteNotFoundError

        doc = Document(create_test_docx())
        doc.insert_footnote("Only footnote", at="test document")

        with pytest.raises(NoteNotFoundError):
            doc.edit_footnote(99, "New text")

    def test_edit_footnote_persists_after_save(self):
        """Test that edited footnote persists after save/reload."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Original text", at="test document")
        doc.edit_footnote(1, "Edited text")

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "edited_footnote.docx"
            doc.save(output_path)

            reloaded = Document(output_path)
            footnotes = reloaded.footnotes
            assert len(footnotes) == 1
            assert footnotes[0].text == "Edited text"

    def test_edit_endnote_basic(self):
        """Test basic endnote text editing."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Original text", at="test document")
        assert doc.endnotes[0].text == "Original text"

        doc.edit_endnote(1, "Updated text")

        endnotes = doc.endnotes
        assert len(endnotes) == 1
        assert endnotes[0].text == "Updated text"

    def test_edit_endnote_preserves_id(self):
        """Test that editing preserves the endnote ID."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Original", at="test document")
        doc.edit_endnote(1, "Edited")

        endnote = doc.get_endnote(1)
        assert endnote.id == "1"
        assert endnote.text == "Edited"

    def test_edit_endnote_not_found(self):
        """Test NoteNotFoundError when editing non-existent endnote."""
        from python_docx_redline import NoteNotFoundError

        doc = Document(create_test_docx())
        doc.insert_endnote("Only endnote", at="test document")

        with pytest.raises(NoteNotFoundError):
            doc.edit_endnote(99, "New text")

    def test_edit_endnote_persists_after_save(self):
        """Test that edited endnote persists after save/reload."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Original text", at="test document")
        doc.edit_endnote(1, "Edited text")

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "edited_endnote.docx"
            doc.save(output_path)

            reloaded = Document(output_path)
            endnotes = reloaded.endnotes
            assert len(endnotes) == 1
            assert endnotes[0].text == "Edited text"


class TestNoteNotFoundError:
    """Tests for NoteNotFoundError exception."""

    def test_error_includes_note_type(self):
        """Test that error message includes note type."""
        from python_docx_redline import NoteNotFoundError

        error = NoteNotFoundError("footnote", 5, ["1", "2", "3"])
        assert "Footnote" in str(error)

        error = NoteNotFoundError("endnote", 5)
        assert "Endnote" in str(error)

    def test_error_includes_searched_id(self):
        """Test that error message includes the ID that was searched."""
        from python_docx_redline import NoteNotFoundError

        error = NoteNotFoundError("footnote", 42)
        assert "42" in str(error)

    def test_error_includes_available_ids(self):
        """Test that error message includes available IDs."""
        from python_docx_redline import NoteNotFoundError

        error = NoteNotFoundError("footnote", 5, ["1", "2", "3"])
        msg = str(error)
        assert "1" in msg
        assert "2" in msg
        assert "3" in msg

    def test_error_when_no_notes_exist(self):
        """Test error message when no notes exist."""
        from python_docx_redline import NoteNotFoundError

        error = NoteNotFoundError("footnote", 1, [])
        assert "No footnotes exist" in str(error)

    def test_error_attributes(self):
        """Test error attributes are accessible."""
        from python_docx_redline import NoteNotFoundError

        error = NoteNotFoundError("endnote", 5, ["1", "2"])
        assert error.note_type == "endnote"
        assert error.note_id == "5"
        assert error.available_ids == ["1", "2"]


class TestDeletePersistence:
    """Tests for footnote/endnote deletion persistence after save."""

    def test_deleted_footnote_not_in_saved_document(self):
        """Test that deleted footnote doesn't appear after save/reload."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote 1", at="test document")
        doc.insert_footnote("Footnote 2", at="Another paragraph")
        doc.delete_footnote(1)

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "deleted_footnote.docx"
            doc.save(output_path)

            reloaded = Document(output_path)
            footnotes = reloaded.footnotes
            assert len(footnotes) == 1
            assert footnotes[0].text == "Footnote 2"

    def test_deleted_endnote_not_in_saved_document(self):
        """Test that deleted endnote doesn't appear after save/reload."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Endnote 1", at="test document")
        doc.insert_endnote("Endnote 2", at="Another paragraph")
        doc.delete_endnote(1)

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "deleted_endnote.docx"
            doc.save(output_path)

            reloaded = Document(output_path)
            endnotes = reloaded.endnotes
            assert len(endnotes) == 1
            assert endnotes[0].text == "Endnote 2"


class TestFootnoteReferenceLocation:
    """Tests for footnote/endnote reference location functionality."""

    def test_footnote_reference_location_basic(self):
        """Test getting reference location for a footnote."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Test footnote", at="test document")

        footnote = doc.get_footnote(1)
        ref_loc = footnote.reference_location

        assert ref_loc is not None
        assert ref_loc.paragraph is not None
        assert "test document" in ref_loc.paragraph.text
        assert ref_loc.run_element is not None
        assert ref_loc.position_in_paragraph >= 0

    def test_endnote_reference_location_basic(self):
        """Test getting reference location for an endnote."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Test endnote", at="test document")

        endnote = doc.get_endnote(1)
        ref_loc = endnote.reference_location

        assert ref_loc is not None
        assert ref_loc.paragraph is not None
        assert "test document" in ref_loc.paragraph.text
        assert ref_loc.run_element is not None

    def test_footnote_reference_location_via_note_ops(self):
        """Test getting reference location via NoteOperations directly."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Test footnote", at="test document")

        ref_loc = doc._note_ops.get_footnote_reference_location(1)

        assert ref_loc is not None
        assert "test document" in ref_loc.paragraph.text

    def test_endnote_reference_location_via_note_ops(self):
        """Test getting reference location via NoteOperations directly."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Test endnote", at="Another paragraph")

        ref_loc = doc._note_ops.get_endnote_reference_location(1)

        assert ref_loc is not None
        assert "Another paragraph" in ref_loc.paragraph.text

    def test_reference_location_position_calculation(self):
        """Test that position_in_paragraph is calculated correctly."""
        # Create doc with specific text layout
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello world</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        doc = Document(create_test_docx(content))
        doc.insert_footnote("Test footnote", at="world")

        footnote = doc.get_footnote(1)
        ref_loc = footnote.reference_location

        assert ref_loc is not None
        # The reference is inserted after "world" which is at position 6-11
        # in "Hello world", so position should be after existing text
        assert ref_loc.position_in_paragraph >= 0

    def test_reference_location_not_found(self):
        """Test reference_location returns None when document ref is None."""
        from python_docx_redline.models.footnote import Footnote

        # Create a minimal element
        elem = etree.Element(f"{{{WORD_NAMESPACE}}}footnote")
        elem.set(f"{{{WORD_NAMESPACE}}}id", "1")

        # Create footnote with None document
        footnote = Footnote(elem, None)  # type: ignore

        # Should return None when document is None
        assert footnote.reference_location is None


class TestFootnoteModelMethods:
    """Tests for Footnote model edit() and delete() methods."""

    def test_footnote_edit_method(self):
        """Test footnote.edit() method delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Original text", at="test document")
        footnote = doc.get_footnote(1)

        # Edit via model method
        footnote.edit("Updated text")

        # Verify change
        updated_footnote = doc.get_footnote(1)
        assert updated_footnote.text == "Updated text"

    def test_footnote_delete_method(self):
        """Test footnote.delete() method delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote to delete", at="test document")
        footnote = doc.get_footnote(1)

        # Delete via model method
        footnote.delete()

        # Verify deletion
        assert len(doc.footnotes) == 0

    def test_footnote_delete_with_renumber_false(self):
        """Test footnote.delete() with renumber=False."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote 1", at="test document")
        doc.insert_footnote("Footnote 2", at="Another paragraph")

        footnote = doc.get_footnote(1)
        footnote.delete(renumber=False)

        # Remaining footnote should keep original ID
        footnotes = doc.footnotes
        assert len(footnotes) == 1
        assert footnotes[0].id == "2"

    def test_footnote_edit_no_document_raises(self):
        """Test footnote.edit() raises ValueError when no document reference."""
        from python_docx_redline.models.footnote import Footnote

        elem = etree.Element(f"{{{WORD_NAMESPACE}}}footnote")
        elem.set(f"{{{WORD_NAMESPACE}}}id", "1")

        footnote = Footnote(elem, None)  # type: ignore

        with pytest.raises(ValueError) as exc_info:
            footnote.edit("New text")

        assert "no document reference" in str(exc_info.value)

    def test_footnote_delete_no_document_raises(self):
        """Test footnote.delete() raises ValueError when no document reference."""
        from python_docx_redline.models.footnote import Footnote

        elem = etree.Element(f"{{{WORD_NAMESPACE}}}footnote")
        elem.set(f"{{{WORD_NAMESPACE}}}id", "1")

        footnote = Footnote(elem, None)  # type: ignore

        with pytest.raises(ValueError) as exc_info:
            footnote.delete()

        assert "no document reference" in str(exc_info.value)


class TestEndnoteModelMethods:
    """Tests for Endnote model edit() and delete() methods."""

    def test_endnote_edit_method(self):
        """Test endnote.edit() method delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Original text", at="test document")
        endnote = doc.get_endnote(1)

        # Edit via model method
        endnote.edit("Updated text")

        # Verify change
        updated_endnote = doc.get_endnote(1)
        assert updated_endnote.text == "Updated text"

    def test_endnote_delete_method(self):
        """Test endnote.delete() method delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Endnote to delete", at="test document")
        endnote = doc.get_endnote(1)

        # Delete via model method
        endnote.delete()

        # Verify deletion
        assert len(doc.endnotes) == 0

    def test_endnote_delete_with_renumber_false(self):
        """Test endnote.delete() with renumber=False."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Endnote 1", at="test document")
        doc.insert_endnote("Endnote 2", at="Another paragraph")

        endnote = doc.get_endnote(1)
        endnote.delete(renumber=False)

        # Remaining endnote should keep original ID
        endnotes = doc.endnotes
        assert len(endnotes) == 1
        assert endnotes[0].id == "2"

    def test_endnote_edit_no_document_raises(self):
        """Test endnote.edit() raises ValueError when no document reference."""
        from python_docx_redline.models.footnote import Endnote

        elem = etree.Element(f"{{{WORD_NAMESPACE}}}endnote")
        elem.set(f"{{{WORD_NAMESPACE}}}id", "1")

        endnote = Endnote(elem, None)  # type: ignore

        with pytest.raises(ValueError) as exc_info:
            endnote.edit("New text")

        assert "no document reference" in str(exc_info.value)

    def test_endnote_delete_no_document_raises(self):
        """Test endnote.delete() raises ValueError when no document reference."""
        from python_docx_redline.models.footnote import Endnote

        elem = etree.Element(f"{{{WORD_NAMESPACE}}}endnote")
        elem.set(f"{{{WORD_NAMESPACE}}}id", "1")

        endnote = Endnote(elem, None)  # type: ignore

        with pytest.raises(ValueError) as exc_info:
            endnote.delete()

        assert "no document reference" in str(exc_info.value)


class TestFootnoteReferenceDataclass:
    """Tests for FootnoteReference dataclass."""

    def test_footnote_reference_dataclass_creation(self):
        """Test FootnoteReference can be created with expected fields."""
        from python_docx_redline.models.footnote import FootnoteReference
        from python_docx_redline.models.paragraph import Paragraph

        # Create a mock paragraph element
        para_elem = etree.Element(f"{{{WORD_NAMESPACE}}}p")
        run_elem = etree.SubElement(para_elem, f"{{{WORD_NAMESPACE}}}r")
        text_elem = etree.SubElement(run_elem, f"{{{WORD_NAMESPACE}}}t")
        text_elem.text = "Test text"

        paragraph = Paragraph(para_elem)

        ref = FootnoteReference(paragraph=paragraph, run_element=run_elem, position_in_paragraph=5)

        assert ref.paragraph is paragraph
        assert ref.run_element is run_elem
        assert ref.position_in_paragraph == 5


class TestMultiParagraphFootnotes:
    """Tests for multi-paragraph footnote and endnote support."""

    def test_insert_footnote_with_list(self):
        """Test inserting footnote with multiple paragraphs as list."""
        doc = Document(create_test_docx())

        footnote_id = doc.insert_footnote(
            ["First paragraph.", "Second paragraph."],
            at="test document",
        )

        assert footnote_id == 1

        footnote = doc.get_footnote(1)
        paragraphs = footnote.paragraphs
        assert len(paragraphs) == 2
        assert paragraphs[0].text == "First paragraph."
        assert paragraphs[1].text == "Second paragraph."

    def test_insert_endnote_with_list(self):
        """Test inserting endnote with multiple paragraphs as list."""
        doc = Document(create_test_docx())

        endnote_id = doc.insert_endnote(
            ["Para one.", "Para two.", "Para three."],
            at="test document",
        )

        assert endnote_id == 1

        endnote = doc.get_endnote(1)
        paragraphs = endnote.paragraphs
        assert len(paragraphs) == 3
        assert paragraphs[0].text == "Para one."
        assert paragraphs[1].text == "Para two."
        assert paragraphs[2].text == "Para three."

    def test_multi_paragraph_text_property(self):
        """Test that text property joins paragraphs with newlines."""
        doc = Document(create_test_docx())

        doc.insert_footnote(
            ["First.", "Second."],
            at="test document",
        )

        footnote = doc.get_footnote(1)
        assert footnote.text == "First.\nSecond."

    def test_multi_paragraph_persists_after_save(self):
        """Test multi-paragraph footnotes persist after save/reload."""
        doc = Document(create_test_docx())

        doc.insert_footnote(
            ["Paragraph A.", "Paragraph B."],
            at="test document",
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "multi_para.docx"
            doc.save(output_path)

            reloaded = Document(output_path)
            footnote = reloaded.get_footnote(1)
            assert len(footnote.paragraphs) == 2
            assert footnote.paragraphs[0].text == "Paragraph A."
            assert footnote.paragraphs[1].text == "Paragraph B."


class TestMarkdownFormatting:
    """Tests for markdown formatting in footnotes and endnotes."""

    def test_footnote_bold_formatting(self):
        """Test inserting footnote with bold markdown."""
        doc = Document(create_test_docx())

        doc.insert_footnote("This is **bold** text.", at="test document")

        footnote = doc.get_footnote(1)
        formatted = footnote.formatted_text

        # Should have 3 runs: "This is ", "bold", " text."
        assert len(formatted) == 3
        assert formatted[0]["text"] == "This is "
        assert formatted[0]["bold"] is False

        assert formatted[1]["text"] == "bold"
        assert formatted[1]["bold"] is True

        assert formatted[2]["text"] == " text."
        assert formatted[2]["bold"] is False

    def test_footnote_italic_formatting(self):
        """Test inserting footnote with italic markdown."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Some *italic* words.", at="test document")

        footnote = doc.get_footnote(1)
        formatted = footnote.formatted_text

        assert len(formatted) == 3
        assert formatted[1]["text"] == "italic"
        assert formatted[1]["italic"] is True

    def test_footnote_underline_formatting(self):
        """Test inserting footnote with underline markdown."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Some ++underlined++ text.", at="test document")

        footnote = doc.get_footnote(1)
        formatted = footnote.formatted_text

        assert len(formatted) == 3
        assert formatted[1]["text"] == "underlined"
        assert formatted[1]["underline"] is True

    def test_footnote_strikethrough_formatting(self):
        """Test inserting footnote with strikethrough markdown."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Some ~~struck~~ text.", at="test document")

        footnote = doc.get_footnote(1)
        formatted = footnote.formatted_text

        assert len(formatted) == 3
        assert formatted[1]["text"] == "struck"
        assert formatted[1]["strikethrough"] is True

    def test_footnote_mixed_formatting(self):
        """Test inserting footnote with multiple formatting types."""
        doc = Document(create_test_docx())

        doc.insert_footnote("**Bold** and *italic* text.", at="test document")

        footnote = doc.get_footnote(1)
        formatted = footnote.formatted_text

        # Find bold and italic runs
        bold_runs = [r for r in formatted if r["bold"]]
        italic_runs = [r for r in formatted if r["italic"]]

        assert len(bold_runs) == 1
        assert bold_runs[0]["text"] == "Bold"

        assert len(italic_runs) == 1
        assert italic_runs[0]["text"] == "italic"

    def test_endnote_markdown_formatting(self):
        """Test markdown formatting also works in endnotes."""
        doc = Document(create_test_docx())

        doc.insert_endnote("This is **bold** in endnote.", at="test document")

        endnote = doc.get_endnote(1)
        formatted = endnote.formatted_text

        bold_runs = [r for r in formatted if r["bold"]]
        assert len(bold_runs) == 1
        assert bold_runs[0]["text"] == "bold"

    def test_markdown_with_multi_paragraph(self):
        """Test markdown formatting works with multi-paragraph notes."""
        doc = Document(create_test_docx())

        doc.insert_footnote(
            ["First with **bold**.", "Second with *italic*."],
            at="test document",
        )

        footnote = doc.get_footnote(1)
        formatted = footnote.formatted_text

        # Check paragraph indices
        para_0_runs = [r for r in formatted if r["paragraph_index"] == 0]
        para_1_runs = [r for r in formatted if r["paragraph_index"] == 1]

        # First paragraph has bold
        bold_in_p0 = [r for r in para_0_runs if r["bold"]]
        assert len(bold_in_p0) == 1
        assert bold_in_p0[0]["text"] == "bold"

        # Second paragraph has italic
        italic_in_p1 = [r for r in para_1_runs if r["italic"]]
        assert len(italic_in_p1) == 1
        assert italic_in_p1[0]["text"] == "italic"

    def test_markdown_persists_after_save(self):
        """Test that markdown formatting persists after save/reload."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Some **bold** text.", at="test document")

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "formatted.docx"
            doc.save(output_path)

            reloaded = Document(output_path)
            footnote = reloaded.get_footnote(1)
            formatted = footnote.formatted_text

            bold_runs = [r for r in formatted if r["bold"]]
            assert len(bold_runs) == 1
            assert bold_runs[0]["text"] == "bold"


class TestFormattedTextProperty:
    """Tests for formatted_text property on Footnote and Endnote."""

    def test_formatted_text_returns_list(self):
        """Test formatted_text property returns list of dicts."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Simple text", at="test document")

        footnote = doc.get_footnote(1)
        formatted = footnote.formatted_text

        assert isinstance(formatted, list)
        assert len(formatted) == 1
        assert isinstance(formatted[0], dict)

    def test_formatted_text_dict_keys(self):
        """Test formatted_text dict has expected keys."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Text content", at="test document")

        footnote = doc.get_footnote(1)
        formatted = footnote.formatted_text[0]

        assert "text" in formatted
        assert "bold" in formatted
        assert "italic" in formatted
        assert "underline" in formatted
        assert "strikethrough" in formatted
        assert "paragraph_index" in formatted

    def test_formatted_text_paragraph_index(self):
        """Test paragraph_index is correct for multi-paragraph notes."""
        doc = Document(create_test_docx())

        doc.insert_footnote(["Para 0", "Para 1"], at="test document")

        footnote = doc.get_footnote(1)
        formatted = footnote.formatted_text

        assert formatted[0]["paragraph_index"] == 0
        assert formatted[1]["paragraph_index"] == 1

    def test_endnote_formatted_text(self):
        """Test formatted_text property works on endnotes too."""
        doc = Document(create_test_docx())

        doc.insert_endnote("**Bold** endnote", at="test document")

        endnote = doc.get_endnote(1)
        formatted = endnote.formatted_text

        assert len(formatted) >= 1
        bold_runs = [r for r in formatted if r["bold"]]
        assert len(bold_runs) == 1


class TestHtmlProperty:
    """Tests for html property on Footnote and Endnote."""

    def test_html_returns_string(self):
        """Test html property returns string."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Simple text", at="test document")

        footnote = doc.get_footnote(1)
        html = footnote.html

        assert isinstance(html, str)

    def test_html_wraps_paragraph(self):
        """Test html wraps content in <p> tags."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Test content", at="test document")

        footnote = doc.get_footnote(1)
        html = footnote.html

        assert html.startswith("<p>")
        assert html.endswith("</p>")
        assert "Test content" in html

    def test_html_bold_formatting(self):
        """Test html uses <b> for bold text."""
        doc = Document(create_test_docx())

        doc.insert_footnote("**Bold** text", at="test document")

        footnote = doc.get_footnote(1)
        html = footnote.html

        assert "<b>Bold</b>" in html

    def test_html_italic_formatting(self):
        """Test html uses <i> for italic text."""
        doc = Document(create_test_docx())

        doc.insert_footnote("*Italic* text", at="test document")

        footnote = doc.get_footnote(1)
        html = footnote.html

        assert "<i>Italic</i>" in html

    def test_html_underline_formatting(self):
        """Test html uses <u> for underlined text."""
        doc = Document(create_test_docx())

        doc.insert_footnote("++Underline++ text", at="test document")

        footnote = doc.get_footnote(1)
        html = footnote.html

        assert "<u>Underline</u>" in html

    def test_html_strikethrough_formatting(self):
        """Test html uses <s> for strikethrough text."""
        doc = Document(create_test_docx())

        doc.insert_footnote("~~Strike~~ text", at="test document")

        footnote = doc.get_footnote(1)
        html = footnote.html

        assert "<s>Strike</s>" in html

    def test_html_multi_paragraph(self):
        """Test html generates multiple <p> tags for multi-paragraph."""
        doc = Document(create_test_docx())

        doc.insert_footnote(["First para.", "Second para."], at="test document")

        footnote = doc.get_footnote(1)
        html = footnote.html

        assert html.count("<p>") == 2
        assert html.count("</p>") == 2
        assert "First para." in html
        assert "Second para." in html

    def test_html_escapes_special_chars(self):
        """Test html escapes HTML special characters."""
        doc = Document(create_test_docx())

        # Note: markdown parser may consume HTML-like content as inline HTML
        # So we test with & which is always escaped
        doc.insert_footnote("Contains & and > symbols", at="test document")

        footnote = doc.get_footnote(1)
        html = footnote.html

        # Should be escaped
        assert "&amp;" in html
        assert "&gt;" in html

    def test_endnote_html(self):
        """Test html property works on endnotes too."""
        doc = Document(create_test_docx())

        doc.insert_endnote("**Bold** endnote", at="test document")

        endnote = doc.get_endnote(1)
        html = endnote.html

        assert "<b>Bold</b>" in html


class TestTrackedChangesInFootnotes:
    """Tests for tracked changes inside footnotes."""

    def test_insert_tracked_in_footnote_after(self):
        """Test inserting tracked text after anchor in footnote."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Original citation text", at="test document")

        doc.insert_tracked_in_footnote(1, " [updated]", after="citation")

        # Verify insertion created w:ins element
        temp_dir = doc._temp_dir
        footnotes_path = temp_dir / "word" / "footnotes.xml"
        footnotes_xml = footnotes_path.read_text()

        assert "<w:ins" in footnotes_xml
        assert "[updated]" in footnotes_xml

    def test_insert_tracked_in_footnote_before(self):
        """Test inserting tracked text before anchor in footnote."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Original citation text", at="test document")

        doc.insert_tracked_in_footnote(1, "[see] ", before="Original")

        # Verify insertion created w:ins element
        temp_dir = doc._temp_dir
        footnotes_path = temp_dir / "word" / "footnotes.xml"
        footnotes_xml = footnotes_path.read_text()

        assert "<w:ins" in footnotes_xml
        assert "[see]" in footnotes_xml

    def test_delete_tracked_in_footnote(self):
        """Test deleting text with tracked changes in footnote."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Original citation text", at="test document")

        doc.delete_tracked_in_footnote(1, "citation")

        # Verify deletion created w:del element
        temp_dir = doc._temp_dir
        footnotes_path = temp_dir / "word" / "footnotes.xml"
        footnotes_xml = footnotes_path.read_text()

        assert "<w:del" in footnotes_xml
        assert "w:delText" in footnotes_xml or "delText" in footnotes_xml

    def test_replace_tracked_in_footnote(self):
        """Test replacing text with tracked changes in footnote."""
        doc = Document(create_test_docx())

        doc.insert_footnote("See citation 2020", at="test document")

        doc.replace_tracked_in_footnote(1, "2020", "2024")

        # Verify both deletion and insertion created
        temp_dir = doc._temp_dir
        footnotes_path = temp_dir / "word" / "footnotes.xml"
        footnotes_xml = footnotes_path.read_text()

        assert "<w:del" in footnotes_xml
        assert "<w:ins" in footnotes_xml
        assert "2024" in footnotes_xml

    def test_insert_tracked_in_footnote_not_found(self):
        """Test error when anchor text not found in footnote."""
        from python_docx_redline import TextNotFoundError

        doc = Document(create_test_docx())
        doc.insert_footnote("Footnote text", at="test document")

        with pytest.raises(TextNotFoundError):
            doc.insert_tracked_in_footnote(1, " new", after="nonexistent")

    def test_insert_tracked_in_footnote_no_params(self):
        """Test error when neither after nor before specified."""
        doc = Document(create_test_docx())
        doc.insert_footnote("Footnote text", at="test document")

        with pytest.raises(ValueError) as exc_info:
            doc.insert_tracked_in_footnote(1, " new")

        assert "after" in str(exc_info.value).lower() or "before" in str(exc_info.value).lower()

    def test_tracked_footnote_persists_after_save(self):
        """Test that tracked changes in footnotes persist after save/reload."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Original footnote", at="test document")
        doc.insert_tracked_in_footnote(1, " [modified]", after="footnote")

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "tracked_footnote.docx"
            doc.save(output_path)

            reloaded = Document(output_path)

            # Check footnotes.xml contains tracked change
            temp_dir = reloaded._temp_dir
            footnotes_path = temp_dir / "word" / "footnotes.xml"
            footnotes_xml = footnotes_path.read_text()

            assert "<w:ins" in footnotes_xml
            assert "[modified]" in footnotes_xml


class TestTrackedChangesInEndnotes:
    """Tests for tracked changes inside endnotes."""

    def test_insert_tracked_in_endnote_after(self):
        """Test inserting tracked text after anchor in endnote."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Original reference text", at="test document")

        doc.insert_tracked_in_endnote(1, " [see also]", after="reference")

        # Verify insertion created w:ins element
        temp_dir = doc._temp_dir
        endnotes_path = temp_dir / "word" / "endnotes.xml"
        endnotes_xml = endnotes_path.read_text()

        assert "<w:ins" in endnotes_xml
        assert "[see also]" in endnotes_xml

    def test_insert_tracked_in_endnote_before(self):
        """Test inserting tracked text before anchor in endnote."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Original reference text", at="test document")

        doc.insert_tracked_in_endnote(1, "[cf.] ", before="Original")

        # Verify insertion created w:ins element
        temp_dir = doc._temp_dir
        endnotes_path = temp_dir / "word" / "endnotes.xml"
        endnotes_xml = endnotes_path.read_text()

        assert "<w:ins" in endnotes_xml
        assert "[cf.]" in endnotes_xml

    def test_delete_tracked_in_endnote(self):
        """Test deleting text with tracked changes in endnote."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Original reference text", at="test document")

        doc.delete_tracked_in_endnote(1, "reference")

        # Verify deletion created w:del element
        temp_dir = doc._temp_dir
        endnotes_path = temp_dir / "word" / "endnotes.xml"
        endnotes_xml = endnotes_path.read_text()

        assert "<w:del" in endnotes_xml

    def test_replace_tracked_in_endnote(self):
        """Test replacing text with tracked changes in endnote."""
        doc = Document(create_test_docx())

        doc.insert_endnote("See ibid page 100", at="test document")

        doc.replace_tracked_in_endnote(1, "ibid", "op. cit.")

        # Verify both deletion and insertion created
        temp_dir = doc._temp_dir
        endnotes_path = temp_dir / "word" / "endnotes.xml"
        endnotes_xml = endnotes_path.read_text()

        assert "<w:del" in endnotes_xml
        assert "<w:ins" in endnotes_xml
        assert "op. cit." in endnotes_xml

    def test_insert_tracked_in_endnote_not_found(self):
        """Test error when anchor text not found in endnote."""
        from python_docx_redline import TextNotFoundError

        doc = Document(create_test_docx())
        doc.insert_endnote("Endnote text", at="test document")

        with pytest.raises(TextNotFoundError):
            doc.insert_tracked_in_endnote(1, " new", after="nonexistent")

    def test_tracked_endnote_persists_after_save(self):
        """Test that tracked changes in endnotes persist after save/reload."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Original endnote", at="test document")
        doc.replace_tracked_in_endnote(1, "endnote", "citation")

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "tracked_endnote.docx"
            doc.save(output_path)

            reloaded = Document(output_path)

            # Check endnotes.xml contains tracked changes
            temp_dir = reloaded._temp_dir
            endnotes_path = temp_dir / "word" / "endnotes.xml"
            endnotes_xml = endnotes_path.read_text()

            assert "<w:del" in endnotes_xml
            assert "<w:ins" in endnotes_xml


class TestTrackedChangesInNotesViaModel:
    """Tests for tracked changes via Footnote/Endnote model methods."""

    def test_footnote_insert_tracked_method(self):
        """Test footnote.insert_tracked() delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Footnote content here", at="test document")
        footnote = doc.get_footnote(1)

        footnote.insert_tracked(" [added]", after="content")

        # Verify change was made
        temp_dir = doc._temp_dir
        footnotes_path = temp_dir / "word" / "footnotes.xml"
        footnotes_xml = footnotes_path.read_text()

        assert "<w:ins" in footnotes_xml
        assert "[added]" in footnotes_xml

    def test_footnote_delete_tracked_method(self):
        """Test footnote.delete_tracked() delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Delete this word here", at="test document")
        footnote = doc.get_footnote(1)

        footnote.delete_tracked("this word")

        # Verify deletion was made
        temp_dir = doc._temp_dir
        footnotes_path = temp_dir / "word" / "footnotes.xml"
        footnotes_xml = footnotes_path.read_text()

        assert "<w:del" in footnotes_xml

    def test_footnote_replace_tracked_method(self):
        """Test footnote.replace_tracked() delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Replace old with new", at="test document")
        footnote = doc.get_footnote(1)

        footnote.replace_tracked("old", "updated")

        # Verify replacement was made
        temp_dir = doc._temp_dir
        footnotes_path = temp_dir / "word" / "footnotes.xml"
        footnotes_xml = footnotes_path.read_text()

        assert "<w:del" in footnotes_xml
        assert "<w:ins" in footnotes_xml
        assert "updated" in footnotes_xml

    def test_endnote_insert_tracked_method(self):
        """Test endnote.insert_tracked() delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Endnote content here", at="test document")
        endnote = doc.get_endnote(1)

        endnote.insert_tracked(" [added]", after="content")

        # Verify change was made
        temp_dir = doc._temp_dir
        endnotes_path = temp_dir / "word" / "endnotes.xml"
        endnotes_xml = endnotes_path.read_text()

        assert "<w:ins" in endnotes_xml
        assert "[added]" in endnotes_xml

    def test_endnote_delete_tracked_method(self):
        """Test endnote.delete_tracked() delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Delete this word here", at="test document")
        endnote = doc.get_endnote(1)

        endnote.delete_tracked("this word")

        # Verify deletion was made
        temp_dir = doc._temp_dir
        endnotes_path = temp_dir / "word" / "endnotes.xml"
        endnotes_xml = endnotes_path.read_text()

        assert "<w:del" in endnotes_xml

    def test_endnote_replace_tracked_method(self):
        """Test endnote.replace_tracked() delegates to document."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Replace old with new", at="test document")
        endnote = doc.get_endnote(1)

        endnote.replace_tracked("old", "updated")

        # Verify replacement was made
        temp_dir = doc._temp_dir
        endnotes_path = temp_dir / "word" / "endnotes.xml"
        endnotes_xml = endnotes_path.read_text()

        assert "<w:del" in endnotes_xml
        assert "<w:ins" in endnotes_xml
        assert "updated" in endnotes_xml

    def test_footnote_tracked_method_no_document(self):
        """Test footnote tracked methods raise when no document reference."""
        from python_docx_redline.models.footnote import Footnote

        elem = etree.Element(f"{{{WORD_NAMESPACE}}}footnote")
        elem.set(f"{{{WORD_NAMESPACE}}}id", "1")

        footnote = Footnote(elem, None)  # type: ignore

        with pytest.raises(ValueError) as exc_info:
            footnote.insert_tracked("text", after="anchor")

        assert "no document reference" in str(exc_info.value)

        with pytest.raises(ValueError):
            footnote.delete_tracked("text")

        with pytest.raises(ValueError):
            footnote.replace_tracked("old", "new")

    def test_endnote_tracked_method_no_document(self):
        """Test endnote tracked methods raise when no document reference."""
        from python_docx_redline.models.footnote import Endnote

        elem = etree.Element(f"{{{WORD_NAMESPACE}}}endnote")
        elem.set(f"{{{WORD_NAMESPACE}}}id", "1")

        endnote = Endnote(elem, None)  # type: ignore

        with pytest.raises(ValueError) as exc_info:
            endnote.insert_tracked("text", after="anchor")

        assert "no document reference" in str(exc_info.value)

        with pytest.raises(ValueError):
            endnote.delete_tracked("text")

        with pytest.raises(ValueError):
            endnote.replace_tracked("old", "new")


class TestTrackedChangesNoteNotFound:
    """Tests for NoteNotFoundError in tracked changes."""

    def test_insert_tracked_footnote_not_found(self):
        """Test NoteNotFoundError when footnote ID doesn't exist."""
        from python_docx_redline import NoteNotFoundError

        doc = Document(create_test_docx())
        doc.insert_footnote("Test footnote", at="test document")

        with pytest.raises(NoteNotFoundError) as exc_info:
            doc.insert_tracked_in_footnote(99, " new", after="Test")

        assert "99" in str(exc_info.value)

    def test_delete_tracked_endnote_not_found(self):
        """Test NoteNotFoundError when endnote ID doesn't exist."""
        from python_docx_redline import NoteNotFoundError

        doc = Document(create_test_docx())
        doc.insert_endnote("Test endnote", at="test document")

        with pytest.raises(NoteNotFoundError) as exc_info:
            doc.delete_tracked_in_endnote(99, "Test")

        assert "99" in str(exc_info.value)


class TestFindAllIncludesFootnotes:
    """Tests for find_all with include_footnotes parameter."""

    def test_find_all_includes_footnotes_when_requested(self):
        """Test that find_all searches footnotes when include_footnotes=True."""
        doc = Document(create_test_docx())

        # Insert footnote with searchable text
        doc.insert_footnote("This citation is important", at="test document")

        # Search without include_footnotes - should not find
        matches = doc.find_all("citation", include_footnotes=False)
        assert len(matches) == 0

        # Search with include_footnotes - should find
        matches = doc.find_all("citation", include_footnotes=True)
        assert len(matches) == 1
        assert matches[0].text == "citation"

    def test_find_all_match_location_shows_footnote_id(self):
        """Test that footnote matches have location like 'footnote:N'."""
        doc = Document(create_test_docx())

        doc.insert_footnote("First citation here", at="test document")
        doc.insert_footnote("Second citation there", at="Another paragraph")

        matches = doc.find_all("citation", include_footnotes=True)
        assert len(matches) == 2

        # Check location format
        assert matches[0].location == "footnote:1"
        assert matches[1].location == "footnote:2"

    def test_find_all_includes_body_and_footnotes(self):
        """Test that body and footnote matches are both returned."""
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>The word citation appears in body.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Another paragraph for footnote.</w:t></w:r></w:p>
  </w:body>
</w:document>"""

        doc = Document(create_test_docx(content))
        doc.insert_footnote("Citation in footnote", at="footnote")

        matches = doc.find_all("citation", include_footnotes=True, case_sensitive=False)
        assert len(matches) == 2

        # One from body, one from footnote
        locations = [m.location for m in matches]
        assert "body" in locations
        assert "footnote:1" in locations


class TestFindAllIncludesEndnotes:
    """Tests for find_all with include_endnotes parameter."""

    def test_find_all_includes_endnotes_when_requested(self):
        """Test that find_all searches endnotes when include_endnotes=True."""
        doc = Document(create_test_docx())

        # Insert endnote with searchable text
        doc.insert_endnote("This reference is important", at="test document")

        # Search without include_endnotes - should not find
        matches = doc.find_all("reference", include_endnotes=False)
        assert len(matches) == 0

        # Search with include_endnotes - should find
        matches = doc.find_all("reference", include_endnotes=True)
        assert len(matches) == 1
        assert matches[0].text == "reference"

    def test_find_all_match_location_shows_endnote_id(self):
        """Test that endnote matches have location like 'endnote:N'."""
        doc = Document(create_test_docx())

        doc.insert_endnote("First reference here", at="test document")
        doc.insert_endnote("Second reference there", at="Another paragraph")

        matches = doc.find_all("reference", include_endnotes=True)
        assert len(matches) == 2

        # Check location format
        assert matches[0].location == "endnote:1"
        assert matches[1].location == "endnote:2"


class TestScopeFootnoteN:
    """Tests for scope='footnote:N' to limit search to specific footnote."""

    def test_scope_footnote_n_limits_search(self):
        """Test that scope='footnote:N' searches only that footnote."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Citation in footnote one", at="test document")
        doc.insert_footnote("Citation in footnote two", at="Another paragraph")

        # Search all footnotes
        all_matches = doc.find_all("Citation", scope="footnotes")
        assert len(all_matches) == 2

        # Search only footnote 1
        matches_fn1 = doc.find_all("Citation", scope="footnote:1")
        assert len(matches_fn1) == 1
        assert matches_fn1[0].location == "footnote:1"

        # Search only footnote 2
        matches_fn2 = doc.find_all("Citation", scope="footnote:2")
        assert len(matches_fn2) == 1
        assert matches_fn2[0].location == "footnote:2"

    def test_scope_footnote_n_nonexistent(self):
        """Test that searching nonexistent footnote returns empty."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Some text", at="test document")

        # Search nonexistent footnote 99
        matches = doc.find_all("text", scope="footnote:99")
        assert len(matches) == 0


class TestScopeFootnotes:
    """Tests for scope='footnotes' to search all footnotes."""

    def test_scope_footnotes_searches_all_footnotes(self):
        """Test that scope='footnotes' searches all footnotes."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Important citation A", at="test document")
        doc.insert_footnote("Important citation B", at="Another paragraph")

        matches = doc.find_all("citation", scope="footnotes")
        assert len(matches) == 2

        # Both should be in footnotes
        assert all("footnote:" in m.location for m in matches)

    def test_scope_footnotes_does_not_search_body(self):
        """Test that scope='footnotes' doesn't include body matches."""
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Citation in body text.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Another paragraph.</w:t></w:r></w:p>
  </w:body>
</w:document>"""

        doc = Document(create_test_docx(content))
        doc.insert_footnote("Citation in footnote", at="Another paragraph")

        # Search only footnotes - should not find body citation
        matches = doc.find_all("Citation", scope="footnotes")
        assert len(matches) == 1
        assert matches[0].location == "footnote:1"


class TestScopeEndnoteN:
    """Tests for scope='endnote:N' to limit search to specific endnote."""

    def test_scope_endnote_n_limits_search(self):
        """Test that scope='endnote:N' searches only that endnote."""
        doc = Document(create_test_docx())

        doc.insert_endnote("Reference in endnote one", at="test document")
        doc.insert_endnote("Reference in endnote two", at="Another paragraph")

        # Search all endnotes
        all_matches = doc.find_all("Reference", scope="endnotes")
        assert len(all_matches) == 2

        # Search only endnote 1
        matches_en1 = doc.find_all("Reference", scope="endnote:1")
        assert len(matches_en1) == 1
        assert matches_en1[0].location == "endnote:1"

        # Search only endnote 2
        matches_en2 = doc.find_all("Reference", scope="endnote:2")
        assert len(matches_en2) == 1
        assert matches_en2[0].location == "endnote:2"


class TestScopeNotes:
    """Tests for scope='notes' to search both footnotes and endnotes."""

    def test_scope_notes_searches_both_types(self):
        """Test that scope='notes' searches both footnotes and endnotes."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Citation A", at="test document")
        doc.insert_endnote("Citation B", at="Another paragraph")

        matches = doc.find_all("Citation", scope="notes")
        assert len(matches) == 2

        # Should have one from each type
        locations = [m.location for m in matches]
        assert any("footnote:" in loc for loc in locations)
        assert any("endnote:" in loc for loc in locations)


class TestFindInFootnotes:
    """Tests for doc.find_in_footnotes() convenience method."""

    def test_find_in_footnotes_basic(self):
        """Test find_in_footnotes convenience method."""
        doc = Document(create_test_docx())

        doc.insert_footnote("citation text here", at="test document")
        doc.insert_footnote("Another citation there", at="Another paragraph")

        matches = doc.find_in_footnotes("citation")
        assert len(matches) == 2
        assert all("footnote:" in m.location for m in matches)

    def test_find_in_footnotes_case_insensitive(self):
        """Test find_in_footnotes with case_sensitive=False."""
        doc = Document(create_test_docx())

        doc.insert_footnote("CITATION uppercase", at="test document")

        matches = doc.find_in_footnotes("citation", case_sensitive=False)
        assert len(matches) == 1

        matches_sensitive = doc.find_in_footnotes("citation", case_sensitive=True)
        assert len(matches_sensitive) == 0


class TestFindInEndnotes:
    """Tests for doc.find_in_endnotes() convenience method."""

    def test_find_in_endnotes_basic(self):
        """Test find_in_endnotes convenience method."""
        doc = Document(create_test_docx())

        doc.insert_endnote("reference text here", at="test document")
        doc.insert_endnote("Another reference there", at="Another paragraph")

        matches = doc.find_in_endnotes("reference")
        assert len(matches) == 2
        assert all("endnote:" in m.location for m in matches)

    def test_find_in_endnotes_case_insensitive(self):
        """Test find_in_endnotes with case_sensitive=False."""
        doc = Document(create_test_docx())

        doc.insert_endnote("REFERENCE uppercase", at="test document")

        matches = doc.find_in_endnotes("reference", case_sensitive=False)
        assert len(matches) == 1

        matches_sensitive = doc.find_in_endnotes("reference", case_sensitive=True)
        assert len(matches_sensitive) == 0


class TestFindAllBothNotesAndBody:
    """Tests for searching body, footnotes, and endnotes together."""

    def test_find_all_body_footnotes_and_endnotes(self):
        """Test searching body, footnotes, and endnotes simultaneously."""
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Important text in body.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Another paragraph here.</w:t></w:r></w:p>
  </w:body>
</w:document>"""

        doc = Document(create_test_docx(content))
        doc.insert_footnote("Important text in footnote", at="body")
        doc.insert_endnote("Important text in endnote", at="Another paragraph")

        matches = doc.find_all(
            "Important",
            include_footnotes=True,
            include_endnotes=True,
        )

        assert len(matches) == 3

        # Check all three locations are represented
        locations = [m.location for m in matches]
        body_matches = [loc for loc in locations if loc == "body"]
        footnote_matches = [loc for loc in locations if loc.startswith("footnote:")]
        endnote_matches = [loc for loc in locations if loc.startswith("endnote:")]

        assert len(body_matches) == 1
        assert len(footnote_matches) == 1
        assert len(endnote_matches) == 1

    def test_find_all_match_indices_are_sequential(self):
        """Test that match indices are sequential across all sources."""
        doc = Document(create_test_docx())

        doc.insert_footnote("Text in footnote", at="test document")
        doc.insert_endnote("Text in endnote", at="Another paragraph")

        matches = doc.find_all("Text", include_footnotes=True, include_endnotes=True)

        # Indices should be 0, 1 (no body matches in this case)
        indices = [m.index for m in matches]
        assert indices == list(range(len(matches)))
