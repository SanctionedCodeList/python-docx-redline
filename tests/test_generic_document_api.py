"""Tests for Phase 3: Generic Document API methods.

This module tests the generic document editing methods (insert, delete, replace, move)
and verifies that the *_tracked aliases maintain backward compatibility.

Tests cover:
- Generic methods with track=False (default, untracked)
- Generic methods with track=True (tracked)
- Backward compatibility of *_tracked aliases
- Mixed tracked/untracked in same document
- find_all() with include_deleted parameter
"""

import io
import zipfile

import pytest

from python_docx_redline import Document


def create_test_docx(paragraphs: list[str]) -> bytes:
    """Create a minimal valid docx file with the given paragraphs."""
    buffer = io.BytesIO()

    with zipfile.ZipFile(buffer, "w") as zf:
        # [Content_Types].xml
        content_types = b"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
        zf.writestr("[Content_Types].xml", content_types)

        # _rels/.rels
        root_rels = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
        zf.writestr("_rels/.rels", root_rels)

        # word/document.xml with paragraphs
        para_xml = "\n".join(f"    <w:p><w:r><w:t>{p}</w:t></w:r></w:p>" for p in paragraphs)
        document = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
{para_xml}
  </w:body>
</w:document>""".encode()
        zf.writestr("word/document.xml", document)

        # word/_rels/document.xml.rels
        doc_rels = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""
        zf.writestr("word/_rels/document.xml.rels", doc_rels)

    buffer.seek(0)
    return buffer.read()


@pytest.fixture
def sample_docx_bytes() -> bytes:
    """Create a simple Word document for testing."""
    return create_test_docx(
        [
            "This is the first paragraph.",
            "This paragraph contains old text that will be edited.",
            "Another paragraph with some content.",
            "Final paragraph here.",
        ]
    )


@pytest.fixture
def sample_doc(sample_docx_bytes: bytes) -> Document:
    """Create a Document instance for testing."""
    return Document(sample_docx_bytes)


class TestInsertGeneric:
    """Test Document.insert() generic method."""

    def test_insert_untracked_default(self, sample_doc: Document):
        """Test that insert() is untracked by default."""
        sample_doc.insert(" [inserted]", after="first paragraph")

        # Verify text was inserted
        text = sample_doc.get_text()
        assert "[inserted]" in text

        # Verify no tracked changes
        assert not sample_doc.has_tracked_changes()

    def test_insert_tracked_explicit(self, sample_doc: Document):
        """Test insert() with track=True creates tracked change."""
        sample_doc.insert(" [inserted]", after="first paragraph", track=True)

        # Verify text was inserted
        text = sample_doc.get_text()
        assert "[inserted]" in text

        # Verify tracked changes exist
        assert sample_doc.has_tracked_changes()

    def test_insert_before(self, sample_doc: Document):
        """Test insert() with before parameter."""
        sample_doc.insert("[prefix] ", before="first paragraph")

        text = sample_doc.get_text()
        assert "[prefix]" in text
        assert not sample_doc.has_tracked_changes()


class TestDeleteGeneric:
    """Test Document.delete() generic method."""

    def test_delete_untracked_default(self, sample_doc: Document):
        """Test that delete() is untracked by default."""
        sample_doc.delete("old text")

        # Verify text was deleted
        text = sample_doc.get_text()
        assert "old text" not in text

        # Verify no tracked changes
        assert not sample_doc.has_tracked_changes()

    def test_delete_tracked_explicit(self, sample_doc: Document):
        """Test delete() with track=True creates tracked change."""
        sample_doc.delete("old text", track=True)

        # Verify no tracked changes (deleted text is still there, just marked)
        assert sample_doc.has_tracked_changes()

    def test_delete_all_occurrences(self):
        """Test delete() with occurrence='all'."""
        doc_bytes = create_test_docx(
            [
                "Remove this word here.",
                "And remove this word there.",
            ]
        )

        doc = Document(doc_bytes)
        doc.delete("this", occurrence="all")

        text = doc.get_text()
        assert "this" not in text
        assert not doc.has_tracked_changes()


class TestReplaceGeneric:
    """Test Document.replace() generic method."""

    def test_replace_untracked_default(self, sample_doc: Document):
        """Test that replace() is untracked by default."""
        sample_doc.replace("old text", "new text")

        text = sample_doc.get_text()
        assert "old text" not in text
        assert "new text" in text
        assert not sample_doc.has_tracked_changes()

    def test_replace_tracked_explicit(self, sample_doc: Document):
        """Test replace() with track=True creates tracked change."""
        sample_doc.replace("old text", "new text", track=True)

        # Verify tracked changes exist
        assert sample_doc.has_tracked_changes()

    def test_replace_all_occurrences(self):
        """Test replace() with occurrence='all'."""
        doc_bytes = create_test_docx(
            [
                "Replace word one.",
                "Replace word two.",
            ]
        )

        doc = Document(doc_bytes)
        doc.replace("word", "term", occurrence="all")

        text = doc.get_text()
        assert "word" not in text
        assert "term" in text
        assert not doc.has_tracked_changes()


class TestMoveGeneric:
    """Test Document.move() generic method."""

    def test_move_untracked_default(self):
        """Test that move() is untracked by default."""
        doc_bytes = create_test_docx(
            [
                "First paragraph.",
                "Text to move here.",
                "Destination marker.",
            ]
        )

        doc = Document(doc_bytes)
        doc.move("Text to move", after="Destination marker")

        # Verify no tracked changes
        assert not doc.has_tracked_changes()

    def test_move_tracked_explicit(self):
        """Test move() with track=True creates tracked change."""
        doc_bytes = create_test_docx(
            [
                "First paragraph.",
                "Text to move here.",
                "Destination marker.",
            ]
        )

        doc = Document(doc_bytes)
        doc.move("Text to move", after="Destination marker", track=True)

        # Verify tracked changes exist
        assert doc.has_tracked_changes()


class TestTrackedAliasesBackwardCompat:
    """Test that *_tracked methods maintain backward compatibility."""

    def test_insert_tracked_alias(self, sample_doc: Document):
        """Test insert_tracked() works exactly as before."""
        sample_doc.insert_tracked(" [tracked]", after="first paragraph")

        text = sample_doc.get_text()
        assert "[tracked]" in text
        assert sample_doc.has_tracked_changes()

    def test_delete_tracked_alias(self, sample_doc: Document):
        """Test delete_tracked() works exactly as before."""
        sample_doc.delete_tracked("old text")

        assert sample_doc.has_tracked_changes()

    def test_replace_tracked_alias(self, sample_doc: Document):
        """Test replace_tracked() works exactly as before."""
        sample_doc.replace_tracked("old text", "new text")

        assert sample_doc.has_tracked_changes()

    def test_move_tracked_alias(self):
        """Test move_tracked() works exactly as before."""
        doc_bytes = create_test_docx(
            [
                "First paragraph.",
                "Text to move here.",
                "Destination marker.",
            ]
        )

        doc = Document(doc_bytes)
        doc.move_tracked("Text to move", after="Destination marker")

        assert doc.has_tracked_changes()


class TestMixedTrackedUntracked:
    """Test mixed tracked and untracked edits in same document."""

    def test_untracked_then_tracked(self, sample_doc: Document):
        """Test untracked edit followed by tracked edit."""
        # Untracked edit
        sample_doc.replace("first paragraph", "initial paragraph")
        assert not sample_doc.has_tracked_changes()

        # Tracked edit
        sample_doc.replace("old text", "updated text", track=True)
        assert sample_doc.has_tracked_changes()

        text = sample_doc.get_text()
        assert "initial paragraph" in text  # Untracked change is visible
        assert "updated text" in text  # Tracked change is visible

    def test_tracked_then_untracked(self, sample_doc: Document):
        """Test tracked edit followed by untracked edit."""
        # Tracked edit
        sample_doc.delete_tracked("old text")
        assert sample_doc.has_tracked_changes()

        # Untracked edit (should just work)
        sample_doc.insert(" [added]", after="some content")

        text = sample_doc.get_text()
        assert "[added]" in text

    def test_multiple_mixed_edits(self, sample_doc: Document):
        """Test multiple mixed tracked/untracked edits."""
        # Silent cleanup (untracked)
        sample_doc.replace("first", "1st")
        sample_doc.replace("Final", "Last")

        # Substantive changes (tracked)
        sample_doc.replace_tracked("old text", "revised content")
        sample_doc.insert_tracked(" (IMPORTANT)", after="Another paragraph")

        # Verify document state
        assert sample_doc.has_tracked_changes()
        text = sample_doc.get_text()
        assert "1st" in text
        assert "Last" in text
        assert "revised content" in text
        assert "(IMPORTANT)" in text


class TestFindAllIncludeDeleted:
    """Test find_all() with include_deleted parameter."""

    def test_find_all_excludes_deleted_by_default(self, sample_doc: Document):
        """Test that find_all() excludes deleted text by default."""
        # Delete some text (tracked, so it's still in XML)
        sample_doc.delete_tracked("old text")

        # Search for deleted text
        matches = sample_doc.find_all("old text")

        # Should not find deleted text by default
        assert len(matches) == 0

    def test_find_all_includes_deleted_when_requested(self, sample_doc: Document):
        """Test find_all(include_deleted=True) finds deleted text."""
        # Delete some text (tracked)
        sample_doc.delete_tracked("old text")

        # Search with include_deleted=True
        matches = sample_doc.find_all("old text", include_deleted=True)

        # Should find deleted text
        assert len(matches) == 1

    def test_find_all_normal_text_unaffected(self, sample_doc: Document):
        """Test that include_deleted doesn't affect finding normal text."""
        # Search for non-deleted text
        matches_default = sample_doc.find_all("first paragraph")
        matches_explicit = sample_doc.find_all("first paragraph", include_deleted=False)

        assert len(matches_default) == len(matches_explicit)
        assert len(matches_default) > 0


class TestSaveAndReload:
    """Test that changes persist after save and reload."""

    def test_untracked_changes_persist(self, sample_doc: Document):
        """Test that untracked changes are in saved document."""
        sample_doc.replace("old text", "modified text")

        # Save to bytes
        doc_bytes = sample_doc.save_to_bytes()

        # Reload and verify
        reloaded = Document(doc_bytes)
        text = reloaded.get_text()
        assert "modified text" in text
        assert "old text" not in text
        assert not reloaded.has_tracked_changes()

    def test_mixed_changes_persist(self, sample_doc: Document):
        """Test that mixed tracked/untracked changes are in saved document."""
        # Untracked change
        sample_doc.replace("first", "1st")

        # Tracked change
        sample_doc.replace_tracked("old text", "revised content")

        # Save and reload
        doc_bytes = sample_doc.save_to_bytes()
        reloaded = Document(doc_bytes)

        text = reloaded.get_text()
        assert "1st" in text
        assert "revised content" in text
        assert reloaded.has_tracked_changes()
