"""
Tests for move_tracked() functionality.

Move tracking creates linked markers (moveFrom/moveTo) that show text
was relocated rather than deleted and re-added.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_redline import AmbiguousTextError, Document, TextNotFoundError


def create_test_document() -> Path:
    """Create a test document with multiple paragraphs."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>First paragraph with some text.</w:t></w:r></w:p>
<w:p><w:r><w:t>Second paragraph to move.</w:t></w:r></w:p>
<w:p><w:r><w:t>Third paragraph as anchor.</w:t></w:r></w:p>
<w:p><w:r><w:t>Fourth paragraph at the end.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)

    return doc_path


class TestMoveTrackedBasic:
    """Basic tests for move_tracked functionality."""

    def test_move_after_anchor(self) -> None:
        """Test moving text to after an anchor."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            doc.move_tracked(
                "Second paragraph to move.",
                after="Third paragraph as anchor.",
                author="Editor",
            )

            # Check that moveFrom markers exist
            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            move_from_starts = list(doc.xml_root.iter(f"{{{word_ns}}}moveFromRangeStart"))
            move_from_ends = list(doc.xml_root.iter(f"{{{word_ns}}}moveFromRangeEnd"))
            move_froms = list(doc.xml_root.iter(f"{{{word_ns}}}moveFrom"))

            assert len(move_from_starts) == 1
            assert len(move_from_ends) == 1
            assert len(move_froms) == 1

            # Check that moveTo markers exist
            move_to_starts = list(doc.xml_root.iter(f"{{{word_ns}}}moveToRangeStart"))
            move_to_ends = list(doc.xml_root.iter(f"{{{word_ns}}}moveToRangeEnd"))
            move_tos = list(doc.xml_root.iter(f"{{{word_ns}}}moveTo"))

            assert len(move_to_starts) == 1
            assert len(move_to_ends) == 1
            assert len(move_tos) == 1

        finally:
            doc_path.unlink()

    def test_move_before_anchor(self) -> None:
        """Test moving text to before an anchor."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            doc.move_tracked(
                "Fourth paragraph at the end.",
                before="First paragraph",
                author="Editor",
            )

            # Check markers exist
            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            move_from_starts = list(doc.xml_root.iter(f"{{{word_ns}}}moveFromRangeStart"))
            move_to_starts = list(doc.xml_root.iter(f"{{{word_ns}}}moveToRangeStart"))

            assert len(move_from_starts) == 1
            assert len(move_to_starts) == 1

        finally:
            doc_path.unlink()

    def test_move_names_match(self) -> None:
        """Test that moveFrom and moveTo have matching names."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            doc.move_tracked(
                "Second paragraph to move.",
                after="Third paragraph as anchor.",
            )

            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

            move_from_start = list(doc.xml_root.iter(f"{{{word_ns}}}moveFromRangeStart"))[0]
            move_to_start = list(doc.xml_root.iter(f"{{{word_ns}}}moveToRangeStart"))[0]

            from_name = move_from_start.get(f"{{{word_ns}}}name")
            to_name = move_to_start.get(f"{{{word_ns}}}name")

            assert from_name is not None
            assert from_name == to_name
            assert from_name.startswith("move")

        finally:
            doc_path.unlink()

    def test_move_ids_are_linked(self) -> None:
        """Test that range IDs link starts to ends."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            doc.move_tracked(
                "Second paragraph to move.",
                after="Third paragraph as anchor.",
            )

            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

            # Check moveFrom range IDs match
            move_from_start = list(doc.xml_root.iter(f"{{{word_ns}}}moveFromRangeStart"))[0]
            move_from_end = list(doc.xml_root.iter(f"{{{word_ns}}}moveFromRangeEnd"))[0]

            from_start_id = move_from_start.get(f"{{{word_ns}}}id")
            from_end_id = move_from_end.get(f"{{{word_ns}}}id")

            assert from_start_id == from_end_id

            # Check moveTo range IDs match
            move_to_start = list(doc.xml_root.iter(f"{{{word_ns}}}moveToRangeStart"))[0]
            move_to_end = list(doc.xml_root.iter(f"{{{word_ns}}}moveToRangeEnd"))[0]

            to_start_id = move_to_start.get(f"{{{word_ns}}}id")
            to_end_id = move_to_end.get(f"{{{word_ns}}}id")

            assert to_start_id == to_end_id

        finally:
            doc_path.unlink()


class TestMoveTrackedContent:
    """Tests for move content handling."""

    def test_move_preserves_text(self) -> None:
        """Test that moved text appears in both moveFrom and moveTo."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            doc.move_tracked(
                "Second paragraph to move.",
                after="Third paragraph as anchor.",
            )

            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

            # Get text from moveFrom (should use delText)
            move_from = list(doc.xml_root.iter(f"{{{word_ns}}}moveFrom"))[0]
            del_texts = move_from.findall(f".//{{{word_ns}}}delText")
            from_text = "".join(t.text or "" for t in del_texts)

            # Get text from moveTo (should use regular t)
            move_to = list(doc.xml_root.iter(f"{{{word_ns}}}moveTo"))[0]
            t_texts = move_to.findall(f".//{{{word_ns}}}t")
            to_text = "".join(t.text or "" for t in t_texts)

            assert from_text == "Second paragraph to move."
            assert to_text == "Second paragraph to move."

        finally:
            doc_path.unlink()

    def test_move_has_author(self) -> None:
        """Test that move markers have author attribution."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            doc.move_tracked(
                "Second paragraph to move.",
                after="Third paragraph as anchor.",
                author="TestAuthor",
            )

            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

            move_from = list(doc.xml_root.iter(f"{{{word_ns}}}moveFrom"))[0]
            move_to = list(doc.xml_root.iter(f"{{{word_ns}}}moveTo"))[0]
            move_from_start = list(doc.xml_root.iter(f"{{{word_ns}}}moveFromRangeStart"))[0]
            move_to_start = list(doc.xml_root.iter(f"{{{word_ns}}}moveToRangeStart"))[0]

            assert move_from.get(f"{{{word_ns}}}author") == "TestAuthor"
            assert move_to.get(f"{{{word_ns}}}author") == "TestAuthor"
            assert move_from_start.get(f"{{{word_ns}}}author") == "TestAuthor"
            assert move_to_start.get(f"{{{word_ns}}}author") == "TestAuthor"

        finally:
            doc_path.unlink()

    def test_move_has_date(self) -> None:
        """Test that move markers have date attribution."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            doc.move_tracked(
                "Second paragraph to move.",
                after="Third paragraph as anchor.",
            )

            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

            move_from = list(doc.xml_root.iter(f"{{{word_ns}}}moveFrom"))[0]
            move_to = list(doc.xml_root.iter(f"{{{word_ns}}}moveTo"))[0]

            # Dates should be ISO 8601 format
            from_date = move_from.get(f"{{{word_ns}}}date")
            to_date = move_to.get(f"{{{word_ns}}}date")

            assert from_date is not None
            assert "T" in from_date  # ISO 8601 format
            assert to_date is not None

        finally:
            doc_path.unlink()


class TestMoveTrackedErrors:
    """Tests for error handling in move_tracked."""

    def test_move_requires_after_or_before(self) -> None:
        """Test that either after or before must be specified."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError, match="Must specify either"):
                doc.move_tracked("Some text")

        finally:
            doc_path.unlink()

    def test_move_rejects_both_after_and_before(self) -> None:
        """Test that both after and before cannot be specified."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError, match="Cannot specify both"):
                doc.move_tracked(
                    "Some text",
                    after="anchor1",
                    before="anchor2",
                )

        finally:
            doc_path.unlink()

    def test_move_source_not_found(self) -> None:
        """Test error when source text is not found."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(TextNotFoundError):
                doc.move_tracked(
                    "Nonexistent text to move",
                    after="First paragraph",
                )

        finally:
            doc_path.unlink()

    def test_move_destination_not_found(self) -> None:
        """Test error when destination anchor is not found."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            with pytest.raises(TextNotFoundError):
                doc.move_tracked(
                    "Second paragraph to move.",
                    after="Nonexistent anchor",
                )

        finally:
            doc_path.unlink()

    def test_move_ambiguous_source(self) -> None:
        """Test error when source text has multiple matches."""
        doc_path = Path(tempfile.mktemp(suffix=".docx"))

        document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Repeated text here.</w:t></w:r></w:p>
<w:p><w:r><w:t>Repeated text here.</w:t></w:r></w:p>
<w:p><w:r><w:t>Anchor paragraph.</w:t></w:r></w:p>
</w:body>
</w:document>"""

        content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

        root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

        with zipfile.ZipFile(doc_path, "w") as docx:
            docx.writestr("[Content_Types].xml", content_types_xml)
            docx.writestr("_rels/.rels", root_rels)
            docx.writestr("word/document.xml", document_xml)

        try:
            doc = Document(doc_path)

            with pytest.raises(AmbiguousTextError):
                doc.move_tracked(
                    "Repeated text here.",
                    after="Anchor paragraph.",
                )

        finally:
            doc_path.unlink()


class TestMoveTrackedPersistence:
    """Tests for move tracking persistence after save."""

    def test_move_persists_after_save(self) -> None:
        """Test that move markers persist after save and reload."""
        doc_path = create_test_document()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Create and save
            doc = Document(doc_path)
            doc.move_tracked(
                "Second paragraph to move.",
                after="Third paragraph as anchor.",
            )
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)
            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

            move_from_starts = list(doc2.xml_root.iter(f"{{{word_ns}}}moveFromRangeStart"))
            move_to_starts = list(doc2.xml_root.iter(f"{{{word_ns}}}moveToRangeStart"))

            assert len(move_from_starts) == 1
            assert len(move_to_starts) == 1

            # Verify names still match
            from_name = move_from_starts[0].get(f"{{{word_ns}}}name")
            to_name = move_to_starts[0].get(f"{{{word_ns}}}name")
            assert from_name == to_name

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_multiple_moves_unique_names(self) -> None:
        """Test that multiple moves get unique names."""
        doc_path = create_test_document()
        try:
            doc = Document(doc_path)

            # First move
            doc.move_tracked(
                "Second paragraph to move.",
                after="First paragraph",
            )

            # Second move
            doc.move_tracked(
                "Fourth paragraph at the end.",
                before="Third paragraph",
            )

            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

            move_from_starts = list(doc.xml_root.iter(f"{{{word_ns}}}moveFromRangeStart"))
            move_to_starts = list(doc.xml_root.iter(f"{{{word_ns}}}moveToRangeStart"))

            assert len(move_from_starts) == 2
            assert len(move_to_starts) == 2

            # Get all names
            from_names = [m.get(f"{{{word_ns}}}name") for m in move_from_starts]
            to_names = [m.get(f"{{{word_ns}}}name") for m in move_to_starts]

            # Each move should have unique name
            assert len(set(from_names)) == 2  # Two unique names
            assert set(from_names) == set(to_names)  # Sources match destinations

        finally:
            doc_path.unlink()


class TestMoveTrackedWithScope:
    """Tests for move_tracked with scope parameters."""

    def test_move_with_source_scope(self) -> None:
        """Test moving with source scope limitation."""
        doc_path = Path(tempfile.mktemp(suffix=".docx"))

        document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Section One</w:t></w:r></w:p>
<w:p><w:r><w:t>Text in section one.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Section Two</w:t></w:r></w:p>
<w:p><w:r><w:t>Text in section two.</w:t></w:r></w:p>
<w:p><w:r><w:t>Destination anchor.</w:t></w:r></w:p>
</w:body>
</w:document>"""

        content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

        root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

        with zipfile.ZipFile(doc_path, "w") as docx:
            docx.writestr("[Content_Types].xml", content_types_xml)
            docx.writestr("_rels/.rels", root_rels)
            docx.writestr("word/document.xml", document_xml)

        try:
            doc = Document(doc_path)

            # Move text from section one specifically
            doc.move_tracked(
                "Text in section one.",
                after="Destination anchor.",
                source_scope="section:Section One",
            )

            # Should succeed - move markers should exist
            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            move_from_starts = list(doc.xml_root.iter(f"{{{word_ns}}}moveFromRangeStart"))
            assert len(move_from_starts) == 1

        finally:
            doc_path.unlink()
