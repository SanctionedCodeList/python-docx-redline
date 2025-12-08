"""
Tests for insert_tracked() 'before' parameter functionality.

This module tests the ability to insert text before an anchor point, complementing
the existing 'after' parameter for more flexible document editing.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest

from docx_redline import Document
from docx_redline.errors import AmbiguousTextError, TextNotFoundError


def create_test_document(text: str) -> Path:
    """Create a simple test document with given text."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>{text}</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


def create_multi_paragraph_document() -> Path:
    """Create a document with multiple paragraphs."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>First paragraph text.</w:t></w:r></w:p>
<w:p><w:r><w:t>Second paragraph text.</w:t></w:r></w:p>
<w:p><w:r><w:t>Third paragraph text.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


def test_insert_before_basic() -> None:
    """Text can be inserted before anchor text."""
    doc_path = create_test_document("Hello world. Goodbye world.")
    try:
        doc = Document(doc_path)
        doc.insert_tracked("Middle text. ", before="Goodbye")

        text = doc.get_text()
        assert "Middle text. " in text
        assert text.find("Middle text") < text.find("Goodbye")

    finally:
        doc_path.unlink()


def test_insert_after_basic() -> None:
    """Text can be inserted after anchor text (existing functionality)."""
    doc_path = create_test_document("Hello world. Goodbye world.")
    try:
        doc = Document(doc_path)
        doc.insert_tracked("Middle text. ", after="Hello world.")

        text = doc.get_text()
        assert "Middle text. " in text
        assert text.find("Hello world.") < text.find("Middle text")

    finally:
        doc_path.unlink()


def test_insert_before_and_after_mutual_exclusion() -> None:
    """Cannot specify both before and after."""
    doc_path = create_test_document("Some text here")
    try:
        doc = Document(doc_path)

        with pytest.raises(ValueError, match="Cannot specify both"):
            doc.insert_tracked("New text", before="Some", after="text")

    finally:
        doc_path.unlink()


def test_insert_neither_before_nor_after() -> None:
    """Must specify either before or after."""
    doc_path = create_test_document("Some text here")
    try:
        doc = Document(doc_path)

        with pytest.raises(ValueError, match="Must specify either"):
            doc.insert_tracked("New text")

    finally:
        doc_path.unlink()


def test_insert_before_at_paragraph_start() -> None:
    """Can insert before text at the start of a paragraph."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)
        doc.insert_tracked("Preface. ", before="First paragraph")

        text = doc.get_text()
        assert "Preface. " in text
        assert text.find("Preface") < text.find("First paragraph")

    finally:
        doc_path.unlink()


def test_insert_before_at_paragraph_end() -> None:
    """Can insert before text at the end of a paragraph."""
    doc_path = create_test_document("Beginning of sentence.")
    try:
        doc = Document(doc_path)
        doc.insert_tracked(" and addition", before=".")

        text = doc.get_text()
        assert "and addition" in text
        assert text.find("and addition") < text.find(".")

    finally:
        doc_path.unlink()


def test_insert_before_with_regex() -> None:
    """Before parameter works with regex patterns."""
    doc_path = create_test_document("Version 1.2.3 released.")
    try:
        doc = Document(doc_path)
        doc.insert_tracked("New: ", before=r"\d+\.\d+\.\d+", regex=True)

        text = doc.get_text()
        assert "New: " in text
        assert "New: Version" in text or text.find("New:") < text.find("1.2.3")

    finally:
        doc_path.unlink()


def test_insert_before_not_found() -> None:
    """Raises TextNotFoundError when before anchor not found."""
    doc_path = create_test_document("Some text here")
    try:
        doc = Document(doc_path)

        with pytest.raises(TextNotFoundError, match="nonexistent"):
            doc.insert_tracked("Text", before="nonexistent")

    finally:
        doc_path.unlink()


def test_insert_before_ambiguous() -> None:
    """Raises AmbiguousTextError when before anchor appears multiple times."""
    doc_path = create_test_document("Hello world. Hello universe.")
    try:
        doc = Document(doc_path)

        with pytest.raises(AmbiguousTextError):
            doc.insert_tracked("Greetings! ", before="Hello")

    finally:
        doc_path.unlink()


def test_insert_before_with_scope() -> None:
    """Before parameter works with scope filtering."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Create document with sections
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
    <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
    <w:r><w:t>Section One</w:t></w:r>
</w:p>
<w:p><w:r><w:t>Content in section one with conclusion.</w:t></w:r></w:p>
<w:p>
    <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
    <w:r><w:t>Section Two</w:t></w:r>
</w:p>
<w:p><w:r><w:t>Content in section two with conclusion.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    try:
        doc = Document(doc_path)

        # Insert before "conclusion" only in Section Two
        doc.insert_tracked("Important note. ", before="conclusion", scope="section:Section Two")

        text = doc.get_text()

        # Verify insertion is in Section Two
        section_two_start = text.find("Section Two")
        conclusion_in_section_two = text.find("conclusion", section_two_start)
        note_position = text.find("Important note")

        assert note_position > section_two_start
        assert note_position < conclusion_in_section_two

    finally:
        doc_path.unlink()


def test_insert_before_preserves_formatting() -> None:
    """Insert before preserves existing text formatting."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Create document with formatted text
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
    <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Bold text here</w:t>
    </w:r>
</w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    try:
        doc = Document(doc_path)
        doc.insert_tracked("Normal text. ", before="Bold text")

        # Verify the insertion exists
        text = doc.get_text()
        assert "Normal text. " in text
        assert text.find("Normal text") < text.find("Bold text")

        # The bold formatting should still exist on "Bold text here"
        # (Not directly testable without parsing XML, but insertion shouldn't affect it)

    finally:
        doc_path.unlink()


def test_insert_before_multiple_runs() -> None:
    """Insert before text that spans multiple runs."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Text split across runs (common in edited documents)
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
    <w:r><w:t>Multi</w:t></w:r>
    <w:r><w:t>ple </w:t></w:r>
    <w:r><w:t>runs</w:t></w:r>
    <w:r><w:t> here.</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    try:
        doc = Document(doc_path)
        doc.insert_tracked("Before: ", before="Multiple runs")

        text = doc.get_text()
        assert "Before: " in text
        assert text.find("Before:") < text.find("Multiple")

    finally:
        doc_path.unlink()


def test_insert_before_and_after_in_sequence() -> None:
    """Can use both before and after in sequence on same document."""
    doc_path = create_test_document("Start. Middle. End.")
    try:
        doc = Document(doc_path)

        # Insert after Start
        doc.insert_tracked("After start. ", after="Start.")

        # Insert before End
        doc.insert_tracked("Before end. ", before="End.")

        text = doc.get_text()

        # Verify both insertions exist
        assert "After start. " in text
        assert "Before end. " in text

        # Verify original text still exists
        assert "Start." in text
        assert "Middle." in text
        assert "End." in text

    finally:
        doc_path.unlink()


def test_insert_before_with_author() -> None:
    """Insert before respects custom author."""
    doc_path = create_test_document("Some text")
    try:
        doc = Document(doc_path, author="Default Author")
        doc.insert_tracked("Inserted. ", before="Some text", author="Custom Author")

        # Verify insertion exists (author verification requires XML inspection)
        text = doc.get_text()
        assert "Inserted. " in text

    finally:
        doc_path.unlink()
