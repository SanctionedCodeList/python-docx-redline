"""
Tests for text extraction with XML formatting whitespace (Issue #1).

This test file addresses the bug where text split across multiple runs
with intervening XML formatting (newlines, indentation) was not properly
extracted, causing search operations to fail.

See: docs/ISSUE_TEXT_SEARCH_WITH_WHITESPACE.md
"""

import tempfile
import zipfile
from pathlib import Path

from docx_redline import Document


def create_document_with_formatted_xml() -> Path:
    """Create a document with intentional XML formatting whitespace.

    This simulates what happens when Word saves documents with pretty-printed
    XML or when documents have been edited extensively.
    """
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # XML with intentional formatting/whitespace between elements
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>database</w:t>
      </w:r>
      <w:r>
        <w:t> records</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


def create_document_with_multiline_whitespace() -> Path:
    """Create document with multiple newlines between runs.

    This simulates the exact bug from the issue report where
    "database records their property ownership" had newlines between words.
    """
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>It claims merely that a database</w:t>
      </w:r>

      <w:r>
        <w:t> records</w:t>
      </w:r>


      <w:r>
        <w:t> their property ownership</w:t>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


def create_document_with_indented_xml() -> Path:
    """Create document with deeply indented XML structure."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>The quick</w:t>
            </w:r>
            <w:r>
                <w:t> brown fox</w:t>
            </w:r>
            <w:r>
                <w:t> jumps</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


# Text extraction tests


def test_text_extraction_with_formatted_xml() -> None:
    """Text with XML formatting whitespace should extract cleanly."""
    doc_path = create_document_with_formatted_xml()
    try:
        doc = Document(doc_path)
        text = doc.get_text()

        # Should NOT have newlines from XML formatting
        assert text == "database records"
        assert "\n" not in text.replace("\n\n", "")  # Allow paragraph separators

    finally:
        doc_path.unlink()


def test_text_extraction_with_multiline_whitespace() -> None:
    """Text with multiple newlines between runs should extract cleanly."""
    doc_path = create_document_with_multiline_whitespace()
    try:
        doc = Document(doc_path)
        text = doc.get_text()

        # Should be continuous text without XML whitespace
        expected = "It claims merely that a database records their property ownership."
        assert text == expected

        # Verify no extra newlines within the paragraph
        assert text.count("\n") == 0

    finally:
        doc_path.unlink()


def test_text_extraction_with_indented_xml() -> None:
    """Text with deeply indented XML should extract cleanly."""
    doc_path = create_document_with_indented_xml()
    try:
        doc = Document(doc_path)
        text = doc.get_text()

        assert text == "The quick brown fox jumps"
        assert "\n" not in text.replace("\n\n", "")

    finally:
        doc_path.unlink()


def test_paragraph_text_property_with_formatted_xml() -> None:
    """Paragraph.text should handle formatted XML correctly."""
    doc_path = create_document_with_formatted_xml()
    try:
        doc = Document(doc_path)

        # Access paragraph directly
        para = doc.paragraphs[0]
        assert para.text == "database records"
        assert "\n" not in para.text

    finally:
        doc_path.unlink()


# Search operation tests (the original bug report)


def test_search_across_formatted_runs() -> None:
    """Search should work across runs with XML formatting."""
    doc_path = create_document_with_formatted_xml()
    try:
        doc = Document(doc_path)

        # This should succeed - the original bug made this fail
        doc.replace_tracked("database records", "database compiles")

        # Verify the replacement worked
        text = doc.get_text()
        assert "database compiles" in text
        assert "database records" not in text

    finally:
        doc_path.unlink()


def test_search_multiword_phrase_across_runs() -> None:
    """Multi-word phrases split across runs should be findable."""
    doc_path = create_document_with_multiline_whitespace()
    try:
        doc = Document(doc_path)

        # Original bug: this would raise TextNotFoundError
        doc.replace_tracked(
            "database records their property ownership",
            "database compiles their property ownership data",
        )

        text = doc.get_text()
        assert "database compiles their property ownership data" in text

    finally:
        doc_path.unlink()


def test_search_short_phrase_across_runs() -> None:
    """Short phrases across runs should be findable."""
    doc_path = create_document_with_multiline_whitespace()
    try:
        doc = Document(doc_path)

        # Original bug report attempted this - should work now
        doc.replace_tracked("database records", "database compiles")

        text = doc.get_text()
        assert "database compiles" in text

    finally:
        doc_path.unlink()


def test_search_adjacent_words_across_runs() -> None:
    """Adjacent words split across runs should be findable."""
    doc_path = create_document_with_multiline_whitespace()
    try:
        doc = Document(doc_path)

        # Original bug: "records their property" was not found
        doc.replace_tracked("records their property", "documents their asset")

        text = doc.get_text()
        assert "documents their asset" in text

    finally:
        doc_path.unlink()


def test_insert_after_text_in_formatted_xml() -> None:
    """insert_tracked should work with formatted XML."""
    doc_path = create_document_with_formatted_xml()
    try:
        doc = Document(doc_path)

        # Should be able to find and insert after text
        doc.insert_tracked(" their data", after="database records")

        text = doc.get_text()
        assert "database records their data" in text

    finally:
        doc_path.unlink()


def test_delete_text_in_formatted_xml() -> None:
    """delete_tracked should work with formatted XML."""
    doc_path = create_document_with_multiline_whitespace()
    try:
        doc = Document(doc_path)

        # Should be able to find and delete text
        doc.delete_tracked("records their property ownership")

        text = doc.get_text()
        assert "records" not in text
        assert "It claims merely that a database." in text

    finally:
        doc_path.unlink()


# Edge cases


def test_empty_runs_with_formatting() -> None:
    """Empty runs with formatting should not break text extraction."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello</w:t>
      </w:r>
      <w:r>
      </w:r>
      <w:r>
        <w:t> world</w:t>
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
        text = doc.get_text()

        assert text == "Hello world"
        assert "\n" not in text

    finally:
        doc_path.unlink()


def test_single_character_runs_with_formatting() -> None:
    """Single character runs (like Eric White's algorithm creates) should work."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Simulate what Eric White's algorithm creates
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>H</w:t>
      </w:r>
      <w:r>
        <w:t>e</w:t>
      </w:r>
      <w:r>
        <w:t>l</w:t>
      </w:r>
      <w:r>
        <w:t>l</w:t>
      </w:r>
      <w:r>
        <w:t>o</w:t>
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
        text = doc.get_text()

        assert text == "Hello"

    finally:
        doc_path.unlink()


def test_runs_with_mixed_content() -> None:
    """Runs with mixed content (text and other elements) should work."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>Bold</w:t>
      </w:r>
      <w:r>
        <w:t> normal</w:t>
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
        text = doc.get_text()

        assert text == "Bold normal"

    finally:
        doc_path.unlink()
