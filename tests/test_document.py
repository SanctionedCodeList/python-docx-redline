"""
Tests for the Document class.

These tests verify the core functionality of loading documents,
inserting tracked changes, and saving documents.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_redline import (
    AmbiguousTextError,
    Document,
    TextNotFoundError,
)


# Minimal Word document XML structure
MINIMAL_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a test document.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Section 2.1: Introduction</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Some content here.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


def create_test_docx(content: str = MINIMAL_DOCUMENT_XML) -> Path:
    """Create a minimal test .docx file.

    Args:
        content: The document.xml content

    Returns:
        Path to the created .docx file
    """
    # Create a temporary directory
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    # Create the .docx structure
    word_dir = temp_dir / "word"
    word_dir.mkdir(exist_ok=True)

    # Write document.xml
    document_xml = word_dir / "document.xml"
    document_xml.write_text(content)

    # Create the .docx ZIP file
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zip_ref:
        for file in temp_dir.rglob("*"):
            if file.is_file() and file != docx_path:
                arcname = file.relative_to(temp_dir)
                zip_ref.write(file, arcname)

    return docx_path


def test_document_load_docx():
    """Test loading a .docx file."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        assert doc.xml_root is not None
        assert doc.author == "Claude"

        # Verify the document was parsed
        body = doc.xml_root.find(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body"
        )
        assert body is not None
    finally:
        # Cleanup
        docx_path.unlink()


def test_insert_tracked_basic():
    """Test basic tracked insertion."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path, author="TestAuthor")

        # Insert after "Section 2.1"
        doc.insert_tracked("New clause text here.", after="Section 2.1")

        # Save the document
        doc.save(output_path)

        # Verify the output file was created
        assert output_path.exists()

        # Load and verify the modified document
        doc2 = Document(output_path)

        # Check for the insertion element
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        assert len(insertions) > 0, "No insertion elements found"

        # Verify the inserted text
        inserted_text = "".join(insertions[0].itertext())
        assert "New clause text here." in inserted_text

    finally:
        # Cleanup
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_text_not_found():
    """Test TextNotFoundError when anchor text doesn't exist."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        with pytest.raises(TextNotFoundError) as exc_info:
            doc.insert_tracked("New text", after="Nonexistent text")

        assert "Nonexistent text" in str(exc_info.value)
        assert "Could not find" in str(exc_info.value)

    finally:
        docx_path.unlink()


def test_ambiguous_text():
    """Test AmbiguousTextError when multiple matches are found."""
    # Create a document with duplicate text
    duplicate_doc = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Section 1: Introduction</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Section 2: Introduction</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    docx_path = create_test_docx(duplicate_doc)

    try:
        doc = Document(docx_path)

        with pytest.raises(AmbiguousTextError) as exc_info:
            doc.insert_tracked("New text", after="Introduction")

        assert "Introduction" in str(exc_info.value)
        assert "2 occurrences" in str(exc_info.value)

    finally:
        docx_path.unlink()


def test_context_manager():
    """Test Document as a context manager."""
    docx_path = create_test_docx()

    try:
        with Document(docx_path) as doc:
            assert doc.xml_root is not None

    finally:
        docx_path.unlink()


if __name__ == "__main__":
    # Run a simple smoke test
    print("Running smoke test...")
    test_document_load_docx()
    print("✓ Document loading works")

    test_insert_tracked_basic()
    print("✓ Insert tracked works")

    test_text_not_found()
    print("✓ TextNotFoundError works")

    test_ambiguous_text()
    print("✓ AmbiguousTextError works")

    test_context_manager()
    print("✓ Context manager works")

    print("\nAll smoke tests passed! ✓")
