"""
Tests for the Document class.

These tests verify the core functionality of loading documents,
inserting tracked changes, and saving documents.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest

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
    # Create a document with duplicate text (compact XML to avoid whitespace issues)
    duplicate_doc = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Section 1: Introduction</w:t></w:r></w:p><w:p><w:r><w:t>Section 2: Introduction</w:t></w:r></w:p></w:body></w:document>"""

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


def test_delete_tracked():
    """Test tracked deletion."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path, author="TestAuthor")

        # Delete "test document"
        doc.delete_tracked("test document")

        # Save the document
        doc.save(output_path)

        # Verify the output file was created
        assert output_path.exists()

        # Load and verify the modified document
        doc2 = Document(output_path)

        # Check for the deletion element
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(deletions) > 0, "No deletion elements found"

        # Verify the deleted text
        deleted_text = "".join(deletions[0].itertext())
        assert "test document" in deleted_text

    finally:
        # Cleanup
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_replace_tracked():
    """Test tracked replacement."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path, author="TestAuthor")

        # Replace "test document" with "sample file"
        doc.replace_tracked("test document", "sample file")

        # Save the document
        doc.save(output_path)

        # Verify the output file was created
        assert output_path.exists()

        # Load and verify the modified document
        doc2 = Document(output_path)

        # Check for deletion element
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(deletions) > 0, "No deletion elements found"

        # Check for insertion element
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        assert len(insertions) > 0, "No insertion elements found"

        # Verify the deleted text
        deleted_text = "".join(deletions[0].itertext())
        assert "test document" in deleted_text

        # Verify the inserted text
        inserted_text = "".join(insertions[0].itertext())
        assert "sample file" in inserted_text

    finally:
        # Cleanup
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_delete_partial_run():
    """Test deleting text that's part of a longer run."""
    # Create document with longer text in single run (compact XML to avoid whitespace issues)
    partial_doc = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>This is a very long sentence with specific words.</w:t></w:r></w:p></w:body></w:document>"""

    docx_path = create_test_docx(partial_doc)
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Delete just "specific words" which is in the middle of the run
        doc.delete_tracked("specific words")

        # Save and reload
        doc.save(output_path)
        doc2 = Document(output_path)

        # Check for deletion
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(deletions) > 0

        # Verify we still have the before and after text
        all_text = "".join(doc2.xml_root.itertext())
        assert "very long sentence with" in all_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_replace_across_multiple_runs():
    """Test replacing text that spans multiple runs."""
    # Create document with text split across runs (compact XML to avoid whitespace issues)
    multi_run_doc = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>This is </w:t></w:r><w:r><w:t>fragmented</w:t></w:r><w:r><w:t> text here.</w:t></w:r></w:p></w:body></w:document>"""

    docx_path = create_test_docx(multi_run_doc)
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Replace text that spans across runs
        doc.replace_tracked("fragmented text", "unified content")

        # Save and reload
        doc.save(output_path)
        doc2 = Document(output_path)

        # Check for both deletion and insertion
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )

        assert len(deletions) > 0
        assert len(insertions) > 0

        deleted_text = "".join(deletions[0].itertext())
        inserted_text = "".join(insertions[0].itertext())

        assert "fragmented text" in deleted_text
        assert "unified content" in inserted_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_accept_all_changes():
    """Test accepting all tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Add some tracked changes
        doc.insert_tracked("NEW TEXT", after="This is")
        doc.delete_tracked("test document")
        doc.save(output_path)

        # Reload and verify tracked changes exist
        doc2 = Document(output_path)
        insertions_before = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions_before = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions_before) > 0
        assert len(deletions_before) > 0

        # Accept all changes
        doc2.accept_all_changes()
        doc2.save(output_path)

        # Reload and verify no tracked changes remain
        doc3 = Document(output_path)
        insertions_after = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions_after = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions_after) == 0, "Insertions should be unwrapped"
        assert len(deletions_after) == 0, "Deletions should be removed"

        # Verify the inserted text is still present
        all_text = "".join(doc3.xml_root.itertext())
        assert "NEW TEXT" in all_text, "Inserted text should remain after accepting"

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_delete_all_comments():
    """Test deleting all comments from document."""
    # Create document with comment markers
    doc_with_comments = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Text before comment</w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>Commented text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r><w:r><w:t>Text after comment</w:t></w:r></w:p></w:body></w:document>"""

    docx_path = create_test_docx(doc_with_comments)
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Verify comment markers exist
        comment_starts_before = doc.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart"
        )
        comment_ends_before = doc.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd"
        )
        comment_refs_before = doc.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference"
        )

        assert len(comment_starts_before) > 0
        assert len(comment_ends_before) > 0
        assert len(comment_refs_before) > 0

        # Delete all comments
        doc.delete_all_comments()
        doc.save(output_path)

        # Reload and verify no comment markers remain
        doc2 = Document(output_path)
        comment_starts_after = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart"
        )
        comment_ends_after = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd"
        )
        comment_refs_after = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference"
        )

        assert len(comment_starts_after) == 0
        assert len(comment_ends_after) == 0
        assert len(comment_refs_after) == 0

        # Verify the commented text is still present
        all_text = "".join(doc2.xml_root.itertext())
        assert "Commented text" in all_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


# Run tests with: pytest tests/test_document.py -v
# Or from project root: .venv/bin/pytest tests/ -v
