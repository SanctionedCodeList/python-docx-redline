"""
Tests for the Document class.

These tests verify the core functionality of loading documents,
inserting tracked changes, and saving documents.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest

from python_docx_redline import (
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
    """Create a minimal but valid OOXML test .docx file.

    Args:
        content: The document.xml content

    Returns:
        Path to the created .docx file
    """
    # Create a temporary directory
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    # Proper Content_Types.xml
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    # Proper relationships file
    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    # Create the .docx ZIP file with proper structure
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", content)

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
    # Create document with text split across runs
    # Note: xml:space="preserve" required for text with trailing/leading whitespace
    multi_run_doc = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:xml="http://www.w3.org/XML/1998/namespace"><w:body><w:p><w:r><w:t xml:space="preserve">This is </w:t></w:r><w:r><w:t>fragmented</w:t></w:r><w:r><w:t xml:space="preserve"> text here.</w:t></w:r></w:p></w:body></w:document>"""

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


def test_accept_insertions():
    """Test accepting all tracked insertions."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Add some tracked changes - both insertions and deletions
        doc.insert_tracked("NEW TEXT 1", after="This is")
        doc.insert_tracked("NEW TEXT 2", after="Section 2.1")
        doc.delete_tracked("test document")
        doc.save(output_path)

        # Reload and accept only insertions
        doc2 = Document(output_path)
        count = doc2.accept_insertions()
        assert count == 2, "Should have accepted 2 insertions"
        doc2.save(output_path)

        # Verify insertions are unwrapped but deletions remain
        doc3 = Document(output_path)
        insertions = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions) == 0, "All insertions should be unwrapped"
        assert len(deletions) > 0, "Deletions should still exist"

        # Verify inserted text is still present
        all_text = "".join(doc3.xml_root.itertext())
        assert "NEW TEXT 1" in all_text
        assert "NEW TEXT 2" in all_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_reject_insertions():
    """Test rejecting all tracked insertions."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Add tracked changes
        doc.insert_tracked("NEW TEXT 1", after="This is")
        doc.insert_tracked("NEW TEXT 2", after="Section 2.1")
        doc.delete_tracked("test document")
        doc.save(output_path)

        # Reload and reject only insertions
        doc2 = Document(output_path)
        count = doc2.reject_insertions()
        assert count == 2, "Should have rejected 2 insertions"
        doc2.save(output_path)

        # Verify insertions are removed but deletions remain
        doc3 = Document(output_path)
        insertions = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions) == 0, "All insertions should be removed"
        assert len(deletions) > 0, "Deletions should still exist"

        # Verify inserted text is NOT present
        all_text = "".join(doc3.xml_root.itertext())
        assert "NEW TEXT 1" not in all_text
        assert "NEW TEXT 2" not in all_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_accept_deletions():
    """Test accepting all tracked deletions."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Add tracked changes
        doc.insert_tracked("NEW TEXT", after="This is")
        doc.delete_tracked("test document")
        doc.delete_tracked("Section 2.1")
        doc.save(output_path)

        # Reload and accept only deletions
        doc2 = Document(output_path)
        count = doc2.accept_deletions()
        assert count == 2, "Should have accepted 2 deletions"
        doc2.save(output_path)

        # Verify deletions are removed but insertions remain
        doc3 = Document(output_path)
        insertions = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions) > 0, "Insertions should still exist"
        assert len(deletions) == 0, "All deletions should be removed"

        # Verify deleted text is NOT present in regular text
        regular_text = "".join(
            elem.text or ""
            for elem in doc3.xml_root.findall(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
            )
        )
        assert "test document" not in regular_text
        assert "Section 2.1" not in regular_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_reject_deletions():
    """Test rejecting all tracked deletions."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Add tracked changes
        doc.insert_tracked("NEW TEXT", after="This is")
        doc.delete_tracked("test document")
        doc.delete_tracked("Section 2.1")
        doc.save(output_path)

        # Reload and reject only deletions
        doc2 = Document(output_path)
        count = doc2.reject_deletions()
        assert count == 2, "Should have rejected 2 deletions"
        doc2.save(output_path)

        # Verify deletions are unwrapped but insertions remain
        doc3 = Document(output_path)
        insertions = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions = doc3.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions) > 0, "Insertions should still exist"
        assert len(deletions) == 0, "All deletions should be unwrapped"

        # Verify deleted text IS restored in regular text
        regular_text = "".join(
            elem.text or ""
            for elem in doc3.xml_root.findall(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
            )
        )
        assert "test document" in regular_text
        assert "Section 2.1" in regular_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_accept_change_by_id():
    """Test accepting a specific change by ID."""
    # Create document with explicit change IDs
    doc_with_ids = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:xml="http://www.w3.org/XML/1998/namespace"><w:body><w:p><w:r><w:t xml:space="preserve">Text before </w:t></w:r><w:ins w:id="1" w:author="Author1"><w:r><w:t>insertion1</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> middle </w:t></w:r><w:ins w:id="2" w:author="Author2"><w:r><w:t>insertion2</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> text after</w:t></w:r></w:p></w:body></w:document>"""

    docx_path = create_test_docx(doc_with_ids)
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Accept only change ID 1
        doc.accept_change("1")
        doc.save(output_path)

        # Verify only ID 1 is unwrapped
        doc2 = Document(output_path)
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        insertions = doc2.xml_root.findall(f".//{{{word_ns}}}ins")
        assert len(insertions) == 1, "Only 1 insertion should remain"
        assert insertions[0].get(f"{{{word_ns}}}id") == "2"

        # Verify text from change 1 is still present (unwrapped)
        all_text = "".join(doc2.xml_root.itertext())
        assert "insertion1" in all_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_reject_change_by_id():
    """Test rejecting a specific change by ID."""
    doc_with_ids = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:xml="http://www.w3.org/XML/1998/namespace"><w:body><w:p><w:r><w:t xml:space="preserve">Text before </w:t></w:r><w:ins w:id="1" w:author="Author1"><w:r><w:t>insertion1</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> middle </w:t></w:r><w:ins w:id="2" w:author="Author2"><w:r><w:t>insertion2</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> text after</w:t></w:r></w:p></w:body></w:document>"""

    docx_path = create_test_docx(doc_with_ids)
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Reject only change ID 1
        doc.reject_change("1")
        doc.save(output_path)

        # Verify only ID 1 is removed
        doc2 = Document(output_path)
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        insertions = doc2.xml_root.findall(f".//{{{word_ns}}}ins")
        assert len(insertions) == 1, "Only 1 insertion should remain"
        assert insertions[0].get(f"{{{word_ns}}}id") == "2"

        # Verify text from change 1 is NOT present (removed)
        all_text = "".join(doc2.xml_root.itertext())
        assert "insertion1" not in all_text
        assert "insertion2" in all_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_accept_change_by_id_not_found():
    """Test accept_change raises ValueError for non-existent ID."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        with pytest.raises(ValueError) as exc_info:
            doc.accept_change("999")

        assert "999" in str(exc_info.value)
        assert "No tracked change found" in str(exc_info.value)

    finally:
        docx_path.unlink()


def test_reject_change_by_id_not_found():
    """Test reject_change raises ValueError for non-existent ID."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        with pytest.raises(ValueError) as exc_info:
            doc.reject_change("999")

        assert "999" in str(exc_info.value)
        assert "No tracked change found" in str(exc_info.value)

    finally:
        docx_path.unlink()


def test_accept_by_author():
    """Test accepting all changes by a specific author."""
    doc_with_authors = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:xml="http://www.w3.org/XML/1998/namespace"><w:body><w:p><w:r><w:t xml:space="preserve">Text </w:t></w:r><w:ins w:id="1" w:author="Alice"><w:r><w:t>by Alice</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> </w:t></w:r><w:ins w:id="2" w:author="Bob"><w:r><w:t>by Bob</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> </w:t></w:r><w:del w:id="3" w:author="Alice"><w:r><w:delText>deleted by Alice</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> </w:t></w:r><w:del w:id="4" w:author="Bob"><w:r><w:delText>deleted by Bob</w:delText></w:r></w:del></w:p></w:body></w:document>"""

    docx_path = create_test_docx(doc_with_authors)
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Accept all changes by Alice
        count = doc.accept_by_author("Alice")
        assert count == 2, "Should have accepted 2 changes by Alice"
        doc.save(output_path)

        # Verify only Bob's changes remain
        doc2 = Document(output_path)
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        insertions = doc2.xml_root.findall(f".//{{{word_ns}}}ins")
        deletions = doc2.xml_root.findall(f".//{{{word_ns}}}del")

        assert len(insertions) == 1, "Only Bob's insertion should remain"
        assert len(deletions) == 1, "Only Bob's deletion should remain"

        # All remaining changes should be by Bob
        for ins in insertions:
            assert ins.get(f"{{{word_ns}}}author") == "Bob"
        for del_elem in deletions:
            assert del_elem.get(f"{{{word_ns}}}author") == "Bob"

        # Alice's insertion text should still be present (unwrapped)
        all_text = "".join(doc2.xml_root.itertext())
        assert "by Alice" in all_text

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_reject_by_author():
    """Test rejecting all changes by a specific author."""
    doc_with_authors = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:xml="http://www.w3.org/XML/1998/namespace"><w:body><w:p><w:r><w:t xml:space="preserve">Text </w:t></w:r><w:ins w:id="1" w:author="Alice"><w:r><w:t>ALICE_INSERT</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> </w:t></w:r><w:ins w:id="2" w:author="Bob"><w:r><w:t>BOB_INSERT</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> </w:t></w:r><w:del w:id="3" w:author="Alice"><w:r><w:delText>ALICE_DELETED</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> </w:t></w:r><w:del w:id="4" w:author="Bob"><w:r><w:delText>BOB_DELETED</w:delText></w:r></w:del></w:p></w:body></w:document>"""

    docx_path = create_test_docx(doc_with_authors)
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Reject all changes by Alice
        count = doc.reject_by_author("Alice")
        assert count == 2, "Should have rejected 2 changes by Alice"
        doc.save(output_path)

        # Verify only Bob's changes remain
        doc2 = Document(output_path)
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        insertions = doc2.xml_root.findall(f".//{{{word_ns}}}ins")
        deletions = doc2.xml_root.findall(f".//{{{word_ns}}}del")

        assert len(insertions) == 1, "Only Bob's insertion should remain"
        assert len(deletions) == 1, "Only Bob's deletion should remain"

        # Alice's insertion text should NOT be present (removed)
        # Alice's deletion should be restored (unwrapped)
        regular_text = "".join(
            elem.text or "" for elem in doc2.xml_root.findall(f".//{{{word_ns}}}t")
        )
        assert "ALICE_INSERT" not in regular_text
        assert "ALICE_DELETED" in regular_text  # Deletion was unwrapped
        assert "BOB_INSERT" in regular_text  # Still in <w:ins>

    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_accept_by_author_no_matches():
    """Test accept_by_author returns 0 when author has no changes."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        # Add a change by TestAuthor
        doc.insert_tracked("test", after="This is", author="TestAuthor")

        # Try to accept changes by non-existent author
        count = doc.accept_by_author("NonExistentAuthor")
        assert count == 0, "Should return 0 when no changes by author"

    finally:
        docx_path.unlink()


def test_reject_by_author_no_matches():
    """Test reject_by_author returns 0 when author has no changes."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        # Add a change by TestAuthor
        doc.insert_tracked("test", after="This is", author="TestAuthor")

        # Try to reject changes by non-existent author
        count = doc.reject_by_author("NonExistentAuthor")
        assert count == 0, "Should return 0 when no changes by author"

    finally:
        docx_path.unlink()


def test_insert_tracked_multi_run_text():
    """Test insert_tracked finds text fragmented across multiple runs.

    Word documents often have text split across multiple <w:r> (run) elements
    due to formatting changes, editing history, or spell-check boundaries.
    This test verifies the search algorithm correctly handles such fragmentation.

    See: docs/issues/ISSUE_INSERT_TRACKED_RETURNS_NONE.md
    """
    # Create document with text fragmented across 4 runs
    # Simulates Word's behavior with legal citations
    multi_run_document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>See citation </w:t></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:t>09 N.E.3d 390, 397</w:t></w:r>
      <w:r><w:t>-</w:t></w:r>
      <w:r><w:t>99 (Ind. 2018).</w:t></w:r>
      <w:r><w:t> More text here.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

    docx_path = create_test_docx(multi_run_document)

    try:
        doc = Document(docx_path)

        # Verify the text appears correct when extracted
        full_text = doc.get_text()
        assert "109 N.E.3d 390, 397-99 (Ind. 2018)." in full_text

        # Search for text that spans all 4 fragmented runs
        # This should work because TextSearch concatenates runs
        anchor = "109 N.E.3d 390, 397-99 (Ind. 2018)."
        doc.insert_tracked(" [CITATION VERIFIED]", after=anchor)

        # Verify insertion was successful
        new_text = doc.get_text()
        assert "[CITATION VERIFIED]" in new_text
        assert "2018). [CITATION VERIFIED]" in new_text

    finally:
        docx_path.unlink()


def test_insert_tracked_multi_run_with_smart_quotes():
    """Test insert_tracked handles smart quotes in fragmented text.

    Combines two challenges:
    1. Text fragmented across multiple runs
    2. Smart quotes in document but straight quotes in search

    See: docs/issues/ISSUE_INSERT_TRACKED_RETURNS_NONE.md
    """
    # Text with smart quotes fragmented across runs
    smart_quote_multi_run = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>The court held </w:t></w:r>
      <w:r><w:t>\u201cthe game-related </w:t></w:r>
      <w:r><w:t>statistics are </w:t></w:r>
      <w:r><w:t>matters of public </w:t></w:r>
      <w:r><w:t>interest.\u201d 1</w:t></w:r>
      <w:r><w:t>09 N.E.3d.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

    docx_path = create_test_docx(smart_quote_multi_run)

    try:
        doc = Document(docx_path)

        # Search with straight quotes - should match smart quotes via normalization
        anchor = '"the game-related statistics are matters of public interest." 109 N.E.3d.'
        doc.insert_tracked(" [FOUND]", after=anchor)

        new_text = doc.get_text()
        assert "[FOUND]" in new_text

    finally:
        docx_path.unlink()


# Run tests with: pytest tests/test_document.py -v
# Or from project root: .venv/bin/pytest tests/ -v
