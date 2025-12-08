"""
Test that delete_all_comments() properly cleans up all comment-related files.

This test verifies the fix for the bug where delete_all_comments() left orphaned
comments.xml files, causing LibreOffice to reject the document.
"""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from docx_redline import Document


def create_document_with_comments() -> Path:
    """Create a test document with comments."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Minimal Word document with a comment
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r><w:t>This is text with a comment.</w:t></w:r>
  <w:commentRangeEnd w:id="0"/>
  <w:r>
    <w:commentReference w:id="0"/>
  </w:r>
</w:p>
</w:body>
</w:document>"""

    comments_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Test User" w:date="2025-12-08T00:00:00Z">
    <w:p>
      <w:r><w:t>This is a comment.</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>"""

    rels_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/comments.xml", comments_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


def test_delete_all_comments_removes_comment_files() -> None:
    """Test that delete_all_comments() removes comments.xml."""
    doc_path = create_document_with_comments()
    try:
        doc = Document(doc_path)
        doc.delete_all_comments()

        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        # Extract and verify no comments.xml
        with zipfile.ZipFile(output_path, "r") as docx:
            file_list = docx.namelist()
            assert "word/comments.xml" not in file_list, "comments.xml should be removed"

        output_path.unlink()

    finally:
        doc_path.unlink()


def test_delete_all_comments_removes_comment_relationships() -> None:
    """Test that delete_all_comments() removes comment relationships."""
    doc_path = create_document_with_comments()
    try:
        doc = Document(doc_path)
        doc.delete_all_comments()

        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        # Extract and check document.xml.rels
        with zipfile.ZipFile(output_path, "r") as docx:
            rels_content = docx.read("word/_rels/document.xml.rels")
            rels_tree = etree.fromstring(rels_content)

            # Check no comment relationships exist
            for rel in rels_tree:
                rel_type = rel.get("Type")
                assert (
                    "comment" not in rel_type.lower()
                ), f"Comment relationship should be removed: {rel_type}"

        output_path.unlink()

    finally:
        doc_path.unlink()


def test_delete_all_comments_removes_comment_content_types() -> None:
    """Test that delete_all_comments() removes comment content types."""
    doc_path = create_document_with_comments()
    try:
        doc = Document(doc_path)
        doc.delete_all_comments()

        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        # Extract and check [Content_Types].xml
        with zipfile.ZipFile(output_path, "r") as docx:
            ct_content = docx.read("[Content_Types].xml")
            ct_tree = etree.fromstring(ct_content)

            # Check no comment content types
            ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
            for override in ct_tree:
                if override.tag == f"{{{ct_ns}}}Override":
                    part_name = override.get("PartName")
                    assert (
                        "comment" not in part_name.lower()
                    ), f"Comment content type should be removed: {part_name}"

        output_path.unlink()

    finally:
        doc_path.unlink()


def test_delete_all_comments_removes_comment_markers() -> None:
    """Test that delete_all_comments() removes comment markers from document.xml."""
    doc_path = create_document_with_comments()
    try:
        doc = Document(doc_path)

        # Verify comments exist before
        word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        comment_starts = doc.xml_root.xpath(
            ".//w:commentRangeStart", namespaces={"w": word_namespace}
        )
        assert len(comment_starts) == 1, "Should have 1 comment marker before deletion"

        doc.delete_all_comments()

        # Verify comments removed after
        comment_starts = doc.xml_root.xpath(
            ".//w:commentRangeStart", namespaces={"w": word_namespace}
        )
        comment_ends = doc.xml_root.xpath(".//w:commentRangeEnd", namespaces={"w": word_namespace})
        comment_refs = doc.xml_root.xpath(".//w:commentReference", namespaces={"w": word_namespace})

        assert len(comment_starts) == 0, "All comment range starts should be removed"
        assert len(comment_ends) == 0, "All comment range ends should be removed"
        assert len(comment_refs) == 0, "All comment references should be removed"

    finally:
        doc_path.unlink()


def test_delete_all_comments_complete_cleanup() -> None:
    """Test that delete_all_comments() performs complete cleanup."""
    doc_path = create_document_with_comments()
    try:
        doc = Document(doc_path)
        doc.delete_all_comments()

        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        # Verify complete cleanup
        with zipfile.ZipFile(output_path, "r") as docx:
            file_list = docx.namelist()

            # No comment files
            assert not any(
                "comment" in f.lower() for f in file_list
            ), "No comment-related files should exist"

            # No comment relationships
            rels_content = docx.read("word/_rels/document.xml.rels")
            assert b"comment" not in rels_content.lower(), "No comment relationships should exist"

            # No comment content types
            ct_content = docx.read("[Content_Types].xml")
            assert b"comment" not in ct_content.lower(), "No comment content types should exist"

            # No comment markers in document.xml
            doc_content = docx.read("word/document.xml")
            assert b"commentRange" not in doc_content, "No comment markers should exist"
            assert b"commentReference" not in doc_content, "No comment references should exist"

        output_path.unlink()

    finally:
        doc_path.unlink()


def test_delete_all_comments_on_document_without_comments() -> None:
    """Test that delete_all_comments() is safe on documents without comments."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Simple text without comments.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    try:
        doc = Document(doc_path)
        text_before = doc.get_text()

        # Should not raise any errors
        doc.delete_all_comments()

        # Text should be unchanged
        text_after = doc.get_text()
        assert text_before == text_after, "Text should be unchanged"

    finally:
        doc_path.unlink()
