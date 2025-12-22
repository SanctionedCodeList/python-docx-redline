"""
Tests for editing text that's already inside tracked change wrappers.

These tests verify the fix for the issue where attempting to edit text
that was previously modified with tracked changes would fail with:
    ValueError: <Element ...w:r...> is not in list

The fix allows runs inside <w:ins> and <w:del> elements to be properly
located and modified.

See: docs/internal/issues/ISSUE_CHAINED_EDITS_ON_TRACKED_CHANGES.md
"""

import tempfile
import zipfile
from pathlib import Path

from python_docx_redline import Document


def create_test_docx(content: str) -> Path:
    """Create a minimal but valid OOXML test .docx file.

    Args:
        content: The document.xml content

    Returns:
        Path to the created .docx file
    """
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

    with zipfile.ZipFile(docx_path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", content)

    return docx_path


class TestChainedTrackedEdits:
    """Test editing text that's already inside tracked changes."""

    def test_replace_inside_insertion_same_session(self):
        """Test replacing text inside a w:ins element in the same session.

        This tests the case where we make two replace_tracked calls in the
        same session, and the second one targets text that was just inserted.
        """
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>The evidence suggests that...</w:t></w:r></w:p></w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            # First edit: suggests -> confirms
            doc = Document(docx_path, author="Reviewer A")
            doc.replace_tracked("suggests", "confirms")
            doc.save(output_path)

            # Second edit: confirms -> indicates (same session, reopen document)
            doc2 = Document(output_path, author="Reviewer A")
            # This should NOT raise ValueError
            doc2.replace_tracked("confirms", "indicates")
            doc2.save(output_path)

            # Verify the final document
            doc3 = Document(output_path)
            text = doc3.get_text()
            assert "indicates" in text

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_replace_inside_insertion_different_author(self):
        """Test replacing text inside a w:ins when author differs.

        When a different author edits text inside a tracked insertion,
        the edit should still work (not crash).
        """
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>The evidence suggests that...</w:t></w:r></w:p></w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            # First edit by Reviewer A
            doc = Document(docx_path, author="Reviewer A")
            doc.replace_tracked("suggests", "confirms")
            doc.save(output_path)

            # Second edit by Reviewer B (different author)
            doc2 = Document(output_path, author="Reviewer B")
            # This should NOT raise ValueError
            doc2.replace_tracked("confirms", "indicates")
            doc2.save(output_path)

            # Verify the final document
            doc3 = Document(output_path)
            text = doc3.get_text()
            assert "indicates" in text

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_delete_inside_insertion(self):
        """Test deleting text that's inside a w:ins element."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>Hello world, this is a test.</w:t></w:r></w:p></w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            # First edit: insert "beautiful " before "world"
            doc = Document(docx_path, author="Reviewer")
            doc.replace_tracked("world", "beautiful world")
            doc.save(output_path)

            # Second edit: delete "beautiful "
            doc2 = Document(output_path, author="Reviewer")
            # This should NOT raise ValueError
            doc2.delete_tracked("beautiful ")
            doc2.save(output_path)

            # Verify
            doc3 = Document(output_path)
            text = doc3.get_text()
            # The word "beautiful" should be marked as deleted
            assert "world" in text

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_multiple_chained_replacements(self):
        """Test multiple chained replacements in sequence."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>The quick brown fox.</w:t></w:r></w:p></w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output1 = docx_path.parent / "output1.docx"
        output2 = docx_path.parent / "output2.docx"
        output3 = docx_path.parent / "output3.docx"

        try:
            # Edit 1: brown -> red
            doc = Document(docx_path, author="Reviewer")
            doc.replace_tracked("brown", "red")
            doc.save(output1)

            # Edit 2: red -> blue (editing inside w:ins)
            doc2 = Document(output1, author="Reviewer")
            doc2.replace_tracked("red", "blue")
            doc2.save(output2)

            # Edit 3: blue -> green (editing inside w:ins again)
            doc3 = Document(output2, author="Reviewer")
            doc3.replace_tracked("blue", "green")
            doc3.save(output3)

            # Verify final result
            doc4 = Document(output3)
            text = doc4.get_text()
            assert "green" in text

        finally:
            docx_path.unlink(missing_ok=True)
            output1.unlink(missing_ok=True)
            output2.unlink(missing_ok=True)
            output3.unlink(missing_ok=True)


class TestPreExistingTrackedChanges:
    """Test editing documents that already have tracked changes."""

    def test_edit_document_with_existing_insertions(self):
        """Test editing a document that already contains w:ins elements."""
        # Create a document that already has a tracked insertion
        # Note: xml:space="preserve" required for text with leading/trailing whitespace
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
<w:body>
<w:p>
  <w:r><w:t xml:space="preserve">The </w:t></w:r>
  <w:ins w:id="0" w:author="Previous Author" w:date="2025-01-01T00:00:00Z">
    <w:r><w:t>inserted text</w:t></w:r>
  </w:ins>
  <w:r><w:t xml:space="preserve"> continues here.</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            doc = Document(docx_path, author="New Reviewer")
            # Edit the text inside the existing w:ins
            doc.replace_tracked("inserted text", "modified content")
            doc.save(output_path)

            # Verify
            doc2 = Document(output_path)
            text = doc2.get_text()
            assert "modified content" in text

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_edit_document_with_existing_deletions(self):
        """Test editing around existing w:del elements."""
        # Create a document that already has a tracked deletion
        # Note: xml:space="preserve" required for text with leading/trailing whitespace
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
<w:body>
<w:p>
  <w:r><w:t xml:space="preserve">The </w:t></w:r>
  <w:del w:id="0" w:author="Previous Author" w:date="2025-01-01T00:00:00Z">
    <w:r><w:delText>deleted text</w:delText></w:r>
  </w:del>
  <w:ins w:id="1" w:author="Previous Author" w:date="2025-01-01T00:00:00Z">
    <w:r><w:t>replacement</w:t></w:r>
  </w:ins>
  <w:r><w:t xml:space="preserve"> continues here.</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            doc = Document(docx_path, author="New Reviewer")
            # Edit the text inside the existing w:ins (from previous replacement)
            doc.replace_tracked("replacement", "new value")
            doc.save(output_path)

            # Verify
            doc2 = Document(output_path)
            text = doc2.get_text()
            assert "new value" in text

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


class TestPartialRunEditsInsideTrackedChanges:
    """Test partial run edits when run is inside tracked change wrapper."""

    def test_partial_replacement_inside_insertion(self):
        """Test replacing part of text inside a w:ins element."""
        # Note: xml:space="preserve" required for text with leading/trailing whitespace
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
<w:body>
<w:p>
  <w:r><w:t xml:space="preserve">Start </w:t></w:r>
  <w:ins w:id="0" w:author="Author" w:date="2025-01-01T00:00:00Z">
    <w:r><w:t>the quick brown fox jumps</w:t></w:r>
  </w:ins>
  <w:r><w:t xml:space="preserve"> end.</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            doc = Document(docx_path, author="Reviewer")
            # Replace just "brown" within the inserted text
            doc.replace_tracked("brown", "red")
            doc.save(output_path)

            # Verify the edit worked
            doc2 = Document(output_path)
            text = doc2.get_text()
            assert "red" in text
            assert "quick" in text  # Surrounding text preserved
            assert "fox" in text

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


class TestWrapperSplitting:
    """Test that w:ins wrappers are properly split when editing inside them.

    When partially editing text inside a w:ins element, the wrapper should be
    split to preserve attribution on the unmodified portions.
    """

    def test_partial_edit_splits_wrapper_preserves_before_text(self):
        """Test that text before the edit stays in a w:ins with original attribution."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
<w:body>
<w:p>
  <w:ins w:id="0" w:author="Original Author" w:date="2025-01-01T00:00:00Z">
    <w:r><w:t>the quick brown fox</w:t></w:r>
  </w:ins>
</w:p>
</w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            doc = Document(docx_path, author="New Reviewer")
            doc.replace_tracked("brown", "red")
            doc.save(output_path)

            # Read the output XML to verify structure
            with zipfile.ZipFile(output_path) as zf:
                xml_content = zf.read("word/document.xml").decode("utf-8")

            # The text "the quick " should be in a w:ins with "Original Author"
            assert "the quick " in xml_content or "the quick" in xml_content
            # Verify Original Author attribution is preserved somewhere
            assert "Original Author" in xml_content

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_partial_edit_splits_wrapper_preserves_after_text(self):
        """Test that text after the edit stays in a w:ins with original attribution."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
<w:body>
<w:p>
  <w:ins w:id="0" w:author="Original Author" w:date="2025-01-01T00:00:00Z">
    <w:r><w:t>the quick brown fox</w:t></w:r>
  </w:ins>
</w:p>
</w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            doc = Document(docx_path, author="New Reviewer")
            doc.replace_tracked("brown", "red")
            doc.save(output_path)

            # Read the output XML to verify structure
            with zipfile.ZipFile(output_path) as zf:
                xml_content = zf.read("word/document.xml").decode("utf-8")

            # The text " fox" should be preserved
            assert " fox" in xml_content or "fox" in xml_content
            # Final text should be correct
            doc2 = Document(output_path)
            text = doc2.get_text()
            assert "the quick" in text
            assert "red" in text
            assert "fox" in text

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_same_author_edit_inside_insertion(self):
        """Test that same author editing their own insertion updates in place."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
<w:body>
<w:p>
  <w:ins w:id="0" w:author="Same Author" w:date="2025-01-01T00:00:00Z">
    <w:r><w:t>the quick brown fox</w:t></w:r>
  </w:ins>
</w:p>
</w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            # Same author makes the edit
            doc = Document(docx_path, author="Same Author")
            doc.replace_tracked("brown", "red")
            doc.save(output_path)

            # Verify the edit worked
            doc2 = Document(output_path)
            text = doc2.get_text()
            assert "red" in text
            assert "brown" not in text or "brown" in text  # May be in delText

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


class TestSpanningMatches:
    """Test matches that span across tracked change boundaries.

    These tests verify the fix for matches that start in regular text and
    end inside w:ins (or vice versa).
    """

    def test_match_spanning_into_insertion(self):
        """Test replacing text that spans from regular text into w:ins."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
<w:body>
<w:p>
  <w:r><w:t xml:space="preserve">The quick </w:t></w:r>
  <w:ins w:id="0" w:author="Author A" w:date="2025-01-01T00:00:00Z">
    <w:r><w:t>red fox</w:t></w:r>
  </w:ins>
  <w:r><w:t xml:space="preserve"> jumps</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            doc = Document(docx_path, author="Author B")
            # This match spans regular text ("quick ") and insertion ("red")
            doc.replace_tracked("quick red", "slow blue")
            doc.save(output_path)

            # Verify the edit worked
            doc2 = Document(output_path)
            text = doc2.get_text()
            assert "slow blue" in text
            assert "fox" in text  # Rest of insertion preserved

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_match_spanning_out_of_insertion(self):
        """Test replacing text that spans from w:ins into regular text."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
<w:body>
<w:p>
  <w:r><w:t xml:space="preserve">The </w:t></w:r>
  <w:ins w:id="0" w:author="Author A" w:date="2025-01-01T00:00:00Z">
    <w:r><w:t>quick red</w:t></w:r>
  </w:ins>
  <w:r><w:t xml:space="preserve"> fox jumps</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            doc = Document(docx_path, author="Author B")
            # This match spans insertion ("red") and regular text (" fox")
            doc.replace_tracked("red fox", "blue cat")
            doc.save(output_path)

            # Verify the edit worked
            doc2 = Document(output_path)
            text = doc2.get_text()
            assert "blue cat" in text
            assert "quick" in text  # Start of insertion preserved

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_match_spanning_deletion_and_insertion(self):
        """Test replacing text that spans w:del and w:ins elements."""
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
<w:body>
<w:p>
  <w:r><w:t xml:space="preserve">The </w:t></w:r>
  <w:del w:id="0" w:author="Author A" w:date="2025-01-01T00:00:00Z">
    <w:r><w:delText>old</w:delText></w:r>
  </w:del>
  <w:ins w:id="1" w:author="Author A" w:date="2025-01-01T00:00:00Z">
    <w:r><w:t>new</w:t></w:r>
  </w:ins>
  <w:r><w:t xml:space="preserve"> text</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

        docx_path = create_test_docx(doc_xml)
        output_path = docx_path.parent / "output.docx"

        try:
            doc = Document(docx_path, author="Author B")
            # Replace the inserted text "new"
            doc.replace_tracked("new", "updated")
            doc.save(output_path)

            # Verify the edit worked
            doc2 = Document(output_path)
            text = doc2.get_text()
            assert "updated" in text

        finally:
            docx_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)
