"""
Test the comments reading API.

Tests the Comment and CommentRange models, as well as the Document.comments
property and get_comments() method.
"""

import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import pytest

from python_docx_redline import Document


def create_document_with_comments() -> Path:
    """Create a test document with comments."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Word document with multiple comments
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Introduction</w:t></w:r>
</w:p>
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r><w:t>This is the first paragraph with a comment.</w:t></w:r>
  <w:commentRangeEnd w:id="0"/>
  <w:r>
    <w:commentReference w:id="0"/>
  </w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Main Content</w:t></w:r>
</w:p>
<w:p>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t>Another paragraph</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
  <w:r>
    <w:commentReference w:id="1"/>
  </w:r>
  <w:r><w:t> with more text after the comment.</w:t></w:r>
</w:p>
<w:p>
  <w:commentRangeStart w:id="2"/>
  <w:r><w:t>Third paragraph by a different author.</w:t></w:r>
  <w:commentRangeEnd w:id="2"/>
  <w:r>
    <w:commentReference w:id="2"/>
  </w:r>
</w:p>
</w:body>
</w:document>"""

    comments_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="John Doe" w:initials="JD" w:date="2025-01-15T10:30:00Z">
    <w:p>
      <w:r><w:t>Please review this section.</w:t></w:r>
    </w:p>
  </w:comment>
  <w:comment w:id="1" w:author="John Doe" w:initials="JD" w:date="2025-01-15T11:00:00Z">
    <w:p>
      <w:r><w:t>This needs more detail.</w:t></w:r>
    </w:p>
  </w:comment>
  <w:comment w:id="2" w:author="Jane Smith" w:initials="JS" w:date="2025-01-16T09:00:00Z">
    <w:p>
      <w:r><w:t>Looks good to me!</w:t></w:r>
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

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/comments.xml", comments_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)

    return doc_path


def create_document_without_comments() -> Path:
    """Create a test document without comments."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Simple text without comments.</w:t></w:r></w:p>
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


class TestCommentsProperty:
    """Tests for Document.comments property."""

    def test_comments_returns_list(self) -> None:
        """Test that comments property returns a list."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments = doc.comments
            assert isinstance(comments, list)
        finally:
            doc_path.unlink()

    def test_comments_count(self) -> None:
        """Test that all comments are returned."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments = doc.comments
            assert len(comments) == 3
        finally:
            doc_path.unlink()

    def test_comments_empty_for_no_comments(self) -> None:
        """Test that comments returns empty list when no comments exist."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comments = doc.comments
            assert comments == []
        finally:
            doc_path.unlink()

    def test_comment_attributes(self) -> None:
        """Test that comment attributes are correctly extracted."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments = doc.comments

            # Check first comment
            comment = comments[0]
            assert comment.id == "0"
            assert comment.author == "John Doe"
            assert comment.initials == "JD"
            assert comment.text == "Please review this section."
        finally:
            doc_path.unlink()

    def test_comment_date_parsing(self) -> None:
        """Test that comment dates are correctly parsed."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments = doc.comments

            comment = comments[0]
            assert comment.date is not None
            assert isinstance(comment.date, datetime)
            assert comment.date.year == 2025
            assert comment.date.month == 1
            assert comment.date.day == 15
            assert comment.date.hour == 10
            assert comment.date.minute == 30
        finally:
            doc_path.unlink()

    def test_comment_marked_text(self) -> None:
        """Test that marked text is correctly extracted."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments = doc.comments

            # First comment marks "This is the first paragraph with a comment."
            assert comments[0].marked_text == "This is the first paragraph with a comment."

            # Second comment marks "Another paragraph"
            assert comments[1].marked_text == "Another paragraph"

            # Third comment marks "Third paragraph by a different author."
            assert comments[2].marked_text == "Third paragraph by a different author."
        finally:
            doc_path.unlink()

    def test_comment_range_paragraphs(self) -> None:
        """Test that comment range includes paragraph references."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments = doc.comments

            for comment in comments:
                assert comment.range is not None
                assert comment.range.start_paragraph is not None
                assert comment.range.end_paragraph is not None
        finally:
            doc_path.unlink()


class TestGetComments:
    """Tests for Document.get_comments() method."""

    def test_get_comments_no_filter(self) -> None:
        """Test get_comments without filters returns all comments."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments = doc.get_comments()
            assert len(comments) == 3
        finally:
            doc_path.unlink()

    def test_get_comments_filter_by_author(self) -> None:
        """Test filtering comments by author."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)

            john_comments = doc.get_comments(author="John Doe")
            assert len(john_comments) == 2
            assert all(c.author == "John Doe" for c in john_comments)

            jane_comments = doc.get_comments(author="Jane Smith")
            assert len(jane_comments) == 1
            assert jane_comments[0].author == "Jane Smith"
        finally:
            doc_path.unlink()

    def test_get_comments_filter_by_nonexistent_author(self) -> None:
        """Test filtering by non-existent author returns empty list."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments = doc.get_comments(author="Nobody")
            assert comments == []
        finally:
            doc_path.unlink()

    def test_get_comments_filter_by_scope_section(self) -> None:
        """Test filtering comments by section scope."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)

            # Get comments in Introduction section
            intro_comments = doc.get_comments(scope="section:Introduction")
            assert len(intro_comments) == 1
            assert "first paragraph" in intro_comments[0].marked_text

            # Get comments in Main Content section
            main_comments = doc.get_comments(scope="section:Main Content")
            assert len(main_comments) == 2
        finally:
            doc_path.unlink()

    def test_get_comments_combined_filters(self) -> None:
        """Test combining author and scope filters."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)

            # John's comments in Main Content
            comments = doc.get_comments(author="John Doe", scope="section:Main Content")
            assert len(comments) == 1
            assert comments[0].text == "This needs more detail."
        finally:
            doc_path.unlink()


class TestCommentModel:
    """Tests for the Comment model class."""

    def test_comment_repr(self) -> None:
        """Test Comment __repr__ method."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comment = doc.comments[0]
            repr_str = repr(comment)
            assert "Comment" in repr_str
            assert "id=0" in repr_str
            assert "John Doe" in repr_str
        finally:
            doc_path.unlink()

    def test_comment_equality(self) -> None:
        """Test Comment equality based on ID."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments1 = doc.comments
            comments2 = doc.comments  # Get again

            # Same ID should be equal
            assert comments1[0] == comments2[0]

            # Different IDs should not be equal
            assert comments1[0] != comments1[1]
        finally:
            doc_path.unlink()

    def test_comment_hash(self) -> None:
        """Test Comment can be used in sets."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comments = doc.comments

            # Should be usable in a set
            comment_set = set(comments)
            assert len(comment_set) == 3
        finally:
            doc_path.unlink()


class TestCommentRange:
    """Tests for the CommentRange dataclass."""

    def test_comment_range_attributes(self) -> None:
        """Test CommentRange has expected attributes."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)
            comment = doc.comments[0]

            assert comment.range is not None
            assert hasattr(comment.range, "start_paragraph")
            assert hasattr(comment.range, "end_paragraph")
            assert hasattr(comment.range, "marked_text")
        finally:
            doc_path.unlink()


class TestEdgeCases:
    """Tests for edge cases and error handling."""

    def test_comment_without_date(self) -> None:
        """Test handling comments without date attribute."""
        doc_path = Path(tempfile.mktemp(suffix=".docx"))

        document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r><w:t>Text with comment.</w:t></w:r>
  <w:commentRangeEnd w:id="0"/>
  <w:r><w:commentReference w:id="0"/></w:r>
</w:p>
</w:body>
</w:document>"""

        # Comment without date
        comments_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Test User">
    <w:p><w:r><w:t>No date comment.</w:t></w:r></w:p>
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

        root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

        with zipfile.ZipFile(doc_path, "w") as docx:
            docx.writestr("[Content_Types].xml", content_types_xml)
            docx.writestr("_rels/.rels", root_rels)
            docx.writestr("word/document.xml", document_xml)
            docx.writestr("word/comments.xml", comments_xml)
            docx.writestr("word/_rels/document.xml.rels", rels_xml)

        try:
            doc = Document(doc_path)
            comments = doc.comments

            assert len(comments) == 1
            assert comments[0].date is None
            assert comments[0].author == "Test User"
        finally:
            doc_path.unlink()

    def test_comment_without_initials(self) -> None:
        """Test handling comments without initials attribute."""
        doc_path = Path(tempfile.mktemp(suffix=".docx"))

        document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r><w:t>Text.</w:t></w:r>
  <w:commentRangeEnd w:id="0"/>
  <w:r><w:commentReference w:id="0"/></w:r>
</w:p>
</w:body>
</w:document>"""

        comments_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Test User">
    <w:p><w:r><w:t>No initials.</w:t></w:r></w:p>
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

        root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

        with zipfile.ZipFile(doc_path, "w") as docx:
            docx.writestr("[Content_Types].xml", content_types_xml)
            docx.writestr("_rels/.rels", root_rels)
            docx.writestr("word/document.xml", document_xml)
            docx.writestr("word/comments.xml", comments_xml)
            docx.writestr("word/_rels/document.xml.rels", rels_xml)

        try:
            doc = Document(doc_path)
            comments = doc.comments

            assert len(comments) == 1
            assert comments[0].initials is None
        finally:
            doc_path.unlink()

    def test_point_comment_no_range_end(self) -> None:
        """Test handling point comments without range end marker."""
        doc_path = Path(tempfile.mktemp(suffix=".docx"))

        # Document with only commentRangeStart (point comment)
        document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r><w:t>Text here.</w:t></w:r>
  <w:r><w:commentReference w:id="0"/></w:r>
</w:p>
</w:body>
</w:document>"""

        comments_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Test User">
    <w:p><w:r><w:t>Point comment.</w:t></w:r></w:p>
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

        root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

        with zipfile.ZipFile(doc_path, "w") as docx:
            docx.writestr("[Content_Types].xml", content_types_xml)
            docx.writestr("_rels/.rels", root_rels)
            docx.writestr("word/document.xml", document_xml)
            docx.writestr("word/comments.xml", comments_xml)
            docx.writestr("word/_rels/document.xml.rels", rels_xml)

        try:
            doc = Document(doc_path)
            comments = doc.comments

            assert len(comments) == 1
            # Point comments have empty marked text
            assert comments[0].marked_text == ""
        finally:
            doc_path.unlink()

    def test_multiline_comment_text(self) -> None:
        """Test comments with multiple paragraphs."""
        doc_path = Path(tempfile.mktemp(suffix=".docx"))

        document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r><w:t>Marked text.</w:t></w:r>
  <w:commentRangeEnd w:id="0"/>
  <w:r><w:commentReference w:id="0"/></w:r>
</w:p>
</w:body>
</w:document>"""

        # Comment with multiple paragraphs
        comments_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Test User">
    <w:p><w:r><w:t>First line.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second line.</w:t></w:r></w:p>
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

        root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

        with zipfile.ZipFile(doc_path, "w") as docx:
            docx.writestr("[Content_Types].xml", content_types_xml)
            docx.writestr("_rels/.rels", root_rels)
            docx.writestr("word/document.xml", document_xml)
            docx.writestr("word/comments.xml", comments_xml)
            docx.writestr("word/_rels/document.xml.rels", rels_xml)

        try:
            doc = Document(doc_path)
            comments = doc.comments

            assert len(comments) == 1
            # Text from all paragraphs is concatenated
            assert comments[0].text == "First line.Second line."
        finally:
            doc_path.unlink()


class TestAddComment:
    """Tests for Document.add_comment() method."""

    def test_add_comment_basic(self) -> None:
        """Test basic comment creation."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path, author="Test Author")

            comment = doc.add_comment(
                "This is my comment",
                on="Simple text without comments.",
            )

            # Verify returned comment
            assert comment.text == "This is my comment"
            assert comment.author == "Test Author"
            assert comment.marked_text == "Simple text without comments."

            # Verify comment appears in doc.comments
            comments = doc.comments
            assert len(comments) == 1
            assert comments[0].text == "This is my comment"

        finally:
            doc_path.unlink()

    def test_add_comment_custom_author(self) -> None:
        """Test comment creation with custom author."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            comment = doc.add_comment(
                "Review this",
                on="Simple text",
                author="Jane Reviewer",
            )

            assert comment.author == "Jane Reviewer"
            assert comment.initials == "JR"  # Auto-generated

        finally:
            doc_path.unlink()

    def test_add_comment_custom_initials(self) -> None:
        """Test comment creation with custom initials."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            comment = doc.add_comment(
                "My note",
                on="Simple text",
                author="John Doe",
                initials="JXD",
            )

            assert comment.author == "John Doe"
            assert comment.initials == "JXD"

        finally:
            doc_path.unlink()

    def test_add_comment_has_date(self) -> None:
        """Test that created comments have timestamps."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            comment = doc.add_comment(
                "Timestamped comment",
                on="Simple text",
            )

            assert comment.date is not None
            # Should be recent (within last minute)
            from datetime import datetime, timezone

            now = datetime.now(timezone.utc)
            diff = abs((now - comment.date).total_seconds())
            assert diff < 60

        finally:
            doc_path.unlink()

    def test_add_comment_to_document_with_existing_comments(self) -> None:
        """Test adding comment to document that already has comments."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)

            # Should have 3 existing comments
            assert len(doc.comments) == 3

            # Add a new comment
            comment = doc.add_comment(
                "New comment",
                on="more text after",
                author="New Reviewer",
            )

            # Should now have 4 comments
            comments = doc.comments
            assert len(comments) == 4

            # New comment should have ID > existing max (which is 2)
            assert int(comment.id) >= 3

        finally:
            doc_path.unlink()

    def test_add_comment_creates_comments_xml(self) -> None:
        """Test that add_comment creates comments.xml if it doesn't exist."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            # Add comment
            doc.add_comment("Test comment", on="Simple text")

            # Save and verify comments.xml exists
            output_path = Path(tempfile.mktemp(suffix=".docx"))
            doc.save(output_path)

            with zipfile.ZipFile(output_path, "r") as docx:
                assert "word/comments.xml" in docx.namelist()

            output_path.unlink()

        finally:
            doc_path.unlink()

    def test_add_comment_creates_relationship(self) -> None:
        """Test that add_comment creates relationship entry."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            doc.add_comment("Test comment", on="Simple text")

            output_path = Path(tempfile.mktemp(suffix=".docx"))
            doc.save(output_path)

            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/document.xml.rels")
                assert b"comments" in rels_content.lower()

            output_path.unlink()

        finally:
            doc_path.unlink()

    def test_add_comment_creates_content_type(self) -> None:
        """Test that add_comment creates content type entry."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            doc.add_comment("Test comment", on="Simple text")

            output_path = Path(tempfile.mktemp(suffix=".docx"))
            doc.save(output_path)

            with zipfile.ZipFile(output_path, "r") as docx:
                ct_content = docx.read("[Content_Types].xml")
                assert b"comments" in ct_content.lower()

            output_path.unlink()

        finally:
            doc_path.unlink()

    def test_add_comment_inserts_markers(self) -> None:
        """Test that add_comment inserts range markers in document body."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test", on="Simple text")

            # Check document body has markers
            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            range_starts = list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeStart"))
            range_ends = list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeEnd"))
            refs = list(doc.xml_root.iter(f"{{{word_ns}}}commentReference"))

            assert len(range_starts) == 1
            assert len(range_ends) == 1
            assert len(refs) == 1

            # All should have same ID
            assert range_starts[0].get(f"{{{word_ns}}}id") == comment.id
            assert range_ends[0].get(f"{{{word_ns}}}id") == comment.id
            assert refs[0].get(f"{{{word_ns}}}id") == comment.id

        finally:
            doc_path.unlink()

    def test_add_comment_text_not_found(self) -> None:
        """Test that add_comment raises error for missing text."""
        from python_docx_redline import TextNotFoundError

        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            with pytest.raises(TextNotFoundError):
                doc.add_comment("Comment", on="nonexistent text")

        finally:
            doc_path.unlink()

    def test_add_comment_ambiguous_text(self) -> None:
        """Test that add_comment raises error for ambiguous text."""
        from python_docx_redline import AmbiguousTextError

        # Create document with repeated text
        doc_path = Path(tempfile.mktemp(suffix=".docx"))

        document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Same text here.</w:t></w:r></w:p>
<w:p><w:r><w:t>Same text here.</w:t></w:r></w:p>
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
                doc.add_comment("Comment", on="Same text here.")

        finally:
            doc_path.unlink()

    def test_add_comment_with_scope(self) -> None:
        """Test adding comment with scope filter."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)

            # Use scope to target specific section
            comment = doc.add_comment(
                "Section-specific comment",
                on="Another paragraph",
                scope="section:Main Content",
            )

            assert comment.marked_text == "Another paragraph"

        finally:
            doc_path.unlink()

    def test_add_multiple_comments(self) -> None:
        """Test adding multiple comments to same document."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path, author="Reviewer")

            # We need a document with more content
            # Just test that multiple adds work on different text
            comment1 = doc.add_comment("First note", on="Simple")
            comment2 = doc.add_comment("Second note", on="without")

            comments = doc.comments
            assert len(comments) == 2

            # IDs should be sequential
            assert int(comment1.id) == 0
            assert int(comment2.id) == 1

        finally:
            doc_path.unlink()

    def test_add_comment_save_and_reload(self) -> None:
        """Test that added comments persist after save and reload."""
        doc_path = create_document_without_comments()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Create and save
            doc = Document(doc_path)
            doc.add_comment(
                "Persistent comment",
                on="Simple text without comments.",
                author="Author",
            )
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)
            comments = doc2.comments

            assert len(comments) == 1
            assert comments[0].text == "Persistent comment"
            assert comments[0].author == "Author"
            # The marked text is the full text between range markers
            assert comments[0].marked_text == "Simple text without comments."

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()


class TestCommentResolution:
    """Tests for comment resolution (resolve/unresolve) functionality."""

    def test_new_comment_is_not_resolved(self) -> None:
        """Test that newly created comments are not resolved."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test comment", on="Simple text")

            assert not comment.is_resolved

        finally:
            doc_path.unlink()

    def test_resolve_comment(self) -> None:
        """Test marking a comment as resolved."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test comment", on="Simple text")

            comment.resolve()

            assert comment.is_resolved

        finally:
            doc_path.unlink()

    def test_unresolve_comment(self) -> None:
        """Test marking a resolved comment as unresolved."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test comment", on="Simple text")

            comment.resolve()
            assert comment.is_resolved

            comment.unresolve()
            assert not comment.is_resolved

        finally:
            doc_path.unlink()

    def test_resolve_creates_comments_extended_xml(self) -> None:
        """Test that resolve() creates commentsExtended.xml if needed."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test comment", on="Simple text")
            comment.resolve()

            # Save and check file exists
            output_path = Path(tempfile.mktemp(suffix=".docx"))
            doc.save(output_path)

            with zipfile.ZipFile(output_path, "r") as docx:
                assert "word/commentsExtended.xml" in docx.namelist()

            output_path.unlink()

        finally:
            doc_path.unlink()

    def test_resolve_persists_after_save_reload(self) -> None:
        """Test that resolution status persists after save and reload."""
        doc_path = create_document_without_comments()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Create, resolve, and save
            doc = Document(doc_path)
            comment = doc.add_comment("Test comment", on="Simple text")
            comment.resolve()
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)
            comments = doc2.comments
            assert len(comments) == 1
            assert comments[0].is_resolved

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_multiple_comments_independent_resolution(self) -> None:
        """Test that resolving one comment doesn't affect others."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment1 = doc.add_comment("First", on="Simple")
            comment2 = doc.add_comment("Second", on="without")

            # Resolve only the first
            comment1.resolve()

            assert comment1.is_resolved
            assert not comment2.is_resolved

        finally:
            doc_path.unlink()

    def test_comment_para_id_exists(self) -> None:
        """Test that created comments have a paraId."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test", on="Simple text")

            # Check paraId is set
            assert comment.para_id is not None
            assert len(comment.para_id) == 8  # 8 hex chars

        finally:
            doc_path.unlink()

    def test_comment_repr_shows_resolved(self) -> None:
        """Test that resolved comments show [RESOLVED] in repr."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test", on="Simple text")

            assert "[RESOLVED]" not in repr(comment)

            comment.resolve()

            assert "[RESOLVED]" in repr(comment)

        finally:
            doc_path.unlink()


class TestCommentDeletion:
    """Tests for individual comment deletion functionality."""

    def test_delete_comment_basic(self) -> None:
        """Test basic comment deletion."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test comment", on="Simple text")

            assert len(doc.comments) == 1

            comment.delete()

            assert len(doc.comments) == 0

        finally:
            doc_path.unlink()

    def test_delete_removes_markers(self) -> None:
        """Test that delete removes comment markers from document body."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test", on="Simple text")

            # Verify markers exist
            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            assert len(list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeStart"))) == 1
            assert len(list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeEnd"))) == 1
            assert len(list(doc.xml_root.iter(f"{{{word_ns}}}commentReference"))) == 1

            comment.delete()

            # Verify markers removed
            assert len(list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeStart"))) == 0
            assert len(list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeEnd"))) == 0
            assert len(list(doc.xml_root.iter(f"{{{word_ns}}}commentReference"))) == 0

        finally:
            doc_path.unlink()

    def test_delete_one_of_multiple(self) -> None:
        """Test deleting one comment leaves others intact."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment1 = doc.add_comment("First", on="Simple")
            doc.add_comment("Second", on="without")  # comment2 kept for side effect

            assert len(doc.comments) == 2

            comment1.delete()

            comments = doc.comments
            assert len(comments) == 1
            assert comments[0].text == "Second"

        finally:
            doc_path.unlink()

    def test_delete_preserves_document_text(self) -> None:
        """Test that deleting a comment doesn't remove the annotated text."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test", on="Simple text")

            original_text = doc.get_text()

            comment.delete()

            # Text should still be there
            assert "Simple text without comments." in doc.get_text()
            assert doc.get_text() == original_text

        finally:
            doc_path.unlink()

    def test_delete_removes_comments_extended_entry(self) -> None:
        """Test that delete removes commentsExtended.xml entry."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)
            comment = doc.add_comment("Test", on="Simple text")
            para_id = comment.para_id

            # Resolve to create commentsExtended entry
            comment.resolve()

            # Verify entry exists
            comment_ex = doc._get_comment_ex(para_id)
            assert comment_ex is not None

            comment.delete()

            # Verify entry removed
            comment_ex = doc._get_comment_ex(para_id)
            assert comment_ex is None

        finally:
            doc_path.unlink()

    def test_delete_persists_after_save_reload(self) -> None:
        """Test that deletion persists after save and reload."""
        doc_path = create_document_without_comments()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Create, delete, and save
            doc = Document(doc_path)
            comment = doc.add_comment("Test", on="Simple text")
            comment.delete()
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)
            assert len(doc2.comments) == 0

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_delete_from_document_with_existing_comments(self) -> None:
        """Test deleting a comment from document with existing comments."""
        doc_path = create_document_with_comments()
        try:
            doc = Document(doc_path)

            # Should have 3 existing comments
            assert len(doc.comments) == 3

            # Delete the first comment
            doc.comments[0].delete()

            # Should now have 2 comments
            assert len(doc.comments) == 2

        finally:
            doc_path.unlink()


class TestCommentReplies:
    """Tests for comment reply (threading) functionality."""

    def test_add_reply_via_document(self) -> None:
        """Test adding a reply using doc.add_comment with reply_to parameter."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            # Create parent comment
            parent = doc.add_comment("Original comment", on="Simple text")

            # Add reply
            reply = doc.add_comment(
                "Reply to original",
                reply_to=parent,
                author="Responder",
            )

            # Both should be in comments list
            assert len(doc.comments) == 2
            assert reply.text == "Reply to original"
            assert reply.author == "Responder"

        finally:
            doc_path.unlink()

    def test_add_reply_via_comment_method(self) -> None:
        """Test adding a reply using comment.add_reply() method."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            # Create parent comment
            parent = doc.add_comment("Original", on="Simple text", author="Author1")

            # Add reply via method
            reply = parent.add_reply("My reply", author="Author2")

            assert len(doc.comments) == 2
            assert reply.text == "My reply"
            assert reply.author == "Author2"

        finally:
            doc_path.unlink()

    def test_reply_has_no_marked_text(self) -> None:
        """Test that replies don't have marked_text (they don't mark document text)."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            parent = doc.add_comment("Original", on="Simple text")
            reply = parent.add_reply("Reply text")

            # Parent should have marked text
            assert parent.marked_text is not None

            # Reply should not have marked text
            assert reply.marked_text is None

        finally:
            doc_path.unlink()

    def test_reply_parent_property(self) -> None:
        """Test that reply.parent returns the parent comment."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            parent = doc.add_comment("Original", on="Simple text")
            reply = parent.add_reply("Reply text")

            # Check parent property
            assert reply.parent is not None
            assert reply.parent.id == parent.id
            assert reply.parent.text == "Original"

        finally:
            doc_path.unlink()

    def test_top_level_comment_has_no_parent(self) -> None:
        """Test that top-level comments have no parent."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            comment = doc.add_comment("Top level", on="Simple text")

            assert comment.parent is None

        finally:
            doc_path.unlink()

    def test_parent_replies_property(self) -> None:
        """Test that parent.replies returns child comments."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            parent = doc.add_comment("Original", on="Simple text")
            parent.add_reply("First reply")  # Side effect: creates reply
            parent.add_reply("Second reply")  # Side effect: creates reply

            replies = parent.replies
            assert len(replies) == 2
            assert any(r.text == "First reply" for r in replies)
            assert any(r.text == "Second reply" for r in replies)

        finally:
            doc_path.unlink()

    def test_comment_without_replies(self) -> None:
        """Test that comments without replies return empty list."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            comment = doc.add_comment("Lonely comment", on="Simple text")

            assert comment.replies == []

        finally:
            doc_path.unlink()

    def test_reply_to_by_comment_id_string(self) -> None:
        """Test replying using comment ID string."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            parent = doc.add_comment("Original", on="Simple text")
            parent_id = parent.id

            # Reply using ID string
            reply = doc.add_comment("Reply via ID", reply_to=parent_id)

            assert reply.parent.id == parent_id

        finally:
            doc_path.unlink()

    def test_reply_to_by_comment_id_int(self) -> None:
        """Test replying using comment ID as integer."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            parent = doc.add_comment("Original", on="Simple text")
            parent_id = int(parent.id)

            # Reply using ID as int
            reply = doc.add_comment("Reply via int ID", reply_to=parent_id)

            assert reply.parent.id == str(parent_id)

        finally:
            doc_path.unlink()

    def test_reply_to_nonexistent_comment(self) -> None:
        """Test that replying to nonexistent comment raises error."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError, match="not found"):
                doc.add_comment("Reply", reply_to="999")

        finally:
            doc_path.unlink()

    def test_reply_persists_after_save_reload(self) -> None:
        """Test that replies persist after save and reload."""
        doc_path = create_document_without_comments()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Create parent and reply, then save
            doc = Document(doc_path)
            parent = doc.add_comment("Original", on="Simple text", author="Author1")
            parent.add_reply("Reply text", author="Author2")
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)
            comments = doc2.comments

            assert len(comments) == 2

            # Find parent and reply
            parent2 = next(c for c in comments if c.text == "Original")
            reply2 = next(c for c in comments if c.text == "Reply text")

            # Verify threading is preserved
            assert reply2.parent is not None
            assert reply2.parent.id == parent2.id
            assert len(parent2.replies) == 1
            assert parent2.replies[0].id == reply2.id

        finally:
            doc_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_nested_replies(self) -> None:
        """Test replies to replies (nested threading)."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            parent = doc.add_comment("Top level", on="Simple text")
            reply1 = parent.add_reply("First reply")
            reply2 = reply1.add_reply("Reply to reply")

            assert len(doc.comments) == 3

            # Check threading
            assert reply1.parent.id == parent.id
            assert reply2.parent.id == reply1.id

            # Parent has one direct reply
            assert len(parent.replies) == 1
            assert parent.replies[0].id == reply1.id

            # Reply1 has one reply
            assert len(reply1.replies) == 1
            assert reply1.replies[0].id == reply2.id

        finally:
            doc_path.unlink()

    def test_add_comment_requires_on_or_reply_to(self) -> None:
        """Test that add_comment requires either 'on' or 'reply_to'."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError, match="Either 'on' or 'reply_to'"):
                doc.add_comment("Comment without target")

        finally:
            doc_path.unlink()

    def test_reply_creates_comments_extended_xml(self) -> None:
        """Test that replies create commentsExtended.xml with parent linkage."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            parent = doc.add_comment("Original", on="Simple text")
            parent.add_reply("Reply")  # Side effect: creates reply

            # Save and check file structure
            output_path = Path(tempfile.mktemp(suffix=".docx"))
            doc.save(output_path)

            with zipfile.ZipFile(output_path, "r") as docx:
                assert "word/commentsExtended.xml" in docx.namelist()

                # Check content has paraIdParent
                content = docx.read("word/commentsExtended.xml")
                assert b"paraIdParent" in content

            output_path.unlink()

        finally:
            doc_path.unlink()

    def test_reply_does_not_add_document_markers(self) -> None:
        """Test that replies don't add range markers to document body."""
        doc_path = create_document_without_comments()
        try:
            doc = Document(doc_path)

            parent = doc.add_comment("Original", on="Simple text")

            # Count markers before reply
            word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            starts_before = len(list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeStart")))
            ends_before = len(list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeEnd")))

            # Add reply
            parent.add_reply("Reply text")

            # Count markers after reply
            starts_after = len(list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeStart")))
            ends_after = len(list(doc.xml_root.iter(f"{{{word_ns}}}commentRangeEnd")))

            # Should not have added more markers
            assert starts_after == starts_before
            assert ends_after == ends_before

        finally:
            doc_path.unlink()
