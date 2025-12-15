"""
Tests for Document.find_all() method.

Tests the ability to find and preview all text matches with rich metadata
including context, location, and paragraph information.
"""

import re
import tempfile
import zipfile
from pathlib import Path

import pytest

from python_docx_redline import Document
from python_docx_redline.match import Match


def create_test_document(text: str = "This is a test document.") -> Path:
    """Create a simple test document with the given text."""
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
    """Create a test document with multiple paragraphs."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>This is production products text in paragraph 1.</w:t></w:r></w:p>
<w:p><w:r><w:t>Another paragraph with production products mentioned here.</w:t></w:r></w:p>
<w:p><w:r><w:t>A third occurrence of production products in this document.</w:t></w:r></w:p>
<w:p><w:r><w:t>This paragraph mentions Production Products with capital letters.</w:t></w:r></w:p>
<w:p><w:r><w:t>Completely unrelated text without the search term.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


def create_document_with_tables() -> Path:
    """Create a test document with tables."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Body text with keyword here.</w:t></w:r></w:p>
<w:tbl>
<w:tr>
<w:tc><w:p><w:r><w:t>Cell 1 with keyword</w:t></w:r></w:p></w:tc>
<w:tc><w:p><w:r><w:t>Cell 2 text</w:t></w:r></w:p></w:tc>
</w:tr>
<w:tr>
<w:tc><w:p><w:r><w:t>Cell 3 text</w:t></w:r></w:p></w:tc>
<w:tc><w:p><w:r><w:t>Cell 4 with keyword</w:t></w:r></w:p></w:tc>
</w:tr>
</w:tbl>
<w:p><w:r><w:t>More body text with keyword at the end.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


def test_find_all_basic() -> None:
    """Test basic find_all() functionality."""
    doc_path = create_test_document("This is a test document.")
    try:
        doc = Document(doc_path)

        matches = doc.find_all("test document")

        assert len(matches) == 1
        assert isinstance(matches[0], Match)
        assert matches[0].text == "test document"
        assert matches[0].index == 0
        assert "test document" in matches[0].context
        assert matches[0].location == "body"
        assert matches[0].span is not None
    finally:
        doc_path.unlink()


def test_find_all_no_matches() -> None:
    """Test find_all() when no matches are found."""
    doc_path = create_test_document("This is a test document.")
    try:
        doc = Document(doc_path)

        matches = doc.find_all("nonexistent text")

        assert matches == []
        assert len(matches) == 0
    finally:
        doc_path.unlink()


def test_find_all_multiple_matches() -> None:
    """Test find_all() with multiple occurrences."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        # Case-sensitive by default, so finds 3 lowercase matches
        matches = doc.find_all("production products")

        assert len(matches) == 3
        assert all(isinstance(m, Match) for m in matches)
        assert [m.index for m in matches] == [0, 1, 2]
        assert all(m.text == "production products" for m in matches)
    finally:
        doc_path.unlink()


def test_find_all_case_sensitive() -> None:
    """Test case-sensitive vs case-insensitive search."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        # Case-sensitive (default)
        matches_sensitive = doc.find_all("production products", case_sensitive=True)
        assert len(matches_sensitive) == 3  # Misses "Production Products"

        # Case-insensitive
        matches_insensitive = doc.find_all("production products", case_sensitive=False)
        assert len(matches_insensitive) == 4  # Finds all including "Production Products"
    finally:
        doc_path.unlink()


def test_find_all_match_properties() -> None:
    """Test that Match objects have all expected properties."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        matches = doc.find_all("production products")
        assert len(matches) > 0

        match = matches[0]

        # Check all properties exist
        assert hasattr(match, "index")
        assert hasattr(match, "text")
        assert hasattr(match, "context")
        assert hasattr(match, "paragraph_index")
        assert hasattr(match, "paragraph_text")
        assert hasattr(match, "location")
        assert hasattr(match, "span")

        # Check property types
        assert isinstance(match.index, int)
        assert isinstance(match.text, str)
        assert isinstance(match.context, str)
        assert isinstance(match.paragraph_index, int)
        assert isinstance(match.paragraph_text, str)
        assert isinstance(match.location, str)

        # Check property values
        assert match.text == "production products"
        assert "production products" in match.context
        assert match.location == "body"
        assert match.paragraph_index >= 0
    finally:
        doc_path.unlink()


def test_find_all_context() -> None:
    """Test that Match.context contains surrounding text."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        matches = doc.find_all("production products")

        # First match
        assert "This is" in matches[0].context
        assert "production products" in matches[0].context
        assert "text in paragraph 1" in matches[0].context

        # Second match
        assert "Another paragraph" in matches[1].context
        assert "production products" in matches[1].context
        assert "mentioned here" in matches[1].context
    finally:
        doc_path.unlink()


def test_find_all_paragraph_index() -> None:
    """Test that Match.paragraph_index is correct."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        matches = doc.find_all("production products")

        # Should be in paragraphs 0, 1, 2 (case-sensitive)
        assert matches[0].paragraph_index == 0
        assert matches[1].paragraph_index == 1
        assert matches[2].paragraph_index == 2
    finally:
        doc_path.unlink()


def test_find_all_paragraph_text() -> None:
    """Test that Match.paragraph_text contains full paragraph text."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        matches = doc.find_all("production products")

        # Full paragraph text
        assert matches[0].paragraph_text == "This is production products text in paragraph 1."
        assert (
            matches[1].paragraph_text
            == "Another paragraph with production products mentioned here."
        )
    finally:
        doc_path.unlink()


def test_find_all_location_body() -> None:
    """Test that matches in body have correct location."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        matches = doc.find_all("production products")

        # All matches should be in body
        assert all(m.location == "body" for m in matches)
    finally:
        doc_path.unlink()


def test_find_all_location_table() -> None:
    """Test that matches in tables have correct location."""
    doc_path = create_document_with_tables()
    try:
        doc = Document(doc_path)

        matches = doc.find_all("keyword")

        # Should find 4 matches: body, table, table, body
        assert len(matches) == 4

        # First match is in body
        assert matches[0].location == "body"

        # Second and third matches are in table cells
        assert "table" in matches[1].location
        assert "table" in matches[2].location

        # Fourth match is in body
        assert matches[3].location == "body"
    finally:
        doc_path.unlink()


def test_find_all_regex_basic() -> None:
    """Test find_all() with regex patterns."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        # Find "production products" or "Production Products"
        # Need to use [Pp] for both words to match capitals
        matches = doc.find_all(r"[Pp]roduction [Pp]roducts", regex=True)

        # Should find 4 occurrences
        assert len(matches) == 4
    finally:
        doc_path.unlink()


def test_find_all_regex_patterns() -> None:
    """Test find_all() with more complex regex patterns."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>The contract expires in 30 days.</w:t></w:r></w:p>
<w:p><w:r><w:t>Payment due in 45 days from signing.</w:t></w:r></w:p>
<w:p><w:r><w:t>No deadline specified for this item.</w:t></w:r></w:p>
<w:p><w:r><w:t>Another 7 days notice required.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    try:
        doc = Document(doc_path)

        # Find patterns like "30 days", "45 days", "7 days"
        matches = doc.find_all(r"\d+ days", regex=True)

        assert len(matches) == 3
        assert matches[0].text == "30 days"
        assert matches[1].text == "45 days"
        assert matches[2].text == "7 days"
    finally:
        doc_path.unlink()


def test_find_all_regex_invalid_pattern() -> None:
    """Test find_all() with invalid regex pattern."""
    doc_path = create_test_document("This is a test document.")
    try:
        doc = Document(doc_path)

        # Invalid regex pattern should raise re.error
        with pytest.raises(re.error):
            doc.find_all(r"[invalid(", regex=True)
    finally:
        doc_path.unlink()


def test_find_all_custom_context_chars() -> None:
    """Test find_all() with custom context_chars."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    long_text = "x" * 100 + " keyword " + "y" * 100
    document_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>{long_text}</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    try:
        doc = Document(doc_path)

        # Small context window
        matches_small = doc.find_all("keyword", context_chars=10)
        assert len(matches_small) == 1

        # Large context window
        matches_large = doc.find_all("keyword", context_chars=100)
        assert len(matches_large) == 1

        # Large should be longer than small
        assert len(matches_large[0].context) > len(matches_small[0].context)
    finally:
        doc_path.unlink()


def test_find_all_scope_string() -> None:
    """Test find_all() with string scope."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Text in section A with keyword here.</w:t></w:r></w:p>
<w:p><w:r><w:t>Text in section B with keyword too.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    try:
        doc = Document(doc_path)

        # Find "keyword" only in paragraphs containing "section A"
        matches = doc.find_all("keyword", scope="section A")

        # Should only find the one in section A paragraph
        assert len(matches) == 1
        assert "section A" in matches[0].paragraph_text
    finally:
        doc_path.unlink()


def test_find_all_scope_dict() -> None:
    """Test find_all() with dictionary scope."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Introduction with keyword here.</w:t></w:r></w:p>
<w:p><w:r><w:t>Important note: keyword appears here.</w:t></w:r></w:p>
<w:p><w:r><w:t>Conclusion with keyword here too.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    try:
        doc = Document(doc_path)

        # Find "keyword" only in paragraphs containing "Important"
        matches = doc.find_all("keyword", scope={"contains": "Important"})

        # Should only find the one in the "Important note" paragraph
        assert len(matches) == 1
        assert "Important" in matches[0].paragraph_text
    finally:
        doc_path.unlink()


def test_find_all_match_str_repr() -> None:
    """Test Match __str__ and __repr__ methods."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        matches = doc.find_all("production products")

        # __str__ should be user-friendly
        str_repr = str(matches[0])
        assert "[0]" in str_repr  # Index
        assert "body" in str_repr  # Location

        # __repr__ should be detailed
        repr_str = repr(matches[0])
        assert "Match(" in repr_str
        assert "index=0" in repr_str
        assert "location='body'" in repr_str
    finally:
        doc_path.unlink()


def test_find_all_empty_document() -> None:
    """Test find_all() on empty document."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    try:
        doc = Document(doc_path)

        matches = doc.find_all("anything")
        assert matches == []
    finally:
        doc_path.unlink()


def test_find_all_single_match() -> None:
    """Test find_all() when there's only one match."""
    doc_path = create_test_document("Only one unique occurrence of the search term.")
    try:
        doc = Document(doc_path)

        matches = doc.find_all("unique occurrence")

        assert len(matches) == 1
        assert matches[0].index == 0
        assert matches[0].text == "unique occurrence"
    finally:
        doc_path.unlink()


def test_find_all_workflow_preview_before_replace() -> None:
    """Test using find_all() to preview before replace_tracked()."""
    doc_path = create_multi_paragraph_document()
    try:
        doc = Document(doc_path)

        # Step 1: Preview all matches
        matches = doc.find_all("production products")

        # Step 2: User sees 3 matches
        assert len(matches) == 3

        # Step 3: User can inspect each match
        for match in matches:
            # In real use, user would print or display these
            assert match.location == "body"
            assert "production products" in match.context
    finally:
        doc_path.unlink()
