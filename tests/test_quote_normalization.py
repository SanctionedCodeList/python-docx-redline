"""
Tests for smart quote normalization in text search operations.

This module tests the ability to search with straight quotes (keyboard defaults)
and match smart/curly quotes that are common in Word documents.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest

from python_docx_redline import Document
from python_docx_redline.errors import TextNotFoundError
from python_docx_redline.quote_normalization import (
    denormalize_quotes,
    has_smart_quotes,
    has_straight_quotes,
    normalize_quotes,
)


def create_test_document(text: str) -> Path:
    """Create a simple but valid test document with given text."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

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

    document_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>{text}</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document_xml)

    return doc_path


# ============================================================================
# Normalization function tests
# ============================================================================


def test_normalize_quotes_apostrophe() -> None:
    """Normalize converts curly apostrophe to straight apostrophe."""
    # Input uses curly apostrophe (')
    # Output should have straight apostrophe (')
    assert normalize_quotes("plaintiff\u2019s claim") == "plaintiff's claim"
    assert normalize_quotes("don\u2019t") == "don't"
    assert normalize_quotes("can\u2019t") == "can't"


def test_normalize_quotes_double_quote() -> None:
    """Normalize converts curly quotes to straight double quotes."""
    # Input uses curly opening and closing quotes
    # Output should have straight double quotes (")
    assert normalize_quotes("\u201cfree trial\u201d") == '"free trial"'


def test_denormalize_quotes_single() -> None:
    """Denormalize converts curly single quotes to straight apostrophe."""
    # Input with curly apostrophe, output with straight apostrophe
    assert denormalize_quotes("plaintiff\u2019s claim") == "plaintiff's claim"
    # Input with curly quotes, output with straight quotes
    assert denormalize_quotes("\u2018quote\u2019") == "'quote'"


def test_denormalize_quotes_double() -> None:
    """Denormalize converts curly double quotes to straight quotes."""
    # Input with curly quotes, output with straight quotes
    assert denormalize_quotes("\u201cfree trial\u201d") == '"free trial"'


def test_has_smart_quotes_detection() -> None:
    """has_smart_quotes detects presence of curly quotes."""
    # Text with curly apostrophe should be detected
    assert has_smart_quotes("plaintiff\u2019s claim")
    # Text with curly quotes should be detected
    assert has_smart_quotes("\u201cquoted\u201d")
    # Text with no quotes should not be detected
    assert not has_smart_quotes("straight text")
    # Text with straight quotes should not be detected
    assert not has_smart_quotes("plaintiff's claim")


def test_has_straight_quotes_detection() -> None:
    """has_straight_quotes detects presence of straight quotes."""
    assert has_straight_quotes("plaintiff's claim")
    assert has_straight_quotes('"quoted"')
    assert not has_straight_quotes("no quotes here")


def test_normalize_denormalize_roundtrip() -> None:
    """Normalize and denormalize both convert to straight quotes (idempotent)."""
    # Both normalize and denormalize convert to straight quotes, so this test
    # verifies that straight quotes remain unchanged through both operations
    original = 'plaintiff\'s "free trial" claim'
    normalized = normalize_quotes(original)
    assert normalized == original  # Already straight, no change
    denormalized = denormalize_quotes(normalized)
    assert denormalized == original


# ============================================================================
# insert_tracked() with smart quotes
# ============================================================================


def test_insert_tracked_with_smart_quotes_apostrophe() -> None:
    """insert_tracked finds text with smart apostrophes using straight quote."""
    # Document has curly apostrophe (')
    doc_path = create_test_document("The plaintiff\u2019s claim is strong.")
    try:
        doc = Document(doc_path)

        # Search with straight apostrophe (keyboard input)
        doc.insert_tracked("valid ", after="plaintiff's")

        text = doc.get_text()
        assert "valid " in text

    finally:
        doc_path.unlink()


def test_insert_tracked_with_smart_quotes_contraction() -> None:
    """insert_tracked finds contractions with smart apostrophes."""
    # Document has curly apostrophe in "don't"
    doc_path = create_test_document("They don\u2019t have standing.")
    try:
        doc = Document(doc_path)

        # Search with straight apostrophe
        doc.insert_tracked("still ", before="don't")

        text = doc.get_text()
        assert "still " in text

    finally:
        doc_path.unlink()


def test_insert_tracked_with_smart_quotes_double() -> None:
    """insert_tracked finds text with smart double quotes."""
    # Document has curly double quotes
    doc_path = create_test_document("The \u201cfree trial\u201d offer was misleading.")
    try:
        doc = Document(doc_path)

        # Search with straight double quotes
        doc.insert_tracked("so-called ", before='"free trial"')

        text = doc.get_text()
        assert "so-called" in text

    finally:
        doc_path.unlink()


def test_insert_tracked_disable_normalization() -> None:
    """insert_tracked with normalize_special_chars=False requires exact match."""
    doc_path = create_test_document("The plaintiff\u2019s claim is strong.")
    try:
        doc = Document(doc_path)

        # With normalization disabled, straight quote won't match curly quote
        with pytest.raises(TextNotFoundError):
            doc.insert_tracked("valid ", after="plaintiff's", normalize_special_chars=False)

    finally:
        doc_path.unlink()


def test_insert_tracked_normalization_with_regex() -> None:
    """Normalization is disabled when regex=True."""
    doc_path = create_test_document("Version 1.2.3 released.")
    try:
        doc = Document(doc_path)

        # Regex should still work (normalization skipped for regex)
        doc.insert_tracked("New: ", before=r"\d+\.\d+\.\d+", regex=True)

        text = doc.get_text()
        assert "New:" in text

    finally:
        doc_path.unlink()


# ============================================================================
# delete_tracked() with smart quotes
# ============================================================================


def test_delete_tracked_with_smart_quotes_apostrophe() -> None:
    """delete_tracked finds text with smart apostrophes using straight quote."""
    doc_path = create_test_document("The plaintiff\u2019s claim is strong.")
    try:
        doc = Document(doc_path)

        # Delete using straight apostrophe
        doc.delete_tracked("plaintiff's ")

        text = doc.get_text()
        # The text should show deletion markup
        assert "claim" in text

    finally:
        doc_path.unlink()


def test_delete_tracked_with_smart_quotes_double() -> None:
    """delete_tracked finds text with smart double quotes."""
    doc_path = create_test_document("The \u201cfree trial\u201d offer was misleading.")
    try:
        doc = Document(doc_path)

        # Delete using straight double quotes
        doc.delete_tracked('"free trial" ')

        text = doc.get_text()
        assert "offer" in text

    finally:
        doc_path.unlink()


def test_delete_tracked_disable_normalization() -> None:
    """delete_tracked with normalize_special_chars=False requires exact match."""
    doc_path = create_test_document("The plaintiff\u2019s claim is strong.")
    try:
        doc = Document(doc_path)

        # With normalization disabled, straight quote won't match curly quote
        with pytest.raises(TextNotFoundError):
            doc.delete_tracked("plaintiff's", normalize_special_chars=False)

    finally:
        doc_path.unlink()


# ============================================================================
# replace_tracked() with smart quotes
# ============================================================================


def test_replace_tracked_with_smart_quotes_apostrophe() -> None:
    """replace_tracked finds text with smart apostrophes using straight quote."""
    doc_path = create_test_document("The plaintiff\u2019s claim is strong.")
    try:
        doc = Document(doc_path)

        # Replace using straight apostrophe
        doc.replace_tracked("plaintiff's", "defendant's")

        text = doc.get_text()
        # Should show both deletion and insertion
        assert "defendant" in text or "defendant" in text

    finally:
        doc_path.unlink()


def test_replace_tracked_with_smart_quotes_double() -> None:
    """replace_tracked finds text with smart double quotes."""
    doc_path = create_test_document("The \u201cfree trial\u201d offer was misleading.")
    try:
        doc = Document(doc_path)

        # Replace using straight double quotes
        doc.replace_tracked('"free trial"', '"limited-time"')

        text = doc.get_text()
        assert "limited-time" in text or "limited-time" in text

    finally:
        doc_path.unlink()


def test_replace_tracked_disable_normalization() -> None:
    """replace_tracked with normalize_special_chars=False requires exact match."""
    doc_path = create_test_document("The plaintiff\u2019s claim is strong.")
    try:
        doc = Document(doc_path)

        # With normalization disabled, straight quote won't match curly quote
        with pytest.raises(TextNotFoundError):
            doc.replace_tracked("plaintiff's", "defendant's", normalize_special_chars=False)

    finally:
        doc_path.unlink()


# ============================================================================
# Edge cases and mixed quotes
# ============================================================================


def test_mixed_quotes_in_document() -> None:
    """Handle documents with both straight and smart quotes."""
    doc_path = create_test_document("The plaintiff\u2019s claim and defendant\u2019s defense.")
    try:
        doc = Document(doc_path)

        # Should find both
        doc.insert_tracked("strong ", after="plaintiff's")
        doc.insert_tracked("weak ", after="defendant's")

        text = doc.get_text()
        assert "strong" in text
        assert "weak" in text

    finally:
        doc_path.unlink()


def test_normalization_no_quotes() -> None:
    """Normalization handles text without quotes gracefully."""
    doc_path = create_test_document("Simple text without any quotes.")
    try:
        doc = Document(doc_path)

        # Should work normally
        doc.insert_tracked("Additional text. ", after="Simple text")

        text = doc.get_text()
        assert "Additional text" in text

    finally:
        doc_path.unlink()


def test_search_with_smart_quotes_in_query() -> None:
    """Searching with smart quotes in query matches smart quotes in doc."""
    doc_path = create_test_document("The plaintiff\u2019s claim is strong.")
    try:
        doc = Document(doc_path)

        # Search with smart apostrophe (if user copies from document)
        doc.insert_tracked("valid ", after="plaintiff\u2019s")

        text = doc.get_text()
        assert "valid" in text

    finally:
        doc_path.unlink()


def test_multiple_apostrophes_in_text() -> None:
    """Handle text with multiple apostrophes correctly."""
    doc_path = create_test_document("The plaintiff\u2019s attorney\u2019s argument was strong.")
    try:
        doc = Document(doc_path)

        # Even with multiple apostrophes in the text, specific searches should work
        doc.insert_tracked("strong ", after="plaintiff\u2019s attorney\u2019s")
        text = doc.get_text()
        assert "strong " in text

    finally:
        doc_path.unlink()


def test_possessive_vs_contraction() -> None:
    """Correctly handle both possessives and contractions."""
    doc_path = create_test_document(
        "It\u2019s clear that the plaintiff\u2019s claim won\u2019t succeed."
    )
    try:
        doc = Document(doc_path)

        # Test contraction
        doc.insert_tracked("very ", before="clear")

        # Test possessive
        doc.insert_tracked("weak ", after="plaintiff's")

        # Test another contraction
        doc.insert_tracked("likely ", before="won't")

        text = doc.get_text()
        assert "very" in text
        assert "weak" in text
        assert "likely" in text

    finally:
        doc_path.unlink()


# ============================================================================
# Integration tests
# ============================================================================


def test_normalization_with_scope() -> None:
    """Quote normalization works with scope filtering."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

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

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
    <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
    <w:r><w:t>First Section</w:t></w:r>
</w:p>
<w:p><w:r><w:t>The plaintiff\u2019s first claim.</w:t></w:r></w:p>
<w:p>
    <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
    <w:r><w:t>Second Section</w:t></w:r>
</w:p>
<w:p><w:r><w:t>The plaintiff\u2019s second claim.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document_xml)

    try:
        doc = Document(doc_path)

        # Insert only in Second Section using straight apostrophe
        doc.insert_tracked("strong ", after="plaintiff's", scope="section:Second Section")

        text = doc.get_text()
        assert "strong" in text

    finally:
        doc_path.unlink()


def test_all_quote_types_together() -> None:
    """Test document with all quote types."""
    # Document has curly apostrophes and curly double quotes
    doc_path = create_test_document(
        "It\u2019s the plaintiff\u2019s \u201cfree trial\u201d offer that won\u2019t hold up."
    )
    try:
        doc = Document(doc_path)

        # Test each quote type - all using straight quotes in search
        doc.insert_tracked("clearly ", before="the plaintiff's")
        doc.insert_tracked("so-called ", before='"free trial"')
        doc.insert_tracked("likely ", before="won't")

        text = doc.get_text()
        assert "clearly" in text
        assert "so-called" in text
        assert "likely" in text

    finally:
        doc_path.unlink()
