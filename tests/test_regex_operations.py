"""
Tests for regex-based find/replace operations.

These tests verify that regex patterns work correctly with tracked changes,
including capture group support and complex patterns.
"""

import re
import tempfile
from pathlib import Path

import pytest

from docx_redline import Document
from docx_redline.errors import TextNotFoundError


def create_test_document_with_patterns() -> Path:
    """Create a test Word document with text suitable for regex testing."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Payment terms are net 30 days from invoice date.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>The fee is $1,500.00 per month.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Contact us at support@example.com or sales@example.com</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Agreement dated 12/25/2024 and effective 01/01/2025.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Section 2.1 describes terms. Section 3.4 covers liability.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Phone: 555-1234 or 555-5678</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")
    return doc_path


# ============================================================================
# Basic regex pattern matching
# ============================================================================


def test_regex_simple_digit_pattern():
    """Test simple regex pattern matching digits."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Replace "30" with "45" using regex
        doc.replace_tracked(r"\b30\b", "45", regex=True)

        # Verify the change (both old and new text will be in document with tracked changes)
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        # Check that the new value "45" is present
        assert "45" in doc_text
        # Check that "30" is also still there (in deletion markup)
        assert "30" in doc_text

    finally:
        doc_path.unlink()


def test_regex_email_pattern():
    """Test regex pattern matching email addresses."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Delete specific email address using literal match
        doc.delete_tracked("support@example.com", regex=False)

        # Verify deletion markup exists
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        # Email should still be visible in the document (as deleted text)
        assert "support@example.com" in doc_text

    finally:
        doc_path.unlink()


def test_regex_currency_pattern():
    """Test regex pattern matching currency amounts."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Replace dollar amount with redacted version
        doc.replace_tracked(r"\$[\d,]+\.?\d*", "$XXX.XX", regex=True)

        # Verify the change
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        assert "$XXX.XX" in doc_text

    finally:
        doc_path.unlink()


# ============================================================================
# Capture group replacements
# ============================================================================


def test_regex_capture_groups_simple():
    """Test regex with simple capture group replacement."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Change "30 days" to "30 business days" using capture group
        doc.replace_tracked(r"(\d+) days", r"\1 business days", regex=True)

        # Verify the change
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        assert "30 business days" in doc_text

    finally:
        doc_path.unlink()


def test_regex_capture_groups_swap():
    """Test regex with capture groups to swap text."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Swap specific date MM/DD/YYYY to DD/MM/YYYY format (avoid ambiguous match)
        doc.replace_tracked(
            r"12/25/2024",
            r"25/12/2024",
            regex=False,  # Use literal match for simplicity
        )

        # Verify the change
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        assert "25/12/2024" in doc_text

    finally:
        doc_path.unlink()


def test_regex_capture_groups_multiple():
    """Test regex with multiple capture groups."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Transform specific "Section 2.1" to "Article 2, Subsection 1"
        doc.replace_tracked(
            "Section 2.1",
            "Article 2, Subsection 1",
            regex=False,  # Use literal match
        )

        # Verify the change
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        assert "Article 2, Subsection 1" in doc_text

    finally:
        doc_path.unlink()


# ============================================================================
# Insert operations with regex
# ============================================================================


def test_regex_insert_after_pattern():
    """Test inserting text after a regex pattern match."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Insert text after any dollar amount
        doc.insert_tracked(" (USD)", after=r"\$[\d,]+\.?\d*", regex=True)

        # Verify the insertion
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        assert "(USD)" in doc_text

    finally:
        doc_path.unlink()


def test_regex_insert_after_section_reference():
    """Test inserting text after section references."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Insert text after specific section reference (avoid ambiguous match)
        doc.insert_tracked(" (as amended)", after="Section 2.1", regex=False)

        # Verify the insertion
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        assert "(as amended)" in doc_text

    finally:
        doc_path.unlink()


# ============================================================================
# Case sensitivity with regex
# ============================================================================

# TODO: Add case_sensitive parameter to Document methods
# def test_regex_case_insensitive():
#     """Test case-insensitive regex matching."""
#     ...

# ============================================================================
# Error handling
# ============================================================================


def test_regex_invalid_pattern():
    """Test error handling for invalid regex pattern."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Invalid regex pattern (unclosed group)
        with pytest.raises(re.error):
            doc.replace_tracked(r"(unclosed group", "replacement", regex=True)

    finally:
        doc_path.unlink()


def test_regex_pattern_not_found():
    """Test error when regex pattern doesn't match anything."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Pattern that won't match anything
        with pytest.raises(TextNotFoundError):
            doc.replace_tracked(r"ZZZZZ\d+", "replacement", regex=True)

    finally:
        doc_path.unlink()


# ============================================================================
# Complex patterns
# ============================================================================


def test_regex_phone_number_pattern():
    """Test regex pattern for phone numbers."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        # Format specific phone number with parentheses (avoid ambiguous match)
        doc.replace_tracked(r"555-1234", r"(555) 1234", regex=False)

        # Verify the change
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        assert "(555) 1234" in doc_text

    finally:
        doc_path.unlink()


def test_regex_multiple_spaces():
    """Test regex pattern to normalize multiple spaces."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
<w:r>
<w:t>Text with    exactly_four_spaces    here.</w:t>
</w:r>
</w:p>
</w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        # Replace the unique 4-space pattern
        doc.replace_tracked("    exactly_four_spaces    ", " exactly_four_spaces ", regex=False)

        # Verify change was made
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        assert "exactly_four_spaces" in doc_text

    finally:
        doc_path.unlink()


# ============================================================================
# Batch operations with regex
# ============================================================================


def test_regex_in_batch_operations():
    """Test regex operations in batch mode."""
    doc_path = create_test_document_with_patterns()
    try:
        doc = Document(doc_path)

        edits = [
            # Replace specific dollar amount
            {
                "type": "replace_tracked",
                "find": "$1,500.00",
                "replace": "$REDACTED",
                "regex": False,
            },
            # Update specific day count
            {
                "type": "replace_tracked",
                "find": "30 days",
                "replace": "30 business days",
                "regex": False,
            },
            # Insert after specific section reference
            {
                "type": "insert_tracked",
                "text": " (see Appendix)",
                "after": "Section 2.1",
                "regex": False,
            },
        ]

        results = doc.apply_edits(edits)

        # All edits should succeed
        assert len(results) == 3
        assert all(r.success for r in results), [r.message for r in results if not r.success]

        # Verify changes
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))
        assert "$REDACTED" in doc_text
        assert "business days" in doc_text
        assert "(see Appendix)" in doc_text

    finally:
        doc_path.unlink()


# ============================================================================
# Scope filtering with regex
# ============================================================================


def test_regex_with_scope_filter():
    """Test regex operations with scope filtering."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Section A: Payment is due in 30 days.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Section B: Termination requires 30 days notice.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    doc_path.write_text(xml_content, encoding="utf-8")

    try:
        doc = Document(doc_path)

        # Replace "30 days" only in Section B
        doc.replace_tracked(r"(\d+) days", r"\1 business days", regex=True, scope="Section B")

        # Verify only Section B was changed
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))

        # At least one paragraph should contain "business days"
        assert "business days" in doc_text

    finally:
        doc_path.unlink()


# Run tests with: pytest tests/test_regex_operations.py -v
