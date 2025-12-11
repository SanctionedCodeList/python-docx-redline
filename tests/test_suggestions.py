"""
Tests for the SuggestionGenerator class.

These tests verify that helpful suggestions are generated for common
text search issues.
"""

from lxml import etree

from python_docx_redline import SuggestionGenerator

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def create_test_paragraph(text: str) -> etree.Element:
    """Create a test paragraph element with given text."""
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{WORD_NAMESPACE}">
  <w:body>
    <w:p>
      <w:r>
        <w:t>{text}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""
    root = etree.fromstring(xml.encode("utf-8"))
    return list(root.iter(f"{{{WORD_NAMESPACE}}}p"))[0]


def test_suggestion_curly_quotes():
    """Test suggestion for curly quotes."""
    # Document has curly quotes
    para = create_test_paragraph("This is a \u201cquoted\u201d text")

    # User searches with straight quotes
    suggestions = SuggestionGenerator.generate_suggestions('a "quoted" text', [para])

    assert any("curly quotes" in s.lower() for s in suggestions)


def test_suggestion_double_spaces():
    """Test suggestion for double spaces in search."""
    para = create_test_paragraph("This is a test")

    # User searches with double space
    suggestions = SuggestionGenerator.generate_suggestions("is  a", [para])

    assert any("double spaces" in s for s in suggestions)


def test_suggestion_leading_trailing_whitespace():
    """Test suggestion for whitespace in search."""
    para = create_test_paragraph("This is a test")

    # User searches with leading/trailing space
    suggestions = SuggestionGenerator.generate_suggestions(" is a ", [para])

    assert any("whitespace" in s.lower() for s in suggestions)


def test_suggestion_case_sensitivity():
    """Test suggestion for case mismatch."""
    para = create_test_paragraph("This is a Test")

    # User searches with wrong case
    suggestions = SuggestionGenerator.generate_suggestions("this is a test", [para])

    assert any("case" in s.lower() for s in suggestions)


def test_suggestion_not_found():
    """Test generic suggestions when text not found."""
    para = create_test_paragraph("This is a test")

    # User searches for text that doesn't exist
    suggestions = SuggestionGenerator.generate_suggestions("nonexistent", [para])

    assert any("typos" in s.lower() for s in suggestions)
    assert len(suggestions) >= 3  # Should have multiple fallback suggestions


def test_check_common_issues_curly_quotes():
    """Test detection of curly quotes in text."""
    issues = SuggestionGenerator.check_common_issues("Text with \u201ccurly\u201d quotes")

    assert any("curly" in issue.lower() for issue in issues)


def test_check_common_issues_non_breaking_space():
    """Test detection of non-breaking spaces."""
    issues = SuggestionGenerator.check_common_issues("Text with\u00a0non-breaking space")

    assert any("non-breaking" in issue.lower() for issue in issues)


def test_check_common_issues_zero_width_space():
    """Test detection of zero-width spaces."""
    issues = SuggestionGenerator.check_common_issues("Text with\u200bzero-width space")

    assert any("zero-width" in issue.lower() for issue in issues)


def test_check_common_issues_tabs():
    """Test detection of tab characters."""
    issues = SuggestionGenerator.check_common_issues("Text with\ttab")

    assert any("tab" in issue.lower() for issue in issues)


def test_check_common_issues_line_breaks():
    """Test detection of line breaks."""
    issues = SuggestionGenerator.check_common_issues("Text with\nline break")

    assert any("line break" in issue.lower() for issue in issues)


# Run tests with: pytest tests/test_suggestions.py -v
