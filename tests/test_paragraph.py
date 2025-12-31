"""
Tests for Paragraph wrapper class.
"""

import pytest
from lxml import etree

from python_docx_redline.models.paragraph import WORD_NAMESPACE, Paragraph


def create_paragraph_element(text: str = "", style: str | None = None) -> etree._Element:
    """Helper to create a w:p element for testing."""
    p = etree.Element(f"{{{WORD_NAMESPACE}}}p")

    # Add style if specified
    if style:
        p_pr = etree.SubElement(p, f"{{{WORD_NAMESPACE}}}pPr")
        p_style = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}pStyle")
        p_style.set(f"{{{WORD_NAMESPACE}}}val", style)

    # Add text if specified
    if text:
        r = etree.SubElement(p, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(r, f"{{{WORD_NAMESPACE}}}t")
        t.text = text

    return p


def test_paragraph_init():
    """Test Paragraph initialization."""
    elem = create_paragraph_element("Test text")
    para = Paragraph(elem)
    assert para.element is elem


def test_paragraph_init_invalid_element():
    """Test Paragraph raises error for non-paragraph element."""
    elem = etree.Element(f"{{{WORD_NAMESPACE}}}r")
    with pytest.raises(ValueError, match="Expected w:p element"):
        Paragraph(elem)


def test_paragraph_text_getter():
    """Test getting paragraph text."""
    elem = create_paragraph_element("Hello World")
    para = Paragraph(elem)
    assert para.text == "Hello World"


def test_paragraph_text_multiple_runs():
    """Test getting text from paragraph with multiple runs."""
    p = etree.Element(f"{{{WORD_NAMESPACE}}}p")

    # Add multiple runs
    for text in ["Hello", " ", "World"]:
        r = etree.SubElement(p, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(r, f"{{{WORD_NAMESPACE}}}t")
        t.text = text

    para = Paragraph(p)
    assert para.text == "Hello World"


def test_paragraph_text_empty():
    """Test getting text from empty paragraph."""
    elem = create_paragraph_element()
    para = Paragraph(elem)
    assert para.text == ""


def test_paragraph_text_setter():
    """Test setting paragraph text."""
    elem = create_paragraph_element("Old text")
    para = Paragraph(elem)

    para.text = "New text"
    assert para.text == "New text"

    # Should have exactly one run now
    runs = para.runs
    assert len(runs) == 1


def test_paragraph_text_setter_replaces_multiple_runs():
    """Test that setting text replaces all existing runs."""
    p = etree.Element(f"{{{WORD_NAMESPACE}}}p")

    # Add multiple runs
    for text in ["Run1", "Run2", "Run3"]:
        r = etree.SubElement(p, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(r, f"{{{WORD_NAMESPACE}}}t")
        t.text = text

    para = Paragraph(p)
    assert len(para.runs) == 3

    para.text = "Single run"
    assert para.text == "Single run"
    assert len(para.runs) == 1


def test_paragraph_style_getter():
    """Test getting paragraph style."""
    elem = create_paragraph_element("Text", style="Heading1")
    para = Paragraph(elem)
    assert para.style == "Heading1"


def test_paragraph_style_getter_none():
    """Test getting style when none is set."""
    elem = create_paragraph_element("Text")
    para = Paragraph(elem)
    assert para.style is None


def test_paragraph_style_setter():
    """Test setting paragraph style."""
    elem = create_paragraph_element("Text")
    para = Paragraph(elem)

    para.style = "Heading2"
    assert para.style == "Heading2"


def test_paragraph_style_setter_updates_existing():
    """Test that setting style updates existing style."""
    elem = create_paragraph_element("Text", style="Normal")
    para = Paragraph(elem)
    assert para.style == "Normal"

    para.style = "Heading1"
    assert para.style == "Heading1"


def test_paragraph_style_setter_none():
    """Test removing style by setting to None."""
    elem = create_paragraph_element("Text", style="Heading1")
    para = Paragraph(elem)
    assert para.style == "Heading1"

    para.style = None
    assert para.style is None


def test_paragraph_runs_property():
    """Test getting runs from paragraph."""
    p = etree.Element(f"{{{WORD_NAMESPACE}}}p")

    # Add three runs
    for _ in range(3):
        etree.SubElement(p, f"{{{WORD_NAMESPACE}}}r")

    para = Paragraph(p)
    runs = para.runs
    assert len(runs) == 3
    assert all(r.tag == f"{{{WORD_NAMESPACE}}}r" for r in runs)


def test_paragraph_runs_empty():
    """Test runs property on paragraph with no runs."""
    elem = create_paragraph_element()
    para = Paragraph(elem)
    assert para.runs == []


def test_paragraph_is_heading_true():
    """Test is_heading returns True for heading paragraphs."""
    elem = create_paragraph_element("Title", style="Heading1")
    para = Paragraph(elem)
    assert para.is_heading() is True


def test_paragraph_is_heading_false():
    """Test is_heading returns False for non-heading paragraphs."""
    elem = create_paragraph_element("Text", style="Normal")
    para = Paragraph(elem)
    assert para.is_heading() is False


def test_paragraph_is_heading_no_style():
    """Test is_heading returns False when no style is set."""
    elem = create_paragraph_element("Text")
    para = Paragraph(elem)
    assert para.is_heading() is False


def test_paragraph_get_heading_level():
    """Test getting heading level from heading paragraphs."""
    for level in range(1, 10):
        elem = create_paragraph_element("Title", style=f"Heading{level}")
        para = Paragraph(elem)
        assert para.get_heading_level() == level


def test_paragraph_get_heading_level_non_heading():
    """Test get_heading_level returns None for non-headings."""
    elem = create_paragraph_element("Text", style="Normal")
    para = Paragraph(elem)
    assert para.get_heading_level() is None


def test_paragraph_get_heading_level_no_style():
    """Test get_heading_level returns None when no style."""
    elem = create_paragraph_element("Text")
    para = Paragraph(elem)
    assert para.get_heading_level() is None


def test_paragraph_contains_case_sensitive():
    """Test contains method with case sensitivity."""
    elem = create_paragraph_element("The Quick Brown Fox")
    para = Paragraph(elem)

    assert para.contains("Quick") is True
    assert para.contains("quick") is False
    assert para.contains("Fox") is True
    assert para.contains("fox") is False


def test_paragraph_contains_case_insensitive():
    """Test contains method without case sensitivity."""
    elem = create_paragraph_element("The Quick Brown Fox")
    para = Paragraph(elem)

    assert para.contains("Quick", case_sensitive=False) is True
    assert para.contains("quick", case_sensitive=False) is True
    assert para.contains("BROWN", case_sensitive=False) is True
    assert para.contains("FOX", case_sensitive=False) is True


def test_paragraph_contains_not_found():
    """Test contains returns False when text not found."""
    elem = create_paragraph_element("Hello World")
    para = Paragraph(elem)
    assert para.contains("Goodbye") is False


def test_paragraph_parent_section():
    """Test parent_section property."""
    elem = create_paragraph_element("Text")
    para = Paragraph(elem)

    # Initially None
    assert para.parent_section is None

    # Can be set (will be used by Section class)

    mock_section = object()
    para._set_parent_section(mock_section)  # type: ignore
    assert para.parent_section is mock_section


def test_paragraph_repr():
    """Test string representation of paragraph."""
    elem = create_paragraph_element("Hello World", style="Heading1")
    para = Paragraph(elem)
    repr_str = repr(para)

    assert "Paragraph" in repr_str
    assert "Hello World" in repr_str
    assert "Heading1" in repr_str


def test_paragraph_repr_long_text():
    """Test repr truncates long text."""
    long_text = "A" * 100
    elem = create_paragraph_element(long_text)
    para = Paragraph(elem)
    repr_str = repr(para)

    assert "Paragraph" in repr_str
    assert "..." in repr_str
    assert len(repr_str) < len(long_text) + 50  # Should be truncated


def test_paragraph_repr_no_style():
    """Test repr when no style is set."""
    elem = create_paragraph_element("Text")
    para = Paragraph(elem)
    repr_str = repr(para)

    assert "Paragraph" in repr_str
    assert "Text" in repr_str
    assert "style=" not in repr_str


# === Tests for Issue #6: Paragraph.text setter preserving pPr and markdown support ===


def test_paragraph_text_setter_preserves_ppr():
    """Test that setting text preserves paragraph properties (w:pPr)."""
    # Create paragraph with style and other pPr properties
    p = etree.Element(f"{{{WORD_NAMESPACE}}}p")
    p_pr = etree.SubElement(p, f"{{{WORD_NAMESPACE}}}pPr")
    p_style = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}pStyle")
    p_style.set(f"{{{WORD_NAMESPACE}}}val", "Heading1")
    # Add justification property
    jc = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}jc")
    jc.set(f"{{{WORD_NAMESPACE}}}val", "center")
    # Add original text
    r = etree.SubElement(p, f"{{{WORD_NAMESPACE}}}r")
    t = etree.SubElement(r, f"{{{WORD_NAMESPACE}}}t")
    t.text = "Original text"

    para = Paragraph(p)

    # Verify initial state
    assert para.style == "Heading1"

    # Set new text
    para.text = "New text"

    # Verify text was updated
    assert para.text == "New text"

    # Verify pPr was preserved
    assert para.style == "Heading1"

    # Verify justification was also preserved
    p_pr = para.element.find(f"{{{WORD_NAMESPACE}}}pPr")
    assert p_pr is not None
    jc = p_pr.find(f"{{{WORD_NAMESPACE}}}jc")
    assert jc is not None
    assert jc.get(f"{{{WORD_NAMESPACE}}}val") == "center"


def test_paragraph_text_setter_markdown_bold():
    """Test that setting text with **bold** creates bold formatting."""
    elem = create_paragraph_element("Old text")
    para = Paragraph(elem)

    para.text = "This is **bold** text"

    # Should have multiple runs for different formatting
    runs = para.runs
    assert len(runs) == 3  # "This is ", "bold", " text"

    # Second run should have bold formatting
    bold_run = runs[1]
    rpr = bold_run.find(f"{{{WORD_NAMESPACE}}}rPr")
    assert rpr is not None
    b_elem = rpr.find(f"{{{WORD_NAMESPACE}}}b")
    assert b_elem is not None

    # Text should read correctly
    assert para.text == "This is bold text"


def test_paragraph_text_setter_markdown_italic():
    """Test that setting text with *italic* creates italic formatting."""
    elem = create_paragraph_element("Old text")
    para = Paragraph(elem)

    para.text = "This is *italic* text"

    # Should have multiple runs
    runs = para.runs
    assert len(runs) == 3  # "This is ", "italic", " text"

    # Second run should have italic formatting
    italic_run = runs[1]
    rpr = italic_run.find(f"{{{WORD_NAMESPACE}}}rPr")
    assert rpr is not None
    i_elem = rpr.find(f"{{{WORD_NAMESPACE}}}i")
    assert i_elem is not None

    # Text should read correctly
    assert para.text == "This is italic text"


def test_paragraph_text_setter_markdown_mixed():
    """Test that setting text with mixed markdown works."""
    elem = create_paragraph_element("Old text")
    para = Paragraph(elem)

    para.text = "**bold** and *italic*"

    # Text should read correctly
    assert para.text == "bold and italic"

    # Should have multiple runs
    runs = para.runs
    assert len(runs) >= 3  # "bold", " and ", "italic"


def test_paragraph_text_setter_empty_string():
    """Test that setting empty string doesn't break the paragraph."""
    elem = create_paragraph_element("Some text", style="Normal")
    para = Paragraph(elem)

    para.text = ""

    # Text should be empty
    assert para.text == ""

    # Style should be preserved
    assert para.style == "Normal"

    # Should still have at least one run
    runs = para.runs
    assert len(runs) >= 1


def test_paragraph_text_setter_removes_hyperlinks():
    """Test that setting text removes hyperlinks but preserves pPr."""
    p = etree.Element(f"{{{WORD_NAMESPACE}}}p")
    p_pr = etree.SubElement(p, f"{{{WORD_NAMESPACE}}}pPr")
    p_style = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}pStyle")
    p_style.set(f"{{{WORD_NAMESPACE}}}val", "Normal")

    # Add a hyperlink element (simplified)
    hyperlink = etree.SubElement(p, f"{{{WORD_NAMESPACE}}}hyperlink")
    r = etree.SubElement(hyperlink, f"{{{WORD_NAMESPACE}}}r")
    t = etree.SubElement(r, f"{{{WORD_NAMESPACE}}}t")
    t.text = "Link text"

    para = Paragraph(p)

    # Set new text
    para.text = "New text without link"

    # Hyperlink should be removed
    hyperlinks = para.element.findall(f"{{{WORD_NAMESPACE}}}hyperlink")
    assert len(hyperlinks) == 0

    # Style should be preserved
    assert para.style == "Normal"


# Run tests with: pytest tests/test_paragraph.py -v
