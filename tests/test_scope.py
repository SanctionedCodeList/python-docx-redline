"""
Tests for the ScopeEvaluator class.

These tests verify that scope specifications correctly filter paragraphs
in Word documents.
"""

from lxml import etree

from python_docx_redline import ScopeEvaluator

WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def create_test_paragraph(text: str, style: str | None = None) -> etree.Element:
    """Create a test paragraph element with given text and optional style."""
    p_pr_xml = ""
    if style:
        p_pr_xml = f"""
      <w:pPr>
        <w:pStyle w:val="{style}"/>
      </w:pPr>"""

    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{WORD_NAMESPACE}">
  <w:body>
    <w:p>{p_pr_xml}
      <w:r>
        <w:t>{text}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""
    root = etree.fromstring(xml.encode("utf-8"))
    return list(root.iter(f"{{{WORD_NAMESPACE}}}p"))[0]


def create_test_document(paragraphs_data: list[tuple[str, str | None]]) -> list[etree.Element]:
    """Create a test document with multiple paragraphs.

    Args:
        paragraphs_data: List of (text, style) tuples

    Returns:
        List of paragraph Elements
    """
    p_elements = []
    for text, style in paragraphs_data:
        p_pr_xml = ""
        if style:
            p_pr_xml = f"""
        <w:pPr>
          <w:pStyle w:val="{style}"/>
        </w:pPr>"""

        p_elements.append(
            f"""<w:p>{p_pr_xml}
        <w:r>
          <w:t>{text}</w:t>
        </w:r>
      </w:p>"""
        )

    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{WORD_NAMESPACE}">
  <w:body>
    {"".join(p_elements)}
  </w:body>
</w:document>"""
    root = etree.fromstring(xml.encode("utf-8"))
    return list(root.iter(f"{{{WORD_NAMESPACE}}}p"))


def test_scope_none_matches_all():
    """Test that None scope matches all paragraphs."""
    para1 = create_test_paragraph("First paragraph")
    para2 = create_test_paragraph("Second paragraph")
    paragraphs = [para1, para2]

    evaluator = ScopeEvaluator.parse(None)
    filtered = [p for p in paragraphs if evaluator(p)]

    assert len(filtered) == 2
    assert filtered == paragraphs


def test_scope_string_paragraph_containing():
    """Test string scope matches paragraphs containing text."""
    para1 = create_test_paragraph("This is a test")
    para2 = create_test_paragraph("Another paragraph")
    para3 = create_test_paragraph("Third one")
    paragraphs = [para1, para2, para3]

    # Test default (paragraph containing)
    filtered = ScopeEvaluator.filter_paragraphs(paragraphs, "test")
    assert len(filtered) == 1
    assert "test" in "".join(filtered[0].itertext())


def test_scope_string_explicit_paragraph_containing():
    """Test explicit paragraph_containing: prefix."""
    para1 = create_test_paragraph("This contains target text")
    para2 = create_test_paragraph("This does not")
    paragraphs = [para1, para2]

    filtered = ScopeEvaluator.filter_paragraphs(paragraphs, "paragraph_containing:target")
    assert len(filtered) == 1
    assert "target" in "".join(filtered[0].itertext())


def test_scope_section():
    """Test section: scope filters by section heading."""
    paragraphs = create_test_document(
        [
            ("Introduction", "Heading1"),
            ("This is the intro", None),
            ("More intro text", None),
            ("Methodology", "Heading1"),
            ("This is the method", None),
            ("Results", "Heading1"),
            ("These are results", None),
        ]
    )

    # Filter for paragraphs in the Methodology section
    filtered = ScopeEvaluator.filter_paragraphs(paragraphs, "section:Methodology")

    # Should include only paragraphs after "Methodology" heading
    # and before the next heading
    assert len(filtered) == 1
    assert "method" in "".join(filtered[0].itertext())


def test_scope_dict_contains():
    """Test dictionary scope with 'contains' filter."""
    para1 = create_test_paragraph("This has keyword")
    para2 = create_test_paragraph("This does not")
    para3 = create_test_paragraph("keyword here too")
    paragraphs = [para1, para2, para3]

    filtered = ScopeEvaluator.filter_paragraphs(paragraphs, {"contains": "keyword"})
    assert len(filtered) == 2


def test_scope_dict_not_contains():
    """Test dictionary scope with 'not_contains' filter."""
    para1 = create_test_paragraph("This has exclude")
    para2 = create_test_paragraph("This is fine")
    para3 = create_test_paragraph("Also has exclude")
    paragraphs = [para1, para2, para3]

    filtered = ScopeEvaluator.filter_paragraphs(paragraphs, {"not_contains": "exclude"})
    assert len(filtered) == 1
    assert "fine" in "".join(filtered[0].itertext())


def test_scope_dict_combined_filters():
    """Test dictionary scope with multiple filters."""
    para1 = create_test_paragraph("Has include and fine")
    para2 = create_test_paragraph("Has include and exclude")
    para3 = create_test_paragraph("Just random text")
    paragraphs = [para1, para2, para3]

    filtered = ScopeEvaluator.filter_paragraphs(
        paragraphs, {"contains": "include", "not_contains": "exclude"}
    )
    assert len(filtered) == 1
    assert "fine" in "".join(filtered[0].itertext())


def test_scope_callable():
    """Test callable scope."""
    para1 = create_test_paragraph("Apple")
    para2 = create_test_paragraph("Banana")
    para3 = create_test_paragraph("Cherry")
    paragraphs = [para1, para2, para3]

    # Custom filter: paragraphs containing the letter 'a'
    def contains_a(p):
        text = "".join(p.itertext())
        return "a" in text.lower()

    filtered = ScopeEvaluator.filter_paragraphs(paragraphs, contains_a)
    # Should match Apple and Banana (both contain 'a')
    assert len(filtered) == 2
    texts = ["".join(p.itertext()) for p in filtered]
    assert "Apple" in texts[0] or "Banana" in texts[0]
    assert "Apple" in texts[1] or "Banana" in texts[1]


def test_scope_invalid_raises_error():
    """Test that invalid scope raises ValueError."""
    try:
        ScopeEvaluator.parse(123)  # Invalid type
        assert False, "Should have raised ValueError"
    except ValueError as e:
        assert "Invalid scope" in str(e)


def test_is_heading():
    """Test heading detection."""
    heading = create_test_paragraph("Chapter 1", "Heading1")
    normal = create_test_paragraph("Regular text")

    assert ScopeEvaluator._is_heading(heading)
    assert not ScopeEvaluator._is_heading(normal)


def test_section_filter_no_heading():
    """Test section filter when no heading exists."""
    para1 = create_test_paragraph("Just text")
    para2 = create_test_paragraph("More text")
    paragraphs = [para1, para2]

    # Should not match any paragraphs if no heading exists
    filtered = ScopeEvaluator.filter_paragraphs(paragraphs, "section:NonExistent")
    assert len(filtered) == 0


def test_section_filter_multiple_sections():
    """Test section filter with multiple sections."""
    paragraphs = create_test_document(
        [
            ("Section A", "Heading1"),
            ("Content A1", None),
            ("Content A2", None),
            ("Section B", "Heading1"),
            ("Content B1", None),
            ("Section C", "Heading1"),
            ("Content C1", None),
        ]
    )

    # Filter for Section B
    filtered = ScopeEvaluator.filter_paragraphs(paragraphs, "section:Section B")
    assert len(filtered) == 1
    assert "B1" in "".join(filtered[0].itertext())


# Run tests with: pytest tests/test_scope.py -v
