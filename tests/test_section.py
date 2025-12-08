"""
Tests for Section wrapper class.
"""

from lxml import etree

from docx_redline.models.paragraph import WORD_NAMESPACE, Paragraph
from docx_redline.models.section import Section


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


def create_document_root(para_specs: list[tuple[str, str | None]]) -> etree._Element:
    """Create a document root with specified paragraphs.

    Args:
        para_specs: List of (text, style) tuples

    Returns:
        Document root element
    """
    root = etree.Element(f"{{{WORD_NAMESPACE}}}document")
    body = etree.SubElement(root, f"{{{WORD_NAMESPACE}}}body")

    for text, style in para_specs:
        p = create_paragraph_element(text, style)
        body.append(p)

    return root


def test_section_init():
    """Test Section initialization."""
    heading = Paragraph(create_paragraph_element("Heading", "Heading1"))
    para1 = Paragraph(create_paragraph_element("Text 1"))
    para2 = Paragraph(create_paragraph_element("Text 2"))

    section = Section(heading, [heading, para1, para2])

    assert section.heading == heading
    assert len(section.paragraphs) == 3
    assert section.paragraphs[0] == heading
    assert section.paragraphs[1] == para1
    assert section.paragraphs[2] == para2


def test_section_init_no_heading():
    """Test Section with no heading (intro section)."""
    para1 = Paragraph(create_paragraph_element("Text 1"))
    para2 = Paragraph(create_paragraph_element("Text 2"))

    section = Section(None, [para1, para2])

    assert section.heading is None
    assert len(section.paragraphs) == 2


def test_section_sets_parent_on_paragraphs():
    """Test that Section sets itself as parent on paragraphs."""
    heading = Paragraph(create_paragraph_element("Heading", "Heading1"))
    para1 = Paragraph(create_paragraph_element("Text 1"))

    section = Section(heading, [heading, para1])

    assert heading.parent_section == section
    assert para1.parent_section == section


def test_section_heading_property():
    """Test heading property."""
    heading = Paragraph(create_paragraph_element("Section Title", "Heading1"))
    para1 = Paragraph(create_paragraph_element("Content"))

    section = Section(heading, [heading, para1])

    assert section.heading == heading
    assert section.heading is not None
    assert section.heading.text == "Section Title"


def test_section_paragraphs_property():
    """Test paragraphs property."""
    heading = Paragraph(create_paragraph_element("Heading", "Heading1"))
    para1 = Paragraph(create_paragraph_element("Text 1"))
    para2 = Paragraph(create_paragraph_element("Text 2"))
    para3 = Paragraph(create_paragraph_element("Text 3"))

    section = Section(heading, [heading, para1, para2, para3])

    assert len(section.paragraphs) == 4
    assert section.paragraphs == [heading, para1, para2, para3]


def test_section_heading_text():
    """Test heading_text property."""
    heading = Paragraph(create_paragraph_element("Introduction", "Heading1"))
    para1 = Paragraph(create_paragraph_element("Content"))

    section = Section(heading, [heading, para1])

    assert section.heading_text == "Introduction"


def test_section_heading_text_none():
    """Test heading_text when no heading."""
    para1 = Paragraph(create_paragraph_element("Content"))
    section = Section(None, [para1])

    assert section.heading_text is None


def test_section_heading_level():
    """Test heading_level property."""
    for level in [1, 2, 3]:
        heading = Paragraph(create_paragraph_element(f"Heading {level}", f"Heading{level}"))
        para1 = Paragraph(create_paragraph_element("Content"))

        section = Section(heading, [heading, para1])

        assert section.heading_level == level


def test_section_heading_level_none():
    """Test heading_level when no heading."""
    para1 = Paragraph(create_paragraph_element("Content"))
    section = Section(None, [para1])

    assert section.heading_level is None


def test_section_contains_true():
    """Test contains method when text is found."""
    heading = Paragraph(create_paragraph_element("Introduction", "Heading1"))
    para1 = Paragraph(create_paragraph_element("The quick brown fox"))
    para2 = Paragraph(create_paragraph_element("jumps over the lazy dog"))

    section = Section(heading, [heading, para1, para2])

    assert section.contains("Introduction") is True
    assert section.contains("quick brown") is True
    assert section.contains("lazy dog") is True


def test_section_contains_false():
    """Test contains method when text is not found."""
    heading = Paragraph(create_paragraph_element("Introduction", "Heading1"))
    para1 = Paragraph(create_paragraph_element("Content"))

    section = Section(heading, [heading, para1])

    assert section.contains("nonexistent") is False


def test_section_contains_case_sensitive():
    """Test contains with case sensitivity."""
    para1 = Paragraph(create_paragraph_element("The Quick Brown Fox"))
    section = Section(None, [para1])

    assert section.contains("Quick") is True
    assert section.contains("quick") is False


def test_section_contains_case_insensitive():
    """Test contains without case sensitivity."""
    para1 = Paragraph(create_paragraph_element("The Quick Brown Fox"))
    section = Section(None, [para1])

    assert section.contains("Quick", case_sensitive=False) is True
    assert section.contains("quick", case_sensitive=False) is True
    assert section.contains("BROWN", case_sensitive=False) is True


def test_section_find_paragraph():
    """Test find_paragraph method."""
    heading = Paragraph(create_paragraph_element("Section", "Heading1"))
    para1 = Paragraph(create_paragraph_element("First paragraph"))
    para2 = Paragraph(create_paragraph_element("Second paragraph with target"))
    para3 = Paragraph(create_paragraph_element("Third paragraph"))

    section = Section(heading, [heading, para1, para2, para3])

    found = section.find_paragraph("target")
    assert found == para2
    assert found is not None
    assert found.text == "Second paragraph with target"


def test_section_find_paragraph_not_found():
    """Test find_paragraph when text not found."""
    para1 = Paragraph(create_paragraph_element("Content"))
    section = Section(None, [para1])

    assert section.find_paragraph("nonexistent") is None


def test_section_find_paragraph_returns_first():
    """Test find_paragraph returns first match."""
    para1 = Paragraph(create_paragraph_element("Text with target"))
    para2 = Paragraph(create_paragraph_element("Another with target"))
    section = Section(None, [para1, para2])

    found = section.find_paragraph("target")
    assert found == para1  # Should be the first one


def test_section_from_document_no_headings():
    """Test from_document with no headings (single intro section)."""
    root = create_document_root(
        [
            ("Paragraph 1", None),
            ("Paragraph 2", None),
            ("Paragraph 3", None),
        ]
    )

    sections = Section.from_document(root)

    assert len(sections) == 1
    assert sections[0].heading is None
    assert len(sections[0].paragraphs) == 3


def test_section_from_document_with_headings():
    """Test from_document with headings."""
    root = create_document_root(
        [
            ("Introduction", "Heading1"),
            ("Intro paragraph 1", None),
            ("Intro paragraph 2", None),
            ("Methods", "Heading1"),
            ("Methods paragraph 1", None),
            ("Results", "Heading1"),
            ("Results paragraph 1", None),
            ("Results paragraph 2", None),
        ]
    )

    sections = Section.from_document(root)

    assert len(sections) == 3

    # Section 1: Introduction
    assert sections[0].heading_text == "Introduction"
    assert len(sections[0].paragraphs) == 3  # heading + 2 paras

    # Section 2: Methods
    assert sections[1].heading_text == "Methods"
    assert len(sections[1].paragraphs) == 2  # heading + 1 para

    # Section 3: Results
    assert sections[2].heading_text == "Results"
    assert len(sections[2].paragraphs) == 3  # heading + 2 paras


def test_section_from_document_intro_then_sections():
    """Test from_document with intro paragraphs before first heading."""
    root = create_document_root(
        [
            ("Abstract paragraph", None),
            ("Another abstract paragraph", None),
            ("Introduction", "Heading1"),
            ("Intro content", None),
            ("Methods", "Heading1"),
            ("Methods content", None),
        ]
    )

    sections = Section.from_document(root)

    assert len(sections) == 3

    # Intro section (no heading)
    assert sections[0].heading is None
    assert len(sections[0].paragraphs) == 2

    # Introduction section
    assert sections[1].heading_text == "Introduction"
    assert len(sections[1].paragraphs) == 2

    # Methods section
    assert sections[2].heading_text == "Methods"
    assert len(sections[2].paragraphs) == 2


def test_section_from_document_different_heading_levels():
    """Test from_document with different heading levels."""
    root = create_document_root(
        [
            ("Chapter 1", "Heading1"),
            ("Section 1.1", "Heading2"),
            ("Content 1.1", None),
            ("Section 1.2", "Heading2"),
            ("Content 1.2", None),
            ("Chapter 2", "Heading1"),
            ("Content 2", None),
        ]
    )

    sections = Section.from_document(root)

    # All headings create new sections
    # Chapter 1 has just heading (no content before next heading)
    # Section 1.1 has heading + content
    # Section 1.2 has heading + content
    # Chapter 2 has heading + content
    assert len(sections) == 4

    assert sections[0].heading_text == "Chapter 1"
    assert sections[0].heading_level == 1
    assert len(sections[0].paragraphs) == 1  # Just the heading

    assert sections[1].heading_text == "Section 1.1"
    assert sections[1].heading_level == 2
    assert len(sections[1].paragraphs) == 2  # Heading + content

    assert sections[2].heading_text == "Section 1.2"
    assert sections[2].heading_level == 2
    assert len(sections[2].paragraphs) == 2  # Heading + content

    assert sections[3].heading_text == "Chapter 2"
    assert sections[3].heading_level == 1
    assert len(sections[3].paragraphs) == 2  # Heading + content


def test_section_repr_with_heading():
    """Test string representation with heading."""
    heading = Paragraph(create_paragraph_element("Introduction", "Heading1"))
    para1 = Paragraph(create_paragraph_element("Content"))

    section = Section(heading, [heading, para1])

    repr_str = repr(section)
    assert "Section" in repr_str
    assert "Introduction" in repr_str
    assert "paragraphs=2" in repr_str


def test_section_repr_intro():
    """Test string representation for intro section."""
    para1 = Paragraph(create_paragraph_element("Content"))
    para2 = Paragraph(create_paragraph_element("More content"))

    section = Section(None, [para1, para2])

    repr_str = repr(section)
    assert "Section" in repr_str
    assert "intro" in repr_str
    assert "paragraphs=2" in repr_str


def test_section_len():
    """Test __len__ returns paragraph count."""
    heading = Paragraph(create_paragraph_element("Heading", "Heading1"))
    para1 = Paragraph(create_paragraph_element("Text 1"))
    para2 = Paragraph(create_paragraph_element("Text 2"))

    section = Section(heading, [heading, para1, para2])

    assert len(section) == 3


def test_section_empty():
    """Test section with no paragraphs."""
    section = Section(None, [])

    assert len(section) == 0
    assert section.heading is None
    assert section.paragraphs == []


# Run tests with: pytest tests/test_section.py -v
