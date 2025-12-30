"""
Tests for view capabilities (Phase 3).

Tests the paragraphs property, sections property, and get_text() method
that allow agents to read and understand document structure.
"""

import tempfile
import zipfile
from pathlib import Path

from python_docx_redline import Document, Paragraph, Section


def create_document_with_sections() -> Path:
    """Create a test document with multiple sections."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Introduction paragraph before any headings.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Section 1</w:t></w:r></w:p>
<w:p><w:r><w:t>This is the content of section 1.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Section 1.1</w:t></w:r></w:p>
<w:p><w:r><w:t>Content of section 1.1.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Section 2</w:t></w:r></w:p>
<w:p><w:r><w:t>Content of section 2.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


def test_paragraphs_property() -> None:
    """Test that doc.paragraphs returns all paragraphs."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Get all paragraphs
        paragraphs = doc.paragraphs

        # Should have 7 paragraphs total
        assert len(paragraphs) == 7

        # All should be Paragraph objects
        assert all(isinstance(p, Paragraph) for p in paragraphs)

        # Check some text content
        assert "Introduction" in paragraphs[0].text
        assert "Section 1" == paragraphs[1].text
        assert "This is the content of section 1." == paragraphs[2].text

    finally:
        doc_path.unlink()


def test_paragraphs_text_access() -> None:
    """Test accessing paragraph text through property."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Get all paragraph texts
        texts = [p.text for p in doc.paragraphs]

        # Should have 7 texts
        assert len(texts) == 7

        # Check specific texts
        assert texts[0] == "Introduction paragraph before any headings."
        assert texts[1] == "Section 1"
        assert texts[2] == "This is the content of section 1."

    finally:
        doc_path.unlink()


def test_paragraphs_style_access() -> None:
    """Test accessing paragraph styles."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # First paragraph has no style
        assert doc.paragraphs[0].style is None

        # Second paragraph is Heading1
        assert doc.paragraphs[1].style == "Heading1"

        # Fourth paragraph is Heading2
        assert doc.paragraphs[3].style == "Heading2"

    finally:
        doc_path.unlink()


def test_sections_property() -> None:
    """Test that doc.sections returns section structure."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Get all sections
        sections = doc.sections

        # Should have 4 sections:
        # 1. Intro (no heading)
        # 2. Section 1 (Heading1)
        # 3. Section 1.1 (Heading2)
        # 4. Section 2 (Heading1)
        assert len(sections) == 4

        # All should be Section objects
        assert all(isinstance(s, Section) for s in sections)

    finally:
        doc_path.unlink()


def test_sections_intro_section() -> None:
    """Test intro section (paragraphs before first heading)."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)
        sections = doc.sections

        # First section is intro (no heading)
        intro = sections[0]
        assert intro.heading is None
        assert intro.heading_text is None
        assert len(intro.paragraphs) == 1
        assert "Introduction" in intro.paragraphs[0].text

    finally:
        doc_path.unlink()


def test_sections_with_headings() -> None:
    """Test sections with headings."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)
        sections = doc.sections

        # Second section has Heading1
        section1 = sections[1]
        assert section1.heading is not None
        assert section1.heading_text == "Section 1"
        assert section1.heading_level == 1
        assert len(section1.paragraphs) == 2  # heading + content

        # Third section has Heading2
        section11 = sections[2]
        assert section11.heading_text == "Section 1.1"
        assert section11.heading_level == 2

        # Fourth section has Heading1
        section2 = sections[3]
        assert section2.heading_text == "Section 2"
        assert section2.heading_level == 1

    finally:
        doc_path.unlink()


def test_get_text_basic() -> None:
    """Test get_text() returns all text."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Get full text
        text = doc.get_text()

        # Should contain all paragraph text
        assert "Introduction paragraph before any headings." in text
        assert "Section 1" in text
        assert "This is the content of section 1." in text
        assert "Section 1.1" in text
        assert "Content of section 1.1." in text
        assert "Section 2" in text
        assert "Content of section 2." in text

    finally:
        doc_path.unlink()


def test_get_text_paragraph_separation() -> None:
    """Test that get_text() separates paragraphs with double newlines."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        text = doc.get_text()

        # Paragraphs should be separated by \n\n
        assert "\n\n" in text

        # Should have 6 separators (7 paragraphs = 6 separators)
        assert text.count("\n\n") == 6

    finally:
        doc_path.unlink()


def test_get_text_for_search() -> None:
    """Test using get_text() to search document content."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        text = doc.get_text()

        # Case-insensitive search
        assert "section 1" in text.lower()
        assert "introduction" in text.lower()

        # Check for specific phrases
        assert "Content of section 1." in text

        # Can use for conditional logic
        has_section2 = "Section 2" in text
        assert has_section2 is True

        # Check for unique text
        assert "This is the content of section 1." in text

    finally:
        doc_path.unlink()


def test_paragraphs_iteration() -> None:
    """Test iterating through paragraphs."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Count heading paragraphs
        heading_count = sum(1 for p in doc.paragraphs if p.is_heading())
        assert heading_count == 3  # Section 1, Section 1.1, Section 2

        # Find paragraphs containing specific text
        section1_paras = [p for p in doc.paragraphs if "section 1" in p.text.lower()]
        assert len(section1_paras) >= 2  # "Section 1" and "Section 1.1"

    finally:
        doc_path.unlink()


def test_sections_iteration() -> None:
    """Test iterating through sections."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Find section by heading text
        for section in doc.sections:
            if section.heading_text == "Section 1":
                # Found it!
                assert len(section.paragraphs) == 2
                assert section.heading_level == 1
                break
        else:
            assert False, "Section 1 not found"

    finally:
        doc_path.unlink()


def test_section_contains_method() -> None:
    """Test Section.contains() for searching within sections."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Get Section 1
        section1 = doc.sections[1]

        # Should contain its content
        assert section1.contains("This is the content of section 1.")

        # Should not contain content from other sections
        assert not section1.contains("Content of section 2.")

    finally:
        doc_path.unlink()


def test_agent_workflow_example() -> None:
    """Test complete agent workflow: read -> understand -> edit."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Step 1: Agent reads document to understand structure
        text = doc.get_text()
        assert "Section 1" in text

        # Step 2: Agent finds specific section
        target_section = None
        for section in doc.sections:
            if section.heading_text == "Section 1":
                target_section = section
                break

        assert target_section is not None

        # Step 3: Agent checks what's in the section
        assert target_section.contains("This is the content of section 1.")

        # Step 4: Agent makes targeted edit
        doc.insert_tracked(" Additional text.", after="This is the content of section 1.")

        # Step 5: Agent verifies edit
        new_text = doc.get_text()
        assert "Additional text." in new_text

    finally:
        doc_path.unlink()


def test_empty_document() -> None:
    """Test view properties on document with no content."""
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

        # Empty document has no paragraphs
        assert len(doc.paragraphs) == 0

        # Empty document has no sections
        assert len(doc.sections) == 0

        # get_text() returns empty string
        assert doc.get_text() == ""

    finally:
        doc_path.unlink()


def test_get_text_skip_deleted_paragraphs() -> None:
    """Test that get_text() skips paragraphs with deleted paragraph marks by default."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Delete a paragraph with tracking
        doc.delete_ref("p:1", track=True, author="TestAgent")

        # Default behavior: skip deleted paragraphs
        text = doc.get_text()

        # Should not have excessive empty lines from deleted paragraph
        # The deleted paragraph mark means the empty paragraph is skipped
        assert "\n\n\n\n" not in text  # No 3+ consecutive newlines

    finally:
        doc_path.unlink()


def test_get_text_include_deleted_paragraphs() -> None:
    """Test get_text() can include empty lines for deleted paragraphs."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Count paragraphs before deletion
        text_before = doc.get_text(skip_deleted_paragraphs=False)
        para_count_before = text_before.count("\n\n") + 1

        # Delete a paragraph with tracking
        doc.delete_ref("p:1", track=True, author="TestAgent")

        # With skip_deleted_paragraphs=False, we keep the empty line
        text_after = doc.get_text(skip_deleted_paragraphs=False)
        para_count_after = text_after.count("\n\n") + 1

        # Paragraph count should be the same (empty paragraph still counted)
        assert para_count_after == para_count_before

    finally:
        doc_path.unlink()


def test_get_text_skip_multiple_deleted_paragraphs() -> None:
    """Test that get_text() skips multiple deleted paragraphs."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Delete multiple paragraphs with tracking (in reverse order to avoid ref shifting)
        doc.delete_ref("p:3", track=True, author="TestAgent")
        doc.delete_ref("p:2", track=True, author="TestAgent")
        doc.delete_ref("p:1", track=True, author="TestAgent")

        # Default behavior: skip deleted paragraphs
        text = doc.get_text()

        # Count empty lines - with skip_deleted_paragraphs=True,
        # we shouldn't have excessive empty content
        lines = text.split("\n")
        empty_line_count = sum(1 for line in lines if not line.strip())

        # Skip behavior should reduce empty lines significantly
        text_with_empty = doc.get_text(skip_deleted_paragraphs=False)
        lines_with_empty = text_with_empty.split("\n")
        empty_line_count_with = sum(1 for line in lines_with_empty if not line.strip())

        # Should have fewer empty lines when skipping
        assert empty_line_count <= empty_line_count_with

    finally:
        doc_path.unlink()


def test_get_text_preserves_regular_empty_paragraphs() -> None:
    """Test that get_text() preserves intentional empty paragraphs (no deleted mark)."""
    doc_path = create_document_with_sections()
    try:
        doc = Document(doc_path)

        # Get text before any deletions
        text_before = doc.get_text()

        # Should have standard paragraph separations
        assert "\n\n" in text_before

        # Using skip_deleted_paragraphs=True shouldn't change anything
        # when there are no deleted paragraphs
        text_after = doc.get_text(skip_deleted_paragraphs=True)
        assert text_before == text_after

    finally:
        doc_path.unlink()
