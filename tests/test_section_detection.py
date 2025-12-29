"""
Tests for the multi-tier section detection algorithm.

These tests verify:
- Tier 1: Heading style detection
- Tier 2: Outline level property detection
- Tier 3: Heuristic detection (bold, caps, numbered, blank lines)
- Tier 4: Fallback strategies
- Configuration options
- Detection metadata and confidence levels
"""

from lxml import etree

from python_docx_redline.accessibility.sections import (
    DetectedSection,
    DetectionConfidence,
    DetectionMetadata,
    DetectionMethod,
    HeuristicConfig,
    SectionDetectionConfig,
    SectionDetector,
    create_section_nodes,
    detect_sections,
)
from python_docx_redline.accessibility.types import AccessibilityNode, ElementType, Ref

# Test XML documents

DOCUMENT_WITH_HEADING_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Title"/>
      </w:pPr>
      <w:r>
        <w:t>Document Title</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>First Section</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content under first section.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More content here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading2"/>
      </w:pPr>
      <w:r>
        <w:t>Subsection</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Subsection content.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Second Section</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second section content.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_OUTLINE_LEVELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:outlineLvl w:val="0"/>
      </w:pPr>
      <w:r>
        <w:t>Chapter One</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content for chapter one.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More chapter one content.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:outlineLvl w:val="1"/>
      </w:pPr>
      <w:r>
        <w:t>Section 1.1</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Section content.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:outlineLvl w:val="0"/>
      </w:pPr>
      <w:r>
        <w:t>Chapter Two</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Chapter two content.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_BOLD_HEADINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>Introduction</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This is the introduction paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Another paragraph in the introduction.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>Background</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This is the background section.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More background information.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_CAPS_HEADINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>INTRODUCTION</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This is the introduction paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Another paragraph in the introduction.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>BACKGROUND</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This is the background section.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More background information.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_NUMBERED_SECTIONS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>1. Introduction</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This is the introduction paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Another paragraph here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>1.1 Subsection</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Subsection content.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More subsection content.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>2. Background</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This is the background section.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More background.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_ARTICLE_SECTIONS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Article I - Definitions</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This article defines key terms.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More definitions here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Article II - Services</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Services section content.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Additional services information.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_BLANK_LINES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph of section one.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph of section one.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph of section one.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t></w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>First paragraph of section two.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph of section two.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph of section two.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_NO_STRUCTURE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Fourth paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Fifth paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


EMPTY_DOCUMENT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
  </w:body>
</w:document>"""


DOCUMENT_WITH_MIXED_BOLD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>Bold part</w:t>
      </w:r>
      <w:r>
        <w:t> not bold part</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Regular paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Another regular paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


def create_xml_root(xml_content: str) -> etree._Element:
    """Create an XML root element from content."""
    return etree.fromstring(xml_content.encode("utf-8"))


class TestTier1HeadingStyles:
    """Tests for Tier 1: Heading style detection."""

    def test_detects_heading_styles(self) -> None:
        """Test detection of explicit heading styles."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        assert len(sections) == 4  # Title, First Section, Subsection, Second Section

    def test_heading_style_confidence_is_high(self) -> None:
        """Test that heading style detection has HIGH confidence."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        for section in sections:
            assert section.metadata.confidence == DetectionConfidence.HIGH
            assert section.metadata.method == DetectionMethod.HEADING_STYLE

    def test_heading_levels_detected(self) -> None:
        """Test that heading levels are correctly extracted."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        # Title is level 1
        assert sections[0].heading_level == 1
        assert sections[0].heading_text == "Document Title"

        # Heading1 is level 1
        assert sections[1].heading_level == 1
        assert sections[1].heading_text == "First Section"

        # Heading2 is level 2
        assert sections[2].heading_level == 2
        assert sections[2].heading_text == "Subsection"

        # Heading1 is level 1
        assert sections[3].heading_level == 1
        assert sections[3].heading_text == "Second Section"

    def test_section_paragraph_counts(self) -> None:
        """Test paragraph counts in sections."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        # Title section: Title + 0 content = 1 (just heading before next heading)
        assert sections[0].paragraph_count == 1

        # First Section: Heading + 2 content = 3 (until Subsection)
        assert sections[1].paragraph_count == 3

        # Subsection: Heading + 1 content = 2 (until Second Section)
        assert sections[2].paragraph_count == 2

        # Second Section: Heading + 1 content = 2 (until end)
        assert sections[3].paragraph_count == 2

    def test_section_refs(self) -> None:
        """Test that sections have correct refs."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        assert sections[0].heading_ref is not None
        assert sections[0].heading_ref.path == "p:0"

        assert sections[1].heading_ref is not None
        assert sections[1].heading_ref.path == "p:1"


class TestTier2OutlineLevels:
    """Tests for Tier 2: Outline level detection."""

    def test_detects_outline_levels(self) -> None:
        """Test detection via w:outlineLvl property."""
        xml_root = create_xml_root(DOCUMENT_WITH_OUTLINE_LEVELS)
        sections = detect_sections(xml_root)

        assert len(sections) == 3  # Chapter One, Section 1.1, Chapter Two

    def test_outline_level_confidence_is_high(self) -> None:
        """Test that outline level detection has HIGH confidence."""
        xml_root = create_xml_root(DOCUMENT_WITH_OUTLINE_LEVELS)
        sections = detect_sections(xml_root)

        for section in sections:
            assert section.metadata.confidence == DetectionConfidence.HIGH
            assert section.metadata.method == DetectionMethod.OUTLINE_LEVEL

    def test_outline_levels_mapped_correctly(self) -> None:
        """Test that outline levels are mapped to heading levels."""
        xml_root = create_xml_root(DOCUMENT_WITH_OUTLINE_LEVELS)
        sections = detect_sections(xml_root)

        # outlineLvl=0 -> level 1
        assert sections[0].heading_level == 1
        assert sections[0].heading_text == "Chapter One"

        # outlineLvl=1 -> level 2
        assert sections[1].heading_level == 2
        assert sections[1].heading_text == "Section 1.1"

        # outlineLvl=0 -> level 1
        assert sections[2].heading_level == 1
        assert sections[2].heading_text == "Chapter Two"


class TestTier3Heuristics:
    """Tests for Tier 3: Heuristic detection."""

    def test_bold_headings_detected(self) -> None:
        """Test detection of all-bold paragraphs as headings."""
        xml_root = create_xml_root(DOCUMENT_WITH_BOLD_HEADINGS)
        sections = detect_sections(xml_root)

        assert len(sections) == 2
        assert sections[0].heading_text == "Introduction"
        assert sections[1].heading_text == "Background"

    def test_bold_heuristic_confidence_is_medium(self) -> None:
        """Test that bold heuristic has MEDIUM confidence."""
        xml_root = create_xml_root(DOCUMENT_WITH_BOLD_HEADINGS)
        sections = detect_sections(xml_root)

        for section in sections:
            assert section.metadata.confidence == DetectionConfidence.MEDIUM
            assert section.metadata.method == DetectionMethod.BOLD_HEURISTIC

    def test_mixed_bold_not_detected(self) -> None:
        """Test that paragraphs with mixed bold are not detected."""
        xml_root = create_xml_root(DOCUMENT_WITH_MIXED_BOLD)

        # Disable other heuristics to test bold only
        config = SectionDetectionConfig(
            heuristic_config=HeuristicConfig(
                detect_caps_headings=False,
                detect_numbered_sections=False,
                detect_blank_line_breaks=False,
            )
        )

        sections = detect_sections(xml_root, config)

        # Mixed bold should not create sections, fallback should kick in
        assert all(s.metadata.method != DetectionMethod.BOLD_HEURISTIC for s in sections)

    def test_caps_headings_detected(self) -> None:
        """Test detection of ALL CAPS paragraphs as headings."""
        xml_root = create_xml_root(DOCUMENT_WITH_CAPS_HEADINGS)
        sections = detect_sections(xml_root)

        assert len(sections) == 2
        assert sections[0].heading_text == "INTRODUCTION"
        assert sections[1].heading_text == "BACKGROUND"

    def test_caps_heuristic_confidence_is_medium(self) -> None:
        """Test that caps heuristic has MEDIUM confidence."""
        xml_root = create_xml_root(DOCUMENT_WITH_CAPS_HEADINGS)
        sections = detect_sections(xml_root)

        for section in sections:
            assert section.metadata.confidence == DetectionConfidence.MEDIUM
            assert section.metadata.method == DetectionMethod.CAPS_HEURISTIC

    def test_numbered_sections_detected(self) -> None:
        """Test detection of numbered section patterns."""
        xml_root = create_xml_root(DOCUMENT_WITH_NUMBERED_SECTIONS)
        sections = detect_sections(xml_root)

        assert len(sections) == 3
        assert sections[0].heading_text == "1. Introduction"
        assert sections[1].heading_text == "1.1 Subsection"
        assert sections[2].heading_text == "2. Background"

    def test_numbered_heuristic_assigns_levels(self) -> None:
        """Test that numbered patterns assign appropriate levels."""
        xml_root = create_xml_root(DOCUMENT_WITH_NUMBERED_SECTIONS)
        sections = detect_sections(xml_root)

        # 1. -> level 1
        assert sections[0].heading_level == 1

        # 1.1 -> level 2
        assert sections[1].heading_level == 2

        # 2. -> level 1
        assert sections[2].heading_level == 1

    def test_article_patterns_detected(self) -> None:
        """Test detection of Article patterns."""
        xml_root = create_xml_root(DOCUMENT_WITH_ARTICLE_SECTIONS)
        sections = detect_sections(xml_root)

        assert len(sections) == 2
        assert "Article I" in sections[0].heading_text
        assert "Article II" in sections[1].heading_text

    def test_blank_line_detection(self) -> None:
        """Test detection by blank line separators."""
        # Disable other heuristics
        config = SectionDetectionConfig(
            heuristic_config=HeuristicConfig(
                detect_bold_headings=False,
                detect_caps_headings=False,
                detect_numbered_sections=False,
                detect_blank_line_breaks=True,
            )
        )

        xml_root = create_xml_root(DOCUMENT_WITH_BLANK_LINES)
        sections = detect_sections(xml_root, config)

        assert len(sections) == 2
        assert sections[0].metadata.method == DetectionMethod.BLANK_LINE_HEURISTIC
        assert sections[1].metadata.method == DetectionMethod.BLANK_LINE_HEURISTIC


class TestTier4Fallback:
    """Tests for Tier 4: Fallback strategies."""

    def test_single_section_fallback(self) -> None:
        """Test fallback to single section for small documents."""
        xml_root = create_xml_root(DOCUMENT_NO_STRUCTURE)

        # Disable heuristics to force fallback
        config = SectionDetectionConfig(
            use_heuristics=False,
        )

        sections = detect_sections(xml_root, config)

        assert len(sections) == 1
        assert sections[0].metadata.method == DetectionMethod.FALLBACK_SINGLE
        assert sections[0].metadata.confidence == DetectionConfidence.LOW

    def test_chunked_fallback_for_large_documents(self) -> None:
        """Test chunked fallback for large documents."""
        # Create a document with many paragraphs
        paragraphs = "\n".join(
            [
                f"""<w:p>
              <w:r>
                <w:t>Paragraph {i}.</w:t>
              </w:r>
            </w:p>"""
                for i in range(25)
            ]
        )

        xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {paragraphs}
  </w:body>
</w:document>"""

        xml_root = create_xml_root(xml)

        # Disable heuristics, use small chunk size
        config = SectionDetectionConfig(
            use_heuristics=False,
            fallback_chunk_size=10,
        )

        sections = detect_sections(xml_root, config)

        assert len(sections) == 3  # 25 paragraphs / 10 chunk size = 3 chunks
        assert all(s.metadata.method == DetectionMethod.FALLBACK_CHUNKED for s in sections)
        assert all(s.metadata.confidence == DetectionConfidence.LOW for s in sections)


class TestEmptyAndEdgeCases:
    """Tests for edge cases and empty documents."""

    def test_empty_document_returns_no_sections(self) -> None:
        """Test that empty document returns empty list."""
        xml_root = create_xml_root(EMPTY_DOCUMENT)
        sections = detect_sections(xml_root)

        assert sections == []

    def test_no_body_returns_no_sections(self) -> None:
        """Test document without body element."""
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:document>"""

        xml_root = create_xml_root(xml)
        sections = detect_sections(xml_root)

        assert sections == []

    def test_single_paragraph_document(self) -> None:
        """Test document with single paragraph."""
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Only paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

        xml_root = create_xml_root(xml)
        sections = detect_sections(xml_root)

        assert len(sections) == 1
        assert sections[0].paragraph_count == 1


class TestConfiguration:
    """Tests for configuration options."""

    def test_disable_heading_styles(self) -> None:
        """Test disabling heading style detection."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)

        config = SectionDetectionConfig(
            use_heading_styles=False,
            use_outline_levels=False,
            use_heuristics=False,
        )

        sections = detect_sections(xml_root, config)

        # Should use fallback
        assert all(
            s.metadata.method in (DetectionMethod.FALLBACK_SINGLE, DetectionMethod.FALLBACK_CHUNKED)
            for s in sections
        )

    def test_disable_heuristics(self) -> None:
        """Test disabling heuristic detection."""
        xml_root = create_xml_root(DOCUMENT_WITH_BOLD_HEADINGS)

        config = SectionDetectionConfig(
            use_heading_styles=True,
            use_outline_levels=True,
            use_heuristics=False,
        )

        sections = detect_sections(xml_root, config)

        # No heading styles or outline levels, so fallback
        assert all(
            s.metadata.method in (DetectionMethod.FALLBACK_SINGLE, DetectionMethod.FALLBACK_CHUNKED)
            for s in sections
        )

    def test_custom_max_heading_length(self) -> None:
        """Test custom max heading length."""
        # Create document with long bold paragraph
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>This is a very long bold paragraph that exceeds the maximum heading length and should not be detected as a heading.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More content.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

        xml_root = create_xml_root(xml)

        # Default max_heading_length is 100, so this should be detected
        config_long = SectionDetectionConfig(
            heuristic_config=HeuristicConfig(max_heading_length=200)
        )
        sections_long = detect_sections(xml_root, config_long)

        # With shorter max, should not detect bold heuristic
        config_short = SectionDetectionConfig(
            heuristic_config=HeuristicConfig(max_heading_length=50)
        )
        sections_short = detect_sections(xml_root, config_short)

        assert any(s.metadata.method == DetectionMethod.BOLD_HEURISTIC for s in sections_long)
        # Short max should not detect the long bold paragraph as a heading
        assert not any(s.metadata.method == DetectionMethod.BOLD_HEURISTIC for s in sections_short)

    def test_custom_min_section_paragraphs(self) -> None:
        """Test min_section_paragraphs filtering."""
        xml_root = create_xml_root(DOCUMENT_WITH_BOLD_HEADINGS)

        # Default min is 2, so all sections should be detected
        config_low = SectionDetectionConfig(
            heuristic_config=HeuristicConfig(min_section_paragraphs=1)
        )
        sections_low = detect_sections(xml_root, config_low)

        # High min should filter out small sections
        config_high = SectionDetectionConfig(
            heuristic_config=HeuristicConfig(min_section_paragraphs=5)
        )
        sections_high = detect_sections(xml_root, config_high)

        assert len(sections_low) == 2
        # High min might cause fallback
        assert len(sections_high) <= len(sections_low)

    def test_custom_numbering_patterns(self) -> None:
        """Test custom numbering patterns."""
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>[A] First Section</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Content here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More content.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>[B] Second Section</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More content here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Even more content.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

        xml_root = create_xml_root(xml)

        # Default patterns won't match [A], [B]
        config_default = SectionDetectionConfig()
        sections_default = detect_sections(xml_root, config_default)

        # Default config should not detect [A], [B] as numbered sections
        assert not any(
            s.metadata.method == DetectionMethod.NUMBERED_HEURISTIC for s in sections_default
        )

        # Custom pattern should match
        config_custom = SectionDetectionConfig(
            heuristic_config=HeuristicConfig(
                numbering_patterns=[r"^\[[A-Z]\]"],
                detect_bold_headings=False,
                detect_caps_headings=False,
            )
        )
        sections_custom = detect_sections(xml_root, config_custom)

        assert len(sections_custom) == 2
        assert sections_custom[0].metadata.method == DetectionMethod.NUMBERED_HEURISTIC


class TestDetectionMetadata:
    """Tests for detection metadata."""

    def test_metadata_has_method(self) -> None:
        """Test that metadata includes detection method."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        for section in sections:
            assert section.metadata.method is not None
            assert isinstance(section.metadata.method, DetectionMethod)

    def test_metadata_has_confidence(self) -> None:
        """Test that metadata includes confidence level."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        for section in sections:
            assert section.metadata.confidence is not None
            assert isinstance(section.metadata.confidence, DetectionConfidence)

    def test_metadata_has_details(self) -> None:
        """Test that metadata includes descriptive details."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        for section in sections:
            assert section.metadata.details is not None
            assert len(section.metadata.details) > 0


class TestDetectedSectionDataclass:
    """Tests for DetectedSection dataclass."""

    def test_section_has_all_fields(self) -> None:
        """Test that DetectedSection has all required fields."""
        section = DetectedSection(
            heading_text="Test Heading",
            heading_ref=Ref(path="p:0"),
            heading_level=1,
            start_index=0,
            end_index=5,
            paragraph_count=5,
            metadata=DetectionMetadata(
                method=DetectionMethod.HEADING_STYLE,
                confidence=DetectionConfidence.HIGH,
            ),
        )

        assert section.heading_text == "Test Heading"
        assert section.heading_ref.path == "p:0"
        assert section.heading_level == 1
        assert section.start_index == 0
        assert section.end_index == 5
        assert section.paragraph_count == 5

    def test_detection_metadata_optional_fields(self) -> None:
        """Test optional fields in DetectionMetadata."""
        metadata = DetectionMetadata(
            method=DetectionMethod.BOLD_HEURISTIC,
            confidence=DetectionConfidence.MEDIUM,
            score=0.85,
            details="Custom details",
        )

        assert metadata.score == 0.85
        assert metadata.details == "Custom details"


class TestHeuristicConfigDataclass:
    """Tests for HeuristicConfig dataclass."""

    def test_default_values(self) -> None:
        """Test default values for HeuristicConfig."""
        config = HeuristicConfig()

        assert config.detect_bold_headings is True
        assert config.detect_caps_headings is True
        assert config.detect_numbered_sections is True
        assert config.detect_blank_line_breaks is True
        assert config.min_section_paragraphs == 2
        assert config.max_heading_length == 100
        assert len(config.numbering_patterns) > 0

    def test_custom_values(self) -> None:
        """Test custom values for HeuristicConfig."""
        config = HeuristicConfig(
            detect_bold_headings=False,
            min_section_paragraphs=5,
            max_heading_length=50,
        )

        assert config.detect_bold_headings is False
        assert config.min_section_paragraphs == 5
        assert config.max_heading_length == 50


class TestSectionDetectionConfigDataclass:
    """Tests for SectionDetectionConfig dataclass."""

    def test_default_values(self) -> None:
        """Test default values for SectionDetectionConfig."""
        config = SectionDetectionConfig()

        assert config.use_heading_styles is True
        assert config.use_outline_levels is True
        assert config.use_heuristics is True
        assert config.use_fallback is True
        assert config.fallback_chunk_size == 10
        assert config.heuristic_config is not None

    def test_nested_heuristic_config(self) -> None:
        """Test nested HeuristicConfig."""
        config = SectionDetectionConfig(
            heuristic_config=HeuristicConfig(
                detect_bold_headings=False,
            )
        )

        assert config.heuristic_config.detect_bold_headings is False


class TestSectionDetectorClass:
    """Tests for SectionDetector class."""

    def test_detector_with_default_config(self) -> None:
        """Test SectionDetector with default config."""
        detector = SectionDetector()
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)

        sections = detector.detect(xml_root)

        assert len(sections) > 0

    def test_detector_with_custom_config(self) -> None:
        """Test SectionDetector with custom config."""
        config = SectionDetectionConfig(use_heuristics=False)
        detector = SectionDetector(config)
        xml_root = create_xml_root(DOCUMENT_WITH_BOLD_HEADINGS)

        sections = detector.detect(xml_root)

        # Without heuristics, should fall back
        assert all(
            s.metadata.method in (DetectionMethod.FALLBACK_SINGLE, DetectionMethod.FALLBACK_CHUNKED)
            for s in sections
        )


class TestCreateSectionNodes:
    """Tests for create_section_nodes helper function."""

    def test_creates_section_nodes(self) -> None:
        """Test that create_section_nodes creates AccessibilityNodes."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        # Create mock paragraph nodes
        paragraph_nodes = [
            AccessibilityNode(
                ref=Ref(path=f"p:{i}"),
                element_type=ElementType.PARAGRAPH,
                text=f"Paragraph {i}",
            )
            for i in range(8)
        ]

        section_nodes = create_section_nodes(sections, paragraph_nodes)

        assert len(section_nodes) == len(sections)
        for node in section_nodes:
            assert node.element_type == ElementType.SECTION

    def test_section_nodes_have_properties(self) -> None:
        """Test that section nodes have detection properties."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        paragraph_nodes = [
            AccessibilityNode(
                ref=Ref(path=f"p:{i}"),
                element_type=ElementType.PARAGRAPH,
                text=f"Paragraph {i}",
            )
            for i in range(8)
        ]

        section_nodes = create_section_nodes(sections, paragraph_nodes)

        for node in section_nodes:
            assert "detection_method" in node.properties
            assert "confidence" in node.properties
            assert "paragraph_count" in node.properties

    def test_section_nodes_have_children(self) -> None:
        """Test that section nodes contain paragraph children."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        paragraph_nodes = [
            AccessibilityNode(
                ref=Ref(path=f"p:{i}"),
                element_type=ElementType.PARAGRAPH,
                text=f"Paragraph {i}",
            )
            for i in range(8)
        ]

        section_nodes = create_section_nodes(sections, paragraph_nodes)

        # Each section should have children matching paragraph_count
        for i, node in enumerate(section_nodes):
            expected_count = sections[i].paragraph_count
            assert len(node.children) == expected_count


class TestConvenienceFunction:
    """Tests for the detect_sections convenience function."""

    def test_detect_sections_function(self) -> None:
        """Test detect_sections convenience function."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)
        sections = detect_sections(xml_root)

        assert len(sections) > 0
        assert all(isinstance(s, DetectedSection) for s in sections)

    def test_detect_sections_with_config(self) -> None:
        """Test detect_sections with config parameter."""
        xml_root = create_xml_root(DOCUMENT_WITH_HEADING_STYLES)

        config = SectionDetectionConfig(use_heuristics=False)
        sections = detect_sections(xml_root, config)

        assert len(sections) > 0
