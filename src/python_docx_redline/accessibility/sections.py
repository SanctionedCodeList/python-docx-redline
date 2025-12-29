"""
Multi-tier section detection for Word documents.

This module provides algorithms for detecting document sections even when
proper heading styles are not used. It implements a multi-tier approach:

- Tier 1: Heading styles (Heading1, Heading2, etc.) - HIGH confidence
- Tier 2: Outline level property (w:outlineLvl) - HIGH confidence
- Tier 3: Heuristics (bold, caps, numbered sections, blank lines) - MEDIUM confidence
- Tier 4: Fallback (single section or chunking) - LOW confidence
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum, auto
from typing import TYPE_CHECKING

from lxml import etree

from ..constants import WORD_NAMESPACE, w
from .types import AccessibilityNode, ElementType, Ref

if TYPE_CHECKING:
    pass


class DetectionMethod(Enum):
    """Method used to detect a section heading."""

    HEADING_STYLE = auto()  # Explicit heading style (Heading1, Heading2, etc.)
    OUTLINE_LEVEL = auto()  # w:outlineLvl property
    BOLD_HEURISTIC = auto()  # All-bold short paragraph
    CAPS_HEURISTIC = auto()  # ALL CAPS short paragraph
    NUMBERED_HEURISTIC = auto()  # Numbered section pattern (1., Article I, etc.)
    BLANK_LINE_HEURISTIC = auto()  # Section after blank line separator
    FALLBACK_SINGLE = auto()  # Single section fallback
    FALLBACK_CHUNKED = auto()  # Chunked sections fallback


class DetectionConfidence(Enum):
    """Confidence level of section detection."""

    HIGH = "high"  # Explicit styles or outline levels
    MEDIUM = "medium"  # Heuristic detection
    LOW = "low"  # Fallback methods


# Mapping from detection method to confidence
METHOD_TO_CONFIDENCE: dict[DetectionMethod, DetectionConfidence] = {
    DetectionMethod.HEADING_STYLE: DetectionConfidence.HIGH,
    DetectionMethod.OUTLINE_LEVEL: DetectionConfidence.HIGH,
    DetectionMethod.BOLD_HEURISTIC: DetectionConfidence.MEDIUM,
    DetectionMethod.CAPS_HEURISTIC: DetectionConfidence.MEDIUM,
    DetectionMethod.NUMBERED_HEURISTIC: DetectionConfidence.MEDIUM,
    DetectionMethod.BLANK_LINE_HEURISTIC: DetectionConfidence.MEDIUM,
    DetectionMethod.FALLBACK_SINGLE: DetectionConfidence.LOW,
    DetectionMethod.FALLBACK_CHUNKED: DetectionConfidence.LOW,
}


@dataclass
class HeuristicConfig:
    """Configuration for heuristic section detection.

    Controls which heuristics are enabled and their parameters.

    Attributes:
        detect_bold_headings: Detect all-bold short paragraphs as headings
        detect_caps_headings: Detect ALL CAPS short paragraphs as headings
        detect_numbered_sections: Detect numbered section patterns
        detect_blank_line_breaks: Detect section breaks by blank line separators
        min_section_paragraphs: Minimum paragraphs to form a section (avoid micro-sections)
        max_heading_length: Maximum characters for a heuristic heading
        numbering_patterns: Regex patterns for numbered section detection
    """

    detect_bold_headings: bool = True
    detect_caps_headings: bool = True
    detect_numbered_sections: bool = True
    detect_blank_line_breaks: bool = True

    min_section_paragraphs: int = 2
    max_heading_length: int = 100

    numbering_patterns: list[str] = field(
        default_factory=lambda: [
            r"^\d+\.\s",  # 1. Section
            r"^\d+\.\d+\s",  # 1.1 Section
            r"^Article\s+[IVXLCDM]+",  # Article I, Article II
            r"^Article\s+\d+",  # Article 1, Article 2
            r"^Section\s+\d+",  # Section 1, Section 2
            r"^SECTION\s+\d+",  # SECTION 1
            r"^Part\s+[IVXLCDM]+",  # Part I, Part II
            r"^Part\s+\d+",  # Part 1, Part 2
            r"^Chapter\s+\d+",  # Chapter 1
            r"^\([a-z]\)",  # (a), (b), (c)
            r"^\([0-9]+\)",  # (1), (2), (3)
        ]
    )


@dataclass
class SectionDetectionConfig:
    """Configuration for section detection behavior.

    Controls how sections are detected and which tiers are enabled.

    Attributes:
        use_heading_styles: Enable Tier 1 (heading style detection)
        use_outline_levels: Enable Tier 2 (outline level property)
        use_heuristics: Enable Tier 3 (heuristic detection)
        use_fallback: Enable Tier 4 (fallback strategies)
        heuristic_config: Configuration for heuristic detection
        fallback_chunk_size: Paragraphs per chunk in fallback mode
    """

    use_heading_styles: bool = True
    use_outline_levels: bool = True
    use_heuristics: bool = True
    use_fallback: bool = True

    heuristic_config: HeuristicConfig = field(default_factory=HeuristicConfig)

    fallback_chunk_size: int = 10


@dataclass
class DetectionMetadata:
    """Metadata about how a section was detected.

    Attributes:
        method: The detection method used
        confidence: Confidence level of the detection
        score: Optional numeric score (for ranked heuristics)
        details: Additional details about the detection
    """

    method: DetectionMethod
    confidence: DetectionConfidence
    score: float | None = None
    details: str | None = None


@dataclass
class DetectedSection:
    """A detected document section.

    Attributes:
        heading_text: Text of the section heading (if any)
        heading_ref: Ref of the heading paragraph
        heading_level: Detected heading level (1-9)
        start_index: Index of first paragraph in section
        end_index: Index of last paragraph in section (exclusive)
        paragraph_count: Number of paragraphs in section
        metadata: Detection metadata
    """

    heading_text: str
    heading_ref: Ref | None
    heading_level: int
    start_index: int
    end_index: int
    paragraph_count: int
    metadata: DetectionMetadata


class SectionDetector:
    """Multi-tier section detection for Word documents.

    Implements a cascading detection algorithm:
    1. First tries heading styles (highest confidence)
    2. Falls back to outline levels
    3. Uses heuristics for unstyled documents
    4. Final fallback to single section or chunking

    Example:
        >>> detector = SectionDetector()
        >>> sections = detector.detect(xml_root)
        >>> for section in sections:
        ...     print(f"{section.heading_text} (confidence: {section.metadata.confidence})")
    """

    def __init__(self, config: SectionDetectionConfig | None = None) -> None:
        """Initialize the section detector.

        Args:
            config: Configuration for detection behavior
        """
        self.config = config or SectionDetectionConfig()
        self._compiled_patterns: list[re.Pattern[str]] | None = None

    def detect(self, xml_root: etree._Element) -> list[DetectedSection]:
        """Detect sections in a document.

        Args:
            xml_root: Root element of the document XML

        Returns:
            List of detected sections in document order
        """
        # Find the body element
        body = xml_root.find(f".//{w('body')}")
        if body is None:
            return []

        # Get all paragraphs from the body
        paragraphs = list(body.findall(f"./{w('p')}"))
        if not paragraphs:
            return []

        # Try each tier in order until we get sections
        sections: list[DetectedSection] = []

        # Tier 1: Heading styles
        if self.config.use_heading_styles:
            sections = self._detect_by_heading_styles(paragraphs)
            if sections:
                return sections

        # Tier 2: Outline levels
        if self.config.use_outline_levels:
            sections = self._detect_by_outline_levels(paragraphs)
            if sections:
                return sections

        # Tier 3: Heuristics
        if self.config.use_heuristics:
            sections = self._detect_by_heuristics(paragraphs)
            if sections:
                return sections

        # Tier 4: Fallback
        if self.config.use_fallback:
            sections = self._detect_fallback(paragraphs)
            return sections

        return []

    def _detect_by_heading_styles(self, paragraphs: list[etree._Element]) -> list[DetectedSection]:
        """Tier 1: Detect sections by heading styles.

        Args:
            paragraphs: List of paragraph elements

        Returns:
            List of detected sections, empty if no headings found
        """
        # Find all paragraphs with heading styles
        heading_indices: list[tuple[int, int, str]] = []  # (index, level, text)

        for idx, p_elem in enumerate(paragraphs):
            style = self._get_paragraph_style(p_elem)
            if style:
                level = self._get_heading_level_from_style(style)
                if level is not None:
                    text = self._extract_text(p_elem)
                    heading_indices.append((idx, level, text))

        if not heading_indices:
            return []

        return self._build_sections_from_headings(
            paragraphs, heading_indices, DetectionMethod.HEADING_STYLE
        )

    def _detect_by_outline_levels(self, paragraphs: list[etree._Element]) -> list[DetectedSection]:
        """Tier 2: Detect sections by outline level property.

        The w:outlineLvl property in paragraph properties indicates
        a paragraph is part of the document outline.

        Args:
            paragraphs: List of paragraph elements

        Returns:
            List of detected sections, empty if no outline levels found
        """
        heading_indices: list[tuple[int, int, str]] = []

        for idx, p_elem in enumerate(paragraphs):
            level = self._get_outline_level(p_elem)
            if level is not None:
                text = self._extract_text(p_elem)
                heading_indices.append((idx, level + 1, text))  # outlineLvl is 0-based

        if not heading_indices:
            return []

        return self._build_sections_from_headings(
            paragraphs, heading_indices, DetectionMethod.OUTLINE_LEVEL
        )

    def _detect_by_heuristics(self, paragraphs: list[etree._Element]) -> list[DetectedSection]:
        """Tier 3: Detect sections using heuristics.

        Combines multiple heuristics:
        - Bold short paragraphs
        - ALL CAPS short paragraphs
        - Numbered section patterns
        - Blank line separators

        Args:
            paragraphs: List of paragraph elements

        Returns:
            List of detected sections, empty if no heuristic matches
        """
        heuristic_config = self.config.heuristic_config
        heading_indices: list[tuple[int, int, str, DetectionMethod]] = []

        for idx, p_elem in enumerate(paragraphs):
            text = self._extract_text(p_elem)

            # Skip empty paragraphs and long paragraphs
            if not text.strip():
                continue
            if len(text) > heuristic_config.max_heading_length:
                continue

            # Check each heuristic
            method: DetectionMethod | None = None

            # Check numbered sections first (most reliable heuristic)
            if heuristic_config.detect_numbered_sections:
                numbered_level = self._check_numbered_pattern(text)
                if numbered_level is not None:
                    method = DetectionMethod.NUMBERED_HEURISTIC
                    heading_indices.append((idx, numbered_level, text, method))
                    continue

            # Check bold heuristic
            if heuristic_config.detect_bold_headings:
                if self._is_all_bold(p_elem):
                    method = DetectionMethod.BOLD_HEURISTIC
                    heading_indices.append((idx, 1, text, method))
                    continue

            # Check caps heuristic
            if heuristic_config.detect_caps_headings:
                if self._is_all_caps(text):
                    method = DetectionMethod.CAPS_HEURISTIC
                    heading_indices.append((idx, 1, text, method))
                    continue

        if not heading_indices:
            # Try blank line detection as last resort
            if heuristic_config.detect_blank_line_breaks:
                return self._detect_by_blank_lines(paragraphs)
            return []

        # Filter out headings that create micro-sections
        heading_indices = self._filter_micro_sections(
            heading_indices, len(paragraphs), heuristic_config.min_section_paragraphs
        )

        if not heading_indices:
            return []

        # Build sections with method-specific metadata
        return self._build_sections_from_heuristic_headings(paragraphs, heading_indices)

    def _detect_by_blank_lines(self, paragraphs: list[etree._Element]) -> list[DetectedSection]:
        """Detect sections by blank line separators.

        Args:
            paragraphs: List of paragraph elements

        Returns:
            List of detected sections
        """
        sections: list[DetectedSection] = []
        section_start = 0
        section_number = 1

        for idx, p_elem in enumerate(paragraphs):
            text = self._extract_text(p_elem)

            # Check for blank paragraph (section break)
            if not text.strip() and idx > section_start:
                # End current section
                if idx > section_start:
                    # Get first paragraph text as heading
                    first_text = self._extract_text(paragraphs[section_start])
                    heading_text = first_text[:50] if first_text else f"Section {section_number}"

                    sections.append(
                        DetectedSection(
                            heading_text=heading_text,
                            heading_ref=Ref(path=f"p:{section_start}"),
                            heading_level=1,
                            start_index=section_start,
                            end_index=idx,
                            paragraph_count=idx - section_start,
                            metadata=DetectionMetadata(
                                method=DetectionMethod.BLANK_LINE_HEURISTIC,
                                confidence=DetectionConfidence.MEDIUM,
                                details="Detected by blank line separator",
                            ),
                        )
                    )
                    section_number += 1

                # Start new section after blank line
                section_start = idx + 1

        # Handle final section
        if section_start < len(paragraphs):
            first_text = self._extract_text(paragraphs[section_start])
            heading_text = first_text[:50] if first_text else f"Section {section_number}"

            sections.append(
                DetectedSection(
                    heading_text=heading_text,
                    heading_ref=Ref(path=f"p:{section_start}"),
                    heading_level=1,
                    start_index=section_start,
                    end_index=len(paragraphs),
                    paragraph_count=len(paragraphs) - section_start,
                    metadata=DetectionMetadata(
                        method=DetectionMethod.BLANK_LINE_HEURISTIC,
                        confidence=DetectionConfidence.MEDIUM,
                        details="Detected by blank line separator",
                    ),
                )
            )

        # Filter micro-sections
        min_paragraphs = self.config.heuristic_config.min_section_paragraphs
        sections = [s for s in sections if s.paragraph_count >= min_paragraphs]

        return sections if sections else []

    def _detect_fallback(self, paragraphs: list[etree._Element]) -> list[DetectedSection]:
        """Tier 4: Fallback section detection.

        Either treats the entire document as one section or
        chunks into fixed-size sections.

        Args:
            paragraphs: List of paragraph elements

        Returns:
            List of detected sections
        """
        total_paragraphs = len(paragraphs)
        chunk_size = self.config.fallback_chunk_size

        # If document is small enough, treat as single section
        if total_paragraphs <= chunk_size * 2:
            first_text = self._extract_text(paragraphs[0]) if paragraphs else "Document"
            return [
                DetectedSection(
                    heading_text=first_text[:50] if first_text else "Document",
                    heading_ref=Ref(path="p:0"),
                    heading_level=1,
                    start_index=0,
                    end_index=total_paragraphs,
                    paragraph_count=total_paragraphs,
                    metadata=DetectionMetadata(
                        method=DetectionMethod.FALLBACK_SINGLE,
                        confidence=DetectionConfidence.LOW,
                        details="Single section fallback - no structure detected",
                    ),
                )
            ]

        # Chunk into fixed-size sections
        sections: list[DetectedSection] = []
        chunk_number = 1

        for start in range(0, total_paragraphs, chunk_size):
            end = min(start + chunk_size, total_paragraphs)

            first_text = self._extract_text(paragraphs[start])
            heading_text = first_text[:50] if first_text else f"Section {chunk_number}"

            sections.append(
                DetectedSection(
                    heading_text=heading_text,
                    heading_ref=Ref(path=f"p:{start}"),
                    heading_level=1,
                    start_index=start,
                    end_index=end,
                    paragraph_count=end - start,
                    metadata=DetectionMetadata(
                        method=DetectionMethod.FALLBACK_CHUNKED,
                        confidence=DetectionConfidence.LOW,
                        details=f"Chunked fallback (size={chunk_size})",
                    ),
                )
            )
            chunk_number += 1

        return sections

    def _build_sections_from_headings(
        self,
        paragraphs: list[etree._Element],
        heading_indices: list[tuple[int, int, str]],
        method: DetectionMethod,
    ) -> list[DetectedSection]:
        """Build section objects from detected headings.

        Args:
            paragraphs: All paragraphs
            heading_indices: List of (index, level, text) tuples
            method: Detection method used

        Returns:
            List of DetectedSection objects
        """
        sections: list[DetectedSection] = []
        total_paragraphs = len(paragraphs)
        confidence = METHOD_TO_CONFIDENCE[method]

        for i, (idx, level, text) in enumerate(heading_indices):
            # Determine section end
            if i + 1 < len(heading_indices):
                end_idx = heading_indices[i + 1][0]
            else:
                end_idx = total_paragraphs

            sections.append(
                DetectedSection(
                    heading_text=text,
                    heading_ref=Ref(path=f"p:{idx}"),
                    heading_level=level,
                    start_index=idx,
                    end_index=end_idx,
                    paragraph_count=end_idx - idx,
                    metadata=DetectionMetadata(
                        method=method,
                        confidence=confidence,
                        details=f"Detected by {method.name.lower().replace('_', ' ')}",
                    ),
                )
            )

        return sections

    def _build_sections_from_heuristic_headings(
        self,
        paragraphs: list[etree._Element],
        heading_indices: list[tuple[int, int, str, DetectionMethod]],
    ) -> list[DetectedSection]:
        """Build section objects from heuristic-detected headings.

        Args:
            paragraphs: All paragraphs
            heading_indices: List of (index, level, text, method) tuples

        Returns:
            List of DetectedSection objects
        """
        sections: list[DetectedSection] = []
        total_paragraphs = len(paragraphs)

        for i, (idx, level, text, method) in enumerate(heading_indices):
            # Determine section end
            if i + 1 < len(heading_indices):
                end_idx = heading_indices[i + 1][0]
            else:
                end_idx = total_paragraphs

            confidence = METHOD_TO_CONFIDENCE[method]

            sections.append(
                DetectedSection(
                    heading_text=text,
                    heading_ref=Ref(path=f"p:{idx}"),
                    heading_level=level,
                    start_index=idx,
                    end_index=end_idx,
                    paragraph_count=end_idx - idx,
                    metadata=DetectionMetadata(
                        method=method,
                        confidence=confidence,
                        details=f"Detected by {method.name.lower().replace('_', ' ')}",
                    ),
                )
            )

        return sections

    def _filter_micro_sections(
        self,
        heading_indices: list[tuple[int, int, str, DetectionMethod]],
        total_paragraphs: int,
        min_paragraphs: int,
    ) -> list[tuple[int, int, str, DetectionMethod]]:
        """Filter out headings that would create micro-sections.

        Args:
            heading_indices: Candidate headings
            total_paragraphs: Total paragraph count
            min_paragraphs: Minimum paragraphs per section

        Returns:
            Filtered heading indices
        """
        if not heading_indices:
            return []

        filtered: list[tuple[int, int, str, DetectionMethod]] = []

        for i, heading in enumerate(heading_indices):
            idx = heading[0]

            # Calculate section size
            if i + 1 < len(heading_indices):
                next_idx = heading_indices[i + 1][0]
            else:
                next_idx = total_paragraphs

            section_size = next_idx - idx

            # Keep if section is large enough
            if section_size >= min_paragraphs:
                filtered.append(heading)

        return filtered

    def _get_paragraph_style(self, p_elem: etree._Element) -> str | None:
        """Get the style name for a paragraph."""
        para_props = p_elem.find(f"./{w('pPr')}")
        if para_props is not None:
            para_style = para_props.find(f"./{w('pStyle')}")
            if para_style is not None:
                return para_style.get(f"{{{WORD_NAMESPACE}}}val")
        return None

    def _get_heading_level_from_style(self, style: str) -> int | None:
        """Determine heading level from style name."""
        if not style:
            return None

        style_lower = style.lower()

        # Handle "HeadingN" and "Heading N" patterns
        if style_lower.startswith("heading"):
            suffix = style_lower[7:].strip()
            if suffix.isdigit():
                return int(suffix)

        # Handle Title as level 1
        if style_lower == "title":
            return 1

        # Handle Subtitle as level 2
        if style_lower == "subtitle":
            return 2

        return None

    def _get_outline_level(self, p_elem: etree._Element) -> int | None:
        """Get the outline level property of a paragraph.

        Args:
            p_elem: Paragraph element

        Returns:
            Outline level (0-8) or None if not set
        """
        para_props = p_elem.find(f"./{w('pPr')}")
        if para_props is not None:
            outline_lvl = para_props.find(f"./{w('outlineLvl')}")
            if outline_lvl is not None:
                val = outline_lvl.get(f"{{{WORD_NAMESPACE}}}val")
                if val is not None and val.isdigit():
                    return int(val)
        return None

    def _check_numbered_pattern(self, text: str) -> int | None:
        """Check if text matches a numbered section pattern.

        Args:
            text: Paragraph text

        Returns:
            Heading level (1-3 based on pattern) or None if no match
        """
        if self._compiled_patterns is None:
            patterns = self.config.heuristic_config.numbering_patterns
            self._compiled_patterns = [re.compile(p, re.IGNORECASE) for p in patterns]

        text_stripped = text.strip()

        for i, pattern in enumerate(self._compiled_patterns):
            if pattern.match(text_stripped):
                # Assign level based on pattern type
                # Single number (1., Article I, Section 1) -> level 1
                # Sub-number (1.1, (a)) -> level 2
                if i in (1, 9, 10):  # 1.1, (a), (1) patterns
                    return 2
                return 1

        return None

    def _is_all_bold(self, p_elem: etree._Element) -> bool:
        """Check if all text in paragraph is bold.

        Args:
            p_elem: Paragraph element

        Returns:
            True if all text runs are bold
        """
        runs = list(p_elem.iter(w("r")))
        if not runs:
            return False

        for r_elem in runs:
            # Check if run has text
            has_text = False
            for t_elem in r_elem.findall(f"./{w('t')}"):
                if t_elem.text and t_elem.text.strip():
                    has_text = True
                    break

            if not has_text:
                continue

            # Check for bold property
            run_props = r_elem.find(f"./{w('rPr')}")
            if run_props is None:
                return False

            bold_elem = run_props.find(f"./{w('b')}")
            if bold_elem is None:
                return False

            # Check for explicit bold=false
            val = bold_elem.get(f"{{{WORD_NAMESPACE}}}val")
            if val is not None and val.lower() in ("false", "0"):
                return False

        return True

    def _is_all_caps(self, text: str) -> bool:
        """Check if text is all caps (and has letters).

        Args:
            text: Paragraph text

        Returns:
            True if text is all uppercase
        """
        # Must have at least 2 letters
        letters = [c for c in text if c.isalpha()]
        if len(letters) < 2:
            return False

        return text == text.upper()

    def _extract_text(self, elem: etree._Element) -> str:
        """Extract all text from an element.

        Args:
            elem: XML element

        Returns:
            Concatenated text content
        """
        text_parts = []

        for t_elem in elem.iter(w("t")):
            if t_elem.text:
                text_parts.append(t_elem.text)

        return "".join(text_parts)


def detect_sections(
    xml_root: etree._Element,
    config: SectionDetectionConfig | None = None,
) -> list[DetectedSection]:
    """Convenience function to detect sections in a document.

    Args:
        xml_root: Root element of the document XML
        config: Optional detection configuration

    Returns:
        List of detected sections

    Example:
        >>> from lxml import etree
        >>> xml_root = etree.fromstring(document_xml)
        >>> sections = detect_sections(xml_root)
        >>> for section in sections:
        ...     print(f"{section.heading_text}: {section.paragraph_count} paragraphs")
    """
    detector = SectionDetector(config)
    return detector.detect(xml_root)


def create_section_nodes(
    sections: list[DetectedSection],
    paragraph_nodes: list[AccessibilityNode],
) -> list[AccessibilityNode]:
    """Create section AccessibilityNodes from detected sections.

    Args:
        sections: Detected sections
        paragraph_nodes: List of paragraph nodes

    Returns:
        List of section nodes containing their paragraph children

    Example:
        >>> sections = detect_sections(xml_root)
        >>> section_nodes = create_section_nodes(sections, paragraph_nodes)
    """
    section_nodes: list[AccessibilityNode] = []

    for idx, section in enumerate(sections):
        # Create section node
        section_node = AccessibilityNode(
            ref=Ref(path=f"sec:{idx}"),
            element_type=ElementType.SECTION,
            text=section.heading_text,
            level=section.heading_level,
            properties={
                "detection_method": section.metadata.method.name.lower(),
                "confidence": section.metadata.confidence.value,
                "paragraph_count": str(section.paragraph_count),
            },
        )

        # Add paragraph children
        for para_idx in range(section.start_index, section.end_index):
            if para_idx < len(paragraph_nodes):
                section_node.children.append(paragraph_nodes[para_idx])

        section_nodes.append(section_node)

    return section_nodes
