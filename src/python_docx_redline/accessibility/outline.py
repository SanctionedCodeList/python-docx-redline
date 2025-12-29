"""
Outline mode and large document handling for DOCX accessibility.

This module provides progressive disclosure capabilities for large documents,
including outline mode, section expansion, and token budgeting.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from io import StringIO
from typing import TYPE_CHECKING, Literal

from lxml import etree

from ..constants import WORD_NAMESPACE, w
from .types import (
    AccessibilityNode,
    Ref,
    ViewMode,
)

if TYPE_CHECKING:
    pass


# ============================================================================
# Section Detection Configuration
# ============================================================================


@dataclass
class SectionDetectionConfig:
    """Configuration for section detection heuristics.

    Attributes:
        detect_heading_styles: Detect sections from Heading styles (Tier 1)
        detect_outline_level: Detect sections from outline level property (Tier 1)
        detect_bold_first_line: Detect bold short paragraphs as headings (Tier 3)
        detect_caps_first_line: Detect ALL CAPS short paragraphs (Tier 3)
        detect_numbered_sections: Detect numbered sections (1., 2., etc.) (Tier 3)
        min_section_paragraphs: Minimum paragraphs for a section to be valid
        max_heading_length: Maximum character length for detected headings
        numbering_patterns: Regex patterns for numbered section detection
    """

    detect_heading_styles: bool = True
    detect_outline_level: bool = True
    detect_bold_first_line: bool = True
    detect_caps_first_line: bool = True
    detect_numbered_sections: bool = True
    min_section_paragraphs: int = 2
    max_heading_length: int = 100
    numbering_patterns: list[str] = field(
        default_factory=lambda: [
            r"^\d+\.\s",  # 1. Section
            r"^\d+\.\d+\s",  # 1.1 Section
            r"^Article\s+\d+",  # Article 1
            r"^Section\s+\d+",  # Section 1
            r"^ARTICLE\s+[IVXLCDM]+",  # ARTICLE I (Roman numerals)
        ]
    )


@dataclass
class SectionInfo:
    """Information about a document section for outline mode.

    Attributes:
        ref: Section reference (sec:N)
        heading: Heading text for this section
        heading_ref: Reference to the heading paragraph
        heading_level: Heading level (1-9)
        paragraph_count: Number of paragraphs in this section
        table_count: Number of tables in this section
        tracked_change_count: Number of tracked changes in this section
        preview: Preview text from first paragraph (truncated)
        detection_method: How the section was detected
        confidence: Confidence level (high, medium, low)
        children: Nested subsections
    """

    ref: Ref
    heading: str
    heading_ref: Ref
    heading_level: int = 1
    paragraph_count: int = 0
    table_count: int = 0
    tracked_change_count: int = 0
    preview: str = ""
    detection_method: str = "heading_style"
    confidence: str = "high"
    children: list[SectionInfo] = field(default_factory=list)


# ============================================================================
# Document Size Estimation
# ============================================================================


@dataclass
class DocumentSizeInfo:
    """Information about document size for degradation decisions.

    Attributes:
        paragraph_count: Total paragraphs in document
        table_count: Total tables in document
        tracked_change_count: Total tracked changes
        estimated_tokens: Estimated token count for full tree
        recommended_mode: Recommended viewing mode
        degradation_tier: Automatic degradation tier (1-4)
    """

    paragraph_count: int = 0
    table_count: int = 0
    tracked_change_count: int = 0
    estimated_tokens: int = 0
    recommended_mode: Literal["outline", "content", "styling"] = "content"
    degradation_tier: int = 1

    @classmethod
    def from_xml(cls, xml_root: etree._Element) -> DocumentSizeInfo:
        """Calculate document size from XML root.

        Args:
            xml_root: Root element of document XML

        Returns:
            DocumentSizeInfo with calculated metrics
        """
        body = xml_root.find(f".//{w('body')}")
        if body is None:
            return cls()

        # Count paragraphs (direct children of body only)
        paragraphs = list(body.findall(f"./{w('p')}"))
        paragraph_count = len(paragraphs)

        # Count tables
        tables = list(body.findall(f"./{w('tbl')}"))
        table_count = len(tables)

        # Count tracked changes
        insertions = list(body.iter(w("ins")))
        deletions = list(body.iter(w("del")))
        tracked_change_count = len(insertions) + len(deletions)

        # Estimate tokens (rough heuristic: ~100 tokens per paragraph, ~200 per table)
        estimated_tokens = paragraph_count * 100 + table_count * 200

        # Determine degradation tier based on paragraph count
        if paragraph_count < 100:
            degradation_tier = 1
            recommended_mode = "content"
        elif paragraph_count < 300:
            degradation_tier = 2
            recommended_mode = "content"
        elif paragraph_count < 500:
            degradation_tier = 3
            recommended_mode = "outline"
        else:
            degradation_tier = 4
            recommended_mode = "outline"

        return cls(
            paragraph_count=paragraph_count,
            table_count=table_count,
            tracked_change_count=tracked_change_count,
            estimated_tokens=estimated_tokens,
            recommended_mode=recommended_mode,
            degradation_tier=degradation_tier,
        )


# ============================================================================
# Outline Tree
# ============================================================================


@dataclass
class OutlineTree:
    """Outline view of a document with section-level information only.

    OutlineTree provides O(sections) complexity instead of O(paragraphs),
    making it safe for large documents that would overflow context windows.

    Attributes:
        sections: List of top-level sections
        document_path: Path to source document
        size_info: Document size information
        detection_config: Section detection configuration used
    """

    sections: list[SectionInfo]
    document_path: str | None = None
    size_info: DocumentSizeInfo = field(default_factory=DocumentSizeInfo)
    detection_config: SectionDetectionConfig = field(default_factory=SectionDetectionConfig)

    @classmethod
    def from_xml(
        cls,
        xml_root: etree._Element,
        detection_config: SectionDetectionConfig | None = None,
        document_path: str | None = None,
    ) -> OutlineTree:
        """Build an outline tree from document XML.

        This is O(sections) not O(paragraphs) - it only scans for headings.

        Args:
            xml_root: Root element of document XML
            detection_config: Section detection configuration
            document_path: Optional path for display

        Returns:
            OutlineTree with section information
        """
        config = detection_config or SectionDetectionConfig()
        builder = _OutlineBuilder(xml_root, config)
        sections = builder.build()
        size_info = DocumentSizeInfo.from_xml(xml_root)

        return cls(
            sections=sections,
            document_path=document_path,
            size_info=size_info,
            detection_config=config,
        )

    def to_yaml(self) -> str:
        """Serialize the outline to YAML format.

        Returns:
            YAML string representation optimized for LLM consumption
        """
        writer = _OutlineYamlWriter(self)
        return writer.write()

    def get_section(self, ref: str | Ref) -> SectionInfo | None:
        """Find a section by its ref.

        Args:
            ref: Section ref to find

        Returns:
            SectionInfo or None if not found
        """
        ref_str = ref if isinstance(ref, str) else ref.path

        def _search(sections: list[SectionInfo]) -> SectionInfo | None:
            for section in sections:
                if section.ref.path == ref_str:
                    return section
                result = _search(section.children)
                if result:
                    return result
            return None

        return _search(self.sections)

    def __repr__(self) -> str:
        return (
            f"OutlineTree(sections={len(self.sections)}, "
            f"paragraphs={self.size_info.paragraph_count}, "
            f"tier={self.size_info.degradation_tier})"
        )


# ============================================================================
# Section Expansion Results
# ============================================================================


@dataclass
class SectionTree:
    """Expanded content for a single section.

    Returned by expand_section() for targeted section viewing.

    Attributes:
        section_ref: Reference to the expanded section
        heading: Section heading text
        content: List of AccessibilityNodes for section content
        paragraph_count: Number of paragraphs in section
        table_count: Number of tables in section
    """

    section_ref: Ref
    heading: str
    content: list[AccessibilityNode]
    paragraph_count: int = 0
    table_count: int = 0

    def to_yaml(self, verbosity: str = "standard") -> str:
        """Serialize section content to YAML.

        Args:
            verbosity: Output verbosity level

        Returns:
            YAML string representation
        """
        # Use the tree's YAML writer for content
        from .tree import _YamlWriter

        buffer = StringIO()
        buffer.write(f"section [ref={self.section_ref}]:\n")
        buffer.write(f'  heading: "{self.heading}"\n')
        buffer.write(f"  paragraph_count: {self.paragraph_count}\n")
        buffer.write(f"  table_count: {self.table_count}\n")
        buffer.write("  content:\n")

        # Write each content node
        writer = _YamlWriter(verbosity, ViewMode(verbosity=verbosity))
        writer._indent = 2
        for node in self.content:
            writer._write_node(node)

        buffer.write(writer.buffer.getvalue())
        return buffer.getvalue()


@dataclass
class RefTree:
    """Expanded content for specific refs.

    Returned by expand_refs() for targeted ref viewing.

    Attributes:
        refs: List of requested refs
        nodes: Mapping from ref to AccessibilityNode
        not_found: Refs that could not be resolved
    """

    refs: list[str]
    nodes: dict[str, AccessibilityNode]
    not_found: list[str] = field(default_factory=list)

    def to_yaml(self, verbosity: str = "standard") -> str:
        """Serialize expanded refs to YAML.

        Args:
            verbosity: Output verbosity level

        Returns:
            YAML string representation
        """
        from .tree import _YamlWriter

        buffer = StringIO()
        buffer.write("expanded_refs:\n")

        writer = _YamlWriter(verbosity, ViewMode(verbosity=verbosity))
        writer._indent = 1

        for ref in self.refs:
            if ref in self.nodes:
                writer._write_node(self.nodes[ref])
            else:
                buffer.write(f"  - ref: {ref}\n")
                buffer.write("    status: not_found\n")

        buffer.write(writer.buffer.getvalue())
        return buffer.getvalue()


@dataclass
class TableTree:
    """Paginated table content.

    Returned by get_table() for large table handling.

    Attributes:
        table_ref: Reference to the table
        total_rows: Total number of rows in table
        total_cols: Total number of columns
        rows: List of row AccessibilityNodes (may be subset)
        page: Current page number (1-indexed)
        page_size: Rows per page
        has_more: Whether more rows exist
    """

    table_ref: Ref
    total_rows: int
    total_cols: int
    rows: list[AccessibilityNode]
    page: int = 1
    page_size: int = 20
    has_more: bool = False

    def to_yaml(self) -> str:
        """Serialize table to YAML with pagination info.

        Returns:
            YAML string representation
        """
        buffer = StringIO()
        buffer.write(f"table [ref={self.table_ref}]:\n")
        buffer.write(f"  total_rows: {self.total_rows}\n")
        buffer.write(f"  total_cols: {self.total_cols}\n")
        buffer.write(f"  page: {self.page}\n")
        buffer.write(f"  page_size: {self.page_size}\n")
        buffer.write(f"  has_more: {str(self.has_more).lower()}\n")
        buffer.write("  rows:\n")

        for row in self.rows:
            header_str = " [header]" if row.properties.get("header") == "true" else ""
            buffer.write(f"    - row [ref={row.ref}]{header_str}:\n")
            for cell in row.children:
                text = cell.text[:40] + "..." if len(cell.text) > 40 else cell.text
                text = text.replace('"', '\\"')
                buffer.write(f'        - cell: "{text}"\n')

        return buffer.getvalue()


# ============================================================================
# Search Results
# ============================================================================


@dataclass
class SearchResult:
    """A single search result with context.

    Attributes:
        ref: Reference to the containing element
        text: Matched text
        context: Surrounding context text
        section_ref: Reference to containing section (if any)
    """

    ref: Ref
    text: str
    context: str
    section_ref: Ref | None = None


@dataclass
class SearchResults:
    """Results from a document search.

    Attributes:
        query: The search query
        results: List of search results
        total_matches: Total number of matches found
        truncated: Whether results were truncated
    """

    query: str
    results: list[SearchResult]
    total_matches: int = 0
    truncated: bool = False

    def to_yaml(self) -> str:
        """Serialize search results to YAML.

        Returns:
            YAML string representation
        """
        buffer = StringIO()
        buffer.write("search_results:\n")
        buffer.write(f'  query: "{self.query}"\n')
        buffer.write(f"  total_matches: {self.total_matches}\n")
        buffer.write(f"  truncated: {str(self.truncated).lower()}\n")
        buffer.write("  results:\n")

        for result in self.results:
            buffer.write(f"    - ref: {result.ref}\n")
            context = result.context.replace('"', '\\"').replace("\n", " ")
            buffer.write(f'      context: "{context}"\n')
            if result.section_ref:
                buffer.write(f"      section: {result.section_ref}\n")

        return buffer.getvalue()


# ============================================================================
# Internal Builders
# ============================================================================


class _OutlineBuilder:
    """Internal builder for constructing outline from document XML."""

    def __init__(
        self,
        xml_root: etree._Element,
        config: SectionDetectionConfig,
    ) -> None:
        self.xml_root = xml_root
        self.config = config
        self._section_index = 0
        self._paragraph_index = 0

        # Compile numbering patterns
        self._compiled_patterns = [re.compile(pattern) for pattern in config.numbering_patterns]

    def build(self) -> list[SectionInfo]:
        """Build sections from document.

        Returns:
            List of top-level SectionInfo objects
        """
        body = self.xml_root.find(f".//{w('body')}")
        if body is None:
            return []

        sections: list[SectionInfo] = []
        current_section: SectionInfo | None = None
        pending_content: list[etree._Element] = []

        for child in body:
            if child.tag == w("p"):
                heading_info = self._detect_heading(child)

                if heading_info:
                    # Close previous section
                    if current_section:
                        self._finalize_section(current_section, pending_content)
                        sections.append(current_section)

                    # Start new section
                    current_section = self._create_section(child, heading_info)
                    pending_content = []
                else:
                    pending_content.append(child)

                self._paragraph_index += 1

            elif child.tag == w("tbl"):
                pending_content.append(child)

        # Finalize last section
        if current_section:
            self._finalize_section(current_section, pending_content)
            sections.append(current_section)
        elif pending_content:
            # Document has no sections - create implicit section
            implicit_section = SectionInfo(
                ref=Ref(path="sec:0"),
                heading="(Document Content)",
                heading_ref=Ref(path="p:0"),
                heading_level=1,
                detection_method="implicit",
                confidence="low",
            )
            self._finalize_section(implicit_section, pending_content)
            sections.append(implicit_section)

        return sections

    def _detect_heading(self, p_elem: etree._Element) -> dict[str, str | int] | None:
        """Detect if a paragraph is a heading.

        Args:
            p_elem: Paragraph element

        Returns:
            Dict with heading info or None if not a heading
        """
        text = self._extract_text(p_elem)

        # Skip empty paragraphs
        if not text.strip():
            return None

        # Skip paragraphs that are too long to be headings
        if len(text) > self.config.max_heading_length:
            return None

        # Tier 1: Check for heading style
        if self.config.detect_heading_styles:
            style = self._get_paragraph_style(p_elem)
            if style:
                level = self._get_heading_level_from_style(style)
                if level:
                    return {
                        "level": level,
                        "method": "heading_style",
                        "confidence": "high",
                    }

        # Tier 1: Check for outline level property
        if self.config.detect_outline_level:
            outline_level = self._get_outline_level(p_elem)
            if outline_level is not None:
                return {
                    "level": outline_level + 1,  # 0-indexed to 1-indexed
                    "method": "outline_level",
                    "confidence": "high",
                }

        # Tier 3: Heuristic detection
        # Check for numbered sections
        if self.config.detect_numbered_sections:
            for pattern in self._compiled_patterns:
                if pattern.match(text):
                    return {
                        "level": 1,
                        "method": "numbered_pattern",
                        "confidence": "medium",
                    }

        # Check for all-bold short paragraph
        if self.config.detect_bold_first_line:
            if self._is_all_bold(p_elem) and len(text) <= 80:
                return {
                    "level": 1,
                    "method": "bold_heuristic",
                    "confidence": "medium",
                }

        # Check for ALL CAPS
        if self.config.detect_caps_first_line:
            if text.isupper() and len(text) <= 80:
                return {
                    "level": 1,
                    "method": "caps_heuristic",
                    "confidence": "medium",
                }

        return None

    def _create_section(
        self, p_elem: etree._Element, heading_info: dict[str, str | int]
    ) -> SectionInfo:
        """Create a SectionInfo from a heading paragraph.

        Args:
            p_elem: Heading paragraph element
            heading_info: Detected heading information

        Returns:
            New SectionInfo object
        """
        text = self._extract_text(p_elem)
        sec_ref = Ref(path=f"sec:{self._section_index}")
        self._section_index += 1

        return SectionInfo(
            ref=sec_ref,
            heading=text,
            heading_ref=Ref(path=f"p:{self._paragraph_index}"),
            heading_level=int(heading_info.get("level", 1)),
            detection_method=str(heading_info.get("method", "unknown")),
            confidence=str(heading_info.get("confidence", "low")),
        )

    def _finalize_section(self, section: SectionInfo, content: list[etree._Element]) -> None:
        """Finalize section with content counts and preview.

        Args:
            section: Section to finalize
            content: List of content elements in section
        """
        paragraph_count = 0
        table_count = 0
        tracked_changes = 0
        preview_text = ""

        for elem in content:
            if elem.tag == w("p"):
                paragraph_count += 1
                if not preview_text:
                    preview_text = self._extract_text(elem)[:100]
                # Count tracked changes in paragraph
                tracked_changes += len(list(elem.iter(w("ins"))))
                tracked_changes += len(list(elem.iter(w("del"))))
            elif elem.tag == w("tbl"):
                table_count += 1

        section.paragraph_count = paragraph_count
        section.table_count = table_count
        section.tracked_change_count = tracked_changes
        section.preview = preview_text

    def _extract_text(self, elem: etree._Element) -> str:
        """Extract text from element."""
        text_parts = []
        for t_elem in elem.iter(w("t")):
            if t_elem.text:
                text_parts.append(t_elem.text)
        return "".join(text_parts)

    def _get_paragraph_style(self, p_elem: etree._Element) -> str | None:
        """Get style name for paragraph."""
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

        return None

    def _get_outline_level(self, p_elem: etree._Element) -> int | None:
        """Get outline level from paragraph properties."""
        para_props = p_elem.find(f"./{w('pPr')}")
        if para_props is not None:
            outline_lvl = para_props.find(f"./{w('outlineLvl')}")
            if outline_lvl is not None:
                val = outline_lvl.get(f"{{{WORD_NAMESPACE}}}val")
                if val and val.isdigit():
                    return int(val)
        return None

    def _is_all_bold(self, p_elem: etree._Element) -> bool:
        """Check if all text in paragraph is bold."""
        runs = list(p_elem.findall(f".//{w('r')}"))
        if not runs:
            return False

        for run in runs:
            # Check if run has any text
            has_text = any(t.text for t in run.findall(f".//{w('t')}"))
            if not has_text:
                continue

            # Check for bold property
            run_props = run.find(f"./{w('rPr')}")
            if run_props is None:
                return False
            if run_props.find(f"./{w('b')}") is None:
                return False

        return True


class _OutlineYamlWriter:
    """Internal writer for outline YAML output."""

    def __init__(self, outline: OutlineTree) -> None:
        self.outline = outline
        self.buffer = StringIO()
        self._indent = 0

    def write(self) -> str:
        """Write outline to YAML."""
        self._write_header()
        self._write_sections()
        self._write_navigation()
        return self.buffer.getvalue()

    def _write_header(self) -> None:
        """Write document header."""
        self._write_line("document:")
        self._indent += 1

        if self.outline.document_path:
            self._write_line(f'path: "{self.outline.document_path}"')

        self._write_line("mode: outline")

        # Stats
        self._write_line("stats:")
        self._indent += 1
        self._write_line(f"total_paragraphs: {self.outline.size_info.paragraph_count}")
        self._write_line(f"total_sections: {len(self.outline.sections)}")
        self._write_line(f"total_tables: {self.outline.size_info.table_count}")
        self._write_line(f"tracked_changes: {self.outline.size_info.tracked_change_count}")
        self._write_line(f"degradation_tier: {self.outline.size_info.degradation_tier}")
        self._indent -= 1

        self._indent -= 1
        self._write_line("")

    def _write_sections(self) -> None:
        """Write section outline."""
        self._write_line("outline:")
        self._indent += 1

        for section in self.outline.sections:
            self._write_section(section)

        self._indent -= 1

    def _write_section(self, section: SectionInfo, level: int = 0) -> None:
        """Write a single section."""
        # Determine confidence marker
        confidence_marker = "[inferred]" if section.confidence != "high" else "[explicit]"

        self._write_line(f"- section [ref={section.ref}] {confidence_marker}:")
        self._indent += 1

        self._write_line(f"detection: {section.detection_method}")
        self._write_line(f"confidence: {section.confidence}")

        # Escape heading text
        heading = section.heading.replace('"', '\\"')
        self._write_line(f'heading: "{heading}"')
        self._write_line(f"heading_ref: {section.heading_ref}")
        self._write_line(f"heading_level: {section.heading_level}")
        self._write_line(f"paragraph_count: {section.paragraph_count}")

        if section.table_count > 0:
            self._write_line(f"table_count: {section.table_count}")

        if section.tracked_change_count > 0:
            self._write_line(f"tracked_changes: {section.tracked_change_count}")

        if section.preview:
            preview = section.preview.replace('"', '\\"').replace("\n", " ")
            self._write_line(f'preview: "{preview}"')

        # Write children (subsections)
        if section.children:
            self._write_line("subsections:")
            self._indent += 1
            for child in section.children:
                self._write_section(child, level + 1)
            self._indent -= 1

        self._indent -= 1

    def _write_navigation(self) -> None:
        """Write navigation hints."""
        self._write_line("")
        self._write_line("navigation:")
        self._indent += 1
        self._write_line("expand_section: \"doc.expand_section('sec:N')\"")
        self._write_line("expand_refs: \"doc.expand_refs(['p:N', ...])\"")
        self._write_line("search: \"doc.search('pattern')\"")
        self._write_line("get_table: \"doc.get_table('tbl:N', max_rows=20)\"")
        self._indent -= 1

    def _write_line(self, text: str) -> None:
        """Write a line with current indentation."""
        indent = "  " * self._indent
        self.buffer.write(f"{indent}{text}\n")


# ============================================================================
# Token Budgeting
# ============================================================================


def estimate_tokens(text: str) -> int:
    """Estimate token count for text.

    Uses a simple heuristic: ~4 characters per token on average.

    Args:
        text: Text to estimate

    Returns:
        Estimated token count
    """
    return len(text) // 4


def truncate_to_token_budget(
    nodes: list[AccessibilityNode],
    max_tokens: int,
    prioritize_structure: bool = True,
) -> list[AccessibilityNode]:
    """Truncate a list of nodes to fit within token budget.

    Args:
        nodes: List of nodes to potentially truncate
        max_tokens: Maximum tokens to allow
        prioritize_structure: If True, keep headings even when truncating

    Returns:
        Truncated list of nodes
    """
    if max_tokens <= 0:
        return nodes

    result: list[AccessibilityNode] = []
    current_tokens = 0

    for node in nodes:
        # Estimate tokens for this node
        node_tokens = estimate_tokens(node.text) + 20  # Add overhead for YAML structure

        # Always include headings if prioritizing structure
        if prioritize_structure and node.level is not None:
            result.append(node)
            current_tokens += node_tokens
            continue

        # Check if we have budget
        if current_tokens + node_tokens <= max_tokens:
            result.append(node)
            current_tokens += node_tokens
        else:
            # Budget exceeded - stop adding
            break

    return result
