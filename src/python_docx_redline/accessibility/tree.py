"""
AccessibilityTree for DOCX documents.

This module provides the AccessibilityTree class that builds a semantic tree
representation of Word documents with YAML serialization support.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from io import StringIO
from pathlib import Path
from typing import TYPE_CHECKING, Literal

from lxml import etree

from ..constants import WORD_NAMESPACE, w
from .bookmarks import BookmarkRegistry, CrossReferenceRegistry
from .images import ImageExtractor
from .registry import RefRegistry
from .types import (
    AccessibilityNode,
    BookmarkInfo,
    CrossReferenceInfo,
    ElementType,
    HyperlinkInfo,
    ImageInfo,
    Ref,
    ReferenceValidationResult,
    ViewMode,
)

if TYPE_CHECKING:
    from ..document import Document


@dataclass
class DocumentStats:
    """Statistics about a document.

    Attributes:
        paragraphs: Number of paragraphs in the body
        tables: Number of tables in the body
        tracked_changes: Number of tracked changes
        comments: Number of comments
        images: Number of images and embedded objects
        bookmarks: Number of bookmarks
        hyperlinks: Number of hyperlinks
        cross_references: Number of cross-references (REF/PAGEREF/NOTEREF fields)
    """

    paragraphs: int = 0
    tables: int = 0
    tracked_changes: int = 0
    comments: int = 0
    images: int = 0
    bookmarks: int = 0
    hyperlinks: int = 0
    cross_references: int = 0


class AccessibilityTree:
    """A semantic accessibility tree for Word documents.

    AccessibilityTree provides a structured view of document content with
    stable refs for elements, tracked changes, and comments. It supports
    YAML serialization with three verbosity levels: minimal, standard, and full.

    Attributes:
        root: Root node of the accessibility tree
        registry: RefRegistry for resolving refs to elements
        view_mode: Configuration for what to include
        stats: Document statistics

    Example:
        >>> from python_docx_redline import Document
        >>> doc = Document("contract.docx")
        >>> tree = AccessibilityTree.from_document(doc)
        >>> print(tree.to_yaml())
    """

    def __init__(
        self,
        root: AccessibilityNode,
        registry: RefRegistry,
        view_mode: ViewMode | None = None,
        stats: DocumentStats | None = None,
        document_path: str | Path | None = None,
        bookmark_registry: BookmarkRegistry | None = None,
        cross_reference_registry: CrossReferenceRegistry | None = None,
    ) -> None:
        """Initialize an AccessibilityTree.

        Args:
            root: Root node of the tree
            registry: RefRegistry for ref resolution
            view_mode: View configuration (defaults to standard)
            stats: Document statistics
            document_path: Path to source document
            bookmark_registry: BookmarkRegistry for bookmarks and hyperlinks
            cross_reference_registry: CrossReferenceRegistry for cross-references
        """
        self.root = root
        self.registry = registry
        self.view_mode = view_mode or ViewMode()
        self.stats = stats or DocumentStats()
        self.document_path = document_path
        self.bookmark_registry = bookmark_registry or BookmarkRegistry()
        self.cross_reference_registry = cross_reference_registry or CrossReferenceRegistry()

    @property
    def bookmarks(self) -> dict[str, BookmarkInfo]:
        """Get all bookmarks in the document.

        Returns:
            Dictionary of bookmark name to BookmarkInfo
        """
        return self.bookmark_registry.bookmarks

    @property
    def hyperlinks(self) -> list[HyperlinkInfo]:
        """Get all hyperlinks in the document.

        Returns:
            List of HyperlinkInfo objects
        """
        return self.bookmark_registry.hyperlinks

    @property
    def cross_references(self) -> list[CrossReferenceInfo]:
        """Get all cross-references in the document.

        Returns:
            List of CrossReferenceInfo objects
        """
        return self.cross_reference_registry.cross_references

    def get_cross_references_to(self, bookmark_name: str) -> list[CrossReferenceInfo]:
        """Get all cross-references that target a specific bookmark.

        Args:
            bookmark_name: Name of the target bookmark

        Returns:
            List of CrossReferenceInfo objects
        """
        return self.cross_reference_registry.get_by_target(bookmark_name)

    def get_bookmark(self, name: str) -> BookmarkInfo | None:
        """Get a bookmark by name.

        Args:
            name: Bookmark name

        Returns:
            BookmarkInfo or None if not found
        """
        return self.bookmark_registry.get_bookmark(name)

    def validate_references(self) -> ReferenceValidationResult:
        """Validate all document references.

        Checks for broken links, broken cross-references, and orphan bookmarks.

        Returns:
            ReferenceValidationResult with validation details
        """
        result = self.bookmark_registry.validate_references()

        # Add broken cross-references
        broken_xrefs = self.cross_reference_registry.get_broken()
        if broken_xrefs:
            result.broken_cross_references = broken_xrefs
            result.is_valid = False
            result.warnings.append(f"Found {len(broken_xrefs)} broken cross-reference(s)")

            # Add to missing_bookmarks
            for xref in broken_xrefs:
                if xref.target_bookmark not in result.missing_bookmarks:
                    result.missing_bookmarks.append(xref.target_bookmark)

        return result

    def get_images(self) -> list[ImageInfo]:
        """Get all images in the document.

        Returns:
            List of ImageInfo objects for all images in the document
        """
        images: list[ImageInfo] = []
        for node in self.iter_nodes():
            if node.has_images:
                images.extend(node.images)
        return images

    def get_image(self, ref: str) -> ImageInfo | None:
        """Get an image by its ref.

        Args:
            ref: Image ref string (e.g., "img:5/0")

        Returns:
            ImageInfo or None if not found
        """
        for node in self.iter_nodes():
            for image in node.images:
                if image.ref == ref:
                    return image
        return None

    @classmethod
    def from_document(
        cls,
        document: Document,
        view_mode: ViewMode | None = None,
    ) -> AccessibilityTree:
        """Build an accessibility tree from a Document.

        Args:
            document: The python_docx_redline Document
            view_mode: Configuration for what to include

        Returns:
            AccessibilityTree representing the document
        """
        view_mode = view_mode or ViewMode()
        xml_root = document.xml_root
        registry = RefRegistry(xml_root)

        # Build the tree
        builder = _TreeBuilder(xml_root, registry, view_mode)
        root = builder.build()
        stats = builder.stats

        # Extract bookmarks and hyperlinks
        bookmark_registry = BookmarkRegistry.from_xml(xml_root)
        stats.bookmarks = len(bookmark_registry.bookmarks)
        stats.hyperlinks = len(bookmark_registry.hyperlinks)

        # Extract cross-references (pass bookmark_registry for reference resolution)
        cross_reference_registry = CrossReferenceRegistry.from_xml(xml_root, bookmark_registry)
        stats.cross_references = len(cross_reference_registry.cross_references)

        return cls(
            root=root,
            registry=registry,
            view_mode=view_mode,
            stats=stats,
            document_path=getattr(document, "path", None),
            bookmark_registry=bookmark_registry,
            cross_reference_registry=cross_reference_registry,
        )

    @classmethod
    def from_xml(
        cls,
        xml_root: etree._Element,
        view_mode: ViewMode | None = None,
        document_path: str | Path | None = None,
    ) -> AccessibilityTree:
        """Build an accessibility tree from an lxml element.

        Args:
            xml_root: Root element of the document XML
            view_mode: Configuration for what to include
            document_path: Optional path for display

        Returns:
            AccessibilityTree representing the document
        """
        view_mode = view_mode or ViewMode()
        registry = RefRegistry(xml_root)

        builder = _TreeBuilder(xml_root, registry, view_mode)
        root = builder.build()
        stats = builder.stats

        # Extract bookmarks and hyperlinks
        bookmark_registry = BookmarkRegistry.from_xml(xml_root)
        stats.bookmarks = len(bookmark_registry.bookmarks)
        stats.hyperlinks = len(bookmark_registry.hyperlinks)

        # Extract cross-references (pass bookmark_registry for reference resolution)
        cross_reference_registry = CrossReferenceRegistry.from_xml(xml_root, bookmark_registry)
        stats.cross_references = len(cross_reference_registry.cross_references)

        return cls(
            root=root,
            registry=registry,
            view_mode=view_mode,
            stats=stats,
            document_path=document_path,
            bookmark_registry=bookmark_registry,
            cross_reference_registry=cross_reference_registry,
        )

    def find_by_ref(self, ref: str | Ref) -> AccessibilityNode | None:
        """Find a node by its ref.

        Args:
            ref: Ref to search for

        Returns:
            The node or None if not found
        """
        return self.root.find_by_ref(ref)

    def find_all(
        self,
        element_type: ElementType | None = None,
        heading_level: int | None = None,
        has_changes: bool | None = None,
        text_contains: str | None = None,
    ) -> list[AccessibilityNode]:
        """Find all nodes matching criteria.

        Args:
            element_type: Filter by element type
            heading_level: Filter by heading level
            has_changes: Filter by whether node has tracked changes
            text_contains: Filter by text content substring

        Returns:
            List of matching nodes
        """
        results: list[AccessibilityNode] = []

        def _matches(node: AccessibilityNode) -> bool:
            if element_type is not None and node.element_type != element_type:
                return False
            if heading_level is not None and node.level != heading_level:
                return False
            if has_changes is not None and node.has_changes != has_changes:
                return False
            if text_contains is not None and text_contains not in node.text:
                return False
            return True

        def _walk(node: AccessibilityNode) -> None:
            if _matches(node):
                results.append(node)
            for child in node.children:
                _walk(child)

        _walk(self.root)
        return results

    def iter_nodes(self) -> Iterator[AccessibilityNode]:
        """Iterate over all nodes in the tree.

        Yields:
            AccessibilityNode in document order
        """

        def _walk(node: AccessibilityNode) -> Iterator[AccessibilityNode]:
            yield node
            for child in node.children:
                yield from _walk(child)

        yield from _walk(self.root)

    def to_yaml(self, verbosity: str | None = None) -> str:
        """Serialize the tree to YAML format.

        Args:
            verbosity: Override verbosity level ("minimal", "standard", "full")
                       Uses view_mode.verbosity if not specified.

        Returns:
            YAML string representation

        Example output (standard):
            document:
              path: "contract.docx"
              verbosity: standard
              stats:
                paragraphs: 47
                tables: 2
                tracked_changes: 12

            content:
              - heading [ref=p:0] [level=1]:
                  text: "SERVICES AGREEMENT"
              - paragraph [ref=p:1]:
                  text: "This Agreement..."
        """
        verbosity = verbosity or self.view_mode.verbosity
        writer = _YamlWriter(verbosity, self.view_mode)
        return writer.write(self)

    # =========================================================================
    # Large Document Handling Methods
    # =========================================================================

    def expand_section(
        self,
        section_ref: str,
        mode: Literal["content", "styling"] = "content",
    ) -> SectionTree:
        """Expand a single section to full content.

        Args:
            section_ref: Section reference (e.g., "sec:2")
            mode: Content mode ("content" for text, "styling" for run-level)

        Returns:
            SectionTree with expanded section content
        """
        from ..errors import RefNotFoundError
        from .outline import SectionTree

        if not section_ref.startswith("sec:"):
            raise RefNotFoundError(section_ref, "Expected section ref (sec:N)")

        try:
            section_index = int(section_ref.split(":")[1])
        except (ValueError, IndexError):
            raise RefNotFoundError(section_ref, "Invalid section ref format")

        body = self.registry.xml_root.find(f".//{w('body')}")
        if body is None:
            raise RefNotFoundError(section_ref, "Document body not found")

        sections = self._find_section_boundaries(body)
        if section_index >= len(sections):
            raise RefNotFoundError(
                section_ref,
                f"Section index {section_index} out of bounds (found {len(sections)} sections)",
            )

        start_idx, end_idx, heading = sections[section_index]
        content_nodes: list[AccessibilityNode] = []
        view = ViewMode(
            verbosity="full" if mode == "styling" else "standard",
            include_formatting=(mode == "styling"),
        )

        paragraph_count = 0
        table_count = 0
        current_p_idx = 0
        current_tbl_idx = 0

        for child in body:
            if child.tag == w("p"):
                if start_idx <= current_p_idx < end_idx:
                    node = self._build_node_for_element(child, f"p:{current_p_idx}", view)
                    if node:
                        content_nodes.append(node)
                        paragraph_count += 1
                current_p_idx += 1
            elif child.tag == w("tbl"):
                if current_p_idx >= start_idx and current_p_idx < end_idx:
                    node = self._build_table_node_static(child, f"tbl:{current_tbl_idx}")
                    if node:
                        content_nodes.append(node)
                        table_count += 1
                current_tbl_idx += 1

        return SectionTree(
            section_ref=Ref(path=section_ref),
            heading=heading,
            content=content_nodes,
            paragraph_count=paragraph_count,
            table_count=table_count,
        )

    def expand_refs(
        self,
        refs: list[str],
        mode: Literal["content", "styling"] = "content",
    ) -> RefTree:
        """Expand specific refs to full content.

        Args:
            refs: List of refs to expand
            mode: Content mode
        """
        from ..errors import RefNotFoundError
        from .outline import RefTree

        view = ViewMode(
            verbosity="full" if mode == "styling" else "standard",
            include_formatting=(mode == "styling"),
        )

        nodes: dict[str, AccessibilityNode] = {}
        not_found: list[str] = []

        for ref_str in refs:
            try:
                element = self.registry.resolve_ref(ref_str)
                node = self._build_node_for_element(element, ref_str, view)
                if node:
                    nodes[ref_str] = node
                else:
                    not_found.append(ref_str)
            except RefNotFoundError:
                not_found.append(ref_str)

        return RefTree(refs=refs, nodes=nodes, not_found=not_found)

    def get_table(
        self,
        table_ref: str,
        max_rows: int | None = None,
        page: int = 1,
    ) -> TableTree:
        """Get table content with optional pagination."""
        from ..errors import RefNotFoundError
        from .outline import TableTree

        try:
            tbl_elem = self.registry.resolve_ref(table_ref)
        except Exception as e:
            raise RefNotFoundError(table_ref, str(e))

        if tbl_elem.tag != w("tbl"):
            raise RefNotFoundError(table_ref, "Not a table")

        all_rows = list(tbl_elem.findall(f"./{w('tr')}"))
        total_rows = len(all_rows)
        total_cols = 0

        if all_rows:
            first_row = all_rows[0]
            total_cols = len(list(first_row.findall(f"./{w('tc')}")))

        if max_rows is None:
            start_idx = 0
            end_idx = total_rows
            has_more = False
        else:
            start_idx = (page - 1) * max_rows
            end_idx = min(start_idx + max_rows, total_rows)
            has_more = end_idx < total_rows

        row_nodes: list[AccessibilityNode] = []
        for i, tr_elem in enumerate(all_rows[start_idx:end_idx]):
            row_idx = start_idx + i
            row_ref = Ref(path=f"{table_ref}/row:{row_idx}")
            row_node = self._build_table_row_node_static(tr_elem, row_ref, row_idx == 0)
            row_nodes.append(row_node)

        return TableTree(
            table_ref=Ref(path=table_ref),
            total_rows=total_rows,
            total_cols=total_cols,
            rows=row_nodes,
            page=page,
            page_size=max_rows or total_rows,
            has_more=has_more,
        )

    def search(
        self,
        pattern: str,
        max_results: int = 20,
        case_sensitive: bool = False,
    ) -> SearchResults:
        """Search document for text patterns."""
        import re as re_module

        from .outline import SearchResult, SearchResults

        flags = 0 if case_sensitive else re_module.IGNORECASE
        try:
            regex = re_module.compile(pattern, flags)
        except re_module.error:
            regex = re_module.compile(re_module.escape(pattern), flags)

        results: list[SearchResult] = []
        total_matches = 0

        for node in self.iter_nodes():
            if node.element_type != ElementType.PARAGRAPH:
                continue

            matches = list(regex.finditer(node.text))
            total_matches += len(matches)

            for match in matches:
                if len(results) >= max_results:
                    break

                start = max(0, match.start() - 30)
                end = min(len(node.text), match.end() + 30)
                context = node.text[start:end]
                if start > 0:
                    context = "..." + context
                if end < len(node.text):
                    context = context + "..."

                results.append(
                    SearchResult(
                        ref=node.ref,
                        text=match.group(),
                        context=context,
                    )
                )

            if len(results) >= max_results:
                break

        return SearchResults(
            query=pattern,
            results=results,
            total_matches=total_matches,
            truncated=(total_matches > max_results),
        )

    def _find_section_boundaries(self, body: etree._Element) -> list[tuple[int, int, str]]:
        """Find section boundaries based on headings."""
        sections: list[tuple[int, int, str]] = []
        heading_indices: list[tuple[int, str]] = []

        para_idx = 0
        for child in body:
            if child.tag == w("p"):
                style = self._get_paragraph_style_static(child)
                if style and self._is_heading_style(style):
                    text = self._extract_text_static(child)
                    heading_indices.append((para_idx, text))
                para_idx += 1

        for i, (start_idx, heading) in enumerate(heading_indices):
            if i + 1 < len(heading_indices):
                end_idx = heading_indices[i + 1][0]
            else:
                end_idx = para_idx

            sections.append((start_idx, end_idx, heading))

        if not sections and para_idx > 0:
            sections.append((0, para_idx, "(Document Content)"))

        return sections

    def _is_heading_style(self, style: str) -> bool:
        """Check if a style name indicates a heading."""
        style_lower = style.lower()
        return style_lower.startswith("heading") or style_lower == "title"

    def _build_node_for_element(
        self,
        element: etree._Element,
        ref_str: str,
        view_mode: ViewMode,
    ) -> AccessibilityNode | None:
        """Build a node for a specific element."""
        if element.tag == w("p"):
            return self._build_paragraph_node_static(element, ref_str, view_mode)
        elif element.tag == w("tbl"):
            return self._build_table_node_static(element, ref_str)
        return None

    def _build_paragraph_node_static(
        self,
        p_elem: etree._Element,
        ref_str: str,
        view_mode: ViewMode,
    ) -> AccessibilityNode:
        """Build node for a paragraph."""
        ref = Ref(path=ref_str)
        text = self._extract_text_static(p_elem)
        style = self._get_paragraph_style_static(p_elem)
        level = self._get_heading_level_static(style)

        node = AccessibilityNode(
            ref=ref,
            element_type=ElementType.PARAGRAPH,
            text=text,
            style=style,
            level=level,
            _element=p_elem,
        )

        if view_mode.include_formatting or view_mode.verbosity == "full":
            node.children = self._build_run_nodes_static(p_elem, ref)

        return node

    def _build_table_node_static(
        self,
        tbl_elem: etree._Element,
        ref_str: str,
    ) -> AccessibilityNode:
        """Build node for a table."""
        ref = Ref(path=ref_str)
        rows: list[AccessibilityNode] = []
        row_index = 0

        for tr_elem in tbl_elem.findall(f"./{w('tr')}"):
            row_ref = Ref(path=f"{ref_str}/row:{row_index}")
            row_node = self._build_table_row_node_static(tr_elem, row_ref, row_index == 0)
            rows.append(row_node)
            row_index += 1

        num_rows = len(rows)
        num_cols = len(rows[0].children) if rows else 0

        return AccessibilityNode(
            ref=ref,
            element_type=ElementType.TABLE,
            children=rows,
            properties={"rows": str(num_rows), "cols": str(num_cols)},
            _element=tbl_elem,
        )

    def _build_table_row_node_static(
        self,
        tr_elem: etree._Element,
        ref: Ref,
        is_header: bool = False,
    ) -> AccessibilityNode:
        """Build node for a table row."""
        cells: list[AccessibilityNode] = []
        cell_index = 0

        for tc_elem in tr_elem.findall(f"./{w('tc')}"):
            cell_ref = Ref(path=f"{ref.path}/cell:{cell_index}")
            cell_node = self._build_table_cell_node_static(tc_elem, cell_ref)
            cells.append(cell_node)
            cell_index += 1

        properties: dict[str, str] = {}
        if is_header:
            properties["header"] = "true"

        return AccessibilityNode(
            ref=ref,
            element_type=ElementType.TABLE_ROW,
            children=cells,
            properties=properties,
            _element=tr_elem,
        )

    def _build_table_cell_node_static(
        self,
        tc_elem: etree._Element,
        ref: Ref,
    ) -> AccessibilityNode:
        """Build node for a table cell."""
        text_parts = []
        for p_elem in tc_elem.findall(f"./{w('p')}"):
            text_parts.append(self._extract_text_static(p_elem))
        text = " ".join(text_parts).strip()

        return AccessibilityNode(
            ref=ref,
            element_type=ElementType.TABLE_CELL,
            text=text,
            _element=tc_elem,
        )

    def _build_run_nodes_static(
        self,
        p_elem: etree._Element,
        parent_ref: Ref,
    ) -> list[AccessibilityNode]:
        """Build nodes for runs within a paragraph."""
        runs: list[AccessibilityNode] = []
        run_index = 0

        for r_elem in p_elem.iter(w("r")):
            text = self._extract_run_text_static(r_elem)
            if not text:
                continue

            ref = Ref(path=f"{parent_ref.path}/r:{run_index}")
            run_index += 1

            properties: dict[str, str] = {}
            run_props = r_elem.find(w("rPr"))
            if run_props is not None:
                if run_props.find(w("b")) is not None:
                    properties["bold"] = "true"
                if run_props.find(w("i")) is not None:
                    properties["italic"] = "true"
                if run_props.find(w("u")) is not None:
                    properties["underline"] = "true"

            runs.append(
                AccessibilityNode(
                    ref=ref,
                    element_type=ElementType.RUN,
                    text=text,
                    properties=properties,
                    _element=r_elem,
                )
            )

        return runs

    @staticmethod
    def _extract_text_static(elem: etree._Element) -> str:
        """Extract all text from an element."""
        text_parts = []
        for t_elem in elem.iter(w("t")):
            if t_elem.text:
                text_parts.append(t_elem.text)
        return "".join(text_parts)

    @staticmethod
    def _extract_run_text_static(r_elem: etree._Element) -> str:
        """Extract text from a single run."""
        text_parts = []
        for t_elem in r_elem.findall(f"./{w('t')}"):
            if t_elem.text:
                text_parts.append(t_elem.text)
        return "".join(text_parts)

    @staticmethod
    def _get_paragraph_style_static(p_elem: etree._Element) -> str | None:
        """Get the style name for a paragraph."""
        para_props = p_elem.find(f"./{w('pPr')}")
        if para_props is not None:
            para_style = para_props.find(f"./{w('pStyle')}")
            if para_style is not None:
                return para_style.get(f"{{{WORD_NAMESPACE}}}val")
        return None

    @staticmethod
    def _get_heading_level_static(style: str | None) -> int | None:
        """Determine heading level from style name."""
        if not style:
            return None

        style_lower = style.lower()
        if style_lower.startswith("heading"):
            suffix = style_lower[7:].strip()
            if suffix.isdigit():
                return int(suffix)

        if style_lower == "title":
            return 1

        return None

    def __repr__(self) -> str:
        return (
            f"AccessibilityTree(paragraphs={self.stats.paragraphs}, "
            f"tables={self.stats.tables}, "
            f"tracked_changes={self.stats.tracked_changes})"
        )


# Forward type references
SectionTree = "SectionTree"
RefTree = "RefTree"
TableTree = "TableTree"
SearchResults = "SearchResults"


class _TreeBuilder:
    """Internal class for building the accessibility tree."""

    def __init__(
        self,
        xml_root: etree._Element,
        registry: RefRegistry,
        view_mode: ViewMode,
    ) -> None:
        self.xml_root = xml_root
        self.registry = registry
        self.view_mode = view_mode
        self.stats = DocumentStats()

        # Track indices for building refs
        self._paragraph_index = 0
        self._table_index = 0

        # Image extractor
        self._image_extractor = ImageExtractor(xml_root)

    def build(self) -> AccessibilityNode:
        """Build the tree and return the root node."""
        # Create root document node
        root = AccessibilityNode(
            ref=Ref(path="doc:0"),
            element_type=ElementType.DOCUMENT,
        )

        # Build body content
        if self.view_mode.include_body:
            body = self.xml_root.find(f".//{w('body')}")
            if body is not None:
                root.children = self._build_body(body)

        return root

    def _build_body(self, body: etree._Element) -> list[AccessibilityNode]:
        """Build nodes for the document body."""
        nodes: list[AccessibilityNode] = []

        for child in body:
            if child.tag == w("p"):
                node = self._build_paragraph(child)
                if node:
                    nodes.append(node)
            elif child.tag == w("tbl"):
                node = self._build_table(child)
                if node:
                    nodes.append(node)

        return nodes

    def _build_paragraph(self, p_elem: etree._Element) -> AccessibilityNode | None:
        """Build a node for a paragraph."""
        idx = self._paragraph_index
        self._paragraph_index += 1
        self.stats.paragraphs += 1

        ref = Ref(path=f"p:{idx}")

        # Extract text content
        text = self._extract_text(p_elem)

        # Get style
        style = self._get_paragraph_style(p_elem)

        # Check for heading level
        level = self._get_heading_level(style)

        # Extract tracked changes
        changes = self._extract_tracked_changes(p_elem)
        if changes:
            self.stats.tracked_changes += len(changes)

        # Extract images
        images = self._image_extractor.extract_from_paragraph(p_elem, idx)
        if images:
            self.stats.images += len(images)

        # Create the node
        node = AccessibilityNode(
            ref=ref,
            element_type=ElementType.PARAGRAPH,
            text=text,
            style=style,
            level=level,
            images=images,
            _element=p_elem,
        )

        # Add changes info to properties if present
        if changes and self.view_mode.include_tracked_changes:
            node.properties["has_changes"] = "true"
            node.properties["change_count"] = str(len(changes))

            # Store change details for YAML output
            node.properties["_changes"] = changes  # type: ignore[assignment]

        # Add images info to properties if present
        if images:
            node.properties["has_images"] = "true"
            node.properties["image_count"] = str(len(images))

        # Build runs if full verbosity
        if self.view_mode.verbosity == "full" or self.view_mode.include_formatting:
            node.children = self._build_runs(p_elem, ref)

        return node

    def _build_runs(self, p_elem: etree._Element, parent_ref: Ref) -> list[AccessibilityNode]:
        """Build nodes for runs within a paragraph."""
        runs: list[AccessibilityNode] = []
        run_index = 0

        for r_elem in p_elem.iter(w("r")):
            text = self._extract_text_from_run(r_elem)
            if not text:
                continue

            ref = parent_ref.with_child(ElementType.RUN, run_index)
            run_index += 1

            # Extract run properties
            properties: dict[str, str] = {}
            run_props = r_elem.find(w("rPr"))
            if run_props is not None:
                if run_props.find(w("b")) is not None:
                    properties["bold"] = "true"
                if run_props.find(w("i")) is not None:
                    properties["italic"] = "true"
                if run_props.find(w("u")) is not None:
                    properties["underline"] = "true"

            node = AccessibilityNode(
                ref=ref,
                element_type=ElementType.RUN,
                text=text,
                properties=properties,
                _element=r_elem,
            )
            runs.append(node)

        return runs

    def _build_table(self, tbl_elem: etree._Element) -> AccessibilityNode:
        """Build a node for a table."""
        idx = self._table_index
        self._table_index += 1
        self.stats.tables += 1

        ref = Ref(path=f"tbl:{idx}")

        # Build rows
        rows: list[AccessibilityNode] = []
        row_index = 0

        for tr_elem in tbl_elem.findall(f"./{w('tr')}"):
            row_ref = ref.with_child(ElementType.TABLE_ROW, row_index)
            row_node = self._build_table_row(tr_elem, row_ref, row_index == 0)
            rows.append(row_node)
            row_index += 1

        # Determine table dimensions
        num_rows = len(rows)
        num_cols = len(rows[0].children) if rows else 0

        node = AccessibilityNode(
            ref=ref,
            element_type=ElementType.TABLE,
            children=rows,
            properties={"rows": str(num_rows), "cols": str(num_cols)},
            _element=tbl_elem,
        )

        return node

    def _build_table_row(
        self, tr_elem: etree._Element, ref: Ref, is_header: bool = False
    ) -> AccessibilityNode:
        """Build a node for a table row."""
        cells: list[AccessibilityNode] = []
        cell_index = 0

        for tc_elem in tr_elem.findall(f"./{w('tc')}"):
            cell_ref = ref.with_child(ElementType.TABLE_CELL, cell_index)
            cell_node = self._build_table_cell(tc_elem, cell_ref)
            cells.append(cell_node)
            cell_index += 1

        properties: dict[str, str] = {}
        if is_header:
            properties["header"] = "true"

        return AccessibilityNode(
            ref=ref,
            element_type=ElementType.TABLE_ROW,
            children=cells,
            properties=properties,
            _element=tr_elem,
        )

    def _build_table_cell(self, tc_elem: etree._Element, ref: Ref) -> AccessibilityNode:
        """Build a node for a table cell."""
        # Extract cell text (concatenate all paragraphs)
        text_parts = []
        for p_elem in tc_elem.findall(f"./{w('p')}"):
            text_parts.append(self._extract_text(p_elem))
        text = " ".join(text_parts).strip()

        return AccessibilityNode(
            ref=ref,
            element_type=ElementType.TABLE_CELL,
            text=text,
            _element=tc_elem,
        )

    def _extract_text(self, elem: etree._Element) -> str:
        """Extract all text from an element, including within tracked changes."""
        text_parts = []

        # Get text from regular runs
        for t_elem in elem.iter(w("t")):
            if t_elem.text:
                text_parts.append(t_elem.text)

        # Get deleted text from w:delText
        for del_text_elem in elem.iter(w("delText")):
            if del_text_elem.text:
                text_parts.append(del_text_elem.text)

        return "".join(text_parts)

    def _extract_text_from_run(self, r_elem: etree._Element) -> str:
        """Extract text from a single run."""
        text_parts = []
        for t_elem in r_elem.findall(f"./{w('t')}"):
            if t_elem.text:
                text_parts.append(t_elem.text)
        return "".join(text_parts)

    def _get_paragraph_style(self, p_elem: etree._Element) -> str | None:
        """Get the style name for a paragraph."""
        para_props = p_elem.find(f"./{w('pPr')}")
        if para_props is not None:
            para_style = para_props.find(f"./{w('pStyle')}")
            if para_style is not None:
                return para_style.get(f"{{{WORD_NAMESPACE}}}val")
        return None

    def _get_heading_level(self, style: str | None) -> int | None:
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

    def _extract_tracked_changes(self, elem: etree._Element) -> list[dict[str, str | None]]:
        """Extract tracked change information from an element."""
        changes: list[dict[str, str | None]] = []

        # Find insertions
        for ins_elem in elem.iter(w("ins")):
            change = self._parse_change_element(ins_elem, "insertion")
            if change:
                changes.append(change)

        # Find deletions
        for del_elem in elem.iter(w("del")):
            change = self._parse_change_element(del_elem, "deletion")
            if change:
                changes.append(change)

        return changes

    def _parse_change_element(
        self, elem: etree._Element, change_type: str
    ) -> dict[str, str | None] | None:
        """Parse a tracked change element (w:ins or w:del)."""
        ns = f"{{{WORD_NAMESPACE}}}"

        change_id = elem.get(f"{ns}id")
        author = elem.get(f"{ns}author")
        date_str = elem.get(f"{ns}date")

        # Extract the changed text
        if change_type == "insertion":
            text_parts = []
            for t_elem in elem.iter(w("t")):
                if t_elem.text:
                    text_parts.append(t_elem.text)
            text = "".join(text_parts)
        else:
            # Deletion - look for delText
            text_parts = []
            for del_text_elem in elem.iter(w("delText")):
                if del_text_elem.text:
                    text_parts.append(del_text_elem.text)
            text = "".join(text_parts)

        if not text:
            return None

        return {
            "type": change_type,
            "id": change_id,
            "author": author,
            "date": date_str,
            "text": text,
        }


class _YamlWriter:
    """Internal class for writing YAML output."""

    def __init__(self, verbosity: str, view_mode: ViewMode) -> None:
        self.verbosity = verbosity
        self.view_mode = view_mode
        self.buffer = StringIO()
        self._indent = 0

    def write(self, tree: AccessibilityTree) -> str:
        """Write the tree to YAML."""
        self.buffer = StringIO()

        # Write document header
        self._write_header(tree)

        # Write content
        self._write_line("")
        self._write_line("content:")
        self._indent += 1

        for child in tree.root.children:
            self._write_node(child)

        self._indent -= 1

        # Write tracked changes summary if present
        if tree.stats.tracked_changes > 0 and self.view_mode.include_tracked_changes:
            self._write_tracked_changes_summary(tree)

        # Write bookmarks and links summary
        if tree.stats.bookmarks > 0 or tree.stats.hyperlinks > 0:
            self._write_bookmarks_and_links(tree)

        # Write cross-references summary
        if tree.stats.cross_references > 0:
            self._write_cross_references(tree)

        return self.buffer.getvalue()

    def _write_header(self, tree: AccessibilityTree) -> str:
        """Write the document header section."""
        self._write_line("document:")
        self._indent += 1

        if tree.document_path:
            path_str = str(tree.document_path)
            self._write_line(f'path: "{path_str}"')

        self._write_line(f"verbosity: {self.verbosity}")

        # Stats
        self._write_line("stats:")
        self._indent += 1
        self._write_line(f"paragraphs: {tree.stats.paragraphs}")
        self._write_line(f"tables: {tree.stats.tables}")
        self._write_line(f"tracked_changes: {tree.stats.tracked_changes}")
        if tree.stats.comments > 0:
            self._write_line(f"comments: {tree.stats.comments}")
        if tree.stats.bookmarks > 0:
            self._write_line(f"bookmarks: {tree.stats.bookmarks}")
        if tree.stats.hyperlinks > 0:
            self._write_line(f"hyperlinks: {tree.stats.hyperlinks}")
        if tree.stats.images > 0:
            self._write_line(f"images: {tree.stats.images}")
        if tree.stats.cross_references > 0:
            self._write_line(f"cross_references: {tree.stats.cross_references}")
        self._indent -= 1

        self._indent -= 1
        return self.buffer.getvalue()

    def _write_node(self, node: AccessibilityNode) -> None:
        """Write a single node to YAML."""
        if self.verbosity == "minimal":
            self._write_node_minimal(node)
        elif self.verbosity == "full":
            self._write_node_full(node)
        else:
            self._write_node_standard(node)

    def _write_node_minimal(self, node: AccessibilityNode) -> None:
        """Write a node in minimal format."""
        if node.element_type == ElementType.PARAGRAPH:
            # Format: - h1 "Title" [ref=p:0]  OR  - p "Text..." [ref=p:1]
            prefix = self._get_minimal_prefix(node)
            text = self._truncate_text(node.text, 60)
            self._write_line(f'- {prefix} "{text}" [ref={node.ref}]')

        elif node.element_type == ElementType.TABLE:
            rows = node.properties.get("rows", "?")
            cols = node.properties.get("cols", "?")
            self._write_line(f"- table [ref={node.ref}] [{rows}x{cols}]")

        # Skip other element types in minimal mode

    def _write_node_standard(self, node: AccessibilityNode) -> None:
        """Write a node in standard format."""
        if node.element_type == ElementType.PARAGRAPH:
            self._write_paragraph_standard(node)
        elif node.element_type == ElementType.TABLE:
            self._write_table_standard(node)

    def _write_node_full(self, node: AccessibilityNode) -> None:
        """Write a node in full format."""
        if node.element_type == ElementType.PARAGRAPH:
            self._write_paragraph_full(node)
        elif node.element_type == ElementType.TABLE:
            self._write_table_full(node)

    def _write_paragraph_standard(self, node: AccessibilityNode) -> None:
        """Write a paragraph in standard format."""
        # Determine element label
        if node.level is not None:
            label = f"heading [ref={node.ref}] [level={node.level}]"
        else:
            label = f"paragraph [ref={node.ref}]"

        self._write_line(f"- {label}:")
        self._indent += 1

        # Text content
        text = self._escape_yaml_string(node.text)
        self._write_line(f'text: "{text}"')

        # Style (if present)
        if node.style:
            self._write_line(f"style: {node.style}")

        # Tracked changes
        if node.properties.get("has_changes") == "true":
            self._write_line("has_changes: true")
            changes = node.properties.get("_changes")
            if changes and isinstance(changes, list):
                self._write_changes(changes)

        # Images
        if node.has_images:
            self._write_images(node.images)

        self._indent -= 1

    def _write_paragraph_full(self, node: AccessibilityNode) -> None:
        """Write a paragraph in full format."""
        # Start with standard info
        if node.level is not None:
            label = f"heading [ref={node.ref}] [level={node.level}]"
        else:
            label = f"paragraph [ref={node.ref}]"

        self._write_line(f"- {label}:")
        self._indent += 1

        if node.style:
            self._write_line(f"style: {node.style}")

        # Write runs if present
        if node.children:
            self._write_line("runs:")
            self._indent += 1
            for run in node.children:
                self._write_run_full(run)
            self._indent -= 1
        else:
            text = self._escape_yaml_string(node.text)
            self._write_line(f'text: "{text}"')

        # Tracked changes
        if node.properties.get("has_changes") == "true":
            self._write_line("has_changes: true")
            changes = node.properties.get("_changes")
            if changes and isinstance(changes, list):
                self._write_changes(changes)

        # Images
        if node.has_images:
            self._write_images(node.images, full=True)

        self._indent -= 1

    def _write_run_full(self, node: AccessibilityNode) -> None:
        """Write a run in full format."""
        # Build state annotations
        states = []
        if node.properties.get("bold") == "true":
            states.append("[bold]")
        if node.properties.get("italic") == "true":
            states.append("[italic]")
        if node.properties.get("underline") == "true":
            states.append("[underline]")

        state_str = " ".join(states)
        text = self._escape_yaml_string(node.text)

        if states:
            self._write_line(f'- text "{text}" [ref={node.ref}] {state_str}')
        else:
            self._write_line(f'- text "{text}" [ref={node.ref}]')

    def _write_table_standard(self, node: AccessibilityNode) -> None:
        """Write a table in standard format."""
        rows = node.properties.get("rows", "?")
        cols = node.properties.get("cols", "?")

        self._write_line(f"- table [ref={node.ref}] [rows={rows}] [cols={cols}]:")
        self._indent += 1

        for row_node in node.children:
            self._write_table_row_standard(row_node)

        self._indent -= 1

    def _write_table_row_standard(self, node: AccessibilityNode) -> None:
        """Write a table row in standard format."""
        header_str = " [header]" if node.properties.get("header") == "true" else ""
        self._write_line(f"- row [ref={node.ref}]{header_str}:")
        self._indent += 1

        for cell_node in node.children:
            text = self._truncate_text(cell_node.text, 40)
            self._write_line(f'- cell: "{text}"')

        self._indent -= 1

    def _write_table_full(self, node: AccessibilityNode) -> None:
        """Write a table in full format (same as standard for now)."""
        self._write_table_standard(node)

    def _write_changes(self, changes: list[dict[str, str | None]]) -> None:
        """Write tracked changes list."""
        self._write_line("changes:")
        self._indent += 1

        for change in changes:
            self._write_line(f"- type: {change.get('type', 'unknown')}")
            self._indent += 1

            text = change.get("text", "")
            if text:
                text = self._escape_yaml_string(text)
                self._write_line(f'text: "{text}"')

            author = change.get("author")
            if author:
                self._write_line(f"author: {author}")

            date = change.get("date")
            if date:
                self._write_line(f"date: {date}")

            self._indent -= 1

        self._indent -= 1

    def _write_images(self, images: list[ImageInfo], full: bool = False) -> None:
        """Write images list.

        Args:
            images: List of ImageInfo objects
            full: If True, include more details
        """
        self._write_line("images:")
        self._indent += 1

        for img in images:
            self._write_line(f"- ref: {img.ref}")
            self._indent += 1

            self._write_line(f"type: {img.image_type.name.lower()}")
            pos_type = "inline" if img.is_inline else "floating"
            self._write_line(f"position: {pos_type}")

            if img.name:
                self._write_line(f"name: {img.name}")

            if img.alt_text:
                alt = self._escape_yaml_string(img.alt_text)
                self._write_line(f'alt_text: "{alt}"')

            if img.size:
                self._write_line(f"size: {img.size.to_display_string()}")

            if full:
                if img.relationship_id:
                    self._write_line(f"relationship_id: {img.relationship_id}")
                if img.position and img.is_floating:
                    self._write_line("floating_position:")
                    self._indent += 1
                    if img.position.horizontal:
                        self._write_line(f"horizontal: {img.position.horizontal}")
                    if img.position.vertical:
                        self._write_line(f"vertical: {img.position.vertical}")
                    if img.position.relative_to:
                        self._write_line(f"relative_to: {img.position.relative_to}")
                    if img.position.wrap_type:
                        self._write_line(f"wrap: {img.position.wrap_type}")
                    self._indent -= 1

            self._indent -= 1

        self._indent -= 1

    def _write_tracked_changes_summary(self, tree: AccessibilityTree) -> None:
        """Write summary of tracked changes at the end."""
        self._write_line("")
        self._write_line("tracked_changes:")
        self._indent += 1

        change_index = 0
        for node in tree.iter_nodes():
            changes = node.properties.get("_changes")
            if changes and isinstance(changes, list):
                for change in changes:
                    self._write_line(f"- ref: change:{change_index}")
                    self._indent += 1
                    self._write_line(f"type: {change.get('type', 'unknown')}")

                    change_id = change.get("id")
                    if change_id:
                        self._write_line(f'id: "{change_id}"')

                    author = change.get("author")
                    if author:
                        self._write_line(f"author: {author}")

                    date = change.get("date")
                    if date:
                        self._write_line(f"date: {date}")

                    text = change.get("text", "")
                    if text:
                        text = self._escape_yaml_string(text)
                        self._write_line(f'text: "{text}"')

                    self._write_line(f"location: {node.ref}")
                    self._indent -= 1
                    change_index += 1

        self._indent -= 1

    def _write_bookmarks_and_links(self, tree: AccessibilityTree) -> None:
        """Write bookmarks and links summary."""
        bookmark_data = tree.bookmark_registry.to_yaml_dict()

        # Write bookmarks section
        if "bookmarks" in bookmark_data:
            self._write_line("")
            self._write_line("bookmarks:")
            self._indent += 1

            for bk in bookmark_data["bookmarks"]:
                self._write_line(f"- ref: {bk['ref']}")
                self._indent += 1
                self._write_line(f"name: {bk['name']}")
                self._write_line(f"location: {bk['location']}")
                if bk.get("text_preview"):
                    preview = self._escape_yaml_string(bk["text_preview"])
                    self._write_line(f'text_preview: "{preview}"')
                if bk.get("referenced_by"):
                    refs = bk["referenced_by"]
                    self._write_line("referenced_by:")
                    self._indent += 1
                    for ref in refs:
                        self._write_line(f"- {ref}")
                    self._indent -= 1
                self._indent -= 1

            self._indent -= 1

        # Write links section
        if "links" in bookmark_data:
            self._write_line("")
            self._write_line("links:")
            self._indent += 1

            links_data = bookmark_data["links"]

            # Internal links
            if "internal" in links_data and links_data["internal"]:
                self._write_line("internal:")
                self._indent += 1
                for link in links_data["internal"]:
                    self._write_line(f"- ref: {link['ref']}")
                    self._indent += 1
                    self._write_line(f"from: {link['from']}")
                    self._write_line(f"to: {link['to']}")
                    if link.get("text"):
                        text = self._escape_yaml_string(link["text"])
                        self._write_line(f'text: "{text}"')
                    self._indent -= 1
                self._indent -= 1

            # External links
            if "external" in links_data and links_data["external"]:
                self._write_line("external:")
                self._indent += 1
                for link in links_data["external"]:
                    self._write_line(f"- ref: {link['ref']}")
                    self._indent += 1
                    self._write_line(f"from: {link['from']}")
                    self._write_line(f'url: "{link["url"]}"')
                    if link.get("text"):
                        text = self._escape_yaml_string(link["text"])
                        self._write_line(f'text: "{text}"')
                    self._indent -= 1
                self._indent -= 1

            # Broken links
            if "broken" in links_data and links_data["broken"]:
                self._write_line("broken:")
                self._indent += 1
                for link in links_data["broken"]:
                    self._write_line(f"- ref: {link['ref']}")
                    self._indent += 1
                    self._write_line(f"from: {link['from']}")
                    target = link.get("target", "")
                    if target:
                        self._write_line(f'target: "{target}"')
                    self._write_line(f'error: "{link["error"]}"')
                    self._indent -= 1
                self._indent -= 1

            self._indent -= 1

    def _write_cross_references(self, tree: AccessibilityTree) -> None:
        """Write cross-references summary."""
        xref_data = tree.cross_reference_registry.to_yaml_dict()

        if not xref_data:
            return

        self._write_line("")
        self._write_line("cross_references:")
        self._indent += 1

        # Group by type for organized output
        by_type: dict[str, list[dict]] = {"ref": [], "pageref": [], "noteref": []}
        for xref in xref_data.get("cross_references", []):
            field_type = xref.get("type", "REF").upper()
            if field_type.lower() in by_type:
                by_type[field_type.lower()].append(xref)
            else:
                by_type["ref"].append(xref)

        # Write REF fields (bookmark text references)
        if by_type["ref"]:
            self._write_line("text_references:")
            self._indent += 1
            for xref in by_type["ref"]:
                self._write_cross_reference_entry(xref)
            self._indent -= 1

        # Write PAGEREF fields (page number references)
        if by_type["pageref"]:
            self._write_line("page_references:")
            self._indent += 1
            for xref in by_type["pageref"]:
                self._write_cross_reference_entry(xref)
            self._indent -= 1

        # Write NOTEREF fields (footnote/endnote number references)
        if by_type["noteref"]:
            self._write_line("note_references:")
            self._indent += 1
            for xref in by_type["noteref"]:
                self._write_cross_reference_entry(xref)
            self._indent -= 1

        # Write broken references separately
        broken = [x for x in xref_data.get("cross_references", []) if x.get("is_broken")]
        if broken:
            self._write_line("broken:")
            self._indent += 1
            for xref in broken:
                self._write_line(f"- ref: {xref['ref']}")
                self._indent += 1
                self._write_line(f"target: {xref['target']}")
                if xref.get("error"):
                    self._write_line(f'error: "{xref["error"]}"')
                self._indent -= 1
            self._indent -= 1

        self._indent -= 1

    def _write_cross_reference_entry(self, xref: dict) -> None:
        """Write a single cross-reference entry."""
        self._write_line(f"- ref: {xref['ref']}")
        self._indent += 1

        self._write_line(f"target: {xref['target']}")
        self._write_line(f"from: {xref['from']}")

        if xref.get("display"):
            text = self._escape_yaml_string(xref["display"])
            self._write_line(f'display: "{text}"')

        if xref.get("target_location"):
            self._write_line(f"target_location: {xref['target_location']}")

        if xref.get("is_hyperlink"):
            self._write_line("hyperlink: true")

        if xref.get("is_dirty"):
            self._write_line("needs_update: true")

        self._indent -= 1

    def _write_line(self, text: str) -> None:
        """Write a line with current indentation."""
        indent = "  " * self._indent
        self.buffer.write(f"{indent}{text}\n")

    def _get_minimal_prefix(self, node: AccessibilityNode) -> str:
        """Get the minimal format prefix for a node."""
        if node.level is not None:
            return f"h{node.level}"
        return "p"

    def _truncate_text(self, text: str, max_len: int) -> str:
        """Truncate text with ellipsis if too long."""
        if len(text) <= max_len:
            return text
        return text[: max_len - 3] + "..."

    def _escape_yaml_string(self, text: str) -> str:
        """Escape special characters in a YAML string."""
        # Escape backslashes, quotes, and newlines
        text = text.replace("\\", "\\\\")
        text = text.replace('"', '\\"')
        text = text.replace("\n", "\\n")
        text = text.replace("\r", "\\r")
        text = text.replace("\t", "\\t")
        return text
