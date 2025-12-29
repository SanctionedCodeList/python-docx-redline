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
from typing import TYPE_CHECKING

from lxml import etree

from ..constants import WORD_NAMESPACE, w
from .registry import RefRegistry
from .types import (
    AccessibilityNode,
    ElementType,
    Ref,
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
    """

    paragraphs: int = 0
    tables: int = 0
    tracked_changes: int = 0
    comments: int = 0


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
    ) -> None:
        """Initialize an AccessibilityTree.

        Args:
            root: Root node of the tree
            registry: RefRegistry for ref resolution
            view_mode: View configuration (defaults to standard)
            stats: Document statistics
            document_path: Path to source document
        """
        self.root = root
        self.registry = registry
        self.view_mode = view_mode or ViewMode()
        self.stats = stats or DocumentStats()
        self.document_path = document_path

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

        return cls(
            root=root,
            registry=registry,
            view_mode=view_mode,
            stats=stats,
            document_path=getattr(document, "path", None),
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

        return cls(
            root=root,
            registry=registry,
            view_mode=view_mode,
            stats=stats,
            document_path=document_path,
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

    def __repr__(self) -> str:
        return (
            f"AccessibilityTree(paragraphs={self.stats.paragraphs}, "
            f"tables={self.stats.tables}, "
            f"tracked_changes={self.stats.tracked_changes})"
        )


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

        # Create the node
        node = AccessibilityNode(
            ref=ref,
            element_type=ElementType.PARAGRAPH,
            text=text,
            style=style,
            level=level,
            _element=p_elem,
        )

        # Add changes info to properties if present
        if changes and self.view_mode.include_tracked_changes:
            node.properties["has_changes"] = "true"
            node.properties["change_count"] = str(len(changes))

            # Store change details for YAML output
            node.properties["_changes"] = changes  # type: ignore[assignment]

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
