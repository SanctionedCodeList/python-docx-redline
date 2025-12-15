"""
HeaderFooterOperations class for handling headers and footers.

This module provides a dedicated class for all header/footer operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from ..constants import WORD_NAMESPACE
from ..errors import AmbiguousTextError, TextNotFoundError
from ..suggestions import SuggestionGenerator

if TYPE_CHECKING:
    from ..document import Document
    from ..models.header_footer import Footer, Header


class HeaderFooterOperations:
    """Handles header and footer operations.

    This class encapsulates all header/footer functionality, including:
    - Accessing headers and footers in the document
    - Replacing text in headers/footers with tracked changes
    - Inserting text in headers/footers with tracked changes
    - Loading and saving header/footer XML files

    The class takes a Document reference and operates on its XML structure.

    Example:
        >>> # Usually accessed through Document
        >>> doc = Document("contract.docx")
        >>> for header in doc.headers:
        ...     print(f"{header.type}: {header.text}")
        >>> doc.replace_in_header("Draft", "Final", header_type="default")
    """

    # Relationship namespace URI
    RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def __init__(self, document: Document) -> None:
        """Initialize HeaderFooterOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    @property
    def headers(self) -> list[Header]:
        """Get all headers in the document.

        Headers are linked via relationships in section properties (sectPr).
        Each section can have up to three headers: default, first, even.

        Returns:
            List of Header objects

        Example:
            >>> doc = Document("contract.docx")
            >>> for header in doc.headers:
            ...     print(f"{header.type}: {header.text}")
        """
        from ..models.header_footer import Header, HeaderFooterType

        if not self._document._temp_dir:
            return []

        # Load relationships to map rId -> filename
        rel_map = self._load_document_relationships()

        # Find all header references in sectPr elements
        headers: list[Header] = []
        seen_files: set[str] = set()

        # Header reference element names and their types
        header_ref_types = {
            f"{{{WORD_NAMESPACE}}}headerReference": {
                "default": HeaderFooterType.DEFAULT,
                "first": HeaderFooterType.FIRST,
                "even": HeaderFooterType.EVEN,
            }
        }

        for sect_pr in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}sectPr"):
            for ref_tag, type_map in header_ref_types.items():
                for ref in sect_pr.findall(ref_tag):
                    rel_id = ref.get(f"{{{self.RELATIONSHIP_NAMESPACE}}}id", "")
                    type_attr = ref.get(f"{{{WORD_NAMESPACE}}}type", "default")
                    header_type = type_map.get(type_attr, HeaderFooterType.DEFAULT)

                    if rel_id in rel_map:
                        target = rel_map[rel_id]
                        # Skip if already processed
                        file_key = f"{target}:{type_attr}"
                        if file_key in seen_files:
                            continue
                        seen_files.add(file_key)

                        # Load header XML
                        header_elem = self._load_header_footer_xml(target)
                        if header_elem is not None:
                            headers.append(
                                Header(
                                    element=header_elem,
                                    document=self._document,
                                    header_type=header_type,
                                    rel_id=rel_id,
                                    file_path=target,
                                )
                            )

        return headers

    @property
    def footers(self) -> list[Footer]:
        """Get all footers in the document.

        Footers are linked via relationships in section properties (sectPr).
        Each section can have up to three footers: default, first, even.

        Returns:
            List of Footer objects

        Example:
            >>> doc = Document("contract.docx")
            >>> for footer in doc.footers:
            ...     print(f"{footer.type}: {footer.text}")
        """
        from ..models.header_footer import Footer, HeaderFooterType

        if not self._document._temp_dir:
            return []

        # Load relationships to map rId -> filename
        rel_map = self._load_document_relationships()

        # Find all footer references in sectPr elements
        footers: list[Footer] = []
        seen_files: set[str] = set()

        # Footer reference element names and their types
        footer_ref_types = {
            f"{{{WORD_NAMESPACE}}}footerReference": {
                "default": HeaderFooterType.DEFAULT,
                "first": HeaderFooterType.FIRST,
                "even": HeaderFooterType.EVEN,
            }
        }

        for sect_pr in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}sectPr"):
            for ref_tag, type_map in footer_ref_types.items():
                for ref in sect_pr.findall(ref_tag):
                    rel_id = ref.get(f"{{{self.RELATIONSHIP_NAMESPACE}}}id", "")
                    type_attr = ref.get(f"{{{WORD_NAMESPACE}}}type", "default")
                    footer_type = type_map.get(type_attr, HeaderFooterType.DEFAULT)

                    if rel_id in rel_map:
                        target = rel_map[rel_id]
                        # Skip if already processed
                        file_key = f"{target}:{type_attr}"
                        if file_key in seen_files:
                            continue
                        seen_files.add(file_key)

                        # Load footer XML
                        footer_elem = self._load_header_footer_xml(target)
                        if footer_elem is not None:
                            footers.append(
                                Footer(
                                    element=footer_elem,
                                    document=self._document,
                                    footer_type=footer_type,
                                    rel_id=rel_id,
                                    file_path=target,
                                )
                            )

        return footers

    def _load_document_relationships(self) -> dict[str, str]:
        """Load document.xml.rels and return rId -> Target mapping.

        Returns:
            Dictionary mapping relationship IDs to target filenames
        """
        if not self._document._temp_dir:
            return {}

        rels_path = self._document._temp_dir / "word" / "_rels" / "document.xml.rels"
        if not rels_path.exists():
            return {}

        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        tree = etree.parse(str(rels_path))
        root = tree.getroot()

        rel_map: dict[str, str] = {}
        for rel in root.findall(f"{{{rels_ns}}}Relationship"):
            rel_id = rel.get("Id", "")
            target = rel.get("Target", "")
            if rel_id and target:
                rel_map[rel_id] = target

        return rel_map

    def _load_header_footer_xml(self, target: str) -> etree._Element | None:
        """Load a header or footer XML file.

        Args:
            target: The target path from relationships (e.g., "header1.xml")

        Returns:
            The root element of the header/footer XML, or None if not found
        """
        if not self._document._temp_dir:
            return None

        # Handle relative paths - they're relative to word/
        if not target.startswith("/"):
            file_path = self._document._temp_dir / "word" / target
        else:
            file_path = self._document._temp_dir / target.lstrip("/")

        if not file_path.exists():
            return None

        tree = etree.parse(str(file_path))
        return tree.getroot()

    def _save_header_footer_xml(self, target: str, root: etree._Element) -> None:
        """Save a header or footer XML file.

        Args:
            target: The target path from relationships (e.g., "header1.xml")
            root: The root element to save
        """
        if not self._document._temp_dir:
            return

        # Handle relative paths
        if not target.startswith("/"):
            file_path = self._document._temp_dir / "word" / target
        else:
            file_path = self._document._temp_dir / target.lstrip("/")

        tree = etree.ElementTree(root)
        tree.write(
            str(file_path),
            encoding="utf-8",
            xml_declaration=True,
            standalone=True,
        )

    def _get_header_by_type(self, header_type: str) -> Header | None:
        """Get a header by its type.

        Args:
            header_type: "default", "first", or "even"

        Returns:
            Header object or None if not found
        """
        for header in self.headers:
            if header.type == header_type:
                return header
        return None

    def _get_footer_by_type(self, footer_type: str) -> Footer | None:
        """Get a footer by its type.

        Args:
            footer_type: "default", "first", or "even"

        Returns:
            Footer object or None if not found
        """
        for footer in self.footers:
            if footer.type == footer_type:
                return footer
        return None

    def replace_in_header(
        self,
        find: str,
        replace: str,
        header_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Replace text in a header with tracked changes.

        Args:
            find: Text to find
            replace: Replacement text
            header_type: Type of header ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat 'find' as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching

        Raises:
            TextNotFoundError: If 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
            ValueError: If no header of the specified type exists

        Example:
            >>> doc.replace_in_header("Draft", "Final", header_type="default")
        """
        header = self._get_header_by_type(header_type)
        if header is None:
            raise ValueError(f"No header of type '{header_type}' found in document")

        # Perform the replacement using the header's paragraphs
        self._replace_in_header_footer(
            header=header,
            find=find,
            replace=replace,
            author=author,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
        )

    def replace_in_footer(
        self,
        find: str,
        replace: str,
        footer_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Replace text in a footer with tracked changes.

        Args:
            find: Text to find
            replace: Replacement text
            footer_type: Type of footer ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat 'find' as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching

        Raises:
            TextNotFoundError: If 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
            ValueError: If no footer of the specified type exists

        Example:
            >>> doc.replace_in_footer("Page {PAGE}", "Page {PAGE} of {NUMPAGES}")
        """
        footer = self._get_footer_by_type(footer_type)
        if footer is None:
            raise ValueError(f"No footer of type '{footer_type}' found in document")

        # Perform the replacement using the footer's paragraphs
        self._replace_in_header_footer(
            footer=footer,
            find=find,
            replace=replace,
            author=author,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
        )

    def insert_in_header(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        header_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Insert text in a header with tracked changes.

        Args:
            text: Text to insert
            after: Text to insert after
            before: Text to insert before
            header_type: Type of header ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat anchor as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching

        Raises:
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
            ValueError: If no header of the specified type exists, or if both
                        'after' and 'before' are specified

        Example:
            >>> doc.insert_in_header(" - Final", after="Document Title")
        """
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        header = self._get_header_by_type(header_type)
        if header is None:
            raise ValueError(f"No header of type '{header_type}' found in document")

        anchor = after if after is not None else before
        assert anchor is not None  # Guaranteed by check above
        insert_after = after is not None

        self._insert_in_header_footer(
            header=header,
            text=text,
            anchor=anchor,
            insert_after=insert_after,
            author=author,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
        )

    def insert_in_footer(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        footer_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Insert text in a footer with tracked changes.

        Args:
            text: Text to insert
            after: Text to insert after
            before: Text to insert before
            footer_type: Type of footer ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat anchor as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching

        Raises:
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
            ValueError: If no footer of the specified type exists, or if both
                        'after' and 'before' are specified

        Example:
            >>> doc.insert_in_footer(" - Confidential", after="Page")
        """
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        footer = self._get_footer_by_type(footer_type)
        if footer is None:
            raise ValueError(f"No footer of type '{footer_type}' found in document")

        anchor = after if after is not None else before
        assert anchor is not None  # Guaranteed by check above
        insert_after = after is not None

        self._insert_in_header_footer(
            footer=footer,
            text=text,
            anchor=anchor,
            insert_after=insert_after,
            author=author,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
        )

    def _replace_in_header_footer(
        self,
        find: str,
        replace: str,
        author: str | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
        header: Header | None = None,
        footer: Footer | None = None,
    ) -> None:
        """Replace text in a header or footer with tracked changes.

        Args:
            find: Text to find
            replace: Replacement text
            author: Optional author override
            regex: Whether to treat 'find' as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching
            header: Header object (if replacing in header)
            footer: Footer object (if replacing in footer)
        """
        # Get the element and file path
        if header is not None:
            element = header.element
            file_path = header.file_path
        elif footer is not None:
            element = footer.element
            file_path = footer.file_path
        else:
            raise ValueError("Must provide either header or footer")

        # Get paragraphs from the header/footer
        paragraphs = list(element.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Search for text
        matches = self._document._text_search.find_text(
            find,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(find, paragraphs)
            raise TextNotFoundError(find, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(find, matches)

        match = matches[0]

        # Generate the replacement XML (deletion + insertion)
        deletion_xml = self._document._xml_generator.create_deletion(find, author)
        insertion_xml = self._document._xml_generator.create_insertion(replace, author)

        # Parse the XML elements
        wrapped_deletion = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}{insertion_xml}
</root>"""
        root = etree.fromstring(wrapped_deletion.encode("utf-8"))
        deletion_element = root[0]
        insertion_element = root[1]

        # Replace the match with deletion + insertion
        self._document._replace_match_with_elements(match, [deletion_element, insertion_element])

        # Save the modified header/footer XML
        self._save_header_footer_xml(file_path, element)

    def _insert_in_header_footer(
        self,
        text: str,
        anchor: str,
        insert_after: bool,
        author: str | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
        header: Header | None = None,
        footer: Footer | None = None,
    ) -> None:
        """Insert text in a header or footer with tracked changes.

        Args:
            text: Text to insert
            anchor: Text to find for insertion point
            insert_after: Whether to insert after (True) or before (False) anchor
            author: Optional author override
            regex: Whether to treat anchor as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching
            header: Header object (if inserting in header)
            footer: Footer object (if inserting in footer)
        """
        # Get the element and file path
        if header is not None:
            element = header.element
            file_path = header.file_path
        elif footer is not None:
            element = footer.element
            file_path = footer.file_path
        else:
            raise ValueError("Must provide either header or footer")

        # Get paragraphs from the header/footer
        paragraphs = list(element.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Search for text
        matches = self._document._text_search.find_text(
            anchor,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(anchor, paragraphs)
            raise TextNotFoundError(anchor, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(anchor, matches)

        match = matches[0]

        # Generate the insertion XML
        insertion_xml = self._document._xml_generator.create_insertion(text, author)

        # Parse the insertion XML
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {insertion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        insertion_element = root[0]

        # Insert at the appropriate position
        if insert_after:
            self._document._insert_after_match(match, insertion_element)
        else:
            self._document._insert_before_match(match, insertion_element)

        # Save the modified header/footer XML
        self._save_header_footer_xml(file_path, element)
