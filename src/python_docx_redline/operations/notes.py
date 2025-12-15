"""
NoteOperations class for handling footnotes and endnotes.

This module provides a dedicated class for all footnote/endnote operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from lxml import etree

from ..constants import WORD_NAMESPACE
from ..content_types import ContentTypeManager, ContentTypes
from ..errors import AmbiguousTextError, TextNotFoundError
from ..relationships import RelationshipManager, RelationshipTypes
from ..scope import ScopeEvaluator

if TYPE_CHECKING:
    from ..document import Document
    from ..models.footnote import Endnote, Footnote


class NoteOperations:
    """Handles footnote and endnote operations.

    This class encapsulates all footnote/endnote functionality, including:
    - Accessing footnotes and endnotes in the document
    - Inserting new footnotes and endnotes at specific locations
    - Managing footnote/endnote XML files and relationships

    The class takes a Document reference and operates on its XML structure.

    Example:
        >>> # Usually accessed through Document
        >>> doc = Document("contract.docx")
        >>> footnotes = doc.footnotes
        >>> doc.insert_footnote("Citation text", at="quoted passage")
    """

    def __init__(self, document: Document) -> None:
        """Initialize NoteOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    @property
    def footnotes(self) -> list[Footnote]:
        """Get all footnotes in the document.

        Returns:
            List of Footnote objects
        """
        from ..models.footnote import Footnote

        temp_dir = self._document._temp_dir
        if not temp_dir:
            return []

        footnotes_path = temp_dir / "word" / "footnotes.xml"
        if not footnotes_path.exists():
            return []

        tree = etree.parse(str(footnotes_path))
        root = tree.getroot()

        # Find all footnote elements
        footnote_elems = root.findall(f"{{{WORD_NAMESPACE}}}footnote")

        # Filter out special footnotes (separator, continuationSeparator)
        # These have type attribute and IDs -1, 0, etc.
        return [
            Footnote(elem, self._document)
            for elem in footnote_elems
            if elem.get(f"{{{WORD_NAMESPACE}}}type") is None
        ]

    @property
    def endnotes(self) -> list[Endnote]:
        """Get all endnotes in the document.

        Returns:
            List of Endnote objects
        """
        from ..models.footnote import Endnote

        temp_dir = self._document._temp_dir
        if not temp_dir:
            return []

        endnotes_path = temp_dir / "word" / "endnotes.xml"
        if not endnotes_path.exists():
            return []

        tree = etree.parse(str(endnotes_path))
        root = tree.getroot()

        # Find all endnote elements
        endnote_elems = root.findall(f"{{{WORD_NAMESPACE}}}endnote")

        # Filter out special endnotes (separator, continuationSeparator)
        return [
            Endnote(elem, self._document)
            for elem in endnote_elems
            if elem.get(f"{{{WORD_NAMESPACE}}}type") is None
        ]

    def insert_footnote(
        self,
        text: str,
        at: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Insert a footnote reference at specific text location.

        Args:
            text: The footnote text content
            at: Text to search for where footnote reference should be inserted
            author: Optional author (uses document author if None)
            scope: Optional scope to limit search

        Returns:
            The footnote ID

        Raises:
            TextNotFoundError: If 'at' text not found
            AmbiguousTextError: If multiple occurrences of 'at' text found

        Example:
            >>> doc.insert_footnote("See Smith (2020) for details", at="original study")
        """
        if not self._document._is_zip or not self._document._temp_dir:
            raise ValueError("Cannot add footnotes to non-ZIP documents")

        author_name = author if author is not None else self._document.author

        # Find location for footnote reference
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(at, paragraphs)

        if not matches:
            raise TextNotFoundError(at, scope)

        if len(matches) > 1:
            raise AmbiguousTextError(at, matches)

        match = matches[0]

        # Generate new footnote ID
        footnote_id = self._get_next_footnote_id()

        # Add footnote content to footnotes.xml
        self._add_footnote_to_xml(footnote_id, text, author_name)

        # Insert footnote reference in document
        self._insert_footnote_reference(match, footnote_id)

        return footnote_id

    def insert_endnote(
        self,
        text: str,
        at: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Insert an endnote reference at specific text location.

        Args:
            text: The endnote text content
            at: Text to search for where endnote reference should be inserted
            author: Optional author (uses document author if None)
            scope: Optional scope to limit search

        Returns:
            The endnote ID

        Raises:
            TextNotFoundError: If 'at' text not found
            AmbiguousTextError: If multiple occurrences of 'at' text found

        Example:
            >>> doc.insert_endnote("Additional details here", at="main conclusion")
        """
        if not self._document._is_zip or not self._document._temp_dir:
            raise ValueError("Cannot add endnotes to non-ZIP documents")

        author_name = author if author is not None else self._document.author

        # Find location for endnote reference
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(at, paragraphs)

        if not matches:
            raise TextNotFoundError(at, scope)

        if len(matches) > 1:
            raise AmbiguousTextError(at, matches)

        match = matches[0]

        # Generate new endnote ID
        endnote_id = self._get_next_endnote_id()

        # Add endnote content to endnotes.xml
        self._add_endnote_to_xml(endnote_id, text, author_name)

        # Insert endnote reference in document
        self._insert_endnote_reference(match, endnote_id)

        return endnote_id

    def _get_next_footnote_id(self) -> int:
        """Get the next available footnote ID.

        Returns:
            Integer ID for new footnote
        """
        temp_dir = self._document._temp_dir
        footnotes_path = temp_dir / "word" / "footnotes.xml"

        if not footnotes_path.exists():
            return 1

        tree = etree.parse(str(footnotes_path))
        root = tree.getroot()

        # Find all footnote IDs
        footnote_elems = root.findall(f"{{{WORD_NAMESPACE}}}footnote")
        ids = []

        for elem in footnote_elems:
            id_str = elem.get(f"{{{WORD_NAMESPACE}}}id")
            if id_str:
                try:
                    ids.append(int(id_str))
                except ValueError:
                    pass

        return max(ids) + 1 if ids else 1

    def _get_next_endnote_id(self) -> int:
        """Get the next available endnote ID.

        Returns:
            Integer ID for new endnote
        """
        temp_dir = self._document._temp_dir
        endnotes_path = temp_dir / "word" / "endnotes.xml"

        if not endnotes_path.exists():
            return 1

        tree = etree.parse(str(endnotes_path))
        root = tree.getroot()

        # Find all endnote IDs
        endnote_elems = root.findall(f"{{{WORD_NAMESPACE}}}endnote")
        ids = []

        for elem in endnote_elems:
            id_str = elem.get(f"{{{WORD_NAMESPACE}}}id")
            if id_str:
                try:
                    ids.append(int(id_str))
                except ValueError:
                    pass

        return max(ids) + 1 if ids else 1

    def _add_footnote_to_xml(self, footnote_id: int, text: str, author: str) -> None:
        """Add a footnote to footnotes.xml, creating the file if needed.

        Args:
            footnote_id: The footnote ID
            text: Footnote text content
            author: Author name (for tracking if needed)
        """
        temp_dir = self._document._temp_dir
        footnotes_path = temp_dir / "word" / "footnotes.xml"

        # Load or create footnotes.xml
        if footnotes_path.exists():
            footnotes_tree = etree.parse(str(footnotes_path))
            footnotes_root = footnotes_tree.getroot()
        else:
            # Create new footnotes.xml with separators
            footnotes_root = etree.Element(
                f"{{{WORD_NAMESPACE}}}footnotes",
                nsmap={"w": WORD_NAMESPACE},
            )
            footnotes_tree = etree.ElementTree(footnotes_root)

            # Add standard footnote separators (required by Word)
            # Separator (ID -1)
            sep = etree.SubElement(footnotes_root, f"{{{WORD_NAMESPACE}}}footnote")
            sep.set(f"{{{WORD_NAMESPACE}}}id", "-1")
            sep.set(f"{{{WORD_NAMESPACE}}}type", "separator")
            sep_p = etree.SubElement(sep, f"{{{WORD_NAMESPACE}}}p")
            sep_r = etree.SubElement(sep_p, f"{{{WORD_NAMESPACE}}}r")
            etree.SubElement(sep_r, f"{{{WORD_NAMESPACE}}}separator")

            # Continuation separator (ID 0)
            cont_sep = etree.SubElement(footnotes_root, f"{{{WORD_NAMESPACE}}}footnote")
            cont_sep.set(f"{{{WORD_NAMESPACE}}}id", "0")
            cont_sep.set(f"{{{WORD_NAMESPACE}}}type", "continuationSeparator")
            cont_sep_p = etree.SubElement(cont_sep, f"{{{WORD_NAMESPACE}}}p")
            cont_sep_r = etree.SubElement(cont_sep_p, f"{{{WORD_NAMESPACE}}}r")
            etree.SubElement(cont_sep_r, f"{{{WORD_NAMESPACE}}}continuationSeparator")

            # Need to add relationship and content type
            self._ensure_footnotes_relationship()
            self._ensure_footnotes_content_type()

        # Create footnote element
        footnote_elem = etree.SubElement(footnotes_root, f"{{{WORD_NAMESPACE}}}footnote")
        footnote_elem.set(f"{{{WORD_NAMESPACE}}}id", str(footnote_id))

        # Add paragraph with text
        para = etree.SubElement(footnote_elem, f"{{{WORD_NAMESPACE}}}p")
        run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = text

        # Write footnotes.xml
        footnotes_tree.write(
            str(footnotes_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _add_endnote_to_xml(self, endnote_id: int, text: str, author: str) -> None:
        """Add an endnote to endnotes.xml, creating the file if needed.

        Args:
            endnote_id: The endnote ID
            text: Endnote text content
            author: Author name (for tracking if needed)
        """
        temp_dir = self._document._temp_dir
        endnotes_path = temp_dir / "word" / "endnotes.xml"

        # Load or create endnotes.xml
        if endnotes_path.exists():
            endnotes_tree = etree.parse(str(endnotes_path))
            endnotes_root = endnotes_tree.getroot()
        else:
            # Create new endnotes.xml with separators
            endnotes_root = etree.Element(
                f"{{{WORD_NAMESPACE}}}endnotes",
                nsmap={"w": WORD_NAMESPACE},
            )
            endnotes_tree = etree.ElementTree(endnotes_root)

            # Add standard endnote separators
            sep = etree.SubElement(endnotes_root, f"{{{WORD_NAMESPACE}}}endnote")
            sep.set(f"{{{WORD_NAMESPACE}}}id", "-1")
            sep.set(f"{{{WORD_NAMESPACE}}}type", "separator")
            sep_p = etree.SubElement(sep, f"{{{WORD_NAMESPACE}}}p")
            sep_r = etree.SubElement(sep_p, f"{{{WORD_NAMESPACE}}}r")
            etree.SubElement(sep_r, f"{{{WORD_NAMESPACE}}}separator")

            cont_sep = etree.SubElement(endnotes_root, f"{{{WORD_NAMESPACE}}}endnote")
            cont_sep.set(f"{{{WORD_NAMESPACE}}}id", "0")
            cont_sep.set(f"{{{WORD_NAMESPACE}}}type", "continuationSeparator")
            cont_sep_p = etree.SubElement(cont_sep, f"{{{WORD_NAMESPACE}}}p")
            cont_sep_r = etree.SubElement(cont_sep_p, f"{{{WORD_NAMESPACE}}}r")
            etree.SubElement(cont_sep_r, f"{{{WORD_NAMESPACE}}}continuationSeparator")

            # Need to add relationship and content type
            self._ensure_endnotes_relationship()
            self._ensure_endnotes_content_type()

        # Create endnote element
        endnote_elem = etree.SubElement(endnotes_root, f"{{{WORD_NAMESPACE}}}endnote")
        endnote_elem.set(f"{{{WORD_NAMESPACE}}}id", str(endnote_id))

        # Add paragraph with text
        para = etree.SubElement(endnote_elem, f"{{{WORD_NAMESPACE}}}p")
        run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = text

        # Write endnotes.xml
        endnotes_tree.write(
            str(endnotes_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _insert_footnote_reference(self, match: Any, footnote_id: int) -> None:
        """Insert a footnote reference at the matched text location.

        Args:
            match: TextSpan object indicating where to insert
            footnote_id: The footnote ID to reference
        """
        # Find the run where the match ends
        end_run = match.runs[-1] if match.runs else None
        if end_run is None:
            return

        # Create a new run with the footnote reference
        new_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")

        # Add footnote reference
        footnote_ref = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}footnoteReference")
        footnote_ref.set(f"{{{WORD_NAMESPACE}}}id", str(footnote_id))

        # Insert the new run after the last run of the match
        parent = end_run.getparent()
        index = list(parent).index(end_run)
        parent.insert(index + 1, new_run)

    def _insert_endnote_reference(self, match: Any, endnote_id: int) -> None:
        """Insert an endnote reference at the matched text location.

        Args:
            match: TextSpan object indicating where to insert
            endnote_id: The endnote ID to reference
        """
        # Find the run where the match ends
        end_run = match.runs[-1] if match.runs else None
        if end_run is None:
            return

        # Create a new run with the endnote reference
        new_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")

        # Add endnote reference
        endnote_ref = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}endnoteReference")
        endnote_ref.set(f"{{{WORD_NAMESPACE}}}id", str(endnote_id))

        # Insert the new run after the last run of the match
        parent = end_run.getparent()
        index = list(parent).index(end_run)
        parent.insert(index + 1, new_run)

    def _ensure_footnotes_relationship(self) -> None:
        """Ensure footnotes.xml relationship exists in document.xml.rels."""
        package = self._document._package
        if not package:
            return

        rel_mgr = RelationshipManager(package, "word/document.xml")
        rel_mgr.add_relationship(RelationshipTypes.FOOTNOTES, "footnotes.xml")
        rel_mgr.save()

    def _ensure_endnotes_relationship(self) -> None:
        """Ensure endnotes.xml relationship exists in document.xml.rels."""
        package = self._document._package
        if not package:
            return

        rel_mgr = RelationshipManager(package, "word/document.xml")
        rel_mgr.add_relationship(RelationshipTypes.ENDNOTES, "endnotes.xml")
        rel_mgr.save()

    def _ensure_footnotes_content_type(self) -> None:
        """Ensure footnotes.xml content type exists in [Content_Types].xml."""
        package = self._document._package
        if not package:
            return

        ct_mgr = ContentTypeManager(package)
        ct_mgr.add_override("/word/footnotes.xml", ContentTypes.FOOTNOTES)
        ct_mgr.save()

    def _ensure_endnotes_content_type(self) -> None:
        """Ensure endnotes.xml content type exists in [Content_Types].xml."""
        package = self._document._package
        if not package:
            return

        ct_mgr = ContentTypeManager(package)
        ct_mgr.add_override("/word/endnotes.xml", ContentTypes.ENDNOTES)
        ct_mgr.save()
