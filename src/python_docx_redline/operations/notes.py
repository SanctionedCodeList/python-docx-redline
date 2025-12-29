"""
NoteOperations class for handling footnotes and endnotes.

This module provides a dedicated class for all footnote/endnote operations,
extracted from the main Document class to improve separation of concerns.

Rich content support:
- Multi-paragraph notes: Pass a list of strings to create multiple paragraphs
- Markdown formatting: **bold**, *italic*, ++underline++, ~~strikethrough~~
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Any

from lxml import etree

from ..constants import WORD_NAMESPACE
from ..content_types import ContentTypeManager, ContentTypes
from ..errors import AmbiguousTextError, NoteNotFoundError, TextNotFoundError
from ..markdown_parser import parse_markdown
from ..relationships import RelationshipManager, RelationshipTypes
from ..scope import ScopeEvaluator
from ..text_search import TextSpan

if TYPE_CHECKING:
    from ..document import Document
    from ..models.footnote import Endnote, Footnote, FootnoteReference


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

    def get_footnote(self, note_id: str | int) -> Footnote:
        """Get a specific footnote by ID.

        Args:
            note_id: The footnote ID to retrieve

        Returns:
            The Footnote object

        Raises:
            NoteNotFoundError: If the footnote ID is not found

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> print(footnote.text)
        """
        note_id_str = str(note_id)
        footnotes = self.footnotes

        for footnote in footnotes:
            if footnote.id == note_id_str:
                return footnote

        available_ids = [fn.id for fn in footnotes]
        raise NoteNotFoundError("footnote", note_id_str, available_ids)

    def get_endnote(self, note_id: str | int) -> Endnote:
        """Get a specific endnote by ID.

        Args:
            note_id: The endnote ID to retrieve

        Returns:
            The Endnote object

        Raises:
            NoteNotFoundError: If the endnote ID is not found

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> print(endnote.text)
        """
        note_id_str = str(note_id)
        endnotes = self.endnotes

        for endnote in endnotes:
            if endnote.id == note_id_str:
                return endnote

        available_ids = [en.id for en in endnotes]
        raise NoteNotFoundError("endnote", note_id_str, available_ids)

    def delete_footnote(self, note_id: str | int, renumber: bool = True) -> None:
        """Delete a footnote by ID.

        This removes both the footnote content from footnotes.xml and the
        corresponding footnote reference from the document body.

        Args:
            note_id: The footnote ID to delete
            renumber: If True, renumber remaining footnotes sequentially (default)

        Raises:
            NoteNotFoundError: If the footnote ID is not found

        Example:
            >>> doc.delete_footnote(2)  # Delete footnote 2 and renumber
            >>> doc.delete_footnote(2, renumber=False)  # Delete without renumbering
        """
        note_id_str = str(note_id)

        # Verify footnote exists (will raise NoteNotFoundError if not found)
        self.get_footnote(note_id_str)

        temp_dir = self._document._temp_dir
        if not temp_dir:
            raise ValueError("Cannot delete footnotes from non-ZIP documents")

        footnotes_path = temp_dir / "word" / "footnotes.xml"

        # Remove from footnotes.xml
        if footnotes_path.exists():
            tree = etree.parse(str(footnotes_path))
            root = tree.getroot()

            # Find and remove the footnote element
            for fn_elem in root.findall(f"{{{WORD_NAMESPACE}}}footnote"):
                if fn_elem.get(f"{{{WORD_NAMESPACE}}}id") == note_id_str:
                    root.remove(fn_elem)
                    break

            tree.write(
                str(footnotes_path),
                encoding="utf-8",
                xml_declaration=True,
                pretty_print=True,
            )

        # Remove footnoteReference from document.xml
        self._remove_footnote_reference(note_id_str)

        # Renumber if requested
        if renumber:
            self._renumber_footnotes()

    def delete_endnote(self, note_id: str | int, renumber: bool = True) -> None:
        """Delete an endnote by ID.

        This removes both the endnote content from endnotes.xml and the
        corresponding endnote reference from the document body.

        Args:
            note_id: The endnote ID to delete
            renumber: If True, renumber remaining endnotes sequentially (default)

        Raises:
            NoteNotFoundError: If the endnote ID is not found

        Example:
            >>> doc.delete_endnote(2)  # Delete endnote 2 and renumber
            >>> doc.delete_endnote(2, renumber=False)  # Delete without renumbering
        """
        note_id_str = str(note_id)

        # Verify endnote exists (will raise NoteNotFoundError if not found)
        self.get_endnote(note_id_str)

        temp_dir = self._document._temp_dir
        if not temp_dir:
            raise ValueError("Cannot delete endnotes from non-ZIP documents")

        endnotes_path = temp_dir / "word" / "endnotes.xml"

        # Remove from endnotes.xml
        if endnotes_path.exists():
            tree = etree.parse(str(endnotes_path))
            root = tree.getroot()

            # Find and remove the endnote element
            for en_elem in root.findall(f"{{{WORD_NAMESPACE}}}endnote"):
                if en_elem.get(f"{{{WORD_NAMESPACE}}}id") == note_id_str:
                    root.remove(en_elem)
                    break

            tree.write(
                str(endnotes_path),
                encoding="utf-8",
                xml_declaration=True,
                pretty_print=True,
            )

        # Remove endnoteReference from document.xml
        self._remove_endnote_reference(note_id_str)

        # Renumber if requested
        if renumber:
            self._renumber_endnotes()

    def edit_footnote(
        self,
        note_id: str | int,
        new_text: str,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Edit the text of a footnote.

        Args:
            note_id: The footnote ID to edit
            new_text: The new text content for the footnote
            track: If True, track the edit as a change (Phase 3 feature)
            author: Author name for tracked changes (uses document author if None)

        Raises:
            NoteNotFoundError: If the footnote ID is not found

        Example:
            >>> doc.edit_footnote(1, "Updated citation text")
        """
        note_id_str = str(note_id)

        # Verify footnote exists (will raise NoteNotFoundError if not found)
        self.get_footnote(note_id_str)

        temp_dir = self._document._temp_dir
        if not temp_dir:
            raise ValueError("Cannot edit footnotes in non-ZIP documents")

        footnotes_path = temp_dir / "word" / "footnotes.xml"

        if footnotes_path.exists():
            tree = etree.parse(str(footnotes_path))
            root = tree.getroot()

            # Find the footnote element
            for fn_elem in root.findall(f"{{{WORD_NAMESPACE}}}footnote"):
                if fn_elem.get(f"{{{WORD_NAMESPACE}}}id") == note_id_str:
                    self._replace_note_content(fn_elem, new_text)
                    break

            tree.write(
                str(footnotes_path),
                encoding="utf-8",
                xml_declaration=True,
                pretty_print=True,
            )

    def edit_endnote(
        self,
        note_id: str | int,
        new_text: str,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Edit the text of an endnote.

        Args:
            note_id: The endnote ID to edit
            new_text: The new text content for the endnote
            track: If True, track the edit as a change (Phase 3 feature)
            author: Author name for tracked changes (uses document author if None)

        Raises:
            NoteNotFoundError: If the endnote ID is not found

        Example:
            >>> doc.edit_endnote(1, "Updated citation text")
        """
        note_id_str = str(note_id)

        # Verify endnote exists (will raise NoteNotFoundError if not found)
        self.get_endnote(note_id_str)

        temp_dir = self._document._temp_dir
        if not temp_dir:
            raise ValueError("Cannot edit endnotes in non-ZIP documents")

        endnotes_path = temp_dir / "word" / "endnotes.xml"

        if endnotes_path.exists():
            tree = etree.parse(str(endnotes_path))
            root = tree.getroot()

            # Find the endnote element
            for en_elem in root.findall(f"{{{WORD_NAMESPACE}}}endnote"):
                if en_elem.get(f"{{{WORD_NAMESPACE}}}id") == note_id_str:
                    self._replace_note_content(en_elem, new_text)
                    break

            tree.write(
                str(endnotes_path),
                encoding="utf-8",
                xml_declaration=True,
                pretty_print=True,
            )

    def _replace_note_content(self, note_elem: etree._Element, new_text: str) -> None:
        """Replace the content of a note element with new text.

        Args:
            note_elem: The footnote or endnote XML element
            new_text: The new text content
        """
        # Remove all existing paragraphs
        for para in note_elem.findall(f"{{{WORD_NAMESPACE}}}p"):
            note_elem.remove(para)

        # Add new paragraph with text
        para = etree.SubElement(note_elem, f"{{{WORD_NAMESPACE}}}p")
        run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = new_text

    def _remove_footnote_reference(self, note_id: str) -> None:
        """Remove a footnote reference from the document body.

        Args:
            note_id: The footnote ID to remove reference for
        """
        # Find and remove footnoteReference elements with matching ID
        for ref in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}footnoteReference"):
            if ref.get(f"{{{WORD_NAMESPACE}}}id") == note_id:
                parent = ref.getparent()
                if parent is not None:
                    # Remove the containing run if it only contains the reference
                    if parent.tag == f"{{{WORD_NAMESPACE}}}r":
                        children = list(parent)
                        if len(children) == 1:
                            grandparent = parent.getparent()
                            if grandparent is not None:
                                grandparent.remove(parent)
                        else:
                            parent.remove(ref)
                    else:
                        parent.remove(ref)
                break

    def _remove_endnote_reference(self, note_id: str) -> None:
        """Remove an endnote reference from the document body.

        Args:
            note_id: The endnote ID to remove reference for
        """
        # Find and remove endnoteReference elements with matching ID
        for ref in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}endnoteReference"):
            if ref.get(f"{{{WORD_NAMESPACE}}}id") == note_id:
                parent = ref.getparent()
                if parent is not None:
                    # Remove the containing run if it only contains the reference
                    if parent.tag == f"{{{WORD_NAMESPACE}}}r":
                        children = list(parent)
                        if len(children) == 1:
                            grandparent = parent.getparent()
                            if grandparent is not None:
                                grandparent.remove(parent)
                        else:
                            parent.remove(ref)
                    else:
                        parent.remove(ref)
                break

    def _renumber_footnotes(self) -> None:
        """Renumber all footnotes sequentially starting from 1.

        This updates both the footnote IDs in footnotes.xml and the
        corresponding references in document.xml.

        Note: IDs -1 and 0 are reserved for separator footnotes and
        are never modified.
        """
        temp_dir = self._document._temp_dir
        if not temp_dir:
            return

        footnotes_path = temp_dir / "word" / "footnotes.xml"
        if not footnotes_path.exists():
            return

        tree = etree.parse(str(footnotes_path))
        root = tree.getroot()

        # Collect user footnotes (not separators)
        user_footnotes = []
        for fn_elem in root.findall(f"{{{WORD_NAMESPACE}}}footnote"):
            if fn_elem.get(f"{{{WORD_NAMESPACE}}}type") is None:
                old_id = fn_elem.get(f"{{{WORD_NAMESPACE}}}id")
                user_footnotes.append((fn_elem, old_id))

        # Sort by current ID (numeric order)
        user_footnotes.sort(key=lambda x: int(x[1]) if x[1] else 0)

        # Build mapping of old ID to new ID
        id_mapping: dict[str, str] = {}
        for new_id, (fn_elem, old_id) in enumerate(user_footnotes, start=1):
            if old_id and old_id != str(new_id):
                id_mapping[old_id] = str(new_id)
                fn_elem.set(f"{{{WORD_NAMESPACE}}}id", str(new_id))

        # Save footnotes.xml
        tree.write(
            str(footnotes_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

        # Update references in document.xml
        if id_mapping:
            for ref in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}footnoteReference"):
                old_id = ref.get(f"{{{WORD_NAMESPACE}}}id")
                if old_id in id_mapping:
                    ref.set(f"{{{WORD_NAMESPACE}}}id", id_mapping[old_id])

    def _renumber_endnotes(self) -> None:
        """Renumber all endnotes sequentially starting from 1.

        This updates both the endnote IDs in endnotes.xml and the
        corresponding references in document.xml.

        Note: IDs -1 and 0 are reserved for separator endnotes and
        are never modified.
        """
        temp_dir = self._document._temp_dir
        if not temp_dir:
            return

        endnotes_path = temp_dir / "word" / "endnotes.xml"
        if not endnotes_path.exists():
            return

        tree = etree.parse(str(endnotes_path))
        root = tree.getroot()

        # Collect user endnotes (not separators)
        user_endnotes = []
        for en_elem in root.findall(f"{{{WORD_NAMESPACE}}}endnote"):
            if en_elem.get(f"{{{WORD_NAMESPACE}}}type") is None:
                old_id = en_elem.get(f"{{{WORD_NAMESPACE}}}id")
                user_endnotes.append((en_elem, old_id))

        # Sort by current ID (numeric order)
        user_endnotes.sort(key=lambda x: int(x[1]) if x[1] else 0)

        # Build mapping of old ID to new ID
        id_mapping: dict[str, str] = {}
        for new_id, (en_elem, old_id) in enumerate(user_endnotes, start=1):
            if old_id and old_id != str(new_id):
                id_mapping[old_id] = str(new_id)
                en_elem.set(f"{{{WORD_NAMESPACE}}}id", str(new_id))

        # Save endnotes.xml
        tree.write(
            str(endnotes_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

        # Update references in document.xml
        if id_mapping:
            for ref in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}endnoteReference"):
                old_id = ref.get(f"{{{WORD_NAMESPACE}}}id")
                if old_id in id_mapping:
                    ref.set(f"{{{WORD_NAMESPACE}}}id", id_mapping[old_id])

    def insert_footnote(
        self,
        text: str | list[str],
        at: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Insert a footnote reference at specific text location.

        Supports rich content including multiple paragraphs and markdown formatting.

        Args:
            text: The footnote text content. Can be:
                - A string for single paragraph
                - A list of strings for multiple paragraphs
                Supports markdown: **bold**, *italic*, ++underline++, ~~strikethrough~~
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
            >>> doc.insert_footnote(
            ...     ["First paragraph.", "Second paragraph with **bold**."],
            ...     at="citation needed"
            ... )
        """
        if not self._document._is_zip or not self._document._temp_dir:
            raise ValueError("Cannot add footnotes to non-ZIP documents")

        author_name = author if author is not None else self._document.author

        # Find location for footnote reference
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(at, paragraphs)

        if not matches:
            scope_str = str(scope) if scope is not None and not isinstance(scope, str) else scope
            raise TextNotFoundError(at, scope_str)

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
        text: str | list[str],
        at: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Insert an endnote reference at specific text location.

        Supports rich content including multiple paragraphs and markdown formatting.

        Args:
            text: The endnote text content. Can be:
                - A string for single paragraph
                - A list of strings for multiple paragraphs
                Supports markdown: **bold**, *italic*, ++underline++, ~~strikethrough~~
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
            >>> doc.insert_endnote(
            ...     ["First paragraph.", "Second paragraph with *italic*."],
            ...     at="see notes"
            ... )
        """
        if not self._document._is_zip or not self._document._temp_dir:
            raise ValueError("Cannot add endnotes to non-ZIP documents")

        author_name = author if author is not None else self._document.author

        # Find location for endnote reference
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(at, paragraphs)

        if not matches:
            scope_str = str(scope) if scope is not None and not isinstance(scope, str) else scope
            raise TextNotFoundError(at, scope_str)

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
        if temp_dir is None:
            return 1
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
        if temp_dir is None:
            return 1
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

    def _add_footnote_to_xml(self, footnote_id: int, text: str | list[str], author: str) -> None:
        """Add a footnote to footnotes.xml, creating the file if needed.

        Args:
            footnote_id: The footnote ID
            text: Footnote text content (string or list of strings for multi-paragraph).
                  Supports markdown formatting.
            author: Author name (for tracking if needed)
        """
        temp_dir = self._document._temp_dir
        if temp_dir is None:
            return
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
            self._ensure_footnote_styles()

        # Create footnote element
        footnote_elem = etree.SubElement(footnotes_root, f"{{{WORD_NAMESPACE}}}footnote")
        footnote_elem.set(f"{{{WORD_NAMESPACE}}}id", str(footnote_id))

        # Add paragraphs with rich content
        self._create_note_content(footnote_elem, text, note_type="footnote")

        # Write footnotes.xml
        footnotes_tree.write(
            str(footnotes_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _add_endnote_to_xml(self, endnote_id: int, text: str | list[str], author: str) -> None:
        """Add an endnote to endnotes.xml, creating the file if needed.

        Args:
            endnote_id: The endnote ID
            text: Endnote text content (string or list of strings for multi-paragraph).
                  Supports markdown formatting.
            author: Author name (for tracking if needed)
        """
        temp_dir = self._document._temp_dir
        if temp_dir is None:
            return
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
            self._ensure_endnote_styles()

        # Create endnote element
        endnote_elem = etree.SubElement(endnotes_root, f"{{{WORD_NAMESPACE}}}endnote")
        endnote_elem.set(f"{{{WORD_NAMESPACE}}}id", str(endnote_id))

        # Add paragraphs with rich content
        self._create_note_content(endnote_elem, text, note_type="endnote")

        # Write endnotes.xml
        endnotes_tree.write(
            str(endnotes_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _create_note_content(
        self,
        note_elem: etree._Element,
        text: str | list[str],
        note_type: str = "footnote",
    ) -> None:
        """Create rich content for a footnote or endnote element.

        Creates one or more paragraphs with formatted runs based on the text input.
        Follows Word's structure:
        - Paragraph has FootnoteText/EndnoteText style
        - First run contains footnoteRef/endnoteRef with reference style
        - Content follows with a leading space

        Supports markdown formatting: **bold**, *italic*, ++underline++, ~~strikethrough~~

        Args:
            note_elem: The footnote or endnote XML element to populate
            text: Content as a single string or list of strings for multiple paragraphs
            note_type: Either "footnote" or "endnote" (default: "footnote")
        """
        # Normalize to list of paragraphs
        paragraphs = [text] if isinstance(text, str) else text

        # Determine style names based on note type
        if note_type == "footnote":
            para_style = "FootnoteText"
            ref_style = "FootnoteReference"
            ref_tag = "footnoteRef"
        else:
            para_style = "EndnoteText"
            ref_style = "EndnoteReference"
            ref_tag = "endnoteRef"

        for i, para_text in enumerate(paragraphs):
            para = etree.SubElement(note_elem, f"{{{WORD_NAMESPACE}}}p")

            # Add paragraph properties with FootnoteText/EndnoteText style
            ppr = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}pPr")
            pstyle = etree.SubElement(ppr, f"{{{WORD_NAMESPACE}}}pStyle")
            pstyle.set(f"{{{WORD_NAMESPACE}}}val", para_style)

            # First paragraph gets the footnoteRef/endnoteRef marker
            if i == 0:
                # Create run with footnoteRef/endnoteRef (shows the note number)
                ref_run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
                ref_rpr = etree.SubElement(ref_run, f"{{{WORD_NAMESPACE}}}rPr")
                ref_rstyle = etree.SubElement(ref_rpr, f"{{{WORD_NAMESPACE}}}rStyle")
                ref_rstyle.set(f"{{{WORD_NAMESPACE}}}val", ref_style)
                etree.SubElement(ref_run, f"{{{WORD_NAMESPACE}}}{ref_tag}")

                # Add space after footnoteRef in separate unformatted run (Word convention)
                space_run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
                space_t = etree.SubElement(space_run, f"{{{WORD_NAMESPACE}}}t")
                space_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                space_t.text = " "

            # Parse markdown and create runs for content
            segments = parse_markdown(para_text)

            for segment in segments:
                run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")

                # Add run properties if any formatting is applied
                if segment.has_formatting():
                    rpr = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}rPr")
                    if segment.bold:
                        etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}b")
                    if segment.italic:
                        etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}i")
                    if segment.underline:
                        u_elem = etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}u")
                        u_elem.set(f"{{{WORD_NAMESPACE}}}val", "single")
                    if segment.strikethrough:
                        etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}strike")

                # Handle linebreak segments
                if segment.is_linebreak:
                    etree.SubElement(run, f"{{{WORD_NAMESPACE}}}br")
                else:
                    t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                    # Preserve whitespace if needed
                    if segment.text and (segment.text[0].isspace() or segment.text[-1].isspace()):
                        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t.text = segment.text

    def _insert_footnote_reference(self, match: TextSpan, footnote_id: int) -> None:
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

        # Add run properties with FootnoteReference style (for superscript formatting)
        rpr = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}rPr")
        rstyle = etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}rStyle")
        rstyle.set(f"{{{WORD_NAMESPACE}}}val", "FootnoteReference")

        # Add footnote reference
        footnote_ref = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}footnoteReference")
        footnote_ref.set(f"{{{WORD_NAMESPACE}}}id", str(footnote_id))

        # Insert the new run after the last run of the match
        parent = end_run.getparent()
        index = list(parent).index(end_run)
        parent.insert(index + 1, new_run)

    def _insert_endnote_reference(self, match: TextSpan, endnote_id: int) -> None:
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

        # Add run properties with EndnoteReference style (for superscript formatting)
        rpr = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}rPr")
        rstyle = etree.SubElement(rpr, f"{{{WORD_NAMESPACE}}}rStyle")
        rstyle.set(f"{{{WORD_NAMESPACE}}}val", "EndnoteReference")

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

    def _ensure_footnote_styles(self) -> None:
        """Ensure required footnote styles exist in the document.

        Creates FootnoteReference, FootnoteText, and FootnoteTextChar styles
        if they don't already exist. These styles are required for proper
        footnote display in Word (superscript numbers, proper text formatting).

        Also ensures the styles.xml relationship and content type exist.
        """
        from ..style_templates import ensure_standard_styles

        ensure_standard_styles(
            self._document.styles,
            "FootnoteReference",
            "FootnoteText",
            "FootnoteTextChar",
        )
        # Ensure styles.xml is properly registered if it was created
        self._ensure_styles_relationship()
        self._ensure_styles_content_type()
        self._document.styles.save()

    def _ensure_endnote_styles(self) -> None:
        """Ensure required endnote styles exist in the document.

        Creates EndnoteReference, EndnoteText, and EndnoteTextChar styles
        if they don't already exist. These styles are required for proper
        endnote display in Word (superscript numbers, proper text formatting).

        Also ensures the styles.xml relationship and content type exist.
        """
        from ..style_templates import ensure_standard_styles

        ensure_standard_styles(
            self._document.styles,
            "EndnoteReference",
            "EndnoteText",
            "EndnoteTextChar",
        )
        # Ensure styles.xml is properly registered if it was created
        self._ensure_styles_relationship()
        self._ensure_styles_content_type()
        self._document.styles.save()

    def _ensure_styles_relationship(self) -> None:
        """Ensure styles.xml relationship exists in document.xml.rels."""
        package = self._document._package
        if not package:
            return

        rel_mgr = RelationshipManager(package, "word/document.xml")
        rel_mgr.add_relationship(RelationshipTypes.STYLES, "styles.xml")
        rel_mgr.save()

    def _ensure_styles_content_type(self) -> None:
        """Ensure styles.xml content type exists in [Content_Types].xml."""
        package = self._document._package
        if not package:
            return

        ct_mgr = ContentTypeManager(package)
        ct_mgr.add_override("/word/styles.xml", ContentTypes.STYLES)
        ct_mgr.save()

    def _find_footnote_reference(
        self, footnote_id: str | int
    ) -> tuple[etree._Element, etree._Element] | None:
        """Find a footnoteReference element and its containing run/paragraph.

        Args:
            footnote_id: The footnote ID to search for

        Returns:
            Tuple of (run_element, paragraph_element) if found, None otherwise
        """
        note_id_str = str(footnote_id)

        for para in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"):
            for run in para.iter(f"{{{WORD_NAMESPACE}}}r"):
                for ref in run.iter(f"{{{WORD_NAMESPACE}}}footnoteReference"):
                    if ref.get(f"{{{WORD_NAMESPACE}}}id") == note_id_str:
                        return (run, para)
        return None

    def _find_endnote_reference(
        self, endnote_id: str | int
    ) -> tuple[etree._Element, etree._Element] | None:
        """Find an endnoteReference element and its containing run/paragraph.

        Args:
            endnote_id: The endnote ID to search for

        Returns:
            Tuple of (run_element, paragraph_element) if found, None otherwise
        """
        note_id_str = str(endnote_id)

        for para in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"):
            for run in para.iter(f"{{{WORD_NAMESPACE}}}r"):
                for ref in run.iter(f"{{{WORD_NAMESPACE}}}endnoteReference"):
                    if ref.get(f"{{{WORD_NAMESPACE}}}id") == note_id_str:
                        return (run, para)
        return None

    def _calculate_position_in_paragraph(
        self, para_elem: etree._Element, run_elem: etree._Element
    ) -> int:
        """Calculate the character position of a run within a paragraph.

        Args:
            para_elem: The paragraph element
            run_elem: The run element containing the reference

        Returns:
            Character offset where the run starts in the paragraph text
        """
        position = 0

        for run in para_elem.iter(f"{{{WORD_NAMESPACE}}}r"):
            if run is run_elem:
                break
            # Count text in this run
            for text_elem in run.findall(f".//{{{WORD_NAMESPACE}}}t"):
                position += len(text_elem.text or "")
            for text_elem in run.findall(f".//{{{WORD_NAMESPACE}}}delText"):
                position += len(text_elem.text or "")

        return position

    def get_footnote_reference_location(self, note_id: str | int) -> FootnoteReference | None:
        """Get the location where a footnote is referenced in the document.

        Args:
            note_id: The footnote ID

        Returns:
            FootnoteReference with paragraph, run element, and position,
            or None if reference not found

        Example:
            >>> ref_loc = doc._note_ops.get_footnote_reference_location(1)
            >>> if ref_loc:
            ...     print(f"Referenced in: {ref_loc.paragraph.text[:50]}")
        """
        from ..models.footnote import FootnoteReference
        from ..models.paragraph import Paragraph

        result = self._find_footnote_reference(note_id)
        if result is None:
            return None

        run_elem, para_elem = result
        position = self._calculate_position_in_paragraph(para_elem, run_elem)

        return FootnoteReference(
            paragraph=Paragraph(para_elem),
            run_element=run_elem,
            position_in_paragraph=position,
        )

    def get_endnote_reference_location(self, note_id: str | int) -> FootnoteReference | None:
        """Get the location where an endnote is referenced in the document.

        Args:
            note_id: The endnote ID

        Returns:
            FootnoteReference with paragraph, run element, and position,
            or None if reference not found

        Example:
            >>> ref_loc = doc._note_ops.get_endnote_reference_location(1)
            >>> if ref_loc:
            ...     print(f"Referenced in: {ref_loc.paragraph.text[:50]}")
        """
        from ..models.footnote import FootnoteReference
        from ..models.paragraph import Paragraph

        result = self._find_endnote_reference(note_id)
        if result is None:
            return None

        run_elem, para_elem = result
        position = self._calculate_position_in_paragraph(para_elem, run_elem)

        return FootnoteReference(
            paragraph=Paragraph(para_elem),
            run_element=run_elem,
            position_in_paragraph=position,
        )

    # ==================== Tracked Changes in Footnotes/Endnotes ====================

    def _get_note_element(self, note_type: str, note_id: str | int) -> tuple[etree._Element, Path]:
        """Get the note element and its XML file path.

        Args:
            note_type: Either "footnote" or "endnote"
            note_id: The note ID

        Returns:
            Tuple of (note_element, xml_file_path)

        Raises:
            NoteNotFoundError: If note not found
            ValueError: For non-ZIP documents
        """

        note_id_str = str(note_id)
        temp_dir = self._document._temp_dir

        if not temp_dir:
            raise ValueError(f"Cannot access {note_type}s in non-ZIP documents")

        if note_type == "footnote":
            xml_path = temp_dir / "word" / "footnotes.xml"
            tag_name = "footnote"
        else:
            xml_path = temp_dir / "word" / "endnotes.xml"
            tag_name = "endnote"

        if not xml_path.exists():
            available: list[str] = []
            raise NoteNotFoundError(note_type, note_id_str, available)

        tree = etree.parse(str(xml_path))
        root = tree.getroot()

        # Find the note element
        for note_elem in root.findall(f"{{{WORD_NAMESPACE}}}{tag_name}"):
            if note_elem.get(f"{{{WORD_NAMESPACE}}}id") == note_id_str:
                return note_elem, xml_path

        # Note not found - get available IDs for error message
        available = []
        for note_elem in root.findall(f"{{{WORD_NAMESPACE}}}{tag_name}"):
            if note_elem.get(f"{{{WORD_NAMESPACE}}}type") is None:
                elem_id = note_elem.get(f"{{{WORD_NAMESPACE}}}id")
                if elem_id:
                    available.append(elem_id)

        raise NoteNotFoundError(note_type, note_id_str, available)

    def _get_note_paragraphs(self, note_type: str, note_id: str | int) -> list[Any]:
        """Get all paragraphs from a footnote or endnote for text search.

        Args:
            note_type: Either "footnote" or "endnote"
            note_id: The note ID

        Returns:
            List of paragraph elements from the note

        Raises:
            NoteNotFoundError: If note not found
        """
        note_elem, _ = self._get_note_element(note_type, note_id)
        return list(note_elem.findall(f"{{{WORD_NAMESPACE}}}p"))

    def _save_note_xml(self, note_type: str, note_id: str | int) -> None:
        """Save the modified note XML back to the file.

        Args:
            note_type: Either "footnote" or "endnote"
            note_id: The note ID
        """
        temp_dir = self._document._temp_dir
        if not temp_dir:
            return

        if note_type == "footnote":
            xml_path = temp_dir / "word" / "footnotes.xml"
        else:
            xml_path = temp_dir / "word" / "endnotes.xml"

        if xml_path.exists():
            tree = etree.parse(str(xml_path))
            tree.write(
                str(xml_path),
                encoding="utf-8",
                xml_declaration=True,
                pretty_print=True,
            )

    def insert_tracked_in_footnote(
        self,
        note_id: str | int,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
    ) -> None:
        """Insert text with tracked changes inside a footnote.

        This method searches for anchor text within the footnote and inserts
        new text as a tracked insertion (w:ins) either after or before it.

        Args:
            note_id: The footnote ID to edit
            text: The text to insert (supports markdown formatting)
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If both or neither of after/before specified
            NoteNotFoundError: If footnote not found
            TextNotFoundError: If anchor text not found in footnote
            AmbiguousTextError: If anchor text found multiple times

        Example:
            >>> doc.insert_tracked_in_footnote(1, " [updated]", after="citation")
        """
        self._insert_tracked_in_note("footnote", note_id, text, after, before, author)

    def insert_tracked_in_endnote(
        self,
        note_id: str | int,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
    ) -> None:
        """Insert text with tracked changes inside an endnote.

        This method searches for anchor text within the endnote and inserts
        new text as a tracked insertion (w:ins) either after or before it.

        Args:
            note_id: The endnote ID to edit
            text: The text to insert (supports markdown formatting)
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If both or neither of after/before specified
            NoteNotFoundError: If endnote not found
            TextNotFoundError: If anchor text not found in endnote
            AmbiguousTextError: If anchor text found multiple times

        Example:
            >>> doc.insert_tracked_in_endnote(1, " [see also]", after="reference")
        """
        self._insert_tracked_in_note("endnote", note_id, text, after, before, author)

    def _insert_tracked_in_note(
        self,
        note_type: str,
        note_id: str | int,
        text: str,
        after: str | None,
        before: str | None,
        author: str | None,
    ) -> None:
        """Internal implementation for tracked insertion in notes."""
        # Validate parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        anchor = after if after is not None else before
        insert_after = after is not None

        # Get note paragraphs for text search
        note_elem, xml_path = self._get_note_element(note_type, note_id)
        paragraphs = list(note_elem.findall(f"{{{WORD_NAMESPACE}}}p"))

        if not paragraphs:
            raise TextNotFoundError(anchor)  # type: ignore[arg-type]

        # Search for anchor text
        matches = self._document._text_search.find_text(anchor, paragraphs)  # type: ignore[arg-type]

        if not matches:
            raise TextNotFoundError(anchor)  # type: ignore[arg-type]

        if len(matches) > 1:
            raise AmbiguousTextError(anchor, matches)  # type: ignore[arg-type]

        match = matches[0]

        # Create tracked insertion XML
        insertion_xml = self._document._xml_generator.create_insertion(text, author)
        insertion_element = self._parse_xml_element(insertion_xml)

        # Insert at the match location
        if insert_after:
            self._insert_after_match(match, insertion_element)
        else:
            self._insert_before_match(match, insertion_element)

        # Save the modified XML
        tree = etree.ElementTree(note_elem.getparent())
        tree.write(
            str(xml_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def delete_tracked_in_footnote(
        self,
        note_id: str | int,
        text: str,
        author: str | None = None,
    ) -> None:
        """Delete text with tracked changes inside a footnote.

        This method searches for text within the footnote and marks it
        as a tracked deletion (w:del).

        Args:
            note_id: The footnote ID to edit
            text: The text to delete
            author: Optional author override (uses document author if None)

        Raises:
            NoteNotFoundError: If footnote not found
            TextNotFoundError: If text not found in footnote
            AmbiguousTextError: If text found multiple times

        Example:
            >>> doc.delete_tracked_in_footnote(1, "outdated reference")
        """
        self._delete_tracked_in_note("footnote", note_id, text, author)

    def delete_tracked_in_endnote(
        self,
        note_id: str | int,
        text: str,
        author: str | None = None,
    ) -> None:
        """Delete text with tracked changes inside an endnote.

        This method searches for text within the endnote and marks it
        as a tracked deletion (w:del).

        Args:
            note_id: The endnote ID to edit
            text: The text to delete
            author: Optional author override (uses document author if None)

        Raises:
            NoteNotFoundError: If endnote not found
            TextNotFoundError: If text not found in endnote
            AmbiguousTextError: If text found multiple times

        Example:
            >>> doc.delete_tracked_in_endnote(1, "obsolete citation")
        """
        self._delete_tracked_in_note("endnote", note_id, text, author)

    def _delete_tracked_in_note(
        self,
        note_type: str,
        note_id: str | int,
        text: str,
        author: str | None,
    ) -> None:
        """Internal implementation for tracked deletion in notes."""
        # Get note paragraphs for text search
        note_elem, xml_path = self._get_note_element(note_type, note_id)
        paragraphs = list(note_elem.findall(f"{{{WORD_NAMESPACE}}}p"))

        if not paragraphs:
            raise TextNotFoundError(text)

        # Search for text to delete
        matches = self._document._text_search.find_text(text, paragraphs)

        if not matches:
            raise TextNotFoundError(text)

        if len(matches) > 1:
            raise AmbiguousTextError(text, matches)

        match = matches[0]

        # Create tracked deletion XML
        deletion_xml = self._document._xml_generator.create_deletion(match.text, author)
        deletion_element = self._parse_xml_element(deletion_xml)

        # Replace the matched text with deletion
        self._replace_match_with_element(match, deletion_element)

        # Save the modified XML
        tree = etree.ElementTree(note_elem.getparent())
        tree.write(
            str(xml_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def replace_tracked_in_footnote(
        self,
        note_id: str | int,
        find: str,
        replace: str,
        author: str | None = None,
    ) -> None:
        """Replace text with tracked changes inside a footnote.

        This method searches for text within the footnote and replaces it
        showing both the deletion of old text and insertion of new text
        as tracked changes.

        Args:
            note_id: The footnote ID to edit
            find: The text to find and replace
            replace: The replacement text (supports markdown formatting)
            author: Optional author override (uses document author if None)

        Raises:
            NoteNotFoundError: If footnote not found
            TextNotFoundError: If find text not found in footnote
            AmbiguousTextError: If find text found multiple times

        Example:
            >>> doc.replace_tracked_in_footnote(1, "2020", "2024")
        """
        self._replace_tracked_in_note("footnote", note_id, find, replace, author)

    def replace_tracked_in_endnote(
        self,
        note_id: str | int,
        find: str,
        replace: str,
        author: str | None = None,
    ) -> None:
        """Replace text with tracked changes inside an endnote.

        This method searches for text within the endnote and replaces it
        showing both the deletion of old text and insertion of new text
        as tracked changes.

        Args:
            note_id: The endnote ID to edit
            find: The text to find and replace
            replace: The replacement text (supports markdown formatting)
            author: Optional author override (uses document author if None)

        Raises:
            NoteNotFoundError: If endnote not found
            TextNotFoundError: If find text not found in endnote
            AmbiguousTextError: If find text found multiple times

        Example:
            >>> doc.replace_tracked_in_endnote(1, "ibid", "op. cit.")
        """
        self._replace_tracked_in_note("endnote", note_id, find, replace, author)

    def _replace_tracked_in_note(
        self,
        note_type: str,
        note_id: str | int,
        find: str,
        replace: str,
        author: str | None,
    ) -> None:
        """Internal implementation for tracked replacement in notes."""
        # Get note paragraphs for text search
        note_elem, xml_path = self._get_note_element(note_type, note_id)
        paragraphs = list(note_elem.findall(f"{{{WORD_NAMESPACE}}}p"))

        if not paragraphs:
            raise TextNotFoundError(find)

        # Search for text to replace
        matches = self._document._text_search.find_text(find, paragraphs)

        if not matches:
            raise TextNotFoundError(find)

        if len(matches) > 1:
            raise AmbiguousTextError(find, matches)

        match = matches[0]

        # Create tracked deletion + insertion XML
        deletion_xml = self._document._xml_generator.create_deletion(match.text, author)
        insertion_xml = self._document._xml_generator.create_insertion(replace, author)

        # Parse both elements
        elements = self._parse_xml_elements(f"{deletion_xml}\n{insertion_xml}")

        # Replace the matched text with deletion + insertion
        self._replace_match_with_elements(match, elements)

        # Save the modified XML
        tree = etree.ElementTree(note_elem.getparent())
        tree.write(
            str(xml_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    # ==================== Helper Methods for XML Manipulation ====================

    def _parse_xml_element(self, xml_content: str) -> Any:
        """Parse XML content into a single lxml element.

        Args:
            xml_content: The XML string to parse

        Returns:
            Parsed lxml Element
        """
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {xml_content}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        return list(root)[0]

    def _parse_xml_elements(self, xml_content: str) -> list[Any]:
        """Parse XML content into multiple lxml elements.

        Args:
            xml_content: The XML string(s) to parse

        Returns:
            List of parsed lxml Elements
        """
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {xml_content}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        return list(root)

    def _insert_after_match(self, match: TextSpan, insertion_element: Any) -> None:
        """Insert XML element after a matched text span.

        Args:
            match: TextSpan object representing where to insert
            insertion_element: The lxml Element to insert
        """
        paragraph = match.paragraph
        end_run = match.runs[match.end_run_index]
        run_index = list(paragraph).index(end_run)
        paragraph.insert(run_index + 1, insertion_element)

    def _insert_before_match(self, match: TextSpan, insertion_element: Any) -> None:
        """Insert XML element before a matched text span.

        Args:
            match: TextSpan object representing where to insert
            insertion_element: The lxml Element to insert
        """
        paragraph = match.paragraph
        start_run = match.runs[match.start_run_index]
        run_index = list(paragraph).index(start_run)
        paragraph.insert(run_index, insertion_element)

    def _replace_match_with_element(self, match: TextSpan, replacement_element: Any) -> None:
        """Replace matched text with a single XML element.

        This handles text potentially spanning multiple runs.

        Args:
            match: TextSpan object representing the text to replace
            replacement_element: The lxml Element to insert in place
        """
        paragraph = match.paragraph

        if match.start_run_index == match.end_run_index:
            # Single run case
            run = match.runs[match.start_run_index]
            run_text = self._get_run_text(run)

            actual_parent = run.getparent()
            if actual_parent is None:
                actual_parent = paragraph

            if match.start_offset == 0 and match.end_offset == len(run_text):
                # Replace entire run
                try:
                    run_index = list(actual_parent).index(run)
                except ValueError:
                    run_index = list(paragraph).index(run)
                    actual_parent = paragraph
                actual_parent.remove(run)
                actual_parent.insert(run_index, replacement_element)
            else:
                # Partial run - split and replace
                self._split_and_replace_in_run(
                    paragraph, run, match.start_offset, match.end_offset, replacement_element
                )
        else:
            # Multiple runs
            start_run = match.runs[match.start_run_index]
            end_run = match.runs[match.end_run_index]

            actual_parent = start_run.getparent()
            if actual_parent is None:
                actual_parent = paragraph

            try:
                start_run_index = list(actual_parent).index(start_run)
            except ValueError:
                start_run_index = list(paragraph).index(start_run)
                actual_parent = paragraph

            first_run_text = self._get_run_text(start_run)
            before_text = first_run_text[: match.start_offset]

            last_run_text = self._get_run_text(end_run)
            after_text = last_run_text[match.end_offset :]

            # Remove all runs in the match
            for i in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[i]
                run_parent = run.getparent()
                if run_parent is not None and run in run_parent:
                    run_parent.remove(run)

            # Build and insert replacement elements
            new_elements = self._build_split_elements(
                start_run, before_text, after_text, [replacement_element]
            )

            for i, elem in enumerate(new_elements):
                actual_parent.insert(start_run_index + i, elem)

    def _replace_match_with_elements(
        self, match: TextSpan, replacement_elements: list[Any]
    ) -> None:
        """Replace matched text with multiple XML elements.

        Args:
            match: TextSpan object representing the text to replace
            replacement_elements: List of lxml Elements to insert
        """
        paragraph = match.paragraph

        if match.start_run_index == match.end_run_index:
            # Single run case
            run = match.runs[match.start_run_index]
            run_text = self._get_run_text(run)

            actual_parent = run.getparent()
            if actual_parent is None:
                actual_parent = paragraph

            if match.start_offset == 0 and match.end_offset == len(run_text):
                # Replace entire run
                try:
                    run_index = list(actual_parent).index(run)
                except ValueError:
                    run_index = list(paragraph).index(run)
                    actual_parent = paragraph
                actual_parent.remove(run)
                for i, elem in enumerate(replacement_elements):
                    actual_parent.insert(run_index + i, elem)
            else:
                # Partial run - split and replace
                self._split_and_replace_in_run_multiple(
                    paragraph, run, match.start_offset, match.end_offset, replacement_elements
                )
        else:
            # Multiple runs - same logic as single element but insert all
            start_run = match.runs[match.start_run_index]
            end_run = match.runs[match.end_run_index]

            actual_parent = start_run.getparent()
            if actual_parent is None:
                actual_parent = paragraph

            try:
                start_run_index = list(actual_parent).index(start_run)
            except ValueError:
                start_run_index = list(paragraph).index(start_run)
                actual_parent = paragraph

            first_run_text = self._get_run_text(start_run)
            before_text = first_run_text[: match.start_offset]

            last_run_text = self._get_run_text(end_run)
            after_text = last_run_text[match.end_offset :]

            # Remove all runs in the match
            for i in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[i]
                run_parent = run.getparent()
                if run_parent is not None and run in run_parent:
                    run_parent.remove(run)

            # Build and insert replacement elements
            new_elements = self._build_split_elements(
                start_run, before_text, after_text, replacement_elements
            )

            for i, elem in enumerate(new_elements):
                actual_parent.insert(start_run_index + i, elem)

    def _get_run_text(self, run: Any) -> str:
        """Extract text content from a run.

        Args:
            run: A w:r (run) Element

        Returns:
            Text content of the run
        """
        text_elements = run.findall(f".//{{{WORD_NAMESPACE}}}t")
        deltext_elements = run.findall(f".//{{{WORD_NAMESPACE}}}delText")
        all_text_elements = text_elements + deltext_elements
        return "".join(elem.text or "" for elem in all_text_elements)

    def _create_text_run(self, text: str, source_run: Any) -> Any:
        """Create a new text run with properties from source run.

        Args:
            text: The text content for the new run
            source_run: The run to copy properties from

        Returns:
            A new w:r element
        """
        new_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
        run_props = source_run.find(f"{{{WORD_NAMESPACE}}}rPr")
        if run_props is not None:
            new_run.append(etree.fromstring(etree.tostring(run_props)))
        text_elem = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}t")
        if text and (text[0].isspace() or text[-1].isspace()):
            text_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_elem.text = text
        return new_run

    def _build_split_elements(
        self, run: Any, before_text: str, after_text: str, replacement_elements: list[Any]
    ) -> list[Any]:
        """Build list of elements for a run split operation.

        Args:
            run: The original run being split
            before_text: Text before the replacement
            after_text: Text after the replacement
            replacement_elements: Elements to insert in the middle

        Returns:
            List of elements to insert
        """
        new_elements = []
        if before_text:
            new_elements.append(self._create_text_run(before_text, run))
        new_elements.extend(replacement_elements)
        if after_text:
            new_elements.append(self._create_text_run(after_text, run))
        return new_elements

    def _split_and_replace_in_run(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_element: Any,
    ) -> None:
        """Split a run and replace a portion with a new element.

        Args:
            paragraph: The paragraph containing the run
            run: The run to split
            start_offset: Character offset where match starts
            end_offset: Character offset where match ends
            replacement_element: Element to insert
        """
        run_text = self._get_run_text(run)
        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        new_elements = self._build_split_elements(
            run, before_text, after_text, [replacement_element]
        )
        self._replace_run_with_elements(paragraph, run, new_elements)

    def _split_and_replace_in_run_multiple(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_elements: list[Any],
    ) -> None:
        """Split a run and replace a portion with multiple elements.

        Args:
            paragraph: The paragraph containing the run
            run: The run to split
            start_offset: Character offset where match starts
            end_offset: Character offset where match ends
            replacement_elements: Elements to insert
        """
        run_text = self._get_run_text(run)
        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        new_elements = self._build_split_elements(
            run, before_text, after_text, replacement_elements
        )
        self._replace_run_with_elements(paragraph, run, new_elements)

    def _replace_run_with_elements(self, paragraph: Any, run: Any, new_elements: list[Any]) -> None:
        """Replace a run with a list of new elements.

        Args:
            paragraph: The paragraph containing the run
            run: The run to replace
            new_elements: Elements to insert in place of the run
        """
        actual_parent = run.getparent()
        if actual_parent is None:
            actual_parent = paragraph

        try:
            run_index = list(actual_parent).index(run)
        except ValueError:
            run_index = list(paragraph).index(run)
            actual_parent = paragraph

        actual_parent.remove(run)
        for i, elem in enumerate(new_elements):
            actual_parent.insert(run_index + i, elem)
