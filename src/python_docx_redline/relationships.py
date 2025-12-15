"""
RelationshipManager class for managing .rels files in OOXML packages.

This module provides a clean abstraction for managing relationships between
parts in an OOXML package, such as the relationship between document.xml
and comments.xml, footnotes.xml, etc.
"""

import logging
from pathlib import Path

from lxml import etree

from .package import OOXMLPackage

logger = logging.getLogger(__name__)

# OOXML relationship namespace
RELS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"


class RelationshipManager:
    """Manages .rels files in OOXML packages.

    This class handles the low-level operations of:
    - Reading relationship files (.rels)
    - Adding new relationships with auto-generated IDs
    - Removing relationships by type
    - Persisting changes back to the package

    A relationship links one part to another using a unique ID (rId),
    a relationship type URI, and a target path.

    Example:
        >>> rel_mgr = RelationshipManager(package, "word/document.xml")
        >>> rel_id = rel_mgr.add_relationship(
        ...     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        ...     "comments.xml"
        ... )
        >>> print(f"Added comments relationship: {rel_id}")
        >>> rel_mgr.save()

    Attributes:
        package: The OOXMLPackage containing this relationship file
        part_name: The part this relationship file is for (e.g., "word/document.xml")
    """

    def __init__(self, package: OOXMLPackage, part_name: str) -> None:
        """Initialize a RelationshipManager for a specific part.

        Args:
            package: The OOXMLPackage containing the relationship file
            part_name: The part this relationship file is for.
                      For example, "word/document.xml" -> "word/_rels/document.xml.rels"
        """
        self._package = package
        self._part_name = part_name
        self._rels_path = self._compute_rels_path(part_name)
        self._root: etree._Element | None = None
        self._tree: etree._ElementTree | None = None
        self._modified = False

    def _compute_rels_path(self, part_name: str) -> Path:
        """Compute the .rels file path for a given part.

        For example:
        - "word/document.xml" -> "word/_rels/document.xml.rels"
        - "[Content_Types].xml" -> "_rels/[Content_Types].xml.rels"

        Args:
            part_name: The part name to compute rels path for

        Returns:
            Path to the .rels file
        """
        part_path = Path(part_name)
        parent = part_path.parent
        filename = part_path.name

        return self._package.temp_dir / parent / "_rels" / f"{filename}.rels"

    def _ensure_loaded(self) -> None:
        """Ensure the relationship XML is loaded into memory."""
        if self._root is not None:
            return

        if self._rels_path.exists():
            parser = etree.XMLParser(remove_blank_text=False)
            self._tree = etree.parse(str(self._rels_path), parser)
            self._root = self._tree.getroot()
        else:
            # Create new rels file structure
            self._root = etree.Element(
                f"{{{RELS_NAMESPACE}}}Relationships",
                nsmap={None: RELS_NAMESPACE},
            )
            self._tree = etree.ElementTree(self._root)
            self._modified = True

    def get_relationship(self, rel_type: str) -> str | None:
        """Get the relationship ID for a given type.

        Args:
            rel_type: The relationship type URI to search for

        Returns:
            The relationship ID (e.g., "rId3") if found, None otherwise
        """
        self._ensure_loaded()
        assert self._root is not None

        for rel in self._root:
            if rel.get("Type") == rel_type:
                return rel.get("Id")

        return None

    def get_relationship_target(self, rel_type: str) -> str | None:
        """Get the target path for a relationship type.

        Args:
            rel_type: The relationship type URI to search for

        Returns:
            The target path if found, None otherwise
        """
        self._ensure_loaded()
        assert self._root is not None

        for rel in self._root:
            if rel.get("Type") == rel_type:
                return rel.get("Target")

        return None

    def has_relationship(self, rel_type: str) -> bool:
        """Check if a relationship of the given type exists.

        Args:
            rel_type: The relationship type URI to check for

        Returns:
            True if a relationship of this type exists
        """
        return self.get_relationship(rel_type) is not None

    def add_relationship(self, rel_type: str, target: str) -> str:
        """Add a new relationship or return existing one.

        If a relationship of the given type already exists, returns its ID.
        Otherwise, creates a new relationship with an auto-generated ID.

        Args:
            rel_type: The relationship type URI
            target: The target path (relative to the part's directory)

        Returns:
            The relationship ID (e.g., "rId3")
        """
        self._ensure_loaded()
        assert self._root is not None

        # Check if relationship already exists
        existing_id = self.get_relationship(rel_type)
        if existing_id is not None:
            logger.debug(f"Relationship {rel_type} already exists: {existing_id}")
            return existing_id

        # Find next available rId
        next_id = self._next_available_id()

        # Add the new relationship
        rel_elem = etree.SubElement(self._root, f"{{{RELS_NAMESPACE}}}Relationship")
        rel_elem.set("Id", f"rId{next_id}")
        rel_elem.set("Type", rel_type)
        rel_elem.set("Target", target)

        self._modified = True
        logger.debug(f"Added relationship rId{next_id}: {rel_type} -> {target}")

        return f"rId{next_id}"

    def add_unique_relationship(self, rel_type: str, target: str) -> str:
        """Add a new relationship, always creating a new ID.

        Unlike add_relationship(), this method does not check for existing
        relationships of the same type. Use this for relationship types
        that can have multiple instances, like images.

        Args:
            rel_type: The relationship type URI
            target: The target path (relative to the part's directory)

        Returns:
            The new relationship ID (e.g., "rId3")
        """
        self._ensure_loaded()
        assert self._root is not None

        # Find next available rId
        next_id = self._next_available_id()

        # Add the new relationship
        rel_elem = etree.SubElement(self._root, f"{{{RELS_NAMESPACE}}}Relationship")
        rel_elem.set("Id", f"rId{next_id}")
        rel_elem.set("Type", rel_type)
        rel_elem.set("Target", target)

        self._modified = True
        logger.debug(f"Added unique relationship rId{next_id}: {rel_type} -> {target}")

        return f"rId{next_id}"

    def remove_relationship(self, rel_type: str) -> bool:
        """Remove a relationship by type.

        Args:
            rel_type: The relationship type URI to remove

        Returns:
            True if a relationship was removed, False if not found
        """
        self._ensure_loaded()
        assert self._root is not None

        for rel in list(self._root):
            if rel.get("Type") == rel_type:
                self._root.remove(rel)
                self._modified = True
                logger.debug(f"Removed relationship: {rel_type}")
                return True

        return False

    def remove_relationships(self, rel_types: list[str]) -> int:
        """Remove multiple relationships by type.

        Args:
            rel_types: List of relationship type URIs to remove

        Returns:
            Number of relationships removed
        """
        self._ensure_loaded()
        assert self._root is not None

        removed = 0
        rel_types_set = set(rel_types)

        for rel in list(self._root):
            rel_type = rel.get("Type")
            if rel_type in rel_types_set:
                self._root.remove(rel)
                removed += 1
                logger.debug(f"Removed relationship: {rel_type}")

        if removed > 0:
            self._modified = True

        return removed

    def _next_available_id(self) -> int:
        """Find the next available relationship ID number.

        Returns:
            The next available ID number (e.g., if rId1, rId2 exist, returns 3)
        """
        assert self._root is not None

        existing_ids: set[int] = set()
        for rel in self._root:
            rel_id = rel.get("Id", "")
            if rel_id.startswith("rId"):
                try:
                    existing_ids.add(int(rel_id[3:]))
                except ValueError:
                    pass

        # Find next available
        next_id = 1
        while next_id in existing_ids:
            next_id += 1

        return next_id

    def save(self) -> None:
        """Persist changes to the .rels file.

        Only writes if modifications were made. Creates the _rels
        directory if it doesn't exist.
        """
        if not self._modified or self._tree is None:
            return

        # Ensure parent directory exists
        self._rels_path.parent.mkdir(parents=True, exist_ok=True)

        # Write the file
        self._tree.write(
            str(self._rels_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

        self._modified = False
        logger.debug(f"Saved relationship file: {self._rels_path}")

    @property
    def is_modified(self) -> bool:
        """Check if there are unsaved modifications."""
        return self._modified


# Common relationship type constants for convenience
class RelationshipTypes:
    """Common OOXML relationship type URIs."""

    # Standard Office Open XML relationships
    COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    FONT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

    # Microsoft Office extensions
    COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
    COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
    COMMENTS_EXTENSIBLE = (
        "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
    )
