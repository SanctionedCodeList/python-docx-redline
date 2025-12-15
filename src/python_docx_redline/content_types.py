"""
ContentTypeManager class for managing [Content_Types].xml in OOXML packages.

This module provides a clean abstraction for managing content type declarations
in an OOXML package. Content types define the MIME type for each part in the
package, such as comments.xml, footnotes.xml, etc.
"""

import logging

from lxml import etree

from .constants import CONTENT_TYPES_NAMESPACE
from .package import OOXMLPackage

logger = logging.getLogger(__name__)


class ContentTypeManager:
    """Manages [Content_Types].xml in OOXML packages.

    This class handles the low-level operations of:
    - Reading the content types file
    - Adding new Override entries for package parts
    - Removing Override entries
    - Persisting changes back to the package

    Content types in OOXML use two mechanisms:
    - Default: Maps file extensions to content types (e.g., .xml -> application/xml)
    - Override: Maps specific part names to content types (e.g., /word/comments.xml)

    This class focuses on Override management since that's what's needed for
    adding new parts like comments, footnotes, etc.

    Example:
        >>> ct_mgr = ContentTypeManager(package)
        >>> ct_mgr.add_override(
        ...     "/word/comments.xml",
        ...     "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
        ... )
        >>> ct_mgr.save()

    Attributes:
        package: The OOXMLPackage containing this content types file
    """

    def __init__(self, package: OOXMLPackage) -> None:
        """Initialize a ContentTypeManager for a package.

        Args:
            package: The OOXMLPackage containing the [Content_Types].xml file
        """
        self._package = package
        self._content_types_path = package.temp_dir / "[Content_Types].xml"
        self._root: etree._Element | None = None
        self._tree: etree._ElementTree | None = None
        self._modified = False

    def _ensure_loaded(self) -> None:
        """Ensure the content types XML is loaded into memory."""
        if self._root is not None:
            return

        if self._content_types_path.exists():
            parser = etree.XMLParser(remove_blank_text=False)
            self._tree = etree.parse(str(self._content_types_path), parser)
            self._root = self._tree.getroot()
        else:
            # Create new content types file structure (shouldn't happen for valid docx)
            self._root = etree.Element(
                f"{{{CONTENT_TYPES_NAMESPACE}}}Types",
                nsmap={None: CONTENT_TYPES_NAMESPACE},
            )
            self._tree = etree.ElementTree(self._root)
            self._modified = True

    def get_content_type(self, part_name: str) -> str | None:
        """Get the content type for a specific part.

        Args:
            part_name: The part name to look up (e.g., "/word/comments.xml")

        Returns:
            The content type string if found, None otherwise
        """
        self._ensure_loaded()
        assert self._root is not None

        for override in self._root:
            if override.tag == f"{{{CONTENT_TYPES_NAMESPACE}}}Override":
                if override.get("PartName") == part_name:
                    return override.get("ContentType")

        return None

    def has_override(self, part_name: str) -> bool:
        """Check if an override exists for the given part name.

        Args:
            part_name: The part name to check (e.g., "/word/comments.xml")

        Returns:
            True if an override exists for this part
        """
        return self.get_content_type(part_name) is not None

    def add_override(self, part_name: str, content_type: str) -> bool:
        """Add a content type override for a part.

        If an override already exists for the part, this is a no-op.

        Args:
            part_name: The part name (e.g., "/word/comments.xml")
            content_type: The content type (e.g., "application/...comments+xml")

        Returns:
            True if a new override was added, False if it already existed
        """
        self._ensure_loaded()
        assert self._root is not None

        # Check if override already exists
        if self.has_override(part_name):
            logger.debug(f"Content type override already exists for {part_name}")
            return False

        # Add the new override
        override = etree.SubElement(self._root, f"{{{CONTENT_TYPES_NAMESPACE}}}Override")
        override.set("PartName", part_name)
        override.set("ContentType", content_type)

        self._modified = True
        logger.debug(f"Added content type override: {part_name} -> {content_type}")

        return True

    def remove_override(self, part_name: str) -> bool:
        """Remove a content type override by part name.

        Args:
            part_name: The part name to remove (e.g., "/word/comments.xml")

        Returns:
            True if an override was removed, False if not found
        """
        self._ensure_loaded()
        assert self._root is not None

        for override in list(self._root):
            if override.tag == f"{{{CONTENT_TYPES_NAMESPACE}}}Override":
                if override.get("PartName") == part_name:
                    self._root.remove(override)
                    self._modified = True
                    logger.debug(f"Removed content type override: {part_name}")
                    return True

        return False

    def remove_overrides(self, part_names: list[str]) -> int:
        """Remove multiple content type overrides by part name.

        Args:
            part_names: List of part names to remove

        Returns:
            Number of overrides removed
        """
        self._ensure_loaded()
        assert self._root is not None

        removed = 0
        part_names_set = set(part_names)

        for override in list(self._root):
            if override.tag == f"{{{CONTENT_TYPES_NAMESPACE}}}Override":
                part_name = override.get("PartName")
                if part_name in part_names_set:
                    self._root.remove(override)
                    removed += 1
                    logger.debug(f"Removed content type override: {part_name}")

        if removed > 0:
            self._modified = True

        return removed

    def save(self) -> None:
        """Persist changes to the [Content_Types].xml file.

        Only writes if modifications were made.
        """
        if not self._modified or self._tree is None:
            return

        # Write the file
        self._tree.write(
            str(self._content_types_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

        self._modified = False
        logger.debug(f"Saved content types file: {self._content_types_path}")

    @property
    def is_modified(self) -> bool:
        """Check if there are unsaved modifications."""
        return self._modified


# Common content type constants for convenience
class ContentTypes:
    """Common OOXML content type strings."""

    # Word document parts
    COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
    COMMENTS_EXTENDED = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
    )
    COMMENTS_IDS = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
    COMMENTS_EXTENSIBLE = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"
    )
    FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
    ENDNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
    DOCUMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
    SETTINGS = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
    NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
    FONT_TABLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
    HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
