"""
XML generation for tracked changes in Word documents.

This module provides the TrackedXMLGenerator class which automatically generates
proper OOXML for tracked insertions and deletions with all required attributes.
"""

import random
from datetime import datetime, timezone
from typing import Any


class TrackedXMLGenerator:
    """Generates OOXML for tracked changes with auto-managed attributes.

    This class handles the complex task of generating valid <w:ins> and <w:del>
    XML elements with all required attributes:
    - Auto-incrementing change IDs
    - ISO 8601 timestamps
    - RSID (Revision Save ID)
    - Author information
    - xml:space preservation for leading/trailing whitespace
    """

    def __init__(
        self,
        doc: Any | None = None,
        author: str = "Claude",
        rsid: str | None = None,
    ) -> None:
        """Initialize the XML generator.

        Args:
            doc: Optional document object to extract settings from
            author: Author name for tracked changes (default: "Claude")
            rsid: Revision Save ID - 8 hex characters (auto-generated if None)
        """
        self.doc = doc
        self.author = author if doc is None else getattr(doc, "author", author)
        self.rsid = rsid if rsid else self._generate_rsid()

        # Start change IDs from max existing + 1, or 1 if no doc provided
        if doc is not None:
            self.next_change_id = self._get_max_change_id(doc) + 1
        else:
            self.next_change_id = 1

    def create_insertion(self, text: str, author: str | None = None) -> str:
        """Generate <w:ins> XML for a tracked insertion.

        Args:
            text: The text to insert
            author: Override author (uses default if None)

        Returns:
            Complete OOXML string for the insertion
        """
        author = author if author is not None else self.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self.next_change_id
        self.next_change_id += 1

        # Handle xml:space for leading/trailing whitespace
        xml_space = (
            ' xml:space="preserve"' if (text and (text[0].isspace() or text[-1].isspace())) else ""
        )

        # Escape XML special characters
        escaped_text = self._escape_xml(text)

        # Generate the OOXML
        xml = (
            f'<w:ins w:id="{change_id}" w:author="{author}" '
            f'w:date="{timestamp}" w16du:dateUtc="{timestamp}">\n'
            f'  <w:r w:rsidR="{self.rsid}">\n'
            f"    <w:t{xml_space}>{escaped_text}</w:t>\n"
            f"  </w:r>\n"
            f"</w:ins>"
        )

        return xml

    def create_deletion(self, text: str, author: str | None = None) -> str:
        """Generate <w:del> XML for a tracked deletion.

        Args:
            text: The text being deleted
            author: Override author (uses default if None)

        Returns:
            Complete OOXML string for the deletion
        """
        author = author if author is not None else self.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self.next_change_id
        self.next_change_id += 1

        # Handle xml:space for leading/trailing whitespace
        xml_space = (
            ' xml:space="preserve"' if (text and (text[0].isspace() or text[-1].isspace())) else ""
        )

        # Escape XML special characters
        escaped_text = self._escape_xml(text)

        # Generate the OOXML
        # Note: deletions use <w:delText> instead of <w:t>
        xml = (
            f'<w:del w:id="{change_id}" w:author="{author}" '
            f'w:date="{timestamp}" w16du:dateUtc="{timestamp}">\n'
            f'  <w:r w:rsidDel="{self.rsid}">\n'
            f"    <w:delText{xml_space}>{escaped_text}</w:delText>\n"
            f"  </w:r>\n"
            f"</w:del>"
        )

        return xml

    @staticmethod
    def _escape_xml(text: str) -> str:
        """Escape XML special characters.

        Args:
            text: Raw text to escape

        Returns:
            XML-safe text
        """
        return (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;")
        )

    @staticmethod
    def _generate_rsid() -> str:
        """Generate a random 8-character hex RSID.

        Returns:
            8-character hex string (e.g., "F3F4F4B4")
        """
        return "".join(random.choices("0123456789ABCDEF", k=8))

    @staticmethod
    def _get_max_change_id(doc: Any) -> int:
        """Find the maximum change ID in the document.

        Args:
            doc: Document object with parsed XML

        Returns:
            Maximum change ID found, or 0 if none exist
        """
        # This will be implemented when we have the Document class
        # For now, return 0 to start from ID 1
        # TODO: Scan document.xml for all w:id attributes on w:ins/w:del elements
        return 0
