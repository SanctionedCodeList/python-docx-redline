"""
XML generation for tracked changes in Word documents.

This module provides the TrackedXMLGenerator class which automatically generates
proper OOXML for tracked insertions and deletions with all required attributes.
"""

import random
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from python_docx_redline.author import AuthorIdentity


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

        # Check if document has MS365 identity
        self._author_identity: AuthorIdentity | None = None
        if doc is not None:
            self._author_identity = getattr(doc, "_author_identity", None)

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
            Complete OOXML string for the insertion with MS365 identity if available
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

        # Build MS365 identity attributes if available
        identity_attrs = ""
        if self._author_identity:
            if self._author_identity.guid:
                identity_attrs += f' w15:userId="{self._author_identity.guid}"'
            identity_attrs += f' w15:providerId="{self._author_identity.provider_id}"'

        # Generate the OOXML
        xml = (
            f'<w:ins w:id="{change_id}" w:author="{author}" '
            f'w:date="{timestamp}" w16du:dateUtc="{timestamp}"{identity_attrs}>\n'
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
            Complete OOXML string for the deletion with MS365 identity if available
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

        # Build MS365 identity attributes if available
        identity_attrs = ""
        if self._author_identity:
            if self._author_identity.guid:
                identity_attrs += f' w15:userId="{self._author_identity.guid}"'
            identity_attrs += f' w15:providerId="{self._author_identity.provider_id}"'

        # Generate the OOXML
        # Note: deletions use <w:delText> instead of <w:t>
        xml = (
            f'<w:del w:id="{change_id}" w:author="{author}" '
            f'w:date="{timestamp}" w16du:dateUtc="{timestamp}"{identity_attrs}>\n'
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

    def create_move_from(
        self,
        text: str,
        move_name: str,
        author: str | None = None,
    ) -> tuple[str, int, int]:
        """Generate moveFrom XML for the source location of a move.

        Creates the complete moveFrom container including:
        - moveFromRangeStart with unique ID and move name
        - moveFrom with the moved text
        - moveFromRangeEnd

        Args:
            text: The text being moved
            move_name: Name linking source to destination (e.g., "move1")
            author: Override author (uses default if None)

        Returns:
            Tuple of (XML string, range_id, move_id)
        """
        author = author if author is not None else self.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        # Generate unique IDs for range markers and move element
        range_id = self.next_change_id
        self.next_change_id += 1
        move_id = self.next_change_id
        self.next_change_id += 1

        # Handle xml:space for leading/trailing whitespace
        xml_space = (
            ' xml:space="preserve"' if (text and (text[0].isspace() or text[-1].isspace())) else ""
        )

        # Escape XML special characters
        escaped_text = self._escape_xml(text)

        # Build MS365 identity attributes if available
        identity_attrs = ""
        if self._author_identity:
            if self._author_identity.guid:
                identity_attrs += f' w15:userId="{self._author_identity.guid}"'
            identity_attrs += f' w15:providerId="{self._author_identity.provider_id}"'

        # Generate the OOXML for moveFrom container
        xml = (
            f'<w:moveFromRangeStart w:id="{range_id}" w:name="{move_name}" '
            f'w:author="{author}" w:date="{timestamp}"/>\n'
            f'<w:moveFrom w:id="{move_id}" w:author="{author}" '
            f'w:date="{timestamp}"{identity_attrs}>\n'
            f'  <w:r w:rsidDel="{self.rsid}">\n'
            f"    <w:delText{xml_space}>{escaped_text}</w:delText>\n"
            f"  </w:r>\n"
            f"</w:moveFrom>\n"
            f'<w:moveFromRangeEnd w:id="{range_id}"/>'
        )

        return xml, range_id, move_id

    def create_move_to(
        self,
        text: str,
        move_name: str,
        author: str | None = None,
    ) -> tuple[str, int, int]:
        """Generate moveTo XML for the destination location of a move.

        Creates the complete moveTo container including:
        - moveToRangeStart with unique ID and move name
        - moveTo with the moved text
        - moveToRangeEnd

        Args:
            text: The text being moved
            move_name: Name linking source to destination (must match moveFrom)
            author: Override author (uses default if None)

        Returns:
            Tuple of (XML string, range_id, move_id)
        """
        author = author if author is not None else self.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        # Generate unique IDs for range markers and move element
        range_id = self.next_change_id
        self.next_change_id += 1
        move_id = self.next_change_id
        self.next_change_id += 1

        # Handle xml:space for leading/trailing whitespace
        xml_space = (
            ' xml:space="preserve"' if (text and (text[0].isspace() or text[-1].isspace())) else ""
        )

        # Escape XML special characters
        escaped_text = self._escape_xml(text)

        # Build MS365 identity attributes if available
        identity_attrs = ""
        if self._author_identity:
            if self._author_identity.guid:
                identity_attrs += f' w15:userId="{self._author_identity.guid}"'
            identity_attrs += f' w15:providerId="{self._author_identity.provider_id}"'

        # Generate the OOXML for moveTo container
        xml = (
            f'<w:moveToRangeStart w:id="{range_id}" w:name="{move_name}" '
            f'w:author="{author}" w:date="{timestamp}"/>\n'
            f'<w:moveTo w:id="{move_id}" w:author="{author}" '
            f'w:date="{timestamp}"{identity_attrs}>\n'
            f'  <w:r w:rsidR="{self.rsid}">\n'
            f"    <w:t{xml_space}>{escaped_text}</w:t>\n"
            f"  </w:r>\n"
            f"</w:moveTo>\n"
            f'<w:moveToRangeEnd w:id="{range_id}"/>'
        )

        return xml, range_id, move_id

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
