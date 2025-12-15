"""
XML generation for tracked changes in Word documents.

This module provides the TrackedXMLGenerator class which automatically generates
proper OOXML for tracked insertions and deletions with all required attributes.

Supports markdown formatting in inserted text:
- **bold** -> <w:b/>
- *italic* -> <w:i/>
- ++underline++ -> <w:u w:val="single"/>
- ~~strikethrough~~ -> <w:strike/>
"""

import random
from copy import deepcopy
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Any

from lxml import etree

from .constants import w as _w
from .constants import w15 as _w15

if TYPE_CHECKING:
    from python_docx_redline.author import AuthorIdentity
    from python_docx_redline.markdown_parser import TextSegment


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

        The text parameter supports markdown formatting:
        - **bold** -> bold text
        - *italic* -> italic text
        - ++underline++ -> underlined text
        - ~~strikethrough~~ -> strikethrough text

        Args:
            text: The text to insert (supports markdown formatting)
            author: Override author (uses default if None)

        Returns:
            Complete OOXML string for the insertion with MS365 identity if available
        """
        from python_docx_redline.markdown_parser import parse_markdown

        author = author if author is not None else self.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self.next_change_id
        self.next_change_id += 1

        # Build MS365 identity attributes if available
        identity_attrs = ""
        if self._author_identity:
            if self._author_identity.guid:
                identity_attrs += f' w15:userId="{self._author_identity.guid}"'
            identity_attrs += f' w15:providerId="{self._author_identity.provider_id}"'

        # Parse markdown to get formatted segments
        segments = parse_markdown(text)

        # Generate runs for each segment
        runs_xml = self._generate_runs(segments)

        # Generate the OOXML
        xml = (
            f'<w:ins w:id="{change_id}" w:author="{author}" '
            f'w:date="{timestamp}" w16du:dateUtc="{timestamp}"{identity_attrs}>\n'
            f"{runs_xml}"
            f"</w:ins>"
        )

        return xml

    def _generate_runs(self, segments: list["TextSegment"]) -> str:
        """Generate <w:r> elements for a list of text segments.

        Args:
            segments: List of TextSegment objects with formatting info

        Returns:
            XML string containing all runs
        """
        runs = []
        for segment in segments:
            run_xml = self._generate_run(segment)
            runs.append(run_xml)
        return "".join(runs)

    def _generate_run(self, segment: "TextSegment") -> str:
        """Generate a single <w:r> element for a text segment.

        Args:
            segment: TextSegment with text and formatting

        Returns:
            XML string for the run
        """
        # Handle linebreak segments - emit <w:br/> instead of <w:t>
        if segment.is_linebreak:
            return f'  <w:r w:rsidR="{self.rsid}">\n' f"    <w:br/>\n" f"  </w:r>\n"

        text = segment.text

        # Handle xml:space for leading/trailing whitespace
        xml_space = (
            ' xml:space="preserve"' if (text and (text[0].isspace() or text[-1].isspace())) else ""
        )

        # Escape XML special characters
        escaped_text = self._escape_xml(text)

        # Generate run properties if any formatting is applied
        rpr_xml = self._generate_run_properties(segment)

        # Build the run
        if rpr_xml:
            return (
                f'  <w:r w:rsidR="{self.rsid}">\n'
                f"{rpr_xml}"
                f"    <w:t{xml_space}>{escaped_text}</w:t>\n"
                f"  </w:r>\n"
            )
        else:
            return (
                f'  <w:r w:rsidR="{self.rsid}">\n'
                f"    <w:t{xml_space}>{escaped_text}</w:t>\n"
                f"  </w:r>\n"
            )

    def _generate_run_properties(self, segment: "TextSegment") -> str:
        """Generate <w:rPr> element for formatting.

        Args:
            segment: TextSegment with formatting flags

        Returns:
            XML string for run properties, or empty string if no formatting
        """
        if not segment.has_formatting():
            return ""

        props = []
        if segment.bold:
            props.append("      <w:b/>\n")
        if segment.italic:
            props.append("      <w:i/>\n")
        if segment.underline:
            props.append('      <w:u w:val="single"/>\n')
        if segment.strikethrough:
            props.append("      <w:strike/>\n")

        if props:
            return "    <w:rPr>\n" + "".join(props) + "    </w:rPr>\n"
        return ""

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

    def create_run_property_change(
        self,
        previous_rpr: etree._Element | None,
        author: str | None = None,
    ) -> tuple[etree._Element, int]:
        """Generate <w:rPrChange> element for tracking run property changes.

        This element should be appended as the last child of the current <w:rPr>.
        It stores the previous state of run properties before the change.

        Args:
            previous_rpr: The <w:rPr> element representing the previous state,
                         or None for empty previous state
            author: Override author (uses default if None)

        Returns:
            Tuple of (<w:rPrChange> element, change_id)

        Example:
            >>> change, change_id = generator.create_run_property_change(old_rpr)
            >>> current_rpr.append(change)
        """
        author = author if author is not None else self.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self.next_change_id
        self.next_change_id += 1

        # Create the rPrChange element
        rpr_change = etree.Element(_w("rPrChange"))
        rpr_change.set(_w("id"), str(change_id))
        rpr_change.set(_w("author"), author)
        rpr_change.set(_w("date"), timestamp)

        # Add MS365 identity attributes if available
        if self._author_identity:
            if self._author_identity.guid:
                rpr_change.set(_w15("userId"), self._author_identity.guid)
            rpr_change.set(_w15("providerId"), self._author_identity.provider_id)

        # Add the previous rPr state as a child
        if previous_rpr is not None:
            # Deep copy to avoid modifying the original
            prev_copy = deepcopy(previous_rpr)
            # Remove any existing rPrChange from the copy (shouldn't nest)
            for existing_change in prev_copy.findall(_w("rPrChange")):
                prev_copy.remove(existing_change)
            rpr_change.append(prev_copy)
        else:
            # Empty previous state
            rpr_change.append(etree.Element(_w("rPr")))

        return rpr_change, change_id

    def create_paragraph_property_change(
        self,
        previous_ppr: etree._Element | None,
        author: str | None = None,
    ) -> tuple[etree._Element, int]:
        """Generate <w:pPrChange> element for tracking paragraph property changes.

        This element should be appended as the last child of the current <w:pPr>.
        It stores the previous state of paragraph properties before the change.

        Args:
            previous_ppr: The <w:pPr> element representing the previous state,
                         or None for empty previous state
            author: Override author (uses default if None)

        Returns:
            Tuple of (<w:pPrChange> element, change_id)

        Example:
            >>> change, change_id = generator.create_paragraph_property_change(old_ppr)
            >>> current_ppr.append(change)
        """
        author = author if author is not None else self.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self.next_change_id
        self.next_change_id += 1

        # Create the pPrChange element
        ppr_change = etree.Element(_w("pPrChange"))
        ppr_change.set(_w("id"), str(change_id))
        ppr_change.set(_w("author"), author)
        ppr_change.set(_w("date"), timestamp)

        # Add MS365 identity attributes if available
        if self._author_identity:
            if self._author_identity.guid:
                ppr_change.set(_w15("userId"), self._author_identity.guid)
            ppr_change.set(_w15("providerId"), self._author_identity.provider_id)

        # Add the previous pPr state as a child
        if previous_ppr is not None:
            # Deep copy to avoid modifying the original
            prev_copy = deepcopy(previous_ppr)
            # Remove any existing pPrChange from the copy (shouldn't nest)
            for existing_change in prev_copy.findall(_w("pPrChange")):
                prev_copy.remove(existing_change)
            # Also remove rPr from pPr copy (run props tracked separately)
            for rpr in prev_copy.findall(_w("rPr")):
                prev_copy.remove(rpr)
            ppr_change.append(prev_copy)
        else:
            # Empty previous state
            ppr_change.append(etree.Element(_w("pPr")))

        return ppr_change, change_id

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
