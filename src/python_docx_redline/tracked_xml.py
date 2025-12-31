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
            return f'  <w:r w:rsidR="{self.rsid}">\n    <w:br/>\n  </w:r>\n'

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

    def create_plain_run(
        self, text: str, source_run: etree._Element | None = None
    ) -> etree._Element:
        """Generate a plain <w:r> element without tracked change wrapper.

        This method creates a standard Word run element that can be used for
        untracked edits, preserving formatting from an optional source run.

        Args:
            text: The text content for the run
            source_run: Optional run element to copy formatting (w:rPr) from

        Returns:
            lxml Element for the run (<w:r>)

        Example:
            >>> gen = TrackedXMLGenerator(author="Editor")
            >>> run = gen.create_plain_run("new text")
            >>> # Returns <w:r><w:t>new text</w:t></w:r>
        """

        # Create the run element
        run = etree.Element(_w("r"))
        run.set(_w("rsidR"), self.rsid)

        # Copy run properties from source run if provided
        if source_run is not None:
            source_rpr = source_run.find(_w("rPr"))
            if source_rpr is not None:
                run.append(deepcopy(source_rpr))

        # Create the text element
        text_elem = etree.SubElement(run, _w("t"))

        # Handle xml:space for leading/trailing whitespace
        if text and (text[0].isspace() or text[-1].isspace()):
            text_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        text_elem.text = text

        return run

    def create_plain_runs(
        self, text: str, source_run: etree._Element | None = None
    ) -> list[etree._Element]:
        """Generate plain runs, handling markdown formatting.

        This method parses markdown-formatted text and creates multiple runs
        as needed to represent the formatting. Unlike create_insertion(), this
        does not wrap the runs in a <w:ins> element.

        Supports the same markdown as create_insertion:
        - **bold** -> <w:b/>
        - *italic* -> <w:i/>
        - ++underline++ -> <w:u/>
        - ~~strikethrough~~ -> <w:strike/>

        Args:
            text: The text content (may include markdown formatting)
            source_run: Optional run element to copy base formatting from

        Returns:
            List of lxml Elements for the runs

        Example:
            >>> gen = TrackedXMLGenerator(author="Editor")
            >>> runs = gen.create_plain_runs("This is **bold** text")
            >>> # Returns 3 runs: "This is ", "bold" (with <w:b/>), " text"
        """
        from python_docx_redline.markdown_parser import parse_markdown

        # Parse markdown to get formatted segments
        segments = parse_markdown(text)

        runs = []
        for segment in segments:
            run = self._create_plain_run_from_segment(segment, source_run)
            runs.append(run)

        return runs

    def _create_plain_run_from_segment(
        self,
        segment: "TextSegment",
        source_run: etree._Element | None = None,
    ) -> etree._Element:
        """Create a plain run element from a TextSegment.

        Args:
            segment: TextSegment with text and formatting flags
            source_run: Optional run element to copy base formatting from

        Returns:
            lxml Element for the run
        """
        # Handle linebreak segments - emit <w:br/> instead of <w:t>
        if segment.is_linebreak:
            run = etree.Element(_w("r"))
            run.set(_w("rsidR"), self.rsid)
            etree.SubElement(run, _w("br"))
            return run

        text = segment.text

        # Create the run element
        run = etree.Element(_w("r"))
        run.set(_w("rsidR"), self.rsid)

        # Build run properties
        rpr = None

        # First, copy base formatting from source run if provided
        if source_run is not None:
            source_rpr = source_run.find(_w("rPr"))
            if source_rpr is not None:
                rpr = deepcopy(source_rpr)
                run.append(rpr)

        # Then add markdown formatting on top
        if segment.has_formatting():
            if rpr is None:
                rpr = etree.SubElement(run, _w("rPr"))

            if segment.bold:
                # Only add if not already present
                if rpr.find(_w("b")) is None:
                    etree.SubElement(rpr, _w("b"))
            if segment.italic:
                if rpr.find(_w("i")) is None:
                    etree.SubElement(rpr, _w("i"))
            if segment.underline:
                if rpr.find(_w("u")) is None:
                    u_elem = etree.SubElement(rpr, _w("u"))
                    u_elem.set(_w("val"), "single")
            if segment.strikethrough:
                if rpr.find(_w("strike")) is None:
                    etree.SubElement(rpr, _w("strike"))

        # Create the text element
        text_elem = etree.SubElement(run, _w("t"))

        # Handle xml:space for leading/trailing whitespace
        if text and (text[0].isspace() or text[-1].isspace()):
            text_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        text_elem.text = text

        return run

    @staticmethod
    def _get_max_change_id(doc: Any) -> int:
        """Find the maximum change ID in the document.

        Scans document.xml for all w:id attributes on w:ins, w:del, w:moveFrom,
        w:moveTo, and w:pPrChange elements to find the maximum existing ID.

        Args:
            doc: Document object with parsed XML (must have xml_root attribute)

        Returns:
            Maximum change ID found, or 0 if none exist
        """
        max_id = 0

        # Get the XML root from the document
        xml_root = getattr(doc, "xml_root", None)
        if xml_root is None:
            return 0

        # Elements that use w:id for tracked changes
        change_tags = [
            _w("ins"),
            _w("del"),
            _w("moveFrom"),
            _w("moveTo"),
            _w("pPrChange"),
            _w("rPrChange"),
            _w("sectPrChange"),
            _w("tblPrChange"),
            _w("trPrChange"),
            _w("tcPrChange"),
            _w("customXmlInsRangeStart"),
            _w("customXmlDelRangeStart"),
            _w("customXmlMoveFromRangeStart"),
            _w("customXmlMoveToRangeStart"),
        ]

        # Find all elements with w:id attribute
        for tag in change_tags:
            for elem in xml_root.iter(tag):
                id_attr = elem.get(_w("id"))
                if id_attr is not None:
                    try:
                        id_val = int(id_attr)
                        if id_val > max_id:
                            max_id = id_val
                    except ValueError:
                        # Non-integer ID, skip
                        pass

        return max_id
