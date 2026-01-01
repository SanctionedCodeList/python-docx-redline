"""
Bookmark, hyperlink, and cross-reference management for Word documents.

This module provides extraction and management of bookmarks, hyperlinks,
and cross-references, including bidirectional reference tracking and
broken link/reference detection.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from ..types import Element

from .types import (
    BookmarkInfo,
    CrossReferenceInfo,
    FieldType,
    HyperlinkInfo,
    LinkType,
    ReferenceValidationResult,
)


class BookmarkRegistry:
    """Registry for bookmarks and hyperlinks in a document.

    Provides:
    - Bookmark extraction and lookup
    - Hyperlink extraction (internal and external)
    - Bidirectional reference tracking
    - Broken link detection
    """

    def __init__(self) -> None:
        self.bookmarks: dict[str, BookmarkInfo] = {}
        self.hyperlinks: list[HyperlinkInfo] = []
        self._bookmark_id_counter = 0

    @classmethod
    def from_xml(
        cls,
        root: Element,
        relationships: dict[str, str] | None = None,
    ) -> BookmarkRegistry:
        """Create a BookmarkRegistry from document XML.

        Args:
            root: The document root element
            relationships: Mapping of relationship IDs to URLs

        Returns:
            BookmarkRegistry with extracted bookmarks and hyperlinks
        """
        registry = cls()
        registry._extract_bookmarks(root)
        registry._extract_hyperlinks(root, relationships or {})
        registry._resolve_references()
        return registry

    def _extract_bookmarks(self, root: Element) -> None:
        """Extract bookmarks from document XML."""
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        # Find all bookmarkStart elements
        for para_idx, para in enumerate(root.findall(".//w:p", ns)):
            for bookmark_start in para.findall(".//w:bookmarkStart", ns):
                name = bookmark_start.get(f"{{{ns['w']}}}name", "")
                bookmark_id = bookmark_start.get(f"{{{ns['w']}}}id", "")

                # Skip internal Word bookmarks (start with _)
                if name.startswith("_"):
                    continue

                # Get text content
                text_parts = []
                for t in para.findall(".//w:t", ns):
                    if t.text:
                        text_parts.append(t.text)
                text_preview = "".join(text_parts)[:100]

                # Find matching bookmarkEnd for span detection
                span_end_location = None
                end_id = bookmark_id
                for end_para_idx, end_para in enumerate(root.findall(".//w:p", ns)):
                    for bookmark_end in end_para.findall(".//w:bookmarkEnd", ns):
                        if bookmark_end.get(f"{{{ns['w']}}}id") == end_id:
                            if end_para_idx != para_idx:
                                span_end_location = f"p:{end_para_idx}"
                            break

                self.bookmarks[name] = BookmarkInfo(
                    name=name,
                    ref=f"bk:{name}",
                    location=f"p:{para_idx}",
                    bookmark_id=bookmark_id,
                    text_preview=text_preview,
                    span_end_location=span_end_location,
                    referenced_by=[],
                )

    def _extract_hyperlinks(
        self,
        root: Element,
        relationships: dict[str, str],
    ) -> None:
        """Extract hyperlinks from document XML."""
        ns = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }

        link_idx = 0
        for para_idx, para in enumerate(root.findall(".//w:p", ns)):
            for hyperlink in para.findall(".//w:hyperlink", ns):
                # Get link text
                text_parts = []
                for t in hyperlink.findall(".//w:t", ns):
                    if t.text:
                        text_parts.append(t.text)
                text = "".join(text_parts)

                # Check if internal (anchor) or external (r:id)
                anchor = hyperlink.get(f"{{{ns['w']}}}anchor")
                rel_id = hyperlink.get(f"{{{ns['r']}}}id")

                if anchor:
                    # Internal link to bookmark
                    target_bookmark = self.bookmarks.get(anchor)
                    target_location = target_bookmark.location if target_bookmark else None
                    is_broken = target_bookmark is None
                    error = "Bookmark not found" if is_broken else None

                    self.hyperlinks.append(
                        HyperlinkInfo(
                            ref=f"lnk:{link_idx}",
                            text=text,
                            link_type=LinkType.INTERNAL,
                            anchor=anchor,
                            target_location=target_location,
                            from_location=f"p:{para_idx}",
                            is_broken=is_broken,
                            error=error,
                        )
                    )
                elif rel_id:
                    # External link
                    target = relationships.get(rel_id)
                    is_broken = target is None
                    error = "Relationship not found" if is_broken else None

                    self.hyperlinks.append(
                        HyperlinkInfo(
                            ref=f"lnk:{link_idx}",
                            text=text,
                            link_type=LinkType.EXTERNAL,
                            target=target,
                            relationship_id=rel_id,
                            from_location=f"p:{para_idx}",
                            is_broken=is_broken,
                            error=error,
                        )
                    )

                link_idx += 1

    def _resolve_references(self) -> None:
        """Build bidirectional reference tracking."""
        for link in self.hyperlinks:
            if link.link_type == LinkType.INTERNAL and link.anchor:
                bookmark = self.bookmarks.get(link.anchor)
                if bookmark:
                    bookmark.referenced_by.append(link.ref)

    def get_bookmark(self, name: str) -> BookmarkInfo | None:
        """Get a bookmark by name."""
        return self.bookmarks.get(name)

    def get_orphan_bookmarks(self) -> list[BookmarkInfo]:
        """Get bookmarks that have no references to them."""
        return [bk for bk in self.bookmarks.values() if not bk.referenced_by]

    def get_broken_links(self) -> list[HyperlinkInfo]:
        """Get all broken hyperlinks."""
        return [link for link in self.hyperlinks if link.is_broken]

    def get_internal_links(self) -> list[HyperlinkInfo]:
        """Get all internal links (to bookmarks)."""
        return [link for link in self.hyperlinks if link.link_type == LinkType.INTERNAL]

    def get_external_links(self) -> list[HyperlinkInfo]:
        """Get all external links (URLs)."""
        return [link for link in self.hyperlinks if link.link_type == LinkType.EXTERNAL]

    def validate_references(self) -> ReferenceValidationResult:
        """Validate all references in the document."""
        broken_links = self.get_broken_links()
        orphan_bookmarks = self.get_orphan_bookmarks()

        missing_bookmarks = set()
        for link in broken_links:
            if link.link_type == LinkType.INTERNAL and link.anchor:
                missing_bookmarks.add(link.anchor)

        warnings = []
        if broken_links:
            warnings.append(f"Found {len(broken_links)} broken link(s)")
        if orphan_bookmarks:
            warnings.append(f"Found {len(orphan_bookmarks)} orphan bookmark(s)")

        return ReferenceValidationResult(
            is_valid=len(broken_links) == 0,
            broken_links=broken_links,
            orphan_bookmarks=orphan_bookmarks,
            missing_bookmarks=list(missing_bookmarks),
            warnings=warnings,
        )

    def to_yaml_dict(self) -> dict:
        """Convert to YAML-serializable dictionary."""
        result = {}

        if self.bookmarks:
            result["bookmarks"] = [
                {
                    "ref": bk.ref,
                    "name": bk.name,
                    "location": bk.location,
                    "text_preview": bk.text_preview,
                    **({"referenced_by": bk.referenced_by} if bk.referenced_by else {}),
                    **({"span_end": bk.span_end_location} if bk.span_end_location else {}),
                }
                for bk in self.bookmarks.values()
            ]

        if self.hyperlinks:
            internal = [
                lnk
                for lnk in self.hyperlinks
                if lnk.link_type == LinkType.INTERNAL and not lnk.is_broken
            ]
            external = [
                lnk
                for lnk in self.hyperlinks
                if lnk.link_type == LinkType.EXTERNAL and not lnk.is_broken
            ]
            broken = [lnk for lnk in self.hyperlinks if lnk.is_broken]

            links = {}
            if internal:
                links["internal"] = [
                    {
                        "ref": lnk.ref,
                        "from": lnk.from_location,
                        "to": f"bk:{lnk.anchor}",
                        "text": lnk.text,
                    }
                    for lnk in internal
                ]
            if external:
                links["external"] = [
                    {
                        "ref": lnk.ref,
                        "from": lnk.from_location,
                        "target": lnk.target,
                        "text": lnk.text,
                    }
                    for lnk in external
                ]
            if broken:
                links["broken"] = [
                    {
                        "ref": lnk.ref,
                        "from": lnk.from_location,
                        "error": lnk.error,
                        "text": lnk.text,
                    }
                    for lnk in broken
                ]

            if links:
                result["links"] = links

        return result


def add_bookmark(
    root: Element,
    name: str,
    location: str,
) -> BookmarkInfo:
    """Add a bookmark to the document.

    Args:
        root: The document root element
        name: The bookmark name
        location: The paragraph ref (e.g., "p:0")

    Returns:
        The created BookmarkInfo

    Raises:
        ValueError: If bookmark name already exists or location is invalid
    """
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    # Check for existing bookmark
    for bookmark_start in root.findall(f".//{w}bookmarkStart", ns):
        if bookmark_start.get(f"{w}name") == name:
            raise ValueError(f"Bookmark '{name}' already exists")

    # Parse location
    if not location.startswith("p:"):
        raise ValueError(f"Invalid location format: {location}")

    para_idx = int(location.split(":")[1])
    paragraphs = root.findall(f".//{w}p", ns)

    if para_idx >= len(paragraphs):
        raise ValueError(f"Paragraph at {location} not found")

    para = paragraphs[para_idx]

    # Find max bookmark ID
    max_id = -1
    for bk in root.findall(f".//{w}bookmarkStart", ns):
        try:
            bk_id = int(bk.get(f"{w}id", "0"))
            max_id = max(max_id, bk_id)
        except ValueError:
            pass

    new_id = str(max_id + 1)

    # Create bookmark elements
    bookmark_start = etree.Element(f"{w}bookmarkStart")
    bookmark_start.set(f"{w}id", new_id)
    bookmark_start.set(f"{w}name", name)

    bookmark_end = etree.Element(f"{w}bookmarkEnd")
    bookmark_end.set(f"{w}id", new_id)

    # Insert at beginning of paragraph
    para.insert(0, bookmark_start)
    para.append(bookmark_end)

    # Get text preview
    text_parts = []
    for t in para.findall(f".//{w}t", ns):
        if t.text:
            text_parts.append(t.text)
    text_preview = "".join(text_parts)[:100]

    return BookmarkInfo(
        name=name,
        ref=f"bk:{name}",
        location=location,
        bookmark_id=new_id,
        text_preview=text_preview,
        span_end_location=None,
        referenced_by=[],
    )


def rename_bookmark(
    root: Element,
    old_name: str,
    new_name: str,
) -> int:
    """Rename a bookmark and update all references to it.

    Args:
        root: The document root element
        old_name: The current bookmark name
        new_name: The new bookmark name

    Returns:
        Number of hyperlink references updated

    Raises:
        ValueError: If old_name not found or new_name already exists
    """
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    # Find the bookmark
    found = False
    for bookmark_start in root.findall(f".//{w}bookmarkStart", ns):
        name = bookmark_start.get(f"{w}name")
        if name == new_name:
            raise ValueError(f"Bookmark '{new_name}' already exists")
        if name == old_name:
            found = True

    if not found:
        raise ValueError(f"Bookmark '{old_name}' not found")

    # Rename the bookmark
    for bookmark_start in root.findall(f".//{w}bookmarkStart", ns):
        if bookmark_start.get(f"{w}name") == old_name:
            bookmark_start.set(f"{w}name", new_name)
            break

    # Update hyperlink references
    updated = 0
    for hyperlink in root.findall(f".//{w}hyperlink", ns):
        anchor = hyperlink.get(f"{w}anchor")
        if anchor == old_name:
            hyperlink.set(f"{w}anchor", new_name)
            updated += 1

    return updated


class CrossReferenceRegistry:
    """Registry for cross-references (field codes) in a document.

    Provides:
    - Cross-reference extraction (REF, PAGEREF, NOTEREF fields)
    - Field code parsing with switch detection
    - Broken cross-reference detection
    - Bidirectional reference tracking with bookmarks
    """

    # Regex to parse field instructions like "REF _Ref123456 \\h \\r"
    _FIELD_PATTERN = re.compile(r"^\s*(REF|PAGEREF|NOTEREF)\s+(\S+)\s*(.*?)\s*$", re.IGNORECASE)

    def __init__(self) -> None:
        self.cross_references: list[CrossReferenceInfo] = []
        self._xref_counter = 0

    @classmethod
    def from_xml(
        cls,
        root: Element,
        bookmark_registry: BookmarkRegistry | None = None,
    ) -> CrossReferenceRegistry:
        """Create a CrossReferenceRegistry from document XML.

        Args:
            root: The document root element
            bookmark_registry: Optional BookmarkRegistry for reference resolution

        Returns:
            CrossReferenceRegistry with extracted cross-references
        """
        registry = cls()
        registry._extract_cross_references(root, bookmark_registry)
        return registry

    def _extract_cross_references(
        self,
        root: Element,
        bookmark_registry: BookmarkRegistry | None,
    ) -> None:
        """Extract cross-references from document XML.

        Handles both simple fields (w:fldSimple) and complex fields
        (w:fldChar with w:instrText).
        """
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        w_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

        # Track paragraph indices
        paragraphs = root.findall(".//w:p", ns)
        para_map = {id(p): idx for idx, p in enumerate(paragraphs)}

        # Extract from simple fields (w:fldSimple)
        for fld_simple in root.findall(".//w:fldSimple", ns):
            instr = fld_simple.get(f"{w_ns}instr", "")
            xref = self._parse_field_instruction(instr)
            if xref:
                # Get display value
                text_parts = []
                for t in fld_simple.findall(".//w:t", ns):
                    if t.text:
                        text_parts.append(t.text)
                display_value = "".join(text_parts)

                # Find containing paragraph
                para = fld_simple
                while para is not None and para.tag != f"{w_ns}p":
                    para = para.getparent()

                para_idx = para_map.get(id(para), 0) if para is not None else 0

                # Check dirty flag
                is_dirty = fld_simple.get(f"{w_ns}dirty", "").lower() == "true"

                # Resolve target
                target_location = None
                is_broken = True
                error = None
                if bookmark_registry:
                    bookmark = bookmark_registry.get_bookmark(xref["target"])
                    if bookmark:
                        target_location = bookmark.location
                        is_broken = False
                        bookmark.referenced_by.append(f"xref:{self._xref_counter}")
                    else:
                        error = f"Bookmark '{xref['target']}' not found"

                self.cross_references.append(
                    CrossReferenceInfo(
                        ref=f"xref:{self._xref_counter}",
                        field_type=xref["field_type"],
                        target_bookmark=xref["target"],
                        from_location=f"p:{para_idx}",
                        display_value=display_value,
                        is_dirty=is_dirty,
                        is_hyperlink=xref["is_hyperlink"],
                        show_position=xref["show_position"],
                        number_format=xref["number_format"],
                        switches=xref["switches"],
                        target_location=target_location,
                        is_broken=is_broken,
                        error=error,
                    )
                )
                self._xref_counter += 1

        # Extract from complex fields (w:fldChar begin/separate/end with w:instrText)
        self._extract_complex_fields(root, ns, w_ns, para_map, bookmark_registry)

    def _extract_complex_fields(
        self,
        root: Element,
        ns: dict,
        w_ns: str,
        para_map: dict,
        bookmark_registry: BookmarkRegistry | None,
    ) -> None:
        """Extract cross-references from complex field constructs."""
        # Complex fields are structured as:
        # <w:fldChar w:fldCharType="begin"/>
        # <w:instrText>REF _Ref123 \h</w:instrText>
        # <w:fldChar w:fldCharType="separate"/>
        # <w:t>display value</w:t>
        # <w:fldChar w:fldCharType="end"/>

        in_field = False
        field_instr = []
        field_display = []
        field_start_para = None
        is_dirty = False

        for elem in root.iter():
            if elem.tag == f"{w_ns}fldChar":
                fld_type = elem.get(f"{w_ns}fldCharType", "")
                if fld_type == "begin":
                    in_field = True
                    field_instr = []
                    field_display = []
                    is_dirty = elem.get(f"{w_ns}dirty", "").lower() == "true"
                    # Find containing paragraph
                    para = elem
                    while para is not None and para.tag != f"{w_ns}p":
                        para = para.getparent()
                    field_start_para = para
                elif fld_type == "separate":
                    # Switch from collecting instruction to collecting display
                    pass
                elif fld_type == "end":
                    if in_field and field_instr:
                        instr_text = "".join(field_instr)
                        xref = self._parse_field_instruction(instr_text)
                        if xref:
                            display_value = "".join(field_display)
                            para_idx = (
                                para_map.get(id(field_start_para), 0) if field_start_para else 0
                            )

                            # Resolve target
                            target_location = None
                            is_broken = True
                            error = None
                            if bookmark_registry:
                                bookmark = bookmark_registry.get_bookmark(xref["target"])
                                if bookmark:
                                    target_location = bookmark.location
                                    is_broken = False
                                    bookmark.referenced_by.append(f"xref:{self._xref_counter}")
                                else:
                                    error = f"Bookmark '{xref['target']}' not found"

                            self.cross_references.append(
                                CrossReferenceInfo(
                                    ref=f"xref:{self._xref_counter}",
                                    field_type=xref["field_type"],
                                    target_bookmark=xref["target"],
                                    from_location=f"p:{para_idx}",
                                    display_value=display_value,
                                    is_dirty=is_dirty,
                                    is_hyperlink=xref["is_hyperlink"],
                                    show_position=xref["show_position"],
                                    number_format=xref["number_format"],
                                    switches=xref["switches"],
                                    target_location=target_location,
                                    is_broken=is_broken,
                                    error=error,
                                )
                            )
                            self._xref_counter += 1

                    in_field = False
                    field_instr = []
                    field_display = []
                    field_start_para = None

            elif in_field:
                if elem.tag == f"{w_ns}instrText" and elem.text:
                    field_instr.append(elem.text)
                elif elem.tag == f"{w_ns}t" and elem.text:
                    # Only collect display text after separate
                    field_display.append(elem.text)

    def _parse_field_instruction(self, instr: str) -> dict | None:
        """Parse a field instruction string.

        Args:
            instr: Field instruction like "REF _Ref123456 \\h \\r"

        Returns:
            Dict with field_type, target, switches, etc. or None if not a cross-reference
        """
        match = self._FIELD_PATTERN.match(instr)
        if not match:
            return None

        field_type_str = match.group(1).upper()
        target = match.group(2)
        switches = match.group(3).strip()

        # Map to FieldType enum
        field_type_map = {
            "REF": FieldType.REF,
            "PAGEREF": FieldType.PAGEREF,
            "NOTEREF": FieldType.NOTEREF,
        }
        field_type = field_type_map.get(field_type_str)
        if not field_type:
            return None

        # Parse switches
        is_hyperlink = "\\h" in switches.lower()
        show_position = "\\p" in switches.lower()

        # Determine number format
        number_format = None
        if "\\w" in switches.lower():
            number_format = "full"
        elif "\\r" in switches.lower():
            number_format = "relative"
        elif "\\n" in switches.lower():
            number_format = "no_context"

        return {
            "field_type": field_type,
            "target": target,
            "switches": switches,
            "is_hyperlink": is_hyperlink,
            "show_position": show_position,
            "number_format": number_format,
        }

    def get_all(self) -> list[CrossReferenceInfo]:
        """Get all cross-references."""
        return self.cross_references

    def get_broken(self) -> list[CrossReferenceInfo]:
        """Get all broken cross-references."""
        return [xref for xref in self.cross_references if xref.is_broken]

    def get_by_target(self, bookmark_name: str) -> list[CrossReferenceInfo]:
        """Get all cross-references to a specific bookmark."""
        return [xref for xref in self.cross_references if xref.target_bookmark == bookmark_name]

    def get_dirty(self) -> list[CrossReferenceInfo]:
        """Get all cross-references marked as dirty (need update)."""
        return [xref for xref in self.cross_references if xref.is_dirty]

    def to_yaml_dict(self) -> dict:
        """Convert to YAML-serializable dictionary."""
        if not self.cross_references:
            return {}

        result = {"cross_references": []}

        for xref in self.cross_references:
            xref_dict = {
                "ref": xref.ref,
                "type": xref.field_type.name,
                "target": xref.target_bookmark,
                "from": xref.from_location,
            }

            if xref.display_value:
                xref_dict["display"] = xref.display_value

            if xref.target_location:
                xref_dict["target_location"] = xref.target_location

            if xref.is_hyperlink:
                xref_dict["is_hyperlink"] = True

            if xref.show_position:
                xref_dict["show_position"] = True

            if xref.number_format:
                xref_dict["number_format"] = xref.number_format

            if xref.is_dirty:
                xref_dict["is_dirty"] = True

            if xref.is_broken:
                xref_dict["is_broken"] = True
                if xref.error:
                    xref_dict["error"] = xref.error

            result["cross_references"].append(xref_dict)

        return result
