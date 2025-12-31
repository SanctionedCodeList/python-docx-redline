"""
Bookmark and hyperlink management for Word documents.

This module provides extraction and management of bookmarks and hyperlinks,
including bidirectional reference tracking and broken link detection.

Note: This is a minimal stub implementation. Full implementation pending.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from ..types import Element

from .types import BookmarkInfo, HyperlinkInfo, LinkType, ReferenceValidationResult


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
