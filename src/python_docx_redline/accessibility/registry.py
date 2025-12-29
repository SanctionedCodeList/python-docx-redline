"""
Reference registry for resolving and managing document element refs.

This module provides the RefRegistry class which maintains mappings between
refs (like "p:5") and actual lxml elements, supporting both ordinal and
fingerprint-based resolution.
"""

from __future__ import annotations

import base64
import hashlib
from functools import lru_cache
from typing import TYPE_CHECKING
from weakref import WeakValueDictionary

from lxml import etree

from ..constants import WORD_NAMESPACE, w
from ..errors import RefNotFoundError, StaleRefError
from .types import ELEMENT_TYPE_TO_PREFIX, PREFIX_TO_ELEMENT_TYPE, ElementType, Ref

if TYPE_CHECKING:
    pass


class RefRegistry:
    """Registry for resolving refs to document elements and vice versa.

    The RefRegistry maintains mappings between ref paths and lxml elements,
    supporting both ordinal-based refs (p:5) and fingerprint-based refs (p:~xK4mNp2q).

    Key features:
    - Ordinal resolution: Maps element type + index to elements
    - Fingerprint generation: Creates stable content-based identifiers
    - Fingerprint resolution: Resolves fingerprint refs back to elements
    - Cache invalidation: Handles structural document changes

    Attributes:
        xml_root: Root element of the document XML

    Example:
        >>> registry = RefRegistry(doc.xml_root)
        >>> element = registry.resolve_ref("p:5")
        >>> ref = registry.get_ref(element)
        >>> ref.path
        'p:5'
    """

    # Tag names for each element type
    ELEMENT_TYPE_TAGS: dict[ElementType, str] = {
        ElementType.PARAGRAPH: w("p"),
        ElementType.RUN: w("r"),
        ElementType.TABLE: w("tbl"),
        ElementType.TABLE_ROW: w("tr"),
        ElementType.TABLE_CELL: w("tc"),
        ElementType.HEADER: w("hdr"),
        ElementType.FOOTER: w("ftr"),
        ElementType.INSERTION: w("ins"),
        ElementType.DELETION: w("del"),
    }

    # Reverse mapping from tag to element type
    TAG_TO_ELEMENT_TYPE: dict[str, ElementType] = {v: k for k, v in ELEMENT_TYPE_TAGS.items()}

    def __init__(self, xml_root: etree._Element) -> None:
        """Initialize the registry with a document root.

        Args:
            xml_root: Root element of the document XML tree
        """
        self.xml_root = xml_root

        # Caches for performance
        self._ordinal_cache: dict[str, etree._Element] = {}
        self._fingerprint_cache: dict[str, str] = {}  # fingerprint -> ref path
        self._element_to_ref: WeakValueDictionary[int, str] = WeakValueDictionary()

        # Version counter for cache invalidation
        self._version = 0

    def invalidate(self) -> None:
        """Invalidate all caches after structural changes.

        Call this method after making changes to the document structure
        (inserting/deleting paragraphs, tables, etc.) to ensure refs
        are recalculated correctly.
        """
        self._ordinal_cache.clear()
        self._fingerprint_cache.clear()
        self._element_to_ref.clear()
        self._version += 1

        # Clear LRU caches
        self._get_all_elements_by_type.cache_clear()

    @lru_cache(maxsize=32)
    def _get_all_elements_by_type(
        self, element_type: ElementType, version: int
    ) -> list[etree._Element]:
        """Get all elements of a given type in document order.

        Args:
            element_type: Type of element to find
            version: Cache version (for invalidation)

        Returns:
            List of matching elements in document order
        """
        tag = self.ELEMENT_TYPE_TAGS.get(element_type)
        if not tag:
            return []

        # Find the body element
        body = self.xml_root.find(f".//{w('body')}")
        if body is None:
            return []

        # For paragraphs, we need to handle nested paragraphs (in table cells)
        # We get top-level paragraphs only from body
        if element_type == ElementType.PARAGRAPH:
            # Get all paragraphs that are direct children of body
            # (not nested in tables, etc.)
            return list(body.findall(f"./{tag}"))

        # For tables, get direct children of body
        if element_type == ElementType.TABLE:
            return list(body.findall(f"./{tag}"))

        # For other elements, use recursive find
        return list(body.findall(f".//{tag}"))

    def resolve_ref(self, ref: str | Ref) -> etree._Element:
        """Resolve a ref to its corresponding XML element.

        Args:
            ref: Ref path or Ref object

        Returns:
            The lxml element referenced

        Raises:
            RefNotFoundError: If the ref cannot be resolved
            StaleRefError: If the ref points to a deleted element
        """
        if isinstance(ref, str):
            ref = Ref.parse(ref)

        segments = ref.segments
        if not segments:
            raise RefNotFoundError(ref.path, "Empty ref path")

        # Start resolution from the document body
        current_element = self.xml_root.find(f".//{w('body')}")
        if current_element is None:
            raise RefNotFoundError(ref.path, "Document body not found")

        # Process each segment
        for i, (prefix, identifier) in enumerate(segments):
            element_type = PREFIX_TO_ELEMENT_TYPE[prefix]

            # Determine the parent context for this segment
            if i == 0:
                # First segment - search in document body
                context = current_element
            else:
                # Subsequent segments - search in current element
                context = current_element

            # Resolve the identifier
            if identifier.startswith("~"):
                # Fingerprint-based resolution
                current_element = self._resolve_fingerprint(
                    context, element_type, identifier[1:], ref.path
                )
            else:
                # Ordinal-based resolution
                ordinal = int(identifier)
                current_element = self._resolve_ordinal(context, element_type, ordinal, ref.path)

        return current_element

    def _resolve_ordinal(
        self,
        context: etree._Element,
        element_type: ElementType,
        ordinal: int,
        ref_path: str,
    ) -> etree._Element:
        """Resolve an ordinal ref within a context.

        Args:
            context: Parent element to search within
            element_type: Type of element to find
            ordinal: 0-based index
            ref_path: Full ref path for error messages

        Returns:
            The element at the given ordinal

        Raises:
            RefNotFoundError: If the ordinal is out of bounds
        """
        tag = self.ELEMENT_TYPE_TAGS.get(element_type)
        if not tag:
            raise RefNotFoundError(ref_path, f"Unsupported element type: {element_type}")

        # Check if context is the body (top-level search)
        if context.tag == w("body"):
            elements = self._get_all_elements_by_type(element_type, self._version)
        else:
            # Search within the context
            elements = list(context.findall(f"./{tag}"))

        if ordinal < 0 or ordinal >= len(elements):
            msg = (
                f"Ordinal {ordinal} out of bounds "
                f"(found {len(elements)} elements of type {element_type.name})"
            )
            raise RefNotFoundError(ref_path, msg)

        return elements[ordinal]

    def _resolve_fingerprint(
        self,
        context: etree._Element,
        element_type: ElementType,
        fingerprint: str,
        ref_path: str,
    ) -> etree._Element:
        """Resolve a fingerprint ref within a context.

        Args:
            context: Parent element to search within
            element_type: Type of element to find
            fingerprint: Content fingerprint (without leading ~)
            ref_path: Full ref path for error messages

        Returns:
            The element matching the fingerprint

        Raises:
            RefNotFoundError: If no element matches the fingerprint
            StaleRefError: If the element was modified
        """
        tag = self.ELEMENT_TYPE_TAGS.get(element_type)
        if not tag:
            raise RefNotFoundError(ref_path, f"Unsupported element type: {element_type}")

        # Search for element with matching fingerprint
        if context.tag == w("body"):
            elements = self._get_all_elements_by_type(element_type, self._version)
        else:
            elements = list(context.findall(f".//{tag}"))

        for element in elements:
            if self._compute_fingerprint(element) == fingerprint:
                return element

        # Fingerprint not found - check if it's in our cache
        # If it was previously valid, this is a stale ref
        cache_key = f"{element_type.name}:~{fingerprint}"
        if cache_key in self._fingerprint_cache:
            raise StaleRefError(ref_path, "Element content has changed")

        raise RefNotFoundError(ref_path, f"No element found with fingerprint {fingerprint}")

    def get_ref(self, element: etree._Element, use_fingerprint: bool = False) -> Ref:
        """Get a ref for an element.

        Args:
            element: The lxml element
            use_fingerprint: If True, generate a fingerprint-based ref

        Returns:
            Ref object for the element

        Raises:
            RefNotFoundError: If the element type is not supported
        """
        # Determine element type
        element_type = self.TAG_TO_ELEMENT_TYPE.get(element.tag)
        if element_type is None:
            raise RefNotFoundError(
                "<unknown>",
                f"Unsupported element tag: {element.tag}",
            )

        # Build the ref path
        path_parts = []

        # Walk up the tree to build the full path
        current = element
        while current is not None:
            current_type = self.TAG_TO_ELEMENT_TYPE.get(current.tag)

            if current_type is not None:
                # Get ordinal or fingerprint for this element
                if use_fingerprint and current_type == element_type:
                    identifier = f"~{self._compute_fingerprint(current)}"
                else:
                    ordinal = self._get_ordinal(current, current_type)
                    identifier = str(ordinal)

                prefix = ELEMENT_TYPE_TO_PREFIX[current_type]
                path_parts.insert(0, f"{prefix}:{identifier}")

            # Move to parent
            parent = current.getparent()
            if parent is not None and parent.tag == w("body"):
                break
            current = parent

        if not path_parts:
            raise RefNotFoundError("<unknown>", "Could not determine ref path")

        return Ref(path="/".join(path_parts))

    def _get_ordinal(self, element: etree._Element, element_type: ElementType) -> int:
        """Get the ordinal index of an element among siblings of the same type.

        Args:
            element: The element to find
            element_type: Type of the element

        Returns:
            0-based ordinal index
        """
        parent = element.getparent()
        if parent is None:
            return 0

        # If parent is body, use document-level ordinals
        if parent.tag == w("body"):
            elements = self._get_all_elements_by_type(element_type, self._version)
        else:
            tag = self.ELEMENT_TYPE_TAGS[element_type]
            elements = list(parent.findall(f"./{tag}"))

        try:
            return elements.index(element)
        except ValueError:
            return 0

    def _compute_fingerprint(self, element: etree._Element) -> str:
        """Compute a content fingerprint for an element.

        The fingerprint is based on:
        - First 200 chars of text content
        - Element style (if any)
        - Parent element type

        Args:
            element: Element to fingerprint

        Returns:
            8-character base64 fingerprint
        """
        # Extract text content
        text_content = self._get_text_content(element)[:200]

        # Get style if available
        style = ""
        style_elem = element.find(f".//{w('pStyle')}")
        if style_elem is not None:
            style = style_elem.get(f"{{{WORD_NAMESPACE}}}val", "")

        # Get parent type
        parent = element.getparent()
        parent_type = parent.tag if parent is not None else ""

        # Combine and hash
        content = f"{text_content}|{style}|{parent_type}"
        hash_bytes = hashlib.sha256(content.encode("utf-8")).digest()

        # Take first 6 bytes for 8 base64 chars
        return base64.urlsafe_b64encode(hash_bytes[:6]).decode("ascii").rstrip("=")

    def _get_text_content(self, element: etree._Element) -> str:
        """Extract all text content from an element.

        Args:
            element: Element to extract text from

        Returns:
            Concatenated text content
        """
        text_parts = []

        # Find all text nodes
        for text_elem in element.iter(w("t")):
            if text_elem.text:
                text_parts.append(text_elem.text)

        return "".join(text_parts)

    def get_all_refs(self, element_type: ElementType) -> list[Ref]:
        """Get refs for all elements of a given type.

        Args:
            element_type: Type of elements to enumerate

        Returns:
            List of Ref objects in document order
        """
        elements = self._get_all_elements_by_type(element_type, self._version)
        prefix = ELEMENT_TYPE_TO_PREFIX[element_type]

        return [Ref(path=f"{prefix}:{i}") for i in range(len(elements))]

    def count_elements(self, element_type: ElementType) -> int:
        """Count elements of a given type.

        Args:
            element_type: Type to count

        Returns:
            Number of elements
        """
        return len(self._get_all_elements_by_type(element_type, self._version))
