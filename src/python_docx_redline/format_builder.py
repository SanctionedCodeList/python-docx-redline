"""
Property builders for format tracking in Word documents.

This module provides utilities for building, merging, and comparing
run properties (<w:rPr>) and paragraph properties (<w:pPr>) elements.
"""

from copy import deepcopy
from typing import Any

from lxml import etree

from .constants import WORD_NAMESPACE
from .constants import w as _w

# Unit conversion utilities


def points_to_twips(points: float) -> int:
    """Convert points to twips (1 point = 20 twips)."""
    return int(points * 20)


def twips_to_points(twips: int) -> float:
    """Convert twips to points."""
    return twips / 20


def inches_to_twips(inches: float) -> int:
    """Convert inches to twips (1 inch = 1440 twips)."""
    return int(inches * 1440)


def twips_to_inches(twips: int) -> float:
    """Convert twips to inches."""
    return twips / 1440


def points_to_half_points(points: float) -> int:
    """Convert points to half-points for font size (w:sz uses half-points)."""
    return int(points * 2)


def half_points_to_points(half_points: int) -> float:
    """Convert half-points to points."""
    return half_points / 2


class RunPropertyBuilder:
    """Build and manipulate <w:rPr> (run properties) elements.

    This class handles conversion between Python parameters and OOXML
    run property elements, including proper unit conversions.

    Supported properties:
        - bold: bool
        - italic: bool
        - underline: bool or str (style name)
        - strikethrough: bool
        - double_strikethrough: bool
        - font_name: str
        - font_size: float (in points)
        - color: str (hex "#RRGGBB" or "auto")
        - highlight: str (color name)
        - superscript: bool
        - subscript: bool
        - small_caps: bool
        - all_caps: bool
    """

    # Map Python property names to OOXML element names and value handling
    # Format: property_name -> (element_name, value_attr, converter_to_xml, converter_from_xml)
    PROPERTY_MAP: dict[str, tuple[str, str | None, Any, Any]] = {
        "bold": ("b", None, None, None),
        "italic": ("i", None, None, None),
        "underline": ("u", "val", lambda v: "single" if v is True else v, None),
        "strikethrough": ("strike", None, None, None),
        "double_strikethrough": ("dstrike", None, None, None),
        "font_size": ("sz", "val", points_to_half_points, half_points_to_points),
        "font_size_cs": ("szCs", "val", points_to_half_points, half_points_to_points),
        "color": ("color", "val", lambda v: v.lstrip("#") if v != "auto" else v, None),
        "highlight": ("highlight", "val", None, None),
        "small_caps": ("smallCaps", None, None, None),
        "all_caps": ("caps", None, None, None),
    }

    # Font properties need special handling (multiple attributes)
    FONT_ATTRS = ["ascii", "hAnsi", "cs", "eastAsia"]

    # Vertical alignment (superscript/subscript) needs special handling
    VERT_ALIGN_VALUES = {"superscript": "superscript", "subscript": "subscript"}

    @classmethod
    def build(cls, nsmap: dict[str, str] | None = None, **kwargs: Any) -> etree._Element:
        """Build a <w:rPr> element from keyword arguments.

        Args:
            nsmap: Namespace map to use (defaults to NSMAP)
            **kwargs: Property values to set

        Returns:
            New <w:rPr> element with specified properties

        Example:
            >>> rpr = RunPropertyBuilder.build(bold=True, font_size=14)
        """
        if nsmap is None:
            nsmap = {"w": WORD_NAMESPACE}

        rpr = etree.Element(_w("rPr"), nsmap=nsmap)
        cls._apply_properties(rpr, kwargs)
        return rpr

    @classmethod
    def merge(cls, base: etree._Element | None, updates: dict[str, Any]) -> etree._Element:
        """Merge updates into existing <w:rPr>, returning new element.

        Creates a deep copy of base (if provided) and applies updates.
        Does not modify the original element.

        Args:
            base: Existing <w:rPr> element (or None for empty)
            updates: Property values to update

        Returns:
            New <w:rPr> element with merged properties
        """
        if base is not None:
            rpr = deepcopy(base)
            # Remove any existing rPrChange from the copy
            for change in rpr.findall(_w("rPrChange")):
                rpr.remove(change)
        else:
            rpr = etree.Element(_w("rPr"))

        cls._apply_properties(rpr, updates)
        return rpr

    @classmethod
    def _apply_properties(cls, rpr: etree._Element, props: dict[str, Any]) -> None:
        """Apply property values to an <w:rPr> element (in-place).

        Args:
            rpr: The <w:rPr> element to modify
            props: Property values to apply
        """
        for prop_name, value in props.items():
            if value is None:
                continue

            # Handle font_name specially
            if prop_name == "font_name":
                cls._set_font_name(rpr, value)
                continue

            # Handle superscript/subscript specially
            if prop_name in ("superscript", "subscript"):
                cls._set_vert_align(rpr, prop_name if value else None)
                continue

            # Handle standard properties
            if prop_name in cls.PROPERTY_MAP:
                elem_name, val_attr, to_xml, _ = cls.PROPERTY_MAP[prop_name]

                # Find or create the element
                elem = rpr.find(_w(elem_name))

                if value is False:
                    # For boolean toggles, use w:val="0" to explicitly turn off
                    # This overrides style defaults (just removing element reverts to style)
                    if val_attr is None:
                        # Boolean property - set w:val="0" to explicitly disable
                        if elem is None:
                            elem = etree.SubElement(rpr, _w(elem_name))
                        elem.set(_w("val"), "0")
                    else:
                        # Value property - remove it entirely
                        if elem is not None:
                            rpr.remove(elem)
                else:
                    if elem is None:
                        elem = etree.SubElement(rpr, _w(elem_name))

                    if val_attr is not None:
                        # Apply converter if present
                        xml_val = to_xml(value) if to_xml else value
                        elem.set(_w(val_attr), str(xml_val))
                    else:
                        # Boolean property with True - remove any w:val="0"
                        if _w("val") in elem.attrib:
                            del elem.attrib[_w("val")]

                # Handle font_size_cs automatically when font_size is set
                if prop_name == "font_size" and value is not False:
                    cs_elem = rpr.find(_w("szCs"))
                    if cs_elem is None:
                        cs_elem = etree.SubElement(rpr, _w("szCs"))
                    xml_val = points_to_half_points(value)
                    cs_elem.set(_w("val"), str(xml_val))

    @classmethod
    def _set_font_name(cls, rpr: etree._Element, font_name: str | bool) -> None:
        """Set or remove font name in <w:rFonts>."""
        rfonts = rpr.find(_w("rFonts"))

        if font_name is False:
            if rfonts is not None:
                rpr.remove(rfonts)
            return

        if rfonts is None:
            rfonts = etree.SubElement(rpr, _w("rFonts"))

        for attr in cls.FONT_ATTRS:
            rfonts.set(_w(attr), font_name)

    @classmethod
    def _set_vert_align(cls, rpr: etree._Element, value: str | None) -> None:
        """Set or remove vertical alignment (superscript/subscript)."""
        vert_align = rpr.find(_w("vertAlign"))

        if value is None:
            if vert_align is not None:
                rpr.remove(vert_align)
            return

        if vert_align is None:
            vert_align = etree.SubElement(rpr, _w("vertAlign"))

        vert_align.set(_w("val"), cls.VERT_ALIGN_VALUES.get(value, value))

    @classmethod
    def extract(cls, rpr: etree._Element | None) -> dict[str, Any]:
        """Extract Python dict from <w:rPr> element.

        Args:
            rpr: The <w:rPr> element to extract from (or None)

        Returns:
            Dictionary of property values
        """
        if rpr is None:
            return {}

        result: dict[str, Any] = {}

        # Extract standard properties
        for prop_name, (elem_name, val_attr, _, from_xml) in cls.PROPERTY_MAP.items():
            elem = rpr.find(_w(elem_name))
            if elem is not None:
                if val_attr is not None:
                    raw_val = elem.get(_w(val_attr))
                    if raw_val is not None:
                        if from_xml:
                            result[prop_name] = from_xml(int(raw_val))
                        else:
                            result[prop_name] = raw_val
                else:
                    # Boolean property (presence = True)
                    # Check for explicit w:val="false" or w:val="0"
                    val = elem.get(_w("val"))
                    if val in ("false", "0"):
                        result[prop_name] = False
                    else:
                        result[prop_name] = True

        # Extract font name
        rfonts = rpr.find(_w("rFonts"))
        if rfonts is not None:
            font_name = rfonts.get(_w("ascii"))
            if font_name:
                result["font_name"] = font_name

        # Extract vertical alignment
        vert_align = rpr.find(_w("vertAlign"))
        if vert_align is not None:
            val = vert_align.get(_w("val"))
            if val == "superscript":
                result["superscript"] = True
            elif val == "subscript":
                result["subscript"] = True

        return result

    @classmethod
    def diff(
        cls, old: etree._Element | None, new: etree._Element | None
    ) -> dict[str, tuple[Any, Any]]:
        """Return dict of properties that differ between old and new.

        Args:
            old: Original <w:rPr> element (or None)
            new: New <w:rPr> element (or None)

        Returns:
            Dictionary mapping property names to (old_value, new_value) tuples
        """
        old_props = cls.extract(old)
        new_props = cls.extract(new)

        all_keys = set(old_props.keys()) | set(new_props.keys())
        diff_result: dict[str, tuple[Any, Any]] = {}

        for key in all_keys:
            old_val = old_props.get(key)
            new_val = new_props.get(key)
            if old_val != new_val:
                diff_result[key] = (old_val, new_val)

        return diff_result

    @classmethod
    def has_changes(cls, old: etree._Element | None, new: etree._Element | None) -> bool:
        """Check if there are any property differences.

        Args:
            old: Original <w:rPr> element (or None)
            new: New <w:rPr> element (or None)

        Returns:
            True if properties differ, False if identical
        """
        return bool(cls.diff(old, new))


class ParagraphPropertyBuilder:
    """Build and manipulate <w:pPr> (paragraph properties) elements.

    This class handles conversion between Python parameters and OOXML
    paragraph property elements, including proper unit conversions.

    Supported properties:
        - alignment: str ("left", "center", "right", "justify", "both")
        - spacing_before: float (in points)
        - spacing_after: float (in points)
        - line_spacing: float (multiplier, e.g., 1.5 for 1.5x)
        - indent_left: float (in inches)
        - indent_right: float (in inches)
        - indent_first_line: float (in inches)
        - indent_hanging: float (in inches)
    """

    # OOXML schema requires pPr child elements in this specific order
    # See ISO/IEC 29500-1:2016 section 17.3.1.26 (pPr)
    PPR_ELEMENT_ORDER = [
        "pStyle",
        "keepNext",
        "keepLines",
        "pageBreakBefore",
        "framePr",
        "widowControl",
        "numPr",
        "suppressLineNumbers",
        "pBdr",
        "shd",
        "tabs",
        "suppressAutoHyphens",
        "kinsoku",
        "wordWrap",
        "overflowPunct",
        "topLinePunct",
        "autoSpaceDE",
        "autoSpaceDN",
        "bidi",
        "adjustRightInd",
        "snapToGrid",
        "spacing",  # Must come before jc
        "ind",
        "contextualSpacing",
        "mirrorIndents",
        "suppressOverlap",
        "jc",  # Alignment
        "textDirection",
        "textAlignment",
        "textboxTightWrap",
        "outlineLvl",
        "divId",
        "cnfStyle",
        "rPr",
        "sectPr",
        "pPrChange",  # Tracked change must be last
    ]

    # Alignment value mapping
    ALIGNMENT_VALUES = {
        "left": "left",
        "center": "center",
        "right": "right",
        "justify": "both",
        "both": "both",
    }

    @classmethod
    def _get_insert_position(cls, ppr: etree._Element, element_name: str) -> int:
        """Get the correct position to insert an element in pPr.

        OOXML requires pPr child elements in a specific order.
        This method finds the correct index to insert a new element.

        Args:
            ppr: The parent pPr element
            element_name: Local name of element to insert (without namespace)

        Returns:
            Index position where the element should be inserted
        """
        try:
            target_order = cls.PPR_ELEMENT_ORDER.index(element_name)
        except ValueError:
            # Unknown element, append at end (but before pPrChange)
            return len(ppr)

        for i, child in enumerate(ppr):
            child_name = etree.QName(child.tag).localname
            try:
                child_order = cls.PPR_ELEMENT_ORDER.index(child_name)
                if child_order > target_order:
                    return i
            except ValueError:
                # Unknown child, keep looking
                continue

        return len(ppr)

    @classmethod
    def build(cls, nsmap: dict[str, str] | None = None, **kwargs: Any) -> etree._Element:
        """Build a <w:pPr> element from keyword arguments.

        Args:
            nsmap: Namespace map to use
            **kwargs: Property values to set

        Returns:
            New <w:pPr> element with specified properties
        """
        if nsmap is None:
            nsmap = {"w": WORD_NAMESPACE}

        ppr = etree.Element(_w("pPr"), nsmap=nsmap)
        cls._apply_properties(ppr, kwargs)
        return ppr

    @classmethod
    def merge(cls, base: etree._Element | None, updates: dict[str, Any]) -> etree._Element:
        """Merge updates into existing <w:pPr>, returning new element.

        Creates a deep copy of base (if provided) and applies updates.
        Does not modify the original element.

        Args:
            base: Existing <w:pPr> element (or None for empty)
            updates: Property values to update

        Returns:
            New <w:pPr> element with merged properties
        """
        if base is not None:
            ppr = deepcopy(base)
            # Remove any existing pPrChange from the copy
            for change in ppr.findall(_w("pPrChange")):
                ppr.remove(change)
        else:
            ppr = etree.Element(_w("pPr"))

        cls._apply_properties(ppr, updates)
        return ppr

    @classmethod
    def _apply_properties(cls, ppr: etree._Element, props: dict[str, Any]) -> None:
        """Apply property values to a <w:pPr> element (in-place).

        Args:
            ppr: The <w:pPr> element to modify
            props: Property values to apply
        """
        for prop_name, value in props.items():
            if value is None:
                continue

            if prop_name == "alignment":
                cls._set_alignment(ppr, value)
            elif prop_name in ("spacing_before", "spacing_after", "line_spacing"):
                cls._set_spacing(ppr, prop_name, value)
            elif prop_name in (
                "indent_left",
                "indent_right",
                "indent_first_line",
                "indent_hanging",
            ):
                cls._set_indent(ppr, prop_name, value)

    @classmethod
    def _set_alignment(cls, ppr: etree._Element, value: str | bool) -> None:
        """Set or remove paragraph alignment."""
        jc = ppr.find(_w("jc"))

        if value is False:
            if jc is not None:
                ppr.remove(jc)
            return

        if jc is None:
            jc = etree.Element(_w("jc"))
            pos = cls._get_insert_position(ppr, "jc")
            ppr.insert(pos, jc)

        # At this point value is a string (bool False was handled above)
        str_value = str(value) if not isinstance(value, str) else value
        xml_val = cls.ALIGNMENT_VALUES.get(str_value, str_value)
        jc.set(_w("val"), xml_val)

    @classmethod
    def _set_spacing(cls, ppr: etree._Element, prop_name: str, value: float | bool) -> None:
        """Set or modify spacing properties."""
        spacing = ppr.find(_w("spacing"))

        if value is False:
            # Remove specific attribute, not whole element
            if spacing is not None:
                attr_map = {
                    "spacing_before": "before",
                    "spacing_after": "after",
                    "line_spacing": "line",
                }
                attr = attr_map.get(prop_name)
                if attr and _w(attr) in spacing.attrib:
                    del spacing.attrib[_w(attr)]
            return

        if spacing is None:
            spacing = etree.Element(_w("spacing"))
            pos = cls._get_insert_position(ppr, "spacing")
            ppr.insert(pos, spacing)

        if prop_name == "spacing_before":
            spacing.set(_w("before"), str(points_to_twips(value)))
        elif prop_name == "spacing_after":
            spacing.set(_w("after"), str(points_to_twips(value)))
        elif prop_name == "line_spacing":
            # Line spacing in OOXML is in 240ths of a line
            # 240 = single space (1.0), 360 = 1.5 space, 480 = double space
            spacing.set(_w("line"), str(int(value * 240)))
            spacing.set(_w("lineRule"), "auto")

    @classmethod
    def _set_indent(cls, ppr: etree._Element, prop_name: str, value: float | bool) -> None:
        """Set or modify indentation properties."""
        ind = ppr.find(_w("ind"))

        if value is False:
            if ind is not None:
                attr_map = {
                    "indent_left": "left",
                    "indent_right": "right",
                    "indent_first_line": "firstLine",
                    "indent_hanging": "hanging",
                }
                attr = attr_map.get(prop_name)
                if attr and _w(attr) in ind.attrib:
                    del ind.attrib[_w(attr)]
            return

        if ind is None:
            ind = etree.Element(_w("ind"))
            pos = cls._get_insert_position(ppr, "ind")
            ppr.insert(pos, ind)

        twips = inches_to_twips(value)

        if prop_name == "indent_left":
            ind.set(_w("left"), str(twips))
        elif prop_name == "indent_right":
            ind.set(_w("right"), str(twips))
        elif prop_name == "indent_first_line":
            # Remove hanging if setting first line
            if _w("hanging") in ind.attrib:
                del ind.attrib[_w("hanging")]
            ind.set(_w("firstLine"), str(twips))
        elif prop_name == "indent_hanging":
            # Remove firstLine if setting hanging
            if _w("firstLine") in ind.attrib:
                del ind.attrib[_w("firstLine")]
            ind.set(_w("hanging"), str(twips))

    @classmethod
    def extract(cls, ppr: etree._Element | None) -> dict[str, Any]:
        """Extract Python dict from <w:pPr> element.

        Args:
            ppr: The <w:pPr> element to extract from (or None)

        Returns:
            Dictionary of property values
        """
        if ppr is None:
            return {}

        result: dict[str, Any] = {}

        # Extract alignment
        jc = ppr.find(_w("jc"))
        if jc is not None:
            val = jc.get(_w("val"))
            if val:
                # Reverse map "both" to "justify" for consistency
                result["alignment"] = "justify" if val == "both" else val

        # Extract spacing
        spacing = ppr.find(_w("spacing"))
        if spacing is not None:
            before = spacing.get(_w("before"))
            if before:
                result["spacing_before"] = twips_to_points(int(before))

            after = spacing.get(_w("after"))
            if after:
                result["spacing_after"] = twips_to_points(int(after))

            line = spacing.get(_w("line"))
            if line:
                result["line_spacing"] = int(line) / 240

        # Extract indentation
        ind = ppr.find(_w("ind"))
        if ind is not None:
            left = ind.get(_w("left"))
            if left:
                result["indent_left"] = twips_to_inches(int(left))

            right = ind.get(_w("right"))
            if right:
                result["indent_right"] = twips_to_inches(int(right))

            first_line = ind.get(_w("firstLine"))
            if first_line:
                result["indent_first_line"] = twips_to_inches(int(first_line))

            hanging = ind.get(_w("hanging"))
            if hanging:
                result["indent_hanging"] = twips_to_inches(int(hanging))

        return result

    @classmethod
    def diff(
        cls, old: etree._Element | None, new: etree._Element | None
    ) -> dict[str, tuple[Any, Any]]:
        """Return dict of properties that differ between old and new.

        Args:
            old: Original <w:pPr> element (or None)
            new: New <w:pPr> element (or None)

        Returns:
            Dictionary mapping property names to (old_value, new_value) tuples
        """
        old_props = cls.extract(old)
        new_props = cls.extract(new)

        all_keys = set(old_props.keys()) | set(new_props.keys())
        diff_result: dict[str, tuple[Any, Any]] = {}

        for key in all_keys:
            old_val = old_props.get(key)
            new_val = new_props.get(key)
            if old_val != new_val:
                diff_result[key] = (old_val, new_val)

        return diff_result

    @classmethod
    def has_changes(cls, old: etree._Element | None, new: etree._Element | None) -> bool:
        """Check if there are any property differences.

        Args:
            old: Original <w:pPr> element (or None)
            new: New <w:pPr> element (or None)

        Returns:
            True if properties differ, False if identical
        """
        return bool(cls.diff(old, new))


# Run splitting utilities for format_tracked


def split_run_at_offset(run: etree._Element, offset: int) -> tuple[etree._Element, etree._Element]:
    """Split a run at the specified character offset.

    Creates two runs: one containing text before the offset, one containing
    text from the offset onwards. Both runs inherit the original run's properties.

    Handles runs with multiple <w:t> nodes by finding which node the offset
    falls within and splitting appropriately.

    Args:
        run: The <w:r> element to split
        offset: Character position to split at (0-based, across all w:t nodes)

    Returns:
        Tuple of (before_run, after_run)

    Example:
        If run contains "Hello World" and offset=6:
        Returns (run with "Hello "), (run with "World")
    """
    # Get all text elements and concatenate
    text_elems = run.findall(_w("t"))
    if not text_elems:
        # No text to split - return original and empty
        empty_run = deepcopy(run)
        return run, empty_run

    full_text = "".join(elem.text or "" for elem in text_elems)

    if offset <= 0:
        # Split at beginning - return empty before, original after
        empty_run = deepcopy(run)
        for t in empty_run.findall(_w("t")):
            t.text = ""
        return empty_run, run
    if offset >= len(full_text):
        # Split at end - return original before, empty after
        empty_run = deepcopy(run)
        for t in empty_run.findall(_w("t")):
            t.text = ""
        return run, empty_run

    # Find which w:t node the offset falls within
    cumulative = 0
    split_node_idx = 0
    local_offset = 0

    for i, elem in enumerate(text_elems):
        elem_text = elem.text or ""
        if cumulative + len(elem_text) > offset:
            split_node_idx = i
            local_offset = offset - cumulative
            break
        cumulative += len(elem_text)
    else:
        # Offset at very end
        split_node_idx = len(text_elems) - 1
        local_offset = len(text_elems[-1].text or "")

    # Create before run - keep text up to and including split point
    before_run = deepcopy(run)
    before_text_elems = before_run.findall(_w("t"))
    for i, t in enumerate(before_text_elems):
        if i < split_node_idx:
            # Keep entire text
            pass
        elif i == split_node_idx:
            # Split this node
            before_text = (t.text or "")[:local_offset]
            t.text = before_text
            if before_text and (before_text[0].isspace() or before_text[-1].isspace()):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        else:
            # Clear text after split point
            t.text = ""

    # Create after run - keep text from split point onwards
    after_run = deepcopy(run)
    after_text_elems = after_run.findall(_w("t"))
    for i, t in enumerate(after_text_elems):
        if i < split_node_idx:
            # Clear text before split point
            t.text = ""
        elif i == split_node_idx:
            # Split this node
            after_text = (t.text or "")[local_offset:]
            t.text = after_text
            if after_text and (after_text[0].isspace() or after_text[-1].isspace()):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        else:
            # Keep entire text
            pass

    return before_run, after_run


def get_run_text(run: etree._Element) -> str:
    """Extract text content from a run element.

    Args:
        run: The <w:r> element

    Returns:
        Text content of all <w:t> elements in the run
    """
    text_elements = run.findall(_w("t"))
    return "".join(elem.text or "" for elem in text_elements)
