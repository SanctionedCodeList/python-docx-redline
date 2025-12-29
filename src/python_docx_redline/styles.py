"""
StyleManager class for managing word/styles.xml in OOXML packages.

This module provides a clean abstraction for managing style definitions
in a Word document, including paragraph styles, character styles,
table styles, and numbering styles.
"""

from __future__ import annotations

import logging
from collections.abc import Iterator
from typing import TYPE_CHECKING

from lxml import etree

from .constants import WORD_NAMESPACE, w
from .models.style import (
    ParagraphFormatting,
    RunFormatting,
    Style,
    StyleType,
)

if TYPE_CHECKING:
    from .package import OOXMLPackage

logger = logging.getLogger(__name__)

# Path to styles.xml within the package
STYLES_PATH = "word/styles.xml"


class StyleManager:
    """Manages word/styles.xml in OOXML packages.

    This class handles the low-level operations of:
    - Reading the styles file
    - Parsing style definitions into Style objects
    - Creating minimal styles when none exist
    - Tracking modifications for save operations

    Example:
        >>> style_mgr = StyleManager(package)
        >>> for style_id, style in style_mgr._styles.items():
        ...     print(f"{style_id}: {style.name} ({style.style_type.value})")

    Attributes:
        package: The OOXMLPackage containing this styles file
    """

    def __init__(self, package: OOXMLPackage) -> None:
        """Initialize a StyleManager for a package.

        Args:
            package: The OOXMLPackage containing the word/styles.xml file
        """
        self._package = package
        self._styles_path = package.temp_dir / STYLES_PATH
        self._root: etree._Element | None = None
        self._tree: etree._ElementTree | None = None
        self._styles: dict[str, Style] = {}
        self._modified = False

        # Load styles on initialization
        self._load()

    def _load(self) -> None:
        """Load and parse the styles.xml file.

        Reads word/styles.xml from the package and parses it. If the file
        doesn't exist, creates a minimal default styles structure.
        """
        if self._styles_path.exists():
            parser = etree.XMLParser(remove_blank_text=False)
            self._tree = etree.parse(str(self._styles_path), parser)
            self._root = self._tree.getroot()
            logger.debug(f"Loaded styles from {self._styles_path}")
        else:
            # Create minimal styles structure
            self._root = self._create_minimal_styles()
            self._tree = etree.ElementTree(self._root)
            self._modified = True
            logger.debug("Created minimal styles structure")

        # Parse all style elements
        self._parse_styles()

    def _parse_styles(self) -> None:
        """Parse all w:style elements from the loaded XML.

        Iterates through all w:style elements in the styles XML and
        converts each to a Style object, storing them in the _styles dict
        keyed by style_id.
        """
        if self._root is None:
            return

        self._styles.clear()

        # Find all w:style elements
        for style_elem in self._root.findall(w("style"), namespaces=None):
            style = self._element_to_style(style_elem)
            if style is not None:
                self._styles[style.style_id] = style
                logger.debug(f"Parsed style: {style.style_id}")

    def _element_to_style(self, element: etree._Element) -> Style | None:
        """Convert a w:style XML element to a Style object.

        Args:
            element: The w:style XML element to convert

        Returns:
            A Style object populated with data from the XML, or None if
            the element lacks required attributes (style_id)
        """
        # Extract style_id (required)
        style_id = element.get(w("styleId"))
        if not style_id:
            logger.warning("Skipping style element without styleId attribute")
            return None

        # Extract style type
        style_type_str = element.get(w("type"), "paragraph")
        try:
            style_type = StyleType(style_type_str)
        except ValueError:
            logger.warning(f"Unknown style type '{style_type_str}' for {style_id}")
            style_type = StyleType.PARAGRAPH

        # Extract name from w:name element
        name_elem = element.find(w("name"), namespaces=None)
        name = name_elem.get(w("val"), style_id) if name_elem is not None else style_id

        # Extract based_on from w:basedOn element
        based_on_elem = element.find(w("basedOn"), namespaces=None)
        based_on = based_on_elem.get(w("val")) if based_on_elem is not None else None

        # Extract next_style from w:next element
        next_elem = element.find(w("next"), namespaces=None)
        next_style = next_elem.get(w("val")) if next_elem is not None else None

        # Extract linked_style from w:link element
        link_elem = element.find(w("link"), namespaces=None)
        linked_style = link_elem.get(w("val")) if link_elem is not None else None

        # Extract UI properties
        ui_priority_elem = element.find(w("uiPriority"), namespaces=None)
        ui_priority = None
        if ui_priority_elem is not None:
            try:
                ui_priority = int(ui_priority_elem.get(w("val"), "99"))
            except ValueError:
                ui_priority = 99

        # Boolean properties - presence indicates True
        quick_format = element.find(w("qFormat"), namespaces=None) is not None
        semi_hidden = element.find(w("semiHidden"), namespaces=None) is not None
        unhide_when_used = element.find(w("unhideWhenUsed"), namespaces=None) is not None

        # Parse run formatting from w:rPr
        rpr_elem = element.find(w("rPr"), namespaces=None)
        run_formatting = self._parse_run_formatting(rpr_elem)

        # Parse paragraph formatting from w:pPr
        ppr_elem = element.find(w("pPr"), namespaces=None)
        paragraph_formatting = self._parse_paragraph_formatting(ppr_elem)

        return Style(
            style_id=style_id,
            name=name,
            style_type=style_type,
            based_on=based_on,
            next_style=next_style,
            linked_style=linked_style,
            run_formatting=run_formatting,
            paragraph_formatting=paragraph_formatting,
            ui_priority=ui_priority,
            quick_format=quick_format,
            semi_hidden=semi_hidden,
            unhide_when_used=unhide_when_used,
            _element=element,
        )

    def _parse_run_formatting(self, rpr_elem: etree._Element | None) -> RunFormatting:
        """Parse run (character) formatting from a w:rPr element.

        Args:
            rpr_elem: The w:rPr XML element to parse, or None

        Returns:
            A RunFormatting object populated with data from the XML
        """
        if rpr_elem is None:
            return RunFormatting()

        # Boolean properties - presence indicates True, unless w:val="0"
        bold = self._parse_bool_property(rpr_elem, "b")
        italic = self._parse_bool_property(rpr_elem, "i")
        strikethrough = self._parse_bool_property(rpr_elem, "strike")
        small_caps = self._parse_bool_property(rpr_elem, "smallCaps")
        all_caps = self._parse_bool_property(rpr_elem, "caps")

        # Underline - w:u with w:val attribute
        underline_elem = rpr_elem.find(w("u"), namespaces=None)
        underline: bool | str | None = None
        if underline_elem is not None:
            underline_val = underline_elem.get(w("val"), "single")
            if underline_val == "none":
                underline = False
            elif underline_val == "single":
                underline = True
            else:
                underline = underline_val

        # Font name from w:rFonts
        fonts_elem = rpr_elem.find(w("rFonts"), namespaces=None)
        font_name = None
        if fonts_elem is not None:
            # Try ascii first, then hAnsi
            font_name = fonts_elem.get(w("ascii")) or fonts_elem.get(w("hAnsi"))

        # Font size from w:sz (in half-points, so divide by 2)
        sz_elem = rpr_elem.find(w("sz"), namespaces=None)
        font_size = None
        if sz_elem is not None:
            try:
                half_points = int(sz_elem.get(w("val"), "0"))
                font_size = half_points / 2.0
            except ValueError:
                pass

        # Color from w:color
        color_elem = rpr_elem.find(w("color"), namespaces=None)
        color = None
        if color_elem is not None:
            color_val = color_elem.get(w("val"))
            if color_val and color_val != "auto":
                # Normalize to #RRGGBB format
                if len(color_val) == 6:
                    color = f"#{color_val.upper()}"
                else:
                    color = color_val

        # Highlight from w:highlight
        highlight_elem = rpr_elem.find(w("highlight"), namespaces=None)
        highlight = highlight_elem.get(w("val")) if highlight_elem is not None else None

        # Superscript/subscript from w:vertAlign
        vert_align_elem = rpr_elem.find(w("vertAlign"), namespaces=None)
        superscript = None
        subscript = None
        if vert_align_elem is not None:
            vert_val = vert_align_elem.get(w("val"))
            if vert_val == "superscript":
                superscript = True
            elif vert_val == "subscript":
                subscript = True

        return RunFormatting(
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            font_name=font_name,
            font_size=font_size,
            color=color,
            highlight=highlight,
            superscript=superscript,
            subscript=subscript,
            small_caps=small_caps,
            all_caps=all_caps,
        )

    def _parse_paragraph_formatting(self, ppr_elem: etree._Element | None) -> ParagraphFormatting:
        """Parse paragraph formatting from a w:pPr element.

        Args:
            ppr_elem: The w:pPr XML element to parse, or None

        Returns:
            A ParagraphFormatting object populated with data from the XML
        """
        if ppr_elem is None:
            return ParagraphFormatting()

        # Alignment from w:jc
        jc_elem = ppr_elem.find(w("jc"), namespaces=None)
        alignment = None
        if jc_elem is not None:
            jc_val = jc_elem.get(w("val"))
            # Map OOXML values to our simplified names
            alignment_map = {
                "left": "left",
                "start": "left",
                "center": "center",
                "right": "right",
                "end": "right",
                "both": "justify",
                "distribute": "justify",
            }
            alignment = alignment_map.get(jc_val, jc_val)

        # Spacing from w:spacing
        spacing_elem = ppr_elem.find(w("spacing"), namespaces=None)
        spacing_before = None
        spacing_after = None
        line_spacing = None
        if spacing_elem is not None:
            # before and after are in twentieths of a point
            before_val = spacing_elem.get(w("before"))
            if before_val:
                try:
                    spacing_before = int(before_val) / 20.0
                except ValueError:
                    pass

            after_val = spacing_elem.get(w("after"))
            if after_val:
                try:
                    spacing_after = int(after_val) / 20.0
                except ValueError:
                    pass

            # Line spacing depends on lineRule
            line_val = spacing_elem.get(w("line"))
            line_rule = spacing_elem.get(w("lineRule"), "auto")
            if line_val:
                try:
                    line_num = int(line_val)
                    if line_rule == "auto":
                        # Value is in 240ths of a line
                        line_spacing = line_num / 240.0
                    # For "exact" or "atLeast", value is in twentieths of a point
                    # We could convert, but for now leave as None for these
                except ValueError:
                    pass

        # Indentation from w:ind
        ind_elem = ppr_elem.find(w("ind"), namespaces=None)
        indent_left = None
        indent_right = None
        indent_first_line = None
        indent_hanging = None
        if ind_elem is not None:
            # Values are in twentieths of a point, convert to inches (1 inch = 1440 twips)
            left_val = ind_elem.get(w("left"))
            if left_val:
                try:
                    indent_left = int(left_val) / 1440.0
                except ValueError:
                    pass

            right_val = ind_elem.get(w("right"))
            if right_val:
                try:
                    indent_right = int(right_val) / 1440.0
                except ValueError:
                    pass

            first_line_val = ind_elem.get(w("firstLine"))
            if first_line_val:
                try:
                    indent_first_line = int(first_line_val) / 1440.0
                except ValueError:
                    pass

            hanging_val = ind_elem.get(w("hanging"))
            if hanging_val:
                try:
                    indent_hanging = int(hanging_val) / 1440.0
                except ValueError:
                    pass

        # Keep properties
        keep_next = ppr_elem.find(w("keepNext"), namespaces=None) is not None
        keep_lines = ppr_elem.find(w("keepLines"), namespaces=None) is not None

        # Outline level from w:outlineLvl
        outline_lvl_elem = ppr_elem.find(w("outlineLvl"), namespaces=None)
        outline_level = None
        if outline_lvl_elem is not None:
            try:
                outline_level = int(outline_lvl_elem.get(w("val"), "0"))
            except ValueError:
                pass

        return ParagraphFormatting(
            alignment=alignment,
            spacing_before=spacing_before,
            spacing_after=spacing_after,
            line_spacing=line_spacing,
            indent_left=indent_left,
            indent_right=indent_right,
            indent_first_line=indent_first_line,
            indent_hanging=indent_hanging,
            keep_next=keep_next if keep_next else None,
            keep_lines=keep_lines if keep_lines else None,
            outline_level=outline_level,
        )

    def _parse_bool_property(self, parent: etree._Element, tag_name: str) -> bool | None:
        """Parse a boolean property element.

        In OOXML, boolean properties like w:b (bold) can be:
        - Present with no value: True
        - Present with w:val="1" or w:val="true": True
        - Present with w:val="0" or w:val="false": False
        - Absent: None (inherit from parent)

        Args:
            parent: The parent element to search in
            tag_name: The local tag name (without namespace)

        Returns:
            True, False, or None based on the element state
        """
        elem = parent.find(w(tag_name), namespaces=None)
        if elem is None:
            return None

        val = elem.get(w("val"))
        if val is None:
            # Element present without value means True
            return True

        # Check for explicit false values
        return val.lower() not in ("0", "false", "off")

    def _create_minimal_styles(self) -> etree._Element:
        """Create a minimal styles.xml root element.

        Creates the minimum required structure for a valid styles.xml file,
        including docDefaults and a Normal style.

        Returns:
            The root w:styles element with minimal content
        """
        # Create root element with required namespaces
        nsmap = {
            "w": WORD_NAMESPACE,
            "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }
        root = etree.Element(w("styles"), nsmap=nsmap)

        # Create docDefaults
        doc_defaults = etree.SubElement(root, w("docDefaults"))

        # Run properties default
        rpr_default = etree.SubElement(doc_defaults, w("rPrDefault"))
        rpr = etree.SubElement(rpr_default, w("rPr"))

        # Default font
        fonts = etree.SubElement(rpr, w("rFonts"))
        fonts.set(w("ascii"), "Calibri")
        fonts.set(w("hAnsi"), "Calibri")

        # Default size (22 half-points = 11pt)
        sz = etree.SubElement(rpr, w("sz"))
        sz.set(w("val"), "22")
        sz_cs = etree.SubElement(rpr, w("szCs"))
        sz_cs.set(w("val"), "22")

        # Paragraph properties default
        ppr_default = etree.SubElement(doc_defaults, w("pPrDefault"))
        ppr = etree.SubElement(ppr_default, w("pPr"))

        # Default spacing
        spacing = etree.SubElement(ppr, w("spacing"))
        spacing.set(w("after"), "200")
        spacing.set(w("line"), "276")
        spacing.set(w("lineRule"), "auto")

        # Create Normal style (required)
        normal_style = etree.SubElement(root, w("style"))
        normal_style.set(w("type"), "paragraph")
        normal_style.set(w("default"), "1")
        normal_style.set(w("styleId"), "Normal")

        normal_name = etree.SubElement(normal_style, w("name"))
        normal_name.set(w("val"), "Normal")

        etree.SubElement(normal_style, w("qFormat"))

        # Create DefaultParagraphFont style (required for character styles)
        dpf_style = etree.SubElement(root, w("style"))
        dpf_style.set(w("type"), "character")
        dpf_style.set(w("default"), "1")
        dpf_style.set(w("styleId"), "DefaultParagraphFont")

        dpf_name = etree.SubElement(dpf_style, w("name"))
        dpf_name.set(w("val"), "Default Paragraph Font")

        ui_priority = etree.SubElement(dpf_style, w("uiPriority"))
        ui_priority.set(w("val"), "1")

        etree.SubElement(dpf_style, w("semiHidden"))
        etree.SubElement(dpf_style, w("unhideWhenUsed"))

        return root

    # -------------------------------------------------------------------------
    # Read Operations
    # -------------------------------------------------------------------------

    def get(self, style_id: str) -> Style | None:
        """Get a style by its ID.

        Args:
            style_id: The style identifier (e.g., "Normal", "Heading1",
                "FootnoteReference")

        Returns:
            The Style object if found, None otherwise

        Example:
            >>> styles = StyleManager(package)
            >>> normal = styles.get("Normal")
            >>> if normal:
            ...     print(f"Normal style: {normal.name}")
        """
        return self._styles.get(style_id)

    def get_by_name(self, name: str) -> Style | None:
        """Get a style by its display name (case-insensitive).

        Searches for a style by its display name rather than its internal ID.
        The comparison is case-insensitive to match Word's behavior.

        Args:
            name: The display name to search for (e.g., "Normal",
                "Heading 1", "footnote reference")

        Returns:
            The Style object if found, None otherwise

        Example:
            >>> styles = StyleManager(package)
            >>> heading = styles.get_by_name("heading 1")  # Case-insensitive
            >>> if heading:
            ...     print(f"Found: {heading.style_id}")
        """
        name_lower = name.lower()
        for style in self._styles.values():
            if style.name.lower() == name_lower:
                return style
        return None

    def list(
        self,
        style_type: StyleType | None = None,
        include_hidden: bool = False,
    ) -> list[Style]:
        """List all styles, optionally filtered by type.

        Returns styles in the document, with optional filtering by style type
        and visibility.

        Args:
            style_type: If provided, only return styles of this type
            include_hidden: If False (default), exclude semi_hidden styles

        Returns:
            List of Style objects matching the criteria

        Example:
            >>> styles = StyleManager(package)
            >>> # Get all visible styles
            >>> all_styles = styles.list()
            >>> # Get only paragraph styles
            >>> para_styles = styles.list(style_type=StyleType.PARAGRAPH)
            >>> # Get all styles including hidden ones
            >>> all_including_hidden = styles.list(include_hidden=True)
        """
        result = []
        for style in self._styles.values():
            # Filter by type if specified
            if style_type is not None and style.style_type != style_type:
                continue
            # Filter hidden styles if not included
            if not include_hidden and style.semi_hidden:
                continue
            result.append(style)
        return result

    def __contains__(self, style_id: str) -> bool:
        """Check if a style exists by ID.

        Enables the use of 'in' operator for checking style existence.

        Args:
            style_id: The style identifier to check

        Returns:
            True if the style exists, False otherwise

        Example:
            >>> styles = StyleManager(package)
            >>> if "FootnoteReference" in styles:
            ...     print("Style exists")
        """
        return style_id in self._styles

    def __iter__(self) -> Iterator[Style]:
        """Iterate over all styles.

        Enables iteration over StyleManager to get all Style objects.

        Yields:
            Style objects in the manager

        Example:
            >>> styles = StyleManager(package)
            >>> for style in styles:
            ...     print(f"{style.style_id}: {style.name}")
        """
        return iter(self._styles.values())

    def __len__(self) -> int:
        """Return the number of styles.

        Enables len() on StyleManager.

        Returns:
            The number of styles in the manager

        Example:
            >>> styles = StyleManager(package)
            >>> print(f"Document has {len(styles)} styles")
        """
        return len(self._styles)

    # -------------------------------------------------------------------------
    # Properties
    # -------------------------------------------------------------------------

    @property
    def is_modified(self) -> bool:
        """Check if there are unsaved modifications."""
        return self._modified

    # -------------------------------------------------------------------------
    # Write Operations
    # -------------------------------------------------------------------------

    def _style_to_element(self, style: Style) -> etree._Element:
        """Convert a Style object to a w:style XML element.

        Creates a complete w:style element with all properties from the Style
        object, including name, type, based_on, run formatting, paragraph
        formatting, and UI properties.

        Args:
            style: The Style object to convert

        Returns:
            A w:style XML element ready to be appended to the styles root
        """
        # Create the w:style element
        style_elem = etree.Element(w("style"))
        style_elem.set(w("type"), style.style_type.value)
        style_elem.set(w("styleId"), style.style_id)

        # Add w:name element (required)
        name_elem = etree.SubElement(style_elem, w("name"))
        name_elem.set(w("val"), style.name)

        # Add optional metadata elements
        if style.based_on:
            based_on_elem = etree.SubElement(style_elem, w("basedOn"))
            based_on_elem.set(w("val"), style.based_on)

        if style.next_style:
            next_elem = etree.SubElement(style_elem, w("next"))
            next_elem.set(w("val"), style.next_style)

        if style.linked_style:
            link_elem = etree.SubElement(style_elem, w("link"))
            link_elem.set(w("val"), style.linked_style)

        # Add UI properties
        if style.ui_priority is not None:
            ui_priority_elem = etree.SubElement(style_elem, w("uiPriority"))
            ui_priority_elem.set(w("val"), str(style.ui_priority))

        if style.quick_format:
            etree.SubElement(style_elem, w("qFormat"))

        if style.semi_hidden:
            etree.SubElement(style_elem, w("semiHidden"))

        if style.unhide_when_used:
            etree.SubElement(style_elem, w("unhideWhenUsed"))

        # Add paragraph formatting (w:pPr)
        ppr_elem = self._paragraph_formatting_to_element(style.paragraph_formatting)
        if ppr_elem is not None:
            style_elem.append(ppr_elem)

        # Add run formatting (w:rPr)
        rpr_elem = self._run_formatting_to_element(style.run_formatting)
        if rpr_elem is not None:
            style_elem.append(rpr_elem)

        return style_elem

    def _run_formatting_to_element(self, fmt: RunFormatting) -> etree._Element | None:
        """Convert RunFormatting to a w:rPr XML element.

        Args:
            fmt: The RunFormatting object to convert

        Returns:
            A w:rPr element if any formatting is set, None otherwise
        """
        rpr = etree.Element(w("rPr"))
        has_content = False

        # Bold
        if fmt.bold is not None:
            b_elem = etree.SubElement(rpr, w("b"))
            if not fmt.bold:
                b_elem.set(w("val"), "0")
            has_content = True

        # Italic
        if fmt.italic is not None:
            i_elem = etree.SubElement(rpr, w("i"))
            if not fmt.italic:
                i_elem.set(w("val"), "0")
            has_content = True

        # Underline
        if fmt.underline is not None:
            u_elem = etree.SubElement(rpr, w("u"))
            if fmt.underline is True:
                u_elem.set(w("val"), "single")
            elif fmt.underline is False:
                u_elem.set(w("val"), "none")
            else:
                # String value like "double", "wave", etc.
                u_elem.set(w("val"), str(fmt.underline))
            has_content = True

        # Strikethrough
        if fmt.strikethrough is not None:
            strike_elem = etree.SubElement(rpr, w("strike"))
            if not fmt.strikethrough:
                strike_elem.set(w("val"), "0")
            has_content = True

        # Font name
        if fmt.font_name is not None:
            fonts_elem = etree.SubElement(rpr, w("rFonts"))
            fonts_elem.set(w("ascii"), fmt.font_name)
            fonts_elem.set(w("hAnsi"), fmt.font_name)
            has_content = True

        # Font size (convert points to half-points)
        if fmt.font_size is not None:
            half_points = int(fmt.font_size * 2)
            sz_elem = etree.SubElement(rpr, w("sz"))
            sz_elem.set(w("val"), str(half_points))
            sz_cs_elem = etree.SubElement(rpr, w("szCs"))
            sz_cs_elem.set(w("val"), str(half_points))
            has_content = True

        # Color
        if fmt.color is not None:
            color_elem = etree.SubElement(rpr, w("color"))
            # Strip # prefix if present
            color_val = fmt.color.lstrip("#")
            color_elem.set(w("val"), color_val)
            has_content = True

        # Highlight
        if fmt.highlight is not None:
            highlight_elem = etree.SubElement(rpr, w("highlight"))
            highlight_elem.set(w("val"), fmt.highlight)
            has_content = True

        # Superscript/subscript (via w:vertAlign)
        if fmt.superscript:
            vert_elem = etree.SubElement(rpr, w("vertAlign"))
            vert_elem.set(w("val"), "superscript")
            has_content = True
        elif fmt.subscript:
            vert_elem = etree.SubElement(rpr, w("vertAlign"))
            vert_elem.set(w("val"), "subscript")
            has_content = True

        # Small caps
        if fmt.small_caps is not None:
            small_caps_elem = etree.SubElement(rpr, w("smallCaps"))
            if not fmt.small_caps:
                small_caps_elem.set(w("val"), "0")
            has_content = True

        # All caps
        if fmt.all_caps is not None:
            caps_elem = etree.SubElement(rpr, w("caps"))
            if not fmt.all_caps:
                caps_elem.set(w("val"), "0")
            has_content = True

        return rpr if has_content else None

    def _paragraph_formatting_to_element(self, fmt: ParagraphFormatting) -> etree._Element | None:
        """Convert ParagraphFormatting to a w:pPr XML element.

        Args:
            fmt: The ParagraphFormatting object to convert

        Returns:
            A w:pPr element if any formatting is set, None otherwise
        """
        ppr = etree.Element(w("pPr"))
        has_content = False

        # Alignment (w:jc)
        if fmt.alignment is not None:
            jc_elem = etree.SubElement(ppr, w("jc"))
            # Map our simplified names to OOXML values
            alignment_map = {
                "left": "left",
                "center": "center",
                "right": "right",
                "justify": "both",
            }
            jc_val = alignment_map.get(fmt.alignment, fmt.alignment)
            jc_elem.set(w("val"), jc_val)
            has_content = True

        # Spacing (w:spacing)
        if (
            fmt.spacing_before is not None
            or fmt.spacing_after is not None
            or fmt.line_spacing is not None
        ):
            spacing_elem = etree.SubElement(ppr, w("spacing"))

            # spacing_before/after are in points, convert to twentieths of a point
            if fmt.spacing_before is not None:
                twips = int(fmt.spacing_before * 20)
                spacing_elem.set(w("before"), str(twips))

            if fmt.spacing_after is not None:
                twips = int(fmt.spacing_after * 20)
                spacing_elem.set(w("after"), str(twips))

            # line_spacing is a multiplier, convert to 240ths of a line
            if fmt.line_spacing is not None:
                line_val = int(fmt.line_spacing * 240)
                spacing_elem.set(w("line"), str(line_val))
                spacing_elem.set(w("lineRule"), "auto")

            has_content = True

        # Indentation (w:ind)
        if (
            fmt.indent_left is not None
            or fmt.indent_right is not None
            or fmt.indent_first_line is not None
            or fmt.indent_hanging is not None
        ):
            ind_elem = etree.SubElement(ppr, w("ind"))

            # Convert inches to twentieths of a point (1 inch = 1440 twips)
            if fmt.indent_left is not None:
                twips = int(fmt.indent_left * 1440)
                ind_elem.set(w("left"), str(twips))

            if fmt.indent_right is not None:
                twips = int(fmt.indent_right * 1440)
                ind_elem.set(w("right"), str(twips))

            if fmt.indent_first_line is not None:
                twips = int(fmt.indent_first_line * 1440)
                ind_elem.set(w("firstLine"), str(twips))

            if fmt.indent_hanging is not None:
                twips = int(fmt.indent_hanging * 1440)
                ind_elem.set(w("hanging"), str(twips))

            has_content = True

        # Keep properties
        if fmt.keep_next:
            etree.SubElement(ppr, w("keepNext"))
            has_content = True

        if fmt.keep_lines:
            etree.SubElement(ppr, w("keepLines"))
            has_content = True

        # Outline level
        if fmt.outline_level is not None:
            outline_elem = etree.SubElement(ppr, w("outlineLvl"))
            outline_elem.set(w("val"), str(fmt.outline_level))
            has_content = True

        return ppr if has_content else None

    def add(self, style: Style) -> None:
        """Add a new style to the document.

        Creates a new style definition in the document's styles.xml. The style
        must have a unique style_id that doesn't already exist.

        Args:
            style: The Style object to add

        Raises:
            ValueError: If a style with the same style_id already exists

        Example:
            >>> from python_docx_redline.models.style import Style, StyleType, RunFormatting
            >>> custom = Style(
            ...     style_id="MyStyle",
            ...     name="My Custom Style",
            ...     style_type=StyleType.CHARACTER,
            ...     run_formatting=RunFormatting(bold=True, color="#FF0000"),
            ... )
            >>> styles.add(custom)
            >>> styles.save()
        """
        if style.style_id in self._styles:
            raise ValueError(f"Style with id '{style.style_id}' already exists")

        if self._root is None:
            raise RuntimeError("StyleManager not properly initialized")

        # Convert style to XML element
        style_elem = self._style_to_element(style)

        # Append to root
        self._root.append(style_elem)

        # Store reference to the element in the style
        style._element = style_elem

        # Add to internal dict
        self._styles[style.style_id] = style

        # Mark as modified
        self._modified = True

        logger.debug(f"Added style: {style.style_id}")

    def ensure_style(
        self,
        style_id: str,
        name: str,
        style_type: StyleType,
        *,
        based_on: str | None = None,
        next_style: str | None = None,
        linked_style: str | None = None,
        run_formatting: RunFormatting | None = None,
        paragraph_formatting: ParagraphFormatting | None = None,
        ui_priority: int | None = None,
        quick_format: bool = False,
        semi_hidden: bool = False,
        unhide_when_used: bool = False,
    ) -> Style:
        """Ensure a style exists, creating it if necessary.

        This is the primary method for features that require specific styles.
        If a style with the given style_id already exists, it is returned.
        Otherwise, a new style is created with the provided parameters and added.

        Args:
            style_id: The unique identifier for the style
            name: The display name of the style
            style_type: The type of style (paragraph, character, etc.)
            based_on: Optional style_id of parent style to inherit from
            next_style: Optional style_id of style to apply after pressing Enter
            linked_style: Optional style_id of linked style
            run_formatting: Optional character formatting
            paragraph_formatting: Optional paragraph formatting
            ui_priority: Optional sort order in style gallery
            quick_format: Whether style appears in Quick Style gallery
            semi_hidden: Whether style is hidden from UI
            unhide_when_used: Whether to unhide style when first used

        Returns:
            The existing or newly created Style object

        Example:
            >>> style = styles.ensure_style(
            ...     style_id="FootnoteReference",
            ...     name="footnote reference",
            ...     style_type=StyleType.CHARACTER,
            ...     based_on="DefaultParagraphFont",
            ...     run_formatting=RunFormatting(superscript=True),
            ...     ui_priority=99,
            ...     unhide_when_used=True,
            ... )
        """
        # Return existing style if it exists
        existing = self.get(style_id)
        if existing is not None:
            logger.debug(f"Style '{style_id}' already exists, returning existing")
            return existing

        # Create new style
        style = Style(
            style_id=style_id,
            name=name,
            style_type=style_type,
            based_on=based_on,
            next_style=next_style,
            linked_style=linked_style,
            run_formatting=run_formatting if run_formatting else RunFormatting(),
            paragraph_formatting=(
                paragraph_formatting if paragraph_formatting else ParagraphFormatting()
            ),
            ui_priority=ui_priority,
            quick_format=quick_format,
            semi_hidden=semi_hidden,
            unhide_when_used=unhide_when_used,
        )

        # Add the new style
        self.add(style)

        logger.debug(f"Created new style: {style_id}")
        return style

    def save(self) -> None:
        """Persist changes to the word/styles.xml file.

        Only writes if modifications were made. Creates the word/
        directory if it doesn't exist.
        """
        if not self._modified or self._tree is None:
            return

        # Ensure parent directory exists
        self._styles_path.parent.mkdir(parents=True, exist_ok=True)

        # Write the file
        self._tree.write(
            str(self._styles_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

        self._modified = False
        logger.debug(f"Saved styles file: {self._styles_path}")
