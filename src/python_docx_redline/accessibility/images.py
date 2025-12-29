"""
Image and embedded object extraction for the accessibility layer.

This module provides functionality to extract images, charts, diagrams,
shapes, and other embedded objects from Word documents, creating stable
refs for each.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from ..constants import (
    A_NAMESPACE,
    OFFICE_RELATIONSHIPS_NAMESPACE,
    PIC_NAMESPACE,
    WORD_NAMESPACE,
    WP_NAMESPACE,
    w,
    wp,
)
from .types import (
    ImageInfo,
    ImagePosition,
    ImagePositionType,
    ImageSize,
    ImageType,
)

if TYPE_CHECKING:
    pass


class ImageExtractor:
    """Extracts image and embedded object information from OOXML.

    This class parses w:drawing, w:pict, and w:object elements to create
    ImageInfo objects with stable refs.

    Ref format for images:
    - img:p_idx/img_idx for inline images (e.g., img:5/0)
    - img:p_idx/f:img_idx for floating images (e.g., img:5/f:0)
    - chart:p_idx/chart_idx for charts
    - diagram:p_idx/diag_idx for SmartArt
    - shape:p_idx/shape_idx for shapes
    - vml:p_idx/vml_idx for legacy VML graphics
    - obj:p_idx/obj_idx for OLE objects

    Attributes:
        xml_root: Root element of the document XML
    """

    # Namespace map for XPath queries
    NSMAP = {
        "w": WORD_NAMESPACE,
        "wp": WP_NAMESPACE,
        "a": A_NAMESPACE,
        "pic": PIC_NAMESPACE,
        "r": OFFICE_RELATIONSHIPS_NAMESPACE,
    }

    def __init__(self, xml_root: etree._Element) -> None:
        """Initialize the image extractor.

        Args:
            xml_root: Root element of the document XML tree
        """
        self.xml_root = xml_root

    def extract_from_paragraph(
        self,
        paragraph: etree._Element,
        paragraph_index: int,
    ) -> list[ImageInfo]:
        """Extract all images from a paragraph.

        Args:
            paragraph: The w:p element
            paragraph_index: Index of this paragraph in the document

        Returns:
            List of ImageInfo objects for images in this paragraph
        """
        images: list[ImageInfo] = []

        # Track indices for each image type
        inline_index = 0
        floating_index = 0
        # Note: chart, diagram, shape indices reserved for future use
        vml_index = 0
        obj_index = 0

        paragraph_ref = f"p:{paragraph_index}"

        # Find all w:drawing elements (modern DrawingML)
        for drawing in paragraph.iter(w("drawing")):
            # Check for inline images (wp:inline)
            for inline in drawing.iter(wp("inline")):
                info = self._extract_inline_image(
                    inline, paragraph_index, inline_index, paragraph_ref
                )
                if info:
                    images.append(info)
                inline_index += 1

            # Check for floating/anchored images (wp:anchor)
            for anchor in drawing.iter(wp("anchor")):
                info = self._extract_anchor_image(
                    anchor, paragraph_index, floating_index, paragraph_ref
                )
                if info:
                    images.append(info)
                floating_index += 1

        # Find legacy VML graphics (w:pict)
        for pict in paragraph.iter(w("pict")):
            info = self._extract_vml_image(pict, paragraph_index, vml_index, paragraph_ref)
            if info:
                images.append(info)
            vml_index += 1

        # Find OLE objects (w:object)
        for obj in paragraph.iter(w("object")):
            info = self._extract_ole_object(obj, paragraph_index, obj_index, paragraph_ref)
            if info:
                images.append(info)
            obj_index += 1

        return images

    def _extract_inline_image(
        self,
        inline: etree._Element,
        p_idx: int,
        img_idx: int,
        paragraph_ref: str,
    ) -> ImageInfo | None:
        """Extract info from a wp:inline element.

        Args:
            inline: The wp:inline element
            p_idx: Paragraph index
            img_idx: Image index within paragraph
            paragraph_ref: Ref of containing paragraph

        Returns:
            ImageInfo or None if not a valid image
        """
        # Determine image type and create ref
        image_type = self._determine_image_type(inline)
        ref = self._make_ref(image_type, p_idx, img_idx, floating=False)

        # Extract extent (size)
        size = self._extract_extent(inline)

        # Extract docPr (name and description)
        name, alt_text = self._extract_doc_properties(inline)

        # Get relationship ID for the actual image
        rel_id = self._extract_relationship_id(inline)

        return ImageInfo(
            ref=ref,
            image_type=image_type,
            position_type=ImagePositionType.INLINE,
            name=name,
            alt_text=alt_text,
            size=size,
            relationship_id=rel_id,
            paragraph_ref=paragraph_ref,
            _element=inline,
        )

    def _extract_anchor_image(
        self,
        anchor: etree._Element,
        p_idx: int,
        img_idx: int,
        paragraph_ref: str,
    ) -> ImageInfo | None:
        """Extract info from a wp:anchor element (floating image).

        Args:
            anchor: The wp:anchor element
            p_idx: Paragraph index
            img_idx: Floating image index within paragraph
            paragraph_ref: Ref of containing paragraph

        Returns:
            ImageInfo or None if not a valid image
        """
        # Determine image type and create ref
        image_type = self._determine_image_type(anchor)
        ref = self._make_ref(image_type, p_idx, img_idx, floating=True)

        # Extract extent (size)
        size = self._extract_extent(anchor)

        # Extract docPr (name and description)
        name, alt_text = self._extract_doc_properties(anchor)

        # Get relationship ID
        rel_id = self._extract_relationship_id(anchor)

        # Extract position info
        position = self._extract_position_info(anchor)

        return ImageInfo(
            ref=ref,
            image_type=image_type,
            position_type=ImagePositionType.FLOATING,
            name=name,
            alt_text=alt_text,
            size=size,
            relationship_id=rel_id,
            paragraph_ref=paragraph_ref,
            position=position,
            _element=anchor,
        )

    def _extract_vml_image(
        self,
        pict: etree._Element,
        p_idx: int,
        vml_idx: int,
        paragraph_ref: str,
    ) -> ImageInfo | None:
        """Extract info from a w:pict element (legacy VML).

        Args:
            pict: The w:pict element
            p_idx: Paragraph index
            vml_idx: VML image index within paragraph
            paragraph_ref: Ref of containing paragraph

        Returns:
            ImageInfo or None if not a valid VML graphic
        """
        ref = f"vml:{p_idx}/{vml_idx}"

        # VML images don't have the same metadata as DrawingML
        # We extract what we can from the VML namespace
        return ImageInfo(
            ref=ref,
            image_type=ImageType.VML,
            position_type=ImagePositionType.INLINE,
            name="",
            alt_text="",
            paragraph_ref=paragraph_ref,
            _element=pict,
        )

    def _extract_ole_object(
        self,
        obj: etree._Element,
        p_idx: int,
        obj_idx: int,
        paragraph_ref: str,
    ) -> ImageInfo | None:
        """Extract info from a w:object element (OLE embedded object).

        Args:
            obj: The w:object element
            p_idx: Paragraph index
            obj_idx: Object index within paragraph
            paragraph_ref: Ref of containing paragraph

        Returns:
            ImageInfo or None if not a valid OLE object
        """
        ref = f"obj:{p_idx}/{obj_idx}"

        return ImageInfo(
            ref=ref,
            image_type=ImageType.OLE_OBJECT,
            position_type=ImagePositionType.INLINE,
            name="",
            alt_text="",
            paragraph_ref=paragraph_ref,
            _element=obj,
        )

    def _determine_image_type(self, container: etree._Element) -> ImageType:
        """Determine the type of image from a wp:inline or wp:anchor container.

        Args:
            container: The wp:inline or wp:anchor element

        Returns:
            ImageType enum value
        """
        # Check for chart (c:chart namespace)
        chart_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
        if container.find(f".//{{{chart_ns}}}chart") is not None:
            return ImageType.CHART

        # Check for diagram/SmartArt (dgm namespace)
        dgm_ns = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
        if container.find(f".//{{{dgm_ns}}}relIds") is not None:
            return ImageType.DIAGRAM

        # Check for shape (wps namespace)
        wps_ns = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
        if container.find(f".//{{{wps_ns}}}wsp") is not None:
            return ImageType.SHAPE

        # Check for picture (pic:pic)
        if container.find(f".//{{{PIC_NAMESPACE}}}pic") is not None:
            return ImageType.IMAGE

        # Default to IMAGE for unknown drawing types
        return ImageType.IMAGE

    def _make_ref(
        self,
        image_type: ImageType,
        p_idx: int,
        img_idx: int,
        floating: bool,
    ) -> str:
        """Create a ref string for an image.

        Args:
            image_type: Type of image
            p_idx: Paragraph index
            img_idx: Image index within paragraph
            floating: Whether this is a floating image

        Returns:
            Ref string (e.g., "img:5/0", "img:5/f:0", "chart:12/0")
        """
        prefix_map = {
            ImageType.IMAGE: "img",
            ImageType.CHART: "chart",
            ImageType.DIAGRAM: "diagram",
            ImageType.SHAPE: "shape",
            ImageType.VML: "vml",
            ImageType.OLE_OBJECT: "obj",
        }

        prefix = prefix_map.get(image_type, "img")

        if floating:
            return f"{prefix}:{p_idx}/f:{img_idx}"
        else:
            return f"{prefix}:{p_idx}/{img_idx}"

    def _extract_extent(self, container: etree._Element) -> ImageSize | None:
        """Extract image dimensions from wp:extent.

        Args:
            container: The wp:inline or wp:anchor element

        Returns:
            ImageSize or None if not found
        """
        extent = container.find(wp("extent"))
        if extent is None:
            return None

        cx = extent.get("cx")
        cy = extent.get("cy")

        if cx is None or cy is None:
            return None

        try:
            return ImageSize(width_emu=int(cx), height_emu=int(cy))
        except ValueError:
            return None

    def _extract_doc_properties(self, container: etree._Element) -> tuple[str, str]:
        """Extract name and description from wp:docPr.

        Args:
            container: The wp:inline or wp:anchor element

        Returns:
            Tuple of (name, alt_text)
        """
        doc_pr = container.find(wp("docPr"))
        if doc_pr is None:
            return "", ""

        name = doc_pr.get("name", "")
        alt_text = doc_pr.get("descr", "")

        return name, alt_text

    def _extract_relationship_id(self, container: etree._Element) -> str:
        """Extract the relationship ID for the image.

        Args:
            container: The wp:inline or wp:anchor element

        Returns:
            Relationship ID string or empty string if not found
        """
        # Look for a:blip with r:embed attribute
        blip = container.find(f".//{{{A_NAMESPACE}}}blip")
        if blip is not None:
            rel_id = blip.get(f"{{{OFFICE_RELATIONSHIPS_NAMESPACE}}}embed")
            if rel_id:
                return rel_id

        return ""

    def _extract_position_info(self, anchor: etree._Element) -> ImagePosition:
        """Extract position information from a wp:anchor element.

        Args:
            anchor: The wp:anchor element

        Returns:
            ImagePosition with extracted info
        """
        position = ImagePosition()

        # Extract horizontal position
        pos_h = anchor.find(wp("positionH"))
        if pos_h is not None:
            position.relative_to = pos_h.get("relativeFrom", "")
            # Check for alignment vs offset
            align = pos_h.find(wp("align"))
            if align is not None and align.text:
                position.horizontal = align.text
            else:
                pos_offset = pos_h.find(wp("posOffset"))
                if pos_offset is not None and pos_offset.text:
                    position.horizontal = f"{int(pos_offset.text)} EMU"

        # Extract vertical position
        pos_v = anchor.find(wp("positionV"))
        if pos_v is not None:
            if not position.relative_to:
                position.relative_to = pos_v.get("relativeFrom", "")
            align = pos_v.find(wp("align"))
            if align is not None and align.text:
                position.vertical = align.text
            else:
                pos_offset = pos_v.find(wp("posOffset"))
                if pos_offset is not None and pos_offset.text:
                    position.vertical = f"{int(pos_offset.text)} EMU"

        # Extract wrap type
        wrap_types = [
            ("wrapNone", "none"),
            ("wrapSquare", "square"),
            ("wrapTight", "tight"),
            ("wrapThrough", "through"),
            ("wrapTopAndBottom", "topAndBottom"),
        ]

        for elem_name, wrap_name in wrap_types:
            if anchor.find(wp(elem_name)) is not None:
                position.wrap_type = wrap_name
                break

        return position


def get_images_from_document(
    xml_root: etree._Element,
    scope: str | None = None,
) -> list[ImageInfo]:
    """Get all images from a document.

    Args:
        xml_root: Root element of the document XML
        scope: Optional scope to limit search (e.g., "p:5", "tbl:0")

    Returns:
        List of ImageInfo objects
    """
    extractor = ImageExtractor(xml_root)
    images: list[ImageInfo] = []

    # Find body
    body = xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
    if body is None:
        return images

    # Iterate through body elements, tracking indices separately
    paragraph_index = 0
    table_index = 0
    for child in body:
        if child.tag == w("p"):
            para_images = extractor.extract_from_paragraph(child, paragraph_index)
            images.extend(para_images)
            paragraph_index += 1
        elif child.tag == w("tbl"):
            # Extract images from table cells
            table_images = _extract_images_from_table(extractor, child, table_index)
            images.extend(table_images)
            table_index += 1

    return images


def _extract_images_from_table(
    extractor: ImageExtractor,
    table: etree._Element,
    table_index: int,
) -> list[ImageInfo]:
    """Extract images from all cells in a table.

    Iterates through rows and cells, extracting images from paragraphs
    within each cell. Uses compound refs like tbl:0/row:1/cell:2/img:0.

    Args:
        extractor: ImageExtractor instance
        table: The w:tbl element
        table_index: Index of this table in the document body

    Returns:
        List of ImageInfo objects with compound refs
    """
    images: list[ImageInfo] = []
    table_ref = f"tbl:{table_index}"

    # Find all rows in the table
    for row_index, row in enumerate(table.findall(f"./{w('tr')}")):
        row_ref = f"{table_ref}/row:{row_index}"

        # Find all cells in the row
        for cell_index, cell in enumerate(row.findall(f"./{w('tc')}")):
            cell_ref = f"{row_ref}/cell:{cell_index}"

            # Extract images from paragraphs in this cell
            cell_images = _extract_cell_images(extractor, cell, cell_ref)
            images.extend(cell_images)

    return images


def _extract_cell_images(
    extractor: ImageExtractor,
    cell: etree._Element,
    cell_ref: str,
) -> list[ImageInfo]:
    """Extract images from a single table cell.

    Args:
        extractor: ImageExtractor instance
        cell: The w:tc element
        cell_ref: Base ref for this cell (e.g., "tbl:0/row:1/cell:2")

    Returns:
        List of ImageInfo objects
    """
    images: list[ImageInfo] = []

    # Track indices for each image type within the cell
    inline_index = 0
    floating_index = 0
    vml_index = 0
    obj_index = 0

    # Iterate through paragraphs in the cell
    for paragraph in cell.findall(f"./{w('p')}"):
        # Find all w:drawing elements (modern DrawingML)
        for drawing in paragraph.iter(w("drawing")):
            # Check for inline images (wp:inline)
            for inline in drawing.iter(wp("inline")):
                info = extractor._extract_inline_image(inline, 0, inline_index, cell_ref)
                if info:
                    # Override the ref with our compound table ref
                    info = _create_table_image_info(
                        info, cell_ref, "img", inline_index, floating=False
                    )
                    images.append(info)
                inline_index += 1

            # Check for floating/anchored images (wp:anchor)
            for anchor in drawing.iter(wp("anchor")):
                info = extractor._extract_anchor_image(anchor, 0, floating_index, cell_ref)
                if info:
                    info = _create_table_image_info(
                        info, cell_ref, "img", floating_index, floating=True
                    )
                    images.append(info)
                floating_index += 1

        # Find legacy VML graphics (w:pict)
        for pict in paragraph.iter(w("pict")):
            info = extractor._extract_vml_image(pict, 0, vml_index, cell_ref)
            if info:
                info = _create_table_image_info(info, cell_ref, "vml", vml_index, floating=False)
                images.append(info)
            vml_index += 1

        # Find OLE objects (w:object)
        for obj in paragraph.iter(w("object")):
            info = extractor._extract_ole_object(obj, 0, obj_index, cell_ref)
            if info:
                info = _create_table_image_info(info, cell_ref, "obj", obj_index, floating=False)
                images.append(info)
            obj_index += 1

    return images


def _create_table_image_info(
    source_info: ImageInfo,
    cell_ref: str,
    prefix: str,
    index: int,
    floating: bool,
) -> ImageInfo:
    """Create an ImageInfo with a compound table ref.

    Args:
        source_info: Original ImageInfo from extractor
        cell_ref: Cell ref (e.g., "tbl:0/row:1/cell:2")
        prefix: Image type prefix ("img", "vml", "obj", etc.)
        index: Image index within the cell
        floating: Whether this is a floating image

    Returns:
        New ImageInfo with compound ref
    """
    if floating:
        compound_ref = f"{cell_ref}/{prefix}:f:{index}"
    else:
        compound_ref = f"{cell_ref}/{prefix}:{index}"

    return ImageInfo(
        ref=compound_ref,
        image_type=source_info.image_type,
        position_type=source_info.position_type,
        name=source_info.name,
        alt_text=source_info.alt_text,
        size=source_info.size,
        format=source_info.format,
        relationship_id=source_info.relationship_id,
        paragraph_ref=cell_ref,
        position=source_info.position,
        _element=source_info._element,
    )
