"""
ImageOperations class for handling image insertion with tracked changes.

This module provides a dedicated class for all image operations,
including inserting images into documents with optional tracking.
"""

from __future__ import annotations

import logging
import random
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any

from lxml import etree

from ..constants import (
    A_NAMESPACE,
    OFFICE_RELATIONSHIPS_NAMESPACE,
    PIC_NAMESPACE,
    WORD_NAMESPACE,
    WP_NAMESPACE,
    a,
    pic,
    r,
    w,
    wp,
)
from ..content_types import ContentTypeManager, ContentTypes
from ..errors import AmbiguousTextError, TextNotFoundError
from ..relationships import RelationshipManager, RelationshipTypes
from ..scope import ScopeEvaluator
from ..suggestions import SuggestionGenerator

if TYPE_CHECKING:
    from ..document import Document
    from ..text_search import TextSpan

logger = logging.getLogger(__name__)

# Conversion constants
EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_PIXEL = 9525  # At 96 DPI


def _get_image_dimensions(image_path: Path) -> tuple[int, int] | None:
    """Try to get image dimensions using PIL/Pillow.

    Args:
        image_path: Path to the image file

    Returns:
        Tuple of (width, height) in pixels, or None if PIL is not available
    """
    try:
        from PIL import Image

        with Image.open(image_path) as img:
            return img.size
    except ImportError:
        return None
    except Exception:
        return None


def _generate_docpr_id() -> int:
    """Generate a random ID for docPr element."""
    return random.randint(1, 2147483647)


class ImageOperations:
    """Handles image insertion operations with tracked changes.

    This class encapsulates all image-related functionality, including:
    - Inserting inline images into documents
    - Wrapping image insertions in tracked changes
    - Managing image relationships and content types

    The class takes a Document reference and operates on its XML structure.

    Example:
        >>> # Usually accessed through Document
        >>> doc = Document("contract.docx")
        >>> doc.insert_image("logo.png", after="Company Name")
        >>> doc.insert_image_tracked("signature.png", after="Authorized By:")
    """

    def __init__(self, document: Document) -> None:
        """Initialize ImageOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    def _find_unique_match(
        self,
        text: str,
        scope: str | dict | Any | None,
        regex: bool,
        normalize_special_chars: bool,
    ) -> TextSpan:
        """Find a unique text match in the document.

        Args:
            text: The text or regex pattern to find
            scope: Limit search scope
            regex: Whether to treat text as regex
            normalize_special_chars: Whether to normalize quotes

        Returns:
            The single TextSpan match

        Raises:
            TextNotFoundError: If text is not found
            AmbiguousTextError: If multiple matches found
        """
        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(
            text,
            paragraphs,
            regex=regex,
            normalize_special_chars=normalize_special_chars and not regex,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(text, paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(text, matches)

        return matches[0]

    def _add_image_to_package(self, image_path: Path) -> str:
        """Add an image file to the document package.

        Args:
            image_path: Path to the image file

        Returns:
            The relative path of the image within the package (e.g., "media/image1.png")
        """
        if self._document._package is None:
            raise ValueError("Cannot add images to documents without a package")

        # Create word/media directory if it doesn't exist
        media_dir = self._document._package.temp_dir / "word" / "media"
        media_dir.mkdir(parents=True, exist_ok=True)

        # Find next available image number
        existing_images = list(media_dir.glob("image*"))
        next_num = 1
        for img in existing_images:
            try:
                # Extract number from "image1.png" etc.
                num = int(img.stem.replace("image", ""))
                if num >= next_num:
                    next_num = num + 1
            except ValueError:
                pass

        # Copy image to media folder with new name
        extension = image_path.suffix.lower()
        new_name = f"image{next_num}{extension}"
        dest_path = media_dir / new_name
        shutil.copy2(image_path, dest_path)

        return f"media/{new_name}"

    def _ensure_content_type(self, extension: str) -> None:
        """Ensure the content type for an image extension is registered.

        Args:
            extension: File extension without dot (e.g., "png")
        """
        if self._document._package is None:
            return

        ct_manager = ContentTypeManager(self._document._package)
        ext_lower = extension.lower()

        # Get the content type for this extension
        content_type = ContentTypes.IMAGE_EXTENSION_MAP.get(ext_lower)
        if content_type:
            ct_manager.add_default(ext_lower, content_type)
            ct_manager.save()

    def _add_image_relationship(self, image_target: str) -> str:
        """Add a relationship for the image.

        Args:
            image_target: Relative path to the image (e.g., "media/image1.png")

        Returns:
            The relationship ID (e.g., "rId5")
        """
        if self._document._package is None:
            raise ValueError("Cannot add relationships to documents without a package")

        rel_manager = RelationshipManager(self._document._package, "word/document.xml")
        rel_id = rel_manager.add_unique_relationship(RelationshipTypes.IMAGE, image_target)
        rel_manager.save()

        return rel_id

    def _create_inline_drawing(
        self,
        rel_id: str,
        width_emu: int,
        height_emu: int,
        name: str = "Picture",
        description: str = "",
    ) -> etree._Element:
        """Create the inline drawing XML element for an image.

        Args:
            rel_id: The relationship ID for the image
            width_emu: Width in EMUs (English Metric Units)
            height_emu: Height in EMUs
            name: Name for the image (used in Word UI)
            description: Alt text description

        Returns:
            The w:drawing element containing the inline image
        """
        # Generate unique IDs
        doc_pr_id = _generate_docpr_id()

        # Build the namespace map for the drawing elements
        nsmap = {
            "wp": WP_NAMESPACE,
            "a": A_NAMESPACE,
            "pic": PIC_NAMESPACE,
            "r": OFFICE_RELATIONSHIPS_NAMESPACE,
        }

        # Create <w:drawing>
        drawing = etree.Element(w("drawing"))

        # Create <wp:inline>
        inline = etree.SubElement(
            drawing,
            wp("inline"),
            nsmap=nsmap,
            attrib={
                "distT": "0",
                "distB": "0",
                "distL": "0",
                "distR": "0",
            },
        )

        # <wp:extent cx="..." cy="..."/>
        etree.SubElement(
            inline,
            wp("extent"),
            attrib={"cx": str(width_emu), "cy": str(height_emu)},
        )

        # <wp:effectExtent l="0" t="0" r="0" b="0"/>
        etree.SubElement(
            inline,
            wp("effectExtent"),
            attrib={"l": "0", "t": "0", "r": "0", "b": "0"},
        )

        # <wp:docPr id="..." name="..." descr="..."/>
        doc_pr_attrib = {"id": str(doc_pr_id), "name": name}
        if description:
            doc_pr_attrib["descr"] = description
        etree.SubElement(inline, wp("docPr"), attrib=doc_pr_attrib)

        # <wp:cNvGraphicFramePr>
        cnv_frame_pr = etree.SubElement(inline, wp("cNvGraphicFramePr"))
        etree.SubElement(
            cnv_frame_pr,
            a("graphicFrameLocks"),
            attrib={"noChangeAspect": "1"},
        )

        # <a:graphic>
        graphic = etree.SubElement(inline, a("graphic"))

        # <a:graphicData uri="...">
        graphic_data = etree.SubElement(
            graphic,
            a("graphicData"),
            attrib={"uri": PIC_NAMESPACE},
        )

        # <pic:pic>
        pic_elem = etree.SubElement(graphic_data, pic("pic"))

        # <pic:nvPicPr>
        nv_pic_pr = etree.SubElement(pic_elem, pic("nvPicPr"))
        etree.SubElement(
            nv_pic_pr,
            pic("cNvPr"),
            attrib={"id": str(doc_pr_id), "name": name},
        )
        cnv_pic_pr = etree.SubElement(nv_pic_pr, pic("cNvPicPr"))
        etree.SubElement(cnv_pic_pr, a("picLocks"), attrib={"noChangeAspect": "1"})

        # <pic:blipFill>
        blip_fill = etree.SubElement(pic_elem, pic("blipFill"))
        etree.SubElement(
            blip_fill,
            a("blip"),
            attrib={r("embed"): rel_id},
        )
        stretch = etree.SubElement(blip_fill, a("stretch"))
        etree.SubElement(stretch, a("fillRect"))

        # <pic:spPr>
        sp_pr = etree.SubElement(pic_elem, pic("spPr"))
        xfrm = etree.SubElement(sp_pr, a("xfrm"))
        etree.SubElement(xfrm, a("off"), attrib={"x": "0", "y": "0"})
        etree.SubElement(xfrm, a("ext"), attrib={"cx": str(width_emu), "cy": str(height_emu)})
        prst_geom = etree.SubElement(sp_pr, a("prstGeom"), attrib={"prst": "rect"})
        etree.SubElement(prst_geom, a("avLst"))

        return drawing

    def insert(
        self,
        image_path: str | Path,
        after: str | None = None,
        before: str | None = None,
        width_inches: float | None = None,
        height_inches: float | None = None,
        width_cm: float | None = None,
        height_cm: float | None = None,
        name: str | None = None,
        description: str = "",
        scope: str | dict | Any | None = None,
        regex: bool = False,
    ) -> None:
        """Insert an image into the document.

        This method searches for the anchor text in the document and inserts
        the image either immediately after it or immediately before it.

        Args:
            image_path: Path to the image file (PNG, JPEG, GIF, etc.)
            after: The text to insert after (optional)
            before: The text to insert before (optional)
            width_inches: Width in inches (auto-calculated if not provided)
            height_inches: Height in inches (auto-calculated if not provided)
            width_cm: Width in centimeters (alternative to inches)
            height_cm: Height in centimeters (alternative to inches)
            name: Display name for the image (defaults to filename)
            description: Alt text description for accessibility
            scope: Limit search scope
            regex: Whether to treat anchor as regex pattern

        Raises:
            ValueError: If both 'after' and 'before' are specified, or neither
            TextNotFoundError: If the anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor are found
            FileNotFoundError: If the image file doesn't exist
        """
        # Validate parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"Image file not found: {image_path}")

        # Find the anchor text
        anchor_text = after if after is not None else before
        assert anchor_text is not None  # Type guard
        match = self._find_unique_match(anchor_text, scope, regex, True)

        # Calculate dimensions
        width_emu, height_emu = self._calculate_dimensions(
            image_path, width_inches, height_inches, width_cm, height_cm
        )

        # Set up image name
        if name is None:
            name = image_path.stem

        # Add image to package
        image_target = self._add_image_to_package(image_path)

        # Ensure content type is registered
        extension = image_path.suffix.lstrip(".").lower()
        self._ensure_content_type(extension)

        # Add relationship
        rel_id = self._add_image_relationship(image_target)

        # Create the drawing element
        drawing = self._create_inline_drawing(rel_id, width_emu, height_emu, name, description)

        # Create a run containing the drawing
        run = etree.Element(w("r"))
        run.append(drawing)

        # Insert the run at the appropriate location
        self._insert_element_at_match(run, match, after is not None)

    def insert_tracked(
        self,
        image_path: str | Path,
        after: str | None = None,
        before: str | None = None,
        width_inches: float | None = None,
        height_inches: float | None = None,
        width_cm: float | None = None,
        height_cm: float | None = None,
        name: str | None = None,
        description: str = "",
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
    ) -> None:
        """Insert an image with tracked changes.

        This method wraps the image insertion in a <w:ins> element so it
        appears as a tracked insertion in Word's review pane.

        Args:
            image_path: Path to the image file (PNG, JPEG, GIF, etc.)
            after: The text to insert after (optional)
            before: The text to insert before (optional)
            width_inches: Width in inches (auto-calculated if not provided)
            height_inches: Height in inches (auto-calculated if not provided)
            width_cm: Width in centimeters (alternative to inches)
            height_cm: Height in centimeters (alternative to inches)
            name: Display name for the image (defaults to filename)
            description: Alt text description for accessibility
            author: Author for the tracked change (uses document author if None)
            scope: Limit search scope
            regex: Whether to treat anchor as regex pattern

        Raises:
            ValueError: If both 'after' and 'before' are specified, or neither
            TextNotFoundError: If the anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor are found
            FileNotFoundError: If the image file doesn't exist
        """
        # Validate parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"Image file not found: {image_path}")

        # Find the anchor text
        anchor_text = after if after is not None else before
        assert anchor_text is not None  # Type guard
        match = self._find_unique_match(anchor_text, scope, regex, True)

        # Calculate dimensions
        width_emu, height_emu = self._calculate_dimensions(
            image_path, width_inches, height_inches, width_cm, height_cm
        )

        # Set up image name
        if name is None:
            name = image_path.stem

        # Add image to package
        image_target = self._add_image_to_package(image_path)

        # Ensure content type is registered
        extension = image_path.suffix.lstrip(".").lower()
        self._ensure_content_type(extension)

        # Add relationship
        rel_id = self._add_image_relationship(image_target)

        # Create the drawing element
        drawing = self._create_inline_drawing(rel_id, width_emu, height_emu, name, description)

        # Create a run containing the drawing
        run = etree.Element(w("r"))
        run.append(drawing)

        # Wrap in w:ins for tracked change
        ins_elem = self._create_tracked_insertion(run, author)

        # Insert the element at the appropriate location
        self._insert_element_at_match(ins_elem, match, after is not None)

    def _calculate_dimensions(
        self,
        image_path: Path,
        width_inches: float | None,
        height_inches: float | None,
        width_cm: float | None,
        height_cm: float | None,
    ) -> tuple[int, int]:
        """Calculate image dimensions in EMUs.

        Args:
            image_path: Path to the image
            width_inches: Width in inches (optional)
            height_inches: Height in inches (optional)
            width_cm: Width in cm (optional)
            height_cm: Height in cm (optional)

        Returns:
            Tuple of (width_emu, height_emu)
        """
        # If dimensions provided in cm, convert to inches
        if width_cm is not None:
            width_inches = width_cm / 2.54
        if height_cm is not None:
            height_inches = height_cm / 2.54

        # If dimensions provided, use them
        if width_inches is not None and height_inches is not None:
            return int(width_inches * EMU_PER_INCH), int(height_inches * EMU_PER_INCH)

        # Try to get actual image dimensions
        dimensions = _get_image_dimensions(image_path)
        if dimensions:
            pixel_width, pixel_height = dimensions
            aspect_ratio = pixel_width / pixel_height

            if width_inches is not None:
                # Scale height based on width
                height_inches = width_inches / aspect_ratio
            elif height_inches is not None:
                # Scale width based on height
                width_inches = height_inches * aspect_ratio
            else:
                # Default: use actual pixel size at 96 DPI
                width_emu = pixel_width * EMU_PER_PIXEL
                height_emu = pixel_height * EMU_PER_PIXEL
                return width_emu, height_emu

            return int(width_inches * EMU_PER_INCH), int(height_inches * EMU_PER_INCH)

        # Default dimensions if we can't detect (2 inches square)
        default_size = int(2 * EMU_PER_INCH)
        width_emu = int(width_inches * EMU_PER_INCH) if width_inches else default_size
        height_emu = int(height_inches * EMU_PER_INCH) if height_inches else default_size

        return width_emu, height_emu

    def _create_tracked_insertion(
        self,
        content: etree._Element,
        author: str | None,
    ) -> etree._Element:
        """Create a w:ins element wrapping the content.

        Args:
            content: The element to wrap (typically a w:r)
            author: Author name (uses document author if None)

        Returns:
            The w:ins element containing the content
        """
        author = author if author is not None else self._document.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        # Get next change ID from the XML generator
        change_id = self._document._xml_generator.next_change_id
        self._document._xml_generator.next_change_id += 1

        # Create w:ins element
        ins_elem = etree.Element(w("ins"))
        ins_elem.set(w("id"), str(change_id))
        ins_elem.set(w("author"), author)
        ins_elem.set(w("date"), timestamp)

        # Add content
        ins_elem.append(content)

        return ins_elem

    def _insert_element_at_match(
        self,
        element: etree._Element,
        match: TextSpan,
        insert_after: bool,
    ) -> None:
        """Insert an element at a text match location.

        Args:
            element: The element to insert
            match: The TextSpan match indicating the location
            insert_after: If True, insert after the match; if False, before
        """
        if insert_after:
            # Insert after the last run of the match
            target_run = match.runs[match.end_run_index]
            parent = target_run.getparent()
            if parent is not None:
                idx = list(parent).index(target_run)
                parent.insert(idx + 1, element)
        else:
            # Insert before the first run of the match
            target_run = match.runs[match.start_run_index]
            parent = target_run.getparent()
            if parent is not None:
                idx = list(parent).index(target_run)
                parent.insert(idx, element)
