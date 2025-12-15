"""
Test the image insertion API.

Tests the insert_image() and insert_image_tracked() methods.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from python_docx_redline import Document
from python_docx_redline.errors import TextNotFoundError

# Namespace constants for assertions
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def create_test_document() -> Path:
    """Create a minimal test document."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>Company Name: Test Corp</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>Figure 1: Chart description here.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>Authorized By: _____________</w:t></w:r>
</w:p>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_test_image(suffix: str = ".png") -> Path:
    """Create a minimal test image (1x1 pixel PNG)."""
    image_path = Path(tempfile.mktemp(suffix=suffix))

    # Minimal 1x1 pixel PNG (red)
    png_data = bytes(
        [
            0x89,
            0x50,
            0x4E,
            0x47,
            0x0D,
            0x0A,
            0x1A,
            0x0A,  # PNG signature
            0x00,
            0x00,
            0x00,
            0x0D,
            0x49,
            0x48,
            0x44,
            0x52,  # IHDR chunk
            0x00,
            0x00,
            0x00,
            0x01,
            0x00,
            0x00,
            0x00,
            0x01,  # 1x1 pixel
            0x08,
            0x02,
            0x00,
            0x00,
            0x00,
            0x90,
            0x77,
            0x53,  # 8-bit RGB
            0xDE,
            0x00,
            0x00,
            0x00,
            0x0C,
            0x49,
            0x44,
            0x41,  # IDAT chunk
            0x54,
            0x08,
            0xD7,
            0x63,
            0xF8,
            0xCF,
            0xC0,
            0x00,
            0x00,
            0x00,
            0x03,
            0x00,
            0x01,
            0x00,
            0x05,
            0xFE,
            0xD4,
            0xEF,
            0x00,
            0x00,
            0x00,
            0x00,
            0x49,
            0x45,  # IEND chunk
            0x4E,
            0x44,
            0xAE,
            0x42,
            0x60,
            0x82,
        ]
    )

    with open(image_path, "wb") as f:
        f.write(png_data)

    return image_path


class TestInsertImage:
    """Tests for Document.insert_image() method."""

    def test_insert_image_basic(self) -> None:
        """Test basic image insertion."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:")

            # Verify drawing element exists
            drawings = list(doc.xml_root.iter(f"{{{WORD_NS}}}drawing"))
            assert len(drawings) == 1

            # Verify inline element exists
            inlines = list(doc.xml_root.iter(f"{{{WP_NS}}}inline"))
            assert len(inlines) == 1

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_before(self) -> None:
        """Test inserting image before anchor text."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, before="Figure 1:")

            drawings = list(doc.xml_root.iter(f"{{{WORD_NS}}}drawing"))
            assert len(drawings) == 1

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_with_dimensions_inches(self) -> None:
        """Test image insertion with specified dimensions in inches."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:", width_inches=2.0, height_inches=1.5)

            # Check dimensions in EMUs (914400 per inch)
            extents = list(doc.xml_root.iter(f"{{{WP_NS}}}extent"))
            assert len(extents) == 1

            cx = int(extents[0].get("cx"))
            cy = int(extents[0].get("cy"))

            # 2.0 inches * 914400 = 1828800 EMU
            assert cx == 1828800
            # 1.5 inches * 914400 = 1371600 EMU
            assert cy == 1371600

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_with_dimensions_cm(self) -> None:
        """Test image insertion with specified dimensions in centimeters."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:", width_cm=5.08, height_cm=2.54)

            # 5.08 cm = 2 inches = 1828800 EMU
            # 2.54 cm = 1 inch = 914400 EMU
            extents = list(doc.xml_root.iter(f"{{{WP_NS}}}extent"))
            assert len(extents) == 1

            cx = int(extents[0].get("cx"))
            cy = int(extents[0].get("cy"))

            # Allow small rounding differences
            assert abs(cx - 1828800) < 100
            assert abs(cy - 914400) < 100

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_custom_name(self) -> None:
        """Test image insertion with custom name."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:", name="Logo")

            # Check docPr element has correct name
            doc_prs = list(doc.xml_root.iter(f"{{{WP_NS}}}docPr"))
            assert len(doc_prs) == 1
            assert doc_prs[0].get("name") == "Logo"

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_description(self) -> None:
        """Test image insertion with alt text description."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:", description="Company logo image")

            doc_prs = list(doc.xml_root.iter(f"{{{WP_NS}}}docPr"))
            assert len(doc_prs) == 1
            assert doc_prs[0].get("descr") == "Company logo image"

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_creates_media_folder(self) -> None:
        """Test that image insertion creates word/media folder."""
        doc_path = create_test_document()
        image_path = create_test_image()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:")
            doc.save(output_path)

            # Verify media folder contains image
            with zipfile.ZipFile(output_path, "r") as docx:
                media_files = [n for n in docx.namelist() if n.startswith("word/media/")]
                assert len(media_files) == 1
                assert media_files[0].endswith(".png")

        finally:
            doc_path.unlink()
            image_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_image_adds_content_type(self) -> None:
        """Test that image insertion adds content type for extension."""
        doc_path = create_test_document()
        image_path = create_test_image()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:")
            doc.save(output_path)

            with zipfile.ZipFile(output_path, "r") as docx:
                ct_content = docx.read("[Content_Types].xml").decode("utf-8")
                assert "image/png" in ct_content

        finally:
            doc_path.unlink()
            image_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_image_adds_relationship(self) -> None:
        """Test that image insertion adds relationship entry."""
        doc_path = create_test_document()
        image_path = create_test_image()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:")
            doc.save(output_path)

            with zipfile.ZipFile(output_path, "r") as docx:
                rels_content = docx.read("word/_rels/document.xml.rels").decode("utf-8")
                assert "relationships/image" in rels_content
                assert "media/image" in rels_content

        finally:
            doc_path.unlink()
            image_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_image_text_not_found(self) -> None:
        """Test that insert_image raises error for missing anchor text."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)

            with pytest.raises(TextNotFoundError):
                doc.insert_image(image_path, after="nonexistent text")

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_file_not_found(self) -> None:
        """Test that insert_image raises error for missing image file."""
        doc_path = create_test_document()

        try:
            doc = Document(doc_path)

            with pytest.raises(FileNotFoundError):
                doc.insert_image("/nonexistent/image.png", after="Company Name:")

        finally:
            doc_path.unlink()

    def test_insert_image_requires_anchor(self) -> None:
        """Test that insert_image requires either 'after' or 'before'."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError, match="Must specify"):
                doc.insert_image(image_path)

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_not_both_anchors(self) -> None:
        """Test that insert_image rejects both 'after' and 'before'."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)

            with pytest.raises(ValueError, match="Cannot specify both"):
                doc.insert_image(image_path, after="Company", before="Name")

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_multiple_images(self) -> None:
        """Test inserting multiple images into a document."""
        doc_path = create_test_document()
        image1_path = create_test_image()
        image2_path = create_test_image()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.insert_image(image1_path, after="Company Name:")
            doc.insert_image(image2_path, after="Figure 1:")
            doc.save(output_path)

            # Verify two drawings
            drawings = list(doc.xml_root.iter(f"{{{WORD_NS}}}drawing"))
            assert len(drawings) == 2

            # Verify two media files
            with zipfile.ZipFile(output_path, "r") as docx:
                media_files = [n for n in docx.namelist() if n.startswith("word/media/")]
                assert len(media_files) == 2

        finally:
            doc_path.unlink()
            image1_path.unlink()
            image2_path.unlink()
            if output_path.exists():
                output_path.unlink()


class TestInsertImageTracked:
    """Tests for Document.insert_image_tracked() method."""

    def test_insert_image_tracked_basic(self) -> None:
        """Test tracked image insertion creates w:ins wrapper."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path, author="Test Author")
            doc.insert_image_tracked(image_path, after="Company Name:")

            # Verify w:ins element exists
            insertions = list(doc.xml_root.iter(f"{{{WORD_NS}}}ins"))
            assert len(insertions) == 1

            # Verify drawing is inside ins
            ins_elem = insertions[0]
            drawings = list(ins_elem.iter(f"{{{WORD_NS}}}drawing"))
            assert len(drawings) == 1

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_tracked_has_author(self) -> None:
        """Test that tracked insertion has author attribute."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path, author="Test Author")
            doc.insert_image_tracked(image_path, after="Company Name:")

            insertions = list(doc.xml_root.iter(f"{{{WORD_NS}}}ins"))
            assert len(insertions) == 1

            author = insertions[0].get(f"{{{WORD_NS}}}author")
            assert author == "Test Author"

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_tracked_custom_author(self) -> None:
        """Test tracked insertion with custom author override."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path, author="Doc Author")
            doc.insert_image_tracked(image_path, after="Company Name:", author="Custom Author")

            insertions = list(doc.xml_root.iter(f"{{{WORD_NS}}}ins"))
            assert len(insertions) == 1

            author = insertions[0].get(f"{{{WORD_NS}}}author")
            assert author == "Custom Author"

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_tracked_has_date(self) -> None:
        """Test that tracked insertion has date attribute."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)
            doc.insert_image_tracked(image_path, after="Company Name:")

            insertions = list(doc.xml_root.iter(f"{{{WORD_NS}}}ins"))
            assert len(insertions) == 1

            date = insertions[0].get(f"{{{WORD_NS}}}date")
            assert date is not None
            # Should be ISO format: 2025-01-15T10:30:00Z
            assert "T" in date
            assert date.endswith("Z")

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_tracked_has_id(self) -> None:
        """Test that tracked insertion has change ID."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)
            doc.insert_image_tracked(image_path, after="Company Name:")

            insertions = list(doc.xml_root.iter(f"{{{WORD_NS}}}ins"))
            assert len(insertions) == 1

            change_id = insertions[0].get(f"{{{WORD_NS}}}id")
            assert change_id is not None
            assert change_id.isdigit()

        finally:
            doc_path.unlink()
            image_path.unlink()

    def test_insert_image_tracked_persists(self) -> None:
        """Test that tracked image insertion persists after save/reload."""
        doc_path = create_test_document()
        image_path = create_test_image()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Create and save
            doc = Document(doc_path, author="Test Author")
            doc.insert_image_tracked(image_path, after="Company Name:")
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)
            insertions = list(doc2.xml_root.iter(f"{{{WORD_NS}}}ins"))
            assert len(insertions) == 1

            drawings = list(insertions[0].iter(f"{{{WORD_NS}}}drawing"))
            assert len(drawings) == 1

        finally:
            doc_path.unlink()
            image_path.unlink()
            if output_path.exists():
                output_path.unlink()

    def test_insert_image_tracked_with_dimensions(self) -> None:
        """Test tracked image insertion with custom dimensions."""
        doc_path = create_test_document()
        image_path = create_test_image()

        try:
            doc = Document(doc_path)
            doc.insert_image_tracked(
                image_path, after="Company Name:", width_inches=3.0, height_inches=2.0
            )

            extents = list(doc.xml_root.iter(f"{{{WP_NS}}}extent"))
            assert len(extents) == 1

            cx = int(extents[0].get("cx"))
            cy = int(extents[0].get("cy"))

            # 3.0 inches * 914400 = 2743200 EMU
            assert cx == 2743200
            # 2.0 inches * 914400 = 1828800 EMU
            assert cy == 1828800

        finally:
            doc_path.unlink()
            image_path.unlink()


class TestImageFormats:
    """Tests for different image format support."""

    def test_insert_jpeg_image(self) -> None:
        """Test inserting a JPEG image."""
        doc_path = create_test_document()
        # Create a minimal JPEG (just for testing content type registration)
        image_path = Path(tempfile.mktemp(suffix=".jpg"))
        # Minimal JPEG header
        jpeg_data = bytes(
            [
                0xFF,
                0xD8,
                0xFF,
                0xE0,
                0x00,
                0x10,
                0x4A,
                0x46,
                0x49,
                0x46,
                0x00,
                0x01,
                0x01,
                0x00,
                0x00,
                0x01,
                0x00,
                0x01,
                0x00,
                0x00,
                0xFF,
                0xDB,
                0x00,
                0x43,
            ]
            + [0x00] * 64
            + [0xFF, 0xD9]
        )

        with open(image_path, "wb") as f:
            f.write(jpeg_data)

        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:", width_inches=1.0, height_inches=1.0)
            doc.save(output_path)

            with zipfile.ZipFile(output_path, "r") as docx:
                ct_content = docx.read("[Content_Types].xml").decode("utf-8")
                assert "image/jpeg" in ct_content

        finally:
            doc_path.unlink()
            image_path.unlink()
            if output_path.exists():
                output_path.unlink()


class TestImagePackageIntegrity:
    """Tests for package integrity after image operations."""

    def test_saved_document_opens_in_word(self) -> None:
        """Test that saved document has valid structure for Word."""
        doc_path = create_test_document()
        image_path = create_test_image()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            doc = Document(doc_path)
            doc.insert_image(image_path, after="Company Name:", width_inches=1.0)
            doc.save(output_path)

            # Verify all required parts exist
            with zipfile.ZipFile(output_path, "r") as docx:
                names = docx.namelist()

                # Core parts
                assert "[Content_Types].xml" in names
                assert "_rels/.rels" in names
                assert "word/document.xml" in names
                assert "word/_rels/document.xml.rels" in names

                # Image-related
                media_files = [n for n in names if n.startswith("word/media/")]
                assert len(media_files) == 1

                # Verify document.xml is valid XML
                doc_xml = docx.read("word/document.xml")
                etree.fromstring(doc_xml)  # Should not raise

        finally:
            doc_path.unlink()
            image_path.unlink()
            if output_path.exists():
                output_path.unlink()
