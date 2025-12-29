"""
Tests for image and embedded object refs in the accessibility layer.

These tests verify:
- Image extraction from documents
- ImageInfo dataclass fields
- Image refs in accessibility tree
- YAML serialization of images
- get_images() method
"""

from lxml import etree

from python_docx_redline.accessibility import (
    AccessibilityTree,
    ElementType,
    ImageExtractor,
    ImageInfo,
    ImagePositionType,
    ImageSize,
    ImageType,
    ViewMode,
)

# Test XML documents with images

DOCUMENT_WITH_INLINE_IMAGE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Here is an image: </w:t>
      </w:r>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="1" name="Picture 1" descr="Logo image"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed="rId4"/>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>More text after the image.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_FLOATING_IMAGE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="0" distR="0"
                     simplePos="0" relativeHeight="1" behindDoc="0"
                     locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column">
              <wp:align>left</wp:align>
            </wp:positionH>
            <wp:positionV relativeFrom="paragraph">
              <wp:align>top</wp:align>
            </wp:positionV>
            <wp:extent cx="1828800" cy="914400"/>
            <wp:wrapSquare wrapText="bothSides"/>
            <wp:docPr id="2" name="Floating Image" descr="A floating chart"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed="rId5"/>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
      <w:r>
        <w:t>Text with floating image.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_MULTIPLE_IMAGES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First paragraph with image: </w:t>
      </w:r>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="457200" cy="457200"/>
            <wp:docPr id="1" name="Image 1" descr="First image"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed="rId1"/>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph no images.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="914400" cy="685800"/>
            <wp:docPr id="2" name="Image 2"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed="rId2"/>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="1371600" cy="1143000"/>
            <wp:docPr id="3" name="Image 3" descr="Third image description"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed="rId3"/>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_WITH_VML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p>
      <w:r>
        <w:pict>
          <v:shape id="Shape1" style="width:100pt;height:50pt"/>
        </w:pict>
      </w:r>
      <w:r>
        <w:t>Paragraph with VML shape.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_NO_IMAGES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Simple text paragraph.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Another paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


def create_tree_from_xml(xml_content: str, view_mode: ViewMode | None = None) -> AccessibilityTree:
    """Create an AccessibilityTree from raw XML content."""
    root = etree.fromstring(xml_content.encode("utf-8"))
    return AccessibilityTree.from_xml(root, view_mode=view_mode)


class TestImageInfoDataclass:
    """Tests for the ImageInfo dataclass."""

    def test_image_info_basic(self) -> None:
        """Test basic ImageInfo creation."""
        info = ImageInfo(
            ref="img:0/0",
            image_type=ImageType.IMAGE,
            position_type=ImagePositionType.INLINE,
        )

        assert info.ref == "img:0/0"
        assert info.image_type == ImageType.IMAGE
        assert info.is_inline
        assert not info.is_floating

    def test_image_info_with_metadata(self) -> None:
        """Test ImageInfo with name and alt_text."""
        info = ImageInfo(
            ref="img:1/0",
            image_type=ImageType.IMAGE,
            position_type=ImagePositionType.INLINE,
            name="Company Logo",
            alt_text="Logo of the company",
        )

        assert info.name == "Company Logo"
        assert info.alt_text == "Logo of the company"

    def test_image_info_floating(self) -> None:
        """Test floating image info."""
        info = ImageInfo(
            ref="img:0/f:0",
            image_type=ImageType.IMAGE,
            position_type=ImagePositionType.FLOATING,
        )

        assert info.is_floating
        assert not info.is_inline

    def test_image_info_to_yaml_dict_content(self) -> None:
        """Test to_yaml_dict in content mode."""
        size = ImageSize(width_emu=914400, height_emu=914400)
        info = ImageInfo(
            ref="img:0/0",
            image_type=ImageType.IMAGE,
            position_type=ImagePositionType.INLINE,
            name="Test Image",
            alt_text="Test alt text",
            size=size,
        )

        yaml_dict = info.to_yaml_dict(mode="content")

        assert yaml_dict["ref"] == "img:0/0"
        assert yaml_dict["type"] == "image"
        assert yaml_dict["position_type"] == "inline"
        assert yaml_dict["name"] == "Test Image"
        assert yaml_dict["alt_text"] == "Test alt text"
        assert yaml_dict["size"] == "1.0in x 1.0in"


class TestImageSize:
    """Tests for ImageSize dataclass."""

    def test_image_size_basic(self) -> None:
        """Test basic ImageSize creation."""
        size = ImageSize(width_emu=914400, height_emu=914400)

        assert size.width_emu == 914400
        assert size.height_emu == 914400

    def test_image_size_inches(self) -> None:
        """Test inch conversion."""
        # 914400 EMU = 1 inch
        size = ImageSize(width_emu=914400, height_emu=1828800)

        assert size.width_inches == 1.0
        assert size.height_inches == 2.0

    def test_image_size_cm(self) -> None:
        """Test centimeter conversion."""
        # 360000 EMU = 1 cm
        size = ImageSize(width_emu=360000, height_emu=720000)

        assert size.width_cm == 1.0
        assert size.height_cm == 2.0

    def test_image_size_display_string(self) -> None:
        """Test display string format."""
        size = ImageSize(width_emu=914400, height_emu=685800)

        display = size.to_display_string()
        assert display == "1.0in x 0.8in"


class TestImageExtractor:
    """Tests for ImageExtractor class."""

    def test_extract_inline_image(self) -> None:
        """Test extracting inline image from paragraph."""
        root = etree.fromstring(DOCUMENT_WITH_INLINE_IMAGE.encode("utf-8"))
        extractor = ImageExtractor(root)

        # Get first paragraph
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paragraphs = root.findall(".//w:p", namespaces=nsmap)

        images = extractor.extract_from_paragraph(paragraphs[0], 0)

        assert len(images) == 1
        img = images[0]
        assert img.ref == "img:0/0"
        assert img.image_type == ImageType.IMAGE
        assert img.position_type == ImagePositionType.INLINE
        assert img.name == "Picture 1"
        assert img.alt_text == "Logo image"
        assert img.size is not None
        assert img.size.width_emu == 914400
        assert img.size.height_emu == 914400
        assert img.relationship_id == "rId4"

    def test_extract_floating_image(self) -> None:
        """Test extracting floating image from paragraph."""
        root = etree.fromstring(DOCUMENT_WITH_FLOATING_IMAGE.encode("utf-8"))
        extractor = ImageExtractor(root)

        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paragraphs = root.findall(".//w:p", namespaces=nsmap)

        images = extractor.extract_from_paragraph(paragraphs[0], 0)

        assert len(images) == 1
        img = images[0]
        assert img.ref == "img:0/f:0"
        assert img.image_type == ImageType.IMAGE
        assert img.position_type == ImagePositionType.FLOATING
        assert img.name == "Floating Image"
        assert img.alt_text == "A floating chart"
        assert img.size is not None
        assert img.size.width_emu == 1828800
        assert img.size.height_emu == 914400
        assert img.position is not None
        assert img.position.horizontal == "left"
        assert img.position.vertical == "top"
        assert img.position.wrap_type == "square"

    def test_extract_vml_image(self) -> None:
        """Test extracting VML graphics."""
        root = etree.fromstring(DOCUMENT_WITH_VML.encode("utf-8"))
        extractor = ImageExtractor(root)

        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paragraphs = root.findall(".//w:p", namespaces=nsmap)

        images = extractor.extract_from_paragraph(paragraphs[0], 0)

        assert len(images) == 1
        img = images[0]
        assert img.ref == "vml:0/0"
        assert img.image_type == ImageType.VML
        assert img.position_type == ImagePositionType.INLINE

    def test_extract_no_images(self) -> None:
        """Test extracting from paragraph with no images."""
        root = etree.fromstring(DOCUMENT_NO_IMAGES.encode("utf-8"))
        extractor = ImageExtractor(root)

        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paragraphs = root.findall(".//w:p", namespaces=nsmap)

        images = extractor.extract_from_paragraph(paragraphs[0], 0)

        assert len(images) == 0


class TestAccessibilityTreeWithImages:
    """Tests for images in AccessibilityTree."""

    def test_tree_stats_include_images(self) -> None:
        """Test that tree stats include image count."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INLINE_IMAGE)

        assert tree.stats.images == 1

    def test_tree_stats_multiple_images(self) -> None:
        """Test stats with multiple images."""
        tree = create_tree_from_xml(DOCUMENT_WITH_MULTIPLE_IMAGES)

        # 1 in first paragraph, 2 in third paragraph
        assert tree.stats.images == 3

    def test_tree_stats_no_images(self) -> None:
        """Test stats when no images present."""
        tree = create_tree_from_xml(DOCUMENT_NO_IMAGES)

        assert tree.stats.images == 0

    def test_paragraph_node_has_images(self) -> None:
        """Test that paragraph node has images property."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INLINE_IMAGE)

        paragraphs = tree.find_all(element_type=ElementType.PARAGRAPH)
        assert len(paragraphs) == 2

        # First paragraph has image
        assert paragraphs[0].has_images
        assert len(paragraphs[0].images) == 1
        assert paragraphs[0].properties.get("has_images") == "true"
        assert paragraphs[0].properties.get("image_count") == "1"

        # Second paragraph has no images
        assert not paragraphs[1].has_images
        assert len(paragraphs[1].images) == 0

    def test_get_images_method(self) -> None:
        """Test get_images() returns all images."""
        tree = create_tree_from_xml(DOCUMENT_WITH_MULTIPLE_IMAGES)

        images = tree.get_images()

        assert len(images) == 3
        assert images[0].ref == "img:0/0"
        assert images[1].ref == "img:2/0"
        assert images[2].ref == "img:2/1"

    def test_get_image_by_ref(self) -> None:
        """Test get_image() returns specific image."""
        tree = create_tree_from_xml(DOCUMENT_WITH_MULTIPLE_IMAGES)

        img = tree.get_image("img:2/1")

        assert img is not None
        assert img.ref == "img:2/1"
        assert img.name == "Image 3"
        assert img.alt_text == "Third image description"

    def test_get_image_not_found(self) -> None:
        """Test get_image() returns None for missing ref."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INLINE_IMAGE)

        img = tree.get_image("img:99/99")

        assert img is None


class TestYamlSerializationWithImages:
    """Tests for YAML serialization of images."""

    def test_yaml_stats_include_images(self) -> None:
        """Test that YAML stats section includes images count."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INLINE_IMAGE)

        yaml = tree.to_yaml()

        assert "images: 1" in yaml

    def test_yaml_paragraph_has_images_section(self) -> None:
        """Test that paragraphs with images have images section."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INLINE_IMAGE)

        yaml = tree.to_yaml()

        assert "images:" in yaml
        assert "- ref: img:0/0" in yaml
        assert "type: image" in yaml
        assert "position: inline" in yaml

    def test_yaml_image_includes_name(self) -> None:
        """Test that image name is included in YAML."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INLINE_IMAGE)

        yaml = tree.to_yaml()

        assert "name: Picture 1" in yaml

    def test_yaml_image_includes_alt_text(self) -> None:
        """Test that image alt_text is included in YAML."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INLINE_IMAGE)

        yaml = tree.to_yaml()

        assert 'alt_text: "Logo image"' in yaml

    def test_yaml_image_includes_size(self) -> None:
        """Test that image size is included in YAML."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INLINE_IMAGE)

        yaml = tree.to_yaml()

        assert "size: 1.0in x 1.0in" in yaml

    def test_yaml_floating_image_info(self) -> None:
        """Test that floating images show position info."""
        tree = create_tree_from_xml(
            DOCUMENT_WITH_FLOATING_IMAGE, view_mode=ViewMode(verbosity="full")
        )

        yaml = tree.to_yaml()

        assert "position: floating" in yaml
        assert "floating_position:" in yaml
        assert "horizontal: left" in yaml
        assert "vertical: top" in yaml
        assert "wrap: square" in yaml


class TestImageType:
    """Tests for ImageType enum."""

    def test_image_types_exist(self) -> None:
        """Test that all image types exist."""
        assert ImageType.IMAGE
        assert ImageType.CHART
        assert ImageType.DIAGRAM
        assert ImageType.SHAPE
        assert ImageType.VML
        assert ImageType.OLE_OBJECT


class TestImagePositionType:
    """Tests for ImagePositionType enum."""

    def test_position_types_exist(self) -> None:
        """Test that position types exist."""
        assert ImagePositionType.INLINE
        assert ImagePositionType.FLOATING


class TestImageRefFormat:
    """Tests for image ref format."""

    def test_inline_image_ref_format(self) -> None:
        """Test inline image ref format: img:p_idx/img_idx."""
        tree = create_tree_from_xml(DOCUMENT_WITH_INLINE_IMAGE)

        images = tree.get_images()
        assert len(images) == 1
        # Format: img:paragraph_index/image_index
        assert images[0].ref == "img:0/0"

    def test_floating_image_ref_format(self) -> None:
        """Test floating image ref format: img:p_idx/f:img_idx."""
        tree = create_tree_from_xml(DOCUMENT_WITH_FLOATING_IMAGE)

        images = tree.get_images()
        assert len(images) == 1
        # Format: img:paragraph_index/f:image_index (f: for floating)
        assert images[0].ref == "img:0/f:0"

    def test_multiple_images_same_paragraph(self) -> None:
        """Test refs for multiple images in same paragraph."""
        tree = create_tree_from_xml(DOCUMENT_WITH_MULTIPLE_IMAGES)

        images = tree.get_images()
        assert len(images) == 3

        # First image in p:0
        assert images[0].ref == "img:0/0"
        # First image in p:2
        assert images[1].ref == "img:2/0"
        # Second image in p:2
        assert images[2].ref == "img:2/1"

    def test_vml_ref_format(self) -> None:
        """Test VML ref format."""
        tree = create_tree_from_xml(DOCUMENT_WITH_VML)

        images = tree.get_images()
        assert len(images) == 1
        assert images[0].ref == "vml:0/0"
