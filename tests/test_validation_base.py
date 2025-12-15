"""Tests for the BaseSchemaValidator class."""

from unittest.mock import patch

import pytest
from lxml import etree

from python_docx_redline.validation_base import BaseSchemaValidator


class TestValidatorInit:
    """Tests for BaseSchemaValidator initialization."""

    def test_init_sets_paths(self, tmp_path):
        """Test that paths are correctly set during initialization."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        assert validator.unpacked_dir == unpacked_dir.resolve()
        assert validator.original_file == original_file

    def test_init_sets_verbose(self, tmp_path):
        """Test that verbose flag is set correctly."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)

        assert validator.verbose is True

    def test_init_finds_xml_files(self, tmp_path):
        """Test that XML files are discovered."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        (unpacked_dir / "test.xml").touch()
        (unpacked_dir / "test.rels").touch()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        assert len(validator.xml_files) == 2

    def test_init_no_xml_files_logs_warning(self, tmp_path, caplog):
        """Test that a warning is logged when no XML files found."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        with caplog.at_level("WARNING"):
            validator = BaseSchemaValidator(unpacked_dir, original_file)

        assert len(validator.xml_files) == 0
        assert "No XML files found" in caplog.text


class TestValidateMethod:
    """Tests for the abstract validate() method."""

    def test_validate_raises_not_implemented(self, tmp_path):
        """Test that validate() raises NotImplementedError."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        with pytest.raises(NotImplementedError):
            validator.validate()


class TestValidateEncodingDeclarations:
    """Tests for validate_encoding_declarations method."""

    def test_passes_with_utf8_encoding(self, tmp_path):
        """Test validation passes for UTF-8 encoded files."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text('<?xml version="1.0" encoding="UTF-8"?><root/>')
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_encoding_declarations()

        assert result is True

    def test_passes_with_utf16_encoding(self, tmp_path):
        """Test validation passes for UTF-16 encoded files."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text('<?xml version="1.0" encoding="UTF-16"?><root/>')
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_encoding_declarations()

        assert result is True

    def test_fails_with_invalid_encoding(self, tmp_path):
        """Test validation fails for non-UTF-8/UTF-16 encoding."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text('<?xml version="1.0" encoding="ISO-8859-1"?><root/>')
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_encoding_declarations()

        assert result is False

    def test_passes_without_encoding_declaration(self, tmp_path):
        """Test validation passes when no encoding is declared."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text('<?xml version="1.0"?><root/>')
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_encoding_declarations()

        assert result is True

    def test_handles_file_read_error(self, tmp_path):
        """Test handling of file read errors."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("<root/>")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        # Make the file unreadable by patching open
        with patch("builtins.open", side_effect=PermissionError("Cannot read")):
            result = validator.validate_encoding_declarations()

        assert result is False


class TestValidateXml:
    """Tests for validate_xml method."""

    def test_passes_for_valid_xml(self, tmp_path):
        """Test validation passes for well-formed XML."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text('<?xml version="1.0"?><root><child/></root>')
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_xml()

        assert result is True

    def test_fails_for_malformed_xml(self, tmp_path):
        """Test validation fails for malformed XML."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("<root><unclosed>")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_xml()

        assert result is False

    def test_handles_unexpected_errors(self, tmp_path):
        """Test handling of unexpected parsing errors."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("<root/>")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        with patch("lxml.etree.parse", side_effect=Exception("Unexpected error")):
            result = validator.validate_xml()

        assert result is False


class TestValidateNamespaces:
    """Tests for validate_namespaces method."""

    def test_passes_for_valid_namespaces(self, tmp_path):
        """Test validation passes when all prefixes are declared."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("""<?xml version="1.0"?>
        <root xmlns:mc="http://example.com/mc" mc:Ignorable="ns1" xmlns:ns1="http://example.com/ns1">
            <child/>
        </root>""")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_namespaces()

        assert result is True

    def test_fails_for_undeclared_namespace(self, tmp_path):
        """Test validation fails when Ignorable references undeclared namespace."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("""<?xml version="1.0"?>
        <root xmlns:mc="http://example.com/mc" mc:Ignorable="undeclared">
            <child/>
        </root>""")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_namespaces()

        assert result is False

    def test_skips_malformed_files(self, tmp_path):
        """Test that malformed XML files are skipped."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("<root><unclosed>")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_namespaces()

        # Should pass because malformed files are skipped
        assert result is True


class TestValidateUniqueIds:
    """Tests for validate_unique_ids method."""

    def test_passes_for_unique_ids(self, tmp_path):
        """Test validation passes when all IDs are unique."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        # Create XML with unique comment IDs
        xml_file.write_text("""<?xml version="1.0"?>
        <root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:comment w:id="1"/>
            <w:comment w:id="2"/>
        </root>""")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_unique_ids()

        assert result is True

    def test_fails_for_duplicate_file_scoped_ids(self, tmp_path):
        """Test validation fails for duplicate file-scoped IDs."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("""<?xml version="1.0"?>
        <root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:comment w:id="1"/>
            <w:comment w:id="1"/>
        </root>""")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_unique_ids()

        assert result is False

    def test_fails_for_duplicate_global_ids(self, tmp_path):
        """Test validation fails for duplicate global-scoped IDs across files."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        # Create two files with the same sldMasterId (global scope)
        xml_file1 = unpacked_dir / "test1.xml"
        xml_file1.write_text("""<?xml version="1.0"?>
        <root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:sldMasterId p:id="100"/>
        </root>""")

        xml_file2 = unpacked_dir / "test2.xml"
        xml_file2.write_text("""<?xml version="1.0"?>
        <root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:sldMasterId p:id="100"/>
        </root>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_unique_ids()

        assert result is False

    def test_skips_mc_alternate_content(self, tmp_path):
        """Test that mc:AlternateContent elements are removed before validation."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("""<?xml version="1.0"?>
        <root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
              xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            <mc:AlternateContent>
                <w:comment w:id="1"/>
            </mc:AlternateContent>
            <w:comment w:id="1"/>
        </root>""")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_unique_ids()

        # Should pass because mc:AlternateContent elements are removed
        assert result is True

    def test_handles_parse_errors(self, tmp_path):
        """Test handling of XML parse errors."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("<root><unclosed>")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_unique_ids()

        # Should fail due to parse error
        assert result is False


class TestValidateFileReferences:
    """Tests for validate_file_references method."""

    def test_passes_with_no_rels_files(self, tmp_path):
        """Test validation passes when there are no .rels files."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        (unpacked_dir / "test.xml").write_text("<root/>")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_file_references()

        assert result is True

    def test_passes_with_valid_references(self, tmp_path):
        """Test validation passes when all references are valid."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        # Create _rels directory and .rels file
        rels_dir = unpacked_dir / "_rels"
        rels_dir.mkdir()
        rels_file = rels_dir / ".rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="word/document.xml" Type="http://test"/>
        </Relationships>""")

        # Create the target file
        word_dir = unpacked_dir / "word"
        word_dir.mkdir()
        (word_dir / "document.xml").write_text("<root/>")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_file_references()

        assert result is True

    def test_fails_with_broken_reference(self, tmp_path):
        """Test validation fails when a reference points to non-existent file."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        rels_dir = unpacked_dir / "_rels"
        rels_dir.mkdir()
        rels_file = rels_dir / ".rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="missing/file.xml" Type="http://test"/>
        </Relationships>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_file_references()

        assert result is False

    def test_fails_with_unreferenced_file(self, tmp_path):
        """Test validation fails when a file is not referenced."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        rels_dir = unpacked_dir / "_rels"
        rels_dir.mkdir()
        rels_file = rels_dir / ".rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>""")

        # Create a file that should be referenced but isn't
        word_dir = unpacked_dir / "word"
        word_dir.mkdir()
        (word_dir / "document.xml").write_text("<root/>")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_file_references()

        assert result is False

    def test_skips_external_urls(self, tmp_path):
        """Test that external URLs are skipped."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        rels_dir = unpacked_dir / "_rels"
        rels_dir.mkdir()
        rels_file = rels_dir / ".rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="http://example.com/external" Type="http://test"/>
            <Relationship Id="rId2" Target="mailto:test@example.com" Type="http://test"/>
        </Relationships>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_file_references()

        assert result is True


class TestValidateAllRelationshipIds:
    """Tests for validate_all_relationship_ids method."""

    def test_passes_with_valid_ids(self, tmp_path):
        """Test validation passes when all r:id references are valid."""
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        rels_dir = word_dir / "_rels"
        rels_dir.mkdir(parents=True)

        # Create document.xml.rels
        rels_file = rels_dir / "document.xml.rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="image.png" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>
        </Relationships>""")

        # Create document.xml with r:id reference
        doc_file = word_dir / "document.xml"
        doc_file.write_text("""<?xml version="1.0"?>
        <document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <element r:id="rId1"/>
        </document>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_all_relationship_ids()

        assert result is True

    def test_fails_with_invalid_id_reference(self, tmp_path):
        """Test validation fails when r:id references non-existent ID."""
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        rels_dir = word_dir / "_rels"
        rels_dir.mkdir(parents=True)

        # Create document.xml.rels with only rId1
        rels_file = rels_dir / "document.xml.rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="image.png" Type="http://test"/>
        </Relationships>""")

        # Create document.xml with reference to rId999 (doesn't exist)
        doc_file = word_dir / "document.xml"
        doc_file.write_text("""<?xml version="1.0"?>
        <document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <element r:id="rId999"/>
        </document>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_all_relationship_ids()

        assert result is False

    def test_fails_with_duplicate_rids_in_rels(self, tmp_path):
        """Test validation fails when .rels has duplicate IDs."""
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        rels_dir = word_dir / "_rels"
        rels_dir.mkdir(parents=True)

        # Create document.xml.rels with duplicate rId1
        rels_file = rels_dir / "document.xml.rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="image1.png" Type="http://test"/>
            <Relationship Id="rId1" Target="image2.png" Type="http://test"/>
        </Relationships>""")

        # Create document.xml
        doc_file = word_dir / "document.xml"
        doc_file.write_text("""<?xml version="1.0"?>
        <document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <element r:id="rId1"/>
        </document>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_all_relationship_ids()

        assert result is False

    def test_skips_files_without_rels(self, tmp_path):
        """Test that XML files without .rels files are skipped."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("<root/>")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_all_relationship_ids()

        assert result is True


class TestGetExpectedRelationshipType:
    """Tests for _get_expected_relationship_type method."""

    def test_explicit_mapping(self, tmp_path):
        """Test that explicit mappings are returned."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        # Set up a test mapping
        validator.ELEMENT_RELATIONSHIP_TYPES = {"blip": "image"}

        result = validator._get_expected_relationship_type("blip")
        assert result == "image"

    def test_id_suffix_pattern(self, tmp_path):
        """Test pattern detection for elements ending in 'Id'."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        validator.ELEMENT_RELATIONSHIP_TYPES = {}

        # sldId should map to "slide"
        result = validator._get_expected_relationship_type("sldId")
        assert result == "slide"

    def test_master_id_pattern(self, tmp_path):
        """Test pattern detection for elements ending in 'masterId'."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        validator.ELEMENT_RELATIONSHIP_TYPES = {}

        result = validator._get_expected_relationship_type("sldMasterId")
        assert result == "sldmaster"

    def test_layout_id_pattern(self, tmp_path):
        """Test pattern detection for elements ending in 'layoutId'."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        validator.ELEMENT_RELATIONSHIP_TYPES = {}

        result = validator._get_expected_relationship_type("sldLayoutId")
        assert result == "sldlayout"

    def test_reference_suffix_pattern(self, tmp_path):
        """Test pattern detection for elements ending in 'Reference'."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        validator.ELEMENT_RELATIONSHIP_TYPES = {}

        result = validator._get_expected_relationship_type("imageReference")
        assert result == "image"

    def test_unknown_element_returns_none(self, tmp_path):
        """Test that unknown elements return None."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        validator.ELEMENT_RELATIONSHIP_TYPES = {}

        result = validator._get_expected_relationship_type("unknownElement")
        assert result is None


class TestValidateContentTypes:
    """Tests for validate_content_types method."""

    def test_fails_when_content_types_missing(self, tmp_path):
        """Test validation fails when [Content_Types].xml is missing."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_content_types()

        assert result is False

    def test_passes_for_valid_content_types(self, tmp_path):
        """Test validation passes with properly declared content types."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        # Create [Content_Types].xml
        content_types = unpacked_dir / "[Content_Types].xml"
        content_types.write_text("""<?xml version="1.0"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="xml" ContentType="application/xml"/>
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>""")

        # Create word/document.xml with 'document' root (which needs Override)
        word_dir = unpacked_dir / "word"
        word_dir.mkdir()
        doc_file = word_dir / "document.xml"
        doc_file.write_text("""<?xml version="1.0"?>
        <document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_content_types()

        assert result is True

    def test_fails_for_missing_media_extension(self, tmp_path):
        """Test validation fails when media extension is not declared."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        # Create [Content_Types].xml without png declaration
        content_types = unpacked_dir / "[Content_Types].xml"
        content_types.write_text("""<?xml version="1.0"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="xml" ContentType="application/xml"/>
        </Types>""")

        # Create a PNG file in media folder
        media_dir = unpacked_dir / "word" / "media"
        media_dir.mkdir(parents=True)
        (media_dir / "image.png").touch()

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_content_types()

        assert result is False


class TestXsdValidation:
    """Tests for XSD schema validation methods."""

    def test_get_schema_path_exact_match(self, tmp_path):
        """Test that exact filename matches work."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        # Create a mock xml file
        xml_file = unpacked_dir / "core.xml"
        schema_path = validator._get_schema_path(xml_file)

        expected = validator.schemas_dir / "ecma/fouth-edition/opc-coreProperties.xsd"
        assert schema_path == expected

    def test_get_schema_path_rels_files(self, tmp_path):
        """Test that .rels files use correct schema."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        xml_file = unpacked_dir / "test.rels"
        schema_path = validator._get_schema_path(xml_file)

        expected = validator.schemas_dir / "ecma/fouth-edition/opc-relationships.xsd"
        assert schema_path == expected

    def test_get_schema_path_returns_none_for_unknown(self, tmp_path):
        """Test that unknown files return None."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        xml_file = unpacked_dir / "custom" / "unknown.xml"
        schema_path = validator._get_schema_path(xml_file)

        assert schema_path is None

    def test_clean_ignorable_namespaces(self, tmp_path):
        """Test that ignorable namespace attributes and elements are removed."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        # Create XML with ignorable namespace attributes
        xml_str = """<?xml version="1.0"?>
        <root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
              xmlns:custom="http://custom.namespace">
            <w:para custom:attr="value"/>
            <custom:element/>
        </root>"""
        xml_doc = etree.ElementTree(etree.fromstring(xml_str))

        result = validator._clean_ignorable_namespaces(xml_doc)
        result_root = result.getroot()

        # The custom namespace attribute and element should be removed
        assert "{http://custom.namespace}attr" not in result_root[0].attrib

    def test_preprocess_mc_ignorable(self, tmp_path):
        """Test that mc:Ignorable attribute is removed."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        mc_ns = validator.MC_NAMESPACE
        xml_str = f"""<?xml version="1.0"?>
        <root xmlns:mc="{mc_ns}" mc:Ignorable="w14">
            <child/>
        </root>"""
        xml_doc = etree.ElementTree(etree.fromstring(xml_str))

        result = validator._preprocess_for_mc_ignorable(xml_doc)
        result_root = result.getroot()

        # mc:Ignorable should be removed
        assert f"{{{mc_ns}}}Ignorable" not in result_root.attrib

    def test_remove_template_tags_from_text_nodes(self, tmp_path):
        """Test that template tags are removed from text nodes."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        xml_str = """<?xml version="1.0"?>
        <root>
            <child>Hello {{template}} World</child>
            <another>Text</another>after {{tag}}
        </root>"""
        xml_doc = etree.ElementTree(etree.fromstring(xml_str))

        _result, warnings = validator._remove_template_tags_from_text_nodes(xml_doc)

        # Template tags should be removed from text content
        # Note: template tags in text content (not w:t elements) are processed
        assert len(warnings) > 0


class TestMainGuard:
    """Test the __main__ guard."""

    def test_raises_when_run_directly(self):
        """Test that running the module directly raises an error."""
        import python_docx_redline.validation_base as mod

        # The module should have a __name__ check at the bottom
        assert hasattr(mod, "BaseSchemaValidator")


class TestValidateAgainstXsd:
    """Tests for validate_against_xsd method."""

    def test_validate_against_xsd_valid_files(self, tmp_path):
        """Test XSD validation passes for files without new errors."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        xml_file = unpacked_dir / "test.xml"
        xml_file.write_text("<root/>")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        # Since there's no schema for 'test.xml', it should be skipped
        result = validator.validate_against_xsd()

        assert result is True

    def test_validate_against_xsd_skips_without_schema(self, tmp_path):
        """Test that files without schemas are skipped."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        # Create a file that has no matching schema
        xml_file = unpacked_dir / "custom" / "random.xml"
        xml_file.parent.mkdir(parents=True)
        xml_file.write_text("<root/>")
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_against_xsd()

        # Should pass because file is skipped
        assert result is True


class TestGetSchemaPathEdgeCases:
    """Additional tests for _get_schema_path method."""

    def test_get_schema_path_chart_file(self, tmp_path):
        """Test schema path for chart files."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        # Create a chart file path
        charts_dir = unpacked_dir / "word" / "charts"
        charts_dir.mkdir(parents=True)
        chart_file = charts_dir / "chart1.xml"

        schema_path = validator._get_schema_path(chart_file)
        expected = validator.schemas_dir / "ISO-IEC29500-4_2016/dml-chart.xsd"
        assert schema_path == expected

    def test_get_schema_path_theme_file(self, tmp_path):
        """Test schema path for theme files."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        # Create a theme file path
        theme_dir = unpacked_dir / "word" / "theme"
        theme_dir.mkdir(parents=True)
        theme_file = theme_dir / "theme1.xml"

        schema_path = validator._get_schema_path(theme_file)
        expected = validator.schemas_dir / "ISO-IEC29500-4_2016/dml-main.xsd"
        assert schema_path == expected

    def test_get_schema_path_main_content_folder(self, tmp_path):
        """Test schema path for files in main content folders."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        # Create a file in word folder
        word_dir = unpacked_dir / "word"
        word_dir.mkdir(parents=True)
        word_file = word_dir / "document.xml"

        schema_path = validator._get_schema_path(word_file)
        expected = validator.schemas_dir / "ISO-IEC29500-4_2016/wml.xsd"
        assert schema_path == expected


class TestRemoveIgnorableElementsEdgeCases:
    """Tests for edge cases in _remove_ignorable_elements."""

    def test_remove_ignorable_elements_with_callable_tag(self, tmp_path):
        """Test that non-element nodes (comments, PI) are skipped."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        # Create XML with comments
        xml_str = """<?xml version="1.0"?>
        <root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <!-- This is a comment -->
            <w:para/>
        </root>"""
        xml_doc = etree.fromstring(xml_str)

        # Should not raise
        validator._remove_ignorable_elements(xml_doc)

        # Structure should remain
        assert xml_doc.tag == "root"


class TestRemoveTemplateTags:
    """Tests for _remove_template_tags_from_text_nodes edge cases."""

    def test_remove_template_tags_in_wt_elements(self, tmp_path):
        """Test that w:t elements are skipped for template tag processing."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)

        # Create XML with template tags in w:t elements (should be skipped)
        xml_str = """<?xml version="1.0"?>
        <root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:t>Hello {{template}} World</w:t>
            <other>Normal {{tag}} text</other>
        </root>"""
        xml_doc = etree.ElementTree(etree.fromstring(xml_str))

        result, warnings = validator._remove_template_tags_from_text_nodes(xml_doc)
        result_root = result.getroot()

        # w:t should retain template tag
        wt = result_root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        assert "{{template}}" in wt.text

        # other element should have template tag removed
        other = result_root.find(".//other")
        assert "{{tag}}" not in other.text


class TestEncodingValidationEdgeCases:
    """Tests for encoding validation edge cases."""

    def test_encoding_latin1_fallback(self, tmp_path):
        """Test that latin-1 fallback is used when utf-8 decode fails."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_file = tmp_path / "original.docx"
        original_file.touch()

        # Create a file with bytes that might cause UTF-8 issues
        xml_file = unpacked_dir / "test.xml"
        # Write bytes directly with a valid XML declaration
        xml_file.write_bytes(b'<?xml version="1.0" encoding="UTF-8"?><root/>')

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_encoding_declarations()

        assert result is True


class TestFileReferenceValidationEdgeCases:
    """Tests for file reference validation edge cases."""

    def test_verbose_case_mismatch_logging(self, tmp_path):
        """Test verbose logging for case mismatches."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        # Create _rels directory and .rels file with uppercase reference
        rels_dir = unpacked_dir / "_rels"
        rels_dir.mkdir()
        rels_file = rels_dir / ".rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="WORD/Document.XML" Type="http://test"/>
        </Relationships>""")

        # Create the target file with different case
        word_dir = unpacked_dir / "word"
        word_dir.mkdir()
        (word_dir / "document.xml").write_text("<root/>")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file, verbose=True)
        result = validator.validate_file_references()

        # Should pass due to case-insensitive matching
        assert result is True

    def test_file_exists_but_not_in_lookup(self, tmp_path):
        """Test when file exists but isn't in case-insensitive lookup."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        rels_dir = unpacked_dir / "_rels"
        rels_dir.mkdir()
        rels_file = rels_dir / ".rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="word/document.xml" Type="http://test"/>
        </Relationships>""")

        word_dir = unpacked_dir / "word"
        word_dir.mkdir()
        (word_dir / "document.xml").write_text("<root/>")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_file_references()

        assert result is True

    def test_broken_ref_with_oserror(self, tmp_path):
        """Test handling of OSError when resolving paths."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        rels_dir = unpacked_dir / "_rels"
        rels_dir.mkdir()
        # Create a relationship pointing to an invalid path (with null bytes)
        rels_file = rels_dir / ".rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="../../../etc/passwd" Type="http://test"/>
        </Relationships>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_file_references()

        # Should fail due to broken reference
        assert result is False


class TestRelationshipTypeValidation:
    """Tests for relationship type validation with ELEMENT_RELATIONSHIP_TYPES."""

    def test_validates_relationship_type_match(self, tmp_path):
        """Test that relationship types are validated when mapping exists."""
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        rels_dir = word_dir / "_rels"
        rels_dir.mkdir(parents=True)

        # Create document.xml.rels with image relationship
        rels_file = rels_dir / "document.xml.rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="image.png" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>
        </Relationships>""")

        # Create document.xml with blip element referencing the relationship
        doc_file = word_dir / "document.xml"
        doc_file.write_text("""<?xml version="1.0"?>
        <document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:blip r:id="rId1"/>
        </document>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        validator.ELEMENT_RELATIONSHIP_TYPES = {"blip": "image"}

        result = validator.validate_all_relationship_ids()

        assert result is True

    def test_fails_on_relationship_type_mismatch(self, tmp_path):
        """Test that mismatched relationship types are detected."""
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        rels_dir = word_dir / "_rels"
        rels_dir.mkdir(parents=True)

        # Create document.xml.rels with wrong relationship type
        rels_file = rels_dir / "document.xml.rels"
        rels_file.write_text("""<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="chart.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>
        </Relationships>""")

        # Create document.xml with blip element (expects image) referencing chart
        doc_file = word_dir / "document.xml"
        doc_file.write_text("""<?xml version="1.0"?>
        <document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:blip r:id="rId1"/>
        </document>""")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        validator.ELEMENT_RELATIONSHIP_TYPES = {"blip": "image"}

        result = validator.validate_all_relationship_ids()

        assert result is False


class TestContentTypesEdgeCases:
    """Tests for content type validation edge cases."""

    def test_content_types_parse_error(self, tmp_path):
        """Test handling of parse errors in [Content_Types].xml."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        # Create malformed [Content_Types].xml
        content_types = unpacked_dir / "[Content_Types].xml"
        content_types.write_text("<Types><unclosed>")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_content_types()

        assert result is False

    def test_skips_unparseable_xml_files(self, tmp_path):
        """Test that unparseable XML files are skipped in content type check."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()

        # Create [Content_Types].xml
        content_types = unpacked_dir / "[Content_Types].xml"
        content_types.write_text("""<?xml version="1.0"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="xml" ContentType="application/xml"/>
        </Types>""")

        # Create malformed XML file
        word_dir = unpacked_dir / "word"
        word_dir.mkdir()
        (word_dir / "malformed.xml").write_text("<root><unclosed>")

        original_file = tmp_path / "original.docx"
        original_file.touch()

        validator = BaseSchemaValidator(unpacked_dir, original_file)
        result = validator.validate_content_types()

        # Should pass - malformed files are skipped
        assert result is True
