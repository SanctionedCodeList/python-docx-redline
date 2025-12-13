"""Tests for the validation module."""

import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from lxml import etree

from python_docx_redline.validation import (
    ValidationError,
    validate_document,
    validate_document_file,
)


class TestValidationError:
    """Tests for ValidationError class."""

    def test_init_with_message_only(self):
        """Test initialization with just a message."""
        err = ValidationError("Test error")
        assert str(err) == "Test error"
        assert err.errors == []

    def test_init_with_errors_list(self):
        """Test initialization with detailed error list."""
        errors = ["Error 1", "Error 2", "Error 3"]
        err = ValidationError("Test error", errors)
        assert err.errors == errors

    def test_init_with_none_errors(self):
        """Test that None errors becomes empty list."""
        err = ValidationError("Test error", None)
        assert err.errors == []

    def test_str_with_no_errors(self):
        """Test string representation with no error details."""
        err = ValidationError("Test error")
        assert str(err) == "Test error"

    def test_str_with_errors(self):
        """Test string representation includes all error details."""
        errors = ["Error 1", "Error 2"]
        err = ValidationError("Test error", errors)
        result = str(err)
        assert "Test error" in result
        assert "Error 1" in result
        assert "Error 2" in result
        assert "  - Error 1" in result  # Check formatting


class TestValidateDocument:
    """Tests for validate_document function."""

    def _create_minimal_document_xml(self) -> etree._Element:
        """Create a minimal valid document.xml root element."""
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        nsmap = {"w": word_ns}

        # Create minimal document structure
        root = etree.Element(f"{{{word_ns}}}document", nsmap=nsmap)
        body = etree.SubElement(root, f"{{{word_ns}}}body")
        p = etree.SubElement(body, f"{{{word_ns}}}p")
        r = etree.SubElement(p, f"{{{word_ns}}}r")
        t = etree.SubElement(r, f"{{{word_ns}}}t")
        t.text = "Hello world"

        return root

    def _create_docx(self, path: Path, content: str = "Hello world") -> None:
        """Create a minimal valid .docx file."""
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        # Create document.xml content
        doc_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{word_ns}">
    <w:body>
        <w:p>
            <w:r>
                <w:t>{content}</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""

        # Create [Content_Types].xml
        content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

        # Create .rels
        rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

        # Create word/_rels/document.xml.rels
        doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

        # Package as docx
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("word/document.xml", doc_xml)
            zf.writestr("[Content_Types].xml", content_types)
            zf.writestr("_rels/.rels", rels)
            zf.writestr("word/_rels/document.xml.rels", doc_rels)

    def test_validate_document_schema_passes(self, tmp_path):
        """Test validation when schema validation passes."""
        xml_root = self._create_minimal_document_xml()
        doc_path = tmp_path / "test.docx"
        self._create_docx(doc_path)

        # Mock DOCXSchemaValidator to return True (passes)
        with patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_validator_class:
            mock_validator = MagicMock()
            mock_validator.validate.return_value = True
            mock_validator_class.return_value = mock_validator

            # Should not raise
            validate_document(xml_root, doc_path, verbose=False)

    def test_validate_document_schema_fails(self, tmp_path):
        """Test validation raises when schema validation fails."""
        xml_root = self._create_minimal_document_xml()
        doc_path = tmp_path / "test.docx"
        self._create_docx(doc_path)

        # Mock DOCXSchemaValidator to return False (fails)
        with patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_validator_class:
            mock_validator = MagicMock()
            mock_validator.validate.return_value = False
            mock_validator_class.return_value = mock_validator

            with pytest.raises(ValidationError) as exc_info:
                validate_document(xml_root, doc_path, verbose=False)

            assert "OOXML schema validation failed" in str(exc_info.value)

    def test_validate_document_schema_exception(self, tmp_path):
        """Test validation handles schema validator exceptions."""
        xml_root = self._create_minimal_document_xml()
        doc_path = tmp_path / "test.docx"
        self._create_docx(doc_path)

        # Mock DOCXSchemaValidator to raise exception
        with patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_validator_class:
            mock_validator_class.side_effect = Exception("Schema error")

            with pytest.raises(ValidationError) as exc_info:
                validate_document(xml_root, doc_path, verbose=False)

            assert "Schema validation error" in str(exc_info.value)

    def test_validate_document_with_original(self, tmp_path):
        """Test validation with original document for redlining check."""
        xml_root = self._create_minimal_document_xml()
        doc_path = tmp_path / "modified.docx"
        original_path = tmp_path / "original.docx"
        self._create_docx(doc_path)
        self._create_docx(original_path)

        # Mock both validators to pass
        with (
            patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_schema_class,
            patch("python_docx_redline.validation.RedliningValidator") as mock_redline_class,
        ):
            mock_schema = MagicMock()
            mock_schema.validate.return_value = True
            mock_schema_class.return_value = mock_schema

            mock_redline = MagicMock()
            mock_redline.validate.return_value = True
            mock_redline_class.return_value = mock_redline

            # Should not raise
            validate_document(xml_root, doc_path, original_path, verbose=False)

            # Verify RedliningValidator was called
            mock_redline_class.assert_called_once()

    def test_validate_document_redlining_fails(self, tmp_path):
        """Test validation raises when redlining validation fails."""
        xml_root = self._create_minimal_document_xml()
        doc_path = tmp_path / "modified.docx"
        original_path = tmp_path / "original.docx"
        self._create_docx(doc_path)
        self._create_docx(original_path)

        # Mock schema to pass, redlining to fail
        with (
            patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_schema_class,
            patch("python_docx_redline.validation.RedliningValidator") as mock_redline_class,
        ):
            mock_schema = MagicMock()
            mock_schema.validate.return_value = True
            mock_schema_class.return_value = mock_schema

            mock_redline = MagicMock()
            mock_redline.validate.return_value = False
            mock_redline_class.return_value = mock_redline

            with pytest.raises(ValidationError) as exc_info:
                validate_document(xml_root, doc_path, original_path, verbose=False)

            assert "Redlining validation failed" in str(exc_info.value)

    def test_validate_document_redlining_exception(self, tmp_path):
        """Test validation handles redlining validator exceptions."""
        xml_root = self._create_minimal_document_xml()
        doc_path = tmp_path / "modified.docx"
        original_path = tmp_path / "original.docx"
        self._create_docx(doc_path)
        self._create_docx(original_path)

        # Mock schema to pass, redlining to raise exception
        with (
            patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_schema_class,
            patch("python_docx_redline.validation.RedliningValidator") as mock_redline_class,
        ):
            mock_schema = MagicMock()
            mock_schema.validate.return_value = True
            mock_schema_class.return_value = mock_schema

            mock_redline_class.side_effect = Exception("Redline error")

            with pytest.raises(ValidationError) as exc_info:
                validate_document(xml_root, doc_path, original_path, verbose=False)

            assert "Redlining validation error" in str(exc_info.value)

    def test_validate_document_same_path_skips_redlining(self, tmp_path):
        """Test that redlining is skipped when original_path == document_path."""
        xml_root = self._create_minimal_document_xml()
        doc_path = tmp_path / "test.docx"
        self._create_docx(doc_path)

        # Mock both validators
        with (
            patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_schema_class,
            patch("python_docx_redline.validation.RedliningValidator") as mock_redline_class,
        ):
            mock_schema = MagicMock()
            mock_schema.validate.return_value = True
            mock_schema_class.return_value = mock_schema

            # Should not raise, and redlining should not be called
            validate_document(xml_root, doc_path, doc_path, verbose=False)

            # RedliningValidator should NOT have been called
            mock_redline_class.assert_not_called()


class TestValidateDocumentFile:
    """Tests for validate_document_file function."""

    def _create_docx(self, path: Path, content: str = "Hello world") -> None:
        """Create a minimal valid .docx file."""
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        doc_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{word_ns}">
    <w:body>
        <w:p>
            <w:r>
                <w:t>{content}</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""

        content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

        rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

        doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("word/document.xml", doc_xml)
            zf.writestr("[Content_Types].xml", content_types)
            zf.writestr("_rels/.rels", rels)
            zf.writestr("word/_rels/document.xml.rels", doc_rels)

    def test_validate_document_file_success(self, tmp_path):
        """Test successful file validation."""
        doc_path = tmp_path / "test.docx"
        self._create_docx(doc_path)

        with patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_validator_class:
            mock_validator = MagicMock()
            mock_validator.validate.return_value = True
            mock_validator_class.return_value = mock_validator

            # Should not raise
            validate_document_file(doc_path, verbose=False)

    def test_validate_document_file_unpack_error(self, tmp_path):
        """Test validation fails on invalid zip file."""
        doc_path = tmp_path / "invalid.docx"
        doc_path.write_text("not a zip file")

        with pytest.raises(ValidationError) as exc_info:
            validate_document_file(doc_path, verbose=False)

        assert "Failed to unpack document" in str(exc_info.value)

    def test_validate_document_file_missing_file(self, tmp_path):
        """Test validation fails when file doesn't exist."""
        doc_path = tmp_path / "nonexistent.docx"

        with pytest.raises(ValidationError) as exc_info:
            validate_document_file(doc_path, verbose=False)

        assert "Failed to unpack document" in str(exc_info.value)

    def test_validate_document_file_schema_fails(self, tmp_path):
        """Test validation raises when schema validation fails."""
        doc_path = tmp_path / "test.docx"
        self._create_docx(doc_path)

        with patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_validator_class:
            mock_validator = MagicMock()
            mock_validator.validate.return_value = False
            mock_validator_class.return_value = mock_validator

            with pytest.raises(ValidationError) as exc_info:
                validate_document_file(doc_path, verbose=False)

            assert "OOXML schema validation failed" in str(exc_info.value)

    def test_validate_document_file_schema_exception(self, tmp_path):
        """Test validation handles schema validator exceptions."""
        doc_path = tmp_path / "test.docx"
        self._create_docx(doc_path)

        with patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_validator_class:
            mock_validator_class.side_effect = Exception("Validator crashed")

            with pytest.raises(ValidationError) as exc_info:
                validate_document_file(doc_path, verbose=False)

            assert "Schema validation error" in str(exc_info.value)

    def test_validate_document_file_with_original(self, tmp_path):
        """Test file validation with original for redlining check."""
        doc_path = tmp_path / "modified.docx"
        original_path = tmp_path / "original.docx"
        self._create_docx(doc_path)
        self._create_docx(original_path)

        with (
            patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_schema_class,
            patch("python_docx_redline.validation.RedliningValidator") as mock_redline_class,
        ):
            mock_schema = MagicMock()
            mock_schema.validate.return_value = True
            mock_schema_class.return_value = mock_schema

            mock_redline = MagicMock()
            mock_redline.validate.return_value = True
            mock_redline_class.return_value = mock_redline

            # Should not raise
            validate_document_file(doc_path, original_path, verbose=False)

            # Verify RedliningValidator was called
            mock_redline_class.assert_called_once()

    def test_validate_document_file_redlining_fails(self, tmp_path):
        """Test file validation raises when redlining fails."""
        doc_path = tmp_path / "modified.docx"
        original_path = tmp_path / "original.docx"
        self._create_docx(doc_path)
        self._create_docx(original_path)

        with (
            patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_schema_class,
            patch("python_docx_redline.validation.RedliningValidator") as mock_redline_class,
        ):
            mock_schema = MagicMock()
            mock_schema.validate.return_value = True
            mock_schema_class.return_value = mock_schema

            mock_redline = MagicMock()
            mock_redline.validate.return_value = False
            mock_redline_class.return_value = mock_redline

            with pytest.raises(ValidationError) as exc_info:
                validate_document_file(doc_path, original_path, verbose=False)

            assert "Redlining validation failed" in str(exc_info.value)

    def test_validate_document_file_redlining_exception(self, tmp_path):
        """Test file validation handles redlining exceptions."""
        doc_path = tmp_path / "modified.docx"
        original_path = tmp_path / "original.docx"
        self._create_docx(doc_path)
        self._create_docx(original_path)

        with (
            patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_schema_class,
            patch("python_docx_redline.validation.RedliningValidator") as mock_redline_class,
        ):
            mock_schema = MagicMock()
            mock_schema.validate.return_value = True
            mock_schema_class.return_value = mock_schema

            mock_redline_class.side_effect = Exception("Redline crashed")

            with pytest.raises(ValidationError) as exc_info:
                validate_document_file(doc_path, original_path, verbose=False)

            assert "Redlining validation error" in str(exc_info.value)

    def test_validate_document_file_verbose_mode(self, tmp_path):
        """Test that verbose flag is passed through."""
        doc_path = tmp_path / "test.docx"
        self._create_docx(doc_path)

        with patch("python_docx_redline.validation.DOCXSchemaValidator") as mock_validator_class:
            mock_validator = MagicMock()
            mock_validator.validate.return_value = True
            mock_validator_class.return_value = mock_validator

            validate_document_file(doc_path, verbose=True)

            # Check verbose was passed
            call_kwargs = mock_validator_class.call_args[1]
            assert call_kwargs["verbose"] is True
