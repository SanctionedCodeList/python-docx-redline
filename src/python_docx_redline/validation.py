"""
Full OOXML validation for python_docx_redline documents.

This module provides comprehensive validation using the same validation
suite as the Anthropic docx skill to ensure documents are production-ready
and will not cause issues when opened in Microsoft Word or other consumers.

Validates:
1. XML well-formedness
2. Namespace declarations
3. Unique IDs
4. File references / relationships
5. Content type declarations
6. XSD schema validation
7. Whitespace preservation
8. Deletion structure (w:delText in w:del)
9. Insertion structure (w:t in w:ins)
10. Relationship ID references
11. Content integrity (tracked changes only)
"""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from .errors import ValidationError as BaseValidationError
from .validation_docx import DOCXSchemaValidator
from .validation_redlining import RedliningValidator


class ValidationError(BaseValidationError):
    """Raised when document validation fails with detailed error list."""

    def __init__(self, message: str, errors: list[str] | None = None):
        """Initialize validation error with detailed messages.

        Args:
            message: Primary error message
            errors: List of specific validation errors for bug reporting
        """
        super().__init__(message)
        self.errors = errors or []

    def __str__(self) -> str:
        """Format error with all validation details."""
        if not self.errors:
            return super().__str__()

        error_details = "\n  - " + "\n  - ".join(self.errors)
        return f"{super().__str__()}{error_details}"


def validate_document(
    xml_root: etree._Element,
    document_path: Path,
    original_path: Path | None = None,
    verbose: bool = False,
) -> None:
    """Validate a document using the full OOXML validation suite.

    This runs the same comprehensive validation as the Anthropic docx skill
    to ensure documents are production-ready and spec-compliant.

    Args:
        xml_root: The root element of word/document.xml
        document_path: Path to the current document being validated
        original_path: Optional path to original document (for redlining validation)
        verbose: Enable verbose validation output

    Raises:
        ValidationError: If any validation check fails, with detailed error list
    """
    errors = []

    # Create a temporary docx file for validation
    with tempfile.TemporaryDirectory() as temp_dir:
        _ = Path(temp_dir) / "validation.docx"  # Not used but kept for reference

        # Save the current document to temp location
        # We need to create a proper .docx file structure
        with tempfile.TemporaryDirectory() as unpack_dir:
            unpack_path = Path(unpack_dir)

            # Create the directory structure
            word_dir = unpack_path / "word"
            word_dir.mkdir(parents=True)
            rels_dir = unpack_path / "_rels"
            rels_dir.mkdir(parents=True)

            # Write document.xml
            doc_xml_path = word_dir / "document.xml"
            with open(doc_xml_path, "wb") as f:
                f.write(etree.tostring(xml_root, xml_declaration=True, encoding="UTF-8"))

            # Note: This creates a minimal structure with only document.xml.
            # Full document structure (styles.xml, comments.xml, etc.) support
            # is tracked in beads issue docx_redline-290.

            # Run DOCXSchemaValidator
            try:
                validator = DOCXSchemaValidator(
                    unpacked_dir=unpack_path,
                    original_file=original_path if original_path else document_path,
                    verbose=verbose,
                )

                # Run validation - returns False if any check fails
                # The validator prints errors directly, we need to capture them
                if not validator.validate():
                    errors.append("OOXML schema validation failed. See detailed output above.")
            except Exception as e:
                errors.append(f"Schema validation error: {e}")

            # Run RedliningValidator if we have an original
            if original_path and original_path != document_path:
                try:
                    redline_validator = RedliningValidator(
                        unpacked_dir=unpack_path,
                        original_docx=original_path,
                        verbose=verbose,
                    )
                    if not redline_validator.validate():
                        errors.append(
                            "Redlining validation failed. Content changed outside tracked changes."
                        )
                except Exception as e:
                    errors.append(f"Redlining validation error: {e}")

    if errors:
        raise ValidationError(
            "Document validation failed. Please report this as a bug with the error details below:",
            errors,
        )


def validate_document_file(
    docx_path: Path, original_path: Path | None = None, verbose: bool = False
) -> None:
    """Validate a complete .docx file using full OOXML validation suite.

    Args:
        docx_path: Path to the .docx file to validate
        original_path: Optional path to original document (for redlining validation)
        verbose: Enable verbose validation output

    Raises:
        ValidationError: If any validation check fails
    """
    errors = []

    with tempfile.TemporaryDirectory() as unpack_dir:
        unpack_path = Path(unpack_dir)

        # Unpack the docx
        try:
            with zipfile.ZipFile(docx_path, "r") as zip_ref:
                zip_ref.extractall(unpack_path)
        except Exception as e:
            raise ValidationError(f"Failed to unpack document: {e}")

        # Run DOCXSchemaValidator
        try:
            validator = DOCXSchemaValidator(
                unpacked_dir=unpack_path,
                original_file=original_path if original_path else docx_path,
                verbose=verbose,
            )

            if not validator.validate():
                errors.append("OOXML schema validation failed")
        except Exception as e:
            errors.append(f"Schema validation error: {e}")

        # Run RedliningValidator if we have an original
        if original_path:
            try:
                redline_validator = RedliningValidator(
                    unpacked_dir=unpack_path,
                    original_docx=original_path,
                    verbose=verbose,
                )
                if not redline_validator.validate():
                    errors.append("Redlining validation failed")
            except Exception as e:
                errors.append(f"Redlining validation error: {e}")

    if errors:
        raise ValidationError(
            "Document validation failed. Please report this as a bug with the error details below:",
            errors,
        )
