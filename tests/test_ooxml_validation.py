"""
OOXML validation tests to ensure python_docx_redline produces valid Office documents.

These tests use validation techniques from the docx skill in ~/.agents to verify:
1. XML well-formedness
2. Tracked changes structure (w:ins, w:del)
3. Whitespace preservation (xml:space='preserve')
4. Content integrity (no untracked modifications)
"""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from python_docx_redline import AuthorIdentity, Document

# Namespace constants
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"


def create_test_document() -> Path:
    """Create a simple but valid test document with proper OOXML structure."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Proper Content_Types.xml
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    # Proper relationships file
    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>This is a test document.</w:t></w:r></w:p>
<w:p><w:r><w:t>It has multiple paragraphs.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document_xml)

    return doc_path


def get_document_xml(doc: Document) -> etree._Element:
    """Extract and parse document.xml from a Document object."""
    return doc.xml_root


def validate_xml_wellformed(doc: Document) -> tuple[bool, str]:
    """Validate that document XML is well-formed."""
    try:
        get_document_xml(doc)
        # If we can parse it, it's well-formed
        return True, "XML is well-formed"
    except etree.XMLSyntaxError as e:
        return False, f"XML syntax error: {e}"


def validate_whitespace_preservation(doc: Document) -> tuple[bool, list[str]]:
    """Validate whitespace preservation (same logic as DOCXSchemaValidator)."""
    import re

    errors = []
    root = doc.xml_root

    # Find all w:t elements
    for elem in root.iter(f"{{{WORD_NAMESPACE}}}t"):
        if elem.text:
            text = elem.text
            # Check if text starts or ends with whitespace
            if re.match(r"^\s.*", text) or re.match(r".*\s$", text):
                # Check if xml:space="preserve" attribute exists
                xml_space_attr = f"{{{XML_NAMESPACE}}}space"
                if xml_space_attr not in elem.attrib or elem.attrib[xml_space_attr] != "preserve":
                    text_preview = repr(text)[:50] + "..." if len(repr(text)) > 50 else repr(text)
                    errors.append(
                        f"Line {elem.sourceline}: <w:t> element with whitespace "
                        f"missing xml:space='preserve': {text_preview}"
                    )

    return (len(errors) == 0, errors)


def validate_deletion_structure(doc: Document) -> tuple[bool, list[str]]:
    """Validate deletion structure (same logic as DOCXSchemaValidator)."""
    errors = []
    root = doc.xml_root

    # Find all w:t elements that are descendants of w:del elements
    namespaces = {"w": WORD_NAMESPACE}
    xpath_expression = ".//w:del//w:t"
    problematic_t_elements = root.xpath(xpath_expression, namespaces=namespaces)

    for t_elem in problematic_t_elements:
        if t_elem.text:
            text_preview = (
                repr(t_elem.text)[:50] + "..." if len(repr(t_elem.text)) > 50 else repr(t_elem.text)
            )
            errors.append(
                f"Line {t_elem.sourceline}: <w:t> found within <w:del> "
                f"(should be <w:delText>): {text_preview}"
            )

    return (len(errors) == 0, errors)


def validate_insertion_structure(doc: Document) -> tuple[bool, list[str]]:
    """Validate insertion structure (same logic as DOCXSchemaValidator)."""
    errors = []
    root = doc.xml_root
    namespaces = {"w": WORD_NAMESPACE}

    # Find w:delText in w:ins that are NOT within w:del
    invalid_elements = root.xpath(
        ".//w:ins//w:delText[not(ancestor::w:del)]", namespaces=namespaces
    )

    for elem in invalid_elements:
        text_preview = (
            repr(elem.text or "")[:50] + "..."
            if len(repr(elem.text or "")) > 50
            else repr(elem.text or "")
        )
        errors.append(
            f"Line {elem.sourceline}: <w:delText> found within <w:ins> "
            f"without <w:del> ancestor: {text_preview}"
        )

    return (len(errors) == 0, errors)


def validate_tracked_changes_integrity(
    original_doc: Document, modified_doc: Document
) -> tuple[bool, list[str]]:
    """
    Validate that all non-tracked content remains unchanged.

    From: ~/.agents/.../docx/ooxml/scripts/validation/redlining.py:RedliningValidator
    """
    errors = []

    # Extract text after removing tracked changes
    def extract_text_without_tracked_changes(root):
        """Extract text content, excluding tracked insertions and unwrapping tracked deletions."""
        # Remove w:ins elements
        for ins_elem in root.xpath(".//w:ins", namespaces={"w": WORD_NAMESPACE}):
            ins_elem.getparent().remove(ins_elem)

        # Unwrap w:del elements (keep content but remove the w:del wrapper)
        for del_elem in root.xpath(".//w:del", namespaces={"w": WORD_NAMESPACE}):
            # Convert w:delText to w:t
            for deltext in del_elem.xpath(".//w:delText", namespaces={"w": WORD_NAMESPACE}):
                deltext.tag = f"{{{WORD_NAMESPACE}}}t"
            # Move children to parent
            parent = del_elem.getparent()
            index = list(parent).index(del_elem)
            for child in reversed(list(del_elem)):
                parent.insert(index, child)
            parent.remove(del_elem)

        # Extract text from w:t elements
        paragraphs = []
        for p_elem in root.xpath(".//w:p", namespaces={"w": WORD_NAMESPACE}):
            text_parts = []
            for t_elem in p_elem.xpath(".//w:t", namespaces={"w": WORD_NAMESPACE}):
                if t_elem.text:
                    text_parts.append(t_elem.text)
            paragraph_text = "".join(text_parts)
            if paragraph_text:  # Skip empty paragraphs
                paragraphs.append(paragraph_text)
        return "\n".join(paragraphs)

    # Make copies of the XML trees to avoid modifying originals
    original_root = etree.fromstring(etree.tostring(original_doc.xml_root))
    modified_root = etree.fromstring(etree.tostring(modified_doc.xml_root))

    original_text = extract_text_without_tracked_changes(original_root)
    modified_text = extract_text_without_tracked_changes(modified_root)

    if original_text != modified_text:
        errors.append("Document content changed outside of tracked changes")
        errors.append(f"Original text: {repr(original_text)[:200]}")
        errors.append(f"Modified text: {repr(modified_text)[:200]}")

    return (len(errors) == 0, errors)


# Tests


def test_insert_tracked_produces_valid_xml() -> None:
    """Test that insert_tracked produces well-formed XML."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)
        doc.insert_tracked(" additional text", after="test document")

        valid, message = validate_xml_wellformed(doc)
        assert valid, message

    finally:
        doc_path.unlink()


def test_insert_tracked_whitespace_preservation() -> None:
    """Test that insert_tracked properly handles xml:space='preserve' for whitespace."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        # Insert text with leading space
        doc.insert_tracked(" leading space", after="test document")

        valid, errors = validate_whitespace_preservation(doc)
        assert valid, "Whitespace preservation errors:\n" + "\n".join(errors)

    finally:
        doc_path.unlink()


def test_delete_tracked_uses_deltext() -> None:
    """Test that delete_tracked uses w:delText instead of w:t within w:del."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)
        doc.delete_tracked("multiple")

        valid, errors = validate_deletion_structure(doc)
        assert valid, "Deletion structure errors:\n" + "\n".join(errors)

    finally:
        doc_path.unlink()


def test_replace_tracked_structure() -> None:
    """Test that replace_tracked produces valid tracked change structure."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)
        doc.replace_tracked("test", "sample")

        # Should have proper deletion structure
        valid_del, del_errors = validate_deletion_structure(doc)
        assert valid_del, "Deletion structure errors:\n" + "\n".join(del_errors)

        # Should have proper insertion structure
        valid_ins, ins_errors = validate_insertion_structure(doc)
        assert valid_ins, "Insertion structure errors:\n" + "\n".join(ins_errors)

        # Should preserve whitespace properly
        valid_ws, ws_errors = validate_whitespace_preservation(doc)
        assert valid_ws, "Whitespace preservation errors:\n" + "\n".join(ws_errors)

    finally:
        doc_path.unlink()


def test_tracked_changes_integrity() -> None:
    """Test that tracked changes don't modify non-tracked content."""
    doc_path = create_test_document()
    try:
        original_doc = Document(doc_path)

        # Make a copy for modification
        modified_doc = Document(doc_path)
        modified_doc.insert_tracked(" inserted", after="test document")
        modified_doc.delete_tracked("multiple")

        valid, errors = validate_tracked_changes_integrity(original_doc, modified_doc)
        assert valid, "Content integrity errors:\n" + "\n".join(errors)

    finally:
        doc_path.unlink()


def test_ms365_identity_produces_valid_xml() -> None:
    """Test that MS365 AuthorIdentity produces valid XML."""
    doc_path = create_test_document()
    try:
        identity = AuthorIdentity(
            author="Test User", email="test@example.com", provider_id="AD", guid="test-guid-123"
        )
        doc = Document(doc_path, author=identity)
        doc.insert_tracked(" with identity", after="test document")

        valid, message = validate_xml_wellformed(doc)
        assert valid, message

        valid_ws, ws_errors = validate_whitespace_preservation(doc)
        assert valid_ws, "Whitespace preservation errors:\n" + "\n".join(ws_errors)

    finally:
        doc_path.unlink()


def test_batch_operations_produce_valid_xml() -> None:
    """Test that batch operations produce valid OOXML."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        edits = [
            {"type": "insert_tracked", "text": " A", "after": "document"},
            {"type": "delete_tracked", "text": "multiple"},
            {"type": "replace_tracked", "find": "test", "replace": "sample"},
        ]

        doc.apply_edits(edits)

        # Validate XML structure
        valid, message = validate_xml_wellformed(doc)
        assert valid, message

        # Validate tracked changes structure
        valid_del, del_errors = validate_deletion_structure(doc)
        assert valid_del, "Deletion structure errors:\n" + "\n".join(del_errors)

        valid_ins, ins_errors = validate_insertion_structure(doc)
        assert valid_ins, "Insertion structure errors:\n" + "\n".join(ins_errors)

        valid_ws, ws_errors = validate_whitespace_preservation(doc)
        assert valid_ws, "Whitespace preservation errors:\n" + "\n".join(ws_errors)

    finally:
        doc_path.unlink()


def test_structural_operations_produce_valid_xml() -> None:
    """Test that structural operations (insert_paragraph) produce valid OOXML."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        doc.insert_paragraph(
            "New paragraph with tracked changes", after="test document", track=True
        )

        valid, message = validate_xml_wellformed(doc)
        assert valid, message

        valid_ws, ws_errors = validate_whitespace_preservation(doc)
        assert valid_ws, "Whitespace preservation errors:\n" + "\n".join(ws_errors)

    finally:
        doc_path.unlink()


def test_empty_string_whitespace_handling() -> None:
    """Test that empty strings and whitespace-only text are handled correctly."""
    doc_path = create_test_document()
    try:
        doc = Document(doc_path)

        # Insert text with various whitespace patterns
        doc.insert_tracked("   ", after="test")  # Spaces only
        doc.insert_tracked("\t", after="document")  # Tab

        valid_ws, ws_errors = validate_whitespace_preservation(doc)
        assert valid_ws, "Whitespace preservation errors:\n" + "\n".join(ws_errors)

    finally:
        doc_path.unlink()


def test_save_validates_document() -> None:
    """Test that save() validates the document and raises informative errors."""

    doc_path = create_test_document()
    output_path = Path(tempfile.mktemp(suffix=".docx"))
    try:
        doc = Document(doc_path)
        doc.insert_tracked(" text with leading space", after="test")

        # save() should validate and succeed for valid documents
        doc.save(output_path)
        assert output_path.exists()

        # Verify saved document can be loaded
        loaded_doc = Document(output_path)
        assert "text with leading space" in loaded_doc.get_text()

    finally:
        doc_path.unlink()
        if output_path.exists():
            output_path.unlink()


def test_save_raises_informative_validation_error() -> None:
    """Test that save() raises ValidationError with detailed error list for bug reports."""
    from python_docx_redline.validation import ValidationError

    doc_path = create_test_document()
    output_path = Path(tempfile.mktemp(suffix=".docx"))

    try:
        doc = Document(doc_path)

        # Manually corrupt the document by creating invalid structure
        # Add a w:t element directly within a w:del (should be w:delText)
        root = doc.xml_root
        para = root.find(f".//{{{WORD_NAMESPACE}}}p")
        if para is not None:
            # Create invalid structure: w:del with w:t instead of w:delText
            del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
            run_elem = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            t_elem = etree.Element(f"{{{WORD_NAMESPACE}}}t")
            t_elem.text = "invalid deletion"
            run_elem.append(t_elem)
            del_elem.append(run_elem)
            para.append(del_elem)

            # Attempt to save should raise ValidationError
            try:
                doc.save(output_path)
                assert False, "save() should have raised ValidationError for invalid structure"
            except ValidationError as e:
                # Verify error has informative message
                error_str = str(e)
                assert "validation failed" in error_str.lower()
                assert "bug" in error_str.lower()  # Mentions reporting as bug
                assert hasattr(e, "errors"), "ValidationError should have errors list"
                assert len(e.errors) > 0, "ValidationError should include specific errors"
                # Verify the specific error is in the list
                assert any("<w:t> found within <w:del>" in err for err in e.errors)

    finally:
        doc_path.unlink()
        if output_path.exists():
            output_path.unlink()


def test_validation_case_insensitive_customxml_paths() -> None:
    """Test that validation handles case-insensitive customXml vs customXML paths.

    Microsoft Word sometimes creates documents with inconsistent casing in paths,
    especially for customXml directories. The validation should tolerate this.

    See: docs/issues/ISSUE_CUSTOMXML_CASE_SENSITIVITY.md
    """
    from python_docx_redline.validation_base import BaseSchemaValidator

    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Create a document with case mismatch: reference to /customXML/ but file at customXml/
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/customXml/item1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXml+xml"/>
<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    # Reference with uppercase customXML but file will be at lowercase customXml
    word_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXML/item1.xml"/>
</Relationships>"""

    # customXml items have their own rels file that references itemProps with case mismatch
    custom_xml_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>
</Relationships>"""

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>Test document with customXml.</w:t></w:r></w:p></w:body>
</w:document>"""

    custom_xml_item = """<?xml version="1.0" encoding="UTF-8"?>
<root>Custom data</root>"""

    custom_xml_props = """<?xml version="1.0" encoding="UTF-8"?>
<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{12345678-1234-1234-1234-123456789012}">
<ds:schemaRefs/>
</ds:datastoreItem>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", word_rels)
        # Files at lowercase path, but referenced with uppercase
        docx.writestr("customXml/item1.xml", custom_xml_item)
        docx.writestr("customXml/_rels/item1.xml.rels", custom_xml_rels)
        docx.writestr("customXml/itemProps1.xml", custom_xml_props)

    try:
        # Extract and validate
        with tempfile.TemporaryDirectory() as unpack_dir:
            unpack_path = Path(unpack_dir)
            with zipfile.ZipFile(doc_path, "r") as zip_ref:
                zip_ref.extractall(unpack_path)

            # Create a minimal validator subclass for testing
            class TestValidator(BaseSchemaValidator):
                def validate(self):
                    return self.validate_file_references()

            validator = TestValidator(
                unpacked_dir=unpack_path, original_file=doc_path, verbose=False
            )

            # This should pass due to case-insensitive matching
            result = validator.validate_file_references()
            assert result, (
                "validate_file_references() should pass with case-insensitive matching "
                "for customXml vs customXML paths"
            )

    finally:
        doc_path.unlink()
