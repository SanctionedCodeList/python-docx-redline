"""
OOXML validation tests to ensure docx_redline produces valid Office documents.

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

from docx_redline import AuthorIdentity, Document

# Namespace constants
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"


def create_test_document() -> Path:
    """Create a simple test document."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>This is a test document.</w:t></w:r></w:p>
<w:p><w:r><w:t>It has multiple paragraphs.</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

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
    """
    Validate that w:t elements with leading/trailing whitespace have xml:space='preserve'.

    From: ~/.agents/.../docx/ooxml/scripts/validation/docx.py:validate_whitespace_preservation
    """
    errors = []
    root = get_document_xml(doc)

    # Find all w:t elements
    for elem in root.iter(f"{{{WORD_NAMESPACE}}}t"):
        if elem.text:
            text = elem.text
            # Check if text starts or ends with whitespace
            if text.startswith((" ", "\t", "\n")) or text.endswith((" ", "\t", "\n")):
                # Check if xml:space="preserve" attribute exists
                xml_space_attr = f"{{{XML_NAMESPACE}}}space"
                if xml_space_attr not in elem.attrib or elem.attrib[xml_space_attr] != "preserve":
                    text_preview = repr(text)[:50] + "..." if len(repr(text)) > 50 else repr(text)
                    errors.append(
                        f"Line {elem.sourceline}: w:t element with whitespace "
                        f"missing xml:space='preserve': {text_preview}"
                    )

    return (len(errors) == 0, errors)


def validate_deletion_structure(doc: Document) -> tuple[bool, list[str]]:
    """
    Validate that w:t elements are NOT within w:del elements.
    Deletions must use w:delText, not w:t.

    From: ~/.agents/.../docx/ooxml/scripts/validation/docx.py:validate_deletions
    """
    errors = []
    root = get_document_xml(doc)

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
                f"Line {t_elem.sourceline}: <w:t> found within <w:del> (should be <w:delText>): {text_preview}"
            )

    return (len(errors) == 0, errors)


def validate_insertion_structure(doc: Document) -> tuple[bool, list[str]]:
    """
    Validate that w:delText elements are NOT within w:ins elements (unless nested in w:del).

    From: ~/.agents/.../docx/ooxml/scripts/validation/docx.py:validate_insertions
    """
    errors = []
    root = get_document_xml(doc)
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
