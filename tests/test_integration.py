"""
End-to-end integration tests for docx_redline.

These tests verify that all components work together correctly in
realistic workflows.
"""

import tempfile
import zipfile
from pathlib import Path

from docx_redline import Document


def create_realistic_docx() -> Path:
    """Create a realistic test document with multiple sections."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Create document.xml
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Introduction</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This Agreement is entered into on January 1, 2024 between the Contractor and the Client for the provision of services.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Terms and Conditions</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Section 2.1: Payment Terms. The Contractor shall invoice the Client monthly. Payment is due within 30 days of invoice date.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Section 2.2: Termination. Either party may terminate this agreement with 30 days notice.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Confidentiality</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>The Contractor agrees to maintain confidentiality of all Client information.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    # Create minimal .docx structure
    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')

    return doc_path


def test_complete_workflow():
    """Test a complete realistic workflow with multiple operations."""
    doc_path = create_realistic_docx()

    try:
        # Load document
        doc = Document(doc_path)

        # 1. Update party names in Introduction
        doc.replace_tracked(
            find="Contractor", replace="Service Provider", scope="section:Introduction"
        )
        doc.replace_tracked(find="Client", replace="Customer", scope="section:Introduction")

        # 2. Update payment terms in specific section
        doc.replace_tracked(
            find="30 days",
            replace="45 days",
            scope="paragraph_containing:Payment Terms",
        )

        # 3. Add clarification in Introduction section only
        doc.insert_tracked(
            text=" (hereinafter 'the Agreement')",
            after="This Agreement",
            scope="section:Introduction",
        )

        # 4. Remove termination clause
        doc.delete_tracked("Either party may terminate this agreement with 30 days notice.")

        # 5. Add emphasis to confidentiality (note: still "Client" in this section)
        doc.insert_tracked(
            text=" The Customer acknowledges the importance of protecting sensitive business information.",
            after="Client information.",
        )

        # Save modified document
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        # Verify document was saved
        assert output_path.exists()

        # Reload and verify changes were applied
        modified_doc = Document(output_path)
        doc_text = "".join("".join(p.itertext()) for p in modified_doc.xml_root.iter("{*}p"))

        # Check that replacements worked in Introduction
        assert "Service Provider" in doc_text
        assert "Customer" in doc_text
        # Note: Other occurrences of Contractor/Client still exist in other sections

        # Check insertion
        assert "(hereinafter 'the Agreement')" in doc_text

        # Check that tracked change elements exist
        from lxml import etree

        xml_string = etree.tostring(modified_doc.xml_root, encoding="unicode")
        assert "w:ins" in xml_string  # Has insertions

        # Cleanup
        output_path.unlink()

    finally:
        doc_path.unlink()


def test_batch_workflow_with_yaml():
    """Test applying edits from YAML file."""
    doc_path = create_realistic_docx()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create YAML edit file
        yaml_content = """
edits:
  - type: replace_tracked
    find: "Contractor"
    replace: "Service Provider"
    scope: "section:Introduction"

  - type: insert_tracked
    text: " (revised)"
    after: "Agreement"
    scope: "section:Introduction"

  - type: replace_tracked
    find: "30 days"
    replace: "45 days"
    scope: "Payment Terms"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        # Apply edits from file
        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path)

        # Check all edits succeeded
        assert len(results) == 3
        assert all(r.success for r in results), [r.message for r in results if not r.success]

        # Save and verify
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        modified_doc = Document(output_path)
        doc_text = "".join("".join(p.itertext()) for p in modified_doc.xml_root.iter("{*}p"))

        assert "Service Provider" in doc_text
        assert "45 days" in doc_text
        assert "(revised)" in doc_text

        # Cleanup
        output_path.unlink()

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_scoped_edits_workflow():
    """Test complex scoping scenarios."""
    doc_path = create_realistic_docx()

    try:
        doc = Document(doc_path)

        # Edit only in Introduction
        doc.insert_tracked(
            text=" between parties",
            after="Agreement",
            scope="section:Introduction",
        )

        # Edit only in paragraphs mentioning payment
        doc.replace_tracked(find="30 days", replace="60 days", scope="paragraph_containing:Payment")

        # Edit using dictionary scope
        doc.insert_tracked(
            text=" strictly",
            after="maintain",
            scope={"contains": "confidential", "section": "Confidentiality"},
        )

        # Save
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        # Verify
        modified_doc = Document(output_path)
        doc_text = "".join("".join(p.itertext()) for p in modified_doc.xml_root.iter("{*}p"))

        assert " between parties" in doc_text
        assert "60 days" in doc_text
        assert " strictly" in doc_text

        # Cleanup
        output_path.unlink()

    finally:
        doc_path.unlink()


def test_error_recovery_workflow():
    """Test graceful error handling in batch operations."""
    doc_path = create_realistic_docx()

    try:
        doc = Document(doc_path)

        edits = [
            # This will succeed (using scope to disambiguate)
            {
                "type": "replace_tracked",
                "find": "Contractor",
                "replace": "Provider",
                "scope": "section:Introduction",
            },
            # This will fail - text doesn't exist
            {
                "type": "insert_tracked",
                "text": " note",
                "after": "nonexistent text",
            },
            # This should still execute (continue on error)
            {"type": "insert_tracked", "text": " Updated", "after": "Agreement"},
        ]

        results = doc.apply_edits(edits, stop_on_error=False)

        # Should have 3 results
        assert len(results) == 3

        # First should succeed
        assert results[0].success
        assert "Provider" in results[0].message

        # Second should fail
        assert not results[1].success
        assert "not found" in results[1].message.lower()

        # Third should succeed (didn't stop on error)
        assert results[2].success

        # Verify successful edits were applied
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        modified_doc = Document(output_path)
        doc_text = "".join("".join(p.itertext()) for p in modified_doc.xml_root.iter("{*}p"))

        assert "Provider" in doc_text
        assert " Updated" in doc_text

        # Cleanup
        output_path.unlink()

    finally:
        doc_path.unlink()


def test_context_manager_workflow():
    """Test using Document as context manager."""
    doc_path = create_realistic_docx()
    output_path = Path(tempfile.mktemp(suffix=".docx"))

    try:
        # Use context manager
        with Document(doc_path) as doc:
            doc.replace_tracked("Contractor", "Provider", scope="section:Introduction")
            doc.insert_tracked(" (updated)", after="Agreement")
            doc.save(output_path)

        # Verify file was saved
        assert output_path.exists()

        # Verify changes
        doc = Document(output_path)
        doc_text = "".join("".join(p.itertext()) for p in doc.xml_root.iter("{*}p"))

        assert "Provider" in doc_text
        assert "(updated)" in doc_text

    finally:
        doc_path.unlink()
        if output_path.exists():
            output_path.unlink()


# Run tests with: pytest tests/test_integration.py -v
