"""
End-to-end integration tests for python_docx_redline.

These tests verify that all components work together correctly in
realistic workflows.
"""

import tempfile
import zipfile
from pathlib import Path

from python_docx_redline import Document


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

    # Create proper .docx structure
    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document_xml)

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

        # Apply edits from file (use minimal_edits=False for coarse mode assertions)
        doc = Document(doc_path, minimal_edits=False)
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
        # Use minimal_edits=False for coarse mode assertions
        doc = Document(doc_path, minimal_edits=False)

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


# Phase 2 Integration Tests - Structural Operations


def create_structured_contract() -> Path:
    """Create a contract document with multiple sections for structural operations."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Parties</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This Agreement is between Company A and Company B.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Services</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Company A will provide consulting services.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Payment</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Payment terms are net 30 days.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Outdated Clause</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This clause is no longer applicable.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Termination</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Either party may terminate with notice.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

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

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)

    return doc_path


def test_structural_operations_workflow():
    """Test complete workflow with Phase 2 structural operations."""
    doc_path = create_structured_contract()

    try:
        # Use minimal_edits=False for coarse mode assertions
        doc = Document(doc_path, minimal_edits=False)

        # 1. Delete outdated section completely
        doc.delete_section("Outdated Clause", track=True)

        # 2. Add new section between Services and Payment
        doc.insert_paragraph(
            "Deliverables",
            after="Company A will provide consulting services.",
            style="Heading1",
            track=True,
        )

        # 3. Add content to new section
        doc.insert_paragraphs(
            [
                "Company A shall deliver the following:",
                "- Weekly status reports",
                "- Monthly progress reviews",
                "- Final project documentation",
            ],
            after="Deliverables",
            track=True,
        )

        # 4. Update existing text using Phase 1 operations
        doc.replace_tracked("net 30 days", "net 45 days")

        # Save and verify
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        # Reload and verify structure
        modified_doc = Document(output_path)
        doc_text = "".join("".join(p.itertext()) for p in modified_doc.xml_root.iter("{*}p"))

        # Verify deletions
        assert "Outdated Clause" in doc_text  # Still present in w:del
        assert "no longer applicable" in doc_text

        # Verify new section
        assert "Deliverables" in doc_text
        assert "Weekly status reports" in doc_text
        assert "Final project documentation" in doc_text

        # Verify updates
        assert "net 45 days" in doc_text

        # Verify tracked changes exist
        from lxml import etree

        xml_string = etree.tostring(modified_doc.xml_root, encoding="unicode")
        assert "w:ins" in xml_string
        assert "w:del" in xml_string

        output_path.unlink()

    finally:
        doc_path.unlink()


def test_document_restructuring_workflow():
    """Test major document restructuring with mixed operations."""
    doc_path = create_structured_contract()

    try:
        doc = Document(doc_path)

        # Scenario: Client wants to reorganize and update contract

        # Step 1: Update party names  (do both replacements with more specific text)
        doc.replace_tracked(
            "Company A and Company B", "Acme Corp and Widget Inc", scope="section:Parties"
        )

        # Step 2: Add executive summary at the beginning
        doc.insert_paragraph(
            "Executive Summary",
            before="This Agreement is between",
            style="Heading1",
            track=True,
        )
        doc.insert_paragraph(
            "This is a services agreement between Acme Corp and Widget Inc for consulting services.",
            after="Executive Summary",
            track=True,
        )

        # Step 3: Remove outdated section
        doc.delete_section("Outdated Clause", track=True)

        # Step 4: Add compliance section at the end
        doc.insert_paragraph(
            "Compliance",
            after="Either party may terminate with notice.",
            style="Heading1",
            track=True,
        )
        doc.insert_paragraphs(
            [
                "Both parties agree to comply with all applicable laws.",
                "This includes but is not limited to:",
                "- Data protection regulations",
                "- Employment laws",
                "- Industry-specific requirements",
            ],
            after="Compliance",
            track=True,
        )

        # Save and verify
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        modified_doc = Document(output_path)
        doc_text = "".join("".join(p.itertext()) for p in modified_doc.xml_root.iter("{*}p"))

        # Verify all changes
        assert "Acme Corp" in doc_text
        assert "Widget Inc" in doc_text
        assert "Executive Summary" in doc_text
        assert "Compliance" in doc_text
        assert "Data protection regulations" in doc_text

        output_path.unlink()

    finally:
        doc_path.unlink()


def test_batch_structural_operations():
    """Test Phase 2 operations in batch mode with YAML."""
    doc_path = create_structured_contract()
    yaml_path = Path(tempfile.mktemp(suffix=".yaml"))

    try:
        # Create YAML with mixed Phase 1 and Phase 2 operations
        yaml_content = """
edits:
  # Phase 1 text operations
  - type: replace_tracked
    find: "Company A and Company B"
    replace: "TechCorp and ClientCo"
    scope: "section:Parties"

  # Phase 2 structural operations
  - type: delete_section
    heading: "Outdated Clause"
    track: true

  - type: insert_paragraph
    text: "Warranties"
    after: "Payment terms are net 30 days."
    style: "Heading1"
    track: true

  - type: insert_paragraphs
    texts:
      - "TechCorp warrants that services will be performed professionally."
      - "ClientCo warrants that all information provided is accurate."
    after: "Warranties"
    track: true

  # More Phase 1 operations
  - type: insert_tracked
    text: " in writing"
    after: "may terminate"
"""
        yaml_path.write_text(yaml_content, encoding="utf-8")

        # Apply all edits from YAML
        doc = Document(doc_path)
        results = doc.apply_edit_file(yaml_path)

        # All operations should succeed
        assert len(results) == 5
        assert all(r.success for r in results), [r.message for r in results if not r.success]

        # Verify each type
        assert results[0].edit_type == "replace_tracked"
        assert results[1].edit_type == "delete_section"
        assert results[2].edit_type == "insert_paragraph"
        assert results[3].edit_type == "insert_paragraphs"
        assert results[4].edit_type == "insert_tracked"

        # Save and verify content
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        modified_doc = Document(output_path)
        doc_text = "".join("".join(p.itertext()) for p in modified_doc.xml_root.iter("{*}p"))

        assert "TechCorp" in doc_text
        assert "ClientCo" in doc_text
        assert "Warranties" in doc_text
        assert "performed professionally" in doc_text
        assert "in writing" in doc_text

        output_path.unlink()

    finally:
        doc_path.unlink()
        yaml_path.unlink()


def test_incremental_document_building():
    """Test building up a document incrementally with structural operations."""
    doc_path = create_structured_contract()

    try:
        doc = Document(doc_path)

        # Scenario: Start with basic contract, add sections incrementally

        # Add new sections one by one
        doc.insert_paragraph(
            "Intellectual Property",
            after="Payment terms are net 30 days.",
            style="Heading1",
            track=True,
        )

        doc.insert_paragraphs(
            [
                "All work product shall be owned by ClientCo.",
                "TechCorp retains ownership of pre-existing materials.",
            ],
            after="Intellectual Property",
            track=True,
        )

        doc.insert_paragraph(
            "Liability",
            after="TechCorp retains ownership of pre-existing materials.",
            style="Heading1",
            track=True,
        )

        doc.insert_paragraph(
            "Liability shall be limited to the fees paid under this Agreement.",
            after="Liability",
            track=True,
        )

        doc.insert_paragraph(
            "Governing Law",
            after="Liability shall be limited to the fees paid under this Agreement.",
            style="Heading1",
            track=True,
        )

        doc.insert_paragraph(
            "This Agreement shall be governed by the laws of California.",
            after="Governing Law",
            track=True,
        )

        # Save and verify structure
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        modified_doc = Document(output_path)

        # Verify content includes all new sections
        # (Don't count headings directly because they may be in w:ins elements)

        doc_text = "".join("".join(p.itertext()) for p in modified_doc.xml_root.iter("{*}p"))
        assert "Intellectual Property" in doc_text
        assert "Liability" in doc_text
        assert "Governing Law" in doc_text

        output_path.unlink()

    finally:
        doc_path.unlink()


def test_section_replacement_workflow():
    """Test replacing a section by deleting old and inserting new."""
    doc_path = create_structured_contract()

    try:
        doc = Document(doc_path)

        # Scenario: Replace "Services" section with more detailed version

        # Delete old section
        doc.delete_section("Services", track=True)

        # Add new detailed Services section
        doc.insert_paragraph(
            "Services and Deliverables",
            after="This Agreement is between Company A and Company B.",
            style="Heading1",
            track=True,
        )

        doc.insert_paragraphs(
            [
                "Company A will provide the following services:",
                "1. Strategic consulting",
                "2. Implementation support",
                "3. Training and documentation",
                "4. Ongoing maintenance and support",
            ],
            after="Services and Deliverables",
            track=True,
        )

        # Save and verify
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        doc.save(output_path)

        modified_doc = Document(output_path)
        doc_text = "".join("".join(p.itertext()) for p in modified_doc.xml_root.iter("{*}p"))

        # Old section still visible in tracked deletion
        assert "Services" in doc_text

        # New section content
        assert "Services and Deliverables" in doc_text
        assert "Strategic consulting" in doc_text
        assert "Ongoing maintenance and support" in doc_text

        output_path.unlink()

    finally:
        doc_path.unlink()


# Run tests with: pytest tests/test_integration.py -v
