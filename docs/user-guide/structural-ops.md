# Structural Operations

Add, remove, and reorganize document structure with tracked changes.

## Insert Paragraphs

Add new paragraphs with optional styles:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Insert a single paragraph
doc.insert_paragraph(
    "New Section Heading",
    after="Introduction content",
    style="Heading1",
    track=True
)

# Insert multiple paragraphs at once
doc.insert_paragraphs(
    [
        "First key point",
        "Second key point",
        "Third key point"
    ],
    after="New Section Heading",
    track=True
)
```

## Delete Sections

Remove entire sections by heading:

```python
# Delete section and all content until next heading
doc.delete_section("Outdated Provisions", track=True)
```

## Combining Operations

Mix structural and text operations in a single workflow:

```python
doc = Document("contract.docx")

# Add new section
doc.insert_paragraph(
    "Amendments",
    after="Section 5",
    style="Heading1",
    track=True
)

# Add content to the new section
doc.insert_paragraphs(
    [
        "This Agreement may be amended by mutual written consent.",
        "Any amendments shall be effective upon execution by both parties."
    ],
    after="Amendments",
    track=True
)

# Update terminology throughout
doc.replace_tracked("old term", "new term")

# Remove obsolete section
doc.delete_section("Legacy Terms", track=True)

doc.save("contract_restructured.docx")
```

## Working with Section Objects

Access document structure programmatically:

```python
# Iterate through sections
for section in doc.sections:
    print(f"Section: {section.heading_text}")
    print(f"  Level: {section.heading_level}")
    print(f"  Paragraphs: {len(section.paragraphs)}")

    # Check content
    if section.contains("confidential"):
        print("  Contains confidential information")
```

## Working with Paragraph Objects

Access individual paragraphs:

```python
for para in doc.paragraphs:
    if para.is_heading():
        print(f"Heading: {para.text}")
    else:
        print(f"  {para.text[:50]}...")
```

## Next Steps

- [Viewing Content](viewing-content.md) — More ways to read documents
- [Batch Operations](batch-operations.md) — Apply structural changes from YAML
- [API Reference](../PROPOSED_API.md) — Complete method documentation
