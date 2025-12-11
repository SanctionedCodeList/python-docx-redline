---
name: docx
description: "Word document creation, editing, and tracked changes. Use python-docx for creation, python-docx-redline for tracked changes/redlining."
---

# Word Document Editing Skill

## Installation

If imports fail, install the required packages in the relevant Python environment:

```bash
# For document creation
pip install python-docx

# For tracked changes / redlining
pip install git+https://github.com/parkerhancock/python-docx-redline.git
```

## Quick Decision Tree

```
What do you need to do?
│
├── Creating a new document from scratch?
│   └── Use python-docx (Section 1)
│
├── Adding tracked changes (redlining)?
│   └── Use python-docx-redline (Section 2) ★ RECOMMENDED
│
├── Reading/analyzing document content?
│   └── Use python-docx-redline or python-docx (Section 3)
│
└── Complex edge cases (comments, images, nested changes)?
    └── Use raw OOXML manipulation (Section 4)
```

---

## Section 1: Creating New Documents (python-docx)

For creating new Word documents from scratch:

```python
from docx import Document

# Create new document
doc = Document()

# Add content
doc.add_heading("Contract Agreement", 0)
doc.add_paragraph("This agreement is entered into on...")
doc.add_heading("Terms and Conditions", level=1)
doc.add_paragraph("1. Payment shall be due within 30 days.")

# Add a table
table = doc.add_table(rows=2, cols=2)
table.cell(0, 0).text = "Item"
table.cell(0, 1).text = "Price"

# Save
doc.save("contract.docx")
```

**When to use python-docx:**
- Creating documents from scratch
- Adding paragraphs, headings, tables, images
- Setting styles and formatting
- You don't need tracked changes

---

## Section 2: Tracked Changes / Redlining (python-docx-redline)

For editing documents with tracked changes (redlining):

```python
from python_docx_redline import Document

# Load existing document
doc = Document("contract.docx")

# Insert text with tracked changes
doc.insert_tracked(" (as amended)", after="Section 2.1")

# Replace text with tracked changes
doc.replace_tracked("30 days", "45 days")

# Delete text with tracked changes
doc.delete_tracked("subject to approval")

# Save - changes appear as tracked in Word
doc.save("contract_redlined.docx")
```

### Key Features

**Smart Text Search** - Handles text fragmented across XML runs:
```python
# Works even if "30 days" is split across multiple <w:r> elements
doc.replace_tracked("30 days", "45 days")
```

**Regex Support** - Pattern matching with capture groups:
```python
# Update all occurrences of "X days" to "X business days"
doc.replace_tracked(r"(\d+) days", r"\1 business days", regex=True)
```

**Scoped Edits** - Limit changes to specific sections:
```python
doc.replace_tracked(
    "Client",
    "Customer",
    scope="section:Payment Terms"
)
```

**Batch Operations** - Apply multiple edits from YAML:
```python
doc.apply_edit_file("edits.yaml")
```

**MS365 Identity** - Link changes to real users:
```python
from python_docx_redline import Document, AuthorIdentity

identity = AuthorIdentity(
    author="Parker Hancock",
    email="parker@company.com",
    provider_id="AD",
    guid="c5c513d2-1f51-4d69-ae91-17e5787f9bfc"
)
doc = Document("contract.docx", author=identity)
```

### Integration with python-docx

Create with python-docx, then add tracked changes:

```python
from docx import Document as PythonDocxDocument
from python_docx_redline import from_python_docx

# Create document with python-docx
py_doc = PythonDocxDocument()
py_doc.add_heading("Contract", 0)
py_doc.add_paragraph("Payment terms: net 30 days")

# Convert to python-docx-redline for tracked edits
doc = from_python_docx(py_doc, author="Legal Team")
doc.replace_tracked("net 30 days", "net 45 days")
doc.save("contract_redlined.docx")
```

**When to use python-docx-redline:**
- Adding tracked insertions, deletions, or replacements
- Editing contracts, legal documents, or any document requiring revision history
- Batch processing multiple similar edits
- When you need changes to be visible in Word's Track Changes view

---

## Section 3: Reading/Analyzing Documents

### Programmatic access with python-docx-redline:
```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Get all text
full_text = doc.get_text()

# Iterate paragraphs
for para in doc.paragraphs:
    if para.is_heading():
        print(f"Heading: {para.text}")
    else:
        print(f"  {para.text}")

# Access sections
for section in doc.sections:
    print(f"Section: {section.heading_text}")
    if section.contains("payment"):
        print("  Contains payment terms")
```

### Quick text extraction with pandoc:
```bash
pandoc --track-changes=all document.docx -o output.md
```

---

## Section 4: Raw OOXML Manipulation (Advanced)

Only use raw OOXML for scenarios not supported by the libraries above:
- Adding comments with tracked changes
- Modifying another author's tracked changes
- Inserting images with tracked changes
- Complex nested revision scenarios

### Workflow:

1. **Unpack** the document:
```bash
unzip document.docx -d unpacked/
```

2. **Edit** `word/document.xml` directly with proper OOXML patterns

3. **Repack**:
```bash
cd unpacked && zip -r ../modified.docx *
```

### Tracked Change XML Patterns:

**Insertion:**
```xml
<w:ins w:id="1" w:author="Claude" w:date="2025-01-15T10:00:00Z">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

**Deletion:**
```xml
<w:del w:id="2" w:author="Claude" w:date="2025-01-15T10:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

---

## Comparison Table

| Task | python-docx | python-docx-redline | Raw OOXML |
|------|-------------|---------------------|-----------|
| Create new document | **Best** | - | Possible |
| Add paragraphs/headings | **Best** | - | Possible |
| Add tables | **Best** | - | Possible |
| Insert tracked text | - | **Best** | Possible |
| Delete tracked text | - | **Best** | Possible |
| Replace tracked text | - | **Best** | Possible |
| Regex find/replace | - | **Best** | Manual |
| Scoped edits | - | **Best** | Manual |
| Batch from YAML | - | **Best** | Manual |
| Add comments | - | - | **Required** |
| Modify other's changes | - | - | **Required** |

---

## Examples

### Contract Redlining Workflow

```python
from python_docx_redline import Document

# Load the contract
doc = Document("original_contract.docx", author="Legal Team")

# Apply standard amendments
doc.replace_tracked("net 30 days", "net 45 days")
doc.replace_tracked("Contractor", "Service Provider")
doc.insert_tracked(" (as amended)", after="Agreement dated")
doc.delete_tracked("subject to board approval")

# Save with tracked changes visible
doc.save("contract_redlined.docx")
```

### Batch Processing from YAML

```yaml
# edits.yaml
edits:
  - type: replace_tracked
    find: "net 30 days"
    replace: "net 45 days"

  - type: replace_tracked
    find: "Contractor"
    replace: "Service Provider"

  - type: insert_tracked
    text: " (as amended)"
    after: "Agreement dated"

  - type: delete_tracked
    text: "subject to board approval"
```

```python
from python_docx_redline import Document

doc = Document("contract.docx")
results = doc.apply_edit_file("edits.yaml")

for result in results:
    print(f"{'OK' if result.success else 'FAIL'} {result.message}")

doc.save("contract_edited.docx")
```

### Create + Edit Workflow

```python
from docx import Document as PythonDocxDocument
from python_docx_redline import from_python_docx

# Step 1: Create with python-docx
py_doc = PythonDocxDocument()
py_doc.add_heading("Service Agreement", 0)
py_doc.add_paragraph("This agreement between Company and Contractor...")
py_doc.add_paragraph("Payment: Net 30 days from invoice date.")

# Step 2: Add tracked changes with python-docx-redline
doc = from_python_docx(py_doc, author="Legal Review")
doc.replace_tracked("Contractor", "Service Provider")
doc.replace_tracked("Net 30 days", "Net 45 days")

# Step 3: Save
doc.save("agreement_reviewed.docx")
```
