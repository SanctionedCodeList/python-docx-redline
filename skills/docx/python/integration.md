# python-docx Integration

python-docx-redline integrates seamlessly with python-docx, allowing you to use both libraries together. This guide covers conversion functions, workflow patterns, and when to use each library.

## Quick Reference

| Function | Purpose |
|----------|---------|
| `from_python_docx(doc)` | Convert python-docx → python-docx-redline |
| `to_python_docx(doc)` | Convert python-docx-redline → python-docx |
| `doc.save_to_bytes()` | Get document as bytes for in-memory workflows |

## Converting Between Libraries

### python-docx → python-docx-redline

Use `from_python_docx()` when you've created or modified a document with python-docx and need to add tracked changes or use python-docx-redline's text search:

```python
from docx import Document as PythonDocxDocument
from python_docx_redline import from_python_docx

# Create document with python-docx
py_doc = PythonDocxDocument()
py_doc.add_heading("Service Agreement", 0)
py_doc.add_paragraph("Payment terms: net 30 days")
py_doc.add_paragraph("Contract duration: 12 months")

# Convert to python-docx-redline for editing
doc = from_python_docx(py_doc, author="Legal Team")

# Now use python-docx-redline features
doc.replace_tracked("net 30 days", "net 45 days")
doc.replace("12 months", "24 months")  # Silent edit
doc.add_comment("Please verify", on="Payment terms")

doc.save("agreement_reviewed.docx")
```

### python-docx-redline → python-docx

Use `to_python_docx()` when you need python-docx features after making edits:

```python
from python_docx_redline import Document
from python_docx_redline.compat import to_python_docx

# Edit with python-docx-redline
doc = Document("contract.docx")
doc.replace_tracked("Contractor", "Service Provider")

# Convert back to python-docx for additional operations
py_doc = to_python_docx(doc)

# Use python-docx features
py_doc.add_page_break()
py_doc.add_paragraph("Appendix A: Additional Terms")
py_doc.core_properties.author = "Legal Department"

py_doc.save("contract_final.docx")
```

## In-Memory Workflows

### Bytes-Based Processing

Both libraries support in-memory operations without touching the filesystem:

```python
from python_docx_redline import Document

# Load from bytes
with open("template.docx", "rb") as f:
    doc_bytes = f.read()

doc = Document(doc_bytes)
doc.replace("{{NAME}}", "John Doe")
doc.replace_tracked("30 days", "45 days")

# Get modified document as bytes
output_bytes = doc.save_to_bytes()

# Use for API responses, S3 upload, email attachment, etc.
```

### Round-Trip Workflow

Create with python-docx, edit with python-docx-redline, finish with python-docx:

```python
from docx import Document as PythonDocxDocument
from python_docx_redline import from_python_docx
from python_docx_redline.compat import to_python_docx

# Step 1: Create structure with python-docx
py_doc = PythonDocxDocument()
py_doc.add_heading("Quarterly Report", 0)
py_doc.add_heading("Executive Summary", 1)
py_doc.add_paragraph("Revenue increased by 15% this quarter.")
py_doc.add_heading("Financial Details", 1)
py_doc.add_paragraph("Net profit: $2.5 million")

# Step 2: Edit with python-docx-redline
doc = from_python_docx(py_doc, author="Finance Team")
doc.replace_tracked("15%", "18%")  # Tracked correction
doc.replace_tracked("$2.5 million", "$2.8 million")
doc.add_comment("Verified by accounting", on="Net profit")

# Step 3: Back to python-docx for final touches
py_doc = to_python_docx(doc)
py_doc.add_page_break()
py_doc.add_paragraph("Report generated: " + str(date.today()))
py_doc.save("Q4_Report_Final.docx")
```

## When to Use Each Library

### Use python-docx for:

| Task | Why |
|------|-----|
| **Creating new documents from scratch** | Full control over structure, styles, tables |
| **Adding content at end of document** | No existing content to navigate |
| **Document properties** | `doc.core_properties.author`, `.title`, etc. |
| **Complex table operations** | Creating, styling, merging cells |
| **Adding page breaks, sections** | Structural elements |
| **Style management** | Applying and modifying styles |

```python
from docx import Document

doc = Document()
doc.add_heading("New Document", 0)
doc.add_paragraph("Created programmatically")

# Add a table
table = doc.add_table(rows=3, cols=3)
table.style = "Table Grid"

doc.save("new_doc.docx")
```

### Use python-docx-redline for:

| Task | Why |
|------|-----|
| **Find and replace** | Handles run fragmentation that breaks python-docx |
| **Tracked changes** | Insert, delete, replace with revision marks |
| **Editing existing documents** | Smart text search across run boundaries |
| **Comments** | Adding comments to specific text |
| **Footnotes/endnotes** | Full CRUD and tracked changes in notes |
| **Scoped operations** | Edit within specific sections |

```python
from python_docx_redline import Document

doc = Document("existing.docx")
doc.replace("Old Company", "New Company")  # Handles run boundaries
doc.replace_tracked("30 days", "45 days")  # With revision marks
doc.add_comment("Please review", on="payment terms")
doc.save("edited.docx")
```

### Use Both for:

| Workflow | Pattern |
|----------|---------|
| **Create then edit** | python-docx → from_python_docx → python-docx-redline |
| **Edit then extend** | python-docx-redline → to_python_docx → python-docx |
| **Full round-trip** | Create → Edit → Finalize |

## Common Patterns

### Template Population with Tracked Changes

```python
from docx import Document as PythonDocxDocument
from python_docx_redline import from_python_docx

# Load template with python-docx
py_doc = PythonDocxDocument("template.docx")

# Convert for editing
doc = from_python_docx(py_doc)

# Silent replacements for placeholders
doc.replace("{{CLIENT_NAME}}", "Acme Corporation")
doc.replace("{{DATE}}", "January 15, 2025")

# Tracked changes for negotiated terms
doc.replace_tracked("30 days", "45 days")
doc.replace_tracked("$10,000", "$12,500")

doc.save("contract_populated.docx")
```

### API/Web Service Processing

```python
from flask import Flask, request, send_file
from python_docx_redline import Document
import io

app = Flask(__name__)

@app.route("/process", methods=["POST"])
def process_document():
    # Receive document bytes
    doc_bytes = request.files["document"].read()

    # Process with python-docx-redline
    doc = Document(doc_bytes)
    doc.replace("DRAFT", "FINAL")
    doc.replace_tracked("v1.0", "v2.0")

    # Return processed document
    output = io.BytesIO(doc.save_to_bytes())
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
```

### Batch Processing

```python
from pathlib import Path
from python_docx_redline import Document

input_dir = Path("contracts/pending")
output_dir = Path("contracts/reviewed")

for docx_file in input_dir.glob("*.docx"):
    doc = Document(docx_file)

    # Apply standard edits
    doc.replace("Contractor", "Service Provider", occurrence="all")
    doc.replace_tracked("net 30", "net 45")

    doc.save(output_dir / docx_file.name)
```

## Error Handling

```python
from python_docx_redline import from_python_docx
from python_docx_redline.errors import TextNotFoundError

try:
    doc = from_python_docx(py_doc)
    doc.replace_tracked("specific text", "replacement")
except ImportError:
    print("python-docx not installed: pip install python-docx")
except TypeError as e:
    print(f"Invalid input: {e}")
except TextNotFoundError as e:
    print(f"Text not found: {e}")
```

## API Reference

### from_python_docx()

```python
from python_docx_redline import from_python_docx

doc = from_python_docx(
    python_docx_doc,           # python-docx Document object
    author="Claude"            # Author for tracked changes (str or AuthorIdentity)
)
```

**Returns:** `python_docx_redline.Document`

**Raises:**
- `ImportError` if python-docx not installed
- `TypeError` if input is not a python-docx Document

### to_python_docx()

```python
from python_docx_redline.compat import to_python_docx

py_doc = to_python_docx(
    doc,                       # python_docx_redline Document
    validate=True              # Run OOXML validation (default: True)
)
```

**Returns:** `docx.Document`

**Raises:**
- `ImportError` if python-docx not installed

### save_to_bytes()

```python
doc_bytes = doc.save_to_bytes(
    validate=True              # Run OOXML validation (default: True)
)
```

**Returns:** `bytes` - The document as a .docx file in bytes
