---
name: docx
description: "Comprehensive document creation, editing, and analysis with support for tracked changes, comments, formatting preservation, and text extraction. When Claude needs to work with professional documents (.docx files) for: (1) Creating new documents, (2) Modifying or editing content, (3) Working with tracked changes, (4) Adding comments, or any other document tasks"
---

# DOCX creation, editing, and analysis

## Overview

A user may ask you to create, edit, or analyze the contents of a .docx file. A .docx file is essentially a ZIP archive containing XML files and other resources that you can read or edit. You have different tools and workflows available for different tasks.

## Installation

If imports fail, install the required packages in the relevant Python environment:

```bash
# For document creation
pip install python-docx

# For tracked changes / redlining (recommended for editing)
pip install git+https://github.com/parkerhancock/python-docx-redline.git

# For text extraction
# Install pandoc: brew install pandoc (macOS) or apt-get install pandoc (Linux)
```

## Workflow Decision Tree

### Reading/Analyzing Content
Use "Text extraction" or "Raw XML access" sections below

### Creating New Document
Use python-docx - see "Creating a new Word document" section

### Editing Existing Document
- **Simple tracked changes (insert/delete/replace)**
  Use **python-docx-redline** (recommended) - see "Redlining with python-docx-redline"

- **Complex scenarios (comments, images, nested changes)**
  Use raw OOXML manipulation - see "Advanced: Raw OOXML editing"

- **Legal, academic, business, or government docs**
  Use **python-docx-redline** for tracked changes

## Reading and analyzing content

### Text extraction
If you just need to read the text contents of a document, you should convert the document to markdown using pandoc. Pandoc provides excellent support for preserving document structure and can show tracked changes:

```bash
# Convert document to markdown with tracked changes
pandoc --track-changes=all path-to-file.docx -o output.md
# Options: --track-changes=accept/reject/all
```

### Programmatic text access
Use python-docx-redline for programmatic access:

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

### Raw XML access
You need raw XML access for: comments, complex formatting, document structure, embedded media, and metadata. For any of these features, you'll need to unpack a document and read its raw XML contents.

#### Unpacking a file
```bash
unzip document.docx -d unpacked/
```

#### Key file structures
* `word/document.xml` - Main document contents
* `word/comments.xml` - Comments referenced in document.xml
* `word/media/` - Embedded images and media files
* Tracked changes use `<w:ins>` (insertions) and `<w:del>` (deletions) tags

## Creating a new Word document

When creating a new Word document from scratch, use **python-docx**:

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

### python-docx features
- Add paragraphs, headings, tables, images
- Set styles and formatting
- Create lists (numbered and bulleted)
- Add headers and footers
- Insert page breaks

## Redlining with python-docx-redline

**This is the recommended approach for editing documents with tracked changes.** The python-docx-redline library handles the complexity of OOXML tracked changes automatically.

### Basic Operations

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

### Smart Text Search

The library handles text fragmented across multiple XML runs automatically:

```python
# Works even if "30 days" is split across multiple <w:r> elements
doc.replace_tracked("30 days", "45 days")
```

### Regex Support

Use regular expressions with capture groups:

```python
# Update all occurrences of "X days" to "X business days"
doc.replace_tracked(r"(\d+) days", r"\1 business days", regex=True)

# Redact dollar amounts
doc.replace_tracked(r"\$[\d,]+\.?\d*", "$XXX.XX", regex=True)

# Swap date format from MM/DD/YYYY to DD/MM/YYYY
doc.replace_tracked(r"(\d{2})/(\d{2})/(\d{4})", r"\2/\1/\3", regex=True)
```

### Scoped Edits

Limit changes to specific sections using string shortcuts:

```python
# Only modify in Payment Terms section
doc.replace_tracked(
    "Client",
    "Customer",
    scope="section:Payment Terms"
)

# Only modify paragraphs containing specific text
doc.replace_tracked(
    "30 days",
    "45 days",
    scope="paragraph_containing:payment"
)
```

**Dictionary format** for complex filtering:

```python
# Combine multiple conditions
doc.replace_tracked(
    "Client",
    "Customer",
    scope={
        "contains": "payment",      # Paragraph must contain this
        "not_contains": "excluded", # Paragraph must NOT contain this
        "section": "Terms"          # Must be under this heading
    }
)
```

**Callable scope** for custom logic:

```python
# Custom filter function
def my_filter(paragraph):
    text = "".join(paragraph.itertext())
    return "important" in text.lower() and len(text) > 100

doc.replace_tracked("old", "new", scope=my_filter)
```

### Batch Operations

Apply multiple edits programmatically or from YAML files.

**From a list of dictionaries:**

```python
from python_docx_redline import Document

doc = Document("contract.docx")

edits = [
    {"type": "replace_tracked", "find": "net 30 days", "replace": "net 45 days"},
    {"type": "replace_tracked", "find": "Contractor", "replace": "Service Provider"},
    {"type": "insert_tracked", "text": " (as amended)", "after": "Agreement dated"},
    {"type": "delete_tracked", "text": "subject to board approval"},
]

# Apply all edits, continue even if some fail
results = doc.apply_edits(edits, stop_on_error=False)

# Check results
for i, result in enumerate(results):
    if result.success:
        print(f"Edit {i}: OK - {result.message}")
    else:
        print(f"Edit {i}: FAILED - {result.message}")
        # Access the actual exception for debugging
        if result.error:
            print(f"  Error type: {type(result.error).__name__}")

doc.save("contract_edited.docx")
```

**From YAML file:**

```yaml
# edits.yaml
edits:
  - type: replace_tracked
    find: "net 30 days"
    replace: "net 45 days"

  - type: insert_tracked
    text: " (as amended)"
    after: "Agreement dated"
    scope: "section:Introduction"  # Scoped edit

  # Regex operations
  - type: replace_tracked
    find: "(\d+) days"
    replace: "\\1 business days"
    regex: true
```

```python
results = doc.apply_edit_file("edits.yaml", stop_on_error=False)
print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
```

### Image Insertion

Insert images into documents with optional tracked changes:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Basic image insertion
doc.insert_image("logo.png", after="Company Name:")

# With custom dimensions
doc.insert_image(
    "chart.png",
    after="Figure 1:",
    width_inches=4.0,
    height_inches=3.0
)

# With alt text for accessibility
doc.insert_image(
    "diagram.png",
    after="See diagram:",
    name="Process Diagram",
    description="Workflow diagram showing approval process"
)

# Tracked image insertion (appears in Word's review pane)
doc.insert_image_tracked(
    "signature.png",
    after="Authorized By:",
    author="Legal Team"
)

doc.save("contract_with_images.docx")
```

Supported formats: PNG, JPEG, GIF, BMP, TIFF.

### MS365 Identity Integration

Link changes to real MS365 users with full identity information:

```python
from python_docx_redline import Document, AuthorIdentity

identity = AuthorIdentity(
    author="Parker Hancock",
    email="parker@company.com",
    provider_id="AD",
    guid="c5c513d2-1f51-4d69-ae91-17e5787f9bfc"
)

doc = Document("contract.docx", author=identity)
doc.replace_tracked("old term", "new term")
doc.save("contract_edited.docx")
# Changes now appear in Word with full user profile
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

### In-Memory Workflows

Work without filesystem:

```python
# Load from bytes
with open("contract.docx", "rb") as f:
    doc = Document(f.read())

# Make changes
doc.replace_tracked("old", "new")

# Get bytes for storage/transmission
doc_bytes = doc.save_to_bytes()
```

### Adding Comments

Add comments to any text, including text inside tracked changes:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Add comment on normal text
comment = doc.add_comment("Please review this", on="Section 2.1")

# Add comment on tracked insertion (text inside w:ins)
doc.insert_tracked("new clause", after="Agreement")
doc.add_comment("Review this addition", on="new clause")

# Add comment on tracked deletion (text inside w:del)
doc.delete_tracked("old term")
doc.add_comment("Why was this removed?", on="old term")

# Comments can span tracked/untracked boundaries
doc.add_comment("Check this section", on="normal inserted")

# Access existing comments
for comment in doc.comments:
    print(f"{comment.author}: {comment.text} on '{comment.marked_text}'")

doc.save("contract_with_comments.docx")
```

### Document Management

Accept, reject, or clean up tracked changes and comments:

```python
from python_docx_redline import Document

doc = Document("contract_with_changes.docx")

# Accept all tracked changes (removes revision marks, keeps new text)
result = doc.accept_all_changes()
print(f"Accepted {result.insertions} insertions, {result.deletions} deletions")

# Or reject all tracked changes (removes revision marks, keeps original text)
result = doc.reject_all_changes()

# Delete all comments
doc.delete_all_comments()

doc.save("contract_clean.docx")
```

### Text Formatting with Tracking

Apply formatting changes as tracked revisions:

```python
# Format text with tracked changes (shows in Word's review pane)
doc.format_tracked(
    "IMPORTANT",
    bold=True,
    color="#FF0000",
    occurrence="all"  # Format all occurrences: "first", "last", "all", or index
)

# Paragraph formatting with tracking
doc.format_paragraph_tracked(
    containing="Section 1",
    alignment="center",
    spacing_after=12.0
)
```

## Advanced: Raw OOXML Editing

For scenarios not covered by python-docx-redline, you may need to manipulate the raw OOXML directly:
- Modifying another author's tracked changes
- Complex nested revision scenarios
- Custom XML manipulations

See **[ooxml.md](./ooxml.md)** for comprehensive documentation including:
- XML patterns for tracked changes, tables, formatting
- Schema compliance rules (element ordering, whitespace, RSIDs)
- Helper scripts for unpacking/repacking documents
- Document library (Python) for OOXML manipulation

Quick reference for unpack/repack workflow:
```bash
# Unpack
unzip document.docx -d unpacked/

# Edit word/document.xml

# Repack
cd unpacked && zip -r ../modified.docx *
```

## Converting Documents to Images

### Using python-docx-redline (Recommended)

The easiest way to render documents to images:

```python
from python_docx_redline import Document
from python_docx_redline.rendering import is_rendering_available

if is_rendering_available():
    doc = Document("contract.docx")
    images = doc.render_to_images(output_dir="./images", dpi=150)
    for img in images:
        print(f"Generated: {img}")  # page-1.png, page-2.png, etc.
```

Or use the standalone function:

```python
from python_docx_redline.rendering import render_document_to_images

images = render_document_to_images("contract.docx", dpi=200)
```

**Why render documents?**
- AI agents can visually inspect document layout
- See how tracked changes appear (strikethrough, underlines)
- Verify formatting before sending to clients

### Manual Command-Line Approach

Alternatively, convert using shell commands:

1. **Convert DOCX to PDF**:
   ```bash
   soffice --headless --convert-to pdf document.docx
   ```

2. **Convert PDF pages to images**:
   ```bash
   pdftoppm -png -r 150 document.pdf page
   ```
   Creates files like `page-1.png`, `page-2.png`, etc.

Options:
- `-r 150`: Sets resolution to 150 DPI
- `-png` or `-jpeg`: Output format
- `-f N`: First page to convert
- `-l N`: Last page to convert

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
| Scoped edits (section/dict/callable) | - | **Best** | Manual |
| Batch operations (list or YAML) | - | **Best** | Manual |
| Accept/reject all changes | - | **Best** | Manual |
| Format text w/ tracking | - | **Best** | Manual |
| Render to images | - | **Best** | Manual |
| Add comments | - | **Best** | Possible |
| Comments on tracked text | - | **Best** | Manual |
| Insert images | - | **Best** | Manual |
| Insert images w/ tracking | - | **Best** | Manual |
| Modify other's changes | - | - | **Required** |

## Code Style Guidelines

**IMPORTANT**: When generating code for DOCX operations:
- Write concise code
- Avoid verbose variable names and redundant operations
- Avoid unnecessary print statements

## Dependencies

Required dependencies (install if not available):

- **python-docx**: `pip install python-docx` (for creating new documents)
- **python-docx-redline**: `pip install git+https://github.com/parkerhancock/python-docx-redline.git` (for tracked changes)
- **pandoc**: `brew install pandoc` or `apt-get install pandoc` (for text extraction)
- **LibreOffice**: `brew install --cask libreoffice` or `apt-get install libreoffice` (for PDF conversion)
- **Poppler**: `brew install poppler` or `apt-get install poppler-utils` (for pdftoppm)
