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

Limit changes to specific sections:

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

### Batch Operations from YAML

Apply multiple edits from a configuration file:

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

  # Regex operations
  - type: replace_tracked
    find: "(\d+) days"
    replace: "\\1 business days"
    regex: true
```

```python
from python_docx_redline import Document

doc = Document("contract.docx")
results = doc.apply_edit_file("edits.yaml")

for result in results:
    print(f"{'OK' if result.success else 'FAIL'} {result.message}")

doc.save("contract_edited.docx")
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

## Advanced: Raw OOXML editing

Use raw OOXML manipulation only for scenarios not supported by python-docx-redline:
- Modifying another author's tracked changes
- Inserting images with tracked changes
- Complex nested revision scenarios

### Helper Scripts

This skill includes helper scripts from [Anthropic Skills](https://github.com/anthropics/skills) in the `scripts/` and `ooxml/` directories:

```python
# Using the Document class for comments and tracked changes
from skills.docx.scripts.document import Document

doc = Document('workspace/unpacked', author="Claude")
node = doc["word/document.xml"].get_node(tag="w:del", attrs={"w:id": "1"})
doc.add_comment(start=node, end=node, text="Please review this deletion")
doc.save()
```

See `scripts/README.md` and `ooxml.md` for detailed documentation.

### Workflow

1. **Unpack** the document:
   ```bash
   unzip document.docx -d unpacked/
   # Or use the helper script:
   python -m skills.docx.ooxml.scripts.unpack document.docx unpacked/
   ```

2. **Edit** `word/document.xml` directly with proper OOXML patterns (or use the helper scripts)

3. **Repack**:
   ```bash
   cd unpacked && zip -r ../modified.docx *
   # Or use the helper script:
   python -m skills.docx.ooxml.scripts.pack unpacked/ modified.docx
   ```

### Schema Compliance
- **Element ordering in `<w:pPr>`**: `<w:pStyle>`, `<w:numPr>`, `<w:spacing>`, `<w:ind>`, `<w:jc>`
- **Whitespace**: Add `xml:space='preserve'` to `<w:t>` elements with leading/trailing spaces
- **Unicode**: Escape characters in ASCII content: `"` becomes `&#8220;`
- **RSIDs must be 8-digit hex**: Use values like `00AB1234` (only 0-9, A-F characters)

### Tracked Change XML Patterns

**Text Insertion:**
```xml
<w:ins w:id="1" w:author="Claude" w:date="2025-01-15T10:00:00Z">
  <w:r w:rsidR="00AB1234">
    <w:t>inserted text</w:t>
  </w:r>
</w:ins>
```

**Text Deletion:**
```xml
<w:del w:id="2" w:author="Claude" w:date="2025-01-15T10:00:00Z">
  <w:r w:rsidDel="00AB1234">
    <w:delText>deleted text</w:delText>
  </w:r>
</w:del>
```

**Minimal Edit Principle:**
Only mark text that actually changes. Keep ALL unchanged text outside `<w:del>`/`<w:ins>` tags.

```xml
<!-- BAD - Replaces entire sentence -->
<w:del><w:r><w:delText>The term is 30 days.</w:delText></w:r></w:del>
<w:ins><w:r><w:t>The term is 60 days.</w:t></w:r></w:ins>

<!-- GOOD - Only marks what changed -->
<w:r><w:t>The term is </w:t></w:r>
<w:del><w:r><w:delText>30</w:delText></w:r></w:del>
<w:ins><w:r><w:t>60</w:t></w:r></w:ins>
<w:r><w:t> days.</w:t></w:r>
```

**Deleting Another Author's Insertion:**
```xml
<!-- Nest deletion inside the original insertion -->
<w:ins w:author="Jane Smith" w:id="16">
  <w:del w:author="Claude" w:id="40">
    <w:r><w:delText>monthly</w:delText></w:r>
  </w:del>
</w:ins>
<w:ins w:author="Claude" w:id="41">
  <w:r><w:t>weekly</w:t></w:r>
</w:ins>
```

**Restoring Another Author's Deletion:**
```xml
<!-- Leave their deletion unchanged, add new insertion after it -->
<w:del w:author="Jane Smith" w:id="50">
  <w:r><w:delText>within 30 days</w:delText></w:r>
</w:del>
<w:ins w:author="Claude" w:id="51">
  <w:r><w:t>within 30 days</w:t></w:r>
</w:ins>
```

### Document Content Patterns

**Basic Structure:**
```xml
<w:p>
  <w:r><w:t>Text content</w:t></w:r>
</w:p>
```

**Headings:**
```xml
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Section Heading</w:t></w:r>
</w:p>
```

**Text Formatting:**
```xml
<!-- Bold -->
<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>Bold</w:t></w:r>
<!-- Italic -->
<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>Italic</w:t></w:r>
<!-- Underline -->
<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Underlined</w:t></w:r>
```

**Tables:**
```xml
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="0" w:type="auto"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="4675"/><w:gridCol w:w="4675"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell 1</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell 2</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
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
| Scoped edits | - | **Best** | Manual |
| Batch from YAML | - | **Best** | Manual |
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
