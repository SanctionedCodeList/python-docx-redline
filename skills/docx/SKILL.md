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

## Viewing and Finding Content

Before making edits, you often need to discover where to make changes. The python-docx-redline library provides comprehensive APIs for inspecting document structure, finding text locations, and viewing existing tracked changes.

### Finding Text Locations

The `find_all()` method helps you discover all occurrences of text before making edits. It returns rich `Match` objects with context information:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Find all occurrences of text
matches = doc.find_all("payment terms")

# Examine each match
for match in matches:
    print(f"Match {match.index}: '{match.text}'")
    print(f"  Context: ...{match.context}...")
    print(f"  Paragraph {match.paragraph_index}: {match.paragraph_text[:50]}...")
    print(f"  Location: {match.location}")
    print(f"  Span: characters {match.span[0]}-{match.span[1]}")
    print()

# Case-insensitive search
matches = doc.find_all("IMPORTANT", case_sensitive=False)

# Regex search
matches = doc.find_all(r"\d+ days", regex=True)

# Scoped search
matches = doc.find_all("Client", scope="section:Payment Terms")
```

**Match object properties:**
- `index` - Zero-based occurrence number (0 for first match, 1 for second, etc.)
- `text` - The matched text
- `context` - Surrounding text for context (configurable window)
- `paragraph_index` - Which paragraph contains this match
- `paragraph_text` - Full text of the containing paragraph
- `location` - Human-readable location description
- `span` - Tuple of (start, end) character positions

**Common workflow:**
1. Find all occurrences with `find_all()`
2. Examine match context to identify the correct one
3. Use `occurrence=N` in edit operations to target specific matches

```python
# Find and review all matches
matches = doc.find_all("30 days")
for m in matches:
    print(f"Match {m.index}: {m.context}")

# After reviewing, target specific occurrence
doc.replace_tracked("30 days", "45 days", occurrence=2)  # Replace 3rd match
```

### Inspecting Document Structure

Access document structure programmatically:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Iterate through all paragraphs
for i, para in enumerate(doc.paragraphs):
    if para.is_heading():
        print(f"Paragraph {i}: HEADING - {para.text}")
        print(f"  Style: {para.style_name}")
    else:
        print(f"Paragraph {i}: {para.text[:50]}...")

# Access sections (hierarchical structure based on headings)
for section in doc.sections:
    print(f"Section: {section.heading_text}")
    if section.contains("payment"):
        print("  Contains payment terms")
```

### Working with Tables

Discover and access tables in the document:

```python
# List all tables
print(f"Document contains {len(doc.tables)} tables")

for i, table in enumerate(doc.tables):
    print(f"Table {i}: {len(table.rows)} rows x {len(table.columns)} columns")
    # Access first row
    for j, cell in enumerate(table.rows[0].cells):
        print(f"  Header {j}: {cell.text}")

# Find specific table by content
table = doc.find_table(containing="Price")
if table:
    print(f"Found pricing table with {len(table.rows)} rows")
    # Edit the table
    table.rows[1].cells[2].text = "$50.00"
else:
    print("No table containing 'Price' found")

# Find table by partial content
invoice_table = doc.find_table(containing="Invoice")
schedule_table = doc.find_table(containing="Schedule")
```

### Viewing Existing Tracked Changes

Inspect existing tracked changes before adding new ones:

```python
from python_docx_redline import Document

doc = Document("contract_redlined.docx")

# Check if document has any tracked changes
if doc.has_tracked_changes():
    print("Document contains tracked changes")

    # Get all tracked changes
    changes = doc.get_tracked_changes()

    for change in changes:
        print(f"{change.type}: '{change.text}' by {change.author}")
        print(f"  Date: {change.date}")
        print()

    # Filter by change type
    insertions = doc.get_tracked_changes(change_type="insert")
    print(f"Found {len(insertions)} insertions")

    deletions = doc.get_tracked_changes(change_type="delete")
    print(f"Found {len(deletions)} deletions")

    # Filter by author
    legal_changes = doc.get_tracked_changes(author="Legal Team")
    print(f"Legal Team made {len(legal_changes)} changes")
else:
    print("No tracked changes found")
```

**TrackedChange object properties:**
- `type` - "insert" or "delete"
- `text` - The text that was inserted or deleted
- `author` - Who made the change
- `date` - When the change was made

### Best Practices for Agents

**1. Always discover before editing:**
```python
# Bad: Blindly edit without checking
doc.replace_tracked("payment", "Payment")  # Might fail if ambiguous

# Good: Discover first, then edit with context
matches = doc.find_all("payment")
if len(matches) == 1:
    doc.replace_tracked("payment", "Payment")
elif len(matches) > 1:
    # Use scope or occurrence to be specific
    doc.replace_tracked("payment", "Payment", scope="section:Terms")
```

**2. Check for existing changes:**
```python
# Check what's already been edited
if doc.has_tracked_changes():
    existing = doc.get_tracked_changes()
    print(f"Warning: Document has {len(existing)} existing changes")
    for change in existing:
        print(f"  {change.author}: {change.type} '{change.text}'")
```

**3. Use structural information:**
```python
# Find the right section first
for section in doc.sections:
    if "payment" in section.heading_text.lower():
        print(f"Found payment section: {section.heading_text}")
        # Now edit within that section
        doc.replace_tracked(
            "30 days",
            "45 days",
            scope=f"section:{section.heading_text}"
        )
        break
```

**4. Combine finding with table editing:**
```python
# Find the pricing table
pricing_table = doc.find_table(containing="Price")
if pricing_table:
    # Inspect it
    print(f"Pricing table has {len(pricing_table.rows)} rows")

    # Make targeted edits
    for row in pricing_table.rows[1:]:  # Skip header
        old_price = row.cells[1].text
        print(f"Current price: {old_price}")
        # Edit if needed
```

**5. Use find_all() to avoid ambiguity errors:**
```python
# When you get "AmbiguousTextError", use find_all to investigate
try:
    doc.replace_tracked("Section", "Article")
except Exception as e:
    if "ambiguous" in str(e).lower():
        # Find all occurrences to see what we're dealing with
        matches = doc.find_all("Section")
        print(f"Found {len(matches)} occurrences:")
        for m in matches:
            print(f"  {m.index}: {m.context}")

        # Now use occurrence or scope to be specific
        doc.replace_tracked("Section", "Article", occurrence=0)
```

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

### Smart Quote Handling

Word documents typically contain "smart" or "curly" quotes (`'` `'` `"` `"`) instead of straight quotes (`'` `"`). The library **automatically normalizes quotes** so you can type straight quotes in Python and match curly quotes in documents:

```python
# Document contains: "The Defendant's motion" (with curly apostrophe U+2019)
# You type straight quotes - it just works!
doc.replace_tracked("Defendant's motion", "party's motion")

# Also works with double quotes
# Document contains: "free trial" (with curly quotes U+201C/U+201D)
doc.replace_tracked('"free trial"', '"subscription"')
```

This is enabled by default (`enable_quote_normalization=True`). To require exact quote matching:

```python
# Only match if document has exact same quote characters
doc.replace_tracked("Defendant's", "party's", enable_quote_normalization=False)
```

**Why this matters:** Legal documents use possessives (`Plaintiff's`, `Defendant's`) and contractions (`don't`, `won't`) extensively. Without quote normalization, you'd need to copy-paste exact characters from Word or use Unicode escapes (`\u2019`).

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
| Find all text occurrences | - | **Best** | Manual |
| Get tracked changes | - | **Best** | Manual |
| Inspect document structure | - | **Best** | Manual |
| Find tables by content | - | **Best** | Manual |
| Insert tracked text | - | **Best** | Possible |
| Delete tracked text | - | **Best** | Possible |
| Replace tracked text | - | **Best** | Possible |
| Regex find/replace | - | **Best** | Manual |
| Smart quote normalization | - | **Best** | Manual |
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
