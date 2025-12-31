# Reading and Analyzing Documents

## Text Extraction with Pandoc

For simple text extraction, pandoc provides excellent support:

```bash
# Convert to markdown with tracked changes visible
pandoc --track-changes=all document.docx -o output.md

# Options: --track-changes=accept/reject/all
```

## Programmatic Text Access

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
```

## Finding Text with find_all()

The `find_all()` method discovers all occurrences before editing:

```python
matches = doc.find_all("payment terms")

for match in matches:
    print(f"Match {match.index}: '{match.text}'")
    print(f"  Context: ...{match.context}...")
    print(f"  Paragraph {match.paragraph_index}: {match.paragraph_text[:50]}...")
    print(f"  Location: {match.location}")
    print(f"  Span: characters {match.span[0]}-{match.span[1]}")
```

**Match properties:**
- `index` - Zero-based occurrence number
- `text` - The matched text
- `context` - Surrounding text for context
- `paragraph_index` - Which paragraph contains this match
- `paragraph_text` - Full text of the containing paragraph
- `location` - Human-readable location description
- `span` - Tuple of (start, end) character positions

**Search options:**
```python
# Case-insensitive
matches = doc.find_all("IMPORTANT", case_sensitive=False)

# Regex
matches = doc.find_all(r"\d+ days", regex=True)

# Scoped
matches = doc.find_all("Client", scope="section:Payment Terms")
```

## Document Structure

```python
# Iterate paragraphs with metadata
for i, para in enumerate(doc.paragraphs):
    if para.is_heading():
        print(f"Paragraph {i}: HEADING - {para.text}")
        print(f"  Style: {para.style_name}")
    else:
        print(f"Paragraph {i}: {para.text[:50]}...")

# Access sections (hierarchical based on headings)
for section in doc.sections:
    print(f"Section: {section.heading_text}")
    if section.contains("payment"):
        print("  Contains payment terms")
```

## Working with Tables

```python
# List all tables
print(f"Document contains {len(doc.tables)} tables")

for i, table in enumerate(doc.tables):
    print(f"Table {i}: {len(table.rows)} rows x {len(table.columns)} columns")
    for j, cell in enumerate(table.rows[0].cells):
        print(f"  Header {j}: {cell.text}")

# Find table by content
table = doc.find_table(containing="Price")
if table:
    print(f"Found pricing table with {len(table.rows)} rows")
    table.rows[1].cells[2].text = "$50.00"  # Edit cell
```

## Viewing Existing Tracked Changes

```python
if doc.has_tracked_changes():
    changes = doc.get_tracked_changes()

    for change in changes:
        print(f"{change.type}: '{change.text}' by {change.author}")
        print(f"  Date: {change.date}")

    # Filter by type
    insertions = doc.get_tracked_changes(change_type="insert")
    deletions = doc.get_tracked_changes(change_type="delete")

    # Filter by author
    legal_changes = doc.get_tracked_changes(author="Legal Team")
```

**TrackedChange properties:**
- `type` - "insert" or "delete"
- `text` - The text that was inserted or deleted
- `author` - Who made the change
- `date` - When the change was made

## Raw XML Access

For comments, complex formatting, or metadata, unpack the document:

```bash
unzip document.docx -d unpacked/
```

Key files:
- `word/document.xml` - Main document contents
- `word/comments.xml` - Comments
- `word/media/` - Embedded images

See [ooxml.md](./ooxml.md) for detailed XML manipulation.
