# Tracked Changes

## Basic Operations

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Insert text
doc.insert_tracked(" (as amended)", after="Section 2.1")

# Replace text
doc.replace_tracked("30 days", "45 days")

# Delete text
doc.delete_tracked("subject to approval")

doc.save("contract_redlined.docx")
```

## Paragraph-Level Deletion

Use `delete_paragraph_tracked()` to remove entire paragraphs cleanly. Unlike `delete_tracked()` which only marks text as deleted (leaving empty paragraphs behind), this removes the paragraph element entirely:

```python
# Delete paragraph containing specific text
doc.delete_paragraph_tracked(containing="Some citation text")

# Delete by index (0-based)
doc.delete_paragraph_tracked(paragraph_index=5)

# Delete paragraph object directly
para = doc.paragraphs[5]
doc.delete_paragraph_tracked(paragraph=para)

# Keep for review (strikethrough only, empty para remains after accept)
doc.delete_paragraph_tracked(containing="text", remove_element=False)
```

This is especially useful for:
- Removing bullet points from lists
- Deleting citations from claim charts
- Cleaning up table cell content without leaving empty lines

## Handling Multiple Occurrences

When text appears multiple times, use the `occurrence` parameter:

```python
# Target specific occurrence (1-indexed)
doc.replace_tracked("Section", "Article", occurrence=1)      # First
doc.replace_tracked("Section", "Article", occurrence=2)      # Second
doc.replace_tracked("Section", "Article", occurrence="last") # Last
doc.replace_tracked("Section", "Article", occurrence="all")  # All

# Target multiple specific occurrences
doc.replace_tracked("Section", "Article", occurrence=[1, 3, 5])
```

## Special Character Normalization

Word documents use typographic characters (curly quotes, special bullets, en/em dashes). The library normalizes these automatically:

```python
# Curly quotes → straight quotes (automatic)
doc.replace_tracked("Defendant's motion", "party's motion")  # Works with curly apostrophe
doc.replace_tracked('"free trial"', '"subscription"')        # Works with curly double quotes

# Bullets (•, ·, ◦, ▪, etc.) → standard bullet
doc.delete_tracked("• First item")  # Matches any bullet variant

# Dashes (–, —) → hyphen
doc.replace_tracked("2020-2024", "2020-2025")  # Matches en/em dashes too

# Disable if you need exact matching
doc.replace_tracked("exact's", "match", normalize_special_chars=False)
```

## Regex Support

```python
# Update all "X days" to "X business days"
doc.replace_tracked(r"(\d+) days", r"\1 business days", regex=True)

# Redact dollar amounts
doc.replace_tracked(r"\$[\d,]+\.?\d*", "$XXX.XX", regex=True)

# Swap date format MM/DD/YYYY to DD/MM/YYYY
doc.replace_tracked(r"(\d{2})/(\d{2})/(\d{4})", r"\2/\1/\3", regex=True)
```

## Scoped Edits

### String Shortcuts

```python
# Only in specific section
doc.replace_tracked("Client", "Customer", scope="section:Payment Terms")

# Only paragraphs containing text
doc.replace_tracked("30 days", "45 days", scope="paragraph_containing:payment")
```

### Dictionary Format

```python
doc.replace_tracked(
    "Client", "Customer",
    scope={
        "contains": "payment",      # Must contain this
        "not_contains": "excluded", # Must NOT contain this
        "section": "Terms"          # Under this heading
    }
)
```

### Callable Scope

```python
def my_filter(paragraph):
    text = "".join(paragraph.itertext())
    return "important" in text.lower() and len(text) > 100

doc.replace_tracked("old", "new", scope=my_filter)
```

### Debugging Scope Issues

```python
from python_docx_redline.scope import ScopeEvaluator
from python_docx_redline.constants import WORD_NAMESPACE

all_paragraphs = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
debug_info = ScopeEvaluator.debug_scope(all_paragraphs, "section:Payment Terms")

print(f"Matched {len(debug_info['matched'])} paragraphs")
print(f"Excluded {len(debug_info['excluded'])} paragraphs")
```

## Batch Operations

### From Python List

```python
edits = [
    {"type": "replace_tracked", "find": "net 30 days", "replace": "net 45 days"},
    {"type": "replace_tracked", "find": "Contractor", "replace": "Service Provider"},
    {"type": "insert_tracked", "text": " (as amended)", "after": "Agreement dated"},
    {"type": "delete_tracked", "text": "subject to board approval"},
]

results = doc.apply_edits(edits, stop_on_error=False)

for i, result in enumerate(results):
    status = "OK" if result.success else "FAILED"
    print(f"Edit {i}: {status} - {result.message}")
```

### From YAML File

```yaml
# edits.yaml
edits:
  - type: replace_tracked
    find: "net 30 days"
    replace: "net 45 days"

  - type: insert_tracked
    text: " (as amended)"
    after: "Agreement dated"
    scope: "section:Introduction"

  - type: replace_tracked
    find: "(\d+) days"
    replace: "\\1 business days"
    regex: true
```

```python
results = doc.apply_edit_file("edits.yaml", stop_on_error=False)
```

## Text Formatting with Tracking

```python
doc.format_tracked(
    "IMPORTANT",
    bold=True,
    color="#FF0000",
    occurrence="all"
)

doc.format_paragraph_tracked(
    containing="Section 1",
    alignment="center",
    spacing_after=12.0
)
```

## Image Insertion

```python
# Basic
doc.insert_image("logo.png", after="Company Name:")

# With dimensions
doc.insert_image("chart.png", after="Figure 1:", width_inches=4.0, height_inches=3.0)

# Tracked (shows in review pane)
doc.insert_image_tracked("signature.png", after="Authorized By:", author="Legal Team")
```

## Accept/Reject Changes

```python
# Accept all
result = doc.accept_all_changes()
print(f"Accepted {result.insertions} insertions, {result.deletions} deletions")

# Reject all
result = doc.reject_all_changes()
```

## MS365 Identity Integration

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
# Changes appear with full user profile in Word
```

## Integration with python-docx

```python
from docx import Document as PythonDocxDocument
from python_docx_redline import from_python_docx

# Create with python-docx
py_doc = PythonDocxDocument()
py_doc.add_heading("Contract", 0)
py_doc.add_paragraph("Payment terms: net 30 days")

# Convert for tracked edits
doc = from_python_docx(py_doc, author="Legal Team")
doc.replace_tracked("net 30 days", "net 45 days")
doc.save("contract_redlined.docx")
```

## In-Memory Workflows

```python
# Load from bytes
with open("contract.docx", "rb") as f:
    doc = Document(f.read())

doc.replace_tracked("old", "new")

# Get bytes
doc_bytes = doc.save_to_bytes()
```

## Rendering to Images

```python
from python_docx_redline.rendering import render_document_to_images, is_rendering_available

if is_rendering_available():
    images = render_document_to_images("contract.docx", dpi=150)
    # Returns: page-1.png, page-2.png, etc.
```
