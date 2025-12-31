# Editing Word Documents

Use **python-docx-redline** for all editing tasks. It handles run fragmentation that breaks python-docx find/replace, and supports both untracked (silent) and tracked editing modes.

## When to Use Each Mode

| Mode | Use Case | Parameter |
|------|----------|-----------|
| **Untracked** | Template population, internal updates, bulk changes | `track=False` (default) |
| **Tracked** | Legal redlines, contract negotiations, review workflows | `track=True` |

## Basic Editing (Untracked)

```python
from python_docx_redline import Document

doc = Document("existing.docx")

# Replace text - handles run boundaries automatically
doc.replace("Old Value", "New Value")
doc.replace("{{PLACEHOLDER}}", "Actual Value")

# Insert text after/before anchor
doc.insert(" Inc.", after="Acme Corp")
doc.insert("Dear ", before="Customer")

# Delete text
doc.delete("DRAFT - ")
doc.delete("Please remove this sentence.")

# Move text to new location
doc.move("Section A", after="Table of Contents")

doc.save("modified.docx")
```

## Tracked Editing

Add `track=True` to any operation to show it as a tracked change:

```python
doc = Document("contract.docx")

# Tracked operations - visible in Word's track changes view
doc.replace("30 days", "45 days", track=True)
doc.insert(" (revised)", after="Exhibit A", track=True)
doc.delete("unless otherwise agreed", track=True)
doc.move("Indemnity clause", after="Warranties", track=True)

# Or use explicit *_tracked methods (equivalent)
doc.replace_tracked("the Contractor", "the Service Provider")
doc.insert_tracked(" (amended)", after="Section 2.1")
doc.delete_tracked("subject to approval")
doc.move_tracked("Appendix A", after="Table of Contents")

doc.save("contract_redlined.docx")
```

## Why python-docx-redline Over python-docx

### The Run Fragmentation Problem

Word stores text in "runs" - segments with consistent formatting. A single word can be split across multiple runs unpredictably:

```
"Contract" might be stored as:
  Run 1: "Con"
  Run 2: "tract"
```

This breaks naive find/replace:

```python
# python-docx - FAILS when text spans runs
from docx import Document
doc = Document("file.docx")
for para in doc.paragraphs:
    if "Contract" in para.text:  # True - concatenated text
        para.text = para.text.replace("Contract", "Agreement")  # DESTROYS ALL FORMATTING
```

```python
# python-docx-redline - WORKS regardless of run boundaries
from python_docx_redline import Document
doc = Document("file.docx")
doc.replace("Contract", "Agreement")  # Handles run boundaries, preserves formatting
```

### Additional Benefits

- **Smart quote handling**: Curly quotes match straight quotes automatically
- **Regex support**: `doc.replace(r"(\d+) days", r"\1 business days", regex=True)`
- **Fuzzy matching**: `doc.replace("Contarct", "Agreement", fuzzy=0.9)` for OCR'd docs
- **Occurrence control**: Target specific occurrences with `occurrence=1` or `occurrence="all"`
- **Scoped edits**: `doc.replace("Client", "Customer", scope="section:Payment Terms")`

## Template Population

```python
from python_docx_redline import Document

doc = Document("template.docx")

# Simple placeholders
doc.replace("{{NAME}}", "John Doe")
doc.replace("{{DATE}}", "December 28, 2024")
doc.replace("{{COMPANY}}", "Acme Inc.")

# Multiple occurrences
doc.replace("{{SIGNATURE}}", "________________", occurrence="all")

doc.save("filled_template.docx")
```

## Finding Text Before Editing

```python
# Find all occurrences with context
matches = doc.find_all("payment")
for m in matches:
    print(f"{m.index}: ...{m.context_before}[{m.matched_text}]{m.context_after}...")

# Target specific occurrence
doc.replace("payment", "Payment", occurrence=2)  # Only 2nd occurrence

# Replace all occurrences
doc.replace("payment", "Payment", occurrence="all")
```

## Working with Headers and Footers

```python
# Headers and footers support both modes
doc.replace("DRAFT", "FINAL", scope="headers")
doc.replace("Page X of Y", "Page {{PAGE}}", scope="footers")

# With tracking
doc.replace("v1.0", "v2.0", scope="headers", track=True)
```

## Working with Tables

```python
# Edits in tables work automatically
doc.replace("{{ITEM_1}}", "Widget", scope="tables")

# Or scope to specific table content
doc.replace("TBD", "Confirmed", scope="section:Pricing Table")
```

## Batch Operations

### From Python

```python
edits = [
    {"type": "replace", "find": "{{NAME}}", "replace": "John Doe"},
    {"type": "replace", "find": "{{DATE}}", "replace": "2024-12-28"},
    {"type": "delete", "text": "DRAFT - "},
    {"type": "insert", "text": " (final)", "after": "Agreement"},
]
results = doc.apply_edits(edits)
print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
```

### From YAML File

```yaml
# edits.yaml
default_track: false  # Default for all edits

edits:
  - type: replace
    find: "{{COMPANY}}"
    replace: "Acme Inc."

  - type: replace
    find: "30 days"
    replace: "45 days"
    track: true  # Override: this one is tracked

  - type: delete
    text: "DRAFT"
```

```python
doc.apply_edit_file("edits.yaml")

# Or override the file's default_track
doc.apply_edit_file("edits.yaml", default_track=True)  # Track all by default
```

## Advanced Features

### Regex with Capture Groups

```python
# Update dates
doc.replace(r"(\d{1,2})/(\d{1,2})/(\d{4})", r"\3-\1-\2", regex=True)

# Standardize formatting
doc.replace(r"Section (\d+)", r"Article \1", regex=True, occurrence="all")
```

### Fuzzy Matching

For OCR'd or inconsistently formatted documents:

```python
# Match "Contarct" or "Contract" or "CONTACT" with 85% similarity
doc.replace("Contract", "Agreement", fuzzy=0.85)

# Full fuzzy config
doc.replace("Agreement", "Contract", fuzzy={
    "threshold": 0.9,
    "algorithm": "ratio",  # or "partial_ratio", "token_sort_ratio"
    "normalize_whitespace": True
})
```

### Context Preview

```python
doc.replace("term", "period", show_context=True, context_chars=50)
# Shows: Replacing "term" in context: "...the initial term of this..."
```

## API Reference

### Core Methods

| Method | Description | Default |
|--------|-------------|---------|
| `replace(find, replace_with, track=False)` | Find and replace text | Untracked |
| `insert(text, after=..., track=False)` | Insert after anchor | Untracked |
| `delete(text, track=False)` | Delete text | Untracked |
| `move(text, after=..., track=False)` | Move text to new location | Untracked |

### Tracked Aliases

| Method | Equivalent |
|--------|------------|
| `replace_tracked(find, replace)` | `replace(find, replace, track=True)` |
| `insert_tracked(text, after=...)` | `insert(text, after=..., track=True)` |
| `delete_tracked(text)` | `delete(text, track=True)` |
| `move_tracked(text, after=...)` | `move(text, after=..., track=True)` |

### Common Parameters

| Parameter | Description |
|-----------|-------------|
| `track` | `True` for tracked changes, `False` for silent edit |
| `occurrence` | `1`, `2`, `"first"`, `"last"`, `"all"`, or `[1,3,5]` |
| `scope` | `"section:Name"`, `"headers"`, `"footers"`, `"tables"` |
| `regex` | Treat pattern as regular expression |
| `fuzzy` | `0.9` or `{"threshold": 0.9, ...}` for fuzzy matching |
| `author` | Author name for tracked changes |

## Section Operations

Delete entire sections (heading + all content until the next heading) with a single call.

### delete_section()

```python
from python_docx_redline import Document

doc = Document("report.docx")

# Delete section with tracked changes (default)
doc.delete_section("Methods", track=True)

# Delete section silently (no tracked changes)
doc.delete_section("Appendix B", track=False)

# With custom author for tracked changes
doc.delete_section("Old Section", author="Reviewer")

# Scoped deletion (only search within specific area)
doc.delete_section("Subsection 1.2", scope="section:Chapter 1")

doc.save("cleaned_report.docx")
```

**Signature:**
```python
def delete_section(
    heading: str,
    track: bool = True,
    update_toc: bool = False,
    author: str | None = None,
    scope: str | dict | Any | None = None,
) -> Section
```

**Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `heading` | `str` | required | Text of the heading to find and delete |
| `track` | `bool` | `True` | Delete as tracked change (visible in Word) |
| `update_toc` | `bool` | `False` | No-op, kept for API compatibility. TOC updates require Word. |
| `author` | `str \| None` | `None` | Author name for tracked changes (uses doc default if None) |
| `scope` | `str \| dict \| None` | `None` | Limit search to specific area |

**Returns:** `Section` object representing the deleted section

**Raises:**
- `TextNotFoundError`: If the heading is not found
- `AmbiguousTextError`: If multiple sections match the heading text

### How Section Boundaries Work

A section is defined as:
1. A heading paragraph (using Word heading styles like Heading 1, Heading 2, etc.)
2. All following paragraphs until the next heading paragraph

When you delete a section, **all** content from the heading through the last paragraph before the next heading is removed.

```
Document structure:
  # Introduction        <-- Heading 1
    Paragraph 1
    Paragraph 2
  # Methods             <-- Heading 1 (section boundary)
    Paragraph 3
    Paragraph 4
  # Results             <-- Heading 1 (section boundary)
    Paragraph 5

doc.delete_section("Methods")
# Deletes: "Methods" heading + Paragraphs 3 & 4
# "Results" section is preserved
```

### Common Use Cases

**Removing an outdated appendix:**
```python
doc.delete_section("Appendix A: Legacy Data", track=False)
```

**Redlining a section for review:**
```python
doc.delete_section("Disputed Terms", track=True, author="Legal Review")
# Section shows as deleted in Word's track changes view
```

**Removing multiple sections:**
```python
sections_to_remove = ["Draft Notes", "Internal Comments", "TODO"]
for section_name in sections_to_remove:
    try:
        doc.delete_section(section_name, track=False)
    except TextNotFoundError:
        pass  # Section doesn't exist, skip
```

**Targeted deletion within a chapter:**
```python
# Only delete "Summary" section inside Chapter 2
doc.delete_section("Summary", scope="section:Chapter 2")
```

### Handling Ambiguous Headings

If the same heading text appears multiple times, use scope to disambiguate:

```python
# Fails if "Summary" appears in multiple chapters
doc.delete_section("Summary")  # Raises AmbiguousTextError

# Use scope to target specific occurrence
doc.delete_section("Summary", scope="section:Chapter 1")
```

## When to Use Raw python-docx

python-docx is still appropriate for:

- **Creating new documents from scratch** (no existing content to edit)
- **Adding new content** at end of document (no find/replace needed)
- **Simple property access** like `doc.core_properties.author`

For everything else, use python-docx-redline.
