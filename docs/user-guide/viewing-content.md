# Viewing Content

Read and understand document structure before making edits.

## Reading Paragraphs

Access all paragraphs with metadata:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

for para in doc.paragraphs:
    if para.is_heading():
        print(f"Section: {para.text}")
    else:
        print(f"  {para.text}")
```

## Parsing Sections

Documents are automatically parsed by heading structure:

```python
for section in doc.sections:
    if section.heading:
        print(f"\nSection: {section.heading_text} (Level {section.heading_level})")
        print(f"  {len(section.paragraphs)} paragraphs")

    # Search within a section
    if section.contains("confidential"):
        print("  Contains confidential information")
```

## Extracting Full Text

Get the complete document text:

```python
text = doc.get_text()

if "arbitration" in text.lower():
    print("Document contains arbitration clause")
```

## Finding All Occurrences

Use `find_all()` to locate text with surrounding context:

```python
matches = doc.find_all("confidential", context_chars=50)

for match in matches:
    print(f"Found: ...{match.context_before}[{match.matched_text}]{match.context_after}...")
```

### Find with Occurrence Selection

Select a specific occurrence when there are multiple matches:

```python
# Find the 2nd occurrence
matches = doc.find_all("Agreement", occurrence=2)

# Find the last occurrence
matches = doc.find_all("Agreement", occurrence=-1)
```

### Fuzzy Matching

Find text even with minor differences (requires `rapidfuzz`):

```python
# Match with 85% similarity threshold
matches = doc.find_all("Sectoin 2.1", fuzzy=0.85)  # Finds "Section 2.1"
```

## Agent Workflow Pattern

Read first, then make targeted edits:

```python
doc = Document("contract.docx")

# Step 1: Understand the document
for section in doc.sections:
    print(f"{section.heading_text}: {len(section.paragraphs)} paragraphs")

# Step 2: Find what needs changing
matches = doc.find_all("net 30 days", context_chars=50)
for m in matches:
    print(f"Found in context: ...{m.context_before}[{m.matched_text}]{m.context_after}...")

# Step 3: Make targeted edits
for section in doc.sections:
    if section.heading_text == "Payment Terms":
        if section.contains("net 30 days"):
            doc.replace_tracked(
                "net 30 days",
                "net 45 days",
                scope="section:Payment Terms"
            )

doc.save("contract_updated.docx")
```

## Context-Aware Editing

Preview context and detect potential issues:

```python
import warnings
from python_docx_redline import Document, ContinuityWarning

doc = Document("contract.docx")
warnings.simplefilter("always")

# Show context before making changes
doc.replace_tracked(
    find="old text",
    replace="new text",
    show_context=True,       # Print surrounding text
    context_chars=100,       # Characters to show
    check_continuity=True    # Warn about sentence fragments
)
```

## Next Steps

- [Basic Operations](basic-operations.md) — Make edits after understanding content
- [Batch Operations](batch-operations.md) — Define edits in YAML files
- [find_all() API](../API_FIND_ALL.md) — Complete find_all() documentation
