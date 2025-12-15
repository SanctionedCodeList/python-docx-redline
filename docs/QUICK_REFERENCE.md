# Quick Reference Guide

**For**: Developers using the proposed high-level API
**See also**: PROPOSED_API.md (full specification), IMPLEMENTATION_NOTES.md (technical details)

---

## Basic Usage

```python
from python_docx_redline import Document

# Open document
doc = Document('file.docx', author="Your Name")

# Make edits
doc.insert_tracked(" text", after="target")
doc.replace_tracked("old", "new")
doc.delete_tracked("unwanted text")

# Save
doc.save('output.docx')
```

---

## Common Operations

### Insert Text

```python
# Simple
doc.insert_tracked(" added", after="target")

# With scope
doc.insert_tracked(" added", after="target", scope="section:Argument")

# Before instead of after
doc.insert_tracked("prefix ", before="target")
```

### Replace Text

```python
# Simple replacement
doc.replace_tracked("old", "new")

# With scope
doc.replace_tracked("old", "new", scope="paragraph_containing:specific context")

# Replace only first occurrence
doc.replace_tracked("old", "new", first_only=True)
```

### Delete Text

```python
doc.delete_tracked("text to remove")
doc.delete_tracked("text", scope="section:Introduction")
```

### Insert Paragraphs

```python
# Single paragraph
doc.insert_paragraph(
    "New paragraph text",
    after="existing paragraph ending",
    track=True
)

# Multiple paragraphs
doc.insert_paragraphs(
    ["Para 1", "Para 2", "Para 3"],
    after="target location",
    track=True
)
```

### Delete Section

```python
doc.delete_section(
    "Section Heading",
    track=True,
    update_toc=True
)
```

### Insert Images

```python
# Basic image insertion
doc.insert_image("logo.png", after="Company Name:")

# With dimensions
doc.insert_image("chart.png", after="Figure 1:", width_inches=4.0)

# Tracked insertion (appears in Word's review pane)
doc.insert_image_tracked("signature.png", after="Authorized By:")
```

---

## Scope System

### String Shortcuts

```python
# Paragraph containing text
scope="target text"

# Specific section
scope="section:Argument"

# Under heading
scope="heading:Background"
```

### Dictionary Format

```python
scope={
    "contains": "specific text",
    "section": "Argument",
    "not_in": ["Table of Authorities", "Table of Contents"]
}
```

### Callable

```python
def my_scope(para):
    return "target" in para.text and para.style == "Normal"

doc.insert_tracked("text", after="target", scope=my_scope)
```

---

## Search and Disambiguation

### Find Text

```python
# Find single occurrence
span = doc.find_text("target", scope="section:Argument")
span.insert_after(" added")

# Find all occurrences
results = doc.find_all("target")
for i, r in enumerate(results):
    print(f"{i}: {r.context}")

# Use specific result
results[2].insert_after(" added")
```

### Handle Ambiguity

```python
try:
    doc.insert_tracked(" added", after="common text")
except AmbiguousTextError as e:
    print(e)  # Shows all matches with context
    # Add better scope
    doc.insert_tracked(" added", after="common text", scope="section:Argument")
```

---

## Batch Operations

### From List

```python
edits = [
    {
        "type": "insert_tracked",
        "text": " (interpreting IRPA)",
        "after": "(7th Cir. 2022)",
        "scope": "Huston"
    },
    {
        "type": "replace_tracked",
        "find": "old",
        "replace": "new"
    }
]

results = doc.apply_edits(edits)
print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
```

### From YAML File

```python
from python_docx_redline import apply_edit_file

results = apply_edit_file('edits.yaml')
```

---

## Document Management

### Accept/Reject Changes

```python
# Accept all
result = doc.accept_all_changes()
print(f"Accepted {result.insertions} insertions")

# Reject all
doc.reject_all_changes()

# Delete comments
doc.delete_all_comments()
```

### Comments

```python
# Add comment on text
comment = doc.add_comment("Please review", on="target text")

# Add comment on tracked insertion (text inside w:ins)
doc.insert_tracked("new clause", after="Section 1")
doc.add_comment("Review this addition", on="new clause")

# Add comment on tracked deletion (text inside w:del)
doc.delete_tracked("old term")
doc.add_comment("Why removed?", on="old term")

# Access comments
for c in doc.comments:
    print(f"{c.author}: {c.text} on '{c.marked_text}'")
```

### Validation

```python
# Validate before saving
result = doc.validate()
if result.is_valid:
    doc.save()
else:
    print(result.errors)

# Auto-validate on save (default)
doc.save('output.docx')  # Validates automatically

# Skip validation (faster but risky)
doc.save('output.docx', validate=False)
```

---

## Error Handling

### Text Not Found

```python
from python_docx_redline.errors import TextNotFoundError

try:
    doc.insert_tracked(" text", after="nonexistent")
except TextNotFoundError as e:
    print(e.suggestions)  # Helpful tips
```

### Multiple Matches

```python
from python_docx_redline.errors import AmbiguousTextError

try:
    doc.insert_tracked(" text", after="common")
except AmbiguousTextError as e:
    # Show all matches
    for match in e.matches:
        print(match.context)
```

---

## TextSpan Operations

When you use `find_text()`, you get a TextSpan object:

```python
span = doc.find_text("target")

# Properties
span.text          # The matched text
span.context       # Surrounding text
span.location      # Section, paragraph, line number

# Operations
span.insert_before("prefix ")
span.insert_after(" suffix")
span.replace("new text")
span.delete()
```

---

## Advanced Features

### Custom Author

```python
# Set default author
doc = Document('file.docx', author="Smith, John")

# Override per-edit
doc.insert_tracked(" text", after="target", author="Different Author")
```

### Fuzzy Matching

```python
# Enable fuzzy text matching
doc.insert_tracked(
    " text",
    after="targat",  # Typo
    fuzzy=True
)
```

### Preview Changes

```python
# Preview before applying
span = doc.find_text("target")
print(f"Will insert after: {span.context}")
span.insert_after(" text")
```

---

## YAML Edit File Format

```yaml
document: input.docx
author: Your Name
output: output.docx

preprocessing:
  - accept_all_changes
  - delete_all_comments

edits:
  - type: insert_tracked
    text: " added text"
    after: "target"
    scope: "section:Argument"

  - type: replace_tracked
    find: "old"
    replace: "new"

  - type: delete_section
    heading: "Section to Remove"
    track: true
```

---

## Performance Tips

```python
# SLOW: Individual operations
for edit in edits:
    doc.insert_tracked(edit['text'], after=edit['after'])

# FAST: Batch processing
doc.apply_edits(edits)

# SLOW: Multiple saves
doc.insert_tracked(...)
doc.save('temp1.docx')
doc.insert_tracked(...)
doc.save('temp2.docx')

# FAST: Single save
doc.insert_tracked(...)
doc.insert_tracked(...)
doc.save('final.docx')
```

---

## Debugging

### Enable Logging

```python
import logging
logging.basicConfig(level=logging.DEBUG)

doc = Document('file.docx')
# Will log all operations
```

### Inspect TextSpan

```python
span = doc.find_text("target")
print(f"Text: {span.text}")
print(f"Context: {span.context}")
print(f"Location: {span.location.section}, line {span.location.line_number}")
```

### List All Tracked Changes

```python
changes = doc.list_tracked_changes()
for c in changes:
    print(f"{c.author}: {c.text[:50]}")
```

---

## Migration from Low-Level API

### Before (Low-Level)

```python
from scripts.document import Document
from scripts.utilities import _parse_tag

doc = Document("unpacked", author="Name", rsid="ABC123")
editor = doc["word/document.xml"]

para = editor.get_node(tag="w:p", line_number=1234)
runs = list(para.iter(_parse_tag("w:r")))
for run in runs:
    if "target" in ''.join(run.itertext()):
        editor.insert_after(run, '<w:ins>...</w:ins>')
        break

doc.save()
```

### After (High-Level)

```python
from python_docx_redline import Document

doc = Document('file.docx', author="Name")
doc.insert_tracked(" text", after="target")
doc.save('output.docx')
```

---

## Complete Example

```python
from python_docx_redline import Document

# Open document
doc = Document('motion.docx', author="Hancock, Parker")

# Preprocessing
doc.accept_all_changes()
doc.delete_all_comments()

# Add procedural parentheticals
citations = [
    ("(7th Cir. 2022)", " (interpreting IRPA)", "Huston"),
    ("(N.D. Ill. 2016)", " (granting motion to dismiss)", "Vrdolyak"),
    ("(N.D. Cal. Mar. 1, 2021)", " (dismissing with prejudice)", "Callahan")
]

for cite, parens, scope in citations:
    doc.insert_tracked(parens, after=cite, scope=scope)

# Replace text
doc.replace_tracked(
    "records their property ownership",
    "compiles their property ownership data",
    scope="It claims merely"
)

# Add new section
doc.insert_paragraph(
    "This case is about property records, not property owners...",
    after="mismatch between the allegations and the law",
    track=True
)

# Delete section
doc.delete_section("Enhanced Damages Fail as a Matter of Law", track=True)

# Save
doc.save('motion_final.docx')

print("Done! Applied 7 surgical edits.")
```

---

## Document Rendering

### Render to Images

```python
from python_docx_redline import Document
from python_docx_redline.rendering import is_rendering_available

# Check availability
if is_rendering_available():
    doc = Document("contract.docx")
    images = doc.render_to_images(output_dir="./images", dpi=150)
    for img in images:
        print(f"Page: {img}")
```

### Standalone Rendering

```python
from python_docx_redline.rendering import render_document_to_images

# Render any DOCX directly
images = render_document_to_images("contract.docx", dpi=200)
```

### Requirements

```bash
# macOS
brew install --cask libreoffice && brew install poppler

# Linux
sudo apt install libreoffice poppler-utils
```

---

## Getting Help

- **Full API**: See `PROPOSED_API.md`
- **Implementation**: See `IMPLEMENTATION_NOTES.md`
- **Examples**: See `examples/` directory
- **Errors**: Check exception message for suggestions
- **Debugging**: Enable logging with `logging.basicConfig(level=logging.DEBUG)`

---

**Quick Reference v1.1** • December 2025
