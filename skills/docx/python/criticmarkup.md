# CriticMarkup Workflow

CriticMarkup enables a markdown-based editing workflow for Word documents. Export tracked changes to plain text, edit in any text editor, and import changes back as tracked changes.

## CriticMarkup Syntax Reference

| Syntax | Meaning | DOCX Equivalent |
|--------|---------|-----------------|
| `{++inserted text++}` | Insertion | `<w:ins>` |
| `{--deleted text--}` | Deletion | `<w:del>` |
| `{~~old~>new~~}` | Substitution | `<w:del>` + `<w:ins>` |
| `{>>comment<<}` | Comment | `<w:comment>` |
| `{==marked text==}` | Highlight | Comment range |
| `{==text=={>>comment<<}}` | Highlight + Comment | Comment on text |

## Export: DOCX to CriticMarkup

Convert tracked changes in a Word document to CriticMarkup markdown.

```python
from python_docx_redline import Document

doc = Document("contract_with_changes.docx")

# Export to markdown
markdown = doc.to_criticmarkup()
print(markdown)
# Output:
# The parties agree to {--30--}{++45++} day payment terms.
# {==Section 2.1=={>>Legal review needed<<}}

# Save for editing
with open("contract_review.md", "w") as f:
    f.write(markdown)
```

### Export Options

```python
# Include comments (default: True)
markdown = doc.to_criticmarkup(include_comments=True)

# Export without comments
markdown = doc.to_criticmarkup(include_comments=False)
```

## Import: CriticMarkup to DOCX

Apply CriticMarkup changes back to a document as tracked changes.

```python
from python_docx_redline import Document

doc = Document("contract_original.docx")

# Apply CriticMarkup from edited markdown
markup = """
Payment is due in {--30--}{++45++} days.
{++This clause was added.++}
"""

result = doc.apply_criticmarkup(markup, author="Review Bot")
print(f"Applied {result.successful} of {result.total} changes")

doc.save("contract_updated.docx")
```

### ApplyResult Object

```python
result = doc.apply_criticmarkup(markup)

result.total        # Total operations attempted
result.successful   # Operations that succeeded
result.failed       # Operations that failed
result.success_rate # Percentage (0.0 - 100.0)
result.errors       # List of (CriticOperation, error_message) tuples
```

### Error Handling

```python
# Stop on first error (default: False)
result = doc.apply_criticmarkup(markup, stop_on_error=True)

# Continue processing all operations
result = doc.apply_criticmarkup(markup, stop_on_error=False)

# Check for errors
if result.failed > 0:
    for op, error_msg in result.errors:
        print(f"Failed: {op.type.value} - {error_msg}")
```

## Round-Trip Workflow

Complete workflow: export, edit, import.

```python
from python_docx_redline import Document

# 1. Export existing tracked changes
doc = Document("contract.docx")
markdown = doc.to_criticmarkup()

# Save for editing
with open("contract_review.md", "w") as f:
    f.write(markdown)

# 2. User edits markdown file in any text editor
# They can add {++insertions++}, {--deletions--}, etc.

# 3. Import edited markdown back
with open("contract_review.md") as f:
    edited_markdown = f.read()

# Re-load original document
doc = Document("contract_original.docx")
result = doc.apply_criticmarkup(edited_markdown, author="Reviewer")
doc.save("contract_final.docx")

# All CriticMarkup changes appear as tracked changes in Word
```

## Parser Functions

For advanced use cases, work directly with parsed operations.

### Parse CriticMarkup

```python
from python_docx_redline.criticmarkup import parse_criticmarkup

text = "Hello {++world++}! This is {--old--}{++new++}."
operations = parse_criticmarkup(text)

for op in operations:
    print(f"{op.type.value}: '{op.text}'")
    print(f"  Position: {op.position}-{op.end_position}")
    print(f"  Context: ...{op.context_before[-20:]}|{op.context_after[:20]}...")
```

### Operation Types

```python
from python_docx_redline.criticmarkup import OperationType

OperationType.INSERTION     # {++text++}
OperationType.DELETION      # {--text--}
OperationType.SUBSTITUTION  # {~~old~>new~~}
OperationType.COMMENT       # {>>comment<<}
OperationType.HIGHLIGHT     # {==text==}
```

### Strip CriticMarkup

Remove markup and resolve to final text.

```python
from python_docx_redline.criticmarkup import strip_criticmarkup

# Insertions: keep inserted text
strip_criticmarkup("Hello {++world++}!")  # "Hello world!"

# Deletions: remove deleted text
strip_criticmarkup("Say {--goodbye--}hello")  # "Say hello"

# Substitutions: keep new text
strip_criticmarkup("{~~old~>new~~}")  # "new"

# Comments: remove entirely
strip_criticmarkup("Text {>>note<<} here")  # "Text  here"
```

### Render Operations

Convert parsed operations back to CriticMarkup syntax.

```python
from python_docx_redline.criticmarkup import (
    parse_criticmarkup,
    render_criticmarkup,
    CriticOperation,
    OperationType,
)

# Create operations programmatically
operations = [
    CriticOperation(type=OperationType.INSERTION, text="world", position=6),
]

# Render to markup
result = render_criticmarkup(operations, "Hello !")
# "Hello {++world++}!"
```

## Common Use Cases

### AI Agent Review

Export document for AI review, then import suggestions.

```python
doc = Document("draft.docx")
markdown = doc.to_criticmarkup()

# Send to AI for review (returns CriticMarkup)
reviewed_markdown = ai_review(markdown)

# Apply AI suggestions as tracked changes
doc_copy = Document("draft.docx")
result = doc_copy.apply_criticmarkup(reviewed_markdown, author="AI Reviewer")
doc_copy.save("draft_reviewed.docx")
```

### Version Control

Use CriticMarkup for git-friendly document changes.

```python
# Export to markdown for version control
doc = Document("contract.docx")
with open("contract.md", "w") as f:
    f.write(doc.to_criticmarkup())

# Now changes can be tracked in git
# git add contract.md
# git commit -m "Version 2.1 changes"
```

### Batch Processing

Process multiple documents with the same changes.

```python
from pathlib import Path

markup = """
Replace {~~Contractor~>Service Provider~~} throughout.
Delete {--outdated clause--}.
"""

for docx_file in Path("contracts/").glob("*.docx"):
    doc = Document(str(docx_file))
    result = doc.apply_criticmarkup(markup, author="Batch Update")
    if result.successful > 0:
        doc.save(str(docx_file))
        print(f"{docx_file.name}: {result.successful} changes applied")
```

## Limitations

- **Insertion anchoring**: Insertions require context to determine placement. The library uses the text before the insertion point as an anchor.
- **Comments**: Standalone comments `{>>text<<}` need surrounding context to attach. Use `{==marked=={>>comment<<}}` format for reliable comment placement.
- **Formatting**: CriticMarkup doesn't preserve formatting (bold, italic). Use direct API calls for formatting changes.
- **Tables**: Tracked changes in tables are exported but complex table structures may not round-trip perfectly.

## Reference

- [CriticMarkup Specification](http://criticmarkup.com/)
- [python-docx-redline Documentation](https://github.com/SanctionedCodeList/python-docx-redline)
