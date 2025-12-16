# Basic Operations

The three core operations for editing documents with tracked changes.

## Insert Text

Insert text after (or before) a specific location:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Insert after specific text
doc.insert_tracked(" (as amended)", after="Section 2.1")

# Insert before specific text
doc.insert_tracked("IMPORTANT: ", before="Payment Terms")

doc.save("contract_edited.docx")
```

## Replace Text

Replace existing text with new text:

```python
doc.replace_tracked(
    find="the Contractor",
    replace="the Service Provider"
)

# Replace with context preview
doc.replace_tracked(
    find="30 days",
    replace="45 days",
    show_context=True,      # Shows surrounding text
    context_chars=100       # Characters to display
)
```

## Delete Text

Mark text for deletion:

```python
doc.delete_tracked("unless otherwise agreed")
```

## Regex Operations

Use regular expressions for pattern matching:

```python
# Replace all dollar amounts
doc.replace_tracked(r"\$[\d,]+\.?\d*", "$XXX.XX", regex=True)

# Use capture groups - change "X days" to "X business days"
doc.replace_tracked(r"(\d+) days", r"\1 business days", regex=True)

# Swap date format from MM/DD/YYYY to DD/MM/YYYY
doc.replace_tracked(
    r"(\d{2})/(\d{2})/(\d{4})",
    r"\2/\1/\3",
    regex=True
)

# Insert after any section reference
doc.insert_tracked(" (as amended)", after=r"Section \d+\.\d+", regex=True)
```

## Error Handling

The library provides helpful error messages:

```python
from python_docx_redline import Document, TextNotFoundError

doc = Document("contract.docx")

try:
    doc.insert_tracked("new text", after="nonexistent text")
except TextNotFoundError as e:
    print(e)
    # Output:
    # Could not find 'nonexistent text'
    #
    # Suggestions:
    #   • Check for typos in the search text
    #   • Try searching for a shorter or more unique phrase
    #   • Verify the text exists in the document
```

Common issues detected automatically:

- Curly quotes vs straight quotes
- Double spaces
- Leading/trailing whitespace
- Case sensitivity mismatches
- Special characters (non-breaking spaces, zero-width spaces)

## Minimal Editing Mode

For legal documents, use minimal mode to avoid spurious formatting changes:

```python
doc.replace_tracked(
    find="the Contractor",
    replace="the Service Provider",
    minimal=True  # Clean tracked changes in Word's review pane
)
```

## Next Steps

- [Structural Operations](structural-ops.md) — Insert paragraphs, delete sections
- [Viewing Content](viewing-content.md) — Read documents before editing
- [Advanced Features](advanced.md) — Scopes, MS365 identity, rendering
