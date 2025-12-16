# Batch Operations

Apply multiple edits efficiently from code or configuration files.

## Programmatic Batch Edits

Apply a list of edits at once:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

edits = [
    {
        "type": "insert_tracked",
        "text": " (revised)",
        "after": "Exhibit A"
    },
    {
        "type": "replace_tracked",
        "find": "30 days",
        "replace": "45 days"
    },
    {
        "type": "delete_tracked",
        "text": "unless otherwise agreed"
    }
]

results = doc.apply_edits(edits)

# Check results
for result in results:
    print(result)  # ✓ insert_tracked: Inserted ' (revised)' after 'Exhibit A'

doc.save("contract_edited.docx")
```

## YAML Configuration Files

Define edits in YAML for repeatable workflows:

```yaml
# edits.yaml
edits:
  # Text operations
  - type: insert_tracked
    text: " (as amended)"
    after: "Agreement dated"
    scope: "section:Recitals"

  - type: replace_tracked
    find: "Contractor"
    replace: "Service Provider"

  - type: delete_tracked
    text: "subject to approval"
    scope:
      contains: "termination"

  # Structural operations
  - type: insert_paragraph
    text: "Compliance"
    after: "Section 5"
    style: "Heading1"
    track: true

  - type: insert_paragraphs
    texts:
      - "All parties shall comply with applicable laws."
      - "This includes federal, state, and local regulations."
    after: "Compliance"
    track: true

  - type: delete_section
    heading: "Deprecated Clause"
    track: true

  # Regex operations
  - type: replace_tracked
    find: "(\d+) days"
    replace: "\1 business days"
    regex: true
```

Apply the YAML file:

```python
results = doc.apply_edit_file("edits.yaml")
print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
```

## JSON Configuration

JSON format is also supported:

```json
{
  "edits": [
    {
      "type": "replace_tracked",
      "find": "old term",
      "replace": "new term"
    }
  ]
}
```

```python
results = doc.apply_edit_file("edits.json")
```

## Error Handling in Batches

Control behavior when edits fail:

```python
# Stop on first error
results = doc.apply_edits(edits, stop_on_error=True)

# Continue on errors (default)
results = doc.apply_edits(edits, stop_on_error=False)

# Check individual results
for i, result in enumerate(results):
    if result.success:
        print(f"✓ Edit {i+1}: {result.message}")
    else:
        print(f"✗ Edit {i+1}: {result.message}")
        if result.error:
            print(f"  Error: {result.error}")
```

## Supported Edit Types

| Type | Description | Required Parameters |
|------|-------------|---------------------|
| `insert_tracked` | Insert text | `text`, `after` or `before` |
| `delete_tracked` | Delete text | `text` |
| `replace_tracked` | Replace text | `find`, `replace` |
| `insert_paragraph` | Add paragraph | `text`, `after` or `before` |
| `insert_paragraphs` | Add multiple paragraphs | `texts`, `after` or `before` |
| `delete_section` | Remove section | `heading` |

## Scopes in YAML

Limit operations to specific parts of the document:

```yaml
edits:
  # String scope - paragraphs containing text
  - type: replace_tracked
    find: "Client"
    replace: "Customer"
    scope: "payment terms"

  # Section scope
  - type: insert_tracked
    text: " (NDA)"
    after: "Agreement"
    scope: "section:Definitions"

  # Dictionary scope for complex filters
  - type: replace_tracked
    find: "old"
    replace: "new"
    scope:
      contains: "confidential"
      not_contains: "public"
```

## Next Steps

- [Basic Operations](basic-operations.md) — Individual operation details
- [Advanced Features](advanced.md) — Scopes, regex, MS365 identity
- [API Reference](../PROPOSED_API.md) — Complete method documentation
