# find_all() API Documentation

## Overview

The `find_all()` method searches for text in a Word document and returns all matches with rich location metadata. This allows you to preview what text will be matched before performing operations like `replace_tracked()`, `insert_tracked()`, or `delete_tracked()`.

## Method Signature

```python
def find_all(
    self,
    text: str,
    regex: bool = False,
    case_sensitive: bool = True,
    scope: str | dict | Any | None = None,
    context_chars: int = 40,
) -> list[Match]:
```

## Parameters

- **text** (str): The text or regex pattern to search for
- **regex** (bool): Whether to treat `text` as a regex pattern (default: False)
- **case_sensitive** (bool): Whether to perform case-sensitive search (default: True)
- **scope** (str | dict | None): Limit search scope
  - `None`: Search entire document (default)
  - String: Search specific location (e.g., `"body"`, `"tables"`)
  - Dict: Complex scope criteria (e.g., `{"contains": "text"}`)
- **context_chars** (int): Number of characters to show before/after match in context (default: 40)

## Return Value

Returns a `list[Match]` where each `Match` object has:

- **index** (int): Zero-based index of this match in results
- **text** (str): The matched text
- **context** (str): Surrounding text for disambiguation (with ellipsis if truncated)
- **paragraph_index** (int): Zero-based index of the paragraph containing this match
- **paragraph_text** (str): Full text of the paragraph
- **location** (str): Human-readable location string
  - `"body"`: Main document body
  - `"table:0:row:2:cell:1"`: Table location with indices
  - `"header"`, `"footer"`: Headers and footers
  - `"footnote"`, `"endnote"`: Notes
- **span** (TextSpan): The underlying TextSpan object for advanced use

## Examples

### Basic Search

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Find all occurrences
matches = doc.find_all("production products")

print(f"Found {len(matches)} occurrence(s)")
for match in matches:
    print(f"[{match.index}] {match.location}: {match.context}")
```

Output:
```
Found 2 occurrence(s)
[0] body: ...Therefore, production products utilizing...
[1] table:0:row:45:cell:1: ...Therefore, production products utilizing Adeia's...
```

### Preview Before Replace

```python
# Preview what will be replaced
matches = doc.find_all("30 days")

if len(matches) == 0:
    print("Text not found!")
elif len(matches) == 1:
    print(f"Will replace 1 occurrence at {matches[0].location}")
    doc.replace_tracked("30 days", "45 days")
else:
    print(f"Found {len(matches)} occurrences:")
    for match in matches:
        print(f"  - {match.location}: {match.context}")
    print("Multiple matches - need to be more specific or use occurrence parameter")
```

### Regex Search

```python
# Find all dates in YYYY-MM-DD format
matches = doc.find_all(r"\d{4}-\d{2}-\d{2}", regex=True)

for match in matches:
    print(f"Found date: {match.text} at {match.location}")
```

### Case-Insensitive Search

```python
# Find "IMPORTANT" regardless of case
matches = doc.find_all("important", case_sensitive=False)

# Will match: "important", "IMPORTANT", "Important", etc.
for match in matches:
    print(f"Match: '{match.text}' at {match.location}")
```

### Custom Context Size

```python
# Get more context for disambiguation
matches = doc.find_all("Section", context_chars=100)

for match in matches:
    # Context will show up to 100 chars before/after
    print(match.context)
```

### Search Within Scope

```python
# Search only in tables
matches = doc.find_all("text", scope={"location": "tables"})

# Search only in body (exclude tables)
matches = doc.find_all("text", scope={"location": "body"})
```

### Accessing Match Metadata

```python
matches = doc.find_all("contract term")

for match in matches:
    print(f"Text: {match.text}")
    print(f"Index: {match.index}")
    print(f"Location: {match.location}")
    print(f"Paragraph #{match.paragraph_index}")
    print(f"Full paragraph: {match.paragraph_text}")
    print(f"Context: {match.context}")
    print()
```

## Match Object

The `Match` class represents a single search result:

```python
@dataclass
class Match:
    index: int                # Position in results (0-based)
    text: str                 # The matched text
    context: str              # Surrounding text with ellipsis
    paragraph_index: int      # Paragraph position in document
    paragraph_text: str       # Full paragraph text
    location: str             # Human-readable location
    span: TextSpan           # Underlying TextSpan for advanced use
```

### String Representations

```python
# User-friendly display
print(str(match))
# Output: [0] body: ...Therefore, production products utilizing...

# Detailed representation
print(repr(match))
# Output: Match(index=0, text='production products', location='body', paragraph_index=33)
```

## Location Strings

The `location` attribute provides human-readable location information:

| Location | Description | Example |
|----------|-------------|---------|
| `"body"` | Main document body | `"body"` |
| `"table:X:row:Y:cell:Z"` | Table cell with indices | `"table:0:row:2:cell:1"` |
| `"header"` | Document header | `"header"` |
| `"footer"` | Document footer | `"footer"` |
| `"footnote"` | Footnote | `"footnote"` |
| `"endnote"` | Endnote | `"endnote"` |

## Use Cases

### 1. Understanding Ambiguous Searches

When you're not sure how many occurrences exist:

```python
matches = doc.find_all("the")
print(f"Document contains {len(matches)} occurrences of 'the'")
```

### 2. Selective Replacement

Preview matches to decide which to replace:

```python
matches = doc.find_all("Company")

# Filter to only table occurrences
table_matches = [m for m in matches if "table:" in m.location]
print(f"Found {len(table_matches)} occurrences in tables")

# Or filter by paragraph
relevant_matches = [m for m in matches if "Section 3" in m.paragraph_text]
```

### 3. Validation Before Edits

Ensure text exists before attempting an operation:

```python
matches = doc.find_all("old clause text")

if len(matches) == 0:
    print("Error: Text not found in document")
elif len(matches) > 1:
    print(f"Warning: Found {len(matches)} occurrences:")
    for match in matches:
        print(f"  - {match.location}")
else:
    # Safe to replace
    doc.replace_tracked("old clause text", "new clause text")
```

### 4. Bulk Analysis

Analyze patterns across a document:

```python
# Find all dollar amounts
amounts = doc.find_all(r"\$[\d,]+\.?\d*", regex=True)

print(f"Document contains {len(amounts)} dollar amounts:")
for amount in amounts:
    print(f"  {amount.text} at {amount.location}")
```

### 5. Building an Index

Create a map of where terms appear:

```python
terms = ["confidential", "proprietary", "trade secret"]

for term in terms:
    matches = doc.find_all(term, case_sensitive=False)
    if matches:
        print(f"\n'{term}' appears {len(matches)} time(s):")
        for match in matches:
            print(f"  - Paragraph {match.paragraph_index}: {match.location}")
```

## Error Handling

```python
import re

# Invalid regex raises re.error
try:
    matches = doc.find_all(r"[invalid(", regex=True)
except re.error as e:
    print(f"Invalid regex pattern: {e}")

# No matches returns empty list (not an error)
matches = doc.find_all("nonexistent text")
if not matches:
    print("No matches found")
```

## Integration with Other Methods

### With replace_tracked()

```python
# Preview then replace
matches = doc.find_all("old text")
print(f"Found {len(matches)} matches")

if len(matches) == 1:
    doc.replace_tracked("old text", "new text")
```

### With insert_tracked()

```python
# Find anchor points for insertion
anchors = doc.find_all("Section 2.1")
if anchors:
    print(f"Will insert after: {anchors[0].context}")
    doc.insert_tracked("new clause", after="Section 2.1")
```

### With delete_tracked()

```python
# Verify text exists before deleting
matches = doc.find_all("deprecated clause")
if matches:
    print(f"Will delete {len(matches)} occurrence(s)")
    doc.delete_tracked("deprecated clause")
```

## Performance Notes

- `find_all()` is a read-only operation - it doesn't modify the document
- Searching is efficient even in large documents
- Regex searches may be slower than literal searches
- Context extraction is lazy and only computed when accessed

## Comparison with Other Methods

| Method | Purpose | Returns |
|--------|---------|---------|
| `find_all()` | Preview all matches | `list[Match]` |
| `replace_tracked()` | Find and replace (must be unique) | `None` |
| `insert_tracked()` | Insert at specific location | `None` |
| `delete_tracked()` | Delete specific text | `None` |

The key difference is that `find_all()` is **read-only** and returns **all matches**, while other methods perform **write operations** and typically require **unique matches**.

## See Also

- [Text Search Algorithm](ERIC_WHITE_ALGORITHM.md)
- [Scope Evaluation](SCOPE.md)
- [Replace Tracked API](PROPOSED_API.md#replace_tracked)
