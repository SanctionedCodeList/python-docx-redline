# Feature Request: find_all() Method for Previewing Matches

## Problem

There's no way to preview what text will match before committing to a replacement. Users must run `replace_tracked()`, see if it fails with `AmbiguousTextError` or `TextNotFoundError`, then adjust. This trial-and-error workflow is slow and frustrating.

## Use Case

A user wants to replace "production products" throughout a document but first needs to:
1. See how many occurrences exist
2. View the context around each occurrence
3. Decide which ones to replace and which to leave

## Proposed Solution

Add a `find_all()` method that returns all matches with context:

```python
matches = doc.find_all("production products")

# Returns list of Match objects:
# [
#   Match(
#     index=0,
#     text="production products",
#     context="...Therefore, production products utilizing this technology...",
#     paragraph_index=33,
#     paragraph_text="Therefore, production products utilizing...",
#     location="body"  # or "table:0:row:2:cell:1"
#   ),
#   Match(
#     index=1,
#     text="production products",
#     context="...Therefore, production products utilizing Adeia's...",
#     paragraph_index=2344,
#     paragraph_text="Therefore, production products utilizing Adeia's...",
#     location="table:0:row:45:cell:1"
#   )
# ]

# Print formatted preview
for m in matches:
    print(f"[{m.index}] {m.location}: ...{m.context}...")
```

## Additional Features

### Regex support
```python
matches = doc.find_all(r"production products \w+", regex=True)
```

### Context size control
```python
matches = doc.find_all("production products", context_chars=100)
```

### Filter by location
```python
matches = doc.find_all("production products", location="tables")  # Only in tables
matches = doc.find_all("production products", location="body")    # Exclude tables
```

## Implementation Notes

- Leverage existing `TextSearch` class
- Return rich objects with location metadata
- Include helper method `match.replace(new_text)` to replace that specific occurrence

## Priority

High - Essential for confident editing of complex documents.
