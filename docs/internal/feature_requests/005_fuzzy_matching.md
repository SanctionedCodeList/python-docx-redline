# Feature Request: Fuzzy Matching with Similarity Threshold

## Problem

Text in Word documents often has minor variations that cause exact matches to fail:
- "would not feature" vs "would not likely feature"
- Whitespace differences (regular space vs non-breaking space)
- Punctuation variations ("don't" vs "don't" with curly apostrophe)
- OCR artifacts in scanned documents

Users must grep the document to find the exact text, then copy/paste it precisely.

## Use Case

A user wants to replace "production products would not feature" but the document actually contains "production products would not likely feature" (extra word) or has a non-breaking space somewhere in the phrase.

## Proposed Solution

Add a `fuzzy` parameter to `replace_tracked()`:

```python
# Match with 90% similarity threshold
doc.replace_tracked(
    "production products would not feature",
    "the documented technology would not describe",
    fuzzy=0.9
)

# Match with specific fuzzy algorithm
doc.replace_tracked(
    "production products would not feature",
    "the documented technology would not describe",
    fuzzy={"threshold": 0.85, "algorithm": "levenshtein"}
)
```

## Fuzzy Matching Options

### Similarity Threshold
```python
fuzzy=0.9  # 90% similarity required
```

### Algorithm Selection
```python
fuzzy={"algorithm": "levenshtein"}   # Edit distance
fuzzy={"algorithm": "token_set"}     # Token-based (good for word order variations)
fuzzy={"algorithm": "partial"}       # Substring matching
```

### Whitespace Normalization
```python
fuzzy={"normalize_whitespace": True}  # Treat all whitespace as equivalent
```

### Case Insensitivity
```python
fuzzy={"ignore_case": True}
```

## Safety Features

### Confirmation for Low-Confidence Matches
```python
doc.replace_tracked(
    "production products",
    "documented technology",
    fuzzy=0.8,
    confirm_below=0.95  # Require confirmation for matches below 95% similarity
)
# Raises ConfirmationRequired with match details if similarity is between 0.8 and 0.95
```

### Return Match Quality
```python
result = doc.replace_tracked("old text", "new text", fuzzy=0.9)
print(result.similarity)  # 0.92
print(result.matched_text)  # "old  text" (what actually matched)
```

## Integration with find_all()

```python
matches = doc.find_all("production products", fuzzy=0.85)
# Returns matches with similarity scores:
# [
#   Match(text="production products", similarity=1.0, ...),
#   Match(text="production product", similarity=0.94, ...),
#   Match(text="production-products", similarity=0.89, ...),
# ]
```

## Implementation Notes

- Use `rapidfuzz` or `thefuzz` library for fuzzy matching algorithms
- Apply fuzzy matching after concatenating text across runs
- Default should remain exact matching (fuzzy=None or fuzzy=False)
- Document the performance implications of fuzzy matching on large documents

## Priority

Medium - Very helpful for messy documents but exact matching covers most use cases.
