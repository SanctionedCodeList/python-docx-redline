# Feature Request: Batch Mode with Partial Success

## Problem

When applying multiple edits, a single failure stops the process or requires verbose try/except handling. Users want to:
1. Apply all possible edits
2. Get a clear report of what succeeded and what failed
3. Get actionable suggestions for failed edits

## Use Case

A user has 20 edits to apply to a document. Edit #7 fails because the text wasn't found. Currently:
- If not wrapped in try/except, the script stops at #7
- If wrapped in try/except, the user must manually track success/failure

## Proposed Solution

### Option 1: apply_edits() Method

```python
edits = [
    ("would likely not infringe", "may present distinctions from"),
    ("very likely practiced", "well-documented"),
    ("nonexistent phrase", "replacement"),  # This one will fail
    ("fundamentally designed", "described as designed"),
]

results = doc.apply_edits(edits, continue_on_error=True)

# Returns BatchResult object:
# BatchResult(
#   succeeded=[
#     EditResult(index=0, old="would likely...", new="may present...", status="success"),
#     EditResult(index=1, old="very likely...", new="well-documented", status="success"),
#     EditResult(index=3, old="fundamentally...", new="described as...", status="success"),
#   ],
#   failed=[
#     EditResult(
#       index=2,
#       old="nonexistent phrase",
#       new="replacement",
#       status="not_found",
#       error=TextNotFoundError(...),
#       suggestions=["nonexistent phrases (1 match)", "existent phrase (2 matches)"]
#     ),
#   ],
#   summary="3/4 edits applied successfully"
# )

# Pretty print results
print(results)
# ✓ 3 edits applied
# ✗ 1 edit failed:
#   [2] "nonexistent phrase" → TextNotFoundError
#       Suggestions: "nonexistent phrases" (1 match)
```

### Option 2: Context Manager

```python
with doc.batch_edits(continue_on_error=True) as batch:
    doc.replace_tracked("old1", "new1")
    doc.replace_tracked("old2", "new2")
    doc.replace_tracked("old3", "new3")

print(batch.results)  # Same BatchResult object
```

## Features

### Suggestions for Failed Edits

When an edit fails, provide helpful suggestions:
- For `TextNotFoundError`: Similar strings that do exist
- For `AmbiguousTextError`: List all matches with context

### Dry Run Mode

```python
results = doc.apply_edits(edits, dry_run=True)
# Shows what would happen without making changes
```

### Rollback on Failure

```python
results = doc.apply_edits(edits, rollback_on_error=True)
# If any edit fails, undo all changes
```

## Implementation Notes

- `apply_edits()` should accept list of tuples or list of Edit objects
- Support loading edits from YAML (existing feature) with same error handling
- Results should be serializable for logging/reporting

## Priority

High - Essential for reliable batch processing of documents.
