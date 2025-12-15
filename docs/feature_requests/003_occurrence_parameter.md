# Feature Request: Occurrence Parameter for replace_tracked()

## Problem

When text appears multiple times, `replace_tracked()` raises `AmbiguousTextError`. Sometimes the user knows exactly which occurrence they want to replace (e.g., "the 2nd one") but must craft a unique search string instead.

## Use Case

A document has the phrase "consistently" appearing 5 times. The user wants to:
- Replace the 1st and 3rd occurrences
- Leave the 2nd, 4th, and 5th unchanged

Current workaround requires finding unique surrounding context for each target occurrence.

## Proposed Solution

Add an `occurrence` parameter to `replace_tracked()`:

```python
# Replace only the first occurrence (0-indexed)
doc.replace_tracked("consistently", "regularly", occurrence=0)

# Replace only the third occurrence
doc.replace_tracked("consistently", "regularly", occurrence=2)

# Replace multiple specific occurrences
doc.replace_tracked("consistently", "regularly", occurrence=[0, 2, 4])

# Replace all occurrences (current replace_all behavior, but with tracking)
doc.replace_tracked("consistently", "regularly", occurrence="all")
```

## Interaction with Existing Parameters

```python
# Combine with scope - occurrence is relative to matches within scope
doc.replace_tracked(
    "consistently",
    "regularly",
    scope="heading:Executive Summary",
    occurrence=0  # First occurrence within Executive Summary only
)
```

## Related: replace_all with Individual Tracking

Currently `replace_all=True` replaces all occurrences. It would be useful to have each replacement tracked as a separate change:

```python
# Each replacement is a separate tracked change
results = doc.replace_tracked("old", "new", occurrence="all", track_individually=True)
# Returns: [Change(...), Change(...), Change(...)]
```

## Implementation Notes

- `occurrence` parameter should accept: int, list of ints, or "all"
- When `occurrence` is specified, don't raise `AmbiguousTextError`
- Return information about which occurrences were replaced

## Priority

Medium - Useful shortcut but `find_all()` + targeted replacement is more flexible.
