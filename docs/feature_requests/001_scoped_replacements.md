# Feature Request: Scoped Replacements by Section/Heading

## Problem

When a phrase appears multiple times in a document, `replace_tracked()` raises `AmbiguousTextError`. The user must then manually expand the search string to include surrounding context to make it unique. This is tedious and error-prone.

## Use Case

A legal claim chart document has an "Executive Summary" section and multiple "Limitation Summary" sections. The phrase "fundamentally designed" appears in both the Executive Summary and in the [1.c] Limitation Summary. The user wants to replace only the one in [1.c].

Current workaround:
```python
# Must include enough context to be unique
doc.replace_tracked(
    "However, the technology is fundamentally designed for low-temperature",
    "However, the technology is designed for low-temperature"
)
```

## Proposed Solution

Add a `scope` parameter to `replace_tracked()` that limits the search to specific sections:

```python
# By heading text
doc.replace_tracked(
    "fundamentally designed",
    "described as designed",
    scope="heading:Executive Summary"
)

# By heading level
doc.replace_tracked(
    "fundamentally designed",
    "described as designed",
    scope="heading_level:2"  # Only in H2 sections
)

# By paragraph containing specific text
doc.replace_tracked(
    "fundamentally designed",
    "described as designed",
    scope="paragraph_contains:However, the technology"
)

# By table (for documents with tables)
doc.replace_tracked(
    "fundamentally designed",
    "described as designed",
    scope="table:0"  # First table only
)
```

## Implementation Notes

- The library already has some `scope` functionality mentioned in docs (`scope="section:Payment Terms"`, `scope="paragraph_containing:payment"`). This request is to ensure it works robustly and document it clearly.
- Should support multiple scope types: heading-based, table-based, paragraph-based
- Scope should be combinable: `scope=["heading:Section 1", "table:0"]`

## Priority

High - This is a common pain point when editing structured documents with repeated language.
