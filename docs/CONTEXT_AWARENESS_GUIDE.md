# Context-Aware Text Replacement Guide

This guide explains how to use the context-aware features in `docx_redline` to prevent sentence fragments and preview replacements before making changes.

## Overview

The `replace_tracked()` method now supports two optional features:

1. **Context Preview** (`show_context`) - Display text before and after the match
2. **Fragment Detection** (`check_continuity`) - Automatically detect potential sentence fragments

These features help prevent common editing mistakes where replacing text inadvertently creates grammatical errors.

## Context Preview

### Basic Usage

```python
from docx_redline import Document

doc = Document("contract.docx")
doc.replace_tracked(
    find="old text",
    replace="new text",
    show_context=True,
    context_chars=50  # Optional, default is 50
)
```

### Output Example

When `show_context=True`, the replacement will print:

```
================================================================================
CONTEXT PREVIEW
================================================================================

BEFORE (45 chars):
  '...promotional purposes. BatchLeads functions'

MATCH (87 chars):
  'like the attorney directory in Vrdolyak and the historical records database in Callahan.'

AFTER (46 chars):
  ' The product in question here is property ow...'

REPLACEMENT (216 chars):
  'stands even further from right-of-publicity concerns than the attorney directory in Vrdolyak. There, users searched for attorneys by name—the product was a directory of people.'

================================================================================
```

### Use Cases

1. **Verify Correct Match** - Ensure you're replacing the right text
2. **Check Sentence Continuity** - See what comes after to avoid fragments
3. **Review Context** - Understand the surrounding text before committing

### Parameters

- `show_context` (bool): Enable context preview (default: False)
- `context_chars` (int): Characters to show before/after (default: 50)

## Fragment Detection

### Basic Usage

```python
import warnings
from docx_redline import Document, ContinuityWarning

doc = Document("contract.docx")

# Enable warnings to see fragment detection
warnings.simplefilter("always")

doc.replace_tracked(
    find="First sentence.",
    replace="New sentence.",
    check_continuity=True
)
```

### Detection Heuristics

The fragment detector checks for three common patterns:

#### 1. Lowercase Start

**Detected:**
```python
# Text: "Sentence one. in question here is..."
# Replacing "Sentence one." will leave " in question here is..."
```

**Warning:** "Next text starts with lowercase letter - may be a sentence fragment"

#### 2. Connecting Phrases

**Detected:**
```python
# Text: "The product is X. in question here is..."
# Replacing "The product is X." will leave " in question here is..."
```

**Warning:** "Next text starts with connecting phrase 'in question' - may require preceding context"

**Phrases detected:**
- "in question"
- "of which"
- "that is"
- "to which"
- "which is"
- "who is"
- "whose"
- "wherein"
- "whereby"

#### 3. Continuation Punctuation

**Detected:**
```python
# Text: "First part, and second part..."
# Replacing "First part" will leave ", and second part..."
```

**Warning:** "Next text starts with continuation punctuation - likely a fragment"

**Punctuation detected:** `,` `;` `:` `—` `–`

### Handling Warnings

Warnings include helpful suggestions:

```python
import warnings
from docx_redline import ContinuityWarning

with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")

    doc.replace_tracked(
        "old text",
        "new text",
        check_continuity=True
    )

    # Check if warnings were issued
    if w:
        for warning in w:
            if issubclass(warning.category, ContinuityWarning):
                print(f"Warning: {warning.message}")
                # Suggestions:
                #   • Include more context in your replacement text
                #   • Adjust the 'find' text to include the connecting phrase
                #   • Review the following text to ensure grammatical correctness
```

### Suppressing Warnings

For automated workflows where you've manually verified the text:

```python
import warnings

# Disable continuity warnings
warnings.filterwarnings("ignore", category=ContinuityWarning)

doc.replace_tracked(
    "text",
    "replacement",
    check_continuity=True  # Still checks, but doesn't warn
)
```

## Using Both Features Together

```python
import warnings
from docx_redline import Document, ContinuityWarning

doc = Document("contract.docx")
warnings.simplefilter("always")

doc.replace_tracked(
    find="BatchLeads functions like the attorney directory.",
    replace="BatchLeads stands apart from right-of-publicity concerns.",
    show_context=True,       # See surrounding text
    check_continuity=True,   # Detect fragments
    context_chars=100        # More context
)

# Output will show:
# 1. Context preview with 100 chars before/after
# 2. Warning if next text creates a fragment
```

## Examples

### Example 1: Preventing Fragment

**Before:**
```
"The product in Vrdolyak was an attorney directory. in question here is property data."
```

**Attempted Replacement:**
```python
doc.replace_tracked(
    "The product in Vrdolyak was an attorney directory.",
    "BatchLeads is different.",
    check_continuity=True
)
```

**Warning Issued:**
```
ContinuityWarning: Next text starts with connecting phrase 'in question' - may require preceding context
Next text begins with: ' in question here is property data.'

Suggestions:
  • Include more context in your replacement text
  • Adjust the 'find' text to include the connecting phrase
  • Review the following text to ensure grammatical correctness
```

**Correct Approach:**
```python
# Include more context in the find/replace
doc.replace_tracked(
    "The product in Vrdolyak was an attorney directory. in question here is property data.",
    "BatchLeads is different. The focus here is property data.",
    check_continuity=True
)
```

### Example 2: Using Context Preview

**Scenario:** Replace text in a long paragraph

```python
doc.replace_tracked(
    find="30 days",
    replace="45 days",
    show_context=True,
    context_chars=80
)
```

**Output:**
```
================================================================================
CONTEXT PREVIEW
================================================================================

BEFORE (78 chars):
  '...The Buyer shall have the right to terminate this Agreement within'

MATCH (7 chars):
  '30 days'

AFTER (79 chars):
  'of the Effective Date by providing written notice to the Seller. Upon such...'

REPLACEMENT (7 chars):
  '45 days'

================================================================================
```

This shows you're changing the termination period from 30 to 45 days, with full context visible.

### Example 3: Roman Numerals Not Flagged

```python
# Text: "Section A. i. This is a list item."
doc.replace_tracked(
    "Section A.",
    "Part 1.",
    check_continuity=True
)

# No warning issued - lowercase 'i' is recognized as Roman numeral
```

## API Reference

### `replace_tracked()` Parameters

```python
def replace_tracked(
    self,
    find: str,
    replace: str,
    author: str | None = None,
    scope: str | dict | Any | None = None,
    regex: bool = False,
    enable_quote_normalization: bool = True,
    show_context: bool = False,
    check_continuity: bool = False,
    context_chars: int = 50,
) -> None:
    """Find and replace text with tracked changes.

    Args:
        find: Text or regex pattern to find
        replace: Replacement text
        author: Optional author override
        scope: Limit search scope
        regex: Whether to treat 'find' as regex
        enable_quote_normalization: Auto-convert quotes for matching
        show_context: Show text before/after the match for preview
        check_continuity: Check if replacement may create sentence fragments
        context_chars: Number of characters to show before/after when show_context=True

    Raises:
        TextNotFoundError: If text not found
        AmbiguousTextError: If multiple occurrences found

    Warnings:
        ContinuityWarning: If check_continuity=True and potential fragment detected
    """
```

### `ContinuityWarning` Class

```python
class ContinuityWarning(UserWarning):
    """Warning raised when text replacement may create a sentence fragment.

    Attributes:
        message: Description of the potential continuity issue
        next_text: The text immediately following the replacement
        suggestions: List of suggestions for fixing the issue
    """
```

## Best Practices

1. **Use Context Preview for Complex Edits**
   - Enable `show_context=True` when replacing long passages
   - Increase `context_chars` for more surrounding text

2. **Enable Fragment Detection for Legal Documents**
   - Legal text often has complex sentence structures
   - Fragments can change meaning significantly

3. **Review Warnings Carefully**
   - Warnings are heuristic-based, not perfect
   - False positives are possible (e.g., intentional fragments)

4. **Combine with Scope**
   ```python
   doc.replace_tracked(
       find="text",
       replace="replacement",
       scope={"contains": "Section 2"},
       show_context=True,
       check_continuity=True
   )
   ```

5. **Test in Development**
   - Use warnings in development to catch issues early
   - Suppress warnings in production if validated

## Performance Notes

- **Context Preview**: Minimal overhead (just text extraction)
- **Fragment Detection**: Lightweight heuristic checks
- **Both Features**: Negligible performance impact

## Limitations

1. **Heuristic-Based Detection**
   - Not a full grammar checker
   - May miss some fragments
   - May flag intentional fragments

2. **English Language Only**
   - Detection rules designed for English
   - Other languages may have false positives/negatives

3. **Single Paragraph Context**
   - Context extracted from same paragraph only
   - Cross-paragraph continuity not checked

## Future Enhancements

See `FEATURE_REQUEST_CONTEXT_AWARENESS.md` for potential future features:
- Sentence-aware replacement helpers
- Dry-run/preview mode
- Multi-sentence context expansion
