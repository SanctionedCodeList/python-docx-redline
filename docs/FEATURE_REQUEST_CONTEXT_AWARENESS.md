# Feature Request: Context-Aware Text Replacement

**Date:** 2025-12-08
**Priority:** Medium
**Category:** API Enhancement / Error Prevention
**Status:** ✅ IMPLEMENTED (2025-12-08)

## Implementation Summary

The following features have been successfully implemented:

1. **Context Preview** (`show_context` parameter) - Shows text before/after match
2. **Fragment Detection** (`check_continuity` parameter) - Warns about potential sentence fragments
3. **ContinuityWarning** - New warning class for fragment detection
4. **Comprehensive Tests** - 10 tests covering all edge cases

---

## Problem Statement

When replacing text with `replace_tracked()`, users can inadvertently create sentence fragments if the replacement doesn't account for how the following text connects to it.

### Real-World Example

**Edit 9 from surgical edits workflow:**

**Original text:**
```
BatchLeads functions like the attorney directory in Vrdolyak and the
historical records database in Callahan. The product in question here
is property ownership information...
```

**Intended replacement:**
```python
doc.replace_tracked(
    "BatchLeads functions like the attorney directory in Vrdolyak and the historical records database in Callahan.",
    "BatchLeads stands even further from right-of-publicity concerns than the attorney directory in Vrdolyak. There, users searched for attorneys by name—the product was a directory of people. Here, users search for properties by address, and the owner's name is incidental data attached to the property record."
)
```

**Result:**
```
BatchLeads stands even further from right-of-publicity concerns than
the attorney directory in Vrdolyak. There, users searched for attorneys
by name—the product was a directory of people. Here, users search for
properties by address, and the owner's name is incidental data attached
to the property record. in question here is property ownership information...
                         ^^^^^^^^^ SENTENCE FRAGMENT!
```

**Issue:** The next sentence begins with "in question here" which requires "The product" before it to be grammatically correct.

---

## Root Cause Analysis

The user replaced only the first sentence, not realizing that the second sentence ("The product in question here...") grammatically depends on being a standalone sentence. After the replacement:

1. Old first sentence: "BatchLeads functions like..."
2. New first sentence: "BatchLeads stands even further..." (ends with "property record.")
3. Old second sentence: "The product in question here..." (still present)

But during replacement, somehow "The product" at the start of the second sentence was lost, creating " in question here..." which is a fragment.

**Hypothesis:** The text search may have matched beyond the period, or the replacement affected the following text's structure.

---

## Proposed Solutions

### Option 1: Context Preview (Recommended)

Add a `show_context` parameter that displays text before and after the match:

```python
result = doc.replace_tracked(
    find="BatchLeads functions like the attorney directory in Vrdolyak and the historical records database in Callahan.",
    replace="BatchLeads stands even further from...",
    show_context=True  # Preview before committing
)

# Output:
# Match found at paragraph 45:
#
# BEFORE (50 chars): "...promotional purposes. "
# MATCH: "BatchLeads functions like the attorney directory in Vrdolyak and the historical records database in Callahan."
# AFTER (50 chars): " The product in question here is property ow..."
#
# Replacement will be:
# "BatchLeads stands even further from right-of-publicity concerns than the attorney directory in Vrdolyak. There, users searched for attorneys by name—the product was a directory of people. Here, users search for properties by address, and the owner's name is incidental data attached to the property record."
#
# Next text will be: " The product in question here is..."
#
# Proceed? [y/n]
```

**Benefits:**
- Shows user exactly what comes after the replacement
- Helps catch sentence continuity issues
- Simple API addition
- Non-breaking change (default: False)

### Option 2: Sentence Fragment Detection

Analyze the text immediately following the replacement and warn about potential fragments:

```python
result = doc.replace_tracked(
    find="...",
    replace="...",
    check_continuity=True  # default: True
)

# If next text starts with lowercase or connecting phrase:
# Warning: Potential sentence fragment detected after replacement.
# Next text begins with: " in question here is property..."
# Consider including more context in your replacement or adjusting the following text.
```

**Detection heuristics:**
- Next text starts with lowercase letter (except "i" for Roman numerals)
- Next text starts with connecting phrase: "in question", "of which", "that is", "to which", etc.
- Next text continues a sentence: starts with comma, semicolon, etc.

**Benefits:**
- Proactive error detection
- Helps users catch issues before they occur
- Can be disabled for power users

**Challenges:**
- Heuristics not 100% accurate
- May produce false positives

### Option 3: Multi-Sentence Replacement Helper

Provide a helper that understands sentence boundaries:

```python
# Find sentences around target text
sentences = doc.find_sentences_containing("BatchLeads functions like")

# Returns:
# [
#   Sentence(text="...", start=..., end=...),  # Previous
#   Sentence(text="BatchLeads functions like the attorney directory...", start=..., end=...),  # Target
#   Sentence(text="The product in question here is property ownership...", start=..., end=...)  # Next
# ]

# Replace with awareness of sentence boundaries
doc.replace_sentences(
    find_sentence="BatchLeads functions like...",
    replace_with=[
        "BatchLeads stands even further from right-of-publicity concerns...",
        "There, users searched for attorneys by name—the product was a directory of people.",
        "Here, users search for properties by address, and the owner's name is incidental data attached to the property record."
    ],
    keep_next_sentence=True  # or merge_with_next, delete_next, etc.
)
```

**Benefits:**
- Explicit sentence-level editing
- Clear control over sentence boundaries
- Prevents accidental fragments

**Challenges:**
- More complex API
- Sentence detection may be imperfect (legal text has complex sentence structures)

### Option 4: Dry-Run / Preview Mode

Allow previewing changes before committing:

```python
# Preview mode
preview = doc.replace_tracked(
    find="...",
    replace="...",
    preview=True  # Don't make changes yet
)

print(preview.old_text)  # What will be removed
print(preview.new_text)  # What will be added
print(preview.context_before)  # 100 chars before
print(preview.context_after)   # 100 chars after
print(preview.result_text)     # Full paragraph after change

# Apply if satisfied
if user_approves():
    preview.apply()
```

**Benefits:**
- Non-destructive testing
- Clear visibility into changes
- Aligns with "measure twice, cut once" philosophy

### Option 5: Include Context Expansion

Allow users to easily expand their match to include surrounding text:

```python
# Match with context window
match = doc.find_text_with_context(
    target="BatchLeads functions like...",
    after_chars=50  # Include 50 chars after match
)

print(match.target)     # "BatchLeads functions like..."
print(match.after)      # " The product in question here is..."

# Replace with expanded context
doc.replace_tracked(
    find=match.target + " The product",  # Include start of next sentence
    replace="BatchLeads stands even further... The product"  # Preserve it
)
```

---

## Recommended Implementation

**Short-term (Easy Win):**

Add `show_context` parameter to `replace_tracked()`:

```python
def replace_tracked(
    self,
    find: str,
    replace: str,
    regex: bool = False,
    show_context: bool = False,
    context_chars: int = 50,
    **kwargs
) -> ReplaceResult:
    """Replace text with optional context preview.

    Args:
        find: Text to find
        replace: Replacement text
        regex: Whether find is regex
        show_context: Show surrounding context before replacing
        context_chars: Number of characters to show before/after
        **kwargs: Additional options

    Returns:
        ReplaceResult with success/failure and context info
    """
    if show_context:
        # Find match
        match = self._find_match(find, regex=regex)
        # Show context
        self._print_context(match, context_chars)
        # Ask for confirmation (or just log, depending on use case)

    # Proceed with replacement
    return self._do_replace(find, replace, regex=regex, **kwargs)
```

**Medium-term (Better UX):**

Add sentence fragment detection with warnings:

```python
def replace_tracked(self, find: str, replace: str, check_continuity: bool = True, **kwargs):
    # ... perform replacement ...

    if check_continuity:
        next_text = self._get_text_after_match(match, chars=30)
        warnings = self._check_continuity(replace, next_text)

        if warnings:
            for warning in warnings:
                logger.warning(f"Potential issue: {warning}")
            # Optional: raise ContinuityWarning if user wants strict mode
```

**Long-term (Power Feature):**

Add sentence-aware replacement helpers:

```python
# Find sentence boundaries
sentences = doc.find_sentences_containing("target text")

# Replace with sentence awareness
doc.replace_sentences(
    find_sentence="...",
    replace_with=["...", "...", "..."],
    merge_with_next=False
)
```

---

## Alternative: Fix in Edit Specification

The other approach is to improve the edit specification in `surgical_edits.md` to be more explicit:

**Current (causes issue):**
```markdown
**BEFORE:**
> BatchLeads functions like the attorney directory in Vrdolyak and the historical records database in Callahan.

**AFTER:**
> BatchLeads stands even further from right-of-publicity concerns...
```

**Better (avoids issue):**
```markdown
**BEFORE:**
> BatchLeads functions like the attorney directory in Vrdolyak and the historical records database in Callahan. The product in question here is property ownership information

**AFTER:**
> BatchLeads stands even further from right-of-publicity concerns than the attorney directory in Vrdolyak. There, users searched for attorneys by name—the product was a directory of people. Here, users search for properties by address, and the owner's name is incidental data attached to the property record. The product in question here is property ownership information
```

This makes the replacement include enough context to preserve sentence continuity.

---

## User Experience Comparison

### Current (No Context)
```python
doc.replace_tracked(find, replace)
# ✓ Works
# ✗ Silent failure if creates fragment
# ✗ No preview
```

### With Context Preview
```python
doc.replace_tracked(find, replace, show_context=True)
# ✓ Works
# ✓ Shows before/after context
# ✓ User can spot issues
# ✗ Still requires manual checking
```

### With Fragment Detection
```python
doc.replace_tracked(find, replace, check_continuity=True)
# ✓ Works
# ✓ Warns about fragments
# ✓ Automatic detection
# ~ May have false positives
```

### With Sentence Awareness
```python
doc.replace_sentences(find_sentence, replace_with)
# ✓ Works
# ✓ Explicit sentence handling
# ✓ Prevents fragments by design
# ✗ More complex API
```

---

## Test Cases

```python
def test_replace_with_context_preview(capsys):
    """Context preview shows surrounding text."""
    doc = create_doc("Sentence one. Sentence two. Sentence three.")
    doc.replace_tracked(
        "Sentence two.",
        "New sentence.",
        show_context=True,
        context_chars=20
    )
    captured = capsys.readouterr()
    assert "Sentence one." in captured.out  # Before
    assert "Sentence three." in captured.out  # After

def test_fragment_detection_warns():
    """Warns when replacement creates potential fragment."""
    doc = create_doc("The product functions well. in question here is...")
    with pytest.warns(ContinuityWarning):
        doc.replace_tracked(
            "The product functions well.",
            "It works great.",
            check_continuity=True
        )

def test_sentence_replacement():
    """Sentence-aware replacement preserves boundaries."""
    doc = create_doc("Sentence one. Sentence two. Sentence three.")
    doc.replace_sentences(
        find_sentence="Sentence two.",
        replace_with=["New sentence.", "Another one."]
    )
    assert doc.get_text() == "Sentence one. New sentence. Another one. Sentence three."
```

---

## Priority

**MEDIUM** - This is a UX improvement that would prevent user errors, but workarounds exist:
- Users can check the document after edits (current workflow)
- Edit specifications can be written more carefully
- Manual review catches these issues

However, it would significantly improve the editing experience and prevent subtle bugs.

---

## Related Issues

- Similar to smart IDEs that warn about unreachable code or type mismatches
- Analogous to grammar checkers that detect sentence fragments
- Related to "intention-based" editing where the library understands user goals

---

## Implementation Complexity

| Feature | Effort | Benefit | Priority |
|---------|--------|---------|----------|
| Context preview | Low | High | **High** |
| Fragment detection | Medium | Medium | Medium |
| Sentence helpers | High | High | Low |
| Dry-run mode | Low | Medium | Medium |

**Recommendation:** Start with context preview as it's low-effort and high-value.
