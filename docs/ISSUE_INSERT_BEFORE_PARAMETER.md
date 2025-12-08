# Issue: `insert_tracked()` Lacks `before` Parameter for Inserting Text Before Anchor

**Date:** 2025-12-08
**Severity:** Medium
**Category:** API Completeness / User Experience

---

## Summary

The `insert_tracked()` method only supports inserting text `after` an anchor point, but not `before` an anchor point. This creates an asymmetry in the API and forces users to find alternative anchor text or restructure their edits when they need to insert content before a specific phrase.

---

## Environment

- **docx_redline version:** 0.1.0 (editable install)
- **Python version:** 3.12
- **Use case:** Legal document editing - inserting new paragraphs before existing content

---

## Current Behavior

```python
from docx_redline import Document

doc = Document("brief.docx")

# This works - insert AFTER text
doc.insert_tracked("New paragraph. ", after="Introduction text.")

# This fails - insert BEFORE text
doc.insert_tracked("New paragraph. ", before="Conclusion text.")
# TypeError: Document.insert_tracked() got an unexpected keyword argument 'before'
```

**Error:**
```
TypeError: Document.insert_tracked() got an unexpected keyword argument 'before'
```

---

## Expected Behavior

The method should accept both `after` and `before` parameters, similar to how text editors and DOM manipulation APIs work:

```python
# Insert after anchor (current behavior)
doc.insert_tracked("Text to insert. ", after="anchor text")

# Insert before anchor (desired behavior)
doc.insert_tracked("Text to insert. ", before="anchor text")

# Mutual exclusivity - only one should be specified
doc.insert_tracked("Text", after="foo", before="bar")  # Should raise ValueError
```

---

## Real-World Use Case

**Scenario:** Adding a futility argument before the final sentence of a legal brief.

The document ends with:
```
...existing argumentation about defects.

The Court should dismiss all counts with prejudice.
```

**Goal:** Insert 3 new paragraphs *before* "The Court should dismiss all counts with prejudice."

### Current Workaround Required

```python
# Must find text that comes BEFORE the insertion point
doc.insert_tracked(
    futility_text,
    after="and the Indiana claim fails for lack of commercial value."
)
```

**Problems with this approach:**
1. **Cognitive overhead** - Must think backwards about insertion point
2. **Fragility** - If the prior paragraph changes, the anchor breaks
3. **Semantic mismatch** - Saying "after X" when you mean "before Y" is confusing
4. **Edge cases** - What if there's no suitable "before" text? (e.g., inserting at document start)

### Desired API

```python
# Natural, semantic approach
doc.insert_tracked(
    futility_text,
    before="The Court should dismiss all counts with prejudice."
)
```

**Benefits:**
1. **Clarity** - Intent is obvious from the code
2. **Robustness** - Anchored to the text you care about
3. **Consistency** - Matches user mental model
4. **Symmetry** - Provides both directions like jQuery's `insertAfter()`/`insertBefore()`

---

## API Design Considerations

### Option 1: Add `before` parameter (Recommended)

```python
def insert_tracked(
    self,
    text: str,
    after: Optional[str] = None,
    before: Optional[str] = None,
    scope: Optional[str] = None
) -> InsertResult:
    """Insert text with tracked changes.

    Args:
        text: Text to insert
        after: Insert after this text (optional)
        before: Insert before this text (optional)
        scope: Limit search to section/heading (optional)

    Raises:
        ValueError: If both after and before are specified
        ValueError: If neither after nor before are specified
        TextNotFoundError: If anchor text not found
    """
    if after is not None and before is not None:
        raise ValueError("Cannot specify both 'after' and 'before'")
    if after is None and before is None:
        raise ValueError("Must specify either 'after' or 'before'")

    # Implementation...
```

### Option 2: Separate methods

```python
doc.insert_tracked_after("text", anchor="foo")
doc.insert_tracked_before("text", anchor="bar")
```

**Cons of Option 2:**
- More verbose
- Breaks naming consistency with `replace_tracked()`, `delete_tracked()`
- Creates API bloat

**Recommendation:** Option 1 is cleaner and more intuitive.

---

## Implementation Notes

### Internal Changes Required

The `TextSearch.find_text()` already returns character positions, so the main changes are:

1. **Parameter validation**
   ```python
   if after and before:
       raise ValueError("Cannot specify both 'after' and 'before'")
   ```

2. **Position calculation**
   ```python
   if after:
       insert_position = match_end  # Current behavior
   elif before:
       insert_position = match_start  # New behavior
   ```

3. **Paragraph handling**
   - `after`: Insert in same paragraph if space available, else new paragraph
   - `before`: Insert in same paragraph if space available, else new paragraph before

### Edge Cases to Handle

1. **Anchor at document start**
   ```python
   doc.insert_tracked("Preface. ", before="Introduction")
   # Should insert at very beginning
   ```

2. **Anchor at paragraph boundary**
   ```python
   doc.insert_tracked("Text. ", before="New paragraph starts here")
   # Should insert before the paragraph, not within it
   ```

3. **Anchor in table cell**
   ```python
   doc.insert_tracked("Text", before="cell content")
   # Should work within table structure
   ```

4. **Multiple matches**
   - Same behavior as `after`: error with suggestions (or use `scope` parameter)

---

## Comparison with Other Libraries

### python-docx
Doesn't have tracked changes support, but has symmetrical API:
```python
paragraph.insert_paragraph_before("text")
paragraph.insert_paragraph_after("text")
```

### jQuery DOM Manipulation
```javascript
$(element).insertBefore(target);  // Insert element before target
$(element).insertAfter(target);   // Insert element after target
```

### BeautifulSoup
```python
tag.insert_before("text")
tag.insert_after("text")
```

**Industry standard:** Both directions are provided for completeness.

---

## Related Functionality

This pattern should probably extend to other operations:

### Current API
```python
doc.insert_tracked(text, after="foo")
doc.replace_tracked(find="foo", replace="bar")
doc.delete_tracked(text="foo")
```

### Potential Consistency
```python
# insert supports: after, before
doc.insert_tracked(text, before="foo")  # NEW

# replace could support: before, after for positioning new text?
# (probably not needed - replace is positional by nature)

# delete could support: before, after for context?
# (probably not needed - delete targets specific text)
```

**Conclusion:** This feature makes most sense for `insert_tracked()` since insertion is inherently directional.

---

## Priority

**MEDIUM** - Not blocking core functionality, but:
- Affects user experience and API intuitiveness
- Common enough use case (inserting before conclusions, before sections, etc.)
- Small implementation effort for significant UX improvement
- Should be included before 1.0 release to avoid API breaking changes later

---

## Test Cases Needed

```python
def test_insert_before():
    """Text can be inserted before anchor text."""
    doc = create_test_doc("Hello world. Goodbye world.")
    doc.insert_tracked("Middle text. ", before="Goodbye")
    assert doc.get_text() == "Hello world. Middle text. Goodbye world."

def test_insert_before_new_paragraph():
    """Text inserted before paragraph boundary creates new paragraph."""
    doc = create_test_doc_with_paragraphs(["First para", "Second para"])
    doc.insert_tracked("Inserted para", before="Second para")
    assert len(doc.paragraphs) == 3
    assert doc.paragraphs[1].text == "Inserted para"

def test_insert_before_and_after_mutual_exclusion():
    """Cannot specify both before and after."""
    doc = create_test_doc("Some text")
    with pytest.raises(ValueError, match="Cannot specify both"):
        doc.insert_tracked("text", before="Some", after="text")

def test_insert_neither_before_nor_after():
    """Must specify either before or after."""
    doc = create_test_doc("Some text")
    with pytest.raises(ValueError, match="Must specify either"):
        doc.insert_tracked("text")

def test_insert_before_at_document_start():
    """Can insert at very beginning using before."""
    doc = create_test_doc("First sentence.")
    doc.insert_tracked("Preface. ", before="First sentence")
    text = doc.get_text()
    assert text.startswith("Preface. First sentence")

def test_insert_before_with_scope():
    """Before parameter works with scope filtering."""
    doc = create_multi_section_doc()
    doc.insert_tracked(
        "New text. ",
        before="conclusion",
        scope="section:Analysis"
    )
    # Should insert before first "conclusion" in Analysis section only
```

---

## Documentation Impact

### README.md Update

```markdown
### Inserting Text

You can insert text either after or before an anchor point:

\`\`\`python
# Insert after existing text
doc.insert_tracked("Additional context. ", after="main argument")

# Insert before existing text
doc.insert_tracked("Preliminary note. ", before="The conclusion")

# Use scope to disambiguate
doc.insert_tracked(
    "Section intro. ",
    before="first paragraph",
    scope="section:Analysis"
)
\`\`\`
```

### Quick Reference Update

Add to insertion examples showing both directions.

---

## Workaround (Until Implemented)

Users must find alternative anchor text that comes before the desired insertion point:

```python
# Want: insert before "Conclusion"
# Must do: insert after "last sentence of prior section"
doc.insert_tracked(futility_text, after="prior sentence.")
```

This is functional but unintuitive and fragile.

---

## References

- Triggered by: Legal brief editing workflow (Edit 11 in surgical_edits.md)
- Similar patterns: DOM manipulation APIs, text editor APIs
- Related code: `src/docx_redline/document.py:229` (insert_tracked method)
