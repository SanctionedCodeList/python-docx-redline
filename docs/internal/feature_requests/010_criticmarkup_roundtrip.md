# Feature Request: CriticMarkup Round-Trip Workflow

**Date:** 2025-12-16
**Priority:** Medium
**Category:** New Feature / Workflow
**Effort Estimate:** ~3 days

---

## Summary

Enable a markdown-based editing workflow where:
1. DOCX documents (with existing tracked changes and comments) export to markdown with CriticMarkup syntax
2. Users review/edit in any text editor
3. Changes import back to DOCX as tracked changes

This allows document review in plain text while preserving Word's tracked changes format.

---

## Motivation

### Current Pain Points

1. **Word-only editing**: Users must use Microsoft Word to review tracked changes
2. **No plain-text workflow**: Can't use preferred text editors (VS Code, Vim, etc.)
3. **Version control unfriendly**: Binary DOCX files don't diff well in git
4. **AI agent limitations**: Agents work better with plain text than binary formats

### Proposed Solution

Use [CriticMarkup](http://criticmarkup.com/) as an interchange format:

```markdown
The {--old--}{++new++} contract states payment is due in {~~30~>45~~} days.

{++This paragraph was added.++}

{==Review this section=={>>Please verify these numbers<<}}
```

---

## CriticMarkup Syntax Reference

| Operation | Syntax | DOCX Equivalent |
|-----------|--------|-----------------|
| Insertion | `{++inserted text++}` | `<w:ins>` |
| Deletion | `{--deleted text--}` | `<w:del>` |
| Substitution | `{~~old~>new~~}` | `<w:del>` + `<w:ins>` |
| Comment | `{>>comment text<<}` | `<w:comment>` |
| Highlight | `{==marked text==}` | Comment range |
| Highlight + Comment | `{==text=={>>comment<<}}` | Comment on text |

---

## Workflow Examples

### Example 1: Review Existing Document

```python
from python_docx_redline import Document

# Load document with existing tracked changes
doc = Document("contract_with_changes.docx")

# Export to CriticMarkup markdown
markdown = doc.to_criticmarkup()
print(markdown)
# Output:
# The parties agree to {--30--}{++45++} day payment terms.
# {==Section 2.1=={>>Legal review needed<<}}

# Save for editing
with open("contract_review.md", "w") as f:
    f.write(markdown)
```

### Example 2: Apply Edits from Markdown

```python
# User edited the markdown file, adding more changes...
with open("contract_review.md") as f:
    edited_markdown = f.read()

# Apply new CriticMarkup changes back to DOCX
doc = Document("contract_original.docx")
doc.apply_criticmarkup(edited_markdown)
doc.save("contract_updated.docx")
```

### Example 3: Full Round-Trip

```python
# 1. Export
doc = Document("input.docx")
markdown = doc.to_criticmarkup()

# 2. User edits markdown (adds {++new text++}, {--removes--}, etc.)

# 3. Import back
doc.apply_criticmarkup(edited_markdown)
doc.save("output.docx")

# Result: All CriticMarkup changes appear as tracked changes in Word
```

---

## Technical Design

### Component 1: CriticMarkup Parser

**File:** `src/python_docx_redline/criticmarkup.py`

```python
from dataclasses import dataclass
from enum import Enum
import re

class OperationType(Enum):
    INSERTION = "insertion"
    DELETION = "deletion"
    SUBSTITUTION = "substitution"
    COMMENT = "comment"
    HIGHLIGHT = "highlight"

@dataclass
class CriticOperation:
    """A single CriticMarkup operation."""
    type: OperationType
    text: str                      # Main text (inserted, deleted, or marked)
    replacement: str | None = None # For substitutions: new text
    comment: str | None = None     # For comments: comment text
    position: int = 0              # Character position in source
    context_before: str = ""       # Text before this operation
    context_after: str = ""        # Text after this operation

def parse_criticmarkup(text: str) -> list[CriticOperation]:
    """Parse CriticMarkup syntax into operations.

    Args:
        text: Markdown text with CriticMarkup annotations

    Returns:
        List of CriticOperation objects in document order
    """
    operations = []

    # Regex patterns for each operation type
    patterns = {
        OperationType.INSERTION: r'\{\+\+(.+?)\+\+\}',
        OperationType.DELETION: r'\{--(.+?)--\}',
        OperationType.SUBSTITUTION: r'\{~~(.+?)~>(.+?)~~\}',
        OperationType.COMMENT: r'\{>>(.+?)<<\}',
        OperationType.HIGHLIGHT: r'\{==(.+?)==\}',
    }

    # ... implementation

    return sorted(operations, key=lambda op: op.position)

def render_criticmarkup(operations: list[CriticOperation], base_text: str) -> str:
    """Render operations back to CriticMarkup syntax."""
    # ... implementation
```

### Component 2: DOCX → CriticMarkup Export

**File:** `src/python_docx_redline/criticmarkup.py` (continued)

```python
def docx_to_criticmarkup(doc: Document, include_comments: bool = True) -> str:
    """Export document with tracked changes to CriticMarkup markdown.

    Args:
        doc: Document to export
        include_comments: Whether to include comments as {>>...<<}

    Returns:
        Markdown string with CriticMarkup annotations
    """
    output = []

    for para in doc.paragraphs:
        para_text = _paragraph_to_criticmarkup(para, doc)
        output.append(para_text)

    return "\n\n".join(output)

def _paragraph_to_criticmarkup(para, doc: Document) -> str:
    """Convert a single paragraph to CriticMarkup text."""
    result = []

    # Walk XML elements in order
    for element in para._element.iter():
        tag = element.tag.split('}')[-1]  # Strip namespace

        if tag == "ins":
            # Tracked insertion → {++text++}
            text = _extract_text_from_element(element)
            result.append(f"{{++{text}++}}")

        elif tag == "del":
            # Tracked deletion → {--text--}
            text = _extract_deltext_from_element(element)
            result.append(f"{{--{text}--}}")

        elif tag == "t":
            # Regular text (if not inside ins/del)
            if not _is_inside_tracked_change(element):
                result.append(element.text or "")

    text = "".join(result)

    # Overlay comments
    for comment in doc.comments:
        if _comment_in_paragraph(comment, para):
            if comment.marked_text:
                # {==marked text=={>>comment<<}}
                text = text.replace(
                    comment.marked_text,
                    f"{{=={comment.marked_text}==}}{{>>{comment.text}<<}}"
                )
            else:
                # Standalone comment - append to paragraph
                text += f" {{>>{comment.text}<<}}"

    return text
```

### Component 3: CriticMarkup → DOCX Import

**File:** `src/python_docx_redline/criticmarkup.py` (continued)

```python
def apply_criticmarkup(doc: Document, markup_text: str, author: str | None = None) -> ApplyResult:
    """Apply CriticMarkup changes to document as tracked changes.

    Args:
        doc: Document to modify
        markup_text: Markdown with CriticMarkup annotations
        author: Author name for tracked changes (uses doc default if None)

    Returns:
        ApplyResult with success/failure counts
    """
    operations = parse_criticmarkup(markup_text)
    results = []

    for op in operations:
        try:
            if op.type == OperationType.INSERTION:
                # Find insertion point using context
                anchor = _find_best_anchor(doc, op.context_before, op.context_after)
                doc.insert_tracked(op.text, after=anchor, author=author)

            elif op.type == OperationType.DELETION:
                doc.delete_tracked(op.text, author=author)

            elif op.type == OperationType.SUBSTITUTION:
                doc.replace_tracked(op.text, op.replacement, author=author)

            elif op.type == OperationType.COMMENT:
                # Find the highlighted text to attach comment to
                if op.context_before:  # Has highlight
                    doc.add_comment(op.comment, on=op.text)

            results.append(OperationResult(op, success=True))

        except TextNotFoundError as e:
            results.append(OperationResult(op, success=False, error=str(e)))

    return ApplyResult(results)

def _find_best_anchor(doc: Document, context_before: str, context_after: str) -> str:
    """Find the best anchor text for an insertion point.

    Uses fuzzy matching to handle minor text differences between
    markdown export and original document.
    """
    # Try exact match first
    if context_before:
        anchor = context_before.strip()[-50:]  # Last 50 chars
        matches = doc.find_all(anchor)
        if len(matches) == 1:
            return anchor

    # Fall back to fuzzy matching
    if context_before:
        matches = doc.find_all(anchor, fuzzy=0.85)
        if matches:
            return matches[0].text

    raise TextNotFoundError(f"Cannot find insertion point for context: {context_before[-30:]!r}")
```

### Component 4: Document Class Integration

**File:** `src/python_docx_redline/document.py` (additions)

```python
class Document:
    # ... existing methods ...

    def to_criticmarkup(self, include_comments: bool = True) -> str:
        """Export document with tracked changes to CriticMarkup markdown.

        Existing tracked changes become CriticMarkup syntax:
        - Insertions → {++text++}
        - Deletions → {--text--}
        - Comments → {==marked=={>>comment<<}}

        Args:
            include_comments: Whether to include comments

        Returns:
            Markdown string with CriticMarkup annotations
        """
        from .criticmarkup import docx_to_criticmarkup
        return docx_to_criticmarkup(self, include_comments)

    def apply_criticmarkup(
        self,
        markup_text: str,
        author: str | None = None,
        stop_on_error: bool = False,
    ) -> ApplyResult:
        """Apply CriticMarkup changes as tracked changes.

        CriticMarkup syntax is converted to Word tracked changes:
        - {++text++} → tracked insertion
        - {--text--} → tracked deletion
        - {~~old~>new~~} → tracked replacement
        - {>>comment<<} → Word comment

        Args:
            markup_text: Markdown with CriticMarkup annotations
            author: Author for tracked changes (uses document default if None)
            stop_on_error: Stop on first error vs continue

        Returns:
            ApplyResult with success/failure information
        """
        from .criticmarkup import apply_criticmarkup
        return apply_criticmarkup(self, markup_text, author, stop_on_error)
```

---

## Implementation Tasks

### Task 1: CriticMarkup Parser
- Parse all 5 operation types with regex
- Extract context (surrounding text) for each operation
- Handle nested operations (highlight + comment)
- Unit tests for parser

### Task 2: DOCX → CriticMarkup Export
- Walk paragraph XML elements in order
- Convert `<w:ins>` → `{++...++}`
- Convert `<w:del>` → `{--...--}`
- Overlay comments as `{==...=={>>...<<}}`
- Handle edge cases (empty paragraphs, tables, headers)

### Task 3: CriticMarkup → DOCX Import
- Use parser to extract operations
- Map operations to document locations using context
- Apply as tracked changes via existing methods
- Return detailed results (success/failure per operation)

### Task 4: Integration Tests
- Round-trip test: export → edit → import → verify
- Preserve existing changes through round-trip
- Handle documents with mixed content (tables, images)

### Task 5: SKILL.md Documentation
- Document `to_criticmarkup()` method
- Document `apply_criticmarkup()` method
- Add CriticMarkup syntax reference
- Add workflow examples

---

## Dependencies

- **Existing infrastructure used:**
  - `doc.get_tracked_changes()` - for export
  - `doc.comments` - for comment export
  - `doc.insert_tracked()`, `delete_tracked()`, `replace_tracked()` - for import
  - `doc.add_comment()` - for comment import
  - `doc.find_all()` with fuzzy matching - for context location

---

## Edge Cases to Handle

1. **Overlapping changes**: Multiple edits to same text
2. **Tables**: Tracked changes inside table cells
3. **Headers/footers**: Changes in non-body content
4. **Images**: Skip or mark with placeholder
5. **Formatting-only changes**: `<w:rPrChange>` (may skip initially)
6. **Move operations**: `<w:moveFrom>`, `<w:moveTo>` (complex - phase 2)

---

## Success Criteria

1. Round-trip preserves all insertions/deletions
2. Comments survive round-trip with correct anchoring
3. New CriticMarkup edits apply as tracked changes
4. Clear error messages for unmatched operations
5. Works with real-world legal documents

---

## Future Enhancements

- **Phase 2:** Support move operations (`{>>moved from here>>}...{<<moved to here<<}`)
- **Phase 2:** Preserve formatting (bold, italic) in markdown
- **Phase 2:** Table-aware export (pipe tables with tracked changes)
- **Phase 3:** Real-time sync (watch file for changes)
