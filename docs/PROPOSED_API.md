# Proposed High-Level API for Tracked Changes

**Status**: Proposed Design
**Date**: December 6, 2025
**Author**: Parker Hancock
**Purpose**: Enable surgical document edits with tracked changes without writing XML

---

## Table of Contents

1. [Executive Summary](#executive-summary)
2. [Current Pain Points](#current-pain-points)
3. [Design Principles](#design-principles)
4. [API Reference](#api-reference)
5. [Implementation Architecture](#implementation-architecture)
6. [Examples](#examples)
7. [Migration Guide](#migration-guide)

---

## Executive Summary

### Problem
Currently, making tracked changes requires:
- Writing raw OOXML XML strings
- Manually handling text fragmentation across `<w:r>` elements
- Dealing with complex node searching and disambiguation
- No abstraction for common operations

**Result**: Simple edits take 10-20x longer than manual editing in Word.

### Solution
Provide a high-level API that:
- Hides XML complexity completely
- Automatically handles text fragmentation
- Provides smart search and disambiguation
- Enables batch operations
- Supports declarative edit specifications

**Result**: Surgical edits become as easy as string operations.

### Impact
```python
# BEFORE (current approach)
para = editor.get_node(tag="w:p", line_number=6587)
for run in para.iter(_parse_tag("w:r")):
    if "(7th Cir. 2022)" in ''.join(run.itertext()):
        editor.insert_after(run, '<w:ins><w:r><w:t> (interpreting IRPA)</w:t></w:r></w:ins>')
        break

# AFTER (proposed API)
doc.insert_tracked(" (interpreting IRPA)", after="(7th Cir. 2022)", scope="Huston")
```

**Time savings**: 5-10x faster to write, more maintainable, less error-prone.

---

## Current Pain Points

### 1. XML String Construction
Every edit requires hand-crafted XML:
```python
xml = '<w:ins w:id="4" w:author="Hancock, Parker" w:date="2025-12-06T06:55:52Z" w16du:dateUtc="2025-12-06T06:55:52Z"><w:r w:rsidR="F3F4F4B4"><w:t xml:space="preserve"> (interpreting IRPA)</w:t></w:r></w:ins>'
```

**Issues**:
- Error-prone (wrong quotes, missing attributes)
- Not readable
- No IDE support
- Hard to maintain

### 2. Text Fragmentation
Text in Word appears continuous but is split across multiple runs:

```xml
<!-- What you see: "The Seventh Circuit has made clear" -->
<!-- What exists in XML: -->
<w:r><w:t>The Seventh </w:t></w:r>
<w:r><w:t>Circuit has </w:t></w:r>
<w:r><w:t>made clear</w:t></w:r>
```

**Result**: Every text search requires:
1. Iterate all paragraphs
2. Iterate all runs per paragraph
3. Concatenate text across runs
4. Pattern match
5. Determine insertion point

### 3. Disambiguation Complexity
Same text appears multiple times:
- Table of Authorities
- Main body
- Footnotes

**Current solution**: Manually specify line ranges
```python
para = editor.get_node(tag="w:p", contains="Huston", line_number=range(6600, 6700))
```

**Issues**:
- Line numbers shift after edits
- Requires grep to find line numbers
- Brittle

### 4. Validation Errors Are Cryptic
```
FAILED - Found 3 deletion validation violations:
  word/document.xml: Line 10868: <w:t> found within <w:del>: 'compiles'
```

**Better error**:
```
ValidationError: Incorrect element in deletion
  Problem: Used <w:t> inside <w:del> tag
  Solution: Use <w:delText> instead for deleted text
  Location: Line 10868 in word/document.xml
  Fix: Replace '<w:t>compiles</w:t>' with '<w:delText>compiles</w:delText>'
```

### 5. No Batch Operations
Can't say "apply these 10 edits" - must write custom loop logic each time.

---

## Design Principles

### 1. Progressive Disclosure
Start simple, add complexity only when needed:

```python
# Level 1: Simple (handles 80% of cases)
doc.insert_tracked(" text", after="target")

# Level 2: Add scope when needed
doc.insert_tracked(" text", after="target", scope="section:Argument")

# Level 3: Use search results for disambiguation
result = doc.find_text("target", scope=...)
result.insert_after(" text")

# Level 4: Drop to XML for edge cases (escape hatch)
editor = doc.get_editor("word/document.xml")
editor.insert_after(node, '<w:ins>...</w:ins>')
```

### 2. Sensible Defaults
```python
# Author inherited from Document
doc = Document('file.docx', author="Hancock, Parker")
doc.insert_tracked("text", ...)  # Uses "Hancock, Parker"

# Override per-edit
doc.insert_tracked("text", ..., author="Different Author")
```

### 3. Automatic Complexity Management
Library handles:
- Text fragmentation across runs
- RSID generation
- Timestamp generation
- Change ID sequencing
- Namespace management

User just says what they want, not how to do it.

### 4. Clear, Actionable Errors
```python
TextNotFoundError: Could not find "(7th Cir. 2022)" in scope "Huston"

Suggestions:
  • Text may be fragmented - try fuzzy matching
  • Found similar: "(7th Cir.2022)" (missing space)
  • Found in Table of Authorities (line 452)

Next steps:
  • Expand scope: scope={"section": "Argument"}
  • Use fuzzy match: fuzzy=True
  • Manual search: doc.find_all("7th Cir. 2022")
```

### 5. Composition Over Configuration
Build complex operations from simple ones:

```python
# Atomic operations compose
doc.insert_tracked(...)
doc.replace_tracked(...)
doc.delete_tracked(...)

# Combine into batch
doc.apply_edits([
    {"type": "insert", ...},
    {"type": "replace", ...},
])

# Or use templates
doc.apply_template(citation_template, ...)
```

---

## API Reference

### Document Class

```python
class Document:
    """High-level document manipulation with tracked changes support."""

    def __init__(
        self,
        path: str | Path,
        author: str = "Claude",
        initials: str | None = None,
        rsid: str | None = None,
        track_revisions: bool = False
    ):
        """
        Initialize document for editing.

        Args:
            path: Path to .docx file or unpacked directory
            author: Default author for tracked changes
            initials: Author initials (auto-derived from author if not provided)
            rsid: RSID for edits (auto-generated if not provided)
            track_revisions: Enable track revisions mode in document settings

        Examples:
            >>> doc = Document('file.docx')
            >>> doc = Document('file.docx', author="Hancock, Parker", initials="PH")
        """
```

### Text Operations

```python
def insert_tracked(
    self,
    text: str,
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Callable | None = None,
    author: str | None = None,
    fuzzy: bool = False
) -> TextSpan:
    """
    Insert text with tracked changes.

    Args:
        text: Text to insert
        after: Insert after this text (mutually exclusive with before)
        before: Insert before this text (mutually exclusive with after)
        scope: Limit search scope (see Scope System below)
        author: Override default author for this edit
        fuzzy: Enable fuzzy text matching

    Returns:
        TextSpan representing the inserted text

    Raises:
        TextNotFoundError: If after/before text not found in scope
        AmbiguousTextError: If multiple matches found (provide better scope)

    Examples:
        >>> # Simple insertion
        >>> doc.insert_tracked(" (interpreting IRPA)", after="(7th Cir. 2022)")

        >>> # With scope to disambiguate
        >>> doc.insert_tracked(
        ...     " (granting motion to dismiss)",
        ...     after="(N.D. Ill. 2016)",
        ...     scope="paragraph_containing:Vrdolyak"
        ... )

        >>> # Different author
        >>> doc.insert_tracked(
        ...     " text",
        ...     after="target",
        ...     author="Different Author"
        ... )
    """

def replace_tracked(
    self,
    find: str,
    replace: str,
    scope: str | dict | Callable | None = None,
    author: str | None = None,
    fuzzy: bool = False,
    first_only: bool = False
) -> list[TextSpan]:
    """
    Find and replace text with tracked changes.

    Args:
        find: Text to find
        replace: Replacement text
        scope: Limit search scope
        author: Override default author
        fuzzy: Enable fuzzy matching
        first_only: Only replace first occurrence

    Returns:
        List of TextSpan objects representing replacements

    Examples:
        >>> # Simple replacement
        >>> doc.replace_tracked("records", "compiles")

        >>> # With scope
        >>> doc.replace_tracked(
        ...     "records their property ownership",
        ...     "compiles their property ownership data",
        ...     scope="paragraph_containing:It claims merely"
        ... )

        >>> # Replace only first occurrence
        >>> doc.replace_tracked("the", "THE", first_only=True)
    """

def delete_tracked(
    self,
    text: str,
    scope: str | dict | Callable | None = None,
    author: str | None = None,
    fuzzy: bool = False
) -> list[TextSpan]:
    """
    Delete text with tracked changes.

    Args:
        text: Text to delete
        scope: Limit search scope
        author: Override default author
        fuzzy: Enable fuzzy matching

    Returns:
        List of TextSpan objects representing deletions

    Examples:
        >>> doc.delete_tracked("for example,", scope="perceived endorsement")
    """
```

### Structural Operations

```python
def insert_paragraph(
    self,
    text: str,
    after: str | None = None,
    before: str | None = None,
    after_heading: str | None = None,
    scope: str | dict | None = None,
    track: bool = True,
    style: str = "Normal"
) -> Paragraph:
    """
    Insert a new paragraph.

    Args:
        text: Paragraph text
        after: Insert after paragraph containing this text
        before: Insert before paragraph containing this text
        after_heading: Insert after this heading
        scope: Limit search scope
        track: Insert as tracked change
        style: Paragraph style (default: "Normal")

    Returns:
        Paragraph object representing the inserted paragraph

    Examples:
        >>> doc.insert_paragraph(
        ...     "This case is about property records, not property owners...",
        ...     after="mismatch between the allegations and the law",
        ...     track=True
        ... )

        >>> doc.insert_paragraph(
        ...     "New section content",
        ...     after_heading="Property Ownership Data",
        ...     style="Body Text"
        ... )
    """

def insert_paragraphs(
    self,
    texts: list[str],
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | None = None,
    track: bool = True,
    style: str = "Normal"
) -> list[Paragraph]:
    """
    Insert multiple paragraphs.

    Args:
        texts: List of paragraph texts
        after: Insert after paragraph containing this text
        before: Insert before paragraph containing this text
        scope: Limit search scope
        track: Insert as tracked change
        style: Paragraph style for all paragraphs

    Returns:
        List of Paragraph objects

    Examples:
        >>> doc.insert_paragraphs(
        ...     [
        ...         "First paragraph of futility argument...",
        ...         "Second paragraph...",
        ...         "Third paragraph..."
        ...     ],
        ...     after="dismiss all counts with prejudice",
        ...     track=True
        ... )
    """

def delete_section(
    self,
    heading: str,
    track: bool = True,
    update_toc: bool = False,
    scope: str | dict | None = None
) -> Section:
    """
    Delete an entire section (heading + all content until next heading).

    Args:
        heading: Heading text of section to delete
        track: Delete as tracked change
        update_toc: Automatically update Table of Contents
        scope: Limit search scope

    Returns:
        Section object representing deleted section

    Examples:
        >>> doc.delete_section(
        ...     "Enhanced Damages Fail as a Matter of Law",
        ...     track=True,
        ...     update_toc=True
        ... )
    """

def move_paragraph(
    self,
    from_text: str,
    to_after: str | None = None,
    to_before: str | None = None,
    track: bool = True,
    scope: str | dict | None = None
) -> Paragraph:
    """
    Move a paragraph to a new location.

    Args:
        from_text: Text identifying paragraph to move
        to_after: Move to after this text
        to_before: Move to before this text
        track: Move as tracked change
        scope: Limit search scope

    Returns:
        Paragraph object at new location
    """
```

### Search and Disambiguation

```python
def find_text(
    self,
    text: str,
    scope: str | dict | Callable | None = None,
    fuzzy: bool = False
) -> TextSpan:
    """
    Find text in document, handling fragmentation automatically.

    Args:
        text: Text to find
        scope: Limit search scope
        fuzzy: Enable fuzzy matching

    Returns:
        TextSpan object representing found text

    Raises:
        TextNotFoundError: If text not found
        AmbiguousTextError: If multiple matches (use better scope)

    Examples:
        >>> # Find and manipulate
        >>> span = doc.find_text("(7th Cir. 2022)", scope="Huston")
        >>> span.insert_after(" (interpreting IRPA)")

        >>> # Preview context before editing
        >>> span = doc.find_text("target text")
        >>> print(span.context)  # Shows surrounding text
        >>> span.replace("new text")
    """

def find_all(
    self,
    text: str | None = None,
    contains: str | None = None,
    in_section: str | None = None,
    scope: str | dict | None = None
) -> list[TextSpan]:
    """
    Find all occurrences of text.

    Args:
        text: Exact text to find
        contains: Find paragraphs containing this text
        in_section: Limit to this section
        scope: Additional scope constraints

    Returns:
        List of TextSpan objects

    Examples:
        >>> # Find all citations to disambiguate
        >>> results = doc.find_all(contains="N.D. Ill")
        >>> for i, result in enumerate(results):
        ...     print(f"{i}: {result.context}")

        >>> # Use specific result
        >>> results[1].insert_after(" (granting motion to dismiss)")
    """

def find_paragraph(
    self,
    containing: str | None = None,
    ending_with: str | None = None,
    starting_with: str | None = None,
    style: str | None = None,
    scope: str | dict | None = None
) -> Paragraph:
    """
    Find a paragraph by various criteria.

    Args:
        containing: Paragraph must contain this text
        ending_with: Paragraph must end with this text
        starting_with: Paragraph must start with this text
        style: Paragraph must have this style
        scope: Additional scope constraints

    Returns:
        Paragraph object

    Examples:
        >>> para = doc.find_paragraph(
        ...     ending_with="mismatch between the allegations and the law"
        ... )
        >>> doc.insert_paragraph("New para", after_paragraph=para)
    """
```

### Batch Operations

```python
def apply_edits(
    self,
    edits: list[dict],
    stop_on_error: bool = False
) -> list[EditResult]:
    """
    Apply multiple edits in sequence.

    Args:
        edits: List of edit specifications (see examples)
        stop_on_error: Stop processing on first error

    Returns:
        List of EditResult objects (success/failure per edit)

    Examples:
        >>> edits = [
        ...     {
        ...         "type": "insert_tracked",
        ...         "text": " (interpreting IRPA)",
        ...         "after": "(7th Cir. 2022)",
        ...         "scope": "Huston"
        ...     },
        ...     {
        ...         "type": "replace_tracked",
        ...         "find": "records",
        ...         "replace": "compiles",
        ...         "scope": "It claims merely"
        ...     },
        ...     {
        ...         "type": "delete_section",
        ...         "heading": "Enhanced Damages",
        ...         "update_toc": True
        ...     }
        ... ]
        >>> results = doc.apply_edits(edits)
        >>> print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
    """

def apply_edit_file(
    self,
    path: str | Path,
    format: str = "yaml"
) -> list[EditResult]:
    """
    Apply edits from YAML or JSON file.

    Args:
        path: Path to edit specification file
        format: File format ("yaml" or "json")

    Returns:
        List of EditResult objects

    Examples:
        >>> doc.apply_edit_file('surgical_edits.yaml')
    """
```

### Document Management

```python
def accept_all_changes(self) -> AcceptResult:
    """
    Accept all tracked changes in document.

    Returns:
        AcceptResult with counts of accepted insertions/deletions

    Examples:
        >>> result = doc.accept_all_changes()
        >>> print(f"Accepted {result.insertions} insertions, {result.deletions} deletions")
    """

def reject_all_changes(self) -> RejectResult:
    """Reject all tracked changes in document."""

def delete_all_comments(self) -> int:
    """
    Delete all comments from document.

    Returns:
        Number of comments deleted
    """

def save(
    self,
    path: str | Path | None = None,
    validate: bool = True
) -> None:
    """
    Save document to file.

    Args:
        path: Output path (if None, overwrites original)
        validate: Run validation before saving

    Raises:
        ValidationError: If validation fails
    """

def validate(self) -> ValidationResult:
    """
    Validate document structure without saving.

    Returns:
        ValidationResult with any errors/warnings
    """
```

### TextSpan Class

```python
class TextSpan:
    """
    Represents found text that may span multiple runs.
    Provides methods for manipulating the text.
    """

    @property
    def text(self) -> str:
        """The matched text."""

    @property
    def context(self) -> str:
        """Surrounding context for disambiguation."""

    @property
    def location(self) -> Location:
        """Location info (section, paragraph, line number)."""

    def insert_before(self, text: str, track: bool = True) -> TextSpan:
        """Insert text before this span."""

    def insert_after(self, text: str, track: bool = True) -> TextSpan:
        """Insert text after this span."""

    def replace(self, text: str, track: bool = True) -> TextSpan:
        """Replace this span with new text."""

    def delete(self, track: bool = True) -> None:
        """Delete this span."""

    def highlight(self, color: str = "yellow") -> None:
        """Add highlighting to this text."""
```

---

## Scope System

The scope parameter accepts multiple formats for flexibility:

### String Shortcuts

```python
# Paragraph containing text
scope="Huston"
scope="paragraph_containing:Huston"

# In a specific section
scope="section:Argument"
scope="section:II"  # By number

# Under a heading
scope="heading:Property Ownership Data"

# Not in specific sections
scope="not_in:Table of Authorities"
```

### Dictionary Format

```python
scope={
    "contains": "Huston",           # Paragraph must contain this
    "section": "Argument",          # In this section
    "line_range": (6500, 6700),    # Within these lines
    "not_in": ["TOA", "TOC"],      # Exclude these sections
    "style": "Normal"               # Paragraph style
}
```

### Callable Format

```python
def my_scope(paragraph: Paragraph) -> bool:
    """Custom scope logic."""
    return (
        "Huston" in paragraph.text and
        paragraph.style == "Normal" and
        paragraph.section == "Argument"
    )

scope=my_scope
```

---

## Implementation Architecture

### Core Components

```
docx_redline/
├── document.py          # High-level Document class
├── text_span.py         # TextSpan class for fragmentation handling
├── search.py            # Search and disambiguation logic
├── scope.py             # Scope parsing and evaluation
├── tracked_changes.py   # Tracked change creation
├── structural.py        # Paragraph/section operations
├── batch.py             # Batch edit processing
├── validation.py        # Enhanced validation with better errors
└── xml/                 # Low-level XML manipulation (existing)
    ├── editor.py
    └── utilities.py
```

### Key Algorithms

#### 1. Text Search with Fragmentation Handling

```python
def find_text(doc, text, scope):
    """
    Algorithm:
    1. Filter paragraphs by scope
    2. For each paragraph:
       a. Concatenate all run text
       b. Find all matches of search text
       c. For each match:
          - Determine which runs contain it
          - Record start/end positions within runs
    3. Return TextSpan objects
    """

    # Example: Finding "(7th Cir. 2022)" across fragmented runs
    # Runs: ["(7th Cir. ", "2022)"]
    # Result: TextSpan(runs=[run1, run2], start_offset=0, end_offset=5)
```

#### 2. Surgical Text Replacement

```python
def replace_in_span(span, new_text, track=True):
    """
    Algorithm:
    1. If span is within single run:
       a. Split run into: [before, span, after]
       b. Replace span run with deletion + insertion

    2. If span crosses multiple runs:
       a. Preserve runs before/after span completely
       b. Mark spanned runs for deletion
       c. Insert new text as insertion

    3. Preserve formatting:
       - Copy <w:rPr> from original runs
       - Maintain RSIDs where appropriate
    """
```

#### 3. Scope Evaluation

```python
def evaluate_scope(paragraph, scope):
    """
    Algorithm:
    1. Parse scope specification
    2. Check each criterion:
       - Text containment
       - Section membership
       - Line number range
       - Style matching
       - Exclusions
    3. Return True if all criteria match
    """
```

### Data Structures

```python
@dataclass
class TextSpan:
    """Represents found text across potentially multiple runs."""
    runs: list[Element]          # lxml elements
    start_run_index: int         # Which run starts the span
    end_run_index: int           # Which run ends the span
    start_offset: int            # Character offset in start run
    end_offset: int              # Character offset in end run
    paragraph: Element           # Parent paragraph

@dataclass
class Location:
    """Location information for a text span."""
    section: str | None
    heading: str | None
    paragraph_index: int
    line_number: int | None

@dataclass
class EditResult:
    """Result of applying a single edit."""
    success: bool
    edit_type: str
    message: str
    span: TextSpan | None
```

### Error Handling

```python
class TextNotFoundError(Exception):
    """
    Raised when text cannot be found in scope.
    Includes suggestions for debugging.
    """
    def __init__(self, text, scope, suggestions):
        self.text = text
        self.scope = scope
        self.suggestions = suggestions

    def __str__(self):
        msg = f"Could not find '{self.text}' in scope '{self.scope}'\n\n"
        msg += "Suggestions:\n"
        for s in self.suggestions:
            msg += f"  • {s}\n"
        return msg

class AmbiguousTextError(Exception):
    """
    Raised when text has multiple matches.
    Shows all matches with context.
    """
    def __init__(self, text, matches):
        self.text = text
        self.matches = matches

    def __str__(self):
        msg = f"Found {len(self.matches)} occurrences of '{self.text}'\n\n"
        for i, match in enumerate(self.matches):
            msg += f"{i}: {match.context}\n"
        msg += "\nProvide better scope to disambiguate."
        return msg
```

---

## Examples

### Example 1: Simple Surgical Edits

```python
from docx_redline import Document

# Open document
doc = Document('motion.docx', author="Hancock, Parker")

# Accept client changes
doc.accept_all_changes()
doc.delete_all_comments()

# Add procedural parentheticals
doc.insert_tracked(" (interpreting IRPA)", after="(7th Cir. 2022)", scope="Huston")
doc.insert_tracked(" (granting motion to dismiss)", after="(N.D. Ill. 2016)", scope="Vrdolyak")
doc.insert_tracked(" (dismissing with prejudice)", after="(N.D. Cal. Mar. 1, 2021)", scope="Callahan")

# Text replacement
doc.replace_tracked(
    "records their property ownership",
    "compiles their property ownership data",
    scope="It claims merely"
)

# Save
doc.save('motion_updated.docx')
```

### Example 2: Complex Structural Changes

```python
doc = Document('motion.docx', author="Hancock, Parker")

# Add new section
doc.insert_paragraph(
    "This case is about property records, not property owners. "
    "BatchLeads aggregates public information about real estate—deed histories, "
    "tax assessments, transaction prices—and makes it searchable for investors, "
    "brokers, and lenders. The owners' names appear because they are part of the "
    "property data, not because Defendants seek to exploit them for promotional purposes.",
    after="mismatch between the allegations and the law",
    track=True
)

# Delete entire section
doc.delete_section(
    "Enhanced Damages Fail as a Matter of Law",
    track=True,
    update_toc=True
)

# Add multi-paragraph futility argument
futility_paragraphs = [
    "The Court should dismiss with prejudice because the Complaint's defects are legal...",
    "That is precisely the situation here. Plaintiffs' claims fail...",
    "Controlling precedent forecloses Plaintiffs' theory..."
]

doc.insert_paragraphs(
    futility_paragraphs,
    before="The Court should dismiss all counts with prejudice",
    track=True
)

doc.save('motion_final.docx')
```

### Example 3: Batch Operations from YAML

```yaml
# surgical_edits.yaml
document: motion_draft.docx
author: Hancock, Parker
output: motion_final.docx

preprocessing:
  - accept_all_changes
  - delete_all_comments

edits:
  - type: insert_tracked
    text: " (interpreting IRPA)"
    after: "(7th Cir. 2022)"
    scope: "Huston"

  - type: insert_tracked
    text: " (granting motion to dismiss)"
    after: "(N.D. Ill. 2016)"
    scope: "Vrdolyak"

  - type: replace_tracked
    find: "records their property ownership"
    replace: "compiles their property ownership data"
    scope: "It claims merely"

  - type: insert_paragraph
    text: "This case is about property records, not property owners..."
    after: "mismatch between the allegations and the law"
    track: true

  - type: delete_section
    heading: "Enhanced Damages Fail as a Matter of Law"
    track: true
    update_toc: true
```

```python
from docx_redline import apply_edit_file

results = apply_edit_file('surgical_edits.yaml')
print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
```

### Example 4: Manual Disambiguation

```python
doc = Document('motion.docx')

# Find all occurrences
results = doc.find_all(contains="N.D. Ill")

# Show context for manual selection
for i, result in enumerate(results):
    print(f"{i}: {result.context}")
    print(f"   Location: {result.location.section}, line {result.location.line_number}\n")

# User sees:
# 0: ...206 F. Supp. 3d 1384 [(N.D. Ill)] 2016...
#    Location: Table of Authorities, line 452
#
# 1: ...Vrdolyak v. Avvo, Inc., 206 F. Supp. 3d 1384, 1387-88 [(N.D. Ill). 2016]...
#    Location: Argument, line 8249

# Use the one in main body (index 1)
results[1].insert_after(" (granting motion to dismiss)")
```

### Example 5: Custom Scope Function

```python
def in_main_body(paragraph):
    """Custom scope: main body text, not TOA/TOC."""
    return (
        paragraph.section not in ["Table of Authorities", "Table of Contents"] and
        paragraph.style == "Normal" and
        not paragraph.is_heading
    )

doc.insert_tracked(
    " (interpreting IRPA)",
    after="(7th Cir. 2022)",
    scope=in_main_body
)
```

---

## Migration Guide

### From Current API to Proposed API

#### Before: Adding Parenthetical

```python
# OLD WAY (20 lines)
doc = Document("unpacked", author="Hancock, Parker", rsid="F3F4F4B4")
editor = doc["word/document.xml"]

para = editor.get_node(tag="w:p", line_number=6587)
w_r_tag = _parse_tag("w:r")
runs = list(para.iter(w_r_tag))

for run in runs:
    text = ''.join(run.itertext())
    if "(7th Cir. 2022)" in text:
        editor.insert_after(
            run,
            '<w:ins w:id="4" w:author="Hancock, Parker" w:date="2025-12-06T06:55:52Z">'
            '<w:r w:rsidR="F3F4F4B4"><w:t> (interpreting IRPA)</w:t></w:r></w:ins>'
        )
        break

doc.save()
```

```python
# NEW WAY (3 lines)
doc = Document("file.docx", author="Hancock, Parker")
doc.insert_tracked(" (interpreting IRPA)", after="(7th Cir. 2022)", scope="Huston")
doc.save()
```

#### Before: Replacing Text

```python
# OLD WAY (30+ lines)
para = editor.get_node(tag="w:p", line_number=range(10850, 10900))
para_text = ''.join(para.itertext())

if "It claims merely that a database" in para_text and "records" in para_text:
    for run in para.iter(_parse_tag("w:r")):
        text = ''.join(run.itertext())
        if text == "records":
            replacement = (
                '<w:del><w:r><w:delText>records</w:delText></w:r></w:del>'
                '<w:ins><w:r><w:t>compiles</w:t></w:r></w:ins>'
            )
            editor.replace_node(run, replacement)
            break

    # Then find where to add " data"...
    for run in para.iter(_parse_tag("w:r")):
        text = ''.join(run.itertext())
        if "property ownership" in text:
            editor.insert_after(
                run,
                '<w:ins><w:r><w:t xml:space="preserve"> data</w:t></w:r></w:ins>'
            )
            break
```

```python
# NEW WAY (1 line)
doc.replace_tracked(
    "records their property ownership",
    "compiles their property ownership data",
    scope="It claims merely"
)
```

#### Before: Deleting Section

```python
# OLD WAY (complex, multi-step)
# 1. Find section heading
heading = editor.get_node(tag="w:p", contains="Enhanced Damages", line_range=...)

# 2. Find all paragraphs until next heading
current = heading
paras_to_delete = [heading]
while current.getnext() is not None:
    current = current.getnext()
    if is_heading(current):
        break
    paras_to_delete.append(current)

# 3. Wrap each in deletion
for para in paras_to_delete:
    deletion = create_deletion_wrapper(para)
    para.getparent().replace(para, deletion)

# 4. Update TOC manually
# ... complex TOC update code ...
```

```python
# NEW WAY (1 line)
doc.delete_section("Enhanced Damages Fail as a Matter of Law", track=True, update_toc=True)
```

### Compatibility

The new API is fully compatible with the existing low-level API:

```python
# Mix and match
doc = Document('file.docx')

# High-level
doc.insert_tracked(" text", after="target")

# Low-level for edge cases
editor = doc.get_editor("word/document.xml")
node = editor.get_node(tag="w:customTag", attrs={"custom": "value"})
editor.insert_after(node, '<custom:xml>...</custom:xml>')

# Save works for both
doc.save()
```

---

## Testing Strategy

### Unit Tests

```python
def test_insert_tracked_simple():
    doc = Document('test.docx')
    doc.insert_tracked(" added", after="target text")
    assert " added" in doc.get_text()

def test_insert_tracked_with_scope():
    doc = Document('test.docx')
    doc.insert_tracked(" added", after="target", scope="section:Argument")
    # Verify inserted in correct location, not in TOA

def test_text_fragmentation():
    # Document with fragmented text across runs
    doc = Document('fragmented.docx')
    doc.insert_tracked(" added", after="fragmented text")
    # Should work even though text is split

def test_disambiguation():
    doc = Document('test.docx')
    # Should raise AmbiguousTextError
    with pytest.raises(AmbiguousTextError):
        doc.insert_tracked(" added", after="common text")
```

### Integration Tests

```python
def test_surgical_edits_workflow():
    """Test full workflow from client feedback scenario."""
    doc = Document('client_comments.docx', author="Hancock, Parker")

    # Accept changes
    result = doc.accept_all_changes()
    assert result.insertions > 0

    # Apply edits
    doc.insert_tracked(" (interpreting IRPA)", after="(7th Cir. 2022)")
    doc.replace_tracked("records", "compiles", scope="It claims")
    doc.delete_section("Enhanced Damages")

    # Save and verify
    doc.save('output.docx')

    # Reopen and verify changes are tracked correctly
    doc2 = Document('output.docx')
    changes = doc2.list_tracked_changes()
    assert len(changes) > 0
    assert all(c.author == "Hancock, Parker" for c in changes)
```

---

## Performance Considerations

### Optimizations

1. **Lazy Loading**: Only parse XML when needed
2. **Caching**: Cache paragraph/section lookups
3. **Batch Processing**: Process multiple edits in single pass when possible

```python
# SLOW: Multiple passes
for edit in edits:
    doc.insert_tracked(edit['text'], after=edit['after'])

# FAST: Single pass
doc.apply_edits(edits)  # Optimizes internally
```

### Benchmarks

Target performance (vs. manual editing in Word):
- Simple insertion: <100ms per edit
- Complex replacement: <500ms per edit
- Section deletion: <1s per section
- Batch of 10 edits: <2s total

---

## Future Enhancements

### Phase 2 Features

1. **Template System**
   ```python
   template = EditTemplate.from_yaml('citation_template.yaml')
   doc.apply_template(template, case="Vrdolyak", parens="granting MTD")
   ```

2. **Smart Suggestions**
   ```python
   doc.suggest_edits()  # AI-powered suggestions
   ```

3. **Diff View**
   ```python
   doc.show_diff(original, edited)  # Visual diff
   ```

4. **Undo/Redo**
   ```python
   doc.undo()
   doc.redo()
   ```

5. **Merge Changes**
   ```python
   doc.merge(other_doc)  # Merge tracked changes from another doc
   ```

---

## Questions for Dev Team

1. **Framework preference**: Build on existing lxml code or new library?
2. **Testing strategy**: Unit tests only or also integration tests with real docs?
3. **YAML support**: Include pyyaml dependency or make optional?
4. **Error handling**: Strict (raise exceptions) or permissive (warnings + continue)?
5. **Validation**: Always validate or make optional for performance?
6. **Python version**: Target 3.10+ (match existing) or support older versions?

---

## Success Criteria

Implementation is complete when:

1. ✅ All examples in this doc work as written
2. ✅ 90% reduction in code for common operations vs. current API
3. ✅ No XML writing required for standard operations
4. ✅ Automatic handling of text fragmentation
5. ✅ Clear, actionable error messages
6. ✅ Passes all surgical edits from client feedback scenario
7. ✅ Performance within benchmarks above
8. ✅ Full test coverage (>90%)
9. ✅ Documentation with examples
10. ✅ Backward compatible with existing low-level API

---

**End of Specification**
