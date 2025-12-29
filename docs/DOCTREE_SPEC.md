# DocTree: DOCX Accessibility Layer Specification

> A semantic accessibility layer for Word documents, inspired by browser ARIA trees

**Version**: 0.1.0 (Draft)
**Date**: 2025-12-28
**Status**: Design Complete, Ready for Implementation

---

## Executive Summary

DocTree provides a structured, agent-friendly representation of Word documents that:

1. **Replaces pandoc/markitdown conversion** with a native YAML representation preserving document semantics
2. **Enables precise surgical edits** via stable element references (refs) instead of ambiguous text anchors
3. **Preserves tracked changes and comments** as first-class citizens for legal/contract workflows
4. **Scales to large documents** via progressive disclosure and chunking

This is analogous to how Chrome's accessibility layer converts DOM to ARIA tree snapshots for browser automation.

---

## Table of Contents

1. [Core Concepts](#1-core-concepts)
2. [Reference (Ref) System](#2-reference-ref-system)
3. [Semantic Roles](#3-semantic-roles)
4. [YAML Output Format](#4-yaml-output-format)
5. [Track Changes & Comments](#5-track-changes--comments)
6. [Python API](#6-python-api)
7. [Agent Workflow](#7-agent-workflow)
8. [Large Document Handling](#8-large-document-handling)
9. [Formatting Handling](#9-formatting-handling)
10. [Images and Embedded Objects](#10-images-and-embedded-objects)
11. [Cross-References and Bookmarks](#11-cross-references-and-bookmarks)
12. [Section Detection Algorithm](#12-section-detection-algorithm)
13. [Performance Targets](#13-performance-targets)
14. [Implementation Plan](#14-implementation-plan)

---

## 1. Core Concepts

### 1.1 Design Principles

| Principle | Description |
|-----------|-------------|
| **Stable Refs** | Every element gets a reference ID that survives edits |
| **Semantic Roles** | Elements classified by purpose (heading, paragraph, list) |
| **Progressive Disclosure** | Three verbosity levels (minimal/standard/full) |
| **LLM-Optimized** | YAML format fits in context windows |
| **Round-Trip Safe** | Refs map back to XML elements for editing |

### 1.2 Comparison to Browser Accessibility

| Browser (ARIA) | DocTree (DOCX) |
|----------------|----------------|
| DOM Element | Paragraph, Run, Table, Cell |
| ARIA Role | Semantic role (heading, listitem) |
| Accessible Name | Text content |
| Element Ref | `p:3`, `tbl:0/row:2/cell:1` |
| State `[checked]` | State `[tracked-insert]` |

### 1.3 Key Benefits

- **Unambiguous targeting**: `p:5` vs text that appears 3 times
- **Structural awareness**: Reference "3rd row of table 2"
- **Change tracking**: See insertions/deletions inline
- **Verification**: Compare trees before/after edits

---

## 2. Reference (Ref) System

### 2.1 Ref Format

```
ref = element_type ":" identifier ["/" sub_element]*

Examples:
  p:3                     # 4th paragraph (0-indexed)
  p:~xK4mNp2q             # Paragraph by fingerprint (stable)
  tbl:0/row:2/cell:1      # Table 0, row 2, cell 1
  tbl:0/row:2/cell:1/p:0  # First paragraph in that cell
  ins:42                  # Tracked insertion with id=42
  hdr:default/p:0         # First paragraph in default header
```

### 2.2 Element Type Prefixes

| Prefix | Element | Example |
|--------|---------|---------|
| `p` | Paragraph | `p:5`, `p:~xK4mNp2q` |
| `r` | Run (text span) | `p:3/r:2` |
| `tbl` | Table | `tbl:0` |
| `row` | Table row | `tbl:0/row:2` |
| `cell` | Table cell | `tbl:0/row:2/cell:1` |
| `ins` | Tracked insertion | `ins:42` |
| `del` | Tracked deletion | `del:17` |
| `hdr` | Header | `hdr:default` |
| `ftr` | Footer | `ftr:first` |
| `fn` | Footnote | `fn:3` |
| `cmt` | Comment | `cmt:5` |

### 2.3 Fingerprint-Based Refs (Stable)

For refs that survive document edits, use content fingerprints:

```python
# Fingerprint = hash of (text[:200] + style + parent_type)
p:~xK4mNp2q  # Stable even if paragraphs inserted before
```

**Stability guarantees:**
- Ordinal refs (`p:5`) shift when elements inserted/deleted
- Fingerprint refs (`p:~xK4mNp2q`) remain stable unless content changes

### 2.4 Ref Resolution

```python
# Resolve ref to XML element
element = doc.resolve_ref("p:5")
element = doc.resolve_ref("tbl:0/row:2/cell:1/p:0")

# Get ref for element
ref = doc.get_ref(match.span.paragraph)  # Returns "p:5"
```

---

## 3. Semantic Roles

### 3.1 Role Taxonomy

**Document Roles (Landmarks)**
| Role | Word Source | HTML Equivalent |
|------|-------------|-----------------|
| `document` | Root | `role="document"` |
| `header` | `w:hdr` | `role="banner"` |
| `footer` | `w:ftr` | `role="contentinfo"` |
| `section` | Heading groups | `<section>` |

**Structural Roles**
| Role | Word Source | HTML Equivalent |
|------|-------------|-----------------|
| `heading` | Heading styles | `<h1>`-`<h6>` |
| `paragraph` | `w:p` | `<p>` |
| `blockquote` | Quote styles | `<blockquote>` |
| `list` | `w:numPr` | `<ul>`, `<ol>` |
| `listitem` | List paragraphs | `<li>` |
| `table` | `w:tbl` | `<table>` |
| `row` | `w:tr` | `<tr>` |
| `cell` | `w:tc` | `<td>` |

**Inline Roles**
| Role | Word Source | HTML Equivalent |
|------|-------------|-----------------|
| `text` | `w:t` | Text node |
| `strong` | `w:b` | `<strong>` |
| `emphasis` | `w:i` | `<em>` |
| `link` | `w:hyperlink` | `<a>` |

**Annotation Roles**
| Role | Word Source | HTML Equivalent |
|------|-------------|-----------------|
| `insertion` | `w:ins` | `<ins>` |
| `deletion` | `w:del` | `<del>` |
| `comment` | `w:comment` | N/A |

### 3.2 Style Mapping

```python
STYLE_TO_ROLE = {
    "Title": {"role": "heading", "level": 1},
    "Heading1": {"role": "heading", "level": 1},
    "Heading2": {"role": "heading", "level": 2},
    "Quote": {"role": "blockquote"},
    "Normal": {"role": "paragraph"},
    "ListParagraph": {"role": "listitem"},
}
```

---

## 4. YAML Output Format

### 4.1 Verbosity Levels

**Minimal** - Structure overview, navigation
```yaml
- h1 "SERVICES AGREEMENT" [ref=p:0]
- p "This Agreement is entered into..." [ref=p:1]
- table [ref=tbl:0] [4x3]: "Payment Terms"
```

**Standard** (default) - Full content, edit planning
```yaml
- heading [ref=p:0] [level=1]:
    text: "SERVICES AGREEMENT"
    style: Heading1

- paragraph [ref=p:1]:
    text: "This Agreement is entered into as of..."

- table [ref=tbl:0] [rows=4] [cols=3]:
    - row [ref=tbl:0/row:0] [header]:
        - cell: "Service"
        - cell: "Rate"
        - cell: "SLA"
```

**Full** - Complete fidelity, run-level detail
```yaml
- heading [ref=p:0] [level=1]:
    style: Heading1
    runs:
      - text "SERVICES AGREEMENT" [ref=p:0/r:0] [bold] [caps]

- paragraph [ref=p:1]:
    style: Normal
    runs:
      - text "This " [ref=p:1/r:0]
      - text "Agreement" [ref=p:1/r:1] [bold]
      - text " is entered into..." [ref=p:1/r:2]
```

### 4.2 State Conventions

**States (visual/semantic) - Use `[brackets]`**
- `[bold]` `[italic]` `[underline]`
- `[tracked-insert]` `[tracked-delete]`
- `[header]` - table header row
- `[level=2]` - heading level

**Properties (metadata) - Use `property:` syntax**
```yaml
author: "Jane Doe"
date: "2025-12-28T10:30:00Z"
style: Heading1
```

### 4.3 Document Structure

```yaml
document:
  path: "contract.docx"
  verbosity: standard
  stats:
    paragraphs: 47
    tables: 2
    tracked_changes: 12
    comments: 5

content:
  - heading [ref=p:0] [level=1]: "SERVICES AGREEMENT"
  - paragraph [ref=p:1]: "This Agreement..."
  - table [ref=tbl:0]:
      # ... table content

tracked_changes:
  - ref: change:0
    type: insertion
    author: "Jane Doe"
    text: "quarterly"
    location: p:3

comments:
  - ref: cmt:0
    author: "Legal Team"
    text: "Please review"
    on_text: "indemnification"
```

---

## 5. Track Changes & Comments

### 5.1 View Modes

| Mode | Description |
|------|-------------|
| `final` | Show as if changes accepted |
| `original` | Show as if changes rejected |
| `markup` | Show all changes with markers |

### 5.2 Inline Change Markers

```yaml
- paragraph [ref=p:3]:
    text: "Payment in {--30--}{++45++} days"
    has_changes: true
    change_refs: [change:0, change:1]
```

### 5.3 Change Metadata

```yaml
tracked_changes:
  - ref: change:0
    type: deletion
    id: "3"
    author: "Jane Doe"
    date: "2025-12-28T10:15:00Z"
    text: "30"
    location:
      paragraph_ref: p:3

  - ref: change:1
    type: insertion
    id: "4"
    author: "Jane Doe"
    date: "2025-12-28T10:15:05Z"
    text: "45"
    location:
      paragraph_ref: p:3
```

### 5.4 Comment Threading

```yaml
comments:
  - ref: cmt:0
    author: "Legal Team"
    text: "Should we expand this?"
    on_text: "indemnification clause"
    resolved: false
    replies:
      - ref: cmt:1
        author: "Parker"
        text: "Yes, I've drafted additions"
```

---

## 6. Python API

### 6.1 Core Types

```python
from dataclasses import dataclass
from enum import Enum

class ElementType(Enum):
    PARAGRAPH = auto()
    RUN = auto()
    TABLE = auto()
    TABLE_ROW = auto()
    TABLE_CELL = auto()
    # ...

@dataclass
class Ref:
    path: str  # e.g., "p:5" or "tbl:0/row:2/cell:1"

@dataclass
class AccessibilityNode:
    ref: Ref
    element_type: ElementType
    text: str = ""
    children: list["AccessibilityNode"] = field(default_factory=list)
    style: str | None = None
    change: ChangeInfo | None = None
    comments: list[CommentInfo] = field(default_factory=list)

@dataclass
class ViewMode:
    include_body: bool = True
    include_headers: bool = False
    include_comments: bool = False
    include_tracked_changes: bool = True
    include_formatting: bool = False
```

### 6.2 Document Methods

```python
class Document:
    @property
    def accessibility_tree(self) -> AccessibilityTree:
        """Get accessibility tree (cached, lazy)."""

    def get_accessibility_tree(
        self,
        view_mode: ViewMode | None = None,
    ) -> AccessibilityTree:
        """Get tree with custom view mode."""

    def resolve_ref(self, ref: str | Ref) -> etree._Element:
        """Resolve ref to XML element."""

    def get_ref(self, element: Match | TextSpan | Element) -> Ref | None:
        """Get ref for element."""
```

### 6.3 Ref-Based Editing

```python
class Document:
    def insert_at_ref(
        self,
        ref: str,
        text: str,
        position: Literal["before", "after", "start", "end"] = "after",
        track: bool = False,
    ) -> EditResult:
        """Insert text at ref position."""

    def delete_ref(
        self,
        ref: str,
        track: bool = False,
    ) -> EditResult:
        """Delete element at ref."""

    def replace_at_ref(
        self,
        ref: str,
        new_text: str,
        track: bool = False,
    ) -> EditResult:
        """Replace content at ref."""

    def add_comment_at_ref(
        self,
        ref: str,
        comment_text: str,
    ) -> Comment:
        """Add comment on element at ref."""
```

### 6.4 Error Handling

```python
class RefNotFoundError(DocxRedlineError):
    """Ref cannot be resolved."""

class StaleRefError(DocxRedlineError):
    """Ref points to deleted element."""
```

---

## 7. Agent Workflow

### 7.1 Ref-Based Workflow

```
1. Load document
2. Generate accessibility tree (YAML with refs)
3. Analyze structure semantically
4. Make precise edits using refs
5. Verify changes by comparing trees
6. Save document
```

### 7.2 Example Usage

```python
from python_docx_redline import Document

# 1. Load and get tree
doc = Document("contract.docx")
tree = doc.accessibility_tree
print(tree.to_yaml())

# 2. Find element of interest
headings = tree.find_all(heading_level=1)
payment_section = next(h for h in headings if "Payment" in h.text)

# 3. Edit by ref
doc.insert_at_ref(payment_section.ref, " (Amended)", position="end", track=True)

# 4. Verify
updated = doc.accessibility_tree.to_yaml()
```

### 7.3 Text Search + Ref Editing (Hybrid)

```python
# Use text search to find, ref to edit
matches = doc.find_all("30 days")
target_ref = doc.get_ref(matches[0])
doc.replace_at_ref(target_ref, "45 days", track=True)
```

### 7.4 Error Recovery

```python
try:
    doc.replace_at_ref("p:5", "new text")
except RefNotFoundError:
    # Tree is stale, regenerate
    tree = doc.accessibility_tree  # Forces regeneration
    # Find element again
    matches = tree.find_all(text_contains="old text")
```

---

## 8. Large Document Handling

### 8.1 The Problem

Large documents (100+ pages) can generate trees exceeding 50,000+ tokens, flooding context windows.

| Document Size | Paragraphs | Est. Tokens (Full) | Risk |
|---------------|------------|-------------------|------|
| 10 pages | ~50 | ~5,000 | Safe |
| 50 pages | ~250 | ~25,000 | Moderate |
| 100 pages | ~500 | ~50,000 | High |
| 200+ pages | ~1000+ | ~100,000+ | Overflow |

### 8.2 Solution: Two-Tier Progressive Disclosure

**Tier 1: Outline Mode (Always Safe)**
```python
# Default for any document - ~2000 tokens max
tree = doc.get_accessibility_tree()  # Returns outline
```

**Tier 2: Expand on Demand**
```python
# Expand specific section
section = doc.expand_section("sec:2")

# Expand specific refs
content = doc.expand_refs(["p:45", "p:46", "tbl:2"])

# Get table with pagination
table = doc.get_table("tbl:0", max_rows=20)
```

### 8.3 Outline Mode Output

```yaml
document:
  path: "enterprise_msa.docx"
  mode: outline
  stats:
    total_paragraphs: 523
    total_sections: 24
    tracked_changes: 47

outline:
  - section [ref=sec:1]:
      heading: "1. DEFINITIONS"
      heading_ref: p:5
      paragraph_count: 32
      tracked_changes: 5
      preview: "In this Agreement, unless the context..."

  - section [ref=sec:2]:
      heading: "2. SERVICES"
      paragraph_count: 45
      tables:
        - table [ref=tbl:0] [4x3]: "Service Categories"

navigation:
  expand_section: "doc.expand_section('sec:2')"
  search: "doc.search('indemnification')"
```

### 8.4 Search-First API

Find refs without loading full tree:

```python
# Search returns refs + context, not full tree
results = doc.search("indemnification", max_results=20)
for r in results:
    print(f"{r.ref}: {r.context}")
# p:156: ...The Provider's total indemnification liability...
```

### 8.5 Token Budgeting

```python
# Request tree with token budget
tree = doc.get_accessibility_tree(max_tokens=5000)
# Intelligently truncates, prioritizing structure over content
```

### 8.6 API Methods

```python
class Document:
    def get_accessibility_tree(
        self,
        mode: Literal["outline", "standard", "full"] = "outline",
        max_tokens: int | None = None,
    ) -> AccessibilityTree:
        """Default returns outline for safety."""

    def expand_section(self, section_ref: str) -> SectionTree:
        """Expand single section to full content."""

    def expand_refs(self, refs: list[str]) -> RefTree:
        """Expand specific refs only."""

    def get_table(self, table_ref: str, max_rows: int = None) -> TableTree:
        """Get table with optional pagination."""

    def search(self, pattern: str, max_results: int = 20) -> SearchResults:
        """Search without loading tree."""
```

---

## 9. Content vs Styling Modes

### 9.1 Two Editing Modes

Documents have two distinct editing concerns:

| Mode | Purpose | Information Needed | Token Cost |
|------|---------|-------------------|------------|
| **Content** | Edit text, restructure | Text, paragraphs, sections | Lower |
| **Styling** | Format, polish, finalize | + Run-level formatting | Higher |

### 9.2 Content Mode (Default)

For drafting, reviewing, restructuring - most agent tasks:

```yaml
# Content mode output
- paragraph [ref=p:3]:
    text: "The term Services means professional consulting..."
    style: Normal  # Style NAME only (semantic)
    has_changes: true
    change_refs: [change:5]

- table [ref=tbl:0] [3x4]:
    - row [header]: ["Term", "Definition", "Section"]
    - row: ["Services", "Professional consulting...", "1.1"]
```

**Content editing - formatting inherited automatically:**
```python
doc.replace_at_ref("p:3", "New clause text")  # Keeps existing formatting
doc.insert_at_ref("p:3", " (amended)", position="end")  # Inherits from anchor
```

### 9.3 Styling Mode

For final polish, brand consistency, presentation:

```yaml
# Styling mode output - same paragraph
- paragraph [ref=p:3]:
    text: "The term Services means professional consulting..."
    style: Normal
    formatting:
      font: "Times New Roman"
      size: 12pt
      spacing_after: 8pt
    runs:
      - text "The term " [ref=p:3/r:0]
      - text "Services" [ref=p:3/r:1]:
          bold: true
          font: "Times New Roman"
          size: 12pt
      - text " means professional consulting..." [ref=p:3/r:2]
```

**Styling editing - explicit control:**
```python
doc.format_at_ref("p:3", style="Heading2")  # Change paragraph style
doc.format_at_ref("p:3/r:1", bold=True, color="#0000FF")  # Format specific run
doc.apply_style_to_refs(["p:5", "p:10"], style="BodyText")  # Batch styling
```

### 9.4 Typical Agent Workflow

**Read wide in content mode, zoom in with styling mode:**

```python
# 1. Understand document structure (content mode, low tokens)
tree = doc.get_accessibility_tree(mode="content")
# Agent sees: sections, paragraphs, tables, track changes

# 2. Make content edits
doc.replace_at_ref("p:45", "Updated liability clause")
doc.insert_at_ref("sec:3", "Additional terms paragraph...")

# 3. Zoom in for polish (styling mode, specific refs only)
styling = doc.expand_refs(["p:45", "p:46"], mode="styling")
# Now agent sees run-level formatting for just those paragraphs

# 4. Apply formatting fixes
doc.format_at_ref("p:45/r:2", bold=True)  # Emphasize defined term
doc.format_at_ref("p:46", style="BodyTextIndent")
```

### 9.5 API

```python
class Document:
    def get_accessibility_tree(
        self,
        mode: Literal["content", "styling"] = "content",
        # ... other params
    ) -> AccessibilityTree:
        """Get tree in content or styling mode."""

    def expand_refs(
        self,
        refs: list[str],
        mode: Literal["content", "styling"] = "content",
    ) -> RefTree:
        """Expand specific refs - can use styling mode for targeted polish."""

    def expand_section(
        self,
        section_ref: str,
        mode: Literal["content", "styling"] = "content",
    ) -> SectionTree:
        """Expand section - typically content mode."""
```

### 9.6 Style vs Direct Formatting

Word has two formatting layers:

| Layer | What It Is | Example |
|-------|------------|---------|
| **Style** | Named semantic format | `Heading1`, `BodyText`, `Quote` |
| **Direct** | Override on specific text | This word is **bold** |

In styling mode, both are visible:

```yaml
- paragraph [ref=p:3]:
    style: Normal  # Paragraph style
    runs:
      - text "Important" [ref=p:3/r:0]:
          bold: true  # Direct formatting override
```

### 9.7 Tracked Format Changes

When track changes is on, formatting changes are tracked:

```yaml
- text "important" [ref=p:3/r:1] [bold] [tracked-format]:
    format_change:
      author: "Editor"
      date: "2025-12-28T10:00:00Z"
      before: {bold: false}
      after: {bold: true}
```

### 9.8 Insert Formatting Behavior

| Scenario | Behavior |
|----------|----------|
| Insert in content mode | Inherit formatting from anchor |
| Insert in styling mode | Inherit by default, can override |
| Insert with explicit formatting | Always uses specified formatting |

```python
# Inherit (default)
doc.insert_at_ref("p:3/r:1", " additional", position="after")
# New text inherits bold from r:1

# Explicit override
doc.insert_at_ref("p:3/r:1", " (note)", position="after", bold=False)
# New text is NOT bold even though anchor is
```

---

## 10. Images and Embedded Objects

### 10.1 Supported Image Types

| Container | Namespace | Contents | Position Type |
|-----------|-----------|----------|---------------|
| `w:drawing` | WML | Modern DrawingML graphics | Inline or floating |
| `w:pict` | WML | Legacy VML graphics | Inline only |
| `w:object` | WML | OLE objects (Excel, PDF) | Inline |

### 10.2 Image Ref Formats

| Prefix | Element | Example |
|--------|---------|---------|
| `img` | Image (inline) | `img:5/0` |
| `img/f` | Image (floating) | `img:5/f:0` |
| `chart` | Chart | `chart:12/0` |
| `diagram` | SmartArt | `diagram:8/0` |
| `shape` | Shape | `shape:3/0` |
| `obj` | OLE Object | `obj:15/0` |
| `vml` | Legacy VML | `vml:20/0` |

Compound refs for special contexts:
```
hdr:default/img:0/0      # Image in default header
tbl:0/row:1/cell:2/img:0/0  # Image in table cell
fn:3/img:0/0             # Image in footnote
```

### 10.3 Image YAML Representation

**Content mode (default):**
```yaml
- paragraph [ref=p:5]:
    text: "See the company logo below:"
    images:
      - image [ref=img:5/0] [inline]:
          name: "Company Logo"
          alt_text: "Acme Corp red and blue logo"
          size: "2.5in x 1.0in"
```

**Styling mode:**
```yaml
- paragraph [ref=p:5]:
    images:
      - image [ref=img:5/0] [inline]:
          name: "Company Logo"
          alt_text: "Acme Corp red and blue logo"
          size:
            width_emu: 2286000
            height_emu: 914400
          format: "png"
          relationship_id: "rId7"
```

### 10.4 Floating Images

```yaml
- paragraph [ref=p:6]:
    floating_images:
      - image [ref=img:6/f:0] [floating]:
          name: "Organization Chart"
          alt_text: "Team structure diagram"
          position:
            horizontal: "center"
            vertical: "top"
            relative_to: "page"
          wrap: "topAndBottom"
```

### 10.5 Charts and SmartArt

```yaml
- chart [ref=chart:12/0]:
    name: "Q4 Revenue"
    alt_text: "Bar chart showing quarterly revenue"
    type: "bar"
    size: "5.0in x 3.0in"
    data_source: "embedded"

- diagram [ref=diagram:8/0]:
    name: "Process Flow"
    alt_text: "5-step approval process"
    diagram_type: "process"
    text_content: ["Step 1", "Step 2", "Step 3"]
```

### 10.6 Image API Methods

```python
class Document:
    def get_images(self, scope: str | None = None) -> list[ImageInfo]:
        """Get all images in document or scope."""

    def insert_image_at_ref(
        self,
        ref: str,
        image_path: str | Path,
        position: Literal["before", "after"] = "after",
        width_inches: float | None = None,
        description: str = "",
        track: bool = False,
    ) -> str:
        """Insert image at ref, returns new image ref."""

    def update_image_alt_text(self, ref: str, alt_text: str) -> bool:
        """Update alt text for accessibility."""

    def resize_image_at_ref(
        self,
        ref: str,
        width_inches: float | None = None,
        height_inches: float | None = None,
        maintain_aspect: bool = True,
    ) -> bool:
        """Resize existing image."""
```

---

## 11. Cross-References and Bookmarks

### 11.1 Element Types

| Prefix | Element | Description |
|--------|---------|-------------|
| `bk` | Bookmark | Named location in document |
| `lnk` | Hyperlink | Link (internal or external) |
| `xref` | Cross-reference | Field-based reference |
| `fn` | Footnote | Footnote reference |
| `en` | Endnote | Endnote reference |
| `toc` | TOC Entry | Table of Contents entry |

### 11.2 Bookmark YAML

```yaml
bookmarks:
  - ref: bk:DefinitionsSection
    name: "DefinitionsSection"
    location: p:5
    text_preview: "1. DEFINITIONS"
    referenced_by:
      - lnk:0           # Hyperlink in TOC
      - xref:3          # Cross-reference in body
```

### 11.3 Hyperlink YAML

**In content:**
```yaml
- paragraph [ref=p:3]:
    text: "See Section 2.1 for details."
    links:
      - ref: lnk:0
        text: "Section 2.1"
        target: bk:PaymentTerms
        target_location: p:15
```

**Link summary:**
```yaml
links:
  internal:
    - ref: lnk:0
      from: p:3
      to: bk:PaymentTerms

  external:
    - ref: lnk:1
      from: p:7
      url: "https://example.com"

  broken:
    - ref: lnk:5
      from: p:20
      target: bk:DeletedSection
      error: "Bookmark not found"
```

### 11.4 Bidirectional Reference Tracking

When querying a location, include what references it:

```yaml
- heading [ref=p:15] [level=2]:
    text: "2.1 Payment Terms"
    bookmark: bk:PaymentTerms
    incoming_references:
      - lnk:0              # Link from TOC
      - lnk:3              # Link from body
      - xref:0             # Cross-reference field
```

### 11.5 Bookmark API

```python
class Document:
    @property
    def bookmarks(self) -> list[Bookmark]:
        """Get all bookmarks."""

    def add_bookmark(
        self,
        name: str,
        at_ref: str,
        span_to_ref: str | None = None,
    ) -> Bookmark:
        """Add bookmark at location."""

    def rename_bookmark(
        self,
        old_name: str,
        new_name: str,
        update_references: bool = True,
    ) -> int:
        """Rename bookmark, optionally updating all references."""

    def validate_references(self) -> ReferenceValidationResult:
        """Check for broken links and orphan bookmarks."""
```

---

## 12. Section Detection Algorithm

### 12.1 Problem

Many documents lack proper heading styles:
- Letters and memos
- Plain text imports
- Documents using bold/caps instead of styles
- Documents with numbered sections but no styles

### 12.2 Detection Tiers

```
Tier 1: Heading styles (Heading1, Heading2...) - HIGH confidence
Tier 2: Outline level property (w:outlineLvl) - HIGH confidence
Tier 3: Heuristics - MEDIUM confidence
  - All-bold short paragraphs
  - ALL CAPS short paragraphs
  - Numbered sections (1., 2., Article I)
  - Blank line separators
Tier 4: Fallback (single section or chunking) - LOW confidence
```

### 12.3 Heuristic Configuration

```python
@dataclass
class HeuristicConfig:
    detect_bold_first_line: bool = True
    detect_caps_first_line: bool = True
    detect_numbered_sections: bool = True
    detect_blank_line_breaks: bool = True

    min_section_paragraphs: int = 2
    max_heading_length: int = 100

    numbering_patterns: list[str] = field(default_factory=lambda: [
        r"^\d+\.\s",           # 1. Section
        r"^\d+\.\d+\s",        # 1.1 Section
        r"^Article\s+\d+",     # Article 1
        r"^Section\s+\d+",     # Section 1
    ])
```

### 12.4 YAML Output with Detection Metadata

```yaml
document:
  section_detection:
    method: "heuristic"
    confidence: "medium"

outline:
  - section [ref=sec:0] [inferred]:
      detection: bold_heuristic
      confidence: medium
      heading: "INTRODUCTION"
      paragraph_count: 5

  - section [ref=sec:1] [explicit]:
      detection: heading_style
      confidence: high
      heading: "1. Background"
      heading_level: 1
```

### 12.5 API

```python
class Document:
    def get_accessibility_tree(
        self,
        section_detection: SectionDetectionConfig | None = None,
    ) -> AccessibilityTree:
        """Configure section detection behavior."""
```

---

## 13. Performance Targets

### 13.1 Tree Generation

| Operation | 100 pages | 500 pages | Token Budget |
|-----------|-----------|-----------|--------------|
| Outline Mode | <100ms | <300ms | ~2,000 |
| Content Mode | <300ms | <1,000ms | ~40,000 |
| Styling Mode | <500ms | N/A (section only) | ~150,000 |

### 13.2 Operations

| Operation | Target | Acceptable |
|-----------|--------|------------|
| Ref resolution (warm) | <2ms | <5ms |
| Ref resolution (cold) | <5ms | <10ms |
| Search (literal) | <50ms | <100ms |
| Insert at ref | <20ms | <50ms |
| Batch 10 edits | <200ms | <400ms |

### 13.3 Memory Usage

| Document Size | Pages | Target | Limit |
|--------------|-------|--------|-------|
| Small | <20 | <10 MB | <20 MB |
| Medium | 20-100 | <30 MB | <50 MB |
| Large | 100-300 | <75 MB | <150 MB |
| Very Large | 300-500 | <150 MB | <300 MB |

### 13.4 Caching Strategy

```
Level 0: Raw XML Elements (lxml tree)
Level 1: Ref Registry (ordinal -> element)
Level 2: Fingerprint Index (hash -> ordinal)
Level 3: Computed Properties (LRU cached)
Level 4: Tree Snapshots (optional)
```

### 13.5 Degradation Tiers

| Document Size | Behavior |
|--------------|----------|
| <100 pages | Full content mode |
| 100-300 pages | Content mode with warning |
| 300-500 pages | Outline mode default |
| 500+ pages | Outline mode only, section expansion |

### 13.6 Key Optimizations

1. **Outline mode O(sections) not O(paragraphs)**: Single pass scanning for headings only
2. **Lazy loading**: Only parse document parts when accessed
3. **LRU caching**: Bound cache sizes with `maxsize`
4. **Generator-based iteration**: Avoid materializing full lists
5. **Token budgeting**: `max_tokens` parameter for intelligent truncation

---

## 14. Implementation Plan

### Phase 1: Core Infrastructure
- [ ] `RefRegistry` class with fingerprint generation
- [ ] `AccessibilityTree` class with tree building
- [ ] `AccessibilityNode` dataclass
- [ ] Integration with Document class

### Phase 2: YAML Serialization
- [ ] Three verbosity levels
- [ ] State and property formatting
- [ ] Tracked changes representation
- [ ] Comments representation

### Phase 3: Ref-Based Editing
- [ ] `resolve_ref()` method
- [ ] `insert_at_ref()` method
- [ ] `delete_ref()` method
- [ ] `replace_at_ref()` method

### Phase 4: Agent Integration
- [ ] Update docx skill documentation
- [ ] Error handling and recovery patterns
- [ ] Verification helpers
- [ ] Performance optimization

---

## Critical Implementation Files

| File | Purpose |
|------|---------|
| `document.py` | Core Document class integration |
| `accessibility/tree.py` | AccessibilityTree class |
| `accessibility/resolver.py` | RefResolver class |
| `accessibility/types.py` | Core types (Ref, AccessibilityNode) |
| `match.py` | Add ref property to Match |
| `errors.py` | RefNotFoundError, StaleRefError |

---

## Appendix: Full Example

```yaml
# Contract accessibility snapshot
document:
  path: "service_agreement_v3.docx"
  verbosity: standard
  stats:
    paragraphs: 47
    tracked_changes: 12
    comments: 5

content:
  - heading [ref=p:0] [level=1]:
      text: "MASTER SERVICE AGREEMENT"
      style: Title

  - paragraph [ref=p:1]:
      text: "This Agreement is entered into as of {--January 1--}{++March 15++}, 2025..."
      has_changes: true

  - heading [ref=p:2] [level=2]:
      text: "1. DEFINITIONS"

  - paragraph [ref=p:3]:
      text: "\"Services\" means the professional consulting services..."
      has_comments: true
      comment_refs: [cmt:0]

  - table [ref=tbl:0] [rows=3] [cols=4]:
      - row [ref=tbl:0/row:0] [header]:
          - cell: "Service"
          - cell: "Rate"
          - cell: "SLA"
          - cell: "Notes"
      - row [ref=tbl:0/row:1]:
          - cell: "Consulting"
          - cell: "$200/hr"
          - cell: "24hr response"
          - cell: ""

tracked_changes:
  - ref: change:0
    type: deletion
    author: "Legal Team"
    text: "January 1"
    location: p:1

  - ref: change:1
    type: insertion
    author: "Legal Team"
    text: "March 15"
    location: p:1

comments:
  - ref: cmt:0
    author: "Reviewer"
    text: "Verify exhibit A is attached"
    on_text: "consulting services"
    resolved: false
```
