# DocTree: Document Accessibility Layer

The DocTree accessibility layer provides a semantic, structured view of Word documents that enables:
- **Text/Markdown export** replacing pandoc dependency
- **LLM-friendly YAML output** for document understanding
- **Stable element references** for precise editing
- **Progressive disclosure** for handling large documents
- **Tracked changes and comments** as first-class citizens

This is analogous to how browsers provide ARIA trees for accessibility tools.

## Quick Start

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Plain text or markdown (replaces pandoc)
text = doc.to_text()
md = doc.to_markdown()

# With tracked changes visible
md = doc.to_markdown(tracked_changes="all")
# Output: "Payment in {--30--}{++45++} days"
```

## Text and Markdown Export

Export documents to plain text or markdown with a single line:

```python
# Plain text
text = doc.to_text()

# Markdown
md = doc.to_markdown()

# Options
md = doc.to_markdown(
    tracked_changes="all",      # "accept" | "reject" | "all"
    include_comments=True,      # Include comment annotations
    table_format="markdown",    # "markdown" | "simple" | "grid"
)
```

### Tracked Changes Output

```python
# Accept (default) - final state
doc.to_text()  # "Payment in 45 days"

# Reject - original state
doc.to_text(tracked_changes="reject")  # "Payment in 30 days"

# All - CriticMarkup syntax
doc.to_text(tracked_changes="all")  # "Payment in {--30--}{++45++} days"
```

## YAML Output (for Agent Workflows)

For structured output with element refs, use the AccessibilityTree directly:

```python
from python_docx_redline.accessibility import AccessibilityTree

tree = AccessibilityTree.from_document(doc)
print(tree.to_yaml())
```

**Output:**
```yaml
document:
  stats:
    paragraphs: 45
    tables: 2
    tracked_changes: 8
    comments: 3
  body:
    - p [ref=p:0] [heading]: "SERVICES AGREEMENT"
    - p [ref=p:1]: "This Agreement is entered into..."
    - p [ref=p:2] [heading]: "1. Definitions"
    - p [ref=p:3]: 'The term "Services" means...'
    - tbl [ref=tbl:0]:
        rows: 5
        cols: 3
        header: ["Term", "Definition", "Section"]
```

## Verbosity Levels

Three levels control output detail:

| Level | Use Case | Token Usage |
|-------|----------|-------------|
| `minimal` | Document overview, structure scanning | ~500 tokens/page |
| `standard` | Normal editing tasks | ~1,500 tokens/page |
| `full` | Detailed formatting work | ~3,000 tokens/page |

```python
# Minimal - just structure
print(tree.to_yaml(verbosity="minimal"))

# Standard - balanced (default)
print(tree.to_yaml(verbosity="standard"))

# Full - everything including formatting
print(tree.to_yaml(verbosity="full"))
```

**Minimal output:**
```yaml
document:
  body:
    - p [ref=p:0]: "SERVICES AGREEMENT"
    - p [ref=p:1]: "This Agreement is entered..."
```

**Full output (includes formatting):**
```yaml
document:
  body:
    - p [ref=p:0]:
        text: "SERVICES AGREEMENT"
        style: "Heading 1"
        formatting:
          bold: true
          font_size: 14pt
          alignment: center
```

## Reference (Ref) System

Every element has a stable reference for unambiguous targeting:

```
p:5                     # 5th paragraph (0-indexed)
p:~xK4mNp2q             # Paragraph by fingerprint (stable across edits)
tbl:0/row:2/cell:1      # Table 0, row 2, cell 1
tbl:0/row:2/cell:1/p:0  # First paragraph in that cell
ins:42                  # Tracked insertion with id=42
hdr:default/p:0         # First paragraph in default header
```

### Ref Types

| Prefix | Element | Example |
|--------|---------|---------|
| `p` | Paragraph | `p:5`, `p:~xK4mNp2q` |
| `tbl` | Table | `tbl:0` |
| `row` | Table row | `tbl:0/row:2` |
| `cell` | Table cell | `tbl:0/row:2/cell:1` |
| `ins` | Tracked insertion | `ins:42` |
| `del` | Tracked deletion | `del:17` |
| `hdr` | Header | `hdr:default` |
| `ftr` | Footer | `ftr:first` |
| `img` | Image | `img:5/0` |
| `cmt` | Comment | `cmt:5` |

### Ordinal vs Fingerprint Refs

```python
# Ordinal refs shift when elements are added/removed
p:5  # Was 5th paragraph, now might be 6th

# Fingerprint refs are stable (content-based hash)
p:~xK4mNp2q  # Same paragraph even after insertions
```

### Resolving Refs

```python
from python_docx_redline.accessibility import RefRegistry

registry = tree.registry

# Resolve ref to XML element
element = registry.resolve_ref("p:5")
element = registry.resolve_ref("tbl:0/row:2/cell:1")

# Get cache statistics (for performance tuning)
stats = registry.cache_stats
print(f"Cache hits: {stats.hits}, misses: {stats.misses}")
```

## ViewMode Configuration

Control what's included in the tree:

```python
from python_docx_redline.accessibility import ViewMode

# Default: body with tracked changes
mode = ViewMode()

# Include everything
mode = ViewMode(
    include_body=True,
    include_headers=True,
    include_footers=True,
    include_footnotes=True,
    include_endnotes=True,
    include_comments=True,
    include_tracked_changes=True,
    include_formatting=True,
    verbosity="full"
)

tree = AccessibilityTree.from_document(doc, view_mode=mode)
```

## Tracked Changes in YAML

Tracked changes appear inline with annotations:

```yaml
body:
  - p [ref=p:3]:
      text: "Payment due in 30 days"
      changes:
        - del [ref=del:17]: "30"
          author: "Legal Team"
          date: "2024-01-15"
        - ins [ref=ins:42]: "45"
          author: "Legal Team"
          date: "2024-01-15"
```

## Images and Objects

```python
# Get all images
images = tree.get_images()
for img in images:
    print(f"{img.ref}: {img.image_type.name} at {img.position}")

# Get specific image by ref
img = tree.get_image("img:5/0")
print(f"Size: {img.size.width_emu}x{img.size.height_emu}")
```

**YAML output:**
```yaml
- p [ref=p:5]:
    text: "See diagram below:"
    images:
      - img [ref=img:5/0]:
          type: inline
          width: 3in
          height: 2in
          alt_text: "Architecture diagram"
```

## Bookmarks and Cross-References

```python
# Get all bookmarks
for name, bookmark in tree.bookmarks.items():
    print(f"{name}: {bookmark.ref}")

# Get specific bookmark
bookmark = tree.get_bookmark("Section2")
print(f"Located at: {bookmark.ref}")

# Validate all references (find broken links)
result = tree.validate_references()
if result.broken_links:
    print("Broken links found:", result.broken_links)
if result.orphan_bookmarks:
    print("Unused bookmarks:", result.orphan_bookmarks)
```

## Large Document Handling

For documents exceeding context limits, use OutlineTree:

```python
from python_docx_redline.accessibility import OutlineTree

# Get outline view (section headings only)
outline = OutlineTree.from_document(doc)
print(outline.to_yaml())
```

**Outline output:**
```yaml
outline:
  document_size:
    paragraphs: 450
    tables: 15
    estimated_tokens: 85000
  sections:
    - section [ref=sec:0]:
        heading: "1. Introduction"
        paragraphs: 12
        tables: 0
    - section [ref=sec:1]:
        heading: "2. Terms and Conditions"
        paragraphs: 45
        tables: 3
    - section [ref=sec:2]:
        heading: "3. Payment Terms"
        paragraphs: 28
        tables: 1
```

### Progressive Disclosure

```python
# Expand a specific section
section = outline.expand_section("sec:1")
print(section.to_yaml())

# Expand specific refs
refs = outline.expand_refs(["p:15", "p:16", "p:17"])
print(refs.to_yaml())

# Get a specific table with pagination
table = outline.get_table("tbl:0", page=1, page_size=20)
print(table.to_yaml())

# Search within the document
results = outline.search("payment terms", max_results=10)
print(results.to_yaml())
```

### Token Budgeting

```python
from python_docx_redline.accessibility import estimate_tokens, truncate_to_token_budget

# Estimate tokens for content
tokens = estimate_tokens(tree.to_yaml())

# Truncate to fit budget
truncated = truncate_to_token_budget(tree.to_yaml(), max_tokens=4000)
```

## Section Detection

The library uses a 4-tier cascade algorithm to detect document sections:

1. **Heading styles** (Heading 1, Heading 2, etc.) - HIGH confidence
2. **Outline levels** (w:outlineLvl in style definition) - HIGH confidence
3. **Heuristics** (bold, caps, numbering patterns) - MEDIUM confidence
4. **Fallback** (single section or fixed chunks) - LOW confidence

```python
from python_docx_redline.accessibility import (
    SectionDetector,
    SectionDetectionConfig,
    HeuristicConfig
)

# Configure detection
config = SectionDetectionConfig(
    use_heading_styles=True,
    use_outline_levels=True,
    use_heuristics=True,
    heuristic_config=HeuristicConfig(
        detect_bold_headings=True,
        detect_caps_headings=True,
        detect_numbered_sections=True,
        min_section_paragraphs=2,
        max_heading_length=100
    )
)

# Detect sections
detector = SectionDetector(config)
sections = detector.detect(doc.body_element)

for section in sections:
    print(f"{section.heading_text} (level {section.level})")
    print(f"  Method: {section.metadata.method.name}")
    print(f"  Confidence: {section.metadata.confidence.value}")
```

## Ref-Based Editing

Use refs for precise document edits. These methods operate on entire elements (paragraphs, cells),
not substrings within them.

```python
from python_docx_redline import Document

doc = Document("contract.docx")
tree = AccessibilityTree.from_document(doc)

# Insert text at a position relative to the element
# position: "before", "after" (new paragraph), "start", "end" (within paragraph)
doc.insert_at_ref("p:5", " (AMENDED)", position="end", track=True)

# Replace ENTIRE paragraph/cell content with new text
# Note: This replaces the whole content, not a substring
doc.replace_at_ref("p:10", "Replacement paragraph text", track=True)

# Delete ENTIRE element (note: method is delete_ref, not delete_at_ref)
doc.delete_ref("p:15", track=True)

# Table cell editing - replaces entire cell content
doc.replace_at_ref("tbl:0/row:2/cell:1", "$150", track=True)

doc.save("contract_edited.docx")
```

### Method Signatures

```python
# Insert text at ref location
doc.insert_at_ref(
    ref: str,           # e.g., "p:5", "tbl:0/row:1/cell:0"
    text: str,          # Text to insert (supports markdown: **bold**, *italic*)
    position: str = "after",  # "before", "after", "start", "end"
    track: bool = False,
    author: str | None = None
) -> EditResult

# Replace entire element content
doc.replace_at_ref(
    ref: str,           # Element to replace
    new_text: str,      # Replacement text (replaces ENTIRE content)
    track: bool = False,
    author: str | None = None
) -> EditResult

# Delete entire element
doc.delete_ref(
    ref: str,           # Element to delete
    track: bool = False,
    author: str | None = None
) -> EditResult
```

## Document Statistics

```python
stats = tree.stats
print(f"Paragraphs: {stats.paragraphs}")
print(f"Tables: {stats.tables}")
print(f"Tracked changes: {stats.tracked_changes}")
print(f"Comments: {stats.comments}")
print(f"Images: {stats.images}")
print(f"Bookmarks: {stats.bookmarks}")
print(f"Hyperlinks: {stats.hyperlinks}")
```

## Performance

The accessibility layer is optimized for large documents:

- **LRU caching** for ref resolution (<2ms warm cache)
- **Fingerprint index** for O(1) fingerprint lookups
- **Lazy loading** of document parts
- **Generator-based iteration** to avoid memory spikes

Performance targets:
- Outline mode: <100ms for 100 pages
- Content mode: <300ms for 100 pages
- Ref resolution (warm cache): <2ms

```python
# Check cache performance
stats = tree.registry.cache_stats
print(f"Cache hit rate: {stats.hit_rate:.1%}")
print(f"Avg resolution time: {stats.avg_resolution_time_ms:.2f}ms")
```

## Agent Workflow Example

```python
# 1. Load document and get overview
doc = Document("contract.docx")
tree = AccessibilityTree.from_document(doc)

# 2. Scan structure (minimal verbosity for speed)
print(tree.to_yaml(verbosity="minimal"))

# 3. Find specific content
results = tree.search("indemnification")
print(f"Found {len(results)} matches")

# 4. Expand relevant sections
for result in results[:3]:
    section = tree.expand_section(result.section_ref)
    print(section.to_yaml())

# 5. Make precise edits using refs
# Note: replace_at_ref replaces ENTIRE paragraph content, not a substring.
# For substring replacement, use the text-based editing API instead.
doc.replace_at_ref(results[0].ref, "The parties will indemnify each other...", track=True)

# 6. Verify changes
new_tree = AccessibilityTree.from_document(doc)
print(new_tree.to_yaml())

# 7. Save
doc.save("contract_edited.docx")
```

## See Also

- [reading.md](./reading.md) - Basic document reading
- [editing.md](./editing.md) - Document editing operations
- [tracked-changes.md](./tracked-changes.md) - Tracked changes details
- [DOCTREE_SPEC.md](../../docs/DOCTREE_SPEC.md) - Full specification
