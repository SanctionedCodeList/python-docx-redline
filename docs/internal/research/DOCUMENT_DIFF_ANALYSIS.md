# Document Diff Algorithm Analysis

**Research Task**: `docx_redline-bab`
**Date**: 2025-12-13
**Author**: Research by Claude (Opus 4.5)

---

## Executive Summary

This document analyzes approaches for implementing document comparison (`compare_documents()`) for the python_docx_redline library. After researching standard diff algorithms, library options, and Word's native behavior, we recommend a **hybrid multi-level approach** combining paragraph-level structural alignment with word-level content diffing.

### Key Recommendations

1. **Paragraph-level first, word-level second**: Use two-pass diffing for performance and clarity
2. **Word-level granularity for content**: Match legal redlining conventions
3. **Simple heuristics for move detection**: Detect identical paragraphs appearing in different positions
4. **Formatting as metadata**: Track formatting separately from content changes
5. **Tables as atomic units initially**: Handle table diffing in a later phase

---

## Research Questions Answered

### 1. Paragraph-level vs Word-level Diff?

**Answer: Use both in a hierarchical approach.**

| Approach | Pros | Cons | Best For |
|----------|------|------|----------|
| **Paragraph-level** | Fast, handles structural changes (add/remove sections), detects moved paragraphs | Coarse-grained, shows entire paragraph as changed for small edits | Structural comparison, TOC changes |
| **Word-level** | Precise, legal-redline readable, matches lawyer expectations | Slower, can't detect paragraph moves, verbose for large changes | In-paragraph edits, contract revisions |
| **Character-level** | Most precise | Too noisy for legal documents, hard to read | Rarely appropriate for DOCX |

**Recommended Strategy: Two-Pass Hierarchical Diff**

```
Pass 1: Paragraph Alignment (structural level)
├── Detect added paragraphs
├── Detect deleted paragraphs
├── Detect moved paragraphs (same content, different position)
└── Identify modified paragraphs (content changed)

Pass 2: Word-Level Diff (content level)
├── For each modified paragraph pair:
│   ├── Tokenize to words/punctuation
│   ├── Compute word-level diff
│   └── Generate tracked change hunks
└── Apply legal-style rules (R1-R5 from minimal_diff.py)
```

This matches the existing `minimal_diff.py` implementation, which already uses word-level tokenization with `SequenceMatcher` for within-paragraph changes.

### 2. How to Detect Moved Content vs Delete+Insert?

**Answer: Fingerprint-based matching with position tracking.**

Standard LCS-based algorithms (like Python's `difflib`) cannot detect moves—they only find in-order matches. To detect moves:

**Algorithm: Paragraph Fingerprinting**

```python
def detect_moves(old_paragraphs, new_paragraphs):
    """Detect paragraphs that were moved vs truly added/deleted."""

    # Step 1: Build fingerprint maps
    # Fingerprint = normalized text content (strip formatting)
    old_fingerprints = {fingerprint(p): (i, p) for i, p in enumerate(old_paragraphs)}
    new_fingerprints = {fingerprint(p): (i, p) for i, p in enumerate(new_paragraphs)}

    moves = []

    # Step 2: Find fingerprints in both (potential moves)
    common = set(old_fingerprints.keys()) & set(new_fingerprints.keys())

    for fp in common:
        old_idx, old_para = old_fingerprints[fp]
        new_idx, new_para = new_fingerprints[fp]

        if old_idx != new_idx:
            # Same content, different position = move
            moves.append(Move(old_idx, new_idx, old_para))

    return moves
```

**Fingerprint Design Considerations:**

| Fingerprint Type | Detects Moves? | Tolerance |
|------------------|----------------|-----------|
| Exact text match | Yes | None - any change breaks detection |
| Normalized (whitespace, case) | Yes | Minor formatting |
| Fuzzy hash (SimHash) | Yes | ~10% content change |
| Sentence/phrase overlap | Partial | Allows some editing within moved block |

**Recommendation**: Start with **exact normalized text match** (whitespace normalized, formatting stripped). This matches Microsoft Word's behavior:

> "The Word Compare tool will only recognise 'moved' text when a paragraph is moved without any changes to it."

Word uses the same limitation—only detecting perfect moves. This is a reasonable starting point; enhanced move detection can be added later.

**OOXML Move Markup:**

Word has native markup for moves:
```xml
<!-- Move source (original location) -->
<w:moveFrom w:id="1" w:name="move1" w:author="Author" w:date="...">
  <w:r><w:t>Moved text</w:t></w:r>
</w:moveFrom>

<!-- Move destination (new location) -->
<w:moveTo w:id="2" w:name="move1" w:author="Author" w:date="...">
  <w:r><w:t>Moved text</w:t></w:r>
</w:moveTo>
```

Using native move markup provides better Word compatibility and cleaner redlines than delete+insert.

### 3. How to Handle Formatting Changes?

**Answer: Track formatting separately, with configurable inclusion.**

Formatting changes in OOXML are represented via `<w:rPrChange>` (run property change) and `<w:pPrChange>` (paragraph property change):

```xml
<w:rPr>
  <w:b/>  <!-- Current: bold -->
  <w:rPrChange w:id="1" w:author="Author" w:date="...">
    <w:rPr/>  <!-- Previous: not bold -->
  </w:rPrChange>
</w:rPr>
```

**Comparison Levels:**

| Level | What's Compared | Use Case |
|-------|-----------------|----------|
| **Content Only** | Text content | Legal redlines, focus on words |
| **Content + Structure** | Text + paragraphs/sections | Document restructuring |
| **Content + Formatting** | Text + bold/italic/underline | Design review |
| **Full** | Everything | Complete audit trail |

**Recommendation**: Default to **content-only** comparison (like Word's default), with option to include formatting:

```python
def compare_documents(
    original: Document,
    modified: Document,
    compare_formatting: bool = False,  # Default: ignore formatting
    compare_case: bool = True,
    compare_whitespace: bool = False,
) -> ComparisonResult:
    ...
```

**Formatting Diff Algorithm:**

```python
def diff_run_properties(old_rpr, new_rpr):
    """Compare two w:rPr elements for formatting changes."""

    TRACKED_PROPERTIES = [
        ('w:b', 'bold'),
        ('w:i', 'italic'),
        ('w:u', 'underline'),
        ('w:strike', 'strikethrough'),
        ('w:sz', 'font_size'),
        ('w:rFonts', 'font_family'),
        # ... etc
    ]

    changes = []
    for tag, name in TRACKED_PROPERTIES:
        old_val = get_property(old_rpr, tag)
        new_val = get_property(new_rpr, tag)
        if old_val != new_val:
            changes.append(FormattingChange(name, old_val, new_val))

    return changes
```

### 4. What About Tables, Images, and Other Content?

**Answer: Handle progressively, with tables as the priority.**

| Content Type | Complexity | Approach | Phase |
|--------------|------------|----------|-------|
| **Tables** | High | Specialized table diff | Phase 12 (future) |
| **Images** | Medium | Hash comparison + position | Phase 14 (future) |
| **Lists** | Medium | Treat as paragraph sequence | Phase 12 (future) |
| **Headers/Footers** | Low | Same as body paragraphs | Current |
| **Footnotes/Endnotes** | Low | Same as body paragraphs | Current |

**Table Diff Considerations:**

Tables require specialized handling because they have:
- Row/column structure
- Merged cells
- Cell-level content that can change
- Structural changes (add/delete rows/columns)

**Proposed Table Diff Approach:**

```
1. Structural Comparison
   ├── Compare row counts
   ├── Compare column counts
   └── Identify added/deleted rows and columns

2. Cell Alignment
   ├── Match cells by position (simple)
   └── Or match cells by content fingerprint (handles row moves)

3. Cell Content Diff
   ├── For each matched cell pair:
   │   └── Run paragraph-level diff on cell content
   └── Generate tracked changes
```

**Row Matching Options:**

| Method | Pros | Cons |
|--------|------|------|
| **Position-based** | Simple, fast | Can't detect row moves |
| **Key column** | Detects row reordering | Requires identifying key |
| **Content fingerprint** | Flexible | Slower, may mismatch |

**Recommendation for Tables**: Use **position-based matching** initially. This handles:
- Content changes within cells
- Simple row additions/deletions at the end

Defer row move detection to a future enhancement.

**Image Comparison:**

Images should be compared by:
1. **Content hash** (MD5/SHA256 of image bytes)
2. **Position in document**
3. **Alt text changes**

```python
def compare_images(old_img, new_img):
    if hash(old_img.blob) != hash(new_img.blob):
        return ImageChange.CONTENT_CHANGED
    if old_img.position != new_img.position:
        return ImageChange.MOVED
    return ImageChange.UNCHANGED
```

---

## Algorithm Options Analysis

### Option 1: Python difflib.SequenceMatcher

**Already in use** in `minimal_diff.py` for word-level diffing.

```python
from difflib import SequenceMatcher

matcher = SequenceMatcher(None, old_tokens, new_tokens, autojunk=False)
opcodes = matcher.get_opcodes()
# Returns: [('equal', 0, 5, 0, 5), ('replace', 5, 7, 5, 8), ...]
```

**Pros:**
- Standard library, no dependencies
- Well-tested, performant for typical document sizes
- Produces edit operations directly

**Cons:**
- O(n²) worst case (O(n) typical)
- No move detection
- "Autojunk" heuristic can misfire (disabled in our code)

**Verdict**: Keep using for word-level diffs. Add paragraph-level wrapper.

### Option 2: Google diff-match-patch

[PyPI: diff-match-patch](https://pypi.org/project/diff-match-patch/)

Originally built for Google Docs synchronization.

```python
import diff_match_patch as dmp_module

dmp = dmp_module.diff_match_patch()
diffs = dmp.diff_main(text1, text2)
dmp.diff_cleanupSemantic(diffs)  # Semantic cleanup
```

**Pros:**
- Character-level precision
- Semantic cleanup built-in (aligns to word boundaries)
- Fuzzy patch application
- Battle-tested at Google scale

**Cons:**
- Character-level by default (need wrapper for word-level)
- No move detection
- Library is now archived (though maintained fork exists)

**Verdict**: Consider as alternative to difflib if edge cases arise. The semantic cleanup is valuable.

### Option 3: Paul Heckel's Diff Algorithm

[GitHub Gist: Paul Heckel's Algorithm](https://gist.github.com/ndarville/3166060)

Designed to detect differences including moves.

**Algorithm Overview:**
1. Build symbol table of unique lines
2. Identify lines that appear exactly once in each document (anchors)
3. Use anchors to align documents
4. Propagate alignment from anchors

**Pros:**
- O(n) time and space
- Produces intuitive diffs (matches human perception)
- Can detect some types of moves
- Inline diff vs block diff

**Cons:**
- Struggles with repeated text
- Less mature Python implementations
- More complex to implement correctly

**Verdict**: Interesting for future move detection. Not needed for MVP.

### Option 4: XML-Aware Diff (DeltaXML approach)

[DeltaXML: XML Compare](https://www.deltaxml.com/products/compare/xml-compare/)

Compares XML structure, not just text.

**Approach:**
1. Parse both documents as XML trees
2. Match elements by: tag name, attributes, parent context, child structure
3. Diff matched elements
4. Handle structural changes (added/removed nodes)

**Pros:**
- Preserves document structure
- Handles XML-specific issues (attribute order, whitespace)
- Can track nested changes

**Cons:**
- Complex to implement
- OOXML has deep nesting that complicates matching
- Overkill for text-focused comparison

**Verdict**: Not recommended for MVP. OOXML structure is too complex. Better to operate at paragraph/run level.

---

## Recommended Implementation Strategy

### Phase 1: Basic Document Comparison (MVP)

```python
class ComparisonResult:
    """Result of comparing two documents."""

    added_paragraphs: list[ParagraphChange]    # New paragraphs in modified
    deleted_paragraphs: list[ParagraphChange]  # Removed from original
    modified_paragraphs: list[ParagraphModification]  # Content changed
    moved_paragraphs: list[ParagraphMove]      # Same content, new position

def compare_documents(
    original: Document,
    modified: Document,
    granularity: Literal["paragraph", "word"] = "word",
    detect_moves: bool = True,
) -> ComparisonResult:
    """Compare two documents and return differences."""

    # Step 1: Extract paragraphs from both documents
    old_paras = extract_paragraphs(original)
    new_paras = extract_paragraphs(modified)

    # Step 2: Paragraph-level alignment
    alignment = align_paragraphs(old_paras, new_paras, detect_moves)

    # Step 3: Word-level diff for modified paragraphs
    for match in alignment.modified_pairs:
        diff = compute_minimal_hunks(match.old_text, match.new_text)
        match.hunks = diff.hunks

    return build_result(alignment)
```

### Phase 2: Generate Tracked Changes

```python
def generate_redline(
    original: Document,
    comparison: ComparisonResult,
    author: str = "Compare",
) -> Document:
    """Apply comparison results as tracked changes to original document."""

    redline = original.copy()

    # Apply paragraph-level changes
    for added in comparison.added_paragraphs:
        insert_paragraph_tracked(redline, added, author)

    for deleted in comparison.deleted_paragraphs:
        delete_paragraph_tracked(redline, deleted, author)

    for moved in comparison.moved_paragraphs:
        move_paragraph_tracked(redline, moved, author)

    # Apply word-level changes within modified paragraphs
    for modified in comparison.modified_paragraphs:
        apply_minimal_edits_to_paragraph(
            modified.paragraph,
            modified.hunks,
            xml_generator,
            author
        )

    return redline
```

### Proposed API

```python
# Simple comparison
result = doc1.compare_to(doc2)

# With options
result = doc1.compare_to(
    doc2,
    granularity="word",        # "paragraph" | "word" | "character"
    detect_moves=True,          # Detect moved paragraphs
    ignore_formatting=True,     # Focus on content
    ignore_case=False,
    ignore_whitespace=True,
)

# Generate redline document
redline = doc1.create_redline(doc2, author="Review")
redline.save("comparison_output.docx")

# Alternative: apply diff to original
doc1.apply_comparison(result, author="Review")
doc1.save("redlined.docx")
```

---

## Performance Considerations

### Complexity Analysis

| Operation | Time Complexity | Space Complexity |
|-----------|-----------------|------------------|
| Paragraph extraction | O(n) | O(n) |
| Paragraph fingerprinting | O(n) | O(n) |
| Paragraph alignment (LCS) | O(n²) worst, O(n) typical | O(n) |
| Move detection | O(n) | O(n) |
| Word-level diff per para | O(m²) where m = words | O(m) |
| **Total** | O(n + k·m²) where k = modified paras | O(n + m) |

For typical legal documents (100-500 paragraphs, <100 words each), this is very fast.

### Optimization Opportunities

1. **Early termination**: Skip word-level diff if paragraph is identical
2. **Parallel processing**: Diff modified paragraphs in parallel
3. **Caching**: Cache paragraph fingerprints
4. **Chunking**: For very large documents, process in sections

---

## References

### Academic/Technical
- [Neil Fraser: Diff Strategies](https://neil.fraser.name/writing/diff/) - Multi-level diff approaches
- [Hunt-McIlroy Algorithm (1976)](https://www.cs.dartmouth.edu/~doug/diff.pdf) - Foundational diff paper
- [Paul Heckel's Algorithm](https://gist.github.com/ndarville/3166060) - Move-aware diffing

### Libraries
- [Python difflib](https://docs.python.org/3/library/difflib.html) - Standard library
- [Google diff-match-patch](https://github.com/google/diff-match-patch) - Character-level diff
- [Florian Hartmann: Diffing](https://florian.github.io/diffing/) - Algorithm explanation

### Commercial/Tools
- [Microsoft Word Compare](https://learn.microsoft.com/en-us/office/vba/api/word.document.compare) - VBA API reference
- [Altova DiffDog](https://www.altova.com/diffdog) - OOXML diff tool
- [DeltaXML](https://www.deltaxml.com/products/compare/xml-compare/) - XML-aware comparison

### Existing Code
- `minimal_diff.py` - Word-level diff implementation (lines 125-226)
- `text_search.py` - Text fragmentation handling
- `docs/ERIC_WHITE_ALGORITHM.md` - Run-level text manipulation

---

## Conclusion

For `compare_documents()` implementation, we recommend:

1. **Use the existing `minimal_diff.py` foundation** for word-level comparison
2. **Add paragraph-level alignment** using fingerprint matching
3. **Implement move detection** using exact paragraph matching (like Word)
4. **Defer formatting comparison** to a configurable option
5. **Handle tables atomically** in the MVP, with specialized diffing in Phase 12

This approach balances implementation complexity with feature completeness, aligning with both user expectations (legal redlines) and technical constraints (OOXML complexity).
