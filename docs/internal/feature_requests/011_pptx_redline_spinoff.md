# Feature Request: python-pptx-redline Spinoff Library

**Status**: Idea / Not Planned
**Created**: 2025-12-16
**Priority**: Future consideration

## Overview

This document captures analysis of creating a sibling library for PowerPoint files, applying the same high-level API philosophy used in `python-docx-redline`.

## Motivation

`python-docx-redline` has proven effective at reducing Word document manipulation from 30+ lines of XML to 3 lines of API calls. The same pain points exist for PowerPoint:

- `python-pptx` requires verbose, low-level code
- Finding and replacing text requires traversing complex shape hierarchies
- Template filling is error-prone and tedious
- No high-level abstractions for common operations

## Key Technical Insight: Same Foundation

PowerPoint `.pptx` files use the same OOXML format as Word `.docx` files:

```
presentation.pptx/
├── [Content_Types].xml
├── _rels/
├── ppt/
│   ├── presentation.xml      # Main presentation
│   ├── slides/
│   │   ├── slide1.xml        # Each slide
│   │   └── slide2.xml
│   ├── slideLayouts/         # Layout templates
│   ├── slideMasters/         # Master slides
│   └── notesSlides/          # Speaker notes
└── docProps/
```

All lxml/XML manipulation patterns from `python-docx-redline` would work identically.

## Critical Difference: No Tracked Changes

**PowerPoint does NOT have Word-style tracked changes.** This fundamentally changes the value proposition:

| Feature | Word | PowerPoint |
|---------|------|------------|
| Tracked insertions (`<w:ins>`) | ✅ | ❌ |
| Tracked deletions (`<w:del>`) | ✅ | ❌ |
| Revision history | ✅ Full | ⚠️ Limited |
| Comments | ✅ | ✅ (different format) |
| Compare documents | ✅ Built-in | ❌ |

The core value prop shifts from "tracked changes" to "template automation and bulk text operations."

## Patterns That Transfer Directly

### Code Reuse (~40%)

| Component | Reusable | Notes |
|-----------|----------|-------|
| `text_search.py` | ✅ 100% | Identical algorithm |
| `fuzzy.py` | ✅ 100% | Identical implementation |
| `quote_normalization.py` | ✅ 100% | Identical implementation |
| Operations pattern | ✅ Architecture | Same class structure |
| Error classes | ✅ 90% | `TextNotFoundError`, etc. |
| Test patterns | ✅ 80% | Same pytest approach |

### Architectural Patterns

- High-level API wrapping low-level library
- Separation of concerns (operations, search, models)
- Context extraction for error messages
- Fuzzy matching with configurable threshold

## Proposed API Design

### Current Pain (python-pptx)

```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation('template.pptx')
slide = prs.slides[0]
for shape in slide.shapes:
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if "{{CLIENT_NAME}}" in run.text:
                    run.text = run.text.replace("{{CLIENT_NAME}}", "Acme Corp")
# Hope you didn't miss any shapes...
```

### Proposed High-Level API

```python
from python_pptx_redline import Presentation

prs = Presentation("template.pptx")

# Simple find/replace across entire presentation
prs.replace_all("{{CLIENT_NAME}}", "Acme Corp")
prs.replace_all("{{DATE}}", "December 2025")

# Slide-specific operations
prs.slides[0].replace("Old Title", "New Title")

# With formatting preservation
prs.replace_all("placeholder", "value", preserve_formatting=True)

# Fuzzy matching for OCR'd presentations
prs.replace_all("Sect1on 2.1", "Section 2.1", fuzzy=0.85)

# Bulk template filling
prs.fill_template({
    "{{CLIENT}}": "Acme Corp",
    "{{PROJECT}}": "Q4 Analysis",
    "{{AUTHOR}}": "Jane Smith",
})

# Clone slide with modifications
prs.duplicate_slide(0, replacements={"Product A": "Product B"})

# Extract all text (for AI processing)
text = prs.get_all_text()

# Find text locations
matches = prs.find_all("revenue", context=20)
for m in matches:
    print(f"Slide {m.slide_number}, Shape: {m.shape_name}: {m.context}")

prs.save("output.pptx")
```

## Implementation Difficulty Assessment

| Component | Difficulty | Notes |
|-----------|------------|-------|
| Basic find/replace | **Easy** | Direct transfer from Word patterns |
| Template filling | **Easy** | Common use case, straightforward |
| Fuzzy matching | **Easy** | Copy directly from this library |
| Format preservation | **Medium** | PowerPoint runs work similarly |
| Shape-aware operations | **Medium** | Need to traverse shape tree |
| Table operations | **Medium** | Similar to Word tables |
| SmartArt modification | **Hard** | Complex nested XML |
| Chart data updates | **Hard** | Charts are complex objects |
| Master/layout inheritance | **Hard** | Style inheritance is tricky |

## Recommended Phased Approach

### Phase 1: Core Value (~1 week)

The 80% use case - template filling and text replacement:

```python
prs = Presentation("template.pptx")
prs.replace_all(old, new)
prs.fill_template(dict)
prs.get_all_text()
prs.find_all(pattern)
prs.save("output.pptx")
```

### Phase 2: Slide Operations (~1 week)

```python
prs.slides[0].replace(old, new)
prs.duplicate_slide(index, replacements)
prs.slides[0].get_text()
prs.slides[0].shapes  # High-level shape access
```

### Phase 3: Advanced (~2 weeks)

```python
prs.add_comment(slide, text, position)
prs.slides[0].notes = "Speaker notes here"
prs.slides[0].tables[0].replace(old, new)
```

## Proposed Project Structure

```
python-pptx-redline/
├── src/python_pptx_redline/
│   ├── __init__.py
│   ├── presentation.py      # Main Presentation class
│   ├── slide.py             # Slide operations
│   ├── shape.py             # Shape operations
│   ├── text_search.py       # Copy from docx-redline
│   ├── fuzzy.py             # Copy from docx-redline
│   ├── quote_normalization.py  # Copy from docx-redline
│   └── operations/
│       ├── text.py          # Text replacement
│       ├── template.py      # Template filling
│       └── duplicate.py     # Slide duplication
├── tests/
├── docs/
├── pyproject.toml
└── README.md
```

## Effort Estimates

| Scope | Effort |
|-------|--------|
| MVP (Phase 1) | ~1-2 weeks |
| Full library (Phases 1-3) | ~4-6 weeks |
| Code reuse from docx-redline | ~40% |

## Summary

| Aspect | Assessment |
|--------|------------|
| **Technical feasibility** | High - same OOXML foundation |
| **Code reuse** | ~40% (text search, fuzzy, normalization, patterns) |
| **Tracked changes equivalent** | ❌ Doesn't exist in PowerPoint |
| **Alternative value prop** | Template filling, bulk find/replace, AI-friendly text extraction |
| **Recommendation** | Viable as separate project when need arises |

## Decision

**Not implementing at this time.** This document preserved for future reference if the need arises.

## References

- [python-pptx documentation](https://python-pptx.readthedocs.io/)
- [OOXML PresentationML specification](http://officeopenxml.com/anatomyofOOXML-pptx.php)
- `python-docx-redline` patterns in this repository
