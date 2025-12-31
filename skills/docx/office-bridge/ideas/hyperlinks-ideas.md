---
name: Office Bridge Hyperlinks Ideas
description: Roadmap and feature ideas for hyperlink support in the Office Bridge Word add-in. Includes helper functions, validation, bulk operations, and advanced features.
---

# Office Bridge Hyperlinks - Roadmap Ideas

This document tracks feature ideas and improvements for hyperlink support in the Office Bridge add-in.

## Current State

The Office Bridge has:
- Type definitions for `LinkInfo` and `LinkSummary`
- `SemanticRole.Link` in the accessibility tree
- `lnk:` ref prefix defined for hyperlinks
- `links?: LinkInfo[]` on `AccessibilityNode`

Missing:
- Actual hyperlink extraction in the tree builder
- Helper functions for inserting/editing hyperlinks
- Integration with the ref system for editing
- Validation utilities

## Priority 1: Core Helper Functions

### insertHyperlink()

Create a clean API for inserting hyperlinks that handles edge cases:

```typescript
interface InsertHyperlinkOptions {
  url?: string;              // External URL
  bookmark?: string;         // Internal bookmark name
  text: string;              // Display text
  tooltip?: string;          // Optional tooltip
  position: 'before' | 'after' | 'replace';
  range?: Word.Range;        // Target range (defaults to selection)
}

async function insertHyperlink(
  context: Word.RequestContext,
  options: InsertHyperlinkOptions
): Promise<string> {  // Returns lnk:N ref
  // Implementation handles:
  // - HTML escaping
  // - Table cell edge case
  // - Internal vs external links
  // - Returns ref for further operations
}
```

### getHyperlinks()

Extract all hyperlinks with our ref system:

```typescript
interface HyperlinkData {
  ref: string;           // lnk:N
  text: string;          // Display text
  url?: string;          // External URL
  bookmark?: string;     // Internal bookmark name
  tooltip?: string;      // Tooltip if set
  paragraphRef: string;  // Parent paragraph ref
  rangeIndex: number;    // Position within paragraph
}

async function getHyperlinks(
  context: Word.RequestContext,
  scope?: Word.Range
): Promise<HyperlinkData[]> {
  // Extract all hyperlinks with refs
  // Support optional scope (selection, paragraph, etc.)
}
```

### editHyperlink()

Edit hyperlink by ref:

```typescript
interface EditHyperlinkOptions {
  newUrl?: string;
  newBookmark?: string;
  newText?: string;
  newTooltip?: string;
}

async function editHyperlink(
  context: Word.RequestContext,
  ref: string,  // lnk:N
  options: EditHyperlinkOptions
): Promise<void> {
  // Look up hyperlink by ref
  // Apply changes
}
```

### removeHyperlink()

Remove hyperlink with options:

```typescript
interface RemoveHyperlinkOptions {
  keepText?: boolean;  // Default: true
}

async function removeHyperlink(
  context: Word.RequestContext,
  ref: string,
  options?: RemoveHyperlinkOptions
): Promise<void> {
  // Remove hyperlink
  // Optionally keep or remove text
}
```

## Priority 2: Hyperlink Validation

### validateHyperlinks()

Check all hyperlinks for issues:

```typescript
interface HyperlinkValidationResult {
  ref: string;
  text: string;
  url?: string;
  bookmark?: string;
  status: 'valid' | 'broken' | 'warning';
  issue?: string;
}

async function validateHyperlinks(
  context: Word.RequestContext
): Promise<{
  valid: HyperlinkValidationResult[];
  broken: HyperlinkValidationResult[];
  warnings: HyperlinkValidationResult[];
}> {
  // Check internal links point to existing bookmarks
  // Check external links are well-formed URLs
  // Warn about mailto: links, file: links, etc.
}
```

Validation checks to implement:
1. **Internal link targets** - Does the bookmark exist?
2. **URL format** - Is it a valid URL?
3. **Protocol warnings** - Warn about file://, javascript:, etc.
4. **Empty links** - Hyperlink with no URL or bookmark
5. **Duplicate links** - Same text with different URLs
6. **Long URLs** - URLs that may be truncated

### fixBrokenLinks()

Suggest or auto-fix broken links:

```typescript
interface BrokenLinkFix {
  ref: string;
  originalTarget: string;
  suggestedFix?: string;
  fixType: 'bookmark_rename' | 'url_update' | 'remove';
}

async function suggestBrokenLinkFixes(
  context: Word.RequestContext,
  brokenLinks: HyperlinkValidationResult[]
): Promise<BrokenLinkFix[]> {
  // For internal links: suggest similar bookmark names
  // For external links: suggest corrections if pattern matches
}
```

## Priority 3: Bulk Operations

### updateUrlDomain()

Bulk update URLs matching a pattern:

```typescript
async function updateUrlDomain(
  context: Word.RequestContext,
  oldDomain: string,
  newDomain: string
): Promise<number> {
  // Find all links containing oldDomain
  // Replace with newDomain
  // Return count of updated links
}
```

### convertToAbsoluteUrls()

Convert relative URLs to absolute:

```typescript
async function convertToAbsoluteUrls(
  context: Word.RequestContext,
  baseUrl: string
): Promise<number> {
  // Find relative URLs (starting with / or no protocol)
  // Prepend baseUrl
  // Return count
}
```

### removeAllHyperlinks()

Clear all hyperlinks from document:

```typescript
async function removeAllHyperlinks(
  context: Word.RequestContext,
  options?: {
    keepText?: boolean;
    scope?: 'document' | 'selection' | Word.Range;
  }
): Promise<number> {
  // Remove all hyperlinks
  // Return count
}
```

## Priority 4: Accessibility Tree Integration

### Include Links in Tree

Update the tree builder to extract hyperlinks:

```typescript
// In builder.ts, when processing paragraphs:
async function extractLinksFromParagraph(
  paragraph: Word.Paragraph,
  paragraphRef: string
): Promise<LinkInfo[]> {
  const range = paragraph.getRange(Word.RangeLocation.whole);
  const hyperlinkRanges = range.getHyperlinkRanges();

  hyperlinkRanges.load(['hyperlink', 'text']);
  await context.sync();

  return hyperlinkRanges.items.map((range, index) => {
    const hyperlink = range.hyperlink;
    const isInternal = hyperlink.startsWith('#');

    return {
      ref: `lnk:${paragraphRef}/${index}`,
      text: range.text,
      url: isInternal ? undefined : hyperlink,
      target: isInternal ? `bk:${hyperlink.slice(1)}` : undefined
    };
  });
}
```

### Build Link Summary

Populate the `LinkSummary` in the tree:

```typescript
function buildLinkSummary(allLinks: LinkInfo[]): LinkSummary {
  const internal: LinkInfo[] = [];
  const external: LinkInfo[] = [];
  const broken: Array<LinkInfo & { error: string }> = [];

  for (const link of allLinks) {
    if (link.target) {
      // Internal link - check if bookmark exists
      if (bookmarkExists(link.target)) {
        internal.push(link);
      } else {
        broken.push({ ...link, error: `Bookmark not found: ${link.target}` });
      }
    } else if (link.url) {
      external.push(link);
    } else {
      broken.push({ ...link, error: 'No URL or bookmark target' });
    }
  }

  return { internal, external, broken };
}
```

## Priority 5: Advanced Features

### Cross-Reference Support

Word's cross-references are special hyperlinks to bookmarks on headings, figures, etc.:

```typescript
interface CrossReference {
  ref: string;            // xref:N
  type: 'heading' | 'figure' | 'table' | 'equation' | 'footnote' | 'endnote';
  targetRef: string;      // What it points to
  displayFormat: 'full' | 'number' | 'title' | 'page';
  text: string;           // Current display text
}

async function getCrossReferences(
  context: Word.RequestContext
): Promise<CrossReference[]> {
  // Cross-references use field codes like { REF _Ref123456 \h }
  // Need to parse field codes to identify them
}

async function insertCrossReference(
  context: Word.RequestContext,
  targetRef: string,
  displayFormat: string
): Promise<string> {
  // Insert proper cross-reference field
}
```

### Table of Contents Links

TOC entries are hyperlinks to headings:

```typescript
interface TocLink {
  ref: string;
  headingText: string;
  pageNumber?: number;
  targetHeadingRef: string;
  level: number;
}

async function getTocLinks(
  context: Word.RequestContext
): Promise<TocLink[]> {
  // Parse TOC entries and extract links
}
```

### Smart Link Suggestions

Based on document content, suggest links:

```typescript
interface LinkSuggestion {
  text: string;           // Text that could be linked
  suggestedUrl?: string;  // Suggested URL
  suggestedBookmark?: string;  // Or internal link
  confidence: number;     // 0-1 confidence score
  reason: string;         // Why this is suggested
}

async function suggestLinks(
  context: Word.RequestContext
): Promise<LinkSuggestion[]> {
  // Find text patterns that look like URLs (www., http)
  // Find references to headings that could be internal links
  // Find email patterns for mailto: links
}
```

## Implementation Notes

### Ref Format for Hyperlinks

Use hierarchical refs for precise addressing:

```
lnk:0              - First hyperlink in document body
lnk:p:5/0          - First hyperlink in paragraph 5
lnk:hdr:default/0  - First hyperlink in default header
lnk:fn:3/0         - First hyperlink in footnote 3
```

### Caching Considerations

Hyperlinks are relatively static - cache aggressively:

```typescript
interface HyperlinkCache {
  links: Map<string, HyperlinkData>;  // ref -> data
  timestamp: number;
  documentHash?: string;  // For staleness detection
}

// Invalidate cache on:
// - Any edit operation
// - Document switch
// - After a timeout (30s?)
```

### Error Handling

Define specific error types:

```typescript
class HyperlinkNotFoundError extends Error {
  constructor(public ref: string) {
    super(`Hyperlink not found: ${ref}`);
  }
}

class InvalidHyperlinkError extends Error {
  constructor(public reason: string) {
    super(`Invalid hyperlink: ${reason}`);
  }
}

class TableCellHyperlinkError extends Error {
  constructor() {
    super('Cannot add hyperlink to entire table cell range');
  }
}
```

## Testing Considerations

### Test Cases

1. Basic insert/read/edit/remove cycle
2. Internal links to bookmarks
3. External links with various protocols
4. Links in tables (cell content, not whole cell)
5. Links in headers/footers
6. Links in footnotes/endnotes
7. Multiple links in same paragraph
8. Unicode in display text
9. Very long URLs
10. Special characters in URLs (encoding)

### Platform Testing

- Word Desktop (Windows)
- Word Desktop (Mac)
- Word Online
- Word Mobile (if supported)

Test for:
- API availability differences
- Behavior differences (insertHtml styling)
- Performance differences

## Related Work

### Python Library Parity

Match features from `python_docx_redline.operations.hyperlinks`:
- [x] Insert external hyperlinks
- [x] Insert internal hyperlinks (bookmarks)
- [ ] Insert in headers/footers
- [ ] Insert in footnotes/endnotes
- [x] Edit URL
- [x] Edit text
- [ ] Edit anchor
- [x] Remove hyperlink
- [ ] Get all hyperlinks
- [ ] Get hyperlink by ref
- [ ] Get hyperlinks by URL pattern

### Future Considerations

1. **Tracked changes for hyperlinks** - Office.js doesn't support this directly; may need workaround or document this limitation

2. **Batch operations** - Office.js supports batching; design helpers to take advantage of this

3. **Undo support** - Office.js operations are automatically undoable, but test complex multi-step operations

4. **Real-time collaboration** - How do hyperlink operations behave with multiple editors?
