---
name: office-bridge-footnotes-roadmap
description: "Roadmap and ideas for enhancing footnote/endnote support in Office Bridge. Contains proposed helper functions, integration improvements, and advanced features."
---

# Office Bridge Footnotes: Roadmap and Ideas

This document tracks proposed improvements for footnote and endnote handling in the Office Bridge Word add-in.

## Current State

The Office Bridge currently provides:
- Footnote extraction from OOXML during tree building
- `fn:N` and `en:N` ref system
- `getTextByRef()` for footnote text retrieval
- Paragraph-to-footnote mapping in tree output

## Priority 1: Core Helper Functions

### insertFootnoteByRef

Insert a footnote at a paragraph or range specified by ref.

```typescript
/**
 * Insert a footnote at the end of the specified paragraph.
 * @param context - Word.RequestContext
 * @param ref - Paragraph ref (e.g., "p:5")
 * @param text - Footnote text content
 * @returns The new footnote ref (e.g., "fn:3")
 */
async function insertFootnoteByRef(
  context: WordRequestContext,
  ref: Ref,
  text: string
): Promise<Ref>;

// Usage
const newRef = await DocTree.insertFootnoteByRef(context, "p:5", "Citation text");
// Returns "fn:3" (next available footnote number)
```

**Implementation notes:**
- Resolve paragraph ref to `Word.Paragraph`
- Get range at end of paragraph
- Call `range.insertFootnote(text)`
- Return new footnote ref based on position

### insertEndnoteByRef

Same pattern for endnotes:

```typescript
async function insertEndnoteByRef(
  context: WordRequestContext,
  ref: Ref,
  text: string
): Promise<Ref>;
```

### editFootnote / editEndnote

Update footnote content by ref.

```typescript
/**
 * Edit the text content of a footnote.
 * @param context - Word.RequestContext
 * @param ref - Footnote ref (e.g., "fn:2")
 * @param newText - New text content
 */
async function editFootnote(
  context: WordRequestContext,
  ref: Ref,
  newText: string
): Promise<void>;

// Usage
await DocTree.editFootnote(context, "fn:2", "Updated citation with 2024 data.");
```

**Implementation approach:**
- Parse ref to get footnote index
- Access via `body.footnotes.items[index]`
- Replace body content: `footnote.body.clear(); footnote.body.insertText(newText, "Start")`

### deleteFootnote / deleteEndnote

Delete footnotes by ref.

```typescript
/**
 * Delete a footnote and its reference mark.
 * @param context - Word.RequestContext
 * @param ref - Footnote ref (e.g., "fn:1")
 */
async function deleteFootnote(
  context: WordRequestContext,
  ref: Ref
): Promise<void>;

// Usage
await DocTree.deleteFootnote(context, "fn:3");
```

**Notes:**
- Word auto-renumbers remaining footnotes
- Tree should be rebuilt after deletion for accurate refs

### getFootnotes / getEndnotes

Get all footnotes with metadata.

```typescript
interface FootnoteData {
  ref: Ref;
  id: number;
  text: string;
  referencedFrom?: Ref;
}

async function getFootnotes(
  context: WordRequestContext
): Promise<FootnoteData[]>;

// Usage
const footnotes = await DocTree.getFootnotes(context);
// Returns array of { ref: "fn:1", id: 1, text: "...", referencedFrom: "p:3" }
```

## Priority 2: Navigation and Selection

### selectFootnoteReference

Navigate to footnote reference in document.

```typescript
/**
 * Select the footnote reference mark in the document body.
 * @param context - Word.RequestContext
 * @param ref - Footnote ref (e.g., "fn:2")
 */
async function selectFootnoteReference(
  context: WordRequestContext,
  ref: Ref
): Promise<void>;

// Usage
await DocTree.selectFootnoteReference(context, "fn:2");
// Word scrolls to and selects the superscript "2" in the body
```

### selectFootnoteBody

Navigate to footnote content area.

```typescript
async function selectFootnoteBody(
  context: WordRequestContext,
  ref: Ref
): Promise<void>;
```

### getNextFootnote / getPreviousFootnote

Navigate between footnotes.

```typescript
async function getNextFootnote(
  context: WordRequestContext,
  currentRef: Ref
): Promise<Ref | null>;

// Usage
const next = await DocTree.getNextFootnote(context, "fn:2");
// Returns "fn:3" or null if at last footnote
```

## Priority 3: Batch Operations

### batchInsertFootnotes

Insert multiple footnotes efficiently.

```typescript
interface FootnoteInsert {
  atRef: Ref;      // Paragraph ref
  text: string;    // Footnote content
  position?: 'end' | 'start';  // Where in paragraph (default: end)
}

async function batchInsertFootnotes(
  context: WordRequestContext,
  inserts: FootnoteInsert[]
): Promise<{ refs: Ref[]; errors: Error[] }>;

// Usage
const result = await DocTree.batchInsertFootnotes(context, [
  { atRef: "p:3", text: "First citation" },
  { atRef: "p:7", text: "Second citation" },
  { atRef: "p:12", text: "Third citation" },
]);
```

### batchEditFootnotes

Edit multiple footnotes in one operation.

```typescript
interface FootnoteEdit {
  ref: Ref;        // Footnote ref
  newText: string; // New content
}

async function batchEditFootnotes(
  context: WordRequestContext,
  edits: FootnoteEdit[]
): Promise<{ successCount: number; errors: Error[] }>;
```

### batchDeleteFootnotes

Delete multiple footnotes.

```typescript
async function batchDeleteFootnotes(
  context: WordRequestContext,
  refs: Ref[]
): Promise<{ deletedCount: number; errors: Error[] }>;
```

## Priority 4: Advanced Features

### Tracked Changes in Footnotes

Support tracked edits within footnote content.

```typescript
async function editFootnoteTracked(
  context: WordRequestContext,
  ref: Ref,
  newText: string,
  options?: { author?: string }
): Promise<void>;
```

**Challenge:** Office.js doesn't expose tracked changes API directly for footnote bodies. May require OOXML manipulation.

### Rich Content Footnotes

Support formatted text in footnotes.

```typescript
interface RichFootnoteContent {
  runs: Array<{
    text: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
  }>;
}

async function insertRichFootnote(
  context: WordRequestContext,
  ref: Ref,
  content: RichFootnoteContent
): Promise<Ref>;
```

### Footnote Search

Search within footnotes.

```typescript
interface FootnoteSearchResult {
  ref: Ref;
  matchText: string;
  context: string;  // Surrounding text
}

async function searchInFootnotes(
  context: WordRequestContext,
  query: string,
  options?: { regex?: boolean; caseSensitive?: boolean }
): Promise<FootnoteSearchResult[]>;

// Usage
const matches = await DocTree.searchInFootnotes(context, "Smith");
// Returns matches in all footnotes
```

### Footnote Statistics

Get footnote analytics.

```typescript
interface FootnoteStats {
  totalCount: number;
  averageLength: number;
  longestFootnote: Ref;
  shortestFootnote: Ref;
  emptyFootnotes: Ref[];
  orphanedFootnotes: Ref[];  // Notes not referenced in body
}

async function getFootnoteStats(
  context: WordRequestContext
): Promise<FootnoteStats>;
```

## Integration Points

### With Scope System

Enhance scope parsing for footnote-specific queries:

```typescript
// Existing scopes
"footnotes"     // All footnotes
"footnote:1"    // Specific footnote
"notes"         // Footnotes and endnotes

// Proposed additions
"footnotes:1-5"         // Range of footnotes
"footnotes:text=Smith"  // Footnotes containing text
"footnotes:long"        // Footnotes over N characters
```

### With Batch Edit

Add footnote operations to `batchEdit`:

```typescript
const result = await DocTree.batchEdit(context, [
  // Paragraph operations
  { ref: "p:3", operation: "replace", newText: "Updated" },

  // Footnote operations (proposed)
  { ref: "fn:1", operation: "editFootnote", newText: "New citation" },
  { ref: "p:5", operation: "insertFootnote", text: "Added note" },
  { ref: "fn:3", operation: "deleteFootnote" },
], { track: true });
```

### With YAML Serialization

Enhance footnote representation in YAML output:

```yaml
# Current output
footnotes:
  - ref: fn:1
    id: 1
    text: "Citation text..."
    referencedFrom: p:3

# Proposed enhancements
footnotes:
  - ref: fn:1
    id: 1
    text: "Citation text..."
    referencedFrom: p:3
    wordCount: 12
    hasFormatting: true
    hasTrackedChanges: false
    containsLinks: false
```

## Implementation Considerations

### Performance

1. **Minimize Sync Calls**: Batch footnote operations before `context.sync()`
2. **Lazy Loading**: Don't load footnote bodies until needed
3. **Cache Footnote Collection**: Reuse loaded collection within operation

### Error Handling

```typescript
// Proposed error types
class FootnoteNotFoundError extends Error {
  constructor(public ref: Ref) {
    super(`Footnote not found: ${ref}`);
  }
}

class FootnoteOperationError extends Error {
  constructor(
    public operation: string,
    public ref: Ref,
    public cause: Error
  ) {
    super(`Failed to ${operation} ${ref}: ${cause.message}`);
  }
}
```

### Ref Stability

After footnote insertions/deletions, refs change due to renumbering. Options:

1. **Rebuild Tree**: Full refresh after any modification
2. **Incremental Update**: Track changes and update refs
3. **Stable IDs**: Use Word's internal footnote IDs instead of positional refs

Recommendation: Start with option 1 (rebuild), optimize later if needed.

## Open Questions

1. **Should we use position-based refs (`fn:1`) or ID-based refs (`fn:id:12345`)?**
   - Position is intuitive but changes on insert/delete
   - IDs are stable but less readable

2. **How to handle footnotes in headers/footers?**
   - Office.js may have different behavior for non-body footnotes

3. **Should tracked changes in footnotes be visible in tree output?**
   - Adds complexity but useful for legal documents

4. **Priority for endnote vs footnote features?**
   - Most documents use footnotes; endnotes less common
   - Should parallel implementation be default?

## Related Work

- Python library has full footnote CRUD: `src/python_docx_redline/operations/notes.py`
- Existing scope parsing: `office-bridge/src/accessibility/scope.ts`
- Tree builder footnote extraction: `office-bridge/src/accessibility/builder.ts`
