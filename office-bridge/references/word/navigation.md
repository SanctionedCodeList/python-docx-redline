---
name: office-bridge-navigation
description: "Navigation helpers and document summary functions for the Office Bridge Word Add-in. Use when traversing document structure, finding sections, calculating word counts, or getting document statistics."
---

# Office Bridge: Navigation & Document Summary

This document covers navigation helpers and document summary functions in the Office Bridge DocTree API. These utilities help traverse document structure, find sections, work with ref ranges, and gather document statistics.

## Navigation Helpers

Navigation helpers provide utility functions for working with paragraph refs without requiring async operations or context synchronization.

### getNextRef / getPrevRef

Get adjacent paragraph refs for sequential navigation.

```typescript
// Get next ref
const next = DocTree.getNextRef("p:5");      // Returns "p:6"
const next2 = DocTree.getNextRef("p:5", 10); // Returns "p:6" (within bounds)
const atEnd = DocTree.getNextRef("p:9", 10); // Returns null (out of bounds)

// Get previous ref
const prev = DocTree.getPrevRef("p:5");      // Returns "p:4"
const atStart = DocTree.getPrevRef("p:0");   // Returns null (already at start)
```

**Parameters:**
- `ref`: Current paragraph ref (e.g., "p:5")
- `totalParagraphs` (optional): Total paragraph count for bounds checking

**Returns:** Next/previous ref string, or `null` if at boundary

**Use Cases:**
- Stepping through document sequentially
- Building prev/next navigation UI
- Iterating over paragraphs in a loop

### getSiblingRefs(ref, totalParagraphs?)

Get both previous and next refs in a single call.

```typescript
const { prev, next } = DocTree.getSiblingRefs("p:5");
// prev = "p:4", next = "p:6"

const { prev: p, next: n } = DocTree.getSiblingRefs("p:0", 100);
// p = null (at start), n = "p:1"

const { prev: p2, next: n2 } = DocTree.getSiblingRefs("p:99", 100);
// p2 = "p:98", n2 = null (at end)
```

**Parameters:**
- `ref`: Current paragraph ref
- `totalParagraphs` (optional): Total paragraph count for bounds checking

**Returns:** Object with `prev` and `next` refs (each may be `null`)

### getSectionForRef(tree, ref)

Find the section heading that contains a given paragraph. This walks backwards through the accessibility tree to find the nearest heading-style paragraph.

```typescript
const tree = await DocTree.buildTree(context);

const section = DocTree.getSectionForRef(tree, "p:25");
if (section) {
  console.log(`Paragraph p:25 is in section "${section.headingText}"`);
  console.log(`Section level: ${section.level}`);
  console.log(`Heading ref: ${section.headingRef}`);
}
// Example output:
// Paragraph p:25 is in section "Methods"
// Section level: 2
// Heading ref: p:18
```

**Parameters:**
- `tree`: AccessibilityTree from `buildTree()`
- `ref`: Reference to the paragraph to find section for

**Returns:** Section info object or `null` if not in a section
```typescript
{
  headingRef: Ref;      // Ref to the heading paragraph
  headingText: string;  // Text content of the heading
  level: number;        // Heading level (1-9)
}
```

**Section Detection Algorithm:**

The function uses the following algorithm to detect sections:

1. Parse the input ref to get the paragraph index
2. Walk backwards from that index through `tree.content`
3. For each node, check if it's a heading by:
   - Checking if `node.role === 'heading'`
   - OR checking if `node.style.name` matches `/^Heading\s*\d*$/i`
4. If a heading is found:
   - Extract level from style name (e.g., "Heading 2" -> level 2)
   - Default to level 1 if no number found
   - Return the heading info
5. If no heading found before index 0, return `null`

**Heading Detection Patterns:**
- `Heading 1`, `Heading 2`, ... `Heading 9`
- `heading1`, `heading2`, etc. (case-insensitive)
- Any style with `role: 'heading'` in the accessibility node

### getRefRange(startRef, endRef)

Get all paragraph refs between two refs (inclusive). Useful for selecting a range of paragraphs.

```typescript
const refs = DocTree.getRefRange("p:5", "p:10");
// Returns ["p:5", "p:6", "p:7", "p:8", "p:9", "p:10"]

// Order doesn't matter - automatically handles reversed ranges
const refs2 = DocTree.getRefRange("p:10", "p:5");
// Returns ["p:5", "p:6", "p:7", "p:8", "p:9", "p:10"]

// Single ref range
const single = DocTree.getRefRange("p:5", "p:5");
// Returns ["p:5"]
```

**Parameters:**
- `startRef`: First ref in range
- `endRef`: Last ref in range

**Returns:** Array of refs in ascending order

**Use Cases:**
- Selecting multiple paragraphs for batch operations
- Calculating word count for a section
- Applying formatting to a range

### isRefInRange(ref, startRef, endRef)

Check if a ref falls within a specified range (inclusive).

```typescript
DocTree.isRefInRange("p:7", "p:5", "p:10");   // true
DocTree.isRefInRange("p:5", "p:5", "p:10");   // true (inclusive start)
DocTree.isRefInRange("p:10", "p:5", "p:10");  // true (inclusive end)
DocTree.isRefInRange("p:11", "p:5", "p:10");  // false
DocTree.isRefInRange("p:4", "p:5", "p:10");   // false

// Works with reversed range arguments
DocTree.isRefInRange("p:7", "p:10", "p:5");   // true
```

**Parameters:**
- `ref`: Ref to check
- `startRef`: Start of range
- `endRef`: End of range

**Returns:** `boolean`

## Document Summary

Document summary functions provide statistics and metadata about the document.

### getDocumentSummary(context, tree?)

Get comprehensive document statistics in a single call.

```typescript
await Word.run(async (context) => {
  // Basic usage (without tree)
  const summary = await DocTree.getDocumentSummary(context);

  // With tree for more complete info (tracked changes detection)
  const tree = await DocTree.buildTree(context);
  const summaryWithTree = await DocTree.getDocumentSummary(context, tree);

  console.log(`
    Paragraphs: ${summary.paragraphCount}
    Tables: ${summary.tableCount}
    Words: ${summary.wordCount}
    Characters: ${summary.characterCount}
    Characters (no spaces): ${summary.characterCountNoSpaces}
    Headings: ${summary.headingCount}
    List Items: ${summary.listItemCount}
    Comments: ${summary.commentCount}
    Has Tracked Changes: ${summary.hasTrackedChanges}
    Sections: ${summary.sections.join(", ")}
  `);

  // Breakdown by heading level
  for (const [level, count] of Object.entries(summary.headingsByLevel)) {
    console.log(`Heading ${level}: ${count}`);
  }
});
```

**Parameters:**
- `context`: Word.RequestContext from Office.js
- `tree` (optional): AccessibilityTree for tracked changes detection

**Returns:** `DocumentSummary` object

### DocumentSummary Structure

```typescript
interface DocumentSummary {
  /** Total number of paragraphs */
  paragraphCount: number;

  /** Total number of tables */
  tableCount: number;

  /** Approximate word count (whitespace-delimited) */
  wordCount: number;

  /** Character count including spaces */
  characterCount: number;

  /** Character count excluding spaces */
  characterCountNoSpaces: number;

  /** Number of heading paragraphs */
  headingCount: number;

  /** Number of list item paragraphs */
  listItemCount: number;

  /** Whether document has tracked changes (requires tree) */
  hasTrackedChanges: boolean;

  /** Number of comments in document */
  commentCount: number;

  /** Count of headings by level: { 1: 3, 2: 5, 3: 2 } */
  headingsByLevel: Record<number, number>;

  /** Array of section names from heading text */
  sections: string[];
}
```

**Calculation Notes:**
- **Word count**: Split by whitespace (`/\s+/`), filter empty strings
- **Character count**: Raw string length
- **Headings**: Detected by style name matching `/^Heading\s*\d*$/i`
- **List items**: Detected by style name starting with "List"
- **Tracked changes**: Only detected if `tree` parameter provided
- **Sections**: Extracted from heading text content (trimmed)

### getWordCount(context, refs)

Get word count for specific paragraphs by ref.

```typescript
await Word.run(async (context) => {
  // Single paragraph
  const count1 = await DocTree.getWordCount(context, "p:5");

  // Multiple paragraphs
  const count2 = await DocTree.getWordCount(context, ["p:5", "p:6", "p:7"]);

  // Range of paragraphs
  const range = DocTree.getRefRange("p:10", "p:20");
  const sectionCount = await DocTree.getWordCount(context, range);

  console.log(`Section word count: ${sectionCount}`);
});
```

**Parameters:**
- `context`: Word.RequestContext from Office.js
- `refs`: Single ref or array of refs

**Returns:** Total word count across all specified refs

**Use Cases:**
- Word count for a specific section
- Progress tracking (words edited)
- Document statistics by section

## Best Practices

### Efficient Navigation

```typescript
// GOOD: Use getSiblingRefs for bidirectional navigation
const { prev, next } = DocTree.getSiblingRefs(currentRef, totalParagraphs);

// GOOD: Use getRefRange + batch operations for ranges
const refs = DocTree.getRefRange("p:5", "p:10");
await DocTree.batchDelete(context, refs, { track: true });

// AVOID: Multiple sequential getTextByRef calls
// Use batch loading instead
```

### Caching Document Summary

```typescript
// Cache summary at start of multi-step operation
const tree = await DocTree.buildTree(context);
const summary = await DocTree.getDocumentSummary(context, tree);

// Use cached totalParagraphs for bounds checking
const { prev, next } = DocTree.getSiblingRefs(ref, summary.paragraphCount);
```

### Section-Aware Operations

```typescript
// Find all paragraphs in a section
async function getParagraphsInSection(
  context: WordRequestContext,
  tree: AccessibilityTree,
  sectionName: string
): Promise<Ref[]> {
  const refs: Ref[] = [];
  let inSection = false;
  let currentLevel = 0;

  for (const node of tree.content) {
    const isHeading = node.role === 'heading' ||
      (node.style?.name && /^Heading\s*\d*$/i.test(node.style.name));

    if (isHeading) {
      const match = node.style?.name?.match(/Heading\s*(\d+)/i);
      const level = match?.[1] ? parseInt(match[1], 10) : 1;

      if (node.text?.includes(sectionName)) {
        inSection = true;
        currentLevel = level;
        refs.push(node.ref);
      } else if (inSection && level <= currentLevel) {
        // New section at same or higher level, stop
        break;
      }
    } else if (inSection) {
      refs.push(node.ref);
    }
  }

  return refs;
}
```

### Range Validation

```typescript
// Always validate ranges before operations
function validateRange(startRef: Ref, endRef: Ref, totalParagraphs: number): boolean {
  const refs = DocTree.getRefRange(startRef, endRef);

  for (const ref of refs) {
    const match = ref.match(/^p:(\d+)$/);
    if (!match) return false;

    const index = parseInt(match[1], 10);
    if (index < 0 || index >= totalParagraphs) return false;
  }

  return true;
}
```

### Combining Navigation with Editing

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);
  const summary = await DocTree.getDocumentSummary(context, tree);

  // Find section containing "Conclusion"
  let conclusionRef: Ref | null = null;
  for (const node of tree.content) {
    if (node.text?.toLowerCase().includes('conclusion')) {
      conclusionRef = node.ref;
      break;
    }
  }

  if (conclusionRef) {
    // Get section info
    const section = DocTree.getSectionForRef(tree, conclusionRef);
    console.log(`Found in section: ${section?.headingText}`);

    // Get all refs from conclusion to end
    const lastRef = `p:${summary.paragraphCount - 1}`;
    const conclusionRefs = DocTree.getRefRange(conclusionRef, lastRef);

    // Calculate word count
    const wordCount = await DocTree.getWordCount(context, conclusionRefs);
    console.log(`Conclusion section: ${wordCount} words`);
  }
});
```

## See Also

- [office-bridge.md](../office-bridge.md) - Main Office Bridge documentation
- [selection.md](./selection.md) - Selection and cursor operations
- [tables.md](./tables.md) - Table navigation and editing
