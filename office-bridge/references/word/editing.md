---
name: office-bridge-editing
description: "Ref-based editing API for Word documents via Office Bridge. Use when making targeted edits to Word documents using refs from the DocTree accessibility tree, including replace, insert, delete, format, and batch operations."
---

# Office Bridge: Editing API

This document covers the editing functions in the Office Bridge DocTree API. All editing operations use refs (references) from the accessibility tree to target specific document elements, eliminating ambiguous text matching.

## Overview

The editing API provides:
- **Single-element operations**: Replace, insert, delete, format by ref
- **Batch operations**: Multiple edits in a single transaction
- **Scope-aware editing**: Edit all elements matching a scope filter

All operations support tracked changes via the `EditOptions` interface.

## EditOptions Interface

Most editing functions accept an options object:

```typescript
interface EditOptions {
  /** Enable tracked changes for this edit */
  track?: boolean;
  /** Author name for tracked changes */
  author?: string;
  /** Comment to attach to the change */
  comment?: string;
}
```

When `track: true` is set, Word records the edit as a tracked change that can be accepted or rejected later.

## Ref-Based Editing Functions

### replaceByRef(context, ref, newText, options?)

Replace the entire text content of a paragraph.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.replaceByRef(
    context,
    "p:5",
    "This paragraph has been completely replaced.",
    { track: true }
  );

  if (result.success) {
    console.log("Replaced content at:", result.newRef);
  } else {
    console.error("Replace failed:", result.error);
  }
});
```

**Parameters:**
- `context` - Word.RequestContext from Office.js
- `ref` - Reference to the paragraph (e.g., `"p:5"`, `"tbl:0/row:1/cell:2/p:0"`)
- `newText` - New text content to replace with
- `options` - Optional EditOptions

**Returns:** `EditResult` with `success`, `newRef`, and optional `error`

**Supported ref types:**
- `p:N` - Top-level paragraphs
- `tbl:T/row:R/cell:C/p:P` - Paragraphs inside table cells

### insertAfterRef(context, ref, content, options?)

Insert text at the end of a paragraph (after existing content).

```typescript
await Word.run(async (context) => {
  // Add amendment notation to end of paragraph
  const result = await DocTree.insertAfterRef(
    context,
    "p:5",
    " (amended December 2024)",
    { track: true }
  );
});
```

**Parameters:**
- `context` - Word.RequestContext
- `ref` - Reference to the paragraph
- `content` - Text to insert after existing content
- `options` - Optional EditOptions

### insertBeforeRef(context, ref, content, options?)

Insert text at the beginning of a paragraph (before existing content).

```typescript
await Word.run(async (context) => {
  // Add a note prefix
  const result = await DocTree.insertBeforeRef(
    context,
    "p:5",
    "Note: ",
    { track: true }
  );
});
```

### insertParagraphAfterRef(context, ref, paragraphText, options?)

Insert a completely new paragraph after the referenced paragraph.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.insertParagraphAfterRef(
    context,
    "p:10",
    "This is a new paragraph inserted after paragraph 10.",
    { track: true }
  );

  // New paragraph gets ref p:11 (indices shift)
  console.log("New paragraph at:", result.newRef);
});
```

### insertParagraphBeforeRef(context, ref, paragraphText, options?)

Insert a new paragraph before the referenced paragraph.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.insertParagraphBeforeRef(
    context,
    "p:5",
    "This paragraph appears before the original p:5.",
    { track: true }
  );

  // The new paragraph takes ref p:5, original shifts to p:6
});
```

### deleteByRef(context, ref, options?)

Delete an element by its ref.

```typescript
await Word.run(async (context) => {
  // Delete a paragraph
  const result = await DocTree.deleteByRef(context, "p:5", { track: true });

  // Delete an entire table
  const tableResult = await DocTree.deleteByRef(context, "tbl:0", { track: true });

  // Delete a table row
  const rowResult = await DocTree.deleteByRef(
    context,
    "tbl:0/row:2",
    { track: true }
  );
});
```

**Supported ref types for deletion:**
- `p:N` - Paragraphs
- `tbl:N` - Entire tables
- `tbl:T/row:R` - Table rows
- `tbl:T/row:R/cell:C/p:P` - Paragraphs in table cells

### formatByRef(context, ref, formatting)

Apply formatting to an element.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.formatByRef(context, "p:3", {
    bold: true,
    italic: false,
    underline: true,
    color: "#0000FF",      // Blue text
    highlight: "yellow",   // Yellow highlight
    size: 14,              // 14pt font
    font: "Arial",         // Font family
    style: "Heading 2"     // Paragraph style
  });
});
```

**FormatOptions interface:**

```typescript
interface FormatOptions {
  /** Apply bold */
  bold?: boolean;
  /** Apply italic */
  italic?: boolean;
  /** Apply underline */
  underline?: boolean;
  /** Apply strikethrough */
  strikethrough?: boolean;
  /** Font name */
  font?: string;
  /** Font size in points */
  size?: number;
  /** Font color (hex, e.g., "#FF0000") */
  color?: string;
  /** Highlight color (e.g., "yellow", "green") */
  highlight?: string;
  /** Paragraph style name */
  style?: string;
}
```

### getTextByRef(context, ref)

Retrieve the text content at a ref without modifying it.

```typescript
await Word.run(async (context) => {
  const text = await DocTree.getTextByRef(context, "p:5");
  if (text !== undefined) {
    console.log("Paragraph content:", text);
  } else {
    console.log("Ref not found or has no text");
  }

  // Also works for footnotes
  const footnoteText = await DocTree.getTextByRef(context, "fn:1");
});
```

**Returns:** The text content as a string, or `undefined` if not found.

## Batch Operations

Batch operations execute multiple edits in a single transaction, minimizing round-trips to the Word application and ensuring atomic execution.

### batchEdit(context, operations, options?)

Execute multiple operations of different types in one call.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.batchEdit(context, [
    { ref: "p:3", operation: "replace", newText: "Updated introduction" },
    { ref: "p:7", operation: "replace", newText: "New conclusion text" },
    { ref: "p:5", operation: "insertAfter", insertText: " (amended)" },
    { ref: "p:8", operation: "insertBefore", insertText: "Important: " },
    { ref: "p:12", operation: "delete" },
    { ref: "p:15", operation: "delete" },
  ], { track: true });

  console.log(`Operations: ${result.successCount} succeeded, ${result.failedCount} failed`);

  // Check individual results
  for (let i = 0; i < result.results.length; i++) {
    const r = result.results[i];
    if (!r.success) {
      console.error(`Operation ${i} failed:`, r.error);
    }
  }
});
```

**BatchEditOperation interface:**

```typescript
interface BatchEditOperation {
  /** Reference to the element to edit */
  ref: Ref;
  /** Operation type */
  operation: 'replace' | 'delete' | 'insertAfter' | 'insertBefore';
  /** New text content (required for 'replace') */
  newText?: string;
  /** Text to insert (required for 'insertAfter'/'insertBefore') */
  insertText?: string;
}
```

**BatchEditResult interface:**

```typescript
interface BatchEditResult {
  /** true if all operations succeeded */
  success: boolean;
  /** Number of successful operations */
  successCount: number;
  /** Number of failed operations */
  failedCount: number;
  /** Individual result for each operation */
  results: EditResult[];
}
```

**Important:** Deletions are automatically sorted to execute in reverse index order, preventing index shifting issues.

### batchReplace(context, replacements, options?)

Convenience function for replacing multiple paragraphs with different texts.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.batchReplace(context, [
    { ref: "p:3", text: "First replacement" },
    { ref: "p:7", text: "Second replacement" },
    { ref: "p:12", text: "Third replacement" },
  ], { track: true });

  console.log(`Replaced ${result.successCount} paragraphs`);
});
```

### batchDelete(context, refs, options?)

Convenience function for deleting multiple elements.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.batchDelete(
    context,
    ["p:10", "p:15", "p:20", "p:25"],
    { track: true }
  );

  console.log(`Deleted ${result.successCount} paragraphs`);
});
```

The refs are automatically sorted and processed in reverse order to maintain correct indices.

## Scope-Aware Editing

Scope-aware functions operate on all elements matching a scope filter. See [office-bridge.md](../office-bridge.md) for scope syntax details.

### replaceByScope(context, tree, scope, newText, options?)

Replace text in all nodes matching a scope.

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Replace all paragraphs in "Methods" section with the same text
  const results = await DocTree.replaceByScope(
    context,
    tree,
    "section:Methods",
    "Content has been redacted.",
    { track: true }
  );

  console.log(`Replaced ${results.filter(r => r.success).length} paragraphs`);
});
```

**Returns:** Array of `EditResult` for each matched node.

### deleteByScope(context, tree, scope, options?)

Delete all nodes matching a scope.

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Delete all paragraphs containing "DRAFT"
  const results = await DocTree.deleteByScope(
    context,
    tree,
    "DRAFT",
    { track: true }
  );

  // Delete all paragraphs with tracked changes
  const changedResults = await DocTree.deleteByScope(
    context,
    tree,
    { hasChanges: true },
    { track: true }
  );
});
```

**Note:** Deletions are automatically processed in reverse index order.

### formatByScope(context, tree, scope, formatting)

Apply formatting to all nodes matching a scope.

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Make all headings in "Results" section bold and blue
  const results = await DocTree.formatByScope(
    context,
    tree,
    { section: "Results", role: "heading" },
    { bold: true, color: "#0000FF" }
  );

  // Highlight all paragraphs containing "important"
  const highlightResults = await DocTree.formatByScope(
    context,
    tree,
    "important",
    { highlight: "yellow" }
  );
});
```

### searchReplaceByScope(context, tree, scope, searchText, replaceText, options?)

Find and replace text within nodes matching a scope.

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Replace "Plaintiff" with "Defendant" only in "Parties" section
  const results = await DocTree.searchReplaceByScope(
    context,
    tree,
    "section:Parties",
    "Plaintiff",
    "Defendant",
    { track: true }
  );

  // Case-sensitive replacement in entire document
  const globalResults = await DocTree.searchReplaceByScope(
    context,
    tree,
    {},  // Empty scope matches all
    "ACME Corp",
    "Globex Industries",
    { track: true }
  );
});
```

**Note:** Only nodes whose text contains `searchText` are modified.

## Scope Specification Formats

### String Shortcuts

```typescript
"keyword"                // Paragraphs containing "keyword"
"section:Introduction"   // Paragraphs in "Introduction" section
"role:heading"           // All headings
"style:Normal"           // Paragraphs with "Normal" style
"footnotes"              // All footnotes
"footnote:1"             // Specific footnote by ID
"endnotes"               // All endnotes
"endnote:2"              // Specific endnote by ID
"changes"                // Paragraphs with tracked changes
"comments"               // Paragraphs with comments
"level:2"                // Level 2 headings
```

### Object Format (AND logic)

All specified conditions must match:

```typescript
{
  contains: "payment",        // Text must contain "payment"
  notContains: "Exhibit",     // Text must NOT contain "Exhibit"
  section: "Definitions",     // Must be in "Definitions" section
  role: "paragraph",          // Must be a paragraph
  style: "Normal",            // Must have "Normal" style
  hasChanges: true,           // Must have tracked changes
  hasComments: false,         // Must NOT have comments
  level: [1, 2],              // Must be heading level 1 or 2
  minLength: 50,              // Text length >= 50
  maxLength: 500,             // Text length <= 500
  pattern: "\\$[\\d,]+",      // Must match regex
  refs: ["p:1", "p:2", "p:3"] // Must be one of these refs
}
```

### Predicate Function

For complex custom logic:

```typescript
const customScope = (node: AccessibilityNode) => {
  // Custom matching logic
  return node.text?.startsWith("WHEREAS") &&
         node.level === undefined;  // Not a heading
};

const results = await DocTree.deleteByScope(
  context,
  tree,
  customScope,
  { track: true }
);
```

## Error Handling

All editing functions return result objects with success indicators:

```typescript
interface EditResult {
  /** Whether the operation succeeded */
  success: boolean;
  /** The ref after the operation (may change for inserts) */
  newRef?: Ref;
  /** Error message if failed */
  error?: string;
}
```

### Common Error Patterns

```typescript
await Word.run(async (context) => {
  // Single operation error handling
  const result = await DocTree.replaceByRef(context, "p:999", "text");
  if (!result.success) {
    if (result.error?.includes("out of range")) {
      console.error("Paragraph does not exist");
    } else if (result.error?.includes("Invalid ref")) {
      console.error("Malformed ref format");
    } else {
      console.error("Unknown error:", result.error);
    }
  }

  // Batch operation error handling
  const batchResult = await DocTree.batchEdit(context, operations);
  if (!batchResult.success) {
    // Find which operations failed
    batchResult.results.forEach((r, i) => {
      if (!r.success) {
        console.error(`Operation ${i} failed:`, r.error);
      }
    });

    // Decide whether to proceed with partial success
    if (batchResult.successCount > 0) {
      console.log(`${batchResult.successCount} operations succeeded`);
    }
  }
});
```

### Ref Validation

The system validates refs before operations:

- Invalid format (e.g., `"invalid"`, `"p:"`, `"p:abc"`) - throws error
- Out of range (e.g., `"p:999"` in 50-paragraph doc) - returns error result
- Unsupported type (e.g., `"img:0"` for replace) - returns error result
- Fingerprint refs (`"p:~xK4mNp2q"`) - not yet supported

## Performance Best Practices

### Use Batch Operations

Instead of multiple individual calls:

```typescript
// BAD: Multiple sync calls
for (const ref of refsToUpdate) {
  await DocTree.replaceByRef(context, ref, newText);
}

// GOOD: Single batch call
await DocTree.batchEdit(context,
  refsToUpdate.map(ref => ({ ref, operation: 'replace', newText }))
);
```

### Scope-Aware Functions Are Optimized

All scope-aware functions use batched `context.sync()` calls internally:

```typescript
// This is already optimized - single sync for loading, single sync for operations
await DocTree.replaceByScope(context, tree, "section:Methods", newText);
```

### Minimize Tree Rebuilding

Build the tree once and reuse it:

```typescript
// GOOD: Build once, use multiple times
const tree = await DocTree.buildTree(context);
await DocTree.replaceByScope(context, tree, "section:Intro", text1);
await DocTree.formatByScope(context, tree, "section:Intro", formatting);

// BAD: Rebuilding tree for each operation
await DocTree.replaceByScope(context, await DocTree.buildTree(context), scope, text);
await DocTree.formatByScope(context, await DocTree.buildTree(context), scope, fmt);
```

### Deletion Order

When deleting multiple paragraphs, process in reverse order to avoid index shifting:

```typescript
// batchDelete handles this automatically
await DocTree.batchDelete(context, ["p:5", "p:10", "p:15"]);
// Internally processes as p:15, p:10, p:5

// If doing manual deletions, sort descending
const refs = ["p:5", "p:10", "p:15"];
refs.sort((a, b) => {
  const aIdx = parseInt(a.split(':')[1]);
  const bIdx = parseInt(b.split(':')[1]);
  return bIdx - aIdx;
});
for (const ref of refs) {
  await DocTree.deleteByRef(context, ref);
}
```

## Performance Characteristics

| Operation | Sync Calls | Notes |
|-----------|------------|-------|
| Single edit | 2-3 | Load paragraphs, execute, sync |
| batchEdit (N ops) | 2-3 | Same regardless of N |
| Scope-aware (N matches) | 2-3 | Same regardless of N |
| Tree building | 3-5 | More for large documents |

The batched architecture means editing 100 paragraphs is nearly as fast as editing 1.

## See Also

- [office-bridge.md](../office-bridge.md) - Main Office Bridge API documentation
- [selection.md](./selection.md) - Selection-based operations
- [tables.md](./tables.md) - Table-specific editing
- [footnotes.md](./footnotes.md) - Footnote and endnote editing
