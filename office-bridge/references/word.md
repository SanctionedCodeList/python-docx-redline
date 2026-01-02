# Word Patterns

Word has full accessibility tree and ref-based editing helpers in addition to raw Office.js.

## Auto-Scroll

The add-in includes an "Auto-scroll to edits" toggle (enabled by default). When on, the document automatically scrolls to show the current selection after each code execution. This works for both helper methods and raw Office.js calls.

## Getting Documents

```typescript
const bridge = await connect();
const docs = await bridge.documents();  // Returns WordDocument[]
const doc = docs[0];
```

## Accessibility Tree

Get a semantic representation of the document:

```typescript
// YAML representation
const tree = await doc.getTree({ verbosity: 'minimal' });   // Structure only
const tree = await doc.getTree({ verbosity: 'standard' });  // Content + metadata
const tree = await doc.getTree({ verbosity: 'full' });      // Run-level detail

// Raw object
const treeObj = await doc.getTreeRaw();
```

## Ref-Based Editing

Elements have stable refs like `p:3`, `tbl:0/row:2/cell:1`:

```typescript
// Replace text
await doc.replaceByRef('p:3', 'New paragraph text', { track: true });

// Insert relative to ref
await doc.insertAfterRef('p:5', ' (amended)', { track: true });
await doc.insertBeforeRef('p:5', 'Note: ', { track: true });

// Delete
await doc.deleteByRef('p:7', { track: true });

// Format
await doc.formatByRef('p:3', { bold: true, color: '#0000FF' });

// Read
const text = await doc.getTextByRef('p:3');
```

## Raw Office.js Patterns

For operations not covered by helpers:

```javascript
// Insert text at end
const body = context.document.body;
body.insertText("Hello!", Word.InsertLocation.end);
await context.sync();
return "Done";

// Get all paragraphs
const paragraphs = context.document.body.paragraphs;
paragraphs.load("items/text");
await context.sync();
return paragraphs.items.map(p => p.text);

// Search and replace
const results = context.document.body.search("old text", { matchCase: true });
results.load("items");
await context.sync();
for (const item of results.items) {
  item.insertText("new text", Word.InsertLocation.replace);
}
await context.sync();
return { replaced: results.items.length };

// Apply formatting
const results = context.document.body.search("Important");
results.load("items");
await context.sync();
for (const item of results.items) {
  item.font.bold = true;
  item.font.color = "red";
}
await context.sync();
```

## Tracked Changes

```typescript
// Edit options
{ track: true }           // Enable tracking
{ track: true, author: "Claude" }  // With author

// Manage changes via raw JS
await doc.executeJs(`
  context.document.body.paragraphs.getFirst().getRange().track();
  await context.sync();
`);
```
