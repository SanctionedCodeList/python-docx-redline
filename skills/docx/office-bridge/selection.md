---
name: office-bridge-selection
description: "Office.js Selection and Range API for cursor operations. Use when working with user selection, navigating documents, or performing selection-based edits in Word through Office Bridge."
---

# Office Bridge: Selection and Cursor Operations

This document covers the Office.js APIs for working with selection and cursor position in Word documents through the Office Bridge add-in.

## Overview

Office.js provides two primary mechanisms for working with selected content:

1. **Word.Range** - The main API for document manipulation (preferred for programmatic edits)
2. **Word.Selection** - Desktop-only API for physically changing what the user sees selected (WordApiDesktop 1.4+)

The key guidance from Microsoft: Use Range objects when you do not need to physically change the current selection. Selection is for user-visible navigation.

## Getting the Current Selection

### Basic Selection Retrieval

```typescript
await Word.run(async (context) => {
  // Get the current selection as a Range
  const selection = context.document.getSelection();
  selection.load('text,font,paragraphs');
  await context.sync();

  console.log('Selected text:', selection.text);
  console.log('Is empty:', selection.text === '');
});
```

### Selection Properties

Key properties available on the selection range:

| Property | Type | Description |
|----------|------|-------------|
| `text` | string | The text content of the selection |
| `isEmpty` | boolean | True if selection length is zero (cursor only) |
| `font` | Word.Font | Character formatting |
| `paragraphs` | ParagraphCollection | Paragraphs in selection |
| `start` | number | Starting character position (Desktop only) |
| `end` | number | Ending character position (Desktop only) |

### Checking Selection Context

Determine where the cursor is positioned:

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  const parentBody = selection.parentBody;
  parentBody.load('type');
  await context.sync();

  // Returns: "MainDoc" | "Header" | "Footer" | "Footnote" | etc.
  console.log('Selection is in:', parentBody.type);

  // Check if in a content control
  const cc = selection.parentContentControlOrNullObject;
  cc.load('id,title');
  await context.sync();

  if (!cc.isNullObject) {
    console.log('Inside content control:', cc.title);
  }
});
```

## Programmatically Selecting Content

### Select by Ref (DocTree Pattern)

Select a specific paragraph using our ref system:

```typescript
await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  // Select paragraph at index 5
  const target = paragraphs.items[5];
  const range = target.getRange('Whole');

  // Navigate UI to this location
  range.select(Word.SelectionMode.select);
  await context.sync();
});
```

### SelectionMode Options

Control where the cursor ends up after selection:

| Mode | Effect |
|------|--------|
| `Word.SelectionMode.select` | Select the entire range (default) |
| `Word.SelectionMode.start` | Move cursor to start without selecting |
| `Word.SelectionMode.end` | Move cursor to end without selecting |

```typescript
// Move cursor to end of document without selecting
context.document.body.paragraphs.getLast().select(Word.SelectionMode.end);
```

### Scrolling to Content

The `select()` method automatically scrolls the Word UI to make the selected range visible:

```typescript
await Word.run(async (context) => {
  // Find text and scroll to it
  const results = context.document.body.search('Section 5.2');
  results.load('items');
  await context.sync();

  if (results.items.length > 0) {
    // This scrolls the view and selects the text
    results.items[0].select();
    await context.sync();
  }
});
```

## Insert and Replace at Selection

### Insert Text at Selection

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();

  // Insert at different positions
  selection.insertText('Before ', Word.InsertLocation.before);
  selection.insertText(' After', Word.InsertLocation.after);
  selection.insertText('Replaced', Word.InsertLocation.replace);
  selection.insertText('Start ', Word.InsertLocation.start);
  selection.insertText(' End', Word.InsertLocation.end);

  await context.sync();
});
```

### InsertLocation Options

| Location | Effect |
|----------|--------|
| `before` | Insert before the range, outside it |
| `after` | Insert after the range, outside it |
| `start` | Insert at the start, inside the range |
| `end` | Insert at the end, inside the range |
| `replace` | Replace the entire range content |

### Insert HTML at Selection

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.insertHtml('<b>Bold text</b> and <i>italic</i>', Word.InsertLocation.replace);
  await context.sync();
});
```

### Insert Paragraph at Selection

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();

  // Insert a new paragraph after selection
  const newPara = selection.insertParagraph('New paragraph content', Word.InsertLocation.after);
  newPara.style = 'Normal';

  await context.sync();
});
```

## Common API for Read/Write Selection

The Common API provides simpler methods for basic selection operations:

```typescript
// Read selection as text
Office.context.document.getSelectedDataAsync(
  Office.CoercionType.Text,
  (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log('Selected:', result.value);
    }
  }
);

// Write to selection
Office.context.document.setSelectedDataAsync(
  'Replacement text',
  (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log('Text inserted');
    }
  }
);
```

## Range Expansion and Manipulation

### Expand Selection to Word/Sentence/Paragraph

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();

  // Get text ranges within selection by delimiters
  const sentences = selection.getTextRanges(['.', '?', '!'], true);
  sentences.load('items,text');
  await context.sync();

  console.log(`Selection contains ${sentences.items.length} sentences`);
});
```

### Expand to Cover Another Range

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  // Expand selection to cover from current to a target paragraph
  const targetRange = paragraphs.items[10].getRange('Whole');
  const expandedRange = selection.expandTo(targetRange);

  expandedRange.select();
  await context.sync();
});
```

### Compare Range Locations

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  const targetPara = context.document.body.paragraphs.getFirst().getRange('Whole');

  const comparison = selection.compareLocationWith(targetPara);
  await context.sync();

  // Returns: Before | InsideStart | Inside | InsideEnd | After |
  //          AdjacentBefore | AdjacentAfter | Contains | Equals
  console.log('Selection is', comparison.value, 'target');
});
```

## Navigation with Selection

### Move Cursor by Position

Using Selection class (Desktop only):

```typescript
await Word.run(async (context) => {
  // This requires WordApiDesktop 1.4+
  const selection = context.document.getSelection() as any; // Cast for desktop API

  // Collapse to end
  selection.collapse('End');

  // Move by words
  selection.move('Word', 5);  // Move 5 words forward
  selection.select();

  await context.sync();
});
```

### Navigate to Specific Content

```typescript
await Word.run(async (context) => {
  // Go to next heading (Desktop only)
  const selection = context.document.getSelection() as any;
  const nextHeading = selection.goToNext('Heading');
  nextHeading.select();

  await context.sync();
});
```

## Known Limitations

### Table Selection Bug

When selecting multiple rows/columns (but not the full table), `getSelection()` only returns the last row's data. This is a known Office.js issue.

**Workaround**: Select the entire table or process row by row.

### Word 2016 select() Issue

In Word 2016 on Windows, `range.select()` may not visually change the selection. This works correctly in Word 2019, Word 365, and Word Online.

### No Direct Cursor Position API

Office.js does not provide direct cursor position control. Workarounds:
- Use `select(SelectionMode.end)` to position cursor
- Use search to find and select specific content

### Desktop-Only Features

Many Selection class features require WordApiDesktop 1.4+:
- `start` and `end` properties
- Movement methods (`move`, `moveStart`, `moveEnd`)
- Navigation methods (`goToNext`, `goToPrevious`)
- Clipboard operations (`cut`, `copy`, `paste`)

Check availability:

```typescript
if (Office.context.requirements.isSetSupported('WordApiDesktop', '1.4')) {
  // Desktop features available
}
```

### Multi-Selection Not Supported

Office.js does not support multiple discontinuous selections. Only a single contiguous selection is available.

## Best Practices

### Use Range Over Selection

Prefer Range objects for programmatic edits:

```typescript
// Preferred: Edit without changing visible selection
await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  paragraphs.items[5].insertText('Updated ', 'Start');
  await context.sync();
});

// Avoid: Unnecessarily moving the user's selection
await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  paragraphs.items[5].select();  // Moves user's cursor
  const selection = context.document.getSelection();
  selection.insertText('Updated ', 'Start');
  await context.sync();
});
```

### Batch Selection Operations

When multiple selections are needed, queue them in a single sync:

```typescript
await Word.run(async (context) => {
  // Queue multiple range loads
  const searchResults = context.document.body.search('important');
  searchResults.load('items');
  await context.sync();

  // Highlight all at once
  for (const item of searchResults.items) {
    item.font.highlightColor = 'Yellow';
  }
  await context.sync();  // Single sync for all changes
});
```

### Track Object Lifecycle

Selection objects can become invalid. Always get fresh selections when needed:

```typescript
await Word.run(async (context) => {
  // Get selection, make edits
  let selection = context.document.getSelection();
  selection.insertText('New text', 'Replace');
  await context.sync();

  // Get fresh selection for next operation
  selection = context.document.getSelection();
  selection.load('text');
  await context.sync();
});
```

## Integration with DocTree Refs

Combining selection with ref-based editing:

```typescript
await Word.run(async (context) => {
  // Get current selection
  const selection = context.document.getSelection();
  selection.load('paragraphs');
  await context.sync();

  // Find the ref for the selected paragraph
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  const selectedPara = selection.paragraphs.items[0];
  const refIndex = paragraphs.items.findIndex(p =>
    p._ReferenceId === selectedPara._ReferenceId
  );

  if (refIndex >= 0) {
    const ref = `p:${refIndex}`;
    console.log('Selected paragraph ref:', ref);

    // Now use ref-based editing
    await DocTree.replaceByRef(context, ref, 'Updated content', { track: true });
  }
});
```

## See Also

- [office-bridge.md](../office-bridge.md) - Main Office Bridge documentation
- [editing.md](../editing.md) - Ref-based editing patterns
- [Microsoft Word.Range documentation](https://learn.microsoft.com/en-us/javascript/api/word/word.range)
- [Microsoft Word.Selection documentation](https://learn.microsoft.com/en-us/javascript/api/word/word.selection)
