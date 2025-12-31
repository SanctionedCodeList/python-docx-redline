---
name: office-bridge-footnotes
description: "Office.js footnote and endnote API reference for the Word Add-in. Use when working with footnotes and endnotes through Office Bridge, including insertion, deletion, and navigation operations."
---

# Office Bridge: Footnotes and Endnotes

This document covers working with footnotes and endnotes through Office.js in the Office Bridge Word add-in.

## API Availability

Footnote and endnote support requires **WordApi 1.5** or later. The API works through:

- `Word.NoteItem` - Represents a single footnote or endnote
- `Word.NoteItemCollection` - Collection of notes (accessed via `body.footnotes` or `body.endnotes`)
- `Word.Range.insertFootnote()` / `insertEndnote()` - Insert new notes

## Current Office Bridge Support

### What We Have

The Office Bridge accessibility layer currently provides:

1. **Footnote Extraction** - Extracts footnotes from OOXML during tree building
2. **Footnote Refs** - Uses `fn:1`, `fn:2`, etc. reference system
3. **Endnote Refs** - Uses `en:1`, `en:2`, etc. reference system
4. **Paragraph Mapping** - Tracks which paragraphs contain footnote references
5. **Text Retrieval** - `getTextByRef()` works with footnote refs

### Ref System

```yaml
# Footnotes in the accessibility tree
footnotes:
  - ref: fn:1
    id: 1
    text: "See Smith (2020) for details."
    referencedFrom: p:3
  - ref: fn:2
    id: 2
    text: "Additional methodology notes."
    referencedFrom: p:7

# Paragraph shows footnote refs
- ref: p:3
  role: paragraph
  text: "The study found significant results..."
  fn: [fn:1]  # This paragraph references footnote 1
```

### Getting Footnote Text

```typescript
// Get footnote text via ref
const text = await DocTree.getTextByRef(context, "fn:1");
console.log(text); // "See Smith (2020) for details."
```

## Office.js API Reference

### Word.NoteItem

Represents a footnote or endnote.

| Property | Type | Description |
|----------|------|-------------|
| `body` | `Word.Body` | Text content of the note |
| `reference` | `Word.Range` | Reference mark in the main document |
| `type` | `"Footnote" \| "Endnote"` | Note type |

| Method | Returns | Description |
|--------|---------|-------------|
| `delete()` | `void` | Delete the note |
| `getNext()` | `Word.NoteItem` | Get next note (throws if last) |
| `getNextOrNullObject()` | `Word.NoteItem` | Get next note or null |

### Accessing Collections

```typescript
await Word.run(async (context) => {
  // All footnotes in document body
  const footnotes = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  console.log(`Document has ${footnotes.items.length} footnotes`);

  // All endnotes in document body
  const endnotes = context.document.body.endnotes;
  endnotes.load("items");
  await context.sync();
});
```

### Getting Footnote Content

```typescript
await Word.run(async (context) => {
  const footnotes = context.document.body.footnotes;
  footnotes.load("items/body");
  await context.sync();

  // Get specific footnote (0-indexed)
  const footnoteIndex = 0; // First footnote
  const footnoteBody = footnotes.items[footnoteIndex].body.getRange();
  footnoteBody.load("text");
  await context.sync();

  console.log(`Footnote text: ${footnoteBody.text}`);
});
```

### Selecting a Footnote Reference

Navigate to where a footnote is referenced in the document:

```typescript
await Word.run(async (context) => {
  const footnotes = context.document.body.footnotes;
  footnotes.load("items/reference");
  await context.sync();

  // Select the reference mark for footnote 1
  const reference = footnotes.items[0].reference;
  reference.select();
  await context.sync();
});
```

### Inserting Footnotes

Use `Range.insertFootnote()` to add new footnotes:

```typescript
await Word.run(async (context) => {
  // Insert at current selection
  const selection = context.document.getSelection();
  const footnote = selection.insertFootnote("Citation text here.");
  await context.sync();

  console.log("Inserted footnote");
});

// Insert at specific text
await Word.run(async (context) => {
  const results = context.document.body.search("significant finding");
  results.load("items");
  await context.sync();

  if (results.items.length > 0) {
    const footnote = results.items[0].insertFootnote(
      "See methodology section for details."
    );
    await context.sync();
  }
});
```

### Inserting Endnotes

Same pattern as footnotes:

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  const endnote = selection.insertEndnote("End reference text.");
  await context.sync();
});
```

### Deleting Footnotes

```typescript
await Word.run(async (context) => {
  const footnotes = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  // Delete first footnote
  footnotes.items[0].delete();
  await context.sync();

  console.log("Footnote deleted");
});
```

### Navigating Between Footnotes

```typescript
await Word.run(async (context) => {
  const footnotes = context.document.body.footnotes;
  footnotes.load("items/reference");
  await context.sync();

  // Get next footnote from current position
  const currentIndex = 0;
  const nextFootnote = footnotes.items[currentIndex].getNextOrNullObject();
  nextFootnote.load("reference");
  await context.sync();

  if (nextFootnote.isNullObject) {
    console.log("No more footnotes");
  } else {
    nextFootnote.reference.select();
    console.log("Selected next footnote");
  }
});
```

## Scope-Based Querying

Use scope strings to filter for footnotes:

```typescript
// Parse scope to check for note targeting
import { parseNoteScope, isNoteScope } from './accessibility/scope';

// Check if scope targets notes
if (isNoteScope("footnotes")) {
  // Handle footnote-specific query
}

// Parse specific scopes
parseNoteScope("footnotes");     // { scopeType: 'footnotes' }
parseNoteScope("footnote:1");    // { scopeType: 'footnote', noteId: '1' }
parseNoteScope("endnotes");      // { scopeType: 'endnotes' }
parseNoteScope("endnote:2");     // { scopeType: 'endnote', noteId: '2' }
parseNoteScope("notes");         // { scopeType: 'notes' } (both types)
```

## Limitations and Gotchas

### API Set Requirements

- Footnote APIs require **WordApi 1.5** (released in Word requirement set 1.5)
- Some features only work on desktop via **WordApiDesktop 1.4**
- Web version may have limited support

### Known Issues

1. **UI Flickering**: Updating footnote text programmatically can cause UI flickering in some Word versions (see [GitHub issue #2217](https://github.com/OfficeDev/office-js/issues/2217))

2. **NotImplemented Errors**: Some older Word versions return "NotImplemented" for footnote operations even when API set is declared

3. **Index vs ID**: Footnotes are accessed by 0-based index in the collection, but Word displays them as 1-indexed to users. Our ref system uses 1-indexed (`fn:1`, `fn:2`) to match Word's display.

4. **Renumbering**: When deleting a footnote, Word automatically renumbers remaining footnotes. The ref system may need refreshing after deletions.

5. **Rich Text**: Setting footnote body text via the API may strip some formatting. Complex formatting should be applied separately.

### Platform Differences

| Feature | Word Desktop | Word Online |
|---------|-------------|-------------|
| `body.footnotes` | Full support | Limited |
| `insertFootnote()` | Full support | May vary |
| `getFootnoteBody()` | WordApiDesktop 1.4 | Not available |
| Note navigation | Full support | Full support |

## Best Practices

### 1. Always Load Before Access

```typescript
// Good - load before accessing
footnotes.load("items/body");
await context.sync();
const text = footnotes.items[0].body.text;

// Bad - will throw error
const text = context.document.body.footnotes.items[0].body.text;
```

### 2. Check Collection Length

```typescript
if (footnotes.items.length > index) {
  // Safe to access
  const footnote = footnotes.items[index];
}
```

### 3. Use NullObject Pattern for Navigation

```typescript
const next = footnote.getNextOrNullObject();
next.load("body");
await context.sync();

if (!next.isNullObject) {
  // Process next footnote
}
```

### 4. Refresh Tree After Modifications

After inserting or deleting footnotes, rebuild the accessibility tree to get updated refs:

```typescript
// After footnote operations
const updatedTree = await DocTree.buildTree(context);
```

## Related Resources

- [Word.NoteItem API Reference](https://learn.microsoft.com/en-us/javascript/api/word/word.noteitem?view=word-js-preview)
- [Word.Range API Reference](https://learn.microsoft.com/en-us/javascript/api/word/word.range?view=word-js-preview)
- [WordApi 1.5 Requirement Set](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-1-5-requirement-set?view=common-js-preview)
- [Office.js Footnote Samples](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml)

## See Also

- `skills/docx/footnotes.md` - Python library footnote operations
- `skills/docx/office-bridge.md` - General Office Bridge documentation
