---
name: headers-footers
description: "Best practices for reading and modifying Word document headers and footers using Office.js via the Office Bridge add-in."
---

# Headers and Footers in Office.js

This document covers the Office.js APIs for working with Word document headers and footers through the Office Bridge add-in.

## API Overview

Headers and footers are accessed through the `Word.Section` class. Each section in a Word document can have up to three types of headers and three types of footers.

### Key Methods

```typescript
// Get header for a section
section.getHeader(type: Word.HeaderFooterType): Word.Body

// Get footer for a section
section.getFooter(type: Word.HeaderFooterType): Word.Body
```

Both methods return a `Word.Body` object, allowing full manipulation of the header/footer content using the same APIs available for the document body.

**Requirement Set:** WordApi 1.1+

## Header/Footer Types

The `Word.HeaderFooterType` enum defines three types:

| Type | String Value | Description |
|------|--------------|-------------|
| `primary` | `"Primary"` | Default header/footer for all pages, excluding first page and even pages if those are set differently |
| `firstPage` | `"FirstPage"` | Header/footer for the first page of the section only |
| `evenPages` | `"EvenPages"` | Header/footer for even-numbered pages only |

### Type Hierarchy

The types follow a cascading pattern:

1. **Primary** is the default for all pages
2. **FirstPage** overrides Primary on page 1 of the section
3. **EvenPages** overrides Primary on pages 2, 4, 6, etc.

When `firstPage` or `evenPages` are set, `primary` becomes the "odd pages" header/footer (pages 3, 5, 7, etc., and page 1 if `firstPage` is not set).

## Reading Headers and Footers

### Get Header/Footer Text

```typescript
await Word.run(async (context) => {
  const section = context.document.sections.getFirst();
  const header = section.getHeader(Word.HeaderFooterType.primary);

  // Load paragraphs to get text
  header.paragraphs.load("text");
  await context.sync();

  // Iterate through header paragraphs
  for (const paragraph of header.paragraphs.items) {
    console.log(paragraph.text);
  }
});
```

### Get All Headers for a Section

```typescript
await Word.run(async (context) => {
  const section = context.document.sections.getFirst();

  const primaryHeader = section.getHeader("Primary");
  const firstPageHeader = section.getHeader("FirstPage");
  const evenPagesHeader = section.getHeader("EvenPages");

  // Load all at once
  primaryHeader.paragraphs.load("text");
  firstPageHeader.paragraphs.load("text");
  evenPagesHeader.paragraphs.load("text");

  await context.sync();

  // Process each header type
  console.log("Primary header:", primaryHeader.paragraphs.items.map(p => p.text).join(" "));
  console.log("First page header:", firstPageHeader.paragraphs.items.map(p => p.text).join(" "));
  console.log("Even pages header:", evenPagesHeader.paragraphs.items.map(p => p.text).join(" "));
});
```

## Modifying Headers and Footers

### Insert Text

```typescript
await Word.run(async (context) => {
  const section = context.document.sections.getFirst();
  const footer = section.getFooter(Word.HeaderFooterType.primary);

  // Insert at end of footer
  footer.insertText("Confidential", Word.InsertLocation.end);

  await context.sync();
});
```

### Insert Paragraph

```typescript
await Word.run(async (context) => {
  const section = context.document.sections.getFirst();
  const header = section.getHeader(Word.HeaderFooterType.primary);

  // Insert paragraph at end
  header.insertParagraph("Document Title", "End");

  await context.sync();
});
```

### Clear Header/Footer

```typescript
await Word.run(async (context) => {
  const section = context.document.sections.getFirst();

  // Clear all header types
  section.getHeader("Primary").clear();
  section.getHeader("FirstPage").clear();
  section.getHeader("EvenPages").clear();

  // Clear all footer types
  section.getFooter("Primary").clear();
  section.getFooter("FirstPage").clear();
  section.getFooter("EvenPages").clear();

  await context.sync();
});
```

### Replace Header Content

```typescript
await Word.run(async (context) => {
  const section = context.document.sections.getFirst();
  const header = section.getHeader(Word.HeaderFooterType.primary);

  // Clear existing content
  header.clear();

  // Add new content
  const newParagraph = header.insertParagraph("New Header Text", "Start");
  newParagraph.font.bold = true;
  newParagraph.alignment = Word.Alignment.centered;

  await context.sync();
});
```

## Working with Multiple Sections

```typescript
await Word.run(async (context) => {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  // Update header in all sections
  for (let i = 0; i < sections.items.length; i++) {
    const section = sections.items[i];
    const header = section.getHeader("Primary");
    header.insertParagraph(`Section ${i + 1}`, "End");
  }

  await context.sync();
});
```

## Adding Content Controls

Headers and footers support content controls for structured content:

```typescript
await Word.run(async (context) => {
  const section = context.document.sections.getFirst();
  const footer = section.getFooter(Word.HeaderFooterType.primary);

  // Insert text and wrap in content control
  footer.insertText("Page ", Word.InsertLocation.end);
  footer.insertContentControl();

  await context.sync();
});
```

## Best Practices

### 1. Batch Load Operations

Load all header/footer types in a single sync to minimize round-trips:

```typescript
// Good: Single sync
const primary = section.getHeader("Primary");
const firstPage = section.getHeader("FirstPage");
const evenPages = section.getHeader("EvenPages");
primary.paragraphs.load("text");
firstPage.paragraphs.load("text");
evenPages.paragraphs.load("text");
await context.sync();  // One sync for all

// Bad: Multiple syncs
primary.paragraphs.load("text");
await context.sync();  // Unnecessary sync
firstPage.paragraphs.load("text");
await context.sync();  // Another sync
```

### 2. Check Section Count

Documents may have multiple sections with different headers:

```typescript
const sections = context.document.sections;
sections.load("items");
await context.sync();

if (sections.items.length > 1) {
  console.log(`Document has ${sections.items.length} sections with potentially different headers`);
}
```

### 3. Preserve Formatting

When replacing header content, capture and reapply formatting if needed:

```typescript
const header = section.getHeader("Primary");
header.paragraphs.load("text,font,alignment");
await context.sync();

// Store original formatting
const originalFont = header.paragraphs.items[0]?.font;
const originalAlignment = header.paragraphs.items[0]?.alignment;

// Clear and recreate with same formatting
header.clear();
const newPara = header.insertParagraph("New Text", "Start");
if (originalFont) {
  newPara.font.bold = originalFont.bold;
  newPara.font.size = originalFont.size;
}
if (originalAlignment) {
  newPara.alignment = originalAlignment;
}
```

### 4. Handle Empty Headers/Footers

An empty header/footer still returns a valid Body object. Check paragraph count:

```typescript
header.paragraphs.load("items");
await context.sync();

if (header.paragraphs.items.length === 0 ||
    (header.paragraphs.items.length === 1 && !header.paragraphs.items[0].text.trim())) {
  console.log("Header is empty");
}
```

## Limitations and Gotchas

### 1. No Direct Page Number API

Office.js does not have a direct API for inserting page numbers. Page numbers in headers/footers are typically inserted as fields:

```typescript
// Cannot do: header.insertPageNumber()
// Instead, use insertOoxml with field codes (advanced)
```

### 2. Header/Footer Body vs Document Body

The `Word.Body` returned by `getHeader()` or `getFooter()` is separate from `document.body`. Changes to one do not affect the other.

### 3. Different Headers/Footers Checkbox

For `firstPage` and `evenPages` to work, the corresponding Word checkboxes must be enabled in the section layout. Office.js can read/write content, but enabling the "Different First Page" or "Different Odd & Even" settings may require additional configuration.

### 4. Shapes and Text Boxes

As of WordApi 1.8, shapes and text boxes in headers/footers are not fully supported. Complex header layouts with positioned elements may not be accessible.

### 5. Linked Headers

When sections have linked headers (same as previous section), modifying one affects all linked sections. Office.js does not expose linking status directly.

## References

- [Word.Section class - Microsoft Learn](https://learn.microsoft.com/en-us/javascript/api/word/word.section)
- [Word.HeaderFooterType enum - Microsoft Learn](https://learn.microsoft.com/en-us/javascript/api/word/word.headerfootertype)
- [Office.js Samples - Headers and Footers](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/word/25-paragraph/insert-header-and-footer.yaml)
