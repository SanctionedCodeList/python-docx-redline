---
name: Office Bridge Hyperlinks
description: Best practices for working with hyperlinks via the Office.js Word API in the Office Bridge add-in. Covers inserting, reading, editing, and removing hyperlinks using Range.hyperlink and related APIs.
---

# Office Bridge Hyperlinks

This document covers hyperlink operations using the Office.js Word API through the Office Bridge add-in.

## Office.js Hyperlink APIs

### API Availability

| API | Requirement Set | Notes |
|-----|-----------------|-------|
| `Range.hyperlink` (read/write) | WordApi 1.3 | Get first or set hyperlink on range |
| `Range.getHyperlinkRanges()` | WordApi 1.3 | Get all hyperlink ranges in selection |
| `Range.hyperlinks` (collection) | WordApiDesktop 1.3 | Desktop-only collection |
| `Word.Hyperlink` class | Preview | Not production-ready |
| `Range.insertHtml()` | WordApi 1.1 | Insert hyperlinks via HTML |

### Key Limitations

1. **No dedicated insertHyperlink method** - Set via `Range.hyperlink` property or use `insertHtml()`
2. **Setting hyperlink replaces all existing** - All hyperlinks in the range are deleted when setting a new one
3. **Table cell limitation** - Cannot add hyperlink to range of entire cell (GeneralException)
4. **Desktop vs Web differences** - `HyperlinkCollection` only available on desktop

## Current Office Bridge Support

The Office Bridge currently has **type definitions** for hyperlinks but **limited implementation**:

```typescript
// From types.ts
export interface LinkInfo {
  ref: Ref;           // e.g., "lnk:5"
  text: string;       // Display text
  target?: Ref;       // Internal bookmark ref
  targetLocation?: Ref;
  url?: string;       // External URL
}

export interface LinkSummary {
  internal: LinkInfo[];
  external: LinkInfo[];
  broken: Array<LinkInfo & { error: string }>;
}
```

The `AccessibilityNode` includes a `links?: LinkInfo[]` property for paragraphs containing hyperlinks.

## Inserting Hyperlinks

### Method 1: Set Range.hyperlink Property

The simplest approach for adding a hyperlink to existing text:

```typescript
await Word.run(async (context) => {
  // Select the text you want to make a hyperlink
  const range = context.document.getSelection();

  // Set the hyperlink (format: "url#location" or just "url")
  range.hyperlink = "https://example.com";

  await context.sync();
});
```

For internal links to bookmarks:

```typescript
await Word.run(async (context) => {
  const range = context.document.getSelection();

  // Use # prefix for bookmark links
  range.hyperlink = "#BookmarkName";

  await context.sync();
});
```

With location (subaddress):

```typescript
await Word.run(async (context) => {
  const range = context.document.getSelection();

  // URL with anchor/subaddress
  range.hyperlink = "https://example.com/page#section";

  await context.sync();
});
```

### Method 2: Insert via HTML

For inserting new hyperlink text (not converting existing text):

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();

  // Insert hyperlink using HTML anchor tag
  selection.insertHtml(
    '<a href="https://example.com" title="Tooltip text">Click here</a>',
    Word.InsertLocation.after
  );

  await context.sync();
});
```

Best practices for insertHtml hyperlinks:

```typescript
async function insertHyperlink(
  context: Word.RequestContext,
  url: string,
  displayText: string,
  tooltip?: string,
  insertLocation: Word.InsertLocation = Word.InsertLocation.after
): Promise<void> {
  const selection = context.document.getSelection();

  // Escape HTML entities in display text
  const escapedText = displayText
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  // Escape URL
  const escapedUrl = url.replace(/"/g, '%22');

  // Build HTML
  let html = `<a href="${escapedUrl}"`;
  if (tooltip) {
    html += ` title="${tooltip.replace(/"/g, '&quot;')}"`;
  }
  html += `>${escapedText}</a>`;

  selection.insertHtml(html, insertLocation);
  await context.sync();
}
```

## Reading Hyperlinks

### Get All Hyperlinks in Document

```typescript
await Word.run(async (context) => {
  // Get the entire document body
  const bodyRange = context.document.body.getRange(Word.RangeLocation.whole);

  // Get all hyperlink ranges
  const hyperlinkRanges = bodyRange.getHyperlinkRanges();
  hyperlinkRanges.load("hyperlink");

  await context.sync();

  // Process each hyperlink
  hyperlinkRanges.items.forEach((range, index) => {
    console.log(`Link ${index}: ${range.hyperlink}`);
  });
});
```

### Get Hyperlink Text and URL

```typescript
await Word.run(async (context) => {
  const bodyRange = context.document.body.getRange(Word.RangeLocation.whole);
  const hyperlinkRanges = bodyRange.getHyperlinkRanges();

  // Load both text and hyperlink properties
  hyperlinkRanges.load(["hyperlink", "text"]);

  await context.sync();

  const links: Array<{text: string; url: string}> = [];

  hyperlinkRanges.items.forEach((range) => {
    links.push({
      text: range.text,
      url: range.hyperlink
    });
  });

  return links;
});
```

### Hyperlink in Selection

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("hyperlink");

  await context.sync();

  if (selection.hyperlink) {
    console.log(`Selected hyperlink: ${selection.hyperlink}`);
  } else {
    console.log("Selection is not a hyperlink");
  }
});
```

## Editing Hyperlinks

### Change URL

```typescript
await Word.run(async (context) => {
  const bodyRange = context.document.body.getRange(Word.RangeLocation.whole);
  const hyperlinkRanges = bodyRange.getHyperlinkRanges();
  hyperlinkRanges.load("hyperlink");

  await context.sync();

  // Find and update specific hyperlink
  hyperlinkRanges.items.forEach((range) => {
    if (range.hyperlink.includes("old-domain.com")) {
      range.hyperlink = range.hyperlink.replace(
        "old-domain.com",
        "new-domain.com"
      );
    }
  });

  await context.sync();
});
```

### Change Display Text

```typescript
await Word.run(async (context) => {
  const bodyRange = context.document.body.getRange(Word.RangeLocation.whole);
  const hyperlinkRanges = bodyRange.getHyperlinkRanges();
  hyperlinkRanges.load(["hyperlink", "text"]);

  await context.sync();

  // Find hyperlink by URL and change its text
  const targetUrl = "https://example.com";
  hyperlinkRanges.items.forEach((range) => {
    if (range.hyperlink === targetUrl) {
      range.insertText("Updated Link Text", Word.InsertLocation.replace);
    }
  });

  await context.sync();
});
```

## Removing Hyperlinks

### Remove Hyperlink, Keep Text

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();

  // Setting hyperlink to empty string removes the link but keeps text
  selection.hyperlink = "";

  await context.sync();
});
```

### Remove All Hyperlinks

```typescript
await Word.run(async (context) => {
  const bodyRange = context.document.body.getRange(Word.RangeLocation.whole);
  const hyperlinkRanges = bodyRange.getHyperlinkRanges();
  hyperlinkRanges.load("items");

  await context.sync();

  // Remove each hyperlink (keeps text)
  hyperlinkRanges.items.forEach((range) => {
    range.hyperlink = "";
  });

  await context.sync();
});
```

### Remove Hyperlink and Text

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();

  // Delete the entire range
  selection.delete();

  await context.sync();
});
```

## Integration with Ref System

When building the accessibility tree, hyperlinks should be assigned refs in the `lnk:N` format:

```typescript
function extractHyperlinks(
  paragraph: Word.Paragraph,
  paragraphRef: string
): LinkInfo[] {
  const links: LinkInfo[] = [];
  let linkIndex = 0;

  // Note: This is conceptual - actual implementation requires
  // working with the paragraph's range and getHyperlinkRanges()

  paragraph.getRange(Word.RangeLocation.whole)
    .getHyperlinkRanges()
    .items.forEach((range) => {
      const url = range.hyperlink;
      const isInternal = url.startsWith("#");

      links.push({
        ref: `lnk:${linkIndex++}`,
        text: range.text,
        url: isInternal ? undefined : url,
        target: isInternal ? `bk:${url.slice(1)}` : undefined
      });
    });

  return links;
}
```

## Known Issues and Gotchas

### 1. Table Cell Hyperlinks

Adding a hyperlink to a range that spans an entire table cell causes a GeneralException:

```typescript
// This will fail:
const cell = table.getCell(0, 0);
const cellRange = cell.body.getRange(Word.RangeLocation.whole);
cellRange.hyperlink = "https://example.com"; // GeneralException!

// Workaround: Select content inside the cell, not the whole cell
const cellContent = cell.body.paragraphs.getFirst().getRange();
cellContent.hyperlink = "https://example.com"; // Works
```

### 2. Word Online vs Desktop

- `insertHtml` has limited style support in Word Online
- `HyperlinkCollection` (`Range.hyperlinks`) only available on desktop
- Test thoroughly on both platforms

### 3. Hyperlink Format String

The `hyperlink` property uses a combined format:
- External: `"https://example.com"` or `"https://example.com#section"`
- Internal: `"#BookmarkName"`

Parse carefully when reading:

```typescript
function parseHyperlink(hyperlinkValue: string): {
  url?: string;
  bookmark?: string;
  subAddress?: string;
} {
  if (hyperlinkValue.startsWith("#")) {
    return { bookmark: hyperlinkValue.slice(1) };
  }

  const [url, subAddress] = hyperlinkValue.split("#");
  return { url, subAddress };
}
```

### 4. Setting Hyperlink Clears Others

When you set `range.hyperlink`, ALL hyperlinks within that range are removed:

```typescript
// If paragraph has 3 hyperlinks and you do:
paragraph.getRange(Word.RangeLocation.whole).hyperlink = "https://new.com";
// All 3 original hyperlinks are gone, replaced by one

// To add without affecting others, use insertHtml at a specific point
```

## Comparison: Office Bridge vs Python Library

| Feature | Office Bridge (Office.js) | Python (python-docx-redline) |
|---------|---------------------------|------------------------------|
| Insert external link | `range.hyperlink = url` | `doc.insert_hyperlink(url=...)` |
| Insert internal link | `range.hyperlink = "#bookmark"` | `doc.insert_hyperlink(anchor=...)` |
| Insert with HTML | `insertHtml('<a>...</a>')` | N/A |
| Read all links | `getHyperlinkRanges()` | `doc.hyperlinks` |
| Edit URL | `range.hyperlink = newUrl` | `doc.edit_hyperlink_url(ref, url)` |
| Edit text | `range.insertText(...)` | `doc.edit_hyperlink_text(ref, text)` |
| Remove (keep text) | `range.hyperlink = ""` | `doc.remove_hyperlink(ref)` |
| Headers/Footers | Access via `document.sections` | Dedicated methods |
| Tracked changes | Not directly supported | `track=True` parameter |

## Sources

- [Word.Hyperlink class - Office Add-ins | Microsoft Learn](https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink?view=word-js-preview)
- [Word.Range class - Office Add-ins | Microsoft Learn](https://learn.microsoft.com/en-us/javascript/api/word/word.range?view=word-js-preview)
- [Malformatted hyperlink in Office Online - GitHub Issue #1753](https://github.com/OfficeDev/office-js/issues/1753)
- [Word Tutorial - office-js-docs-pr](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/tutorials/word-tutorial.md)
