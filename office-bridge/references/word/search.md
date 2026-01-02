---
name: office-bridge-search
description: "Text search API for finding and highlighting content in Word documents via Office Bridge. Covers findText, findAndHighlight, regex patterns, scoped search, and match result handling."
---

# Office Bridge: Text Search API

The Office Bridge provides powerful text search capabilities through the DocTree API. This document covers finding text, handling match results, regex patterns, scoped searches, and highlighting functionality.

## Overview

Text search in Office Bridge operates at the paragraph level, returning refs and position information that integrate with the editing API. The search system supports:

- Literal text matching
- Regular expression patterns
- Case-insensitive search
- Whole word matching
- Scoped search within sections or filtered content
- Batch highlighting

## Core API

### findText(context, searchText, tree?, options?)

Search for text across the document, returning refs and match positions.

```typescript
const result = await DocTree.findText(context, searchText, tree?, options?);
```

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `context` | `Word.RequestContext` | Yes | Office.js request context |
| `searchText` | `string` | Yes | Text or regex pattern to search for |
| `tree` | `AccessibilityTree \| null` | No | Required if using scope option |
| `options` | `FindTextOptions` | No | Search configuration |

**Returns:** `FindTextResult`

```typescript
interface FindTextResult {
  /** Total number of matches found */
  count: number;
  /** All matches found */
  matches: TextMatch[];
  /** Number of paragraphs searched */
  paragraphsSearched: number;
}
```

### FindTextOptions

Configuration options for search operations:

```typescript
interface FindTextOptions {
  /** Case insensitive search (default: false) */
  caseInsensitive?: boolean;
  /** Use regex pattern (default: false, treats search as literal text) */
  regex?: boolean;
  /** Match whole words only (default: false) */
  wholeWord?: boolean;
  /** Maximum number of matches to return (default: unlimited) */
  maxMatches?: number;
  /** Only search within a specific scope */
  scope?: ScopeSpec;
}
```

#### Option Details

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `caseInsensitive` | `boolean` | `false` | When `true`, matches regardless of case ("Agreement" matches "agreement") |
| `regex` | `boolean` | `false` | When `true`, treats `searchText` as a regex pattern |
| `wholeWord` | `boolean` | `false` | When `true`, only matches complete words (adds `\b` boundaries) |
| `maxMatches` | `number` | `Infinity` | Limits the number of matches returned |
| `scope` | `ScopeSpec` | `undefined` | Restricts search to matching paragraphs |

### TextMatch Result Structure

Each match returned contains detailed position information:

```typescript
interface TextMatch {
  /** Reference to the paragraph containing the match */
  ref: Ref;
  /** Full text of the paragraph */
  text: string;
  /** Starting position of the match within the paragraph */
  start: number;
  /** Ending position of the match within the paragraph */
  end: number;
  /** The actual matched text (may differ in case if caseInsensitive) */
  matchedText: string;
}
```

**Example match object:**

```typescript
{
  ref: "p:15",
  text: "The Agreement shall be governed by New York law.",
  start: 4,
  end: 13,
  matchedText: "Agreement"
}
```

## Basic Usage Examples

### Simple Text Search

```typescript
await Word.run(async (context) => {
  const result = await DocTree.findText(context, "agreement");

  console.log(`Found ${result.count} matches in ${result.paragraphsSearched} paragraphs`);

  for (const match of result.matches) {
    console.log(`${match.ref}: "${match.matchedText}" at position ${match.start}`);
  }
});
```

### Case-Insensitive Search

```typescript
await Word.run(async (context) => {
  const result = await DocTree.findText(context, "WHEREAS", null, {
    caseInsensitive: true
  });

  // Matches "WHEREAS", "Whereas", "whereas", etc.
  console.log(`Found ${result.count} occurrences`);
});
```

### Whole Word Matching

```typescript
await Word.run(async (context) => {
  // Only match "the" as a complete word, not within "other", "there", etc.
  const result = await DocTree.findText(context, "the", null, {
    wholeWord: true,
    caseInsensitive: true
  });

  console.log(`Found ${result.count} instances of "the" as a word`);
});
```

### Limited Results

```typescript
await Word.run(async (context) => {
  // Get first 10 matches only
  const result = await DocTree.findText(context, "shall", null, {
    maxMatches: 10
  });

  console.log(`Showing first ${result.count} matches`);
});
```

## Regex Search Examples

When `regex: true` is set, the search text is interpreted as a JavaScript regular expression pattern.

### Finding Currency Values

```typescript
await Word.run(async (context) => {
  // Match dollar amounts like $1,234.56
  const result = await DocTree.findText(context, "\\$[\\d,]+\\.\\d{2}", null, {
    regex: true
  });

  for (const match of result.matches) {
    console.log(`Found amount: ${match.matchedText} in ${match.ref}`);
  }
});
```

### Finding Dates

```typescript
await Word.run(async (context) => {
  // Match dates in MM/DD/YYYY format
  const result = await DocTree.findText(context, "\\d{1,2}/\\d{1,2}/\\d{4}", null, {
    regex: true
  });

  console.log(`Found ${result.count} dates`);
});
```

### Finding Email Addresses

```typescript
await Word.run(async (context) => {
  const result = await DocTree.findText(
    context,
    "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}",
    null,
    { regex: true }
  );

  for (const match of result.matches) {
    console.log(`Email found: ${match.matchedText}`);
  }
});
```

### Finding Section References

```typescript
await Word.run(async (context) => {
  // Match "Section 1.2.3" style references
  const result = await DocTree.findText(
    context,
    "Section\\s+\\d+(\\.\\d+)*",
    null,
    { regex: true, caseInsensitive: true }
  );

  console.log(`Found ${result.count} section references`);
});
```

### Capturing Groups (Partial Match Extraction)

```typescript
await Word.run(async (context) => {
  // Find party names in "Party: [Name]" format
  const result = await DocTree.findText(
    context,
    "Party:\\s*(.+?)(?=,|\\.|$)",
    null,
    { regex: true }
  );

  // matchedText contains the full match including "Party:"
  for (const match of result.matches) {
    console.log(`Full match: ${match.matchedText}`);
  }
});
```

### Complex Legal Patterns

```typescript
await Word.run(async (context) => {
  // Find defined terms (capitalized phrases in quotes)
  const result = await DocTree.findText(
    context,
    '"[A-Z][a-zA-Z\\s]+"',
    null,
    { regex: true }
  );

  console.log(`Found ${result.count} potential defined terms`);
});
```

## Scoped Search Examples

Scoped search restricts matches to paragraphs that satisfy a scope condition. This requires passing the accessibility tree.

### Search Within a Section

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Only search in the "Definitions" section
  const result = await DocTree.findText(context, "means", tree, {
    scope: "section:Definitions"
  });

  console.log(`Found ${result.count} "means" in Definitions section`);
});
```

### Search in Paragraphs Containing Text

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Find "indemnify" only in paragraphs that mention "Seller"
  const result = await DocTree.findText(context, "indemnify", tree, {
    scope: { contains: "Seller" }
  });

  console.log(`Found ${result.count} matches in Seller-related paragraphs`);
});
```

### Search by Style

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Search only in heading paragraphs
  const result = await DocTree.findText(context, "Article", tree, {
    scope: "role:heading"
  });

  console.log(`Found ${result.count} article headings`);
});
```

### Combined Scope with Regex

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Find dollar amounts only in the Pricing section
  const result = await DocTree.findText(
    context,
    "\\$[\\d,]+\\.\\d{2}",
    tree,
    {
      regex: true,
      scope: "section:Pricing"
    }
  );

  console.log(`Found ${result.count} amounts in Pricing section`);
});
```

### Exclude Certain Content

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Search everywhere except Exhibit sections
  const result = await DocTree.findText(context, "Confidential", tree, {
    caseInsensitive: true,
    scope: { notContains: "Exhibit" }
  });

  console.log(`Found ${result.count} matches outside exhibits`);
});
```

### Search in Changed Paragraphs

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context, {
    includeTrackedChanges: true
  });

  // Find text only in paragraphs with tracked changes
  const result = await DocTree.findText(context, "payment", tree, {
    scope: { hasChanges: true }
  });

  console.log(`Found ${result.count} matches in changed paragraphs`);
});
```

## Find and Highlight

### findAndHighlight(context, searchText, color, options?)

Search for text and apply highlighting to all matches.

```typescript
const count = await DocTree.findAndHighlight(context, searchText, color, options?);
```

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `context` | `Word.RequestContext` | Yes | Office.js request context |
| `searchText` | `string` | Yes | Text or regex pattern to search for |
| `color` | `string` | Yes | Highlight color name or hex code |
| `options` | `FindTextOptions` | No | Search configuration (same as findText) |

**Returns:** `number` - Count of highlighted occurrences

**Supported highlight colors:**

- Named colors: `"yellow"`, `"green"`, `"cyan"`, `"pink"`, `"blue"`, `"red"`, etc.
- Hex codes: `"#FFFF00"`, `"#00FF00"`, etc.

### Basic Highlighting

```typescript
await Word.run(async (context) => {
  const count = await DocTree.findAndHighlight(context, "important", "yellow");
  console.log(`Highlighted ${count} occurrences`);
});
```

### Case-Insensitive Highlighting

```typescript
await Word.run(async (context) => {
  const count = await DocTree.findAndHighlight(
    context,
    "confidential",
    "pink",
    { caseInsensitive: true }
  );
  console.log(`Highlighted ${count} confidential mentions`);
});
```

### Highlight Patterns

```typescript
await Word.run(async (context) => {
  // Highlight all dates
  const count = await DocTree.findAndHighlight(
    context,
    "\\d{1,2}/\\d{1,2}/\\d{4}",
    "cyan",
    { regex: true }
  );
  console.log(`Highlighted ${count} dates`);
});
```

### Multiple Highlight Colors

```typescript
await Word.run(async (context) => {
  // Highlight different terms with different colors
  await DocTree.findAndHighlight(context, "Buyer", "green");
  await DocTree.findAndHighlight(context, "Seller", "blue");
  await DocTree.findAndHighlight(context, "shall", "yellow", { wholeWord: true });
});
```

## Integration with Editing

Search results integrate seamlessly with the editing API via refs.

### Find and Replace

```typescript
await Word.run(async (context) => {
  // Find all occurrences
  const result = await DocTree.findText(context, "Acme Corp", null, {
    caseInsensitive: true
  });

  // Replace each match using replaceByRef
  for (const match of result.matches) {
    const currentText = await DocTree.getTextByRef(context, match.ref);
    if (currentText) {
      const newText = currentText.replace(/acme corp/gi, "NewCo LLC");
      await DocTree.replaceByRef(context, match.ref, newText, { track: true });
    }
  }
});
```

### Add Comments to Matches

```typescript
await Word.run(async (context) => {
  const result = await DocTree.findText(context, "TBD", null, {
    wholeWord: true
  });

  // Add a comment to each TBD
  for (const match of result.matches) {
    await DocTree.addComment(context, match.ref, "Please provide specific information");
  }

  console.log(`Added comments to ${result.count} TBD placeholders`);
});
```

### Conditional Formatting Based on Content

```typescript
await Word.run(async (context) => {
  // Find warning language
  const result = await DocTree.findText(
    context,
    "WARNING|CAUTION|NOTICE",
    null,
    { regex: true, caseInsensitive: true }
  );

  // Format paragraphs containing warnings
  const formattedRefs = new Set<string>();
  for (const match of result.matches) {
    if (!formattedRefs.has(match.ref)) {
      await DocTree.formatByRef(context, match.ref, {
        bold: true,
        color: "#CC0000"
      });
      formattedRefs.add(match.ref);
    }
  }
});
```

## Performance Notes

### Batched Loading

The `findText` function uses batched paragraph loading for optimal performance:

```typescript
// All paragraphs are loaded in a single sync call
const paragraphs = context.document.body.paragraphs.load('items,text');
await context.sync();
// Then searched synchronously in memory
```

This means:
- Large documents (~500+ paragraphs) search in ~1-2 seconds
- Search time is O(n) where n is paragraph count
- Regex complexity affects per-paragraph search time

### Best Practices

1. **Use `maxMatches` for preview searches:**
   ```typescript
   // Quick preview of matches
   const preview = await DocTree.findText(context, searchText, null, {
     maxMatches: 10
   });
   ```

2. **Use scope to reduce search space:**
   ```typescript
   // Instead of searching entire document
   const result = await DocTree.findText(context, "term", tree, {
     scope: "section:Definitions"
   });
   ```

3. **Combine with tree building strategically:**
   ```typescript
   // Build tree once, reuse for multiple scoped searches
   const tree = await DocTree.buildTree(context);

   const search1 = await DocTree.findText(context, "Buyer", tree, { scope: "section:Parties" });
   const search2 = await DocTree.findText(context, "Seller", tree, { scope: "section:Parties" });
   ```

4. **Use `wholeWord` for common terms:**
   ```typescript
   // Without wholeWord, "a" matches thousands of times
   const result = await DocTree.findText(context, "a", null, {
     wholeWord: true
   });
   ```

5. **Prefer literal search over regex when possible:**
   ```typescript
   // Faster - literal search
   const result1 = await DocTree.findText(context, "Section 3.1");

   // Slower - regex pattern
   const result2 = await DocTree.findText(context, "Section\\s+3\\.1", null, { regex: true });
   ```

### Memory Considerations

- Each `TextMatch` stores the full paragraph text
- For documents with many matches, consider processing in batches:

```typescript
await Word.run(async (context) => {
  let offset = 0;
  const batchSize = 100;

  while (true) {
    const result = await DocTree.findText(context, searchText, null, {
      maxMatches: batchSize
    });

    if (result.count === 0) break;

    // Process this batch
    for (const match of result.matches) {
      // ... process match
    }

    // For large result sets, you may need pagination logic
    // This simple example assumes sequential processing
    break; // Remove this for actual pagination implementation
  }
});
```

## Error Handling

### Invalid Regex Pattern

```typescript
await Word.run(async (context) => {
  try {
    const result = await DocTree.findText(context, "[invalid(regex", null, {
      regex: true
    });
  } catch (e) {
    // Invalid regex returns empty result, doesn't throw
    // result.count will be 0
  }
});
```

### Scope Without Tree

```typescript
await Word.run(async (context) => {
  // This works but scope is ignored without tree
  const result = await DocTree.findText(context, "text", null, {
    scope: "section:Methods"  // Ignored because tree is null
  });

  // Correct approach:
  const tree = await DocTree.buildTree(context);
  const scopedResult = await DocTree.findText(context, "text", tree, {
    scope: "section:Methods"  // Now properly scoped
  });
});
```

## See Also

- [office-bridge.md](../office-bridge.md) - Main Office Bridge documentation
- [scope.md](./ideas/scope.md) - Scope system documentation
- [selection.md](./selection.md) - Selection and cursor operations
