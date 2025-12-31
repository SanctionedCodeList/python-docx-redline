# Office Bridge: Word Add-in DocTree API

The Office Bridge provides a TypeScript/Office.js API for interacting with Word documents via the DocTree accessibility layer. It runs as a Word add-in that connects to a local bridge server, enabling remote execution of document operations.

## Architecture

```
Python Client  <-->  Bridge Server (localhost:3847)  <-->  Word Add-in
                          WebSocket                      Office.js API
```

The add-in exposes a `DocTree` global object with all available functions.

## Quick Start

```typescript
// All operations run inside Word.run()
await Word.run(async (context) => {
  // Build the accessibility tree
  const tree = await DocTree.buildTree(context);

  // Get YAML representation
  const yaml = DocTree.toStandardYaml(tree);
  console.log(yaml);

  // Edit by ref
  await DocTree.replaceByRef(context, "p:5", "New text", { track: true });
});
```

## Tree Building

### buildTree(context, options?)

Build an accessibility tree from the current document.

```typescript
const tree = await DocTree.buildTree(context, {
  includeTrackedChanges: true,  // Detect tracked changes
  includeComments: true,        // Include comments
  changeViewMode: 'markup'      // 'markup' | 'final' | 'original'
});
```

Returns an `AccessibilityTree` with:
- `document`: Metadata (filename, author, dates)
- `content`: Array of `AccessibilityNode` objects
- `stats`: Document statistics

## YAML Serialization

Three verbosity levels for different use cases:

```typescript
// Minimal - just refs and truncated text
DocTree.toMinimalYaml(tree);  // ~500 tokens/page

// Standard - balanced detail (default)
DocTree.toStandardYaml(tree); // ~1,500 tokens/page

// Full - includes all formatting
DocTree.toFullYaml(tree);     // ~3,000 tokens/page
```

## Ref-Based Editing

### replaceByRef(context, ref, newText, options?)

Replace entire paragraph content.

```typescript
await DocTree.replaceByRef(context, "p:5", "Updated paragraph", { track: true });
```

### insertAfterRef / insertBeforeRef

Insert text at paragraph boundaries.

```typescript
await DocTree.insertAfterRef(context, "p:5", " (amended)", { track: true });
await DocTree.insertBeforeRef(context, "p:5", "Note: ", { track: true });
```

### deleteByRef(context, ref, options?)

Delete an element.

```typescript
await DocTree.deleteByRef(context, "p:5", { track: true });
```

### formatByRef(context, ref, formatting)

Apply formatting to an element.

```typescript
await DocTree.formatByRef(context, "p:5", {
  bold: true,
  color: "#0000FF",
  size: 14
});
```

### getTextByRef(context, ref)

Get text content of an element.

```typescript
const text = await DocTree.getTextByRef(context, "p:5");
```

## Batch Operations

For multiple edits, use batch functions to minimize round-trips:

### batchEdit(context, operations, options?)

Execute multiple operations in one call.

```typescript
const result = await DocTree.batchEdit(context, [
  { ref: "p:3", operation: "replace", newText: "Updated intro" },
  { ref: "p:7", operation: "replace", newText: "New conclusion" },
  { ref: "p:12", operation: "delete" },
  { ref: "p:5", operation: "insertAfter", insertText: " (amended)" },
], { track: true });

console.log(`${result.successCount}/${result.results.length} succeeded`);
```

### batchReplace / batchDelete

Convenience wrappers:

```typescript
await DocTree.batchReplace(context, [
  { ref: "p:3", text: "New text 1" },
  { ref: "p:7", text: "New text 2" },
], { track: true });

await DocTree.batchDelete(context, ["p:10", "p:15", "p:20"], { track: true });
```

## Scope-Aware Editing

Edit multiple elements matching a scope:

### replaceByScope(context, tree, scope, newText, options?)

```typescript
// Replace all paragraphs in "Methods" section
await DocTree.replaceByScope(context, tree, "section:Methods", "Updated content");
```

### deleteByScope / formatByScope / searchReplaceByScope

```typescript
// Delete paragraphs containing "DRAFT"
await DocTree.deleteByScope(context, tree, "DRAFT", { track: true });

// Format all headings in a section
await DocTree.formatByScope(context, tree,
  { section: "Results", role: "heading" },
  { bold: true, color: "#0000FF" }
);

// Search and replace within scope
await DocTree.searchReplaceByScope(context, tree,
  "section:Parties",
  "Plaintiff",
  "Defendant",
  { track: true }
);
```

## Text Search

### findText(context, searchText, tree?, options?)

Search for text across the document.

```typescript
const result = await DocTree.findText(context, "agreement");
console.log(`Found ${result.count} matches`);

for (const match of result.matches) {
  console.log(`${match.ref}: "${match.matchedText}" at position ${match.start}`);
}
```

Options:
- `caseInsensitive`: Case-insensitive search
- `regex`: Treat searchText as regex pattern
- `wholeWord`: Match whole words only
- `maxMatches`: Limit results
- `scope`: Only search within scope

```typescript
// Case-insensitive regex search in a section
const result = await DocTree.findText(context, "\\$[\\d,]+\\.\\d{2}", tree, {
  regex: true,
  scope: "section:Pricing"
});
```

### findAndHighlight(context, searchText, color, options?)

Search and apply highlighting.

```typescript
const count = await DocTree.findAndHighlight(context, "important", "yellow");
```

## Tracked Changes

### acceptAllChanges / rejectAllChanges

```typescript
const result = await DocTree.acceptAllChanges(context);
console.log(`Accepted ${result.count} changes`);
```

### acceptNextChange / rejectNextChange

Step through changes one at a time.

```typescript
const result = await DocTree.acceptNextChange(context);
console.log(`${result.count} changes remaining`);
```

### getTrackedChangesInfo(context)

Get information about all tracked changes.

```typescript
const info = await DocTree.getTrackedChangesInfo(context);
console.log(`${info.insertions} insertions, ${info.deletions} deletions`);
```

## Comments

### addComment(context, ref, commentText)

Add a comment to a paragraph.

```typescript
const result = await DocTree.addComment(context, "p:5", "Please review this section");
console.log(`Added comment: ${result.commentId}`);
```

### addCommentToSelection(context, commentText)

Add comment to current selection.

### replyToComment(context, commentId, replyText)

```typescript
await DocTree.replyToComment(context, "comment-123", "I've addressed this");
```

### resolveComment / unresolveComment

```typescript
await DocTree.resolveComment(context, "comment-123");
```

### deleteComment(context, commentId)

```typescript
await DocTree.deleteComment(context, "comment-123");
```

### getComments(context)

List all comments.

```typescript
const result = await DocTree.getComments(context);
for (const c of result.comments) {
  console.log(`${c.author}: ${c.content} (${c.replyCount} replies)`);
}
```

## Navigation Helpers

### getNextRef / getPrevRef

```typescript
const next = DocTree.getNextRef("p:5");      // "p:6"
const prev = DocTree.getPrevRef("p:5");      // "p:4"
const first = DocTree.getPrevRef("p:0");     // null
```

### getSiblingRefs(ref, totalParagraphs?)

```typescript
const { prev, next } = DocTree.getSiblingRefs("p:5");
```

### getSectionForRef(tree, ref)

Find the section heading for a paragraph.

```typescript
const section = DocTree.getSectionForRef(tree, "p:25");
if (section) {
  console.log(`In section "${section.headingText}" (level ${section.level})`);
}
```

### getRefRange(startRef, endRef)

Get all refs between two refs (inclusive).

```typescript
const refs = DocTree.getRefRange("p:5", "p:10");
// ["p:5", "p:6", "p:7", "p:8", "p:9", "p:10"]
```

### isRefInRange(ref, startRef, endRef)

```typescript
DocTree.isRefInRange("p:7", "p:5", "p:10");  // true
```

## Document Summary

### getDocumentSummary(context, tree?)

Get comprehensive document statistics.

```typescript
const summary = await DocTree.getDocumentSummary(context, tree);
console.log(`
  Paragraphs: ${summary.paragraphCount}
  Tables: ${summary.tableCount}
  Words: ${summary.wordCount}
  Characters: ${summary.characterCount}
  Headings: ${summary.headingCount}
  Comments: ${summary.commentCount}
  Sections: ${summary.sections.join(", ")}
`);
```

### getWordCount(context, refs)

Word count for specific refs.

```typescript
const count = await DocTree.getWordCount(context, ["p:5", "p:6", "p:7"]);
```

## Scope System

Scopes filter document elements by various criteria:

### String Shortcuts

```typescript
"keyword"              // Paragraphs containing text
"section:Introduction" // Paragraphs in section
"role:heading"         // By semantic role
"style:Normal"         // By style name
"footnotes"            // All footnotes
"footnote:1"           // Specific footnote
```

### Object Format (AND logic)

```typescript
{
  contains: "payment",
  notContains: "Exhibit",
  section: "Definitions",
  hasChanges: true
}
```

### Scope Functions

```typescript
// Parse scope string
const filter = DocTree.parseScope("section:Methods");

// Resolve scope to matching nodes
const result = DocTree.resolveScope(tree, "section:Methods");
console.log(`Found ${result.nodes.length} nodes`);

// Filter nodes by scope
const filtered = DocTree.filterByScope(tree.content, scope);

// Find first matching node
const first = DocTree.findFirstByScope(tree.content, scope);

// Count matches
const count = DocTree.countByScope(tree.content, scope);

// Get refs for all matches
const refs = DocTree.getRefsByScope(tree.content, scope);
```

## Edit Options

Most editing functions accept an options object:

```typescript
interface EditOptions {
  track?: boolean;    // Enable tracked changes
  author?: string;    // Author name for changes
  comment?: string;   // Comment to attach
}
```

## Error Handling

All functions return result objects with success indicators:

```typescript
const result = await DocTree.replaceByRef(context, "p:999", "text");
if (!result.success) {
  console.error(result.error);
}
```

## Performance Notes

- **Batched operations**: Use `batchEdit`, `batchReplace`, `batchDelete` for multiple edits
- **Scope-aware functions**: Already optimized with batched `context.sync()` calls
- **Tree building**: ~2 seconds for 500-paragraph documents with tracked changes
- **Single sync**: Most operations use minimal `context.sync()` calls

## See Also

- [accessibility.md](./accessibility.md) - Python accessibility API
- [DOCTREE_SPEC.md](../../docs/DOCTREE_SPEC.md) - Full specification
