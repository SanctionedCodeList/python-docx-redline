---
name: tracked-changes
description: "Tracked changes API for Office Bridge Word Add-in. Use when making edits with change tracking, reviewing changes, accepting/rejecting changes, or getting change statistics in live Word documents."
---

# Tracked Changes in Office Bridge

The Office Bridge provides comprehensive tracked changes support for editing Word documents with revision tracking enabled. This enables collaborative editing workflows where changes can be reviewed, accepted, or rejected.

## Making Edits with Tracking

### The track Option

All editing operations accept a `track: true` option to enable tracked changes:

```typescript
await Word.run(async (context) => {
  // Replace with tracking
  await DocTree.replaceByRef(context, "p:5", "Updated text", { track: true });

  // Insert with tracking
  await DocTree.insertAfterRef(context, "p:5", " (amended)", { track: true });
  await DocTree.insertBeforeRef(context, "p:5", "Note: ", { track: true });

  // Delete with tracking
  await DocTree.deleteByRef(context, "p:3", { track: true });
});
```

### EditOptions Interface

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

### Batch Operations with Tracking

Use batch operations for multiple tracked edits:

```typescript
await Word.run(async (context) => {
  const result = await DocTree.batchEdit(context, [
    { ref: "p:3", operation: "replace", newText: "Updated intro" },
    { ref: "p:7", operation: "replace", newText: "New conclusion" },
    { ref: "p:12", operation: "delete" },
    { ref: "p:5", operation: "insertAfter", insertText: " (amended)" },
  ], { track: true });

  console.log(`${result.successCount}/${result.results.length} succeeded`);
});
```

## Accepting and Rejecting Changes

### Accept All Changes

Accept all tracked changes in the document at once:

```typescript
await Word.run(async (context) => {
  const result = await DocTree.acceptAllChanges(context);
  if (result.success) {
    console.log(`Accepted ${result.count} changes`);
  } else {
    console.error(result.error);
  }
});
```

### Reject All Changes

Reject all tracked changes, reverting the document to its original state:

```typescript
await Word.run(async (context) => {
  const result = await DocTree.rejectAllChanges(context);
  console.log(`Rejected ${result.count} changes`);
});
```

### Stepping Through Changes

For review workflows where each change needs individual consideration:

```typescript
await Word.run(async (context) => {
  // Accept the first (next) change
  const result = await DocTree.acceptNextChange(context);
  if (result.success) {
    console.log(`Accepted 1 change, ${result.count} remaining`);
  }
});

await Word.run(async (context) => {
  // Reject the first (next) change
  const result = await DocTree.rejectNextChange(context);
  if (result.success) {
    console.log(`Rejected 1 change, ${result.count} remaining`);
  }
});
```

## Getting Change Information

### getTrackedChangesInfo(context)

Get comprehensive information about all tracked changes in the document:

```typescript
await Word.run(async (context) => {
  const info = await DocTree.getTrackedChangesInfo(context);

  console.log(`API Available: ${info.available}`);
  console.log(`Total Changes: ${info.count}`);
  console.log(`Insertions: ${info.insertions}`);
  console.log(`Deletions: ${info.deletions}`);

  // Iterate through individual changes
  for (const change of info.changes) {
    console.log(`${change.type}: "${change.text}" by ${change.author}`);
  }
});
```

### TrackedChangeResult Structure

All tracked change operations return a `TrackedChangeResult`:

```typescript
interface TrackedChangeResult {
  /** Whether the operation succeeded */
  success: boolean;
  /** Number of changes affected (or remaining for next/step operations) */
  count: number;
  /** Error message if failed */
  error?: string;
}
```

### TrackedChange Type

Individual tracked changes have this structure:

```typescript
interface TrackedChange {
  /** Reference ID for this change (e.g., "change:0") */
  ref: Ref;
  /** Type of change: 'ins' for insertion, 'del' for deletion */
  type: 'ins' | 'del';
  /** Author who made the change */
  author: string;
  /** ISO date string when change was made */
  date: string;
  /** The changed text */
  text: string;
  /** Location information */
  location: {
    paragraphRef: Ref;
  };
}
```

## Tracked Changes in the Accessibility Tree

### Building Tree with Change Detection

Include tracked changes when building the accessibility tree:

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context, {
    includeTrackedChanges: true,
    changeViewMode: 'markup'  // 'markup' | 'final' | 'original'
  });

  // Tree nodes include change information
  for (const node of tree.content) {
    if (node.hasChanges) {
      console.log(`${node.ref}: Has tracked changes`);
      console.log(`  Changes: ${node.changeRefs?.join(', ')}`);
    }
  }
});
```

### AccessibilityNode Change Properties

Nodes with tracked changes include:

```typescript
interface AccessibilityNode {
  // ... other properties

  /** Whether this node has tracked changes */
  hasChanges?: boolean;
  /** Array of TrackedChange objects */
  changes?: TrackedChange[];
  /** Array of change ref IDs for quick reference */
  changeRefs?: Ref[];
}
```

### Markup Text Format

In `markup` mode, the text includes inline change markers:

- Insertions: `{++inserted text++}`
- Deletions: `{--deleted text--}`

```typescript
// Example node.text in markup mode:
"The {--old--} {++new++} paragraph text"
```

## Change View Modes

The `changeViewMode` option controls how text appears:

| Mode | Description | Use Case |
|------|-------------|----------|
| `'markup'` | Shows insertions and deletions with inline markers | Review changes, see what changed |
| `'final'` | Shows document as if all changes accepted | See the end result |
| `'original'` | Shows document before any changes | Compare to original |

```typescript
// See document with all changes accepted
const finalTree = await DocTree.buildTree(context, {
  includeTrackedChanges: true,
  changeViewMode: 'final'
});

// See original document before changes
const originalTree = await DocTree.buildTree(context, {
  includeTrackedChanges: true,
  changeViewMode: 'original'
});

// See both with inline markers
const markupTree = await DocTree.buildTree(context, {
  includeTrackedChanges: true,
  changeViewMode: 'markup'
});
```

## Best Practices for Review Workflows

### 1. Check for Changes First

Before starting a review, check if there are changes to review:

```typescript
await Word.run(async (context) => {
  const info = await DocTree.getTrackedChangesInfo(context);

  if (!info.available) {
    console.log("Tracked changes API not available");
    return;
  }

  if (info.count === 0) {
    console.log("No tracked changes to review");
    return;
  }

  console.log(`Ready to review ${info.count} changes`);
});
```

### 2. Interactive Review Loop

Step through changes one at a time with user confirmation:

```typescript
async function reviewChanges(context: Word.RequestContext) {
  const info = await DocTree.getTrackedChangesInfo(context);

  for (let i = 0; i < info.count; i++) {
    // Get fresh info after each action
    const current = await DocTree.getTrackedChangesInfo(context);
    if (current.count === 0) break;

    const change = current.changes[0];
    console.log(`Change ${i + 1}/${info.count}: ${change.type} "${change.text}"`);

    // Decision logic here (could prompt user)
    const accept = await getUserDecision();

    if (accept) {
      await DocTree.acceptNextChange(context);
    } else {
      await DocTree.rejectNextChange(context);
    }
  }
}
```

### 3. Scope-Aware Change Filtering

Find and work with changes in specific sections:

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context, {
    includeTrackedChanges: true,
    changeViewMode: 'markup'
  });

  // Find all paragraphs with changes in a section
  const changedNodes = tree.content.filter(node =>
    node.hasChanges &&
    DocTree.getSectionForRef(tree, node.ref)?.headingText === "Methods"
  );

  console.log(`${changedNodes.length} changed paragraphs in Methods section`);
});
```

### 4. Bulk Operations by Type

Accept or reject changes selectively based on type or author:

```typescript
await Word.run(async (context) => {
  const info = await DocTree.getTrackedChangesInfo(context);

  // Process changes based on criteria
  for (const change of info.changes) {
    if (change.type === 'ins' && change.author === 'Reviewer A') {
      // Accept insertions from specific author
      await DocTree.acceptNextChange(context);
    } else {
      // Skip to next (would need to track position)
      // For complex filtering, process all then rebuild tree
    }
  }
});
```

### 5. Document Summary Before/After

Compare document state before and after accepting changes:

```typescript
await Word.run(async (context) => {
  // Get stats before accepting
  const beforeTree = await DocTree.buildTree(context, {
    includeTrackedChanges: true,
    changeViewMode: 'original'
  });
  const beforeSummary = await DocTree.getDocumentSummary(context, beforeTree);

  // Accept all changes
  await DocTree.acceptAllChanges(context);

  // Get stats after
  const afterTree = await DocTree.buildTree(context);
  const afterSummary = await DocTree.getDocumentSummary(context, afterTree);

  console.log(`Words: ${beforeSummary.wordCount} -> ${afterSummary.wordCount}`);
});
```

## Error Handling

The tracked changes API may not be available in all Word versions:

```typescript
await Word.run(async (context) => {
  const result = await DocTree.acceptAllChanges(context);

  if (!result.success) {
    if (result.error?.includes('not available')) {
      console.log("Tracked changes API requires Word 2016 or later");
    } else {
      console.error("Error:", result.error);
    }
  }
});
```

## Performance Considerations

1. **Batched Operations**: The implementation uses batched `context.sync()` calls internally for optimal performance
2. **Tree Building**: Building with `includeTrackedChanges: true` adds ~500ms for comparing original/current versions
3. **Large Documents**: For documents with many changes, consider processing in batches or using `acceptAllChanges` instead of stepping through individually

## See Also

- [office-bridge.md](../office-bridge.md) - Main Office Bridge documentation
- [accessibility.md](../accessibility.md) - Python accessibility API
- [comments.md](./comments.md) - Comment handling documentation (if created)
