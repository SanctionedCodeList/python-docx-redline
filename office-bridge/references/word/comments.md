---
name: office-bridge-comments
description: "Complete reference for Word document comment operations via Office Bridge. Covers adding, replying, resolving, and deleting comments programmatically."
---

# Office Bridge: Comments API

This document provides comprehensive documentation for working with comments in Word documents through the Office Bridge add-in. Comments enable collaborative review workflows, allowing reviewers to annotate specific document sections without modifying content.

## Overview

The Office Bridge Comments API provides:

- **Add comments** to paragraphs by ref or to the current selection
- **Reply to comments** for threaded discussions
- **Resolve/unresolve** comments to track review progress
- **Delete comments** when no longer needed
- **List all comments** with metadata (author, content, reply count)

All comment operations return a `CommentOperationResult` indicating success/failure with relevant details.

## Result Structures

### CommentOperationResult

Returned by all comment mutation operations:

```typescript
interface CommentOperationResult {
  /** Whether the operation succeeded */
  success: boolean;
  /** ID of the affected comment (if applicable) */
  commentId?: string;
  /** Error message if failed */
  error?: string;
}
```

### Comment Listing Result

Returned by `getComments()`:

```typescript
interface CommentsResult {
  success: boolean;
  comments: Array<{
    id: string;           // Unique comment identifier
    content: string;      // Comment text
    author: string;       // Author display name
    resolved: boolean;    // Whether marked as resolved
    replyCount: number;   // Number of replies
  }>;
  error?: string;
}
```

## Adding Comments

### addComment(context, ref, commentText)

Add a comment to a paragraph identified by its ref.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.addComment(
    context,
    "p:5",
    "Please review this section for accuracy"
  );

  if (result.success) {
    console.log(`Added comment: ${result.commentId}`);
  } else {
    console.error(`Failed: ${result.error}`);
  }
});
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `context` | `Word.RequestContext` | Office.js request context |
| `ref` | `Ref` | Reference to the paragraph (e.g., `"p:5"`) |
| `commentText` | `string` | The comment content |

**Returns:** `CommentOperationResult` with the new comment's ID on success.

**Notes:**
- Comments can only be added to paragraph refs (type `p:`)
- The comment anchors to the entire paragraph content
- Returns an error if the ref is not a paragraph type

### addCommentToSelection(context, commentText)

Add a comment to the user's current selection in Word.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.addCommentToSelection(
    context,
    "This needs clarification"
  );

  if (result.success) {
    console.log(`Added comment to selection: ${result.commentId}`);
  }
});
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `context` | `Word.RequestContext` | Office.js request context |
| `commentText` | `string` | The comment content |

**Returns:** `CommentOperationResult` with the new comment's ID on success.

**Use Cases:**
- When the user has selected specific text to annotate
- For interactive workflows where selection-based commenting is preferred
- When the exact paragraph ref is not known

## Replying to Comments

### replyToComment(context, commentId, replyText)

Add a reply to an existing comment, creating a threaded discussion.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.replyToComment(
    context,
    "comment-123",
    "I've addressed this concern in the latest revision"
  );

  if (result.success) {
    console.log("Reply added successfully");
  }
});
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `context` | `Word.RequestContext` | Office.js request context |
| `commentId` | `string` | ID of the comment to reply to |
| `replyText` | `string` | The reply content |

**Returns:** `CommentOperationResult`

**Notes:**
- The comment ID is obtained from `addComment()` result or `getComments()` listing
- Replies appear threaded under the parent comment in Word
- Multiple replies can be added to the same comment

## Resolving Comments

### resolveComment(context, commentId)

Mark a comment as resolved, indicating the issue has been addressed.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.resolveComment(context, "comment-123");

  if (result.success) {
    console.log("Comment resolved");
  }
});
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `context` | `Word.RequestContext` | Office.js request context |
| `commentId` | `string` | ID of the comment to resolve |

**Returns:** `CommentOperationResult`

### unresolveComment(context, commentId)

Reopen a previously resolved comment.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.unresolveComment(context, "comment-123");

  if (result.success) {
    console.log("Comment reopened");
  }
});
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `context` | `Word.RequestContext` | Office.js request context |
| `commentId` | `string` | ID of the comment to unresolve |

**Returns:** `CommentOperationResult`

**Use Cases:**
- When a resolved issue needs further discussion
- When resolution was premature
- For reopening items after additional review

## Deleting Comments

### deleteComment(context, commentId)

Permanently remove a comment from the document.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.deleteComment(context, "comment-123");

  if (result.success) {
    console.log("Comment deleted");
  }
});
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `context` | `Word.RequestContext` | Office.js request context |
| `commentId` | `string` | ID of the comment to delete |

**Returns:** `CommentOperationResult`

**Warning:** This operation is irreversible. All replies to the comment are also deleted.

## Listing Comments

### getComments(context)

Retrieve all comments in the document with their metadata.

```typescript
await Word.run(async (context) => {
  const result = await DocTree.getComments(context);

  if (result.success) {
    console.log(`Found ${result.comments.length} comments`);

    for (const comment of result.comments) {
      console.log(`
        ID: ${comment.id}
        Author: ${comment.author}
        Content: ${comment.content}
        Resolved: ${comment.resolved}
        Replies: ${comment.replyCount}
      `);
    }
  }
});
```

**Returns:**

```typescript
{
  success: boolean;
  comments: Array<{
    id: string;
    content: string;
    author: string;
    resolved: boolean;
    replyCount: number;
  }>;
  error?: string;
}
```

**Use Cases:**
- Building a comment summary or dashboard
- Finding unresolved comments for review
- Filtering comments by author
- Counting total comments for document statistics

## Comments in the Accessibility Tree

When building the accessibility tree with `includeComments: true`, comment information is attached to the relevant nodes:

```typescript
const tree = await DocTree.buildTree(context, {
  includeComments: true
});
```

In the YAML output, comments appear with their anchored content:

```yaml
content:
  - ref: "p:5"
    role: paragraph
    text: "This clause requires review..."
    comments:
      - id: "comment-123"
        author: "Jane Smith"
        content: "Please verify the effective date"
        resolved: false
        replyCount: 1
```

### Comment Annotations in Tree Nodes

Each `AccessibilityNode` may include:

| Property | Type | Description |
|----------|------|-------------|
| `comments` | `Comment[]` | Array of comments anchored to this element |
| `commentCount` | `number` | Quick count of attached comments |

## Comment Workflow Examples

### Review Workflow

A typical document review workflow using comments:

```typescript
await Word.run(async (context) => {
  // 1. Build tree to understand document structure
  const tree = await DocTree.buildTree(context, { includeComments: true });

  // 2. Find sections needing review
  const legalSection = DocTree.resolveScope(tree, "section:Legal Terms");

  // 3. Add review comments
  for (const node of legalSection.nodes) {
    if (node.text?.includes("shall")) {
      await DocTree.addComment(
        context,
        node.ref,
        "Consider replacing 'shall' with 'must' per style guide"
      );
    }
  }

  // 4. List all comments for summary
  const allComments = await DocTree.getComments(context);
  console.log(`Added ${allComments.comments.length} review comments`);
});
```

### Processing Resolved Comments

Clean up resolved comments after review completion:

```typescript
await Word.run(async (context) => {
  const result = await DocTree.getComments(context);

  if (result.success) {
    // Find resolved comments
    const resolvedComments = result.comments.filter(c => c.resolved);

    console.log(`Found ${resolvedComments.length} resolved comments to clean up`);

    // Delete resolved comments (with caution)
    for (const comment of resolvedComments) {
      await DocTree.deleteComment(context, comment.id);
    }
  }
});
```

### Bulk Comment Addition

Add comments to multiple paragraphs matching criteria:

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);

  // Find all paragraphs containing specific terms
  const targetNodes = tree.content.filter(
    node => node.text?.toLowerCase().includes("confidential")
  );

  // Add comment to each
  const results = [];
  for (const node of targetNodes) {
    const result = await DocTree.addComment(
      context,
      node.ref,
      "Verify confidentiality classification is appropriate"
    );
    results.push(result);
  }

  const successCount = results.filter(r => r.success).length;
  console.log(`Added ${successCount}/${targetNodes.length} comments`);
});
```

### Comment Threading

Build a comment thread with replies:

```typescript
await Word.run(async (context) => {
  // Add initial comment
  const initial = await DocTree.addComment(
    context,
    "p:10",
    "Is this date correct?"
  );

  if (initial.success && initial.commentId) {
    // Add first reply
    await DocTree.replyToComment(
      context,
      initial.commentId,
      "Yes, verified against the original agreement"
    );

    // Add second reply
    await DocTree.replyToComment(
      context,
      initial.commentId,
      "Marking as resolved"
    );

    // Resolve the thread
    await DocTree.resolveComment(context, initial.commentId);
  }
});
```

## Error Handling

Common errors and how to handle them:

```typescript
await Word.run(async (context) => {
  const result = await DocTree.addComment(context, "tbl:0", "Comment text");

  if (!result.success) {
    switch (true) {
      case result.error?.includes("only be added to paragraphs"):
        console.log("Cannot add comments to tables directly");
        break;
      case result.error?.includes("not available"):
        console.log("Comment API requires Word 2016 or later");
        break;
      case result.error?.includes("not found"):
        console.log("The specified comment does not exist");
        break;
      default:
        console.error(`Comment operation failed: ${result.error}`);
    }
  }
});
```

## API Availability

The Comments API requires:

- **Word 2016** or later (desktop)
- **Word Online** (web version)
- **Office.js** requirement set: `WordApi 1.4` or later

If the API is not available, operations return:

```typescript
{
  success: false,
  error: "Comment API not available in this version of Word"
}
```

## Best Practices

### 1. Check API Availability

Before performing comment operations, verify the API is available:

```typescript
const testResult = await DocTree.getComments(context);
if (!testResult.success && testResult.error?.includes("not available")) {
  console.log("Comment features not supported in this Word version");
  return;
}
```

### 2. Use Meaningful Comment Text

Comments should be actionable and clear:

```typescript
// Good - specific and actionable
"Please verify the effective date matches the signed agreement (p.3)"

// Avoid - vague
"Check this"
```

### 3. Leverage Comment Threading

Use replies to keep related discussion together:

```typescript
// Instead of multiple separate comments
await DocTree.replyToComment(context, originalCommentId, "Follow-up note");
```

### 4. Resolve When Complete

Mark comments as resolved rather than deleting them to preserve review history:

```typescript
await DocTree.resolveComment(context, commentId);
// Only delete if truly not needed for audit trail
```

### 5. Batch Operations Carefully

When adding many comments, consider user experience:

```typescript
// Add comments in a single Word.run() for performance
await Word.run(async (context) => {
  for (const ref of refsToComment) {
    await DocTree.addComment(context, ref, commentText);
  }
  // Single sync happens at end of Word.run()
});
```

## See Also

- [office-bridge.md](../office-bridge.md) - Main Office Bridge documentation
- [tracked-changes.md](./tracked-changes.md) - Tracked changes operations
- [DOCTREE_SPEC.md](../../../docs/DOCTREE_SPEC.md) - Full DocTree specification
