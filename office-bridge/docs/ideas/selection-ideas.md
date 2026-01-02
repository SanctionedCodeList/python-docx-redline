---
name: office-bridge-selection-ideas
description: "Roadmap ideas for selection and cursor operations in Office Bridge. Reference when planning new selection-related features."
---

# Selection API Enhancement Ideas

This document outlines potential enhancements for selection and cursor operations in the Office Bridge add-in.

## Priority 1: Core Selection Helpers

### getSelection()

Get the current selection with rich context.

```typescript
interface SelectionInfo {
  text: string;
  isEmpty: boolean;
  ref: Ref | null;           // p:5 if selection is in a paragraph
  startRef: Ref | null;      // Start paragraph if multi-paragraph
  endRef: Ref | null;        // End paragraph if multi-paragraph
  context: 'MainDoc' | 'Header' | 'Footer' | 'Footnote' | 'Endnote';
  contentControlId?: string; // If inside a content control
  tableRef?: Ref;            // tbl:0 if inside a table
  cellRef?: Ref;             // tbl:0/row:1/cell:2 if in a cell
}

async function getSelection(context: WordRequestContext): Promise<SelectionInfo>;
```

**Use Case**: Agent needs to understand what the user is looking at before taking action.

### selectByRef(context, ref, mode?)

Select content by ref, scrolling it into view.

```typescript
async function selectByRef(
  context: WordRequestContext,
  ref: Ref,
  mode?: 'select' | 'start' | 'end'
): Promise<void>;
```

**Use Case**: After finding relevant content with scope queries, navigate user to it.

```typescript
// Find and navigate to a section
const result = DocTree.resolveScope(tree, 'section:Definitions');
if (result.nodes.length > 0) {
  await DocTree.selectByRef(context, result.nodes[0].ref);
}
```

### insertAtSelection(context, text, options?)

Insert text at current selection with tracked changes support.

```typescript
interface InsertOptions {
  location?: 'before' | 'after' | 'replace' | 'start' | 'end';
  track?: boolean;
  author?: string;
}

async function insertAtSelection(
  context: WordRequestContext,
  text: string,
  options?: InsertOptions
): Promise<{ ref: Ref; range: WordRange }>;
```

**Use Case**: User selects text, agent replaces or augments it.

### replaceSelection(context, newText, options?)

Replace selected content with new text.

```typescript
async function replaceSelection(
  context: WordRequestContext,
  newText: string,
  options?: EditOptions
): Promise<EditResult>;
```

**Use Case**: Direct replacement of user-selected text.

## Priority 2: Selection-Based Workflows

### Selection-Driven Agent Pattern

User selects content, agent operates on it:

```typescript
// Get what user selected
const selection = await DocTree.getSelection(context);

if (selection.isEmpty) {
  // No selection - use cursor position context
  console.log('Cursor at:', selection.ref);
} else {
  // Work with selected content
  const text = selection.text;
  // AI processes text...
  const improved = await improveText(text);
  await DocTree.replaceSelection(context, improved, { track: true });
}
```

### Confirm-Before-Edit Pattern

```typescript
async function confirmAndEdit(
  context: WordRequestContext,
  ref: Ref,
  newText: string
): Promise<EditResult> {
  // Navigate user to the content first
  await DocTree.selectByRef(context, ref);
  await context.sync();

  // In a real add-in, show confirmation dialog
  // For now, proceed with edit
  return DocTree.replaceByRef(context, ref, newText, { track: true });
}
```

### Multi-Location Preview

Navigate through multiple edits before applying:

```typescript
interface EditPreview {
  ref: Ref;
  currentText: string;
  proposedText: string;
}

async function previewEdits(
  context: WordRequestContext,
  edits: EditPreview[]
): Promise<void> {
  for (const edit of edits) {
    await DocTree.selectByRef(context, edit.ref);
    await context.sync();
    // Pause for user to see context
    await delay(2000);
  }
}
```

## Priority 3: Range Expansion Helpers

### expandToWord(context, range?)

Expand selection/range to include full word.

```typescript
async function expandToWord(
  context: WordRequestContext,
  range?: WordRange
): Promise<WordRange>;
```

### expandToSentence(context, range?)

Expand to include full sentence.

```typescript
async function expandToSentence(
  context: WordRequestContext,
  range?: WordRange
): Promise<WordRange>;
```

### expandToParagraph(context, range?)

Expand to include full paragraph.

```typescript
async function expandToParagraph(
  context: WordRequestContext,
  range?: WordRange
): Promise<WordRange>;
```

### expandToSection(context, tree, range?)

Expand to include entire section (from heading to next heading).

```typescript
async function expandToSection(
  context: WordRequestContext,
  tree: AccessibilityTree,
  range?: WordRange
): Promise<WordRange>;
```

**Use Case**: User clicks in a section, agent can easily select the entire section for processing.

## Priority 4: Advanced Selection Features

### getSelectionRefs(context)

Get all paragraph refs touched by selection.

```typescript
async function getSelectionRefs(
  context: WordRequestContext
): Promise<Ref[]>;
```

Returns all refs for paragraphs that overlap with the selection:

```typescript
const refs = await DocTree.getSelectionRefs(context);
// ["p:5", "p:6", "p:7"] if selection spans 3 paragraphs
```

### selectRange(context, startRef, endRef)

Select everything between two refs.

```typescript
async function selectRange(
  context: WordRequestContext,
  startRef: Ref,
  endRef: Ref
): Promise<void>;
```

### selectByScope(context, tree, scope)

Select all content matching a scope.

```typescript
// Note: Office.js doesn't support multi-selection
// This would select the range covering all matches
async function selectByScope(
  context: WordRequestContext,
  tree: AccessibilityTree,
  scope: ScopeSpec
): Promise<{ selected: number; range: WordRange }>;
```

## Priority 5: Cursor Position Utilities

### getCursorPosition(context)

Get detailed cursor position information.

```typescript
interface CursorPosition {
  ref: Ref;                    // Current paragraph
  offsetInParagraph: number;   // Character offset within paragraph
  line: number;                // Line number in document (approximate)
  isAtStart: boolean;          // At start of paragraph
  isAtEnd: boolean;            // At end of paragraph
  precedingWord: string;       // Word before cursor
  followingWord: string;       // Word after cursor
}

async function getCursorPosition(
  context: WordRequestContext
): Promise<CursorPosition>;
```

### moveCursor(context, direction, unit, count?)

Move cursor without selecting.

```typescript
type Direction = 'forward' | 'backward' | 'up' | 'down';
type Unit = 'character' | 'word' | 'sentence' | 'paragraph' | 'line';

async function moveCursor(
  context: WordRequestContext,
  direction: Direction,
  unit: Unit,
  count?: number
): Promise<CursorPosition>;
```

### moveToRef(context, ref, position?)

Move cursor to a specific ref.

```typescript
async function moveToRef(
  context: WordRequestContext,
  ref: Ref,
  position?: 'start' | 'end'
): Promise<void>;
```

## Future Considerations

### Collaborative Editing Awareness

Track other users' selections (in co-authoring scenarios):

```typescript
interface UserSelection {
  userId: string;
  displayName: string;
  color: string;
  refs: Ref[];
}

async function getOtherSelections(
  context: WordRequestContext
): Promise<UserSelection[]>;
```

### Selection History

Track recent selections for undo/redo:

```typescript
interface SelectionHistory {
  push(selection: SelectionInfo): void;
  pop(): SelectionInfo | undefined;
  peek(): SelectionInfo | undefined;
  clear(): void;
}
```

### Smart Selection Suggestions

Suggest what to select based on context:

```typescript
interface SelectionSuggestion {
  ref: Ref;
  reason: string;
  confidence: number;
}

async function suggestSelections(
  context: WordRequestContext,
  tree: AccessibilityTree,
  intent: string
): Promise<SelectionSuggestion[]>;

// Usage
const suggestions = await suggestSelections(context, tree, 'edit definitions');
// Returns: [{ ref: 'p:15', reason: 'Definition section heading', confidence: 0.9 }]
```

## Implementation Notes

### Handling the Table Selection Bug

Office.js has a known bug where `getSelection()` on partial table selections returns only the last row. Mitigation:

```typescript
async function getSelectionSafe(
  context: WordRequestContext
): Promise<SelectionInfo> {
  const selection = context.document.getSelection();
  selection.load('text,paragraphs,tables');
  await context.sync();

  // Check if we're in a table
  if (selection.tables.items.length > 0) {
    // Use table-specific handling
    return getTableSelectionInfo(context, selection);
  }

  return getNormalSelectionInfo(context, selection);
}
```

### Desktop vs. Web Compatibility

Many selection features are desktop-only. Create fallbacks:

```typescript
async function selectByRefCompat(
  context: WordRequestContext,
  ref: Ref
): Promise<void> {
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  const parsed = parseRef(ref);
  if (parsed.type === 'p' && parsed.index < paragraphs.items.length) {
    const range = paragraphs.items[parsed.index].getRange('Whole');

    // This works on all platforms
    range.select(Word.SelectionMode.select);
    await context.sync();
  }
}
```

### Performance Considerations

Selection operations can be expensive. Batch when possible:

```typescript
// Bad: Multiple syncs
for (const ref of refs) {
  await selectByRef(context, ref);
  await context.sync();
}

// Good: Single operation for preview, batch for edits
await selectByRef(context, refs[0]);  // Show first
await batchEdit(context, edits);      // Apply all
```

## Integration with Existing APIs

These selection utilities should integrate with:

- **buildTree()**: Include current selection ref in tree metadata
- **replaceByRef()**: Option to select result after edit
- **findText()**: Option to select first/all matches
- **batchEdit()**: Option to select final edited range

Example integration:

```typescript
const tree = await DocTree.buildTree(context, {
  includeSelectionContext: true  // New option
});

console.log(tree.selection);  // { ref: 'p:5', isEmpty: false }
```

## See Also

- [selection.md](../selection.md) - Current selection API documentation
- [office-bridge.md](../../office-bridge.md) - Main Office Bridge docs
- [accessibility.md](../../accessibility.md) - DocTree accessibility layer
