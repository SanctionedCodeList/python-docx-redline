---
name: tree-building
description: "Accessibility tree building and YAML serialization for the Office Bridge Word Add-in. Covers buildTree() API, AccessibilityTree/AccessibilityNode structures, TreeOptions configuration, and YAML output at minimal/standard/full verbosity levels."
---

# Tree Building and YAML Serialization

The Office Bridge builds an accessibility tree from Word documents using Office.js APIs. This tree provides a semantic, agent-friendly representation with stable refs for precise editing.

## Quick Reference

```typescript
// Build tree with default options
const tree = await DocTree.buildTree(context);

// Build with specific options
const tree = await DocTree.buildTree(context, {
  verbosity: 'standard',
  changeViewMode: 'markup',
  viewMode: {
    includeTrackedChanges: true,
    includeComments: true,
  },
});

// Serialize to YAML
const yaml = DocTree.toStandardYaml(tree);  // ~1,500 tokens/page
```

## buildTree(context, options?)

Builds an accessibility tree from the current Word document.

### Signature

```typescript
async function buildTree(
  context: Word.RequestContext,
  options?: TreeOptions
): Promise<AccessibilityTree>
```

### Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `context` | `Word.RequestContext` | The Office.js request context from `Word.run()` |
| `options` | `TreeOptions` | Optional configuration for tree building |

### TreeOptions Interface

```typescript
interface TreeOptions {
  /** Output verbosity level */
  verbosity?: 'minimal' | 'standard' | 'full';

  /** How to display tracked changes */
  changeViewMode?: 'final' | 'original' | 'markup';

  /** Content vs styling focus */
  contentMode?: 'content' | 'styling';

  /** Fine-grained view mode control */
  viewMode?: {
    includeBody?: boolean;          // Default: true
    includeHeaders?: boolean;       // Default: false
    includeComments?: boolean;      // Default: false
    includeTrackedChanges?: boolean; // Default: true
    includeFormatting?: boolean;    // Default: false
  };

  /** Maximum tokens for output (triggers truncation) */
  maxTokens?: number;

  /** Section detection heuristics */
  sectionDetection?: SectionDetectionConfig;

  /** Filter to specific refs only */
  scopeRefs?: Ref[];

  /** Filter by scope specification */
  scopeFilter?: ScopeSpec;
}
```

### Option Details

#### verbosity

Controls the level of detail in the output:

| Level | Description | Token Estimate |
|-------|-------------|----------------|
| `minimal` | Structure overview, refs, truncated text | ~500 tokens/page |
| `standard` | Full content with styles (default) | ~1,500 tokens/page |
| `full` | Run-level detail with all formatting | ~3,000 tokens/page |

#### changeViewMode

How tracked changes appear in text:

| Mode | Description |
|------|-------------|
| `final` | Show document as if all changes were accepted |
| `original` | Show document as if all changes were rejected |
| `markup` | Show all changes with `[+inserted+]` and `[-deleted-]` markers |

#### viewMode Options

Fine-grained control over included content:

```typescript
// Include everything
const tree = await DocTree.buildTree(context, {
  viewMode: {
    includeBody: true,
    includeHeaders: true,
    includeComments: true,
    includeTrackedChanges: true,
    includeFormatting: true,
  },
});

// Minimal content view (lower token cost)
const tree = await DocTree.buildTree(context, {
  viewMode: {
    includeBody: true,
    includeHeaders: false,
    includeComments: false,
    includeTrackedChanges: false,
    includeFormatting: false,
  },
});
```

### Return Value: AccessibilityTree

```typescript
interface AccessibilityTree {
  /** Document metadata and stats */
  document: DocumentMetadata;

  /** Content nodes (standard/full mode) */
  content?: AccessibilityNode[];

  /** Outline nodes (minimal mode) */
  outline?: AccessibilityNode[];

  /** All tracked changes in document */
  trackedChanges?: TrackedChange[];

  /** All comments in document */
  comments?: Comment[];

  /** All bookmarks */
  bookmarks?: BookmarkInfo[];

  /** All footnotes */
  footnotes?: FootnoteInfo[];

  /** All endnotes */
  endnotes?: EndnoteInfo[];

  /** Navigation hints (minimal mode) */
  navigation?: {
    expandSection: string;
    search: string;
  };
}
```

### DocumentMetadata Structure

```typescript
interface DocumentMetadata {
  /** Document path or filename */
  path?: string;

  /** Verbosity level of this tree */
  verbosity: VerbosityLevel;

  /** Document statistics */
  stats: DocumentStats;

  /** Section detection info (outline mode) */
  sectionDetection?: SectionDetectionInfo;

  /** Current mode */
  mode?: 'outline' | 'content' | 'styling';
}
```

### DocumentStats Structure

```typescript
interface DocumentStats {
  paragraphs: number;
  tables: number;
  trackedChanges: number;
  comments: number;
  footnotes?: number;
  endnotes?: number;
  sections?: number;
}
```

## AccessibilityNode Structure

Each element in the tree is represented as an `AccessibilityNode`:

```typescript
interface AccessibilityNode {
  /** Stable reference (e.g., "p:5", "tbl:0/row:1/cell:2") */
  ref: Ref;

  /** Semantic role */
  role: SemanticRole;

  /** Text content */
  text?: string;

  /** Child nodes (for tables, lists, etc.) */
  children?: AccessibilityNode[];

  /** Style information */
  style?: NodeStyle;

  // --- Tracked Changes ---
  changes?: TrackedChange[];
  changeRefs?: Ref[];
  hasChanges?: boolean;

  // --- Comments ---
  comments?: Comment[];
  commentRefs?: Ref[];
  hasComments?: boolean;

  // --- Images ---
  images?: ImageInfo[];
  floatingImages?: ImageInfo[];

  // --- Links and References ---
  links?: LinkInfo[];
  bookmark?: BookmarkInfo;
  footnoteRefs?: Ref[];
  endnoteRefs?: Ref[];
  incomingReferences?: Ref[];

  // --- Role-Specific Properties ---
  level?: number;           // For headings (1-6)
  isHeader?: boolean;       // For table rows
  dimensions?: { rows: number; cols: number };  // For tables
  runs?: AccessibilityNode[];  // For full verbosity

  // --- Section Properties (outline mode) ---
  headingRef?: Ref;
  paragraphCount?: number;
  tables?: Ref[];
  trackedChangesCount?: number;
  preview?: string;
  detection?: 'heading_style' | 'outline_level' | 'bold_heuristic' | 'caps_heuristic' | 'numbered_section' | 'fallback';
  confidence?: 'high' | 'medium' | 'low';
}
```

### SemanticRole Values

```typescript
enum SemanticRole {
  // Document Landmarks
  Document = 'document',
  Header = 'header',
  Footer = 'footer',
  Section = 'section',

  // Structural Elements
  Heading = 'heading',
  Paragraph = 'paragraph',
  Blockquote = 'blockquote',
  List = 'list',
  ListItem = 'listitem',
  Table = 'table',
  Row = 'row',
  Cell = 'cell',

  // Inline Elements
  Text = 'text',
  Strong = 'strong',
  Emphasis = 'emphasis',
  Link = 'link',

  // Annotations
  Insertion = 'insertion',
  Deletion = 'deletion',
  Comment = 'comment',

  // Objects
  Image = 'image',
  Chart = 'chart',
  Diagram = 'diagram',
  Shape = 'shape',

  // References
  Footnote = 'footnote',
  Endnote = 'endnote',
  Bookmark = 'bookmark',
}
```

### NodeStyle Structure

```typescript
interface NodeStyle {
  /** Named style (e.g., "Heading1", "Normal") */
  name?: string;

  /** Direct formatting overrides */
  formatting?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
    font?: string;
    size?: string;      // e.g., "12pt"
    color?: string;     // e.g., "#0000FF"
    highlight?: string; // e.g., "yellow"
  };

  /** Paragraph formatting */
  paragraph?: {
    alignment?: 'left' | 'center' | 'right' | 'justify';
    spacingBefore?: string;
    spacingAfter?: string;
    lineSpacing?: string;
    indent?: string;
  };
}
```

## Ref Format

Refs are stable identifiers for addressing document elements:

| Pattern | Example | Description |
|---------|---------|-------------|
| `p:N` | `p:5` | Paragraph at index N |
| `tbl:N` | `tbl:0` | Table at index N |
| `tbl:N/row:M` | `tbl:0/row:2` | Row M in table N |
| `tbl:N/row:M/cell:K` | `tbl:0/row:2/cell:1` | Cell K in row M of table N |
| `tbl:N/row:M/cell:K/p:J` | `tbl:0/row:2/cell:1/p:0` | Paragraph J in cell |
| `fn:N` | `fn:1` | Footnote with ID N |
| `cmt:ID` | `cmt:abc123` | Comment with ID |
| `ins:ID` | `ins:42` | Tracked insertion |
| `del:ID` | `del:43` | Tracked deletion |

## YAML Serialization Functions

Three convenience functions for different verbosity levels:

### toMinimalYaml(tree)

Compact structure overview for navigation and planning.

```typescript
const yaml = DocTree.toMinimalYaml(tree);
```

**Output characteristics:**
- ~500 tokens per page
- Refs and truncated text (40 chars max)
- Table dimensions without cell content
- Key states: `[header]`, `[has-changes]`
- Uses `outline:` section instead of `content:`

**Example output:**
```yaml
document:
  verbosity: minimal
  stats:
    paragraphs: 45
    tables: 2
    tracked_changes: 12
    comments: 3

outline:
  - h1 "SERVICES AGREEMENT" [ref=p:0]
  - p "This Agreement is entered into..." [ref=p:1]
  - h2 "1. DEFINITIONS" [ref=p:2]
  - p "1.1 \"Agreement\" means this..." [ref=p:3] [has-changes]
  - table [4x3] [ref=tbl:0]
```

**Best for:**
- Initial document exploration
- Document structure overview
- Navigation planning
- Low-context operations

### toStandardYaml(tree)

Balanced detail for most editing operations (default).

```typescript
const yaml = DocTree.toStandardYaml(tree);
```

**Output characteristics:**
- ~1,500 tokens per page
- Full text content
- Style names
- Change and comment refs
- Inline format for simple paragraphs

**Example output:**
```yaml
document:
  verbosity: standard
  stats:
    paragraphs: 45
    tables: 2
    tracked_changes: 12
    comments: 3

content:
  - heading [ref=p:0] [level=1]: "SERVICES AGREEMENT"
  - paragraph [ref=p:1]: "This Agreement is entered into as of the Effective Date."
  - heading [ref=p:2] [level=2]: "1. DEFINITIONS"
  - paragraph [ref=p:3] [has-changes]:
      text: "1.1 \"Agreement\" means this [+Services+] Agreement."
      style: Normal
      change_refs: [ins:42]
  - table [ref=tbl:0] [rows=4] [cols=3]:
      - row [ref=tbl:0/row:0] [header]:
          - cell: "Term"
          - cell: "Definition"
          - cell: "Section"

tracked_changes:
  - ref: ins:42
    type: insertion
    author: "John Smith"
    date: "2025-01-15T10:30:00Z"
    text: "Services"
    location: p:3
```

**Best for:**
- Content editing and review
- Search and replace operations
- Most document manipulation tasks
- Tracked change review

### toFullYaml(tree)

Complete fidelity with run-level detail.

```typescript
const yaml = DocTree.toFullYaml(tree);
```

**Output characteristics:**
- ~3,000 tokens per page
- Run-level formatting detail
- All style information
- Image size in EMUs
- Bookmark references

**Example output:**
```yaml
document:
  verbosity: full
  stats:
    paragraphs: 45
    tables: 2
    tracked_changes: 12
    comments: 3

content:
  - heading [ref=p:0] [level=1]:
      style: Heading 1
      formatting:
        bold: true
        size: 16pt
        color: "#2F5496"
      runs:
        - text "SERVICES AGREEMENT" [ref=p:0/r:0] [bold]
  - paragraph [ref=p:1]:
      style: Normal
      formatting:
        font: Calibri
        size: 11pt
      runs:
        - text "This Agreement is entered into as of the " [ref=p:1/r:0]
        - text "Effective Date" [ref=p:1/r:1] [bold]
        - text "." [ref=p:1/r:2]
```

**Best for:**
- Precise formatting operations
- Style analysis
- Document comparison
- Format-sensitive edits

### Generic treeToYaml(tree, verbosity)

The underlying function used by convenience wrappers:

```typescript
function treeToYaml(
  tree: AccessibilityTree,
  verbosity: VerbosityLevel = 'standard'
): string
```

## Token Estimates by Document Size

| Document Size | Minimal | Standard | Full |
|--------------|---------|----------|------|
| 10 pages | ~5,000 | ~15,000 | ~30,000 |
| 50 pages | ~25,000 | ~75,000 | ~150,000 |
| 100 pages | ~50,000 | ~150,000 | ~300,000 |

## Best Practices

### Choosing Verbosity Level

1. **Start with minimal** for large or unfamiliar documents
2. **Use standard** for most editing operations
3. **Use full** only when formatting precision matters

### Progressive Disclosure Pattern

```typescript
// Step 1: Get overview
const overview = await DocTree.buildTree(context, { verbosity: 'minimal' });
const minimalYaml = DocTree.toMinimalYaml(overview);

// Step 2: Focus on section of interest
const focused = await DocTree.buildTree(context, {
  verbosity: 'standard',
  scopeFilter: 'section:Definitions',
});

// Step 3: Get full detail for specific paragraph
const detailed = await DocTree.buildTree(context, {
  verbosity: 'full',
  scopeRefs: ['p:3', 'p:4', 'p:5'],
});
```

### Tracked Changes Workflow

```typescript
// See all changes in markup view
const tree = await DocTree.buildTree(context, {
  changeViewMode: 'markup',
  viewMode: { includeTrackedChanges: true },
});

// Preview final document
const final = await DocTree.buildTree(context, {
  changeViewMode: 'final',
});

// Preview original document
const original = await DocTree.buildTree(context, {
  changeViewMode: 'original',
});
```

### Comments Review

```typescript
const tree = await DocTree.buildTree(context, {
  viewMode: {
    includeComments: true,
    includeTrackedChanges: true,
  },
});

// Access comments
for (const comment of tree.comments ?? []) {
  console.log(`${comment.author}: ${comment.text}`);
  console.log(`On: "${comment.onText}"`);
  console.log(`Replies: ${comment.replies?.length ?? 0}`);
}
```

## Performance Notes

### Batched Loading

The tree builder uses batched `context.sync()` calls to minimize round-trips:

- **Paragraphs**: Single sync to load all paragraphs
- **Tables**: 3 syncs regardless of table count (rows, cells, cell paragraphs)
- **Changes/Comments**: Single sync for all tracked changes and comments

### Typical Performance

| Document Size | Build Time |
|--------------|------------|
| 50 paragraphs | ~500ms |
| 200 paragraphs | ~1,000ms |
| 500 paragraphs | ~2,000ms |

### Optimization Tips

1. **Use scopeRefs** to limit tree building to specific paragraphs
2. **Use scopeFilter** to exclude irrelevant content
3. **Avoid full verbosity** unless formatting details are needed
4. **Cache trees** when making multiple edits to same document

```typescript
// Build once, use for multiple operations
const tree = await DocTree.buildTree(context);

// Use tree for multiple scope-aware operations
await DocTree.replaceByScope(context, tree, "section:Terms", newText);
await DocTree.formatByScope(context, tree, "section:Terms", { bold: true });
```

## Helper Functions

### findNodes(tree, predicate)

Find all nodes matching a predicate:

```typescript
const allHeadings = findNodes(tree, node => node.role === SemanticRole.Heading);
const changedNodes = findNodes(tree, node => node.hasChanges === true);
```

### findHeadings(tree, level?)

Find all headings, optionally filtered by level:

```typescript
const allHeadings = findHeadings(tree);
const h2Headings = findHeadings(tree, 2);
```

### findTables(tree)

Find all tables in the tree:

```typescript
const tables = findTables(tree);
```

### findByRef(tree, ref)

Find a specific node by ref:

```typescript
const node = findByRef(tree, "p:5");
```

### getNodeCount(tree)

Get total node count including nested children:

```typescript
const count = getNodeCount(tree);
```

## See Also

- [office-bridge.md](../office-bridge.md) - Full Office Bridge API reference
- [accessibility.md](../accessibility.md) - Python accessibility API
- [editing.md](../editing.md) - Ref-based editing operations
