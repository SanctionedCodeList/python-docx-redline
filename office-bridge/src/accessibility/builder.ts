/**
 * DocTree Accessibility Tree Builder
 *
 * Builds an accessibility tree from a Word document using Office.js APIs.
 * Transforms Word document structure into semantic nodes with stable refs.
 *
 * Based on docs/DOCTREE_SPEC.md sections 4-6.
 */

import type {
  AccessibilityTree,
  AccessibilityNode,
  DocumentMetadata,
  DocumentStats,
  TreeOptions,
  VerbosityLevel,
  SemanticRole,
  Ref,
  TrackedChange,
  Comment,
  FootnoteInfo,
} from './types';
import { SemanticRole as Role } from './types';
import {
  processChangesAndComments,
  applyChangesToNode,
  applyCommentsToNode,
  type ParagraphChanges,
} from './changes';
import { resolveScope } from './scope';

// =============================================================================
// Style to Role Mapping (Section 3.2 of spec)
// =============================================================================

/**
 * Word built-in styles mapped to semantic roles and heading levels.
 */
const STYLE_TO_ROLE: Record<string, { role: SemanticRole; level?: number }> = {
  Title: { role: Role.Heading, level: 1 },
  Heading1: { role: Role.Heading, level: 1 },
  'Heading 1': { role: Role.Heading, level: 1 },
  Heading2: { role: Role.Heading, level: 2 },
  'Heading 2': { role: Role.Heading, level: 2 },
  Heading3: { role: Role.Heading, level: 3 },
  'Heading 3': { role: Role.Heading, level: 3 },
  Heading4: { role: Role.Heading, level: 4 },
  'Heading 4': { role: Role.Heading, level: 4 },
  Heading5: { role: Role.Heading, level: 5 },
  'Heading 5': { role: Role.Heading, level: 5 },
  Heading6: { role: Role.Heading, level: 6 },
  'Heading 6': { role: Role.Heading, level: 6 },
  Quote: { role: Role.Blockquote },
  'Intense Quote': { role: Role.Blockquote },
  'List Paragraph': { role: Role.ListItem },
  ListParagraph: { role: Role.ListItem },
  Normal: { role: Role.Paragraph },
};

/**
 * Extract heading level from style name using pattern matching.
 */
function extractHeadingLevel(styleName: string): number | undefined {
  // Match patterns like "Heading 1", "Heading1", "H1", etc.
  const match = styleName.match(/(?:heading\s*|h)(\d)/i);
  if (match && match[1]) {
    const level = parseInt(match[1], 10);
    if (level >= 1 && level <= 6) {
      return level;
    }
  }
  return undefined;
}

/**
 * Determine semantic role from Word style name.
 */
function getRoleFromStyle(styleName: string): { role: SemanticRole; level?: number } {
  // Check exact match first
  const exact = STYLE_TO_ROLE[styleName];
  if (exact) {
    return exact;
  }

  // Check for heading pattern
  const level = extractHeadingLevel(styleName);
  if (level !== undefined) {
    return { role: Role.Heading, level };
  }

  // Check for quote-related styles
  if (styleName.toLowerCase().includes('quote')) {
    return { role: Role.Blockquote };
  }

  // Check for list-related styles
  if (styleName.toLowerCase().includes('list')) {
    return { role: Role.ListItem };
  }

  // Default to paragraph
  return { role: Role.Paragraph };
}

// =============================================================================
// Ref Generation
// =============================================================================

/**
 * Generate a paragraph ref.
 */
function makeParagraphRef(index: number): Ref {
  return `p:${index}`;
}

/**
 * Generate a table ref.
 */
function makeTableRef(index: number): Ref {
  return `tbl:${index}`;
}

/**
 * Generate a table row ref.
 */
function makeRowRef(tableIndex: number, rowIndex: number): Ref {
  return `tbl:${tableIndex}/row:${rowIndex}`;
}

/**
 * Generate a table cell ref.
 */
function makeCellRef(tableIndex: number, rowIndex: number, cellIndex: number): Ref {
  return `tbl:${tableIndex}/row:${rowIndex}/cell:${cellIndex}`;
}

/**
 * Generate a paragraph ref within a table cell.
 */
function makeCellParagraphRef(
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number
): Ref {
  return `tbl:${tableIndex}/row:${rowIndex}/cell:${cellIndex}/p:${paragraphIndex}`;
}

// =============================================================================
// Office.js Type Declarations (for type safety without full @types/office-js)
// =============================================================================

/**
 * Minimal Office.js Word.RequestContext interface.
 * The actual types come from @types/office-js at runtime.
 */
interface WordRequestContext {
  document: {
    body: {
      paragraphs: WordParagraphCollection;
      tables: WordTableCollection;
      getReviewedText(changeTrackingVersion: string): { value: string };
      getComments(): WordCommentCollection;
    };
  };
  sync(): Promise<void>;
}

interface WordCommentCollection {
  load(properties: string): WordCommentCollection;
  items: WordComment[];
}

interface WordComment {
  id: string;
  authorName: string;
  authorEmail: string;
  content: string;
  createdDate: Date;
  resolved: boolean;
  replies: WordCommentReplyCollection;
  getRange(): WordRange;
  load(properties: string): WordComment;
}

interface WordCommentReplyCollection {
  load(properties: string): WordCommentReplyCollection;
  items: WordCommentReply[];
}

interface WordCommentReply {
  id: string;
  authorName: string;
  authorEmail: string;
  content: string;
  createdDate: Date;
}

interface WordRange {
  text: string;
  load(properties: string): WordRange;
  getReviewedText(changeTrackingVersion: string): { value: string };
}

interface WordParagraphCollection {
  load(properties: string): WordParagraphCollection;
  items: WordParagraph[];
}

interface WordParagraph {
  text: string;
  style: string;
  isListItem: boolean;
  listItem?: WordListItem | null;
  font?: WordFont;
  getReviewedText(changeTrackingVersion: string): { value: string };
}

interface WordListItem {
  level: number;
  listString: string;
}

interface WordFont {
  bold: boolean;
  italic: boolean;
  underline: string;
  name: string;
  size: number;
}

interface WordTableCollection {
  load(properties: string): WordTableCollection;
  items: WordTable[];
}

interface WordTable {
  rowCount: number;
  rows: WordRowCollection;
}

interface WordRowCollection {
  load(properties: string): WordRowCollection;
  items: WordRow[];
}

interface WordRow {
  isHeader: boolean;
  cellCount: number;
  cells: WordCellCollection;
}

interface WordCellCollection {
  load(properties: string): WordCellCollection;
  items: WordCell[];
}

interface WordCell {
  body: {
    paragraphs: WordParagraphCollection;
  };
}

// =============================================================================
// Statistics Collection
// =============================================================================

/**
 * Mutable stats collector during tree building.
 */
interface StatsCollector {
  paragraphs: number;
  tables: number;
  trackedChanges: number;
  comments: number;
  footnotes: number;
}

function createStatsCollector(): StatsCollector {
  return {
    paragraphs: 0,
    tables: 0,
    trackedChanges: 0,
    comments: 0,
    footnotes: 0,
  };
}

function toDocumentStats(collector: StatsCollector): DocumentStats {
  return {
    paragraphs: collector.paragraphs,
    tables: collector.tables,
    trackedChanges: collector.trackedChanges,
    comments: collector.comments,
    footnotes: collector.footnotes > 0 ? collector.footnotes : undefined,
  };
}

// =============================================================================
// Paragraph Processing
// =============================================================================

/**
 * Process a single Word paragraph into an AccessibilityNode.
 */
function processParagraph(
  paragraph: WordParagraph,
  index: number,
  options: TreeOptions,
  stats: StatsCollector
): AccessibilityNode {
  stats.paragraphs++;

  const ref = makeParagraphRef(index);
  const { role, level } = getRoleFromStyle(paragraph.style);

  const node: AccessibilityNode = {
    ref,
    role,
    text: paragraph.text,
  };

  // Add heading level if applicable
  if (role === Role.Heading && level !== undefined) {
    node.level = level;
  }

  // Add style info in standard+ verbosity
  if (options.verbosity !== 'minimal') {
    node.style = {
      name: paragraph.style,
    };
  }

  // Add formatting in full verbosity
  if (options.verbosity === 'full' && paragraph.font) {
    node.style = {
      ...node.style,
      formatting: {
        bold: paragraph.font.bold || undefined,
        italic: paragraph.font.italic || undefined,
        underline: paragraph.font.underline !== 'None' ? true : undefined,
        font: paragraph.font.name || undefined,
        size: paragraph.font.size ? `${paragraph.font.size}pt` : undefined,
      },
    };
  }

  return node;
}

// =============================================================================
// Table Processing
// =============================================================================

/**
 * Process a table cell into an AccessibilityNode.
 * Called after cell paragraphs have been pre-loaded.
 */
function processCellSync(
  cell: WordCell,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  options: TreeOptions,
  stats: StatsCollector
): AccessibilityNode {
  const ref = makeCellRef(tableIndex, rowIndex, cellIndex);
  const paragraphs = cell.body.paragraphs.items;

  // In minimal mode, just concatenate text
  if (options.verbosity === 'minimal') {
    const text = paragraphs.map((p) => p.text).join(' ');
    return {
      ref,
      role: Role.Cell,
      text: text.trim(),
    };
  }

  // In standard/full mode, include child paragraphs
  const children: AccessibilityNode[] = [];
  for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
    const p = paragraphs[pIdx];
    if (!p) continue;
    const pRef = makeCellParagraphRef(tableIndex, rowIndex, cellIndex, pIdx);
    const { role, level } = getRoleFromStyle(p.style);

    const childNode: AccessibilityNode = {
      ref: pRef,
      role,
      text: p.text,
    };

    if (role === Role.Heading && level !== undefined) {
      childNode.level = level;
    }

    if (options.verbosity === 'full') {
      childNode.style = { name: p.style };
    }

    children.push(childNode);
    stats.paragraphs++;
  }

  // For single-paragraph cells, flatten to text
  if (children.length === 1 && children[0]) {
    return {
      ref,
      role: Role.Cell,
      text: children[0].text,
    };
  }

  return {
    ref,
    role: Role.Cell,
    children,
  };
}

/**
 * Process a table row into an AccessibilityNode.
 * Called after row cells have been pre-loaded.
 */
function processRowSync(
  row: WordRow,
  tableIndex: number,
  rowIndex: number,
  options: TreeOptions,
  stats: StatsCollector
): AccessibilityNode {
  const ref = makeRowRef(tableIndex, rowIndex);

  const cells: AccessibilityNode[] = [];
  for (let cellIdx = 0; cellIdx < row.cells.items.length; cellIdx++) {
    const cell = row.cells.items[cellIdx];
    if (!cell) continue;
    const cellNode = processCellSync(
      cell,
      tableIndex,
      rowIndex,
      cellIdx,
      options,
      stats
    );
    cells.push(cellNode);
  }

  const node: AccessibilityNode = {
    ref,
    role: Role.Row,
    children: cells,
  };

  if (row.isHeader) {
    node.isHeader = true;
  }

  return node;
}

/**
 * Process all tables with batched loading.
 *
 * Uses 3 context.sync() calls instead of O(tables * rows * cells):
 * 1. Load all table rows
 * 2. Load all row cells
 * 3. Load all cell paragraphs
 */
async function processAllTables(
  tables: WordTableCollection,
  context: WordRequestContext,
  options: TreeOptions,
  stats: StatsCollector
): Promise<AccessibilityNode[]> {
  if (tables.items.length === 0) {
    return [];
  }

  // Phase 1: Queue loading rows for all tables
  for (const table of tables.items) {
    table.rows.load('items/isHeader,items/cellCount');
  }
  await context.sync();

  // Phase 2: Queue loading cells for all rows
  for (const table of tables.items) {
    for (const row of table.rows.items) {
      row.cells.load('items');
    }
  }
  await context.sync();

  // Phase 3: Queue loading paragraphs for all cells
  for (const table of tables.items) {
    for (const row of table.rows.items) {
      for (const cell of row.cells.items) {
        cell.body.paragraphs.load('items/text,items/style');
      }
    }
  }
  await context.sync();

  // Phase 4: Build nodes synchronously (no more async calls needed)
  const tableNodes: AccessibilityNode[] = [];

  for (let tableIndex = 0; tableIndex < tables.items.length; tableIndex++) {
    const table = tables.items[tableIndex];
    if (!table) continue;

    stats.tables++;
    const ref = makeTableRef(tableIndex);
    const colCount = table.rows.items[0]?.cellCount ?? 0;

    const rows: AccessibilityNode[] = [];
    for (let rowIdx = 0; rowIdx < table.rows.items.length; rowIdx++) {
      const row = table.rows.items[rowIdx];
      if (!row) continue;
      const rowNode = processRowSync(row, tableIndex, rowIdx, options, stats);
      rows.push(rowNode);
    }

    tableNodes.push({
      ref,
      role: Role.Table,
      dimensions: {
        rows: table.rowCount,
        cols: colCount,
      },
      children: rows,
    });
  }

  return tableNodes;
}

// =============================================================================
// Content Order Tracking
// =============================================================================

/**
 * Track content order between paragraphs and tables.
 *
 * Word documents interleave paragraphs and tables in document order.
 * We need to reconstruct this order for correct ref assignment.
 */
interface ContentPosition {
  type: 'paragraph' | 'table';
  index: number;
}

/**
 * Build content ordering from document structure.
 *
 * Note: Office.js doesn't directly expose interleaved order.
 * This is a simplified approach that processes paragraphs then tables.
 * For true document order, we'd need to use body.getRange() and compare positions.
 */
async function getContentOrder(
  context: WordRequestContext
): Promise<ContentPosition[]> {
  // For now, assume paragraphs come before tables in simple documents.
  // A more sophisticated implementation would use Range.compareLocationWith()
  // to determine true document order.

  const paragraphs = context.document.body.paragraphs.load('items');
  const tables = context.document.body.tables.load('items');
  await context.sync();

  const order: ContentPosition[] = [];

  // Add paragraphs
  for (let i = 0; i < paragraphs.items.length; i++) {
    order.push({ type: 'paragraph', index: i });
  }

  // Add tables
  for (let i = 0; i < tables.items.length; i++) {
    order.push({ type: 'table', index: i });
  }

  return order;
}

// =============================================================================
// Tree Builder
// =============================================================================

/**
 * Default tree options.
 */
const DEFAULT_OPTIONS: TreeOptions = {
  verbosity: 'standard',
  changeViewMode: 'markup',
  contentMode: 'content',
  viewMode: {
    includeBody: true,
    includeHeaders: false,
    includeComments: false,
    includeTrackedChanges: true,
    includeFormatting: false,
  },
};

/**
 * Merge user options with defaults.
 */
function mergeOptions(userOptions?: TreeOptions): TreeOptions {
  if (!userOptions) {
    return { ...DEFAULT_OPTIONS };
  }

  return {
    ...DEFAULT_OPTIONS,
    ...userOptions,
    viewMode: {
      ...DEFAULT_OPTIONS.viewMode,
      ...userOptions.viewMode,
    },
  };
}

// =============================================================================
// Footnote Extraction
// =============================================================================

/**
 * Result of footnote extraction from OOXML.
 */
interface FootnoteExtractionResult {
  /** All footnotes found in the document */
  footnotes: FootnoteInfo[];
  /** Map of paragraph index to footnote refs it contains */
  paragraphFootnotes: Map<number, Ref[]>;
}

/**
 * Extract footnotes from document OOXML.
 *
 * Parses the <w:footnotes> section and <w:footnoteReference> elements
 * to build a list of footnotes with their content and references.
 *
 * @param context - Word.RequestContext from Office.js
 * @returns FootnoteExtractionResult with footnotes and paragraph mappings
 */
async function extractFootnotes(
  context: WordRequestContext
): Promise<FootnoteExtractionResult> {
  const footnotes: FootnoteInfo[] = [];
  const paragraphFootnotes = new Map<number, Ref[]>();

  try {
    // Get OOXML which contains footnotes
    const body = context.document.body;
    const ooxml = body.getOoxml();
    await context.sync();

    const xml = ooxml.value;

    // Find the footnotes section
    const footnotesMatch = xml.match(/<w:footnotes[^>]*>([\s\S]*?)<\/w:footnotes>/);

    if (footnotesMatch) {
      // Extract individual footnotes
      const fnPattern = /<w:footnote[^>]*w:id="(\d+)"[^>]*>([\s\S]*?)<\/w:footnote>/g;
      let match;

      while ((match = fnPattern.exec(footnotesMatch[1])) !== null) {
        const id = parseInt(match[1], 10);
        const content = match[2];

        // Skip separator footnotes (id 0 and -1)
        if (id <= 0) continue;

        // Extract text from the footnote
        const textParts = content.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
        const text = textParts
          .map((t) => t.replace(/<[^>]+>/g, ''))
          .join('')
          .trim();

        if (text) {
          footnotes.push({
            ref: `fn:${id}` as Ref,
            id,
            text,
          });
        }
      }
    }

    // Find footnote references in paragraphs to map them
    // The body XML contains <w:footnoteReference w:id="N"/> within paragraphs
    const bodyMatch = xml.match(/<w:body[^>]*>([\s\S]*?)<\/w:body>/);
    if (bodyMatch) {
      // Find all paragraphs and their footnote references
      const paraPattern = /<w:p[^>]*>([\s\S]*?)<\/w:p>/g;
      let paraMatch;
      let paraIndex = 0;

      while ((paraMatch = paraPattern.exec(bodyMatch[1])) !== null) {
        const paraContent = paraMatch[1];
        const fnRefPattern = /<w:footnoteReference[^>]*w:id="(\d+)"[^>]*\/>/g;
        let fnRefMatch;
        const refs: Ref[] = [];

        while ((fnRefMatch = fnRefPattern.exec(paraContent)) !== null) {
          const fnId = parseInt(fnRefMatch[1], 10);
          if (fnId > 0) {
            refs.push(`fn:${fnId}` as Ref);

            // Update the footnote's referencedFrom field
            const footnote = footnotes.find((f) => f.id === fnId);
            if (footnote) {
              footnote.referencedFrom = `p:${paraIndex}` as Ref;
            }
          }
        }

        if (refs.length > 0) {
          paragraphFootnotes.set(paraIndex, refs);
        }

        paraIndex++;
      }
    }
  } catch (err) {
    // Log error but don't fail - footnotes are optional
    console.warn('Failed to extract footnotes:', err);
  }

  return { footnotes, paragraphFootnotes };
}

/**
 * Build an accessibility tree from a Word document.
 *
 * This is the main entry point for tree construction. It:
 * 1. Loads document paragraphs and tables via Office.js
 * 2. Processes each element into AccessibilityNodes with refs
 * 3. Detects heading styles and assigns semantic roles
 * 4. Collects tracked changes and comments
 * 5. Extracts footnotes from OOXML
 * 6. Collects document statistics
 *
 * @param context - Word.RequestContext from Office.js
 * @param options - Tree building options (verbosity, view mode, etc.)
 * @returns AccessibilityTree representing the document
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const tree = await buildTree(context, { verbosity: 'standard' });
 *   console.log(tree.document.stats);
 * });
 * ```
 */
export async function buildTree(
  context: WordRequestContext,
  options?: TreeOptions
): Promise<AccessibilityTree> {
  const opts = mergeOptions(options);
  const stats = createStatsCollector();

  // Load paragraphs with required properties
  const paragraphProps = ['items/text', 'items/style'];
  if (opts.verbosity === 'full') {
    paragraphProps.push('items/font');
  }
  const paragraphs = context.document.body.paragraphs.load(paragraphProps.join(','));

  // Load tables
  const tables = context.document.body.tables.load('items/rowCount');

  // Execute the load
  await context.sync();

  // Build content nodes
  const content: AccessibilityNode[] = [];
  const paragraphRefs: Ref[] = [];

  // Process paragraphs
  for (let i = 0; i < paragraphs.items.length; i++) {
    const para = paragraphs.items[i];
    if (!para) continue;
    const pNode = processParagraph(para, i, opts, stats);
    content.push(pNode);
    paragraphRefs.push(pNode.ref);
  }

  // Process tables with batched loading (3 sync calls instead of O(tables*rows*cells))
  const tableNodes = await processAllTables(tables, context, opts, stats);
  content.push(...tableNodes);

  // Extract footnotes from OOXML
  const { footnotes, paragraphFootnotes } = await extractFootnotes(context);
  stats.footnotes = footnotes.length;

  // Apply footnote refs to paragraph nodes
  for (const [paraIndex, fnRefs] of paragraphFootnotes) {
    const node = content[paraIndex];
    if (node && node.role !== Role.Table) {
      node.footnoteRefs = fnRefs;
    }
  }

  // Process tracked changes and comments
  let trackedChanges: TrackedChange[] = [];
  let comments: Comment[] = [];
  let paragraphChanges: Map<Ref, ParagraphChanges> = new Map();

  const shouldIncludeTrackedChanges = opts.viewMode?.includeTrackedChanges !== false;
  const shouldIncludeComments = opts.viewMode?.includeComments === true;

  if (shouldIncludeTrackedChanges || shouldIncludeComments) {
    const changesResult = await processChangesAndComments(
      context,
      paragraphRefs,
      {
        changeViewMode: opts.changeViewMode ?? 'markup',
        includeTrackedChanges: shouldIncludeTrackedChanges,
        includeComments: shouldIncludeComments,
      }
    );

    trackedChanges = changesResult.trackedChanges;
    comments = changesResult.comments;
    paragraphChanges = changesResult.paragraphChanges;

    // Update stats
    stats.trackedChanges = changesResult.stats.trackedChanges;
    stats.comments = changesResult.stats.comments;
  }

  // Apply changes and comments to nodes
  const changeViewMode = opts.changeViewMode ?? 'markup';
  let processedContent = content.map((node) => {
    let processed = node;

    // Apply tracked changes
    if (shouldIncludeTrackedChanges && paragraphChanges.size > 0) {
      processed = applyChangesToNode(processed, paragraphChanges, changeViewMode);
    }

    // Apply comments
    if (shouldIncludeComments && comments.length > 0) {
      processed = applyCommentsToNode(processed, comments);
    }

    return processed;
  });

  // Apply scope filtering if specified
  if (opts.scopeFilter) {
    // Build temporary tree to run scope resolution
    const tempTree: AccessibilityTree = {
      document: {
        verbosity: opts.verbosity ?? 'standard',
        stats: toDocumentStats(stats),
      },
      content: processedContent,
    };
    const scopeResult = resolveScope(tempTree, opts.scopeFilter);
    processedContent = scopeResult.nodes;

    // Update stats to reflect scoped content
    stats.paragraphs = processedContent.filter(
      (n) => n.role === Role.Paragraph || n.role === Role.Heading
    ).length;
    stats.tables = processedContent.filter((n) => n.role === Role.Table).length;
  }

  // Apply legacy scopeRefs filtering if specified
  if (opts.scopeRefs && opts.scopeRefs.length > 0) {
    const scopeSet = new Set(opts.scopeRefs);
    processedContent = processedContent.filter((n) => scopeSet.has(n.ref));

    // Update stats
    stats.paragraphs = processedContent.filter(
      (n) => n.role === Role.Paragraph || n.role === Role.Heading
    ).length;
    stats.tables = processedContent.filter((n) => n.role === Role.Table).length;
  }

  // Build document metadata
  const metadata: DocumentMetadata = {
    verbosity: opts.verbosity ?? 'standard',
    stats: toDocumentStats(stats),
    mode: opts.verbosity === 'minimal' ? 'outline' : 'content',
  };

  // Build tree based on verbosity
  if (opts.verbosity === 'minimal') {
    return {
      document: metadata,
      outline: processedContent,
      trackedChanges: trackedChanges.length > 0 ? trackedChanges : undefined,
      comments: comments.length > 0 ? comments : undefined,
      footnotes: footnotes.length > 0 ? footnotes : undefined,
    };
  }

  return {
    document: metadata,
    content: processedContent,
    trackedChanges: trackedChanges.length > 0 ? trackedChanges : undefined,
    comments: comments.length > 0 ? comments : undefined,
    footnotes: footnotes.length > 0 ? footnotes : undefined,
  };
}

// =============================================================================
// Helper Functions for Tree Traversal
// =============================================================================

/**
 * Find all nodes matching a predicate.
 */
export function findNodes(
  tree: AccessibilityTree,
  predicate: (node: AccessibilityNode) => boolean
): AccessibilityNode[] {
  const results: AccessibilityNode[] = [];
  const nodes = tree.content ?? tree.outline ?? [];

  function traverse(node: AccessibilityNode): void {
    if (predicate(node)) {
      results.push(node);
    }
    if (node.children) {
      for (const child of node.children) {
        traverse(child);
      }
    }
  }

  for (const node of nodes) {
    traverse(node);
  }

  return results;
}

/**
 * Find all headings in the tree.
 */
export function findHeadings(
  tree: AccessibilityTree,
  level?: number
): AccessibilityNode[] {
  return findNodes(tree, (node) => {
    if (node.role !== Role.Heading) return false;
    if (level !== undefined && node.level !== level) return false;
    return true;
  });
}

/**
 * Find all tables in the tree.
 */
export function findTables(tree: AccessibilityTree): AccessibilityNode[] {
  return findNodes(tree, (node) => node.role === Role.Table);
}

/**
 * Find a node by its ref.
 */
export function findByRef(
  tree: AccessibilityTree,
  ref: Ref
): AccessibilityNode | undefined {
  const results = findNodes(tree, (node) => node.ref === ref);
  return results[0];
}

/**
 * Get the total node count in a tree.
 */
export function getNodeCount(tree: AccessibilityTree): number {
  let count = 0;
  const nodes = tree.content ?? tree.outline ?? [];

  function traverse(node: AccessibilityNode): void {
    count++;
    if (node.children) {
      for (const child of node.children) {
        traverse(child);
      }
    }
  }

  for (const node of nodes) {
    traverse(node);
  }

  return count;
}
