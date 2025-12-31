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
} from './types';
import { SemanticRole as Role } from './types';

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
    };
  };
  sync(): Promise<void>;
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
}

function createStatsCollector(): StatsCollector {
  return {
    paragraphs: 0,
    tables: 0,
    trackedChanges: 0,
    comments: 0,
  };
}

function toDocumentStats(collector: StatsCollector): DocumentStats {
  return {
    paragraphs: collector.paragraphs,
    tables: collector.tables,
    trackedChanges: collector.trackedChanges,
    comments: collector.comments,
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
 */
async function processCell(
  cell: WordCell,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  context: WordRequestContext,
  options: TreeOptions,
  stats: StatsCollector
): Promise<AccessibilityNode> {
  const ref = makeCellRef(tableIndex, rowIndex, cellIndex);

  // Load cell paragraphs
  const paragraphs = cell.body.paragraphs.load('items/text,items/style');
  await context.sync();

  // In minimal mode, just concatenate text
  if (options.verbosity === 'minimal') {
    const text = paragraphs.items.map((p) => p.text).join(' ');
    return {
      ref,
      role: Role.Cell,
      text: text.trim(),
    };
  }

  // In standard/full mode, include child paragraphs
  const children: AccessibilityNode[] = [];
  for (let pIdx = 0; pIdx < paragraphs.items.length; pIdx++) {
    const p = paragraphs.items[pIdx];
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
 */
async function processRow(
  row: WordRow,
  tableIndex: number,
  rowIndex: number,
  context: WordRequestContext,
  options: TreeOptions,
  stats: StatsCollector
): Promise<AccessibilityNode> {
  const ref = makeRowRef(tableIndex, rowIndex);

  // Load cells
  row.cells.load('items');
  await context.sync();

  const cells: AccessibilityNode[] = [];
  for (let cellIdx = 0; cellIdx < row.cells.items.length; cellIdx++) {
    const cell = row.cells.items[cellIdx];
    if (!cell) continue;
    const cellNode = await processCell(
      cell,
      tableIndex,
      rowIndex,
      cellIdx,
      context,
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
 * Process a table into an AccessibilityNode.
 */
async function processTable(
  table: WordTable,
  tableIndex: number,
  context: WordRequestContext,
  options: TreeOptions,
  stats: StatsCollector
): Promise<AccessibilityNode> {
  stats.tables++;
  const ref = makeTableRef(tableIndex);

  // Load rows
  table.rows.load('items/isHeader,items/cellCount');
  await context.sync();

  // Get column count from first row
  const colCount = table.rows.items[0]?.cellCount ?? 0;

  const rows: AccessibilityNode[] = [];
  for (let rowIdx = 0; rowIdx < table.rows.items.length; rowIdx++) {
    const row = table.rows.items[rowIdx];
    if (!row) continue;
    const rowNode = await processRow(row, tableIndex, rowIdx, context, options, stats);
    rows.push(rowNode);
  }

  return {
    ref,
    role: Role.Table,
    dimensions: {
      rows: table.rowCount,
      cols: colCount,
    },
    children: rows,
  };
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

/**
 * Build an accessibility tree from a Word document.
 *
 * This is the main entry point for tree construction. It:
 * 1. Loads document paragraphs and tables via Office.js
 * 2. Processes each element into AccessibilityNodes with refs
 * 3. Detects heading styles and assigns semantic roles
 * 4. Collects document statistics
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

  // Process paragraphs
  for (let i = 0; i < paragraphs.items.length; i++) {
    const para = paragraphs.items[i];
    if (!para) continue;
    const pNode = processParagraph(para, i, opts, stats);
    content.push(pNode);
  }

  // Process tables
  for (let i = 0; i < tables.items.length; i++) {
    const tbl = tables.items[i];
    if (!tbl) continue;
    const tableNode = await processTable(tbl, i, context, opts, stats);
    content.push(tableNode);
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
      outline: content,
    };
  }

  return {
    document: metadata,
    content,
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
