/**
 * DocTree Ref-Based Editing Methods
 *
 * Provides methods for editing Word documents using refs from the accessibility tree.
 * These methods enable precise, targeted edits without ambiguous text matching.
 *
 * Based on docs/DOCTREE_SPEC.md sections 6.3 and 7.
 */

import type {
  Ref,
  EditResult,
  InsertPosition,
  NodeStyle,
  AccessibilityTree,
  ScopeSpec,
} from './types';
import { resolveScope } from './scope';

// =============================================================================
// Office.js Type Declarations
// =============================================================================

/**
 * Minimal Office.js Word.RequestContext interface for editing operations.
 */
interface WordRequestContext {
  document: {
    body: WordBody;
    getSelection(): WordRange;
  };
  sync(): Promise<void>;
  trackedChanges?: {
    enabled: boolean;
  };
}

interface WordBody {
  paragraphs: WordParagraphCollection;
  tables: WordTableCollection;
  search(searchText: string, options?: WordSearchOptions): WordRangeCollection;
}

interface WordSearchOptions {
  matchCase?: boolean;
  matchWholeWord?: boolean;
  matchWildcards?: boolean;
  matchPrefix?: boolean;
  matchSuffix?: boolean;
}

interface WordRangeCollection {
  load(properties: string): WordRangeCollection;
  items: WordRange[];
}

interface WordParagraphCollection {
  load(properties: string): WordParagraphCollection;
  items: WordParagraph[];
  getItem(index: number): WordParagraph;
}

interface WordParagraph {
  text: string;
  style: string;
  font: WordFont;
  getRange(rangeLocation?: RangeLocation): WordRange;
  insertText(text: string, insertLocation: InsertLocation): WordRange;
  insertParagraph(paragraphText: string, insertLocation: InsertLocation): WordParagraph;
  delete(): void;
  load(properties: string): WordParagraph;
}

interface WordFont {
  bold: boolean;
  italic: boolean;
  underline: string;
  strikeThrough: boolean;
  color: string;
  highlightColor: string;
  name: string;
  size: number;
  set(properties: Partial<WordFont>): void;
}

interface WordRange {
  text: string;
  font: WordFont;
  insertText(text: string, insertLocation: InsertLocation): WordRange;
  delete(): void;
  load(properties: string): WordRange;
  select(selectionMode?: SelectionMode): void;
  compareLocationWith(range: WordRange): {
    value: LocationRelation;
  };
  track(): void;
  untrack(): void;
}

interface WordTableCollection {
  load(properties: string): WordTableCollection;
  items: WordTable[];
  getItem(index: number): WordTable;
}

interface WordTable {
  rowCount: number;
  rows: WordRowCollection;
  getCell(rowIndex: number, cellIndex: number): WordTableCell;
  delete(): void;
}

interface WordRowCollection {
  load(properties: string): WordRowCollection;
  items: WordRow[];
  getItem(index: number): WordRow;
}

interface WordRow {
  isHeader: boolean;
  cellCount: number;
  cells: WordCellCollection;
  delete(): void;
}

interface WordCellCollection {
  load(properties: string): WordCellCollection;
  items: WordCell[];
  getItem(index: number): WordCell;
}

interface WordCell {
  body: {
    paragraphs: WordParagraphCollection;
  };
}

interface WordTableCell {
  body: {
    paragraphs: WordParagraphCollection;
    getRange(rangeLocation?: RangeLocation): WordRange;
  };
  value: string;
}

type InsertLocation = 'Before' | 'After' | 'Start' | 'End' | 'Replace';
type RangeLocation = 'Whole' | 'Start' | 'End' | 'Before' | 'After' | 'Content';
type SelectionMode = 'Select' | 'Start' | 'End';
type LocationRelation = 'Before' | 'InsideStart' | 'Inside' | 'InsideEnd' | 'After' | 'AdjacentBefore' | 'AdjacentAfter' | 'Contains' | 'Equals';

// =============================================================================
// Editing Options
// =============================================================================

/**
 * Options for editing operations.
 */
export interface EditOptions {
  /** Enable tracked changes for this edit */
  track?: boolean;
  /** Author name for tracked changes */
  author?: string;
  /** Comment to attach to the change */
  comment?: string;
}

/**
 * Formatting options for formatByRef.
 */
export interface FormatOptions {
  /** Apply bold */
  bold?: boolean;
  /** Apply italic */
  italic?: boolean;
  /** Apply underline */
  underline?: boolean;
  /** Apply strikethrough */
  strikethrough?: boolean;
  /** Font name */
  font?: string;
  /** Font size in points */
  size?: number;
  /** Font color (hex) */
  color?: string;
  /** Highlight color */
  highlight?: string;
  /** Paragraph style name */
  style?: string;
}

// =============================================================================
// Ref Parsing
// =============================================================================

/**
 * Parsed ref structure.
 */
interface ParsedRef {
  type: string;
  index: number;
  subRefs: ParsedRef[];
}

/**
 * Parse a ref string into its components.
 *
 * Examples:
 *   "p:3" -> { type: "p", index: 3, subRefs: [] }
 *   "tbl:0/row:2/cell:1" -> { type: "tbl", index: 0, subRefs: [{ type: "row", index: 2, subRefs: [{ type: "cell", index: 1, subRefs: [] }] }] }
 */
function parseRef(ref: Ref): ParsedRef {
  const parts = ref.split('/');
  const [firstPart, ...restParts] = parts;

  if (!firstPart) {
    throw new Error(`Invalid ref format: ${ref}`);
  }

  const match = firstPart.match(/^([a-z]+):(\d+|~[\w]+)$/);
  if (!match) {
    throw new Error(`Invalid ref format: ${ref}`);
  }

  const type = match[1];
  const indexStr = match[2];

  // Handle fingerprint refs (e.g., p:~xK4mNp2q)
  if (indexStr?.startsWith('~')) {
    // TODO: Implement fingerprint resolution
    throw new Error(`Fingerprint refs not yet supported: ${ref}`);
  }

  const index = parseInt(indexStr ?? '0', 10);

  const subRefs: ParsedRef[] = [];
  if (restParts.length > 0) {
    const subRef = parseRef(restParts.join('/'));
    subRefs.push(subRef);
  }

  return { type: type ?? '', index, subRefs };
}

/**
 * Get the leaf ref (innermost component).
 */
function getLeafRef(parsed: ParsedRef): ParsedRef {
  if (parsed.subRefs.length === 0) {
    return parsed;
  }
  return getLeafRef(parsed.subRefs[0]!);
}

// =============================================================================
// Element Resolution
// =============================================================================

/**
 * Resolve a ref to a Word paragraph.
 */
async function resolveParagraphRef(
  context: WordRequestContext,
  ref: Ref
): Promise<WordParagraph> {
  const parsed = parseRef(ref);

  if (parsed.type === 'p') {
    // Top-level paragraph
    const paragraphs = context.document.body.paragraphs.load('items');
    await context.sync();

    if (parsed.index >= paragraphs.items.length) {
      throw new Error(`Ref not found: ${ref} (paragraph index ${parsed.index} out of range)`);
    }

    const paragraph = paragraphs.items[parsed.index];
    if (!paragraph) {
      throw new Error(`Ref not found: ${ref}`);
    }
    return paragraph;
  }

  if (parsed.type === 'tbl') {
    // Paragraph inside a table cell
    const tables = context.document.body.tables.load('items');
    await context.sync();

    if (parsed.index >= tables.items.length) {
      throw new Error(`Ref not found: ${ref} (table index ${parsed.index} out of range)`);
    }

    const table = tables.items[parsed.index];
    if (!table) {
      throw new Error(`Ref not found: ${ref}`);
    }

    // Navigate to row
    const rowRef = parsed.subRefs[0];
    if (!rowRef || rowRef.type !== 'row') {
      throw new Error(`Invalid table ref: ${ref}`);
    }

    table.rows.load('items');
    await context.sync();

    if (rowRef.index >= table.rows.items.length) {
      throw new Error(`Ref not found: ${ref} (row index ${rowRef.index} out of range)`);
    }

    const row = table.rows.items[rowRef.index];
    if (!row) {
      throw new Error(`Ref not found: ${ref}`);
    }

    // Navigate to cell
    const cellRef = rowRef.subRefs[0];
    if (!cellRef || cellRef.type !== 'cell') {
      throw new Error(`Invalid cell ref: ${ref}`);
    }

    row.cells.load('items');
    await context.sync();

    if (cellRef.index >= row.cells.items.length) {
      throw new Error(`Ref not found: ${ref} (cell index ${cellRef.index} out of range)`);
    }

    const cell = row.cells.items[cellRef.index];
    if (!cell) {
      throw new Error(`Ref not found: ${ref}`);
    }

    // Navigate to paragraph in cell
    const pRef = cellRef.subRefs[0];
    if (!pRef || pRef.type !== 'p') {
      throw new Error(`Invalid paragraph ref in cell: ${ref}`);
    }

    cell.body.paragraphs.load('items');
    await context.sync();

    if (pRef.index >= cell.body.paragraphs.items.length) {
      throw new Error(`Ref not found: ${ref} (paragraph index ${pRef.index} out of range)`);
    }

    const paragraph = cell.body.paragraphs.items[pRef.index];
    if (!paragraph) {
      throw new Error(`Ref not found: ${ref}`);
    }
    return paragraph;
  }

  throw new Error(`Unsupported ref type: ${parsed.type} in ref ${ref}`);
}

/**
 * Resolve a ref to a Word table.
 */
async function resolveTableRef(
  context: WordRequestContext,
  ref: Ref
): Promise<WordTable> {
  const parsed = parseRef(ref);

  if (parsed.type !== 'tbl') {
    throw new Error(`Expected table ref, got: ${ref}`);
  }

  const tables = context.document.body.tables.load('items');
  await context.sync();

  if (parsed.index >= tables.items.length) {
    throw new Error(`Ref not found: ${ref} (table index ${parsed.index} out of range)`);
  }

  const table = tables.items[parsed.index];
  if (!table) {
    throw new Error(`Ref not found: ${ref}`);
  }
  return table;
}

/**
 * Resolve a ref to a Word table row.
 */
async function resolveRowRef(
  context: WordRequestContext,
  ref: Ref
): Promise<WordRow> {
  const parsed = parseRef(ref);

  if (parsed.type !== 'tbl') {
    throw new Error(`Expected table row ref, got: ${ref}`);
  }

  const table = await resolveTableRef(context, `tbl:${parsed.index}`);

  const rowRef = parsed.subRefs[0];
  if (!rowRef || rowRef.type !== 'row') {
    throw new Error(`Expected row ref, got: ${ref}`);
  }

  table.rows.load('items');
  await context.sync();

  if (rowRef.index >= table.rows.items.length) {
    throw new Error(`Ref not found: ${ref} (row index ${rowRef.index} out of range)`);
  }

  const row = table.rows.items[rowRef.index];
  if (!row) {
    throw new Error(`Ref not found: ${ref}`);
  }
  return row;
}

// =============================================================================
// Editing Methods
// =============================================================================

/**
 * Replace the text content at a ref.
 *
 * @param context - Word.RequestContext from Office.js
 * @param ref - Reference to the element to replace
 * @param newText - New text to replace with
 * @param options - Editing options (track changes, author, etc.)
 * @returns EditResult indicating success/failure
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await replaceByRef(context, "p:3", "Updated text", { track: true });
 *   if (result.success) {
 *     console.log("Replaced:", result.newRef);
 *   }
 * });
 * ```
 */
export async function replaceByRef(
  context: WordRequestContext,
  ref: Ref,
  newText: string,
  options?: EditOptions
): Promise<EditResult> {
  try {
    const parsed = parseRef(ref);

    // Handle paragraph replacement
    if (parsed.type === 'p' || (parsed.type === 'tbl' && getLeafRef(parsed).type === 'p')) {
      const paragraph = await resolveParagraphRef(context, ref);

      // Get the current range
      const range = paragraph.getRange('Content');
      range.load('text');
      await context.sync();

      // Enable tracked changes if requested
      if (options?.track && context.trackedChanges) {
        context.trackedChanges.enabled = true;
      }

      // Replace the text
      range.insertText(newText, 'Replace');
      await context.sync();

      return {
        success: true,
        newRef: ref,
      };
    }

    return {
      success: false,
      error: `Unsupported ref type for replacement: ${parsed.type}`,
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * Insert content after a ref.
 *
 * @param context - Word.RequestContext from Office.js
 * @param ref - Reference to the element to insert after
 * @param content - Content to insert
 * @param options - Editing options (track changes, author, etc.)
 * @returns EditResult indicating success/failure
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await insertAfterRef(context, "p:5", " (amended)", { track: true });
 * });
 * ```
 */
export async function insertAfterRef(
  context: WordRequestContext,
  ref: Ref,
  content: string,
  options?: EditOptions
): Promise<EditResult> {
  try {
    const parsed = parseRef(ref);

    if (parsed.type === 'p' || (parsed.type === 'tbl' && getLeafRef(parsed).type === 'p')) {
      const paragraph = await resolveParagraphRef(context, ref);

      // Enable tracked changes if requested
      if (options?.track && context.trackedChanges) {
        context.trackedChanges.enabled = true;
      }

      // Insert text at end of paragraph
      const range = paragraph.getRange('End');
      range.insertText(content, 'After');
      await context.sync();

      return {
        success: true,
        newRef: ref,
      };
    }

    return {
      success: false,
      error: `Unsupported ref type for insertion: ${parsed.type}`,
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * Insert content before a ref.
 *
 * @param context - Word.RequestContext from Office.js
 * @param ref - Reference to the element to insert before
 * @param content - Content to insert
 * @param options - Editing options (track changes, author, etc.)
 * @returns EditResult indicating success/failure
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await insertBeforeRef(context, "p:5", "Note: ", { track: true });
 * });
 * ```
 */
export async function insertBeforeRef(
  context: WordRequestContext,
  ref: Ref,
  content: string,
  options?: EditOptions
): Promise<EditResult> {
  try {
    const parsed = parseRef(ref);

    if (parsed.type === 'p' || (parsed.type === 'tbl' && getLeafRef(parsed).type === 'p')) {
      const paragraph = await resolveParagraphRef(context, ref);

      // Enable tracked changes if requested
      if (options?.track && context.trackedChanges) {
        context.trackedChanges.enabled = true;
      }

      // Insert text at start of paragraph
      const range = paragraph.getRange('Start');
      range.insertText(content, 'Before');
      await context.sync();

      return {
        success: true,
        newRef: ref,
      };
    }

    return {
      success: false,
      error: `Unsupported ref type for insertion: ${parsed.type}`,
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * Delete an element by ref.
 *
 * @param context - Word.RequestContext from Office.js
 * @param ref - Reference to the element to delete
 * @param options - Editing options (track changes, author, etc.)
 * @returns EditResult indicating success/failure
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await deleteByRef(context, "p:3", { track: true });
 * });
 * ```
 */
export async function deleteByRef(
  context: WordRequestContext,
  ref: Ref,
  options?: EditOptions
): Promise<EditResult> {
  try {
    const parsed = parseRef(ref);

    // Enable tracked changes if requested
    if (options?.track && context.trackedChanges) {
      context.trackedChanges.enabled = true;
    }

    // Handle paragraph deletion
    if (parsed.type === 'p' || (parsed.type === 'tbl' && getLeafRef(parsed).type === 'p')) {
      const paragraph = await resolveParagraphRef(context, ref);
      paragraph.delete();
      await context.sync();

      return {
        success: true,
      };
    }

    // Handle table deletion
    if (parsed.type === 'tbl' && parsed.subRefs.length === 0) {
      const table = await resolveTableRef(context, ref);
      table.delete();
      await context.sync();

      return {
        success: true,
      };
    }

    // Handle row deletion
    if (parsed.type === 'tbl' && parsed.subRefs[0]?.type === 'row' && parsed.subRefs[0].subRefs.length === 0) {
      const row = await resolveRowRef(context, ref);
      row.delete();
      await context.sync();

      return {
        success: true,
      };
    }

    return {
      success: false,
      error: `Unsupported ref type for deletion: ${ref}`,
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * Apply formatting to an element by ref.
 *
 * @param context - Word.RequestContext from Office.js
 * @param ref - Reference to the element to format
 * @param formatting - Formatting options to apply
 * @returns EditResult indicating success/failure
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await formatByRef(context, "p:3", {
 *     bold: true,
 *     color: "#0000FF"
 *   });
 * });
 * ```
 */
export async function formatByRef(
  context: WordRequestContext,
  ref: Ref,
  formatting: FormatOptions
): Promise<EditResult> {
  try {
    const parsed = parseRef(ref);

    if (parsed.type === 'p' || (parsed.type === 'tbl' && getLeafRef(parsed).type === 'p')) {
      const paragraph = await resolveParagraphRef(context, ref);
      paragraph.load('font,style');
      await context.sync();

      // Apply character formatting
      const fontProps: Partial<WordFont> = {};
      if (formatting.bold !== undefined) fontProps.bold = formatting.bold;
      if (formatting.italic !== undefined) fontProps.italic = formatting.italic;
      if (formatting.underline !== undefined) fontProps.underline = formatting.underline ? 'Single' : 'None';
      if (formatting.strikethrough !== undefined) fontProps.strikeThrough = formatting.strikethrough;
      if (formatting.font !== undefined) fontProps.name = formatting.font;
      if (formatting.size !== undefined) fontProps.size = formatting.size;
      if (formatting.color !== undefined) fontProps.color = formatting.color;
      if (formatting.highlight !== undefined) fontProps.highlightColor = formatting.highlight;

      if (Object.keys(fontProps).length > 0) {
        paragraph.font.set(fontProps);
      }

      // Apply paragraph style
      if (formatting.style !== undefined) {
        paragraph.style = formatting.style;
      }

      await context.sync();

      return {
        success: true,
        newRef: ref,
      };
    }

    return {
      success: false,
      error: `Unsupported ref type for formatting: ${parsed.type}`,
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

// =============================================================================
// Convenience Methods
// =============================================================================

/**
 * Insert a new paragraph after a ref.
 *
 * @param context - Word.RequestContext from Office.js
 * @param ref - Reference to insert after
 * @param paragraphText - Text for the new paragraph
 * @param options - Editing options
 * @returns EditResult with new paragraph ref
 */
export async function insertParagraphAfterRef(
  context: WordRequestContext,
  ref: Ref,
  paragraphText: string,
  options?: EditOptions
): Promise<EditResult> {
  try {
    const parsed = parseRef(ref);

    if (parsed.type === 'p') {
      const paragraph = await resolveParagraphRef(context, ref);

      // Enable tracked changes if requested
      if (options?.track && context.trackedChanges) {
        context.trackedChanges.enabled = true;
      }

      // Insert new paragraph after
      paragraph.insertParagraph(paragraphText, 'After');
      await context.sync();

      // New paragraph gets the next index
      return {
        success: true,
        newRef: `p:${parsed.index + 1}`,
      };
    }

    return {
      success: false,
      error: `Unsupported ref type for paragraph insertion: ${parsed.type}`,
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * Insert a new paragraph before a ref.
 *
 * @param context - Word.RequestContext from Office.js
 * @param ref - Reference to insert before
 * @param paragraphText - Text for the new paragraph
 * @param options - Editing options
 * @returns EditResult with new paragraph ref
 */
export async function insertParagraphBeforeRef(
  context: WordRequestContext,
  ref: Ref,
  paragraphText: string,
  options?: EditOptions
): Promise<EditResult> {
  try {
    const parsed = parseRef(ref);

    if (parsed.type === 'p') {
      const paragraph = await resolveParagraphRef(context, ref);

      // Enable tracked changes if requested
      if (options?.track && context.trackedChanges) {
        context.trackedChanges.enabled = true;
      }

      // Insert new paragraph before
      paragraph.insertParagraph(paragraphText, 'Before');
      await context.sync();

      // New paragraph takes the current index, existing shifts
      return {
        success: true,
        newRef: ref,
      };
    }

    return {
      success: false,
      error: `Unsupported ref type for paragraph insertion: ${parsed.type}`,
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * Get the text content at a ref.
 *
 * @param context - Word.RequestContext from Office.js
 * @param ref - Reference to read from
 * @returns The text content or undefined if not found
 */
export async function getTextByRef(
  context: WordRequestContext,
  ref: Ref
): Promise<string | undefined> {
  try {
    const parsed = parseRef(ref);

    // Handle footnote refs (fn:1, fn:2, etc.)
    if (parsed.type === 'fn') {
      return await getFootnoteText(context, parsed.index);
    }

    if (parsed.type === 'p' || (parsed.type === 'tbl' && getLeafRef(parsed).type === 'p')) {
      const paragraph = await resolveParagraphRef(context, ref);
      paragraph.load('text');
      await context.sync();
      return paragraph.text;
    }

    return undefined;
  } catch {
    return undefined;
  }
}

/**
 * Get text content of a footnote by ID.
 *
 * @param context - Word.RequestContext from Office.js
 * @param footnoteId - Footnote ID (1-indexed)
 * @returns The footnote text or undefined if not found
 */
async function getFootnoteText(
  context: WordRequestContext,
  footnoteId: number
): Promise<string | undefined> {
  try {
    // Get OOXML which contains footnotes
    const body = context.document.body as unknown as { getOoxml(): { value: string } };
    const ooxml = body.getOoxml();
    await context.sync();

    const xml = ooxml.value;

    // Find the footnotes section
    const footnotesMatch = xml.match(/<w:footnotes[^>]*>([\s\S]*?)<\/w:footnotes>/);
    if (!footnotesMatch) return undefined;

    // Find the specific footnote
    const fnPattern = new RegExp(
      `<w:footnote[^>]*w:id="${footnoteId}"[^>]*>([\\s\\S]*?)<\\/w:footnote>`
    );
    const fnContent = footnotesMatch[1];
    if (!fnContent) return undefined;
    const fnMatch = fnContent.match(fnPattern);
    if (!fnMatch?.[1]) return undefined;

    // Extract text from the footnote
    const textParts = fnMatch[1].match(/<w:t[^>]*>([^<]*)<\/w:t>/g) ?? [];
    const text = textParts
      .map((t) => t.replace(/<[^>]+>/g, ''))
      .join('')
      .trim();

    return text || undefined;
  } catch {
    return undefined;
  }
}

// =============================================================================
// Scope-Aware Editing Methods
// =============================================================================

/**
 * Replace text in all nodes matching a scope.
 *
 * PERFORMANCE: Uses batched paragraph loading and operations.
 * Reduces O(n) sync calls to O(1) by loading all paragraphs once
 * and queuing all operations before a single sync.
 *
 * @param context - Word.RequestContext from Office.js
 * @param tree - Accessibility tree to search in
 * @param scope - Scope specification to match nodes
 * @param newText - New text to replace with
 * @param options - Editing options (track changes, author, etc.)
 * @returns Array of EditResults for each matched node
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const tree = await buildTree(context);
 *   const results = await replaceByScope(
 *     context,
 *     tree,
 *     "section:Methods",
 *     "Updated methods text",
 *     { track: true }
 *   );
 * });
 * ```
 */
export async function replaceByScope(
  context: WordRequestContext,
  tree: AccessibilityTree,
  scope: ScopeSpec,
  newText: string,
  options?: EditOptions
): Promise<EditResult[]> {
  const scopeResult = resolveScope(tree, scope);
  const results: EditResult[] = [];

  if (scopeResult.nodes.length === 0) {
    return results;
  }

  // BATCHED: Load all paragraphs once
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  // Enable tracked changes if requested
  if (options?.track && context.trackedChanges) {
    context.trackedChanges.enabled = true;
  }

  // BATCHED: Queue all range loads
  const ranges: Array<{ ref: Ref; range: WordRange }> = [];
  for (const node of scopeResult.nodes) {
    try {
      const parsed = parseRef(node.ref);
      if (parsed.type === 'p' && parsed.index < paragraphs.items.length) {
        const para = paragraphs.items[parsed.index];
        if (para) {
          const range = para.getRange('Content');
          ranges.push({ ref: node.ref, range });
        }
      }
    } catch {
      results.push({ success: false, error: `Invalid ref: ${node.ref}` });
    }
  }

  // Single sync to load all ranges
  if (ranges.length > 0) {
    await context.sync();
  }

  // BATCHED: Queue all replacements
  for (const { ref, range } of ranges) {
    try {
      range.insertText(newText, 'Replace');
      results.push({ success: true, newRef: ref });
    } catch (err) {
      results.push({ success: false, error: err instanceof Error ? err.message : String(err) });
    }
  }

  // Single sync for all operations
  if (ranges.length > 0) {
    await context.sync();
  }

  return results;
}

/**
 * Delete all nodes matching a scope.
 *
 * PERFORMANCE: Uses batched paragraph loading and deletion.
 * Reduces O(n) sync calls to O(1) by loading all paragraphs once
 * and queuing all deletions before a single sync.
 *
 * @param context - Word.RequestContext from Office.js
 * @param tree - Accessibility tree to search in
 * @param scope - Scope specification to match nodes
 * @param options - Editing options (track changes, author, etc.)
 * @returns Array of EditResults for each deleted node
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const tree = await buildTree(context);
 *   // Delete all paragraphs containing "DRAFT"
 *   const results = await deleteByScope(context, tree, "DRAFT", { track: true });
 * });
 * ```
 */
export async function deleteByScope(
  context: WordRequestContext,
  tree: AccessibilityTree,
  scope: ScopeSpec,
  options?: EditOptions
): Promise<EditResult[]> {
  const scopeResult = resolveScope(tree, scope);
  const results: EditResult[] = [];

  if (scopeResult.nodes.length === 0) {
    return results;
  }

  // Delete in reverse order to preserve indices
  const sortedNodes = [...scopeResult.nodes].sort((a, b) => {
    // Extract index from ref (e.g., "p:5" -> 5)
    const aMatch = a.ref.match(/:(\d+)/);
    const bMatch = b.ref.match(/:(\d+)/);
    const aIndex = aMatch?.[1] ? parseInt(aMatch[1], 10) : 0;
    const bIndex = bMatch?.[1] ? parseInt(bMatch[1], 10) : 0;
    return bIndex - aIndex; // Reverse order
  });

  // BATCHED: Load all paragraphs once
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  // Enable tracked changes if requested
  if (options?.track && context.trackedChanges) {
    context.trackedChanges.enabled = true;
  }

  // BATCHED: Queue all deletions (no sync needed between)
  for (const node of sortedNodes) {
    try {
      const parsed = parseRef(node.ref);
      if (parsed.type === 'p' && parsed.index < paragraphs.items.length) {
        const para = paragraphs.items[parsed.index];
        if (para) {
          para.delete();
          results.push({ success: true });
        }
      }
    } catch (err) {
      results.push({ success: false, error: err instanceof Error ? err.message : String(err) });
    }
  }

  // Single sync for all deletions
  await context.sync();

  return results;
}

/**
 * Apply formatting to all nodes matching a scope.
 *
 * PERFORMANCE: Uses batched paragraph loading and formatting.
 * Reduces O(n) sync calls to O(1) by loading all paragraphs once
 * and queuing all formatting operations before a single sync.
 *
 * @param context - Word.RequestContext from Office.js
 * @param tree - Accessibility tree to search in
 * @param scope - Scope specification to match nodes
 * @param formatting - Formatting options to apply
 * @returns Array of EditResults for each formatted node
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const tree = await buildTree(context);
 *   // Make all headings in "Results" section bold and blue
 *   const results = await formatByScope(
 *     context,
 *     tree,
 *     { section: "Results", role: "heading" },
 *     { bold: true, color: "#0000FF" }
 *   );
 * });
 * ```
 */
export async function formatByScope(
  context: WordRequestContext,
  tree: AccessibilityTree,
  scope: ScopeSpec,
  formatting: FormatOptions
): Promise<EditResult[]> {
  const scopeResult = resolveScope(tree, scope);
  const results: EditResult[] = [];

  if (scopeResult.nodes.length === 0) {
    return results;
  }

  // BATCHED: Load all paragraphs once with font properties
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  // Build font properties object once
  const fontProps: Partial<WordFont> = {};
  if (formatting.bold !== undefined) fontProps.bold = formatting.bold;
  if (formatting.italic !== undefined) fontProps.italic = formatting.italic;
  if (formatting.underline !== undefined) fontProps.underline = formatting.underline ? 'Single' : 'None';
  if (formatting.strikethrough !== undefined) fontProps.strikeThrough = formatting.strikethrough;
  if (formatting.font !== undefined) fontProps.name = formatting.font;
  if (formatting.size !== undefined) fontProps.size = formatting.size;
  if (formatting.color !== undefined) fontProps.color = formatting.color;
  if (formatting.highlight !== undefined) fontProps.highlightColor = formatting.highlight;
  const hasFontProps = Object.keys(fontProps).length > 0;

  // BATCHED: Queue all formatting operations (no sync between)
  for (const node of scopeResult.nodes) {
    try {
      const parsed = parseRef(node.ref);
      if (parsed.type === 'p' && parsed.index < paragraphs.items.length) {
        const para = paragraphs.items[parsed.index];
        if (para) {
          if (hasFontProps) {
            para.font.set(fontProps);
          }
          if (formatting.style !== undefined) {
            para.style = formatting.style;
          }
          results.push({ success: true, newRef: node.ref });
        }
      }
    } catch (err) {
      results.push({ success: false, error: err instanceof Error ? err.message : String(err) });
    }
  }

  // Single sync for all formatting
  await context.sync();

  return results;
}

/**
 * Search and replace text within all nodes matching a scope.
 *
 * PERFORMANCE: Uses batched paragraph loading and replacement.
 * Reduces O(n) sync calls to O(1) by loading all paragraphs once
 * and queuing all replacements before a single sync.
 *
 * @param context - Word.RequestContext from Office.js
 * @param tree - Accessibility tree to search in
 * @param scope - Scope specification to match nodes
 * @param searchText - Text to find within matching nodes
 * @param replaceText - Text to replace with
 * @param options - Editing options (track changes, etc.)
 * @returns Array of EditResults for each modified node
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const tree = await buildTree(context);
 *   // Replace "Plaintiff" with "Defendant" only in section "Parties"
 *   const results = await searchReplaceByScope(
 *     context,
 *     tree,
 *     "section:Parties",
 *     "Plaintiff",
 *     "Defendant",
 *     { track: true }
 *   );
 * });
 * ```
 */
export async function searchReplaceByScope(
  context: WordRequestContext,
  tree: AccessibilityTree,
  scope: ScopeSpec,
  searchText: string,
  replaceText: string,
  options?: EditOptions
): Promise<EditResult[]> {
  const scopeResult = resolveScope(tree, scope);
  const results: EditResult[] = [];

  // Filter to only nodes containing the search text
  const nodesToReplace = scopeResult.nodes.filter(
    (node) => node.text && node.text.includes(searchText)
  );

  if (nodesToReplace.length === 0) {
    return results;
  }

  // BATCHED: Load all paragraphs once
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  // Enable tracked changes if requested
  if (options?.track && context.trackedChanges) {
    context.trackedChanges.enabled = true;
  }

  // BATCHED: Queue all range loads
  const ranges: Array<{ ref: Ref; range: WordRange; newText: string }> = [];
  const searchRegex = new RegExp(escapeRegExp(searchText), 'g');

  for (const node of nodesToReplace) {
    try {
      const parsed = parseRef(node.ref);
      if (parsed.type === 'p' && parsed.index < paragraphs.items.length) {
        const para = paragraphs.items[parsed.index];
        if (para && node.text) {
          const range = para.getRange('Content');
          const newText = node.text.replace(searchRegex, replaceText);
          ranges.push({ ref: node.ref, range, newText });
        }
      }
    } catch {
      results.push({ success: false, error: `Invalid ref: ${node.ref}` });
    }
  }

  // Single sync to load all ranges
  if (ranges.length > 0) {
    await context.sync();
  }

  // BATCHED: Queue all replacements
  for (const { ref, range, newText } of ranges) {
    try {
      range.insertText(newText, 'Replace');
      results.push({ success: true, newRef: ref });
    } catch (err) {
      results.push({ success: false, error: err instanceof Error ? err.message : String(err) });
    }
  }

  // Single sync for all operations
  if (ranges.length > 0) {
    await context.sync();
  }

  return results;
}

/**
 * Escape special regex characters in a string.
 */
function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// =============================================================================
// Batch Operations
// =============================================================================

/**
 * A single edit operation for batch processing.
 */
export interface BatchEditOperation {
  /** Reference to the element to edit */
  ref: Ref;
  /** New text content (for replace operations) */
  newText?: string;
  /** Operation type */
  operation: 'replace' | 'delete' | 'insertAfter' | 'insertBefore';
  /** Text to insert (for insert operations) */
  insertText?: string;
}

/**
 * Result of a batch edit operation.
 */
export interface BatchEditResult {
  /** Overall success (true if all operations succeeded) */
  success: boolean;
  /** Number of successful operations */
  successCount: number;
  /** Number of failed operations */
  failedCount: number;
  /** Individual results for each operation */
  results: EditResult[];
}

/**
 * Execute multiple edit operations in a single batched transaction.
 *
 * PERFORMANCE: Uses batched paragraph loading and operations.
 * All operations are queued and executed with minimal context.sync() calls,
 * regardless of how many edits are requested.
 *
 * @param context - Word.RequestContext from Office.js
 * @param operations - Array of edit operations to perform
 * @param options - Editing options (track changes, author, etc.)
 * @returns BatchEditResult with overall success and individual results
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await batchEdit(context, [
 *     { ref: "p:3", operation: "replace", newText: "Updated intro" },
 *     { ref: "p:7", operation: "replace", newText: "New conclusion" },
 *     { ref: "p:12", operation: "delete" },
 *     { ref: "p:5", operation: "insertAfter", insertText: " (amended)" },
 *   ], { track: true });
 *
 *   console.log(`${result.successCount}/${result.results.length} operations succeeded`);
 * });
 * ```
 */
export async function batchEdit(
  context: WordRequestContext,
  operations: BatchEditOperation[],
  options?: EditOptions
): Promise<BatchEditResult> {
  const results: EditResult[] = [];

  if (operations.length === 0) {
    return { success: true, successCount: 0, failedCount: 0, results };
  }

  // BATCHED: Load all paragraphs once
  const paragraphs = context.document.body.paragraphs.load('items');
  await context.sync();

  // Enable tracked changes if requested
  if (options?.track && context.trackedChanges) {
    context.trackedChanges.enabled = true;
  }

  // Sort operations: deletions should be processed last (in reverse order)
  // to avoid index shifting issues
  const sortedOps = [...operations].sort((a, b) => {
    // Deletions go last
    if (a.operation === 'delete' && b.operation !== 'delete') return 1;
    if (a.operation !== 'delete' && b.operation === 'delete') return -1;

    // Among deletions, process in reverse index order
    if (a.operation === 'delete' && b.operation === 'delete') {
      const aMatch = a.ref.match(/:(\d+)/);
      const bMatch = b.ref.match(/:(\d+)/);
      const aIndex = aMatch?.[1] ? parseInt(aMatch[1], 10) : 0;
      const bIndex = bMatch?.[1] ? parseInt(bMatch[1], 10) : 0;
      return bIndex - aIndex;
    }

    return 0;
  });

  // BATCHED: Queue all range loads for replace/insert operations
  const rangeOps: Array<{
    op: BatchEditOperation;
    range: WordRange;
    para: WordParagraph;
  }> = [];

  for (const op of sortedOps) {
    try {
      const parsed = parseRef(op.ref);
      if (parsed.type !== 'p' || parsed.index >= paragraphs.items.length) {
        results.push({ success: false, error: `Invalid ref: ${op.ref}` });
        continue;
      }

      const para = paragraphs.items[parsed.index];
      if (!para) {
        results.push({ success: false, error: `Paragraph not found: ${op.ref}` });
        continue;
      }

      if (op.operation === 'replace') {
        const range = para.getRange('Content');
        rangeOps.push({ op, range, para });
      } else if (op.operation === 'insertAfter') {
        const range = para.getRange('End');
        rangeOps.push({ op, range, para });
      } else if (op.operation === 'insertBefore') {
        const range = para.getRange('Start');
        rangeOps.push({ op, range, para });
      } else if (op.operation === 'delete') {
        // Deletions don't need range loading, handled separately
        rangeOps.push({ op, range: null as unknown as WordRange, para });
      }
    } catch (err) {
      results.push({
        success: false,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  // Single sync to load all ranges
  if (rangeOps.length > 0) {
    await context.sync();
  }

  // BATCHED: Execute all operations
  for (const { op, range, para } of rangeOps) {
    try {
      switch (op.operation) {
        case 'replace':
          if (op.newText !== undefined) {
            range.insertText(op.newText, 'Replace');
            results.push({ success: true, newRef: op.ref });
          } else {
            results.push({ success: false, error: 'newText required for replace' });
          }
          break;

        case 'insertAfter':
          if (op.insertText !== undefined) {
            range.insertText(op.insertText, 'After');
            results.push({ success: true, newRef: op.ref });
          } else {
            results.push({ success: false, error: 'insertText required for insertAfter' });
          }
          break;

        case 'insertBefore':
          if (op.insertText !== undefined) {
            range.insertText(op.insertText, 'Before');
            results.push({ success: true, newRef: op.ref });
          } else {
            results.push({ success: false, error: 'insertText required for insertBefore' });
          }
          break;

        case 'delete':
          para.delete();
          results.push({ success: true });
          break;
      }
    } catch (err) {
      results.push({
        success: false,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  // Single sync for all operations
  await context.sync();

  const successCount = results.filter((r) => r.success).length;
  const failedCount = results.filter((r) => !r.success).length;

  return {
    success: failedCount === 0,
    successCount,
    failedCount,
    results,
  };
}

/**
 * Convenience function to batch replace multiple refs with different texts.
 *
 * @param context - Word.RequestContext from Office.js
 * @param replacements - Map or array of ref -> newText pairs
 * @param options - Editing options (track changes, author, etc.)
 * @returns BatchEditResult
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await batchReplace(context, [
 *     { ref: "p:3", text: "Updated intro" },
 *     { ref: "p:7", text: "New conclusion" },
 *   ], { track: true });
 * });
 * ```
 */
export async function batchReplace(
  context: WordRequestContext,
  replacements: Array<{ ref: Ref; text: string }>,
  options?: EditOptions
): Promise<BatchEditResult> {
  const operations: BatchEditOperation[] = replacements.map((r) => ({
    ref: r.ref,
    operation: 'replace' as const,
    newText: r.text,
  }));

  return batchEdit(context, operations, options);
}

/**
 * Convenience function to batch delete multiple refs.
 *
 * @param context - Word.RequestContext from Office.js
 * @param refs - Array of refs to delete
 * @param options - Editing options (track changes, author, etc.)
 * @returns BatchEditResult
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await batchDelete(context, ["p:3", "p:7", "p:12"], { track: true });
 * });
 * ```
 */
export async function batchDelete(
  context: WordRequestContext,
  refs: Ref[],
  options?: EditOptions
): Promise<BatchEditResult> {
  const operations: BatchEditOperation[] = refs.map((ref) => ({
    ref,
    operation: 'delete' as const,
  }));

  return batchEdit(context, operations, options);
}

// =============================================================================
// Text Search
// =============================================================================

/**
 * A single match found by findText.
 */
export interface TextMatch {
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

/**
 * Result of a findText operation.
 */
export interface FindTextResult {
  /** Total number of matches found */
  count: number;
  /** All matches found */
  matches: TextMatch[];
  /** Number of paragraphs searched */
  paragraphsSearched: number;
}

/**
 * Options for findText operation.
 */
export interface FindTextOptions {
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

/**
 * Search for text across the document, returning refs and match positions.
 *
 * PERFORMANCE: Uses batched paragraph loading. All paragraphs are loaded
 * in a single sync call, then searched synchronously.
 *
 * @param context - Word.RequestContext from Office.js
 * @param searchText - Text or regex pattern to search for
 * @param tree - Optional accessibility tree (required if using scope option)
 * @param options - Search options (case sensitivity, regex, scope, etc.)
 * @returns FindTextResult with all matches
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   // Simple search
 *   const result = await findText(context, "agreement");
 *   console.log(`Found ${result.count} matches`);
 *
 *   // Case-insensitive search
 *   const result2 = await findText(context, "WHEREAS", null, { caseInsensitive: true });
 *
 *   // Regex search
 *   const result3 = await findText(context, "\\$[\\d,]+\\.\\d{2}", null, { regex: true });
 *
 *   // Scoped search
 *   const tree = await buildTree(context);
 *   const result4 = await findText(context, "plaintiff", tree, { scope: "section:Parties" });
 * });
 * ```
 */
export async function findText(
  context: WordRequestContext,
  searchText: string,
  tree?: AccessibilityTree | null,
  options?: FindTextOptions
): Promise<FindTextResult> {
  const matches: TextMatch[] = [];
  const maxMatches = options?.maxMatches ?? Infinity;

  // If scope is specified, use the tree to filter paragraphs
  let scopedRefs: Set<Ref> | null = null;
  if (options?.scope && tree) {
    const scopeResult = resolveScope(tree, options.scope);
    scopedRefs = new Set(scopeResult.nodes.map((n) => n.ref));
  }

  // BATCHED: Load all paragraphs once
  const paragraphs = context.document.body.paragraphs.load('items,text');
  await context.sync();

  const paragraphsSearched = paragraphs.items.length;

  // Build the search pattern
  let pattern: RegExp;
  if (options?.regex) {
    try {
      const flags = options.caseInsensitive ? 'gi' : 'g';
      pattern = new RegExp(searchText, flags);
    } catch (e) {
      // Invalid regex
      return { count: 0, matches: [], paragraphsSearched };
    }
  } else {
    // Escape special regex characters for literal search
    let escapedSearch = escapeRegExp(searchText);

    // Add word boundary for whole word matching
    if (options?.wholeWord) {
      escapedSearch = `\\b${escapedSearch}\\b`;
    }

    const flags = options?.caseInsensitive ? 'gi' : 'g';
    pattern = new RegExp(escapedSearch, flags);
  }

  // Search through paragraphs
  for (let i = 0; i < paragraphs.items.length && matches.length < maxMatches; i++) {
    const para = paragraphs.items[i];
    if (!para) continue;

    const ref = `p:${i}` as Ref;

    // Skip if not in scope
    if (scopedRefs && !scopedRefs.has(ref)) {
      continue;
    }

    const text = para.text;
    if (!text) continue;

    // Find all matches in this paragraph
    let match: RegExpExecArray | null;
    pattern.lastIndex = 0; // Reset for each paragraph

    while ((match = pattern.exec(text)) !== null && matches.length < maxMatches) {
      matches.push({
        ref,
        text,
        start: match.index,
        end: match.index + match[0].length,
        matchedText: match[0],
      });

      // Prevent infinite loop on zero-length matches
      if (match[0].length === 0) {
        pattern.lastIndex++;
      }
    }
  }

  return {
    count: matches.length,
    matches,
    paragraphsSearched,
  };
}

/**
 * Find and highlight text in the document.
 *
 * Searches for text and applies highlight formatting to all matches.
 *
 * @param context - Word.RequestContext from Office.js
 * @param searchText - Text to search for
 * @param highlightColor - Highlight color (e.g., "yellow", "green", "#FFFF00")
 * @param options - Search options
 * @returns Number of matches highlighted
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const count = await findAndHighlight(context, "important", "yellow");
 *   console.log(`Highlighted ${count} occurrences`);
 * });
 * ```
 */
export async function findAndHighlight(
  context: WordRequestContext,
  searchText: string,
  highlightColor: string,
  options?: FindTextOptions
): Promise<number> {
  // BATCHED: Load all paragraphs once
  const paragraphs = context.document.body.paragraphs.load('items,text');
  await context.sync();

  // Build the search pattern
  let pattern: RegExp;
  if (options?.regex) {
    try {
      const flags = options.caseInsensitive ? 'gi' : 'g';
      pattern = new RegExp(searchText, flags);
    } catch {
      return 0;
    }
  } else {
    let escapedSearch = escapeRegExp(searchText);
    if (options?.wholeWord) {
      escapedSearch = `\\b${escapedSearch}\\b`;
    }
    const flags = options?.caseInsensitive ? 'gi' : 'g';
    pattern = new RegExp(escapedSearch, flags);
  }

  // Collect all ranges to highlight
  interface RangeToHighlight {
    range: WordRange;
    para: WordParagraph;
    start: number;
    length: number;
  }
  const rangesToHighlight: RangeToHighlight[] = [];

  // Find matches and queue range loads
  for (let i = 0; i < paragraphs.items.length; i++) {
    const para = paragraphs.items[i];
    if (!para) continue;

    const text = para.text;
    if (!text) continue;

    pattern.lastIndex = 0;
    let match: RegExpExecArray | null;

    while ((match = pattern.exec(text)) !== null) {
      // Get range for this specific text portion
      // Note: We'll need to use search to find the exact range
      rangesToHighlight.push({
        range: null as unknown as WordRange,
        para,
        start: match.index,
        length: match[0].length,
      });

      if (match[0].length === 0) {
        pattern.lastIndex++;
      }
    }
  }

  if (rangesToHighlight.length === 0) {
    return 0;
  }

  // Use Word's search feature to find and highlight
  // This is more reliable than trying to calculate character positions
  const body = context.document.body;
  const searchResults = body.search(searchText, {
    matchCase: !options?.caseInsensitive,
    matchWholeWord: options?.wholeWord ?? false,
  });
  searchResults.load('items');
  await context.sync();

  // Apply highlight to all found ranges
  for (const item of searchResults.items) {
    (item as WordRange).font.highlightColor = highlightColor;
  }

  await context.sync();

  return searchResults.items.length;
}

// =============================================================================
// Tracked Changes Operations
// =============================================================================

/**
 * Office.js TrackedChange interface
 */
interface WordTrackedChange {
  type: 'Inserted' | 'Deleted';
  text: string;
  author: string;
  date: Date;
  accept(): void;
  reject(): void;
  getRange(): WordRange;
  load(properties: string): WordTrackedChange;
}

interface WordTrackedChangeCollection {
  load(properties: string): WordTrackedChangeCollection;
  items: WordTrackedChange[];
  acceptAll(): void;
  rejectAll(): void;
  getFirst(): WordTrackedChange;
  getFirstOrNullObject(): WordTrackedChange;
}

interface WordDocumentWithTrackedChanges {
  body: WordBody & {
    getTrackedChanges?(): WordTrackedChangeCollection;
  };
  getTrackedChanges?(): WordTrackedChangeCollection;
}

/**
 * Result of tracked change operations.
 */
export interface TrackedChangeResult {
  /** Whether the operation succeeded */
  success: boolean;
  /** Number of changes affected */
  count: number;
  /** Error message if failed */
  error?: string;
}

/**
 * Accept all tracked changes in the document.
 *
 * @param context - Word.RequestContext from Office.js
 * @returns TrackedChangeResult with count of accepted changes
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await acceptAllChanges(context);
 *   console.log(`Accepted ${result.count} changes`);
 * });
 * ```
 */
export async function acceptAllChanges(
  context: WordRequestContext
): Promise<TrackedChangeResult> {
  try {
    const doc = context.document as unknown as WordDocumentWithTrackedChanges;

    // Try document-level tracked changes first
    if (doc.getTrackedChanges) {
      const changes = doc.getTrackedChanges();
      changes.load('items');
      await context.sync();

      const count = changes.items.length;
      if (count > 0) {
        changes.acceptAll();
        await context.sync();
      }

      return { success: true, count };
    }

    // Try body-level tracked changes
    if (doc.body.getTrackedChanges) {
      const changes = doc.body.getTrackedChanges();
      changes.load('items');
      await context.sync();

      const count = changes.items.length;
      if (count > 0) {
        changes.acceptAll();
        await context.sync();
      }

      return { success: true, count };
    }

    return {
      success: false,
      count: 0,
      error: 'Tracked changes API not available in this version of Word',
    };
  } catch (err) {
    return {
      success: false,
      count: 0,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}

/**
 * Reject all tracked changes in the document.
 *
 * @param context - Word.RequestContext from Office.js
 * @returns TrackedChangeResult with count of rejected changes
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await rejectAllChanges(context);
 *   console.log(`Rejected ${result.count} changes`);
 * });
 * ```
 */
export async function rejectAllChanges(
  context: WordRequestContext
): Promise<TrackedChangeResult> {
  try {
    const doc = context.document as unknown as WordDocumentWithTrackedChanges;

    // Try document-level tracked changes first
    if (doc.getTrackedChanges) {
      const changes = doc.getTrackedChanges();
      changes.load('items');
      await context.sync();

      const count = changes.items.length;
      if (count > 0) {
        changes.rejectAll();
        await context.sync();
      }

      return { success: true, count };
    }

    // Try body-level tracked changes
    if (doc.body.getTrackedChanges) {
      const changes = doc.body.getTrackedChanges();
      changes.load('items');
      await context.sync();

      const count = changes.items.length;
      if (count > 0) {
        changes.rejectAll();
        await context.sync();
      }

      return { success: true, count };
    }

    return {
      success: false,
      count: 0,
      error: 'Tracked changes API not available in this version of Word',
    };
  } catch (err) {
    return {
      success: false,
      count: 0,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}

/**
 * Accept the next (first) tracked change in the document.
 *
 * Useful for stepping through changes one at a time.
 *
 * @param context - Word.RequestContext from Office.js
 * @returns TrackedChangeResult indicating success and remaining count
 *
 * @example
 * ```typescript
 * await Word.run(async (context) => {
 *   const result = await acceptNextChange(context);
 *   if (result.success) {
 *     console.log(`Accepted 1 change, ${result.count} remaining`);
 *   }
 * });
 * ```
 */
export async function acceptNextChange(
  context: WordRequestContext
): Promise<TrackedChangeResult> {
  try {
    const doc = context.document as unknown as WordDocumentWithTrackedChanges;

    const getChanges = doc.getTrackedChanges || doc.body.getTrackedChanges;
    if (!getChanges) {
      return {
        success: false,
        count: 0,
        error: 'Tracked changes API not available',
      };
    }

    const changes = getChanges.call(doc.getTrackedChanges ? doc : doc.body);
    const first = changes.getFirstOrNullObject();
    first.load('type');
    await context.sync();

    // Check if there was a change
    if ((first as unknown as { isNullObject?: boolean }).isNullObject) {
      return { success: true, count: 0 };
    }

    first.accept();
    await context.sync();

    // Get remaining count
    changes.load('items');
    await context.sync();

    return { success: true, count: changes.items.length };
  } catch (err) {
    return {
      success: false,
      count: 0,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}

/**
 * Reject the next (first) tracked change in the document.
 *
 * Useful for stepping through changes one at a time.
 *
 * @param context - Word.RequestContext from Office.js
 * @returns TrackedChangeResult indicating success and remaining count
 */
export async function rejectNextChange(
  context: WordRequestContext
): Promise<TrackedChangeResult> {
  try {
    const doc = context.document as unknown as WordDocumentWithTrackedChanges;

    const getChanges = doc.getTrackedChanges || doc.body.getTrackedChanges;
    if (!getChanges) {
      return {
        success: false,
        count: 0,
        error: 'Tracked changes API not available',
      };
    }

    const changes = getChanges.call(doc.getTrackedChanges ? doc : doc.body);
    const first = changes.getFirstOrNullObject();
    first.load('type');
    await context.sync();

    // Check if there was a change
    if ((first as unknown as { isNullObject?: boolean }).isNullObject) {
      return { success: true, count: 0 };
    }

    first.reject();
    await context.sync();

    // Get remaining count
    changes.load('items');
    await context.sync();

    return { success: true, count: changes.items.length };
  } catch (err) {
    return {
      success: false,
      count: 0,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}

/**
 * Get information about tracked changes in the document.
 *
 * @param context - Word.RequestContext from Office.js
 * @returns Information about tracked changes
 */
export async function getTrackedChangesInfo(
  context: WordRequestContext
): Promise<{
  available: boolean;
  count: number;
  insertions: number;
  deletions: number;
  changes: Array<{ type: string; text: string; author: string }>;
}> {
  try {
    const doc = context.document as unknown as WordDocumentWithTrackedChanges;

    const getChanges = doc.getTrackedChanges || doc.body.getTrackedChanges;
    if (!getChanges) {
      return { available: false, count: 0, insertions: 0, deletions: 0, changes: [] };
    }

    const changes = getChanges.call(doc.getTrackedChanges ? doc : doc.body);
    changes.load('items');
    await context.sync();

    // Load details for each change
    for (const change of changes.items) {
      change.load('type,text,author');
    }
    await context.sync();

    const insertions = changes.items.filter((c) => c.type === 'Inserted').length;
    const deletions = changes.items.filter((c) => c.type === 'Deleted').length;

    return {
      available: true,
      count: changes.items.length,
      insertions,
      deletions,
      changes: changes.items.map((c) => ({
        type: c.type,
        text: c.text,
        author: c.author,
      })),
    };
  } catch (err) {
    return {
      available: false,
      count: 0,
      insertions: 0,
      deletions: 0,
      changes: [],
    };
  }
}
