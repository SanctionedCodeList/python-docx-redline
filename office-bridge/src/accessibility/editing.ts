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
} from './types';

// =============================================================================
// Office.js Type Declarations
// =============================================================================

/**
 * Minimal Office.js Word.RequestContext interface for editing operations.
 */
interface WordRequestContext {
  document: {
    body: {
      paragraphs: WordParagraphCollection;
      tables: WordTableCollection;
    };
    getSelection(): WordRange;
  };
  sync(): Promise<void>;
  trackedChanges?: {
    enabled: boolean;
  };
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
    const fnMatch = footnotesMatch[1].match(fnPattern);
    if (!fnMatch) return undefined;

    // Extract text from the footnote
    const textParts = fnMatch[1].match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
    const text = textParts
      .map((t) => t.replace(/<[^>]+>/g, ''))
      .join('')
      .trim();

    return text || undefined;
  } catch {
    return undefined;
  }
}
