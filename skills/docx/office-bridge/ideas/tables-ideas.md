---
name: tables-ideas
description: "Roadmap and ideas for enhanced table manipulation in Office Bridge. Helper functions, table-aware editing, formatting operations, and data import/export capabilities."
---

# Table Manipulation Ideas for Office Bridge

This document captures ideas for enhancing table manipulation capabilities in the Office Bridge Word add-in.

## Helper Functions to Build

### 1. Table Creation Helpers

```typescript
// Create table from array with smart defaults
async function insertTableFromArray(
  context: WordRequestContext,
  data: string[][],
  options?: {
    insertLocation?: "Start" | "End" | { afterRef: Ref };
    style?: string;
    hasHeader?: boolean;
    autoFit?: "Content" | "Window" | "FixedSize";
  }
): Promise<{ ref: Ref; table: Word.Table }>

// Create table from object array (auto-generates headers from keys)
async function insertTableFromObjects(
  context: WordRequestContext,
  objects: Record<string, any>[],
  options?: {
    columns?: string[];  // Subset/order of keys to include
    headers?: string[];  // Custom header names
  }
): Promise<{ ref: Ref; table: Word.Table }>
```

**Use Cases:**
- Quick table creation from API responses
- Converting JSON data to formatted tables
- Template population with tabular data

### 2. Row Operations

```typescript
// Add row by table ref
async function addRow(
  context: WordRequestContext,
  tableRef: Ref,
  cells: string[],
  options?: {
    position?: "Start" | "End" | { afterRow: number | string };
    copyFormatFrom?: number;  // Row index to copy formatting from
    track?: boolean;
  }
): Promise<{ ref: Ref }>

// Add multiple rows efficiently
async function addRows(
  context: WordRequestContext,
  tableRef: Ref,
  rows: string[][],
  options?: {
    position?: "Start" | "End";
    track?: boolean;
  }
): Promise<{ refs: Ref[] }>

// Delete row by ref or text match
async function deleteRow(
  context: WordRequestContext,
  rowRef: Ref | { tableRef: Ref; containsText: string },
  options?: { track?: boolean }
): Promise<{ success: boolean }>

// Swap two rows
async function swapRows(
  context: WordRequestContext,
  rowRef1: Ref,
  rowRef2: Ref
): Promise<{ success: boolean }>

// Sort rows by column
async function sortRows(
  context: WordRequestContext,
  tableRef: Ref,
  options: {
    columnIndex: number;
    ascending?: boolean;
    skipHeader?: boolean;
    numeric?: boolean;
  }
): Promise<{ success: boolean }>
```

### 3. Column Operations

```typescript
// Add column to table
async function addColumn(
  context: WordRequestContext,
  tableRef: Ref,
  cells: string[],
  options?: {
    position?: "Start" | "End" | { afterColumn: number | string };
    header?: string;  // If provided, first cell is header
    track?: boolean;
  }
): Promise<{ columnIndex: number }>

// Delete column by index or header text
async function deleteColumn(
  context: WordRequestContext,
  tableRef: Ref,
  column: number | string,  // Index or header text
  options?: { track?: boolean }
): Promise<{ success: boolean }>

// Rename column header
async function renameColumn(
  context: WordRequestContext,
  tableRef: Ref,
  oldHeader: string,
  newHeader: string,
  options?: { track?: boolean }
): Promise<{ success: boolean }>

// Reorder columns
async function reorderColumns(
  context: WordRequestContext,
  tableRef: Ref,
  newOrder: (number | string)[]  // Indices or header names
): Promise<{ success: boolean }>
```

### 4. Cell Merge/Split Operations

```typescript
// Merge cells in a range
async function mergeCells(
  context: WordRequestContext,
  tableRef: Ref,
  range: {
    startRow: number;
    startCell: number;
    endRow: number;
    endCell: number;
  },
  options?: {
    separator?: string;  // How to join content: ", " | "\n" | ""
  }
): Promise<{ ref: Ref }>

// Merge entire row
async function mergeRow(
  context: WordRequestContext,
  rowRef: Ref
): Promise<{ ref: Ref }>

// Merge entire column
async function mergeColumn(
  context: WordRequestContext,
  tableRef: Ref,
  columnIndex: number
): Promise<{ success: boolean }>

// Split cell
async function splitCell(
  context: WordRequestContext,
  cellRef: Ref,
  rows: number,
  columns: number
): Promise<{ refs: Ref[] }>
```

### 5. Cell Content Operations

```typescript
// Set cell value by ref
async function setCellValue(
  context: WordRequestContext,
  cellRef: Ref,
  value: string,
  options?: { track?: boolean }
): Promise<{ success: boolean }>

// Get cell value by ref
async function getCellValue(
  context: WordRequestContext,
  cellRef: Ref
): Promise<string>

// Set multiple cells efficiently
async function setCellValues(
  context: WordRequestContext,
  updates: Array<{ ref: Ref; value: string }>,
  options?: { track?: boolean }
): Promise<{ successCount: number }>

// Clear cell content (keep cell)
async function clearCell(
  context: WordRequestContext,
  cellRef: Ref
): Promise<{ success: boolean }>
```

## Table-Aware Editing Ideas

### 1. Column-Wide Operations

```typescript
// Edit all cells in a column
async function editColumn(
  context: WordRequestContext,
  tableRef: Ref,
  columnIndex: number | string,  // Index or header
  operation: {
    type: "replace" | "prefix" | "suffix" | "format";
    find?: string;
    replace?: string;
    text?: string;
    formatting?: FormattingOptions;
  },
  options?: {
    skipHeader?: boolean;
    track?: boolean;
  }
): Promise<{ modifiedCount: number }>

// Apply formula to column
async function applyColumnFormula(
  context: WordRequestContext,
  tableRef: Ref,
  targetColumn: number,
  formula: string,  // e.g., "=B2*C2", "=SUM(LEFT)"
  options?: { skipHeader?: boolean }
): Promise<{ success: boolean }>
```

### 2. Row-Wide Operations

```typescript
// Format entire row
async function formatRow(
  context: WordRequestContext,
  rowRef: Ref,
  formatting: {
    backgroundColor?: string;
    fontColor?: string;
    bold?: boolean;
    alignment?: "left" | "center" | "right";
  }
): Promise<{ success: boolean }>

// Duplicate row with modifications
async function duplicateRow(
  context: WordRequestContext,
  rowRef: Ref,
  modifications?: Record<number, string>,  // Column index -> new value
  options?: { insertAfter?: boolean }
): Promise<{ ref: Ref }>
```

### 3. Conditional Operations

```typescript
// Format cells matching condition
async function conditionalFormat(
  context: WordRequestContext,
  tableRef: Ref,
  condition: {
    column: number | string;
    operator: "equals" | "contains" | "greaterThan" | "lessThan" | "regex";
    value: string | number;
  },
  formatting: FormattingOptions,
  options?: { applyTo: "cell" | "row" }
): Promise<{ matchCount: number }>

// Delete rows matching condition
async function deleteRowsWhere(
  context: WordRequestContext,
  tableRef: Ref,
  condition: {
    column: number | string;
    operator: "equals" | "contains" | "isEmpty" | "regex";
    value?: string;
  },
  options?: { track?: boolean }
): Promise<{ deletedCount: number }>

// Filter table (hide non-matching rows)
// Note: Office.js doesn't support row hiding, so this would need workaround
async function filterTable(
  context: WordRequestContext,
  tableRef: Ref,
  filters: Record<string, string>  // Column header -> filter value
): Promise<{ visibleRows: number }>
```

## Table Formatting Operations

### 1. Style Helpers

```typescript
// Apply professional table style
async function applyTableStyle(
  context: WordRequestContext,
  tableRef: Ref,
  style: "professional" | "minimal" | "striped" | "bordered" | "colorful",
  options?: {
    primaryColor?: string;
    accentColor?: string;
    headerStyle?: "bold" | "inverted" | "underlined";
  }
): Promise<{ success: boolean }>

// Copy formatting from another table
async function copyTableFormatting(
  context: WordRequestContext,
  sourceRef: Ref,
  targetRef: Ref
): Promise<{ success: boolean }>
```

### 2. Border Helpers

```typescript
// Set all borders at once
async function setTableBorders(
  context: WordRequestContext,
  tableRef: Ref,
  borders: {
    outside?: BorderStyle;
    inside?: BorderStyle;
    horizontal?: BorderStyle;
    vertical?: BorderStyle;
  }
): Promise<{ success: boolean }>

// Remove all borders
async function removeBorders(
  context: WordRequestContext,
  tableRef: Ref,
  which?: "all" | "inside" | "outside"
): Promise<{ success: boolean }>
```

### 3. Layout Helpers

```typescript
// Auto-size columns to content
async function autoSizeColumns(
  context: WordRequestContext,
  tableRef: Ref
): Promise<{ success: boolean }>

// Set specific column widths
async function setColumnWidths(
  context: WordRequestContext,
  tableRef: Ref,
  widths: number[] | Record<string, number>  // Points or percentages
): Promise<{ success: boolean }>

// Center table on page
async function centerTable(
  context: WordRequestContext,
  tableRef: Ref
): Promise<{ success: boolean }>
```

## Data Import/Export

### 1. Array Conversion

```typescript
// Convert table to 2D array
async function tableToArray(
  context: WordRequestContext,
  tableRef: Ref,
  options?: {
    includeHeader?: boolean;
    trimWhitespace?: boolean;
  }
): Promise<string[][]>

// Convert table to object array (using first row as keys)
async function tableToObjects(
  context: WordRequestContext,
  tableRef: Ref,
  options?: {
    keyTransform?: (header: string) => string;  // e.g., camelCase
  }
): Promise<Record<string, string>[]>
```

### 2. CSV/TSV Import

```typescript
// Import CSV data into table
async function importCSV(
  context: WordRequestContext,
  csvContent: string,
  options?: {
    delimiter?: "," | "\t" | ";";
    hasHeader?: boolean;
    insertLocation?: "Start" | "End" | { afterRef: Ref };
    style?: string;
  }
): Promise<{ ref: Ref; rowCount: number }>

// Export table to CSV format
async function exportToCSV(
  context: WordRequestContext,
  tableRef: Ref,
  options?: {
    delimiter?: "," | "\t" | ";";
    includeHeader?: boolean;
  }
): Promise<string>
```

### 3. Markdown Conversion

```typescript
// Convert table to Markdown format
async function tableToMarkdown(
  context: WordRequestContext,
  tableRef: Ref
): Promise<string>

// Import Markdown table
async function importMarkdownTable(
  context: WordRequestContext,
  markdown: string,
  insertLocation: { afterRef: Ref }
): Promise<{ ref: Ref }>
```

### 4. JSON Integration

```typescript
// Update table from JSON API response
async function updateTableFromJSON(
  context: WordRequestContext,
  tableRef: Ref,
  data: Record<string, any>[],
  options?: {
    matchColumn?: string;  // Column to use for row matching
    addNewRows?: boolean;
    removeUnmatched?: boolean;
    track?: boolean;
  }
): Promise<{
  updated: number;
  added: number;
  removed: number;
}>
```

## Table Analysis

### 1. Structure Analysis

```typescript
// Get table structure info
async function analyzeTable(
  context: WordRequestContext,
  tableRef: Ref
): Promise<{
  rowCount: number;
  columnCount: number;
  hasMergedCells: boolean;
  isUniform: boolean;
  headers: string[];
  estimatedDataType: Record<string, "text" | "number" | "date" | "currency">;
}>
```

### 2. Data Validation

```typescript
// Validate table data
async function validateTable(
  context: WordRequestContext,
  tableRef: Ref,
  rules: Record<string, {
    type?: "number" | "date" | "email" | "phone" | "regex";
    pattern?: string;
    required?: boolean;
    min?: number;
    max?: number;
  }>
): Promise<{
  valid: boolean;
  errors: Array<{
    row: number;
    column: string;
    error: string;
  }>;
}>
```

## Implementation Priority

### Phase 1: Core Helpers (Essential)
1. `tableToArray` / `insertTableFromArray` - Basic data conversion
2. `addRow` / `deleteRow` - Row manipulation
3. `setCellValue` / `getCellValue` - Cell access
4. `applyTableStyle` - Quick formatting

### Phase 2: Column Operations
1. `addColumn` / `deleteColumn` - Column manipulation
2. `editColumn` - Bulk column edits
3. `reorderColumns` - Column reorganization
4. `setColumnWidths` - Layout control

### Phase 3: Advanced Features
1. `mergeCells` / `splitCell` - Cell structure
2. `conditionalFormat` - Conditional formatting
3. `sortRows` - Data sorting
4. `tableToObjects` / `importCSV` - Data interchange

### Phase 4: Analysis & Validation
1. `analyzeTable` - Structure detection
2. `validateTable` - Data validation
3. `updateTableFromJSON` - Dynamic updates

## Integration with DocTree

All table helpers should integrate with the existing DocTree ref system:

```typescript
// After any table modification, update the tree
const result = await DocTree.addRow(context, "tbl:0", ["A", "B", "C"]);
// result.ref = "tbl:0/row:5"

// Tree is automatically updated
const tree = await DocTree.buildTree(context);
// Tree now includes the new row
```

## Notes

- All operations should support the `track: boolean` option where applicable
- Performance: batch operations when possible to minimize `context.sync()` calls
- Error handling: return detailed error messages for debugging
- Consider Office.js API version requirements for each feature
