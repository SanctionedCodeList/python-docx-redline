# Excel Patterns

Excel provides workbook, worksheet, and range manipulation via `Excel.run()`.

## Getting Workbooks

```typescript
const bridge = await connect();
const workbooks = await bridge.excel();  // Returns ExcelSession[]
const wb = workbooks[0];
```

## Execution Context

Code runs inside `Excel.run()` with access to `context`, `Excel`, and `Office`:

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
sheet.load("name");
await context.sync();
return sheet.name;
```

## Common Patterns

### Read Active Sheet Name

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
sheet.load("name");
await context.sync();
return sheet.name;
```

### Read a Range

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1:B10");
range.load("values");
await context.sync();
return range.values;  // 2D array
```

### Write to Cells

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1");
range.values = [["Hello from Claude!"]];
await context.sync();
return "Written";
```

### Write Multiple Cells

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1:B2");
range.values = [
  ["Name", "Value"],
  ["Item 1", 100]
];
await context.sync();
return "Written";
```

### Get Used Range

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const usedRange = sheet.getUsedRange();
usedRange.load("address,values");
await context.sync();
return { address: usedRange.address, values: usedRange.values };
```

### Format Cells

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1:B1");
range.format.font.bold = true;
range.format.fill.color = "#FFFF00";
await context.sync();
return "Formatted";
```

### Add a Worksheet

```javascript
const newSheet = context.workbook.worksheets.add("NewSheet");
newSheet.activate();
await context.sync();
return "Sheet added";
```

### List All Worksheets

```javascript
const sheets = context.workbook.worksheets;
sheets.load("items/name");
await context.sync();
return sheets.items.map(s => s.name);
```

### Create a Table

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const table = sheet.tables.add("A1:C4", true);  // true = has headers
table.name = "SalesTable";
await context.sync();
return "Table created";
```

## Tips

- Always call `await context.sync()` after queueing operations
- Use `.load()` to specify which properties to fetch
- Range addresses use Excel notation: `"A1"`, `"A1:B10"`, `"Sheet2!A1"`
- Values are 2D arrays even for single cells: `[[value]]`
