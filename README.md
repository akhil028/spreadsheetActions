# Pure Spreadsheet JS

A lightweight, pure JavaScript spreadsheet library with support for cell merging, multiple sheets, and Excel-like functionality.

## Features

- ✅ **Multiple Sheets**: Create, manage, and switch between multiple worksheets
- ✅ **Cell Merging**: Merge and unmerge cells with collision detection
- ✅ **Row/Column Operations**: Insert, delete, and manipulate rows and columns
- ✅ **Data Import/Export**: JSON serialization for data persistence
- ✅ **Formula Support**: Basic formula handling (extensible)
- ✅ **Cell Styling**: Style individual cells
- ✅ **Pure JavaScript**: No external dependencies
- ✅ **TypeScript Ready**: Full TypeScript definitions included

## Installation

\`\`\`bash
npm install pure-spreadsheet-js
\`\`\`

## Quick Start

\`\`\`javascript
const { Spreadsheet } = require('pure-spreadsheet-js');

// Create a new spreadsheet
const spreadsheet = new Spreadsheet();

// Add a sheet
spreadsheet.createSheet('Sales Data');

// Get the active sheet
const sheet = spreadsheet.getActiveSheet();

// Set cell values
sheet.setCell(0, 0, 'Product');
sheet.setCell(0, 1, 'Revenue');
sheet.setCell(1, 0, 'Widget A');
sheet.setCell(1, 1, 1000);

// Merge cells
sheet.mergeCells(0, 0, 0, 1); // Merge header row

// Add another sheet
spreadsheet.createSheet('Inventory');

// Export to JSON
const data = spreadsheet.toJSON();
console.log(JSON.stringify(data, null, 2));
\`\`\`

## API Reference

### Spreadsheet Class

#### Methods

##### \`createSheet(name)\`
Creates a new worksheet.

\`\`\`javascript
spreadsheet.createSheet('My Sheet');
\`\`\`

##### \`getSheet(name)\`
Gets a sheet by name.

\`\`\`javascript
const sheet = spreadsheet.getSheet('My Sheet');
\`\`\`

##### \`removeSheet(name)\`
Removes a sheet.

\`\`\`javascript
spreadsheet.removeSheet('My Sheet');
\`\`\`

##### \`setActiveSheet(name)\`
Sets the active sheet.

\`\`\`javascript
spreadsheet.setActiveSheet('My Sheet');
\`\`\`

##### \`getActiveSheet()\`
Gets the currently active sheet.

\`\`\`javascript
const activeSheet = spreadsheet.getActiveSheet();
\`\`\`

### Sheet Class

#### Methods

##### \`setCell(row, col, value)\`
Sets a cell value.

\`\`\`javascript
sheet.setCell(0, 0, 'Hello World');
\`\`\`

##### \`getCell(row, col)\`
Gets a cell object.

\`\`\`javascript
const cell = sheet.getCell(0, 0);
console.log(cell.getValue()); // 'Hello World'
\`\`\`

##### \`mergeCells(startRow, startCol, endRow, endCol)\`
Merges a range of cells.

\`\`\`javascript
sheet.mergeCells(0, 0, 2, 2); // Merge 3x3 area
\`\`\`

##### \`unmergeCells(startRow, startCol, endRow, endCol)\`
Unmerges a range of cells.

\`\`\`javascript
sheet.unmergeCells(0, 0, 2, 2);
\`\`\`

##### \`insertRow(index)\`
Inserts a new row.

\`\`\`javascript
sheet.insertRow(1); // Insert at row 1
\`\`\`

##### \`insertColumn(index)\`
Inserts a new column.

\`\`\`javascript
sheet.insertColumn(1); // Insert at column 1
\`\`\`

##### \`toArray()\`
Exports sheet data as 2D array.

\`\`\`javascript
const data = sheet.toArray();
console.log(data); // [['A1', 'B1'], ['A2', 'B2']]
\`\`\`

## Advanced Usage

### Working with Formulas

\`\`\`javascript
sheet.setCellFormula(2, 0, '=A1+B1');
sheet.setCell(2, 1, '=SUM(A1:A10)');
\`\`\`

### Cell Styling

\`\`\`javascript
const cell = sheet.getCell(0, 0);
cell.setStyle({
  backgroundColor: '#ff0000',
  color: '#ffffff',
  fontWeight: 'bold'
});
\`\`\`

### Data Persistence

\`\`\`javascript
// Export
const data = spreadsheet.toJSON();
localStorage.setItem('spreadsheet', JSON.stringify(data));

// Import
const savedData = JSON.parse(localStorage.getItem('spreadsheet'));
const newSpreadsheet = new Spreadsheet();
newSpreadsheet.fromJSON(savedData);
\`\`\`

## Browser Support

- Chrome 60+
- Firefox 55+
- Safari 11+
- Edge 79+

## License

MIT License
