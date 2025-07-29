<div align="center">

# 📊 Pure Spreadsheet JS

### A powerful, lightweight JavaScript spreadsheet library

[![npm version](https://badge.fury.io/js/pure-spreadsheet-js.svg)](https://badge.fury.io/js/pure-spreadsheet-js)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://img.shields.io/npm/dm/pure-spreadsheet-js.svg)](https://npmjs.org/package/pure-spreadsheet-js)

**Create Excel-like spreadsheets with pure JavaScript • Zero dependencies • Full TypeScript support**

[🚀 Quick Start](#-quick-start) • [💡 Examples](#-examples) • [📖 API Reference](#-api-reference) • [🔧 Advanced Usage](#-advanced-usage)

</div>

---

## ✨ What is Pure Spreadsheet JS?

Pure Spreadsheet JS is a comprehensive JavaScript library that brings Excel-like functionality to your web applications. Built with **zero dependencies**, it provides all the essential spreadsheet features you need including multiple worksheets, cell merging, data manipulation, and much more.

### 🎯 Perfect For
- **📈 Business Dashboards** - Create interactive data visualization tools
- **💼 Enterprise Applications** - Build CRM, ERP, and inventory management systems
- **📊 Financial Tools** - Develop budgeting, forecasting, and accounting software
- **🎓 Educational Platforms** - Create grade books and learning management tools
- **📋 Data Management** - Build dynamic forms and data entry applications

---

## 🌟 Core Features (7 Main Capabilities)

<table>
<tr>
<td width="50%">

### 1️⃣ **Create Sheets**
Create new worksheets in your workbook
workbook.createSheet('Sales Data');

text

### 2️⃣ **Add Sheets to Workbook**
Add new worksheet tabs to existing workbook
workbook.addSheetToWorkbook('Inventory', {
position: 'after',
referenceSheet: 'Sales Data'
});

text

### 3️⃣ **Add Multiple Rows & Columns**
Insert multiple rows and columns at once
sheet.insertMultipleRows(1, 5); // Add 5 rows
sheet.insertMultipleColumns(2, 3); // Add 3 columns

text

### 4️⃣ **Add Headers**
Set styled headers for your data
sheet.setHeaders(['Product', 'Price', 'Stock', 'Total']);

text

</td>
<td width="50%">

### 5️⃣ **Delete Data from Cells**
Clear or delete cell content and entire cells
sheet.clearCell(1, 2); // Clear content
sheet.deleteCell(0, 0); // Delete cell entirely
sheet.clearRange(0, 0, 2, 2); // Clear range

text

### 6️⃣ **Merge Rows & Columns with Separator**
Combine data from multiple cells using separators
const rowData = sheet.mergeRowData(1, ' | ');
const colData = sheet.mergeColumnData(0, ', ');

text

### 7️⃣ **Concatenate Data**
Combine data with rows, columns, and specific cells
sheet.concatenateWithRow(1, 3, 0, ' - ');
sheet.concatenateWithCell(1, 0, 2, 1, ' + ');

text

</td>
</tr>
</table>

### 🔧 Additional Features
- **Cell Merging** - Traditional spreadsheet cell merging
- **Cell Styling** - Apply custom formatting
- **Formula Support** - Basic formula handling
- **JSON Import/Export** - Data persistence
- **Search Functionality** - Find data across sheets
- **TypeScript Ready** - Full type definitions included

---

## 🚀 Quick Start

### Installation

npm install pure-spreadsheet-js

text

### Basic Example

const { Spreadsheet } = require('pure-spreadsheet-js');

// Create a new workbook
const workbook = new Spreadsheet();

// 1. Create a sheet
workbook.createSheet('Sales Report');
const sheet = workbook.getActiveSheet();

// 4. Add headers
sheet.setHeaders(['Product', 'Q1 Sales', 'Q2 Sales', 'Total']);

// 3. Add multiple rows
sheet.insertMultipleRows(1, 3);

// Add sample data
sheet.setCell(1, 0, 'Widget A');
sheet.setCell(1, 1, 1000);
sheet.setCell(1, 2, 1200);
sheet.setCell(1, 3, 2200);

sheet.setCell(2, 0, 'Widget B');
sheet.setCell(2, 1, 800);
sheet.setCell(2, 2, 900);
sheet.setCell(2, 3, 1700);

// 6. Merge row data with separator
const productSummary = sheet.mergeRowData(1, ' | ');
console.log('Product Summary:', productSummary);
// Output: "Widget A | 1000 | 1200 | 2200"

// 7. Concatenate data with specific cell
sheet.concatenateWithRow(1, 3, 4, ' = ');

// 2. Add another sheet to workbook
workbook.addSheetToWorkbook('Inventory', {
headers: ['Item', 'Stock', 'Location']
});

// 5. Delete data from cells
sheet.clearCell(2, 3); // Clear Widget B total

console.log('📊 Workbook created successfully!');
console.log('📋 Available sheets:', workbook.getSheetNames());
console.log('📈 Sales data:', sheet.toArray());

text

**Output:**
Product Summary: Widget A | 1000 | 1200 | 2200
📊 Workbook created successfully!
📋 Available sheets: ['Sales Report', 'Inventory']
📈 Sales data: [
['Product', 'Q1 Sales', 'Q2 Sales', 'Total'],
['Widget A', 1000, 1200, 2200],
['Widget B', 800, 900, ''],
['', '', '', 'Widget A | 1000 | 1200 | 2200 = ']
]

text

---

## 💡 Real-World Examples

### 📊 E-commerce Product Management

const { Spreadsheet } = require('pure-spreadsheet-js');

// Create product management system
const catalog = new Spreadsheet();

// 1. Create main products sheet
catalog.createSheet('Products');
const productSheet = catalog.getActiveSheet();

// 4. Set up headers
productSheet.setHeaders(['SKU', 'Name', 'Category', 'Price', 'Stock', 'Status']);

// 3. Add multiple rows for products
productSheet.insertMultipleRows(1, 5);

// Add product data
const products = [
['SKU001', 'Wireless Mouse', 'Electronics', 25.99, 150, 'In Stock'],
['SKU002', 'USB Cable', 'Electronics', 12.99, 5, 'Low Stock'],
['SKU003', 'Laptop Stand', 'Accessories', 45.99, 30, 'In Stock'],
['SKU004', 'Bluetooth Speaker', 'Electronics', 89.99, 0, 'Out of Stock']
];

products.forEach((product, index) => {
product.forEach((value, col) => {
productSheet.setCell(index + 1, col, value);
});
});

// 2. Add inventory summary sheet
catalog.addSheetToWorkbook('Inventory Summary', {
position: 'after',
referenceSheet: 'Products',
headers: ['Category', 'Total Items', 'Low Stock Items', 'Total Value']
});

const summarySheet = catalog.getSheet('Inventory Summary');

// 6. Merge category data
const categories = productSheet.mergeColumnData(2, ', ');
summarySheet.setCell(1, 0, 'Electronics');
summarySheet.setCell(1, 1, 3);
summarySheet.setCell(1, 2, 1); // USB Cable is low stock

// 7. Concatenate product summary
const productSummary = productSheet.mergeRowData(1, ' - ');
summarySheet.setCell(2, 0, 'Best Seller');
summarySheet.setCell(2, 1, productSummary);

// 5. Clear out of stock item price (for discount calculation)
productSheet.clearCell(4, 3);

console.log('🛍️ Product catalog created!');
console.log('📦 Products:', productSheet.toArray());
console.log('📊 Summary:', summarySheet.toArray());

text

### 📈 Financial Budget Tracker

const { Spreadsheet } = require('pure-spreadsheet-js');

// Create budget tracking workbook
const budget = new Spreadsheet();

// 1. Create monthly budget sheet
budget.createSheet('Monthly Budget');
const budgetSheet = budget.getActiveSheet();

// 4. Add financial headers
budgetSheet.setHeaders(['Category', 'Planned', 'Actual', 'Difference', 'Status']);

// 3. Add rows for different categories
budgetSheet.insertMultipleRows(1, 8);

// Add budget categories
const categories = [
['Housing', 2000, 1950, -50, 'Under Budget'],
['Food', 600, 720, 120, 'Over Budget'],
['Transportation', 400, 380, -20, 'Under Budget'],
['Entertainment', 300, 450, 150, 'Over Budget'],
['Utilities', 250, 245, -5, 'Under Budget'],
['Healthcare', 200, 180, -20, 'Under Budget'],
['Savings', 500, 500, 0, 'On Target']
];

categories.forEach((category, index) => {
category.forEach((value, col) => {
budgetSheet.setCell(index + 1, col, value);
});
});

// 2. Add summary analysis sheet
budget.addSheetToWorkbook('Budget Analysis', {
headers: ['Analysis', 'Value', 'Recommendation']
});

const analysisSheet = budget.getSheet('Budget Analysis');

// 6. Merge budget data for analysis
const plannedTotal = budgetSheet.mergeColumnData(1, ' + ');
const actualTotal = budgetSheet.mergeColumnData(2, ' + ');

analysisSheet.setCell(1, 0, 'Total Planned');
analysisSheet.setCell(1, 1, '$4,250');
analysisSheet.setCell(1, 2, 'Stick to plan');

analysisSheet.setCell(2, 0, 'Total Actual');
analysisSheet.setCell(2, 1, '$4,425');
analysisSheet.setCell(2, 2, 'Reduce spending');

// 7. Concatenate over-budget categories
const overBudgetSummary = budgetSheet.mergeRowData(2, ' | '); // Food row
analysisSheet.setCell(3, 0, 'Over Budget Alert');
analysisSheet.setCell(3, 1, overBudgetSummary);

// 5. Clear planned amount for categories that need revision
budgetSheet.clearCell(2, 1); // Clear food planned amount for revision

console.log('💰 Budget tracker ready!');
console.log('📊 Budget data:', budgetSheet.toArray());
console.log('📈 Analysis:', analysisSheet.toArray());

text

---

## 📖 Complete API Reference

### 🗂️ Spreadsheet Class

#### Core Sheet Management

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `createSheet(name)` | Create new worksheet | `name: string` | `Spreadsheet` | `workbook.createSheet('Data')` |
| `addSheetToWorkbook(name, options)` | Add sheet with options | `name: string, options: object` | `Spreadsheet` | `workbook.addSheetToWorkbook('Report', { headers: ['A', 'B'] })` |
| `getSheet(name)` | Get sheet by name | `name: string` | `Sheet` | `const sheet = workbook.getSheet('Data')` |
| `getActiveSheet()` | Get current active sheet | - | `Sheet` | `const active = workbook.getActiveSheet()` |
| `setActiveSheet(name)` | Switch to different sheet | `name: string` | `Spreadsheet` | `workbook.setActiveSheet('Report')` |
| `getSheetNames()` | Get all sheet names in order | - | `string[]` | `const names = workbook.getSheetNames()` |
| `getSheetCount()` | Get total number of sheets | - | `number` | `const count = workbook.getSheetCount()` |

#### Advanced Sheet Operations

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `removeSheet(name)` | Remove sheet from workbook | `name: string` | `Spreadsheet` | `workbook.removeSheet('OldData')` |
| `renameSheet(oldName, newName)` | Rename existing sheet | `oldName: string, newName: string` | `Spreadsheet` | `workbook.renameSheet('Sheet1', 'Sales')` |
| `cloneSheet(source, target, position)` | Duplicate sheet with data | `source: string, target: string, position?: string` | `Spreadsheet` | `workbook.cloneSheet('Template', 'NewData')` |
| `findInWorkbook(searchValue, options)` | Search across all sheets | `searchValue: any, options?: object` | `object[]` | `workbook.findInWorkbook('Widget')` |
| `getWorkbookInfo()` | Get workbook statistics | - | `object` | `const info = workbook.getWorkbookInfo()` |

#### Data Import/Export

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `toJSON()` | Export workbook to JSON | - | `object` | `const data = workbook.toJSON()` |
| `fromJSON(data)` | Import workbook from JSON | `data: object` | `Spreadsheet` | `workbook.fromJSON(savedData)` |
| `clear()` | Clear all sheets and data | - | `Spreadsheet` | `workbook.clear()` |

### 📋 Sheet Class

#### Basic Cell Operations

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `setCell(row, col, value)` | Set cell value | `row: number, col: number, value: any` | `Sheet` | `sheet.setCell(0, 0, 'Hello')` |
| `getCell(row, col)` | Get cell object | `row: number, col: number` | `Cell` | `const cell = sheet.getCell(0, 0)` |
| `setCellFormula(row, col, formula)` | Set cell formula | `row: number, col: number, formula: string` | `Sheet` | `sheet.setCellFormula(0, 0, '=A1+B1')` |

#### Row and Column Operations

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `insertRow(index)` | Insert single row | `index: number` | `Sheet` | `sheet.insertRow(1)` |
| `insertColumn(index)` | Insert single column | `index: number` | `Sheet` | `sheet.insertColumn(2)` |
| `insertMultipleRows(start, count)` | Insert multiple rows | `start: number, count: number` | `Sheet` | `sheet.insertMultipleRows(1, 5)` |
| `insertMultipleColumns(start, count)` | Insert multiple columns | `start: number, count: number` | `Sheet` | `sheet.insertMultipleColumns(2, 3)` |

#### Header Management

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `setHeaders(headers)` | Set styled headers | `headers: string[]` | `Sheet` | `sheet.setHeaders(['Name', 'Age'])` |
| `getHeaders()` | Get header values | - | `string[]` | `const headers = sheet.getHeaders()` |

#### Data Deletion

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `clearCell(row, col)` | Clear cell content | `row: number, col: number` | `Sheet` | `sheet.clearCell(1, 2)` |
| `clearRange(startRow, startCol, endRow, endCol)` | Clear cell range | `startRow: number, startCol: number, endRow: number, endCol: number` | `Sheet` | `sheet.clearRange(0, 0, 2, 2)` |
| `deleteCell(row, col)` | Delete cell entirely | `row: number, col: number` | `Sheet` | `sheet.deleteCell(0, 0)` |

#### Data Merging with Separators

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `mergeRowData(rowIndex, separator)` | Merge row data with separator | `rowIndex: number, separator?: string` | `string` | `sheet.mergeRowData(1, ' \| ')` |
| `mergeColumnData(colIndex, separator)` | Merge column data with separator | `colIndex: number, separator?: string` | `string` | `sheet.mergeColumnData(0, ', ')` |
| `mergeRangeData(startRow, startCol, endRow, endCol, separator)` | Merge range data with separator | `startRow: number, startCol: number, endRow: number, endCol: number, separator?: string` | `string` | `sheet.mergeRangeData(0, 0, 2, 2, ' - ')` |

#### Data Concatenation

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `concatenateWithRow(sourceRow, targetRow, targetCol, separator)` | Concatenate row data into cell | `sourceRow: number, targetRow: number, targetCol: number, separator?: string` | `Sheet` | `sheet.concatenateWithRow(1, 3, 0, ' = ')` |
| `concatenateWithColumn(sourceCol, targetRow, targetCol, separator)` | Concatenate column data into cell | `sourceCol: number, targetRow: number, targetCol: number, separator?: string` | `Sheet` | `sheet.concatenateWithColumn(0, 2, 4, ' + ')` |
| `concatenateWithCell(sourceRow, sourceCol, targetRow, targetCol, separator)` | Concatenate two cells | `sourceRow: number, sourceCol: number, targetRow: number, targetCol: number, separator?: string` | `Sheet` | `sheet.concatenateWithCell(1, 0, 2, 1, ' & ')` |

#### Cell Merging (Visual)

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `mergeCells(startRow, startCol, endRow, endCol)` | Merge cell range visually | `startRow: number, startCol: number, endRow: number, endCol: number` | `Sheet` | `sheet.mergeCells(0, 0, 2, 2)` |
| `unmergeCells(startRow, startCol, endRow, endCol)` | Unmerge cell range | `startRow: number, startCol: number, endRow: number, endCol: number` | `Sheet` | `sheet.unmergeCells(0, 0, 2, 2)` |
| `getMergedRanges()` | Get all merged ranges | - | `object[]` | `const ranges = sheet.getMergedRanges()` |

#### Utility Methods

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `toArray()` | Export sheet as 2D array | - | `any[][]` | `const data = sheet.toArray()` |
| `clear()` | Clear all sheet data | - | `Sheet` | `sheet.clear()` |
| `toJSON()` | Export sheet to JSON | - | `object` | `const data = sheet.toJSON()` |

### 🔤 Cell Class

| Method | Description | Parameters | Returns | Example |
|--------|-------------|------------|---------|---------|
| `setValue(value)` | Set cell value | `value: any` | `Cell` | `cell.setValue('Hello')` |
| `getValue()` | Get cell value | - | `any` | `const value = cell.getValue()` |
| `setFormula(formula)` | Set cell formula | `formula: string` | `Cell` | `cell.setFormula('=A1+B1')` |
| `getFormula()` | Get cell formula | - | `string \| null` | `const formula = cell.getFormula()` |
| `setStyle(styleObj)` | Apply cell styling | `styleObj: object` | `Cell` | `cell.setStyle({ fontWeight: 'bold' })` |
| `getStyle()` | Get cell styles | - | `object` | `const styles = cell.getStyle()` |
| `clone()` | Clone cell with all properties | - | `Cell` | `const newCell = cell.clone()` |

---

## 🔧 Advanced Usage

### 🎨 Cell Styling

// Get a cell and apply multiple styles
const headerCell = sheet.getCell(0, 0);
headerCell.setStyle({
fontWeight: 'bold',
backgroundColor: '#4CAF50',
color: 'white',
textAlign: 'center',
fontSize: '14px',
border: '1px solid #ddd'
});

// Chain operations
headerCell.setValue('TOTAL SALES')
.setStyle({ fontWeight: 'bold' })
.setStyle({ backgroundColor: '#f0f0f0' });

text

### 📐 Formula Handling

// Set formulas
sheet.setCellFormula(5, 0, '=A1+A2+A3');
sheet.setCellFormula(5, 1, '=SUM(B1:B4)');
sheet.setCellFormula(5, 2, '=AVERAGE(C1:C4)');

// Get formulas
const formula = sheet.getCell(5, 0).getFormula();
console.log(formula); // '=A1+A2+A3'

text

### 💾 Data Persistence

// Save to localStorage (Browser)
const workbookData = workbook.toJSON();
localStorage.setItem('mySpreadsheet', JSON.stringify(workbookData));

// Load from localStorage
const savedData = JSON.parse(localStorage.getItem('mySpreadsheet'));
const newWorkbook = new Spreadsheet();
newWorkbook.fromJSON(savedData);

// Save to file (Node.js)
const fs = require('fs');
fs.writeFileSync('spreadsheet.json', JSON.stringify(workbookData, null, 2));

// Load from file (Node.js)
const fileData = JSON.parse(fs.readFileSync('spreadsheet.json', 'utf8'));
workbook.fromJSON(fileData);

text

### 🔍 Search and Find

// Search across all sheets
const searchResults = workbook.findInWorkbook('Widget', {
caseSensitive: false
});

console.log('Search Results:');
searchResults.forEach(result => {
console.log(Found "${result.value}" in ${result.sheet} at ${result.coordinate});
});

// Output:
// Found "Widget A" in Products at Products!A2
// Found "Widget B" in Products at Products!A3

text

### 📊 Advanced Data Operations

// Complex data merging example
const productSheet = workbook.getSheet('Products');

// Merge all product names with custom separator
const allProducts = productSheet.mergeColumnData(0, ' - ');
console.log('All Products:', allProducts);
// Output: "Product - Widget A - Widget B - Widget C"

// Merge specific row with different separator
const productDetails = productSheet.mergeRowData(1, ' | ');
console.log('Product Details:', productDetails);
// Output: "Widget A | 25.99 | 150 | In Stock"

// Concatenate data from multiple sources
productSheet.concatenateWithRow(1, 5, 0, ' = SUMMARY: ');
productSheet.concatenateWithCell(1, 0, 2, 0, ' & ');

// Complex range merging
const quarterData = productSheet.mergeRangeData(1, 1, 1, 3, ' + ');
console.log('Quarter Total:', quarterData);

text

---

## 🌐 Framework Integration

### ⚛️ React Integration

import React, { useState, useEffect } from 'react';
import { Spreadsheet } from 'pure-spreadsheet-js';

function SpreadsheetManager() {
const [workbook, setWorkbook] = useState(null);
const [activeSheetData, setActiveSheetData] = useState([]);
const [sheetNames, setSheetNames] = useState([]);

useEffect(() => {
// Initialize workbook
const wb = new Spreadsheet();
wb.createSheet('Sales Data');

text
const sheet = wb.getActiveSheet();
sheet.setHeaders(['Product', 'Revenue', 'Units']);
sheet.setCell(1, 0, 'Widget A');
sheet.setCell(1, 1, 15000);
sheet.setCell(1, 2, 150);

setWorkbook(wb);
setActiveSheetData(sheet.toArray());
setSheetNames(wb.getSheetNames());
}, []);

const addNewSheet = () => {
if (workbook) {
const newSheetName = Sheet ${workbook.getSheetCount() + 1};
workbook.addSheetToWorkbook(newSheetName);
setSheetNames(workbook.getSheetNames());
}
};

const switchSheet = (sheetName) => {
if (workbook) {
workbook.setActiveSheet(sheetName);
setActiveSheetData(workbook.getActiveSheet().toArray());
}
};

return (
<div>
<h2>Spreadsheet Manager</h2>

text
  {/* Sheet tabs */}
  <div style={{ marginBottom: '10px' }}>
    {sheetNames.map(name => (
      <button key={name} onClick={() => switchSheet(name)}>
        {name}
      </button>
    ))}
    <button onClick={addNewSheet}>+ Add Sheet</button>
  </div>

  {/* Data table */}
  <table border="1" style={{ borderCollapse: 'collapse' }}>
    <tbody>
      {activeSheetData.map((row, i) => (
        <tr key={i}>
          {row.map((cell, j) => (
            <td key={j} style={{ padding: '8px' }}>
              {cell}
            </td>
          ))}
        </tr>
      ))}
    </tbody>
  </table>
</div>
);
}

export default SpreadsheetManager;

text

### 🟢 Vue.js Integration

<template> <div class="spreadsheet-component"> <h2>My Spreadsheet App</h2>
text
<!-- Sheet Navigation -->
<div class="sheet-tabs">
  <button 
    v-for="sheet in sheetNames" 
    :key="sheet"
    @click="switchToSheet(sheet)"
    :class="{ active: sheet === activeSheet }"
  >
    {{ sheet }}
  </button>
  <button @click="addNewSheet" class="add-sheet">+ Add Sheet</button>
</div>

<!-- Data Display -->
<table class="data-table">
  <tr v-for="(row, i) in currentSheetData" :key="i">
    <td v-for="(cell, j) in row" :key="j">
      {{ cell }}
    </td>
  </tr>
</table>

<!-- Actions -->
<div class="actions">
  <button @click="exportData">Export JSON</button>
  <button @click="mergeSampleData">Merge Sample Data</button>
</div>
</div> </template> <script> import { Spreadsheet } from 'pure-spreadsheet-js'; export default { name: 'SpreadsheetComponent', data() { return { workbook: null, currentSheetData: [], sheetNames: [], activeSheet: null }; }, mounted() { this.initializeSpreadsheet(); }, methods: { initializeSpreadsheet() { this.workbook = new Spreadsheet(); this.workbook.createSheet('Products'); const sheet = this.workbook.getActiveSheet(); sheet.setHeaders(['Product', 'Price', 'Stock', 'Status']); sheet.insertMultipleRows(1, 3); // Add sample data sheet.setCell(1, 0, 'Laptop'); sheet.setCell(1, 1, '$999'); sheet.setCell(1, 2, 15); sheet.setCell(1, 3, 'Available'); this.updateView(); }, updateView() { this.currentSheetData = this.workbook.getActiveSheet().toArray(); this.sheetNames = this.workbook.getSheetNames(); this.activeSheet = this.workbook.activeSheetName; }, switchToSheet(sheetName) { this.workbook.setActiveSheet(sheetName); this.updateView(); }, addNewSheet() { const newName = `Sheet ${this.workbook.getSheetCount() + 1}`; this.workbook.addSheetToWorkbook(newName, { headers: ['Column A', 'Column B', 'Column C'] }); this.updateView(); }, exportData() { const data = this.workbook.toJSON(); console.log('Exported data:', data); // You can download or save the data here }, mergeSampleData() { const sheet = this.workbook.getActiveSheet(); const mergedRow = sheet.mergeRowData(1, ' | '); alert(`Merged data: ${mergedRow}`); } } }; </script> <style scoped> .sheet-tabs button { margin-right: 5px; padding: 8px 12px; border: 1px solid #ddd; background: #f5f5f5; cursor: pointer; } .sheet-tabs button.active { background: #007bff; color: white; } .data-table { border-collapse: collapse; margin: 20px 0; } .data-table td { border: 1px solid #ddd; padding: 8px; min-width: 100px; } .actions button { margin-right: 10px; padding: 10px 15px; background: #28a745; color: white; border: none; cursor: pointer; } </style>
text

---

## 📱 Browser Support

| Browser | Version | Notes |
|---------|---------|-------|
| **Chrome** | 60+ | ✅ Full support |
| **Firefox** | 55+ | ✅ Full support |
| **Safari** | 11+ | ✅ Full support |
| **Edge** | 79+ | ✅ Full support |
| **Internet Explorer** | ❌ | Not supported |

---

## 🤝 Contributing

We welcome contributions from the community! Here's how you can help:

### 🐛 Bug Reports
Found a bug? [Open an issue](https://github.com/akhil028/spreadsheetActions/issues) with:
- Clear description of the problem
- Steps to reproduce
- Expected vs actual behavior
- Code example (if applicable)
- Browser/Node.js version

### 💡 Feature Requests
Have an idea? [Create a feature request](https://github.com/akhil028/spreadsheetActions/issues) with:
- Use case description
- Proposed API design
- Benefits to other users
- Implementation suggestions

### 🔧 Code Contributions
1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Write tests for new functionality
4. Ensure all tests pass: `npm test`
5. Follow the existing code style
6. Submit a pull request with clear description

### 📝 Documentation
Help improve documentation by:
- Fixing typos or unclear sections
- Adding more examples
- Translating to other languages
- Creating tutorials or guides

---

## 📄 License

**MIT License** - You're free to use this library in commercial and personal projects.

See the [LICENSE](https://github.com/akhil028/spreadsheetActions/blob/main/LICENSE) file for full details.

---

## 🙋‍♂️ Support & Community

### 📧 Get Help
- **Email Support**: [your.email@example.com](mailto:your.email@example.com)
- **GitHub Issues**: [Report bugs or request features](https://github.com/akhil028/spreadsheetActions/issues)
- **GitHub Discussions**: [Community Q&A](https://github.com/akhil028/spreadsheetActions/discussions)

### 📚 Resources
- **📖 Full Documentation**: [GitHub Repository](https://github.com/akhil028/spreadsheetActions)
- **📦 NPM Package**: [pure-spreadsheet-js](https://npmjs.org/package/pure-spreadsheet-js)
- **🔗 Source Code**: [GitHub Repository](https://github.com/akhil028/spreadsheetActions)
- **📋 Changelog**: [Release Notes](https://github.com/akhil028/spreadsheetActions/releases)

### 🌟 Show Your Support
If this library helped you, please:
- ⭐ **Star the repository** on GitHub
- 📢 **Share with your network**
- 💬 **Leave feedback** in discussions
- 🐛 **Report issues** you find
- 🤝 **Contribute** improvements

---

<div align="center">

**Made with ❤️ by developers, for developers**

⭐ **Star this repo if it helped you!** ⭐

[📦 NPM Package](https://npmjs.org/package/pure-spreadsheet-js) • [🔗 GitHub](https://github.com/akhil028/spreadsheetActions) • [📖 Documentation](https://github.com/akhil028/spreadsheetActions#readme)

**Version 1.0.0** • **Zero Dependencies** • **MIT Licensed**

</div>