<div align="center">

# 📊 Pure Spreadsheet JS

### A powerful, lightweight JavaScript spreadsheet library

[![npm version](https://badge.fury.io/js/pure-spreadsheet-js.svg)](https://badge.fury.io/js/pure-spreadsheet-js)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://img.shields.io/npm/dm/pure-spreadsheet-js.svg)](https://npmjs.org/package/pure-spreadsheet-js)

</div>

---

### ✨ What is Pure Spreadsheet JS?

Pure Spreadsheet JS is a comprehensive JavaScript library that brings Excel-like functionality to your web applications. Built with **zero dependencies**, it provides all the essential spreadsheet features you need including multiple worksheets, cell merging, data manipulation, and much more.

### 🎯 Perfect For
- **📈 Business Dashboards** - Create interactive data visualization tools
- **💼 Enterprise Applications** - Build CRM, ERP, and inventory management systems
- **📊 Financial Tools** - Develop budgeting, forecasting, and accounting software
- **🎓 Educational Platforms** - Create grade books and learning management tools
- **📋 Data Management** - Build dynamic forms and data entry applications

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
