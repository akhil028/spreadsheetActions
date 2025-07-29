const { Spreadsheet } = require('../src');

// Create and populate spreadsheet
const spreadsheet = new Spreadsheet();

// Create sheets
spreadsheet.createSheet('Sales');
spreadsheet.createSheet('Inventory');

// Work with sales sheet
spreadsheet.setActiveSheet('Sales');
const salesSheet = spreadsheet.getActiveSheet();

// Add headers
salesSheet.setCell(0, 0, 'Product');
salesSheet.setCell(0, 1, 'Q1');
salesSheet.setCell(0, 2, 'Q2');
salesSheet.setCell(0, 3, 'Total');

// Add data
salesSheet.setCell(1, 0, 'Widget A');
salesSheet.setCell(1, 1, 1000);
salesSheet.setCell(1, 2, 1200);
salesSheet.setCellFormula(1, 3, '=B1+C1');

// Merge header cells
salesSheet.mergeCells(0, 1, 0, 2);

console.log('Sales data:');
console.log(salesSheet.toArray());

console.log('\nMerged ranges:');
console.log(salesSheet.getMergedRanges());

// Export to JSON
const data = spreadsheet.toJSON();
console.log('\nExported JSON structure available');
