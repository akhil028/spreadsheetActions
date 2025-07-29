const { Spreadsheet } = require('../src');

console.log('=== Excel-like Workbook Management ===\n');

// Create a new workbook
const workbook = new Spreadsheet();

// 1. Create initial sheet
workbook.createSheet('Sales Data');
console.log('Created initial sheet: Sales Data');

// 2. Add multiple sheets to the workbook (this is the clarified point 2)
workbook.addSheetToWorkbook('Inventory', {
  headers: ['Item', 'Stock', 'Location', 'Reorder Level']
});

workbook.addSheetToWorkbook('Employee Data', {
  position: 'after',
  referenceSheet: 'Sales Data',
  headers: ['Name', 'Department', 'Salary', 'Start Date']
});

workbook.addSheetToWorkbook('Reports', {
  position: 'end'
});

console.log('Added sheets to workbook:', workbook.getSheetNames());
console.log('Total sheets in workbook:', workbook.getSheetCount());

// Work with Sales Data sheet
workbook.setActiveSheet('Sales Data');
const salesSheet = workbook.getActiveSheet();

// 4. Add headers
salesSheet.setHeaders(['Product', 'Q1 Sales', 'Q2 Sales', 'Total', 'Notes']);

// 3. Add multiple rows and columns
salesSheet.insertMultipleRows(1, 3); // Add 3 data rows
salesSheet.insertMultipleColumns(5, 1); // Add 1 additional column

// Add sample data
salesSheet.setCell(1, 0, 'Widget A');
salesSheet.setCell(1, 1, 1000);
salesSheet.setCell(1, 2, 1200);
salesSheet.setCell(1, 4, 'Excellent performance');

salesSheet.setCell(2, 0, 'Widget B');
salesSheet.setCell(2, 1, 800);
salesSheet.setCell(2, 2, 900);
salesSheet.setCell(2, 4, 'Good growth');

salesSheet.setCell(3, 0, 'Widget C');
salesSheet.setCell(3, 1, 600);
salesSheet.setCell(3, 2, 750);
salesSheet.setCell(3, 4, 'Improving');

console.log('\nSales Data:');
console.log(salesSheet.toArray());

// 6. Merge row data with separator
const productSummary = salesSheet.mergeRowData(1, ' | ');
console.log('\nProduct A Summary:', productSummary);

// 7. Concatenate data examples
salesSheet.concatenateWithRow(2, 4, 0, ' - '); // Concatenate Widget B data
console.log('Concatenated data in cell (4,0):', salesSheet.getCell(4, 0).getValue());

// Work with Employee Data sheet
workbook.setActiveSheet('Employee Data');
const empSheet = workbook.getActiveSheet();

// Add employee data
empSheet.setCell(1, 0, 'John Doe');
empSheet.setCell(1, 1, 'Engineering');
empSheet.setCell(1, 2, 75000);
empSheet.setCell(1, 3, '2023-01-15');

empSheet.setCell(2, 0, 'Jane Smith');
empSheet.setCell(2, 1, 'Marketing');
empSheet.setCell(2, 2, 65000);
empSheet.setCell(2, 3, '2023-03-20');

console.log('\nEmployee Data:');
console.log(empSheet.toArray());

// 5. Delete data from specific cells
empSheet.clearCell(2, 2); // Clear Jane's salary
console.log('Cleared salary for Jane Smith');

// Move sheets around
workbook.moveSheet('Reports', 1); // Move Reports to second position
console.log('\nSheet order after moving Reports:', workbook.getSheetNames());

// Clone a sheet
workbook.cloneSheet('Sales Data', 'Sales Data Copy');
console.log('Cloned Sales Data sheet');

// Workbook information
console.log('\nWorkbook Information:');
const workbookInfo = workbook.getWorkbookInfo();
console.log(JSON.stringify(workbookInfo, null, 2));

// Search across all sheets
const searchResults = workbook.findInWorkbook('Widget');
console.log('\nSearch results for "Widget":');
searchResults.forEach(result => {
  console.log(`  - ${result.coordinate}: ${result.value}`);
});

// Export and save workbook
const workbookData = workbook.toJSON();
console.log('\nWorkbook exported to JSON (structure available)');

// Demonstrate importing
const newWorkbook = new Spreadsheet();
newWorkbook.fromJSON(workbookData);
console.log('Workbook imported successfully');
console.log('Imported sheets:', newWorkbook.getSheetNames());
