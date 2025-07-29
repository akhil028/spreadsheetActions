const { Spreadsheet } = require('../src');

describe('Workbook Management', () => {
  let workbook;

  beforeEach(() => {
    workbook = new Spreadsheet();
  });

  afterEach(() => {
    if (workbook) {
      workbook.clear();
      workbook = null;
    }
  });

  test('should create and manage multiple sheets', () => {
    // 1. Create initial sheet
    workbook.createSheet('Sheet1');
    expect(workbook.getSheetNames()).toEqual(['Sheet1']);

    // 2. Add sheets to workbook
    workbook.addSheetToWorkbook('Sheet2');
    workbook.addSheetToWorkbook('Sheet3', { position: 'start' });
    
    expect(workbook.getSheetNames()).toEqual(['Sheet3', 'Sheet1', 'Sheet2']);
    expect(workbook.getSheetCount()).toBe(3);
  });

  test('should add sheet with different positions', () => {
    workbook.createSheet('A');
    workbook.addSheetToWorkbook('B');
    workbook.addSheetToWorkbook('C', { position: 'after', referenceSheet: 'A' });
    workbook.addSheetToWorkbook('D', { position: 'before', referenceSheet: 'B' });
    
    expect(workbook.getSheetNames()).toEqual(['A', 'C', 'D', 'B']);
  });

  test('should add sheet with headers and copy structure', () => {
    workbook.createSheet('Template');
    const template = workbook.getActiveSheet();
    template.setHeaders(['Name', 'Age', 'City']);
    
    workbook.addSheetToWorkbook('Copy', { 
      copyFrom: 'Template',
      headers: ['Product', 'Price', 'Stock']
    });
    
    const copy = workbook.getSheet('Copy');
    expect(copy.getHeaders()).toEqual(['Product', 'Price', 'Stock']);
  });

  test('should move sheets in workbook', () => {
    workbook.createSheet('A');
    workbook.addSheetToWorkbook('B');
    workbook.addSheetToWorkbook('C');
    
    workbook.moveSheet('C', 0); // Move C to first position
    expect(workbook.getSheetNames()).toEqual(['C', 'A', 'B']);
  });

  test('should not allow removing last sheet', () => {
    workbook.createSheet('OnlySheet');
    
    expect(() => {
      workbook.removeSheet('OnlySheet');
    }).toThrow('Cannot remove the last sheet in workbook');
  });

  test('should clone sheet with all data', () => {
    workbook.createSheet('Original');
    const sheet = workbook.getActiveSheet();
    
    sheet.setHeaders(['Col1', 'Col2']);
    sheet.setCell(1, 0, 'Data1');
    sheet.setCell(1, 1, 'Data2');
    sheet.mergeCells(2, 0, 2, 1);
    
    workbook.cloneSheet('Original', 'Clone');
    const clone = workbook.getSheet('Clone');
    
    expect(clone.getHeaders()).toEqual(['Col1', 'Col2']);
    expect(clone.getCell(1, 0).getValue()).toBe('Data1');
    expect(clone.getMergedRanges()).toHaveLength(1);
  });

  test('should search across all sheets', () => {
    workbook.createSheet('Sheet1');
    workbook.addSheetToWorkbook('Sheet2');
    
    workbook.getSheet('Sheet1').setCell(0, 0, 'Test Value');
    workbook.getSheet('Sheet2').setCell(1, 1, 'Another Test');
    
    const results = workbook.findInWorkbook('Test');
    expect(results).toHaveLength(2);
    expect(results[0].sheet).toBe('Sheet1');
    expect(results[1].sheet).toBe('Sheet2');
  });

  test('should maintain sheet order in JSON export/import', () => {
    workbook.createSheet('First');
    workbook.addSheetToWorkbook('Second');
    workbook.addSheetToWorkbook('Third', { position: 'start' });
    
    const originalOrder = workbook.getSheetNames();
    const json = workbook.toJSON();
    
    const newWorkbook = new Spreadsheet();
    newWorkbook.fromJSON(json);
    
    expect(newWorkbook.getSheetNames()).toEqual(originalOrder);
  });
});
