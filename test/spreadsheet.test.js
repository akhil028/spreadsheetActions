const { Spreadsheet } = require('../src');

describe('Spreadsheet', () => {
  let spreadsheet;

  beforeEach(() => {
    spreadsheet = new Spreadsheet();
  });

  test('should create a new spreadsheet', () => {
    expect(spreadsheet).toBeInstanceOf(Spreadsheet);
    expect(spreadsheet.getSheetNames()).toHaveLength(0);
  });

  test('should create a new sheet', () => {
    spreadsheet.createSheet('Sheet1');
    expect(spreadsheet.getSheetNames()).toContain('Sheet1');
    expect(spreadsheet.activeSheetName).toBe('Sheet1');
  });

  test('should set and get cell values', () => {
    spreadsheet.createSheet('Sheet1');
    const sheet = spreadsheet.getActiveSheet();
    
    sheet.setCell(0, 0, 'Hello');
    sheet.setCell(0, 1, 'World');
    
    expect(sheet.getCell(0, 0).getValue()).toBe('Hello');
    expect(sheet.getCell(0, 1).getValue()).toBe('World');
  });

  test('should merge cells', () => {
    spreadsheet.createSheet('Sheet1');
    const sheet = spreadsheet.getActiveSheet();
    
    sheet.setCell(0, 0, 'Merged');
    sheet.mergeCells(0, 0, 1, 1);
    
    const mergedRanges = sheet.getMergedRanges();
    expect(mergedRanges).toHaveLength(1);
    expect(mergedRanges[0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1
    });
  });

  test('should throw error for overlapping merges', () => {
    spreadsheet.createSheet('Sheet1');
    const sheet = spreadsheet.getActiveSheet();
    
    sheet.mergeCells(0, 0, 1, 1);
    
    expect(() => {
      sheet.mergeCells(1, 1, 2, 2);
    }).toThrow('Cannot merge: range overlaps with existing merged cells');
  });

  test('should export and import JSON', () => {
    spreadsheet.createSheet('Sheet1');
    const sheet = spreadsheet.getActiveSheet();
    
    sheet.setCell(0, 0, 'Test');
    sheet.mergeCells(1, 1, 2, 2);
    
    const json = spreadsheet.toJSON();
    const newSpreadsheet = new Spreadsheet();
    newSpreadsheet.fromJSON(json);
    
    const importedSheet = newSpreadsheet.getActiveSheet();
    expect(importedSheet.getCell(0, 0).getValue()).toBe('Test');
    expect(importedSheet.getMergedRanges()).toHaveLength(1);
  });
});
