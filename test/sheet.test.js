const { Sheet } = require('../src');

describe('Sheet', () => {
  let sheet;

  beforeEach(() => {
    sheet = new Sheet('TestSheet');
  });

  test('should insert rows correctly', () => {
    sheet.setCell(0, 0, 'A1');
    sheet.setCell(1, 0, 'A2');
    
    sheet.insertRow(1);
    
    expect(sheet.getCell(0, 0).getValue()).toBe('A1');
    expect(sheet.getCell(1, 0).getValue()).toBe('');
    expect(sheet.getCell(2, 0).getValue()).toBe('A2');
  });

  test('should insert columns correctly', () => {
    sheet.setCell(0, 0, 'A1');
    sheet.setCell(0, 1, 'B1');
    
    sheet.insertColumn(1);
    
    expect(sheet.getCell(0, 0).getValue()).toBe('A1');
    expect(sheet.getCell(0, 1).getValue()).toBe('');
    expect(sheet.getCell(0, 2).getValue()).toBe('B1');
  });

  test('should convert to array correctly', () => {
    sheet.setCell(0, 0, 'A1');
    sheet.setCell(0, 1, 'B1');
    sheet.setCell(1, 0, 'A2');
    sheet.setCell(1, 1, 'B2');
    
    const array = sheet.toArray();
    expect(array).toEqual([
      ['A1', 'B1'],
      ['A2', 'B2']
    ]);
  });
});
