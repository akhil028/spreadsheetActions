const { Cell } = require('../src');

describe('Cell', () => {
  test('should create empty cell', () => {
    const cell = new Cell();
    expect(cell.getValue()).toBe('');
    expect(cell.getFormula()).toBeNull();
  });

  test('should set and get value', () => {
    const cell = new Cell();
    cell.setValue('Hello');
    expect(cell.getValue()).toBe('Hello');
  });

  test('should set and get formula', () => {
    const cell = new Cell();
    cell.setFormula('=A1+B1');
    expect(cell.getFormula()).toBe('=A1+B1');
  });

  test('should clone cell correctly', () => {
    const original = new Cell('Test', '=SUM(A1:A10)');
    original.setStyle({ color: 'red' });
    
    const clone = original.clone();
    expect(clone.getValue()).toBe('Test');
    expect(clone.getFormula()).toBe('=SUM(A1:A10)');
    expect(clone.getStyle()).toEqual({ color: 'red' });
  });
});
