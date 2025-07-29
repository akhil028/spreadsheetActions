declare class Cell {
  constructor(value?: any, formula?: string | null);
  setValue(value: any): Cell;
  getValue(): any;
  setFormula(formula: string): Cell;
  getFormula(): string | null;
  setStyle(style: object): Cell;
  getStyle(): object;
  clone(): Cell;
  toJSON(): object;
}

declare class Sheet {
  constructor(name: string);
  name: string;
  setCell(row: number, col: number, value: any): Sheet;
  getCell(row: number, col: number): Cell;
  setHeaders(headers: string[]): Sheet;
  getHeaders(): string[];
  insertMultipleRows(startIndex: number, count: number): Sheet;
  insertMultipleColumns(startIndex: number, count: number): Sheet;
  mergeCells(startRow: number, startCol: number, endRow: number, endCol: number): Sheet;
  mergeRowData(rowIndex: number, separator?: string): string;
  mergeColumnData(colIndex: number, separator?: string): string;
  concatenateWithRow(sourceRow: number, targetRow: number, targetCol: number, separator?: string): Sheet;
  concatenateWithColumn(sourceCol: number, targetRow: number, targetCol: number, separator?: string): Sheet;
  concatenateWithCell(sourceRow: number, sourceCol: number, targetRow: number, targetCol: number, separator?: string): Sheet;
  clearCell(row: number, col: number): Sheet;
  deleteCell(row: number, col: number): Sheet;
  toArray(): any[][];
  toJSON(): object;
}

declare class Spreadsheet {
  constructor();
  createSheet(name: string): Spreadsheet;
  addSheetToWorkbook(name: string, options?: object): Spreadsheet;
  getSheet(name: string): Sheet;
  getActiveSheet(): Sheet;
  setActiveSheet(name: string): Spreadsheet;
  getSheetNames(): string[];
  getSheetCount(): number;
  removeSheet(name: string): Spreadsheet;
  renameSheet(oldName: string, newName: string): Spreadsheet;
  cloneSheet(sourceName: string, targetName: string, position?: string): Spreadsheet;
  findInWorkbook(searchValue: any, options?: object): object[];
  getWorkbookInfo(): object;
  toJSON(): object;
  fromJSON(data: object): Spreadsheet;
  clear(): Spreadsheet;
}

export { Spreadsheet, Sheet, Cell };
export default Spreadsheet;
