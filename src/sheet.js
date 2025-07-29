const Cell = require('./cell');
const { validateCoordinates } = require('./utils/helpers');

class Sheet {
  constructor(name) {
    this.name = name;
    this.cells = new Map(); // Uses "row,col" as key
    this.mergedRanges = [];
    this.rowCount = 0;
    this.colCount = 0;
    this.metadata = {
      created: new Date(),
      modified: new Date()
    };
  }

  // Get cell at specific coordinates
  getCell(row, col) {
    validateCoordinates(row, col);
    const key = `${row},${col}`;
    return this.cells.get(key) || new Cell();
  }

  // Set cell value
  setCell(row, col, value) {
    validateCoordinates(row, col);
    const key = `${row},${col}`;
    let cell = this.cells.get(key);
    
    if (!cell) {
      cell = new Cell();
      this.cells.set(key, cell);
    }
    
    cell.setValue(value);
    this._updateDimensions(row, col);
    this.metadata.modified = new Date();
    
    return this;
  }

  // Set cell formula
  setCellFormula(row, col, formula) {
    validateCoordinates(row, col);
    const key = `${row},${col}`;
    let cell = this.cells.get(key);
    
    if (!cell) {
      cell = new Cell();
      this.cells.set(key, cell);
    }
    
    cell.setFormula(formula);
    this._updateDimensions(row, col);
    this.metadata.modified = new Date();
    
    return this;
  }

  // Merge cells in a range
  mergeCells(startRow, startCol, endRow, endCol) {
    validateCoordinates(startRow, startCol);
    validateCoordinates(endRow, endCol);

    if (startRow > endRow || startCol > endCol) {
      throw new Error('Invalid merge range: start coordinates must be less than or equal to end coordinates');
    }

    // Check for overlapping merges
    if (this._hasOverlappingMerge(startRow, startCol, endRow, endCol)) {
      throw new Error('Cannot merge: range overlaps with existing merged cells');
    }

    const mergeInfo = { startRow, startCol, endRow, endCol };
    
    // Mark all cells in range as merged
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const key = `${row},${col}`;
        let cell = this.cells.get(key);
        
        if (!cell) {
          cell = new Cell();
          this.cells.set(key, cell);
        }
        
        cell.setMerged(mergeInfo);
      }
    }

    this.mergedRanges.push(mergeInfo);
    this.metadata.modified = new Date();
    
    return this;
  }

  // Unmerge cells
  unmergeCells(startRow, startCol, endRow, endCol) {
    const mergeIndex = this.mergedRanges.findIndex(range => 
      range.startRow === startRow && 
      range.startCol === startCol && 
      range.endRow === endRow && 
      range.endCol === endCol
    );

    if (mergeIndex === -1) {
      throw new Error('No merged range found at specified coordinates');
    }

    // Clear merge info from all cells in range
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const key = `${row},${col}`;
        const cell = this.cells.get(key);
        if (cell) {
          cell.clearMerge();
        }
      }
    }

    this.mergedRanges.splice(mergeIndex, 1);
    this.metadata.modified = new Date();
    
    return this;
  }

  // Get all merged ranges
  getMergedRanges() {
    return [...this.mergedRanges];
  }

  // Insert row
  insertRow(rowIndex) {
    if (rowIndex < 0) throw new Error('Row index must be non-negative');

    const newCells = new Map();
    
    // Shift existing cells down
    for (const [key, cell] of this.cells) {
      const [row, col] = key.split(',').map(Number);
      if (row >= rowIndex) {
        newCells.set(`${row + 1},${col}`, cell);
      } else {
        newCells.set(key, cell);
      }
    }

    this.cells = newCells;
    this.rowCount++;
    this._updateMergedRangesAfterRowInsert(rowIndex);
    this.metadata.modified = new Date();
    
    return this;
  }

  // Insert column
  insertColumn(colIndex) {
    if (colIndex < 0) throw new Error('Column index must be non-negative');

    const newCells = new Map();
    
    // Shift existing cells right
    for (const [key, cell] of this.cells) {
      const [row, col] = key.split(',').map(Number);
      if (col >= colIndex) {
        newCells.set(`${row},${col + 1}`, cell);
      } else {
        newCells.set(key, cell);
      }
    }

    this.cells = newCells;
    this.colCount++;
    this._updateMergedRangesAfterColInsert(colIndex);
    this.metadata.modified = new Date();
    
    return this;
  }

  // 3. Add multiple rows/columns at once
  insertMultipleRows(startIndex, count = 1) {
    if (startIndex < 0 || count < 1) {
      throw new Error('Invalid parameters for row insertion');
    }
    
    for (let i = 0; i < count; i++) {
      this.insertRow(startIndex);
    }
    return this;
  }

  insertMultipleColumns(startIndex, count = 1) {
    if (startIndex < 0 || count < 1) {
      throw new Error('Invalid parameters for column insertion');
    }
    
    for (let i = 0; i < count; i++) {
      this.insertColumn(startIndex);
    }
    return this;
  }

  // 4. Header management
  setHeaders(headers) {
    if (!Array.isArray(headers)) {
      throw new Error('Headers must be an array');
    }
    
    for (let i = 0; i < headers.length; i++) {
      this.setCell(0, i, headers[i]);
      // Style headers differently
      this.getCell(0, i).setStyle({
        fontWeight: 'bold',
        backgroundColor: '#f0f0f0',
        borderBottom: '2px solid #333'
      });
    }
    
    this.metadata.hasHeaders = true;
    this.metadata.modified = Date.now();
    return this;
  }

  getHeaders() {
    const headers = [];
    for (let col = 0; col < this.colCount; col++) {
      headers.push(this.getCell(0, col).getValue());
    }
    return headers;
  }

  // 5. Delete data from cells
  clearCell(row, col) {
    validateCoordinates(row, col);
    const key = `${row},${col}`;
    const cell = this.cells.get(key);
    if (cell) {
      cell.setValue('');
      cell.setFormula(null);
    }
    this.metadata.modified = Date.now();
    return this;
  }

  clearRange(startRow, startCol, endRow, endCol) {
    validateCoordinates(startRow, startCol);
    validateCoordinates(endRow, endCol);
    
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        this.clearCell(row, col);
      }
    }
    return this;
  }

  deleteCell(row, col) {
    validateCoordinates(row, col);
    const key = `${row},${col}`;
    this.cells.delete(key);
    this.metadata.modified = Date.now();
    return this;
  }

  // 6. Merge rows/columns with separator (combine data, not cells)
  mergeRowData(rowIndex, separator = ' ') {
    if (rowIndex < 0 || rowIndex >= this.rowCount) {
      throw new Error('Invalid row index');
    }
    
    const values = [];
    for (let col = 0; col < this.colCount; col++) {
      const cellValue = this.getCell(rowIndex, col).getValue();
      if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
        values.push(String(cellValue));
      }
    }
    
    return values.join(separator);
  }

  mergeColumnData(colIndex, separator = ' ') {
    if (colIndex < 0 || colIndex >= this.colCount) {
      throw new Error('Invalid column index');
    }
    
    const values = [];
    for (let row = 0; row < this.rowCount; row++) {
      const cellValue = this.getCell(row, colIndex).getValue();
      if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
        values.push(String(cellValue));
      }
    }
    
    return values.join(separator);
  }

  mergeRangeData(startRow, startCol, endRow, endCol, separator = ' ') {
    validateCoordinates(startRow, startCol);
    validateCoordinates(endRow, endCol);
    
    const values = [];
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cellValue = this.getCell(row, col).getValue();
        if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
          values.push(String(cellValue));
        }
      }
    }
    
    return values.join(separator);
  }

  // 7. Concatenate data utilities
  concatenateWithRow(sourceRow, targetRow, targetCol, separator = ' ') {
    const mergedData = this.mergeRowData(sourceRow, separator);
    this.setCell(targetRow, targetCol, mergedData);
    return this;
  }

  concatenateWithColumn(sourceCol, targetRow, targetCol, separator = ' ') {
    const mergedData = this.mergeColumnData(sourceCol, separator);
    this.setCell(targetRow, targetCol, mergedData);
    return this;
  }

  concatenateWithCell(sourceRow, sourceCol, targetRow, targetCol, separator = ' ') {
    const sourceValue = this.getCell(sourceRow, sourceCol).getValue();
    const targetValue = this.getCell(targetRow, targetCol).getValue();
    
    const concatenated = [targetValue, sourceValue]
      .filter(val => val !== '' && val !== null && val !== undefined)
      .join(separator);
    
    this.setCell(targetRow, targetCol, concatenated);
    return this;
  }


  // Get data as 2D array
  toArray() {
    const result = [];
    
    for (let row = 0; row < this.rowCount; row++) {
      const rowData = [];
      for (let col = 0; col < this.colCount; col++) {
        const cell = this.getCell(row, col);
        rowData.push(cell.getValue());
      }
      result.push(rowData);
    }
    
    return result;
  }

  // Clear all data
  clear() {
    this.cells.clear();
    this.mergedRanges = [];
    this.rowCount = 0;
    this.colCount = 0;
    this.metadata.modified = new Date();
    return this;
  }

  // Private helper methods
  _updateDimensions(row, col) {
    this.rowCount = Math.max(this.rowCount, row + 1);
    this.colCount = Math.max(this.colCount, col + 1);
  }

  _hasOverlappingMerge(startRow, startCol, endRow, endCol) {
    return this.mergedRanges.some(range => {
      return !(endRow < range.startRow || 
               startRow > range.endRow || 
               endCol < range.startCol || 
               startCol > range.endCol);
    });
  }

  _updateMergedRangesAfterRowInsert(insertIndex) {
    this.mergedRanges.forEach(range => {
      if (range.startRow >= insertIndex) range.startRow++;
      if (range.endRow >= insertIndex) range.endRow++;
    });
  }

  _updateMergedRangesAfterColInsert(insertIndex) {
    this.mergedRanges.forEach(range => {
      if (range.startCol >= insertIndex) range.startCol++;
      if (range.endCol >= insertIndex) range.endCol++;
    });
  }

  // Convert to JSON
  toJSON() {
    const cellsObj = {};
    for (const [key, cell] of this.cells) {
      cellsObj[key] = cell.toJSON();
    }

    return {
      name: this.name,
      cells: cellsObj,
      mergedRanges: this.mergedRanges,
      rowCount: this.rowCount,
      colCount: this.colCount,
      metadata: this.metadata
    };
  }
}

module.exports = Sheet;
