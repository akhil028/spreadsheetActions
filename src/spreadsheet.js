const Sheet = require('./sheet');
const Cell = require('./cell');

class Spreadsheet {
  constructor() {
    this.sheets = new Map();
    this.activeSheetName = null;
    this.sheetOrder = []; // Track sheet order for tabs
    this.metadata = {
      created: Date.now(),
      modified: Date.now(),
      version: '1.0.0',
      author: '',
      title: 'Untitled Workbook'
    };
  }

  // 1. Create a new sheet (workbook tab)
  createSheet(name) {
    if (this.sheets.has(name)) {
      throw new Error(`Sheet with name "${name}" already exists`);
    }

    const sheet = new Sheet(name);
    this.sheets.set(name, sheet);
    this.sheetOrder.push(name);
    
    // Set as active if it's the first sheet
    if (!this.activeSheetName) {
      this.activeSheetName = name;
    }
    
    this.metadata.modified = Date.now();
    return this;
  }

  // 2. Add new sheet to current workbook (this is the main method for point 2)
  addSheetToWorkbook(name, options = {}) {
    const {
      position = 'end', // 'start', 'end', 'after', 'before'
      referenceSheet = null,
      copyFrom = null, // Copy structure from existing sheet
      headers = null
    } = options;

    if (this.sheets.has(name)) {
      throw new Error(`Sheet with name "${name}" already exists`);
    }

    const sheet = new Sheet(name);

    // Copy structure from another sheet if specified
    if (copyFrom && this.sheets.has(copyFrom)) {
      const sourceSheet = this.sheets.get(copyFrom);
      sheet.rowCount = sourceSheet.rowCount;
      sheet.colCount = sourceSheet.colCount;
      
      // Copy headers only
      if (sourceSheet.metadata.hasHeaders) {
        const headers = sourceSheet.getHeaders();
        sheet.setHeaders(headers);
      }
    }

    // Set headers if provided
    if (headers) {
      sheet.setHeaders(headers);
    }

    this.sheets.set(name, sheet);

    // Handle sheet positioning
    this._insertSheetAtPosition(name, position, referenceSheet);
    
    this.metadata.modified = Date.now();
    return this;
  }

  // Helper method to handle sheet positioning
  _insertSheetAtPosition(sheetName, position, referenceSheet) {
    switch (position) {
      case 'start':
        this.sheetOrder.unshift(sheetName);
        break;
      case 'end':
        this.sheetOrder.push(sheetName);
        break;
      case 'after':
        if (referenceSheet && this.sheets.has(referenceSheet)) {
          const index = this.sheetOrder.indexOf(referenceSheet);
          this.sheetOrder.splice(index + 1, 0, sheetName);
        } else {
          this.sheetOrder.push(sheetName);
        }
        break;
      case 'before':
        if (referenceSheet && this.sheets.has(referenceSheet)) {
          const index = this.sheetOrder.indexOf(referenceSheet);
          this.sheetOrder.splice(index, 0, sheetName);
        } else {
          this.sheetOrder.unshift(sheetName);
        }
        break;
      default:
        this.sheetOrder.push(sheetName);
    }
  }

  // Get sheet by name
  getSheet(name) {
    const sheet = this.sheets.get(name);
    if (!sheet) {
      throw new Error(`Sheet "${name}" not found`);
    }
    return sheet;
  }

  // Remove sheet from workbook
  removeSheet(name) {
    if (!this.sheets.has(name)) {
      throw new Error(`Sheet "${name}" not found`);
    }

    if (this.sheets.size === 1) {
      throw new Error('Cannot remove the last sheet in workbook');
    }

    this.sheets.delete(name);
    
    // Remove from order
    const orderIndex = this.sheetOrder.indexOf(name);
    if (orderIndex > -1) {
      this.sheetOrder.splice(orderIndex, 1);
    }
    
    // Update active sheet if necessary
    if (this.activeSheetName === name) {
      this.activeSheetName = this.sheetOrder.length > 0 ? 
        this.sheetOrder[0] : null;
    }
    
    this.metadata.modified = Date.now();
    return this;
  }

  // Move sheet to different position
  moveSheet(sheetName, newPosition) {
    if (!this.sheets.has(sheetName)) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    const currentIndex = this.sheetOrder.indexOf(sheetName);
    if (currentIndex === -1) return this;

    // Remove from current position
    this.sheetOrder.splice(currentIndex, 1);

    // Insert at new position
    const targetIndex = Math.max(0, Math.min(newPosition, this.sheetOrder.length));
    this.sheetOrder.splice(targetIndex, 0, sheetName);

    this.metadata.modified = Date.now();
    return this;
  }

  // Get all sheet names in order
  getSheetNames() {
    return [...this.sheetOrder];
  }

  // Get sheet count
  getSheetCount() {
    return this.sheets.size;
  }

  // Set active sheet (switch tabs)
  setActiveSheet(name) {
    if (!this.sheets.has(name)) {
      throw new Error(`Sheet "${name}" not found`);
    }
    
    this.activeSheetName = name;
    return this;
  }

  // Get active sheet
  getActiveSheet() {
    if (!this.activeSheetName) {
      throw new Error('No active sheet');
    }
    return this.getSheet(this.activeSheetName);
  }

  // Rename sheet
  renameSheet(oldName, newName) {
    if (!this.sheets.has(oldName)) {
      throw new Error(`Sheet "${oldName}" not found`);
    }
    
    if (this.sheets.has(newName)) {
      throw new Error(`Sheet with name "${newName}" already exists`);
    }

    const sheet = this.sheets.get(oldName);
    sheet.name = newName;
    
    // Update in maps and order
    this.sheets.delete(oldName);
    this.sheets.set(newName, sheet);
    
    const orderIndex = this.sheetOrder.indexOf(oldName);
    if (orderIndex > -1) {
      this.sheetOrder[orderIndex] = newName;
    }
    
    if (this.activeSheetName === oldName) {
      this.activeSheetName = newName;
    }
    
    this.metadata.modified = Date.now();
    return this;
  }

  // Clone entire sheet (including data)
  cloneSheet(sourceName, targetName, position = 'after') {
    if (!this.sheets.has(sourceName)) {
      throw new Error(`Sheet "${sourceName}" not found`);
    }
    
    if (this.sheets.has(targetName)) {
      throw new Error(`Sheet with name "${targetName}" already exists`);
    }

    const sourceSheet = this.sheets.get(sourceName);
    const targetSheet = new Sheet(targetName);
    
    // Copy all cells
    for (const [key, cell] of sourceSheet.cells) {
      targetSheet.cells.set(key, cell.clone());
    }
    
    // Copy properties
    targetSheet.mergedRanges = [...sourceSheet.mergedRanges];
    targetSheet.rowCount = sourceSheet.rowCount;
    targetSheet.colCount = sourceSheet.colCount;
    targetSheet.metadata = { ...sourceSheet.metadata };
    targetSheet.metadata.created = Date.now();
    
    this.sheets.set(targetName, targetSheet);
    this._insertSheetAtPosition(targetName, position, sourceName);
    
    this.metadata.modified = Date.now();
    return this;
  }

  // Workbook-wide operations
  findInWorkbook(searchValue, options = {}) {
    const results = [];
    
    for (const sheetName of this.sheetOrder) {
      const sheet = this.sheets.get(sheetName);
      for (const [key, cell] of sheet.cells) {
        const [row, col] = key.split(',').map(Number);
        const cellValue = String(cell.getValue());
        
        let searchStr = String(searchValue);
        let compareValue = cellValue;
        
        if (!options.caseSensitive) {
          compareValue = compareValue.toLowerCase();
          searchStr = searchStr.toLowerCase();
        }
        
        if (compareValue.includes(searchStr)) {
          results.push({
            sheet: sheetName,
            row,
            col,
            value: cellValue,
            coordinate: `${sheetName}!${this._indicesToCoordinate(row, col)}`
          });
        }
      }
    }
    
    return results;
  }

  // Export workbook structure info
  getWorkbookInfo() {
    const sheetsInfo = this.sheetOrder.map(name => {
      const sheet = this.sheets.get(name);
      return {
        name: name,
        isActive: name === this.activeSheetName,
        rowCount: sheet.rowCount,
        colCount: sheet.colCount,
        hasHeaders: sheet.metadata.hasHeaders || false,
        cellCount: sheet.cells.size,
        mergedRanges: sheet.mergedRanges.length
      };
    });

    return {
      metadata: this.metadata,
      sheetCount: this.sheets.size,
      activeSheet: this.activeSheetName,
      sheets: sheetsInfo
    };
  }

  // Export to JSON (maintaining sheet order)
  toJSON() {
    const sheetsObj = {};
    for (const name of this.sheetOrder) {
      const sheet = this.sheets.get(name);
      sheetsObj[name] = sheet.toJSON();
    }

    return {
      sheets: sheetsObj,
      sheetOrder: this.sheetOrder,
      activeSheetName: this.activeSheetName,
      metadata: this.metadata
    };
  }

  // Import from JSON (preserving sheet order)
  fromJSON(data) {
    this.sheets.clear();
    this.sheetOrder = [];
    
    // Use sheet order if available, otherwise use object keys
    const orderedSheetNames = data.sheetOrder || Object.keys(data.sheets);
    
    for (const name of orderedSheetNames) {
      const sheetData = data.sheets[name];
      if (!sheetData) continue;
      
      const sheet = new Sheet(name);
      
      // Restore cells
      for (const [key, cellData] of Object.entries(sheetData.cells || {})) {
        const cell = new Cell(cellData.value, cellData.formula);
        cell.style = cellData.style || {};
        cell.isMerged = cellData.isMerged || false;
        cell.mergeInfo = cellData.mergeInfo || null;
        sheet.cells.set(key, cell);
      }
      
      sheet.mergedRanges = sheetData.mergedRanges || [];
      sheet.rowCount = sheetData.rowCount || 0;
      sheet.colCount = sheetData.colCount || 0;
      sheet.metadata = sheetData.metadata || { created: Date.now(), modified: Date.now() };
      
      this.sheets.set(name, sheet);
      this.sheetOrder.push(name);
    }
    
    this.activeSheetName = data.activeSheetName;
    this.metadata = data.metadata || { 
      created: Date.now(), 
      modified: Date.now(), 
      version: '1.0.0',
      author: '',
      title: 'Untitled Workbook'
    };
    
    return this;
  }

  // Helper method for coordinate conversion
  _indicesToCoordinate(row, col) {
    let colStr = '';
    let tempCol = col + 1;
    
    while (tempCol > 0) {
      tempCol -= 1;
      colStr = String.fromCharCode('A'.charCodeAt(0) + (tempCol % 26)) + colStr;
      tempCol = Math.floor(tempCol / 26);
    }
    
    return colStr + (row + 1);
  }

  // Clear all data
  clear() {
    this.sheets.clear();
    this.sheetOrder = [];
    this.activeSheetName = null;
    this.metadata.modified = Date.now();
    return this;
  }
}

module.exports = Spreadsheet;
