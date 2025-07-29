// Validate row and column coordinates
function validateCoordinates(row, col) {
  if (typeof row !== 'number' || typeof col !== 'number') {
    throw new Error('Row and column must be numbers');
  }
  
  if (row < 0 || col < 0) {
    throw new Error('Row and column must be non-negative');
  }
  
  if (!Number.isInteger(row) || !Number.isInteger(col)) {
    throw new Error('Row and column must be integers');
  }
}

// Convert Excel-style coordinate (e.g., "A1") to indices
function coordinateToIndices(coordinate) {
  const match = coordinate.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error('Invalid coordinate format. Use format like "A1", "B2", etc.');
  }
  
  const colStr = match[1];
  const rowStr = match[2];
  
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  col -= 1; // Convert to 0-based index
  
  const row = parseInt(rowStr) - 1; // Convert to 0-based index
  
  return { row, col };
}

// Convert indices to Excel-style coordinate
function indicesToCoordinate(row, col) {
  validateCoordinates(row, col);
  
  let colStr = '';
  let tempCol = col + 1; // Convert to 1-based
  
  while (tempCol > 0) {
    tempCol -= 1;
    colStr = String.fromCharCode('A'.charCodeAt(0) + (tempCol % 26)) + colStr;
    tempCol = Math.floor(tempCol / 26);
  }
  
  return colStr + (row + 1); // Convert row to 1-based
}

// Parse range string (e.g., "A1:B2") to coordinates
function parseRange(rangeStr) {
  const parts = rangeStr.split(':');
  if (parts.length !== 2) {
    throw new Error('Invalid range format. Use format like "A1:B2"');
  }
  
  const start = coordinateToIndices(parts[0]);
  const end = coordinateToIndices(parts[1]);
  
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col
  };
}

// Format range coordinates to string
function formatRange(startRow, startCol, endRow, endCol) {
  const startCoord = indicesToCoordinate(startRow, startCol);
  const endCoord = indicesToCoordinate(endRow, endCol);
  return `${startCoord}:${endCoord}`;
}

module.exports = {
  validateCoordinates,
  coordinateToIndices,
  indicesToCoordinate,
  parseRange,
  formatRange
};
