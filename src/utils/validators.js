// Validate sheet name
function validateSheetName(name) {
  if (typeof name !== 'string') {
    throw new Error('Sheet name must be a string');
  }
  
  if (name.trim().length === 0) {
    throw new Error('Sheet name cannot be empty');
  }
  
  if (name.length > 31) {
    throw new Error('Sheet name cannot exceed 31 characters');
  }
  
  // Check for invalid characters
  const invalidChars = /[\\/?*[\]]/;
  if (invalidChars.test(name)) {
    throw new Error('Sheet name contains invalid characters');
  }
  
  return true;
}

// Validate cell value
function validateCellValue(value) {
  // Allow null, undefined, string, number, boolean
  const validTypes = ['string', 'number', 'boolean', 'undefined'];
  
  if (value !== null && !validTypes.includes(typeof value)) {
    throw new Error('Cell value must be string, number, boolean, null, or undefined');
  }
  
  return true;
}

// Validate formula
function validateFormula(formula) {
  if (formula !== null && typeof formula !== 'string') {
    throw new Error('Formula must be a string or null');
  }
  
  if (typeof formula === 'string' && !formula.startsWith('=')) {
    throw new Error('Formula must start with "="');
  }
  
  return true;
}

module.exports = {
  validateSheetName,
  validateCellValue,
  validateFormula
};
