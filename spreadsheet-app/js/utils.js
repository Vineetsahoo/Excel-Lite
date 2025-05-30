/**
 * Excel Lite Utility Functions
 * Collection of helper functions used throughout the application
 */

/**
 * Format a number with commas as thousands separators
 * @param {number} num - Number to format
 * @returns {string} - Formatted number
 */
function formatNumber(num) {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/**
 * Delay execution for specified milliseconds
 * @param {number} ms - Milliseconds to delay
 * @returns {Promise} - Promise that resolves after delay
 */
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Generate a unique ID
 * @returns {string} - Unique ID
 */
function generateUniqueId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

/**
 * Convert column index to letter (0 = A, 1 = B, etc.)
 * @param {number} index - Column index
 * @returns {string} - Column letter
 */
function indexToColumn(index) {
  let column = '';
  let temp = index + 1;
  
  while (temp > 0) {
    let remainder = (temp - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    temp = Math.floor((temp - remainder) / 26);
  }
  
  return column;
}

/**
 * Convert column letter to index (A = 0, B = 1, etc.)
 * @param {string} column - Column letter
 * @returns {number} - Column index
 */
function columnToIndex(column) {
  let index = 0;
  
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + column.charCodeAt(i) - 64;
  }
  
  return index - 1;
}

/**
 * Parse cell reference (e.g., "A1" -> {row: 0, col: 0})
 * @param {string} cellRef - Cell reference
 * @returns {object} - Row and column indices
 */
function parseCellReference(cellRef) {
  const match = cellRef.match(/^([A-Z]+)([0-9]+)$/);
  if (!match) return null;
  
  const colStr = match[1];
  const rowNum = parseInt(match[2], 10);
  
  const colIndex = columnToIndex(colStr);
  const rowIndex = rowNum - 1;
  
  return { row: rowIndex, col: colIndex };
}

/**
 * Create cell reference from row and column indices
 * @param {number} row - Row index
 * @param {number} col - Column index
 * @returns {string} - Cell reference (e.g., "A1")
 */
function createCellReference(row, col) {
  return indexToColumn(col) + (row + 1);
}

/**
 * Debounce function to limit how often a function can be called
 * @param {Function} func - Function to debounce
 * @param {number} wait - Milliseconds to wait
 * @returns {Function} - Debounced function
 */
function debounce(func, wait) {
  let timeout;
  
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

/**
 * Save data to localStorage
 * @param {string} key - Storage key
 * @param {*} data - Data to store
 */
function saveToLocalStorage(key, data) {
  try {
    localStorage.setItem(key, JSON.stringify(data));
  } catch (error) {
    console.error('Error saving to localStorage:', error);
  }
}

/**
 * Load data from localStorage
 * @param {string} key - Storage key
 * @returns {*} - Stored data or null if not found
 */
function loadFromLocalStorage(key) {
  try {
    const data = localStorage.getItem(key);
    return data ? JSON.parse(data) : null;
  } catch (error) {
    console.error('Error loading from localStorage:', error);
    return null;
  }
}

// Add utility functions to window
window.utils = {
  formatNumber,
  delay,
  generateUniqueId,
  indexToColumn,
  columnToIndex,
  parseCellReference,
  createCellReference,
  debounce,
  saveToLocalStorage,
  loadFromLocalStorage
};
