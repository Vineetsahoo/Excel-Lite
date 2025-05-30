/**
 * Excel Lite Formula Engine
 * Provides Excel-like formula calculation capabilities
 */
class FormulaEngine {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.errorValue = '#ERROR!';
    this.refErrorValue = '#REF!';
    this.divZeroErrorValue = '#DIV/0!';
    this.nameErrorValue = '#NAME?';
    this.valueErrorValue = '#VALUE!';
    this.formulaCache = new Map();
    this.cellDependencies = new Map();
  }

  /**
   * Evaluate a formula in the context of the spreadsheet
   * @param {string} formula - The formula to evaluate
   * @param {number} row - Current row index
   * @param {number} col - Current column index
   * @returns {*} - The result of the formula evaluation
   */
  evaluate(formula, row, col) {
    // Remove the equals sign
    formula = formula.substring(1).trim();
    
    try {
      // Track cell dependencies for recalculation
      this.trackDependencies(formula, row, col);
      
      // Parse and evaluate the formula
      return this.parseFormula(formula, row, col);
    } catch (error) {
      console.error('Formula evaluation error:', error);
      return this.errorValue;
    }
  }

  /**
   * Parse and evaluate a formula
   * @param {string} formula - Formula without the equals sign
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {*} - Evaluated result
   */
  parseFormula(formula, currentRow, currentCol) {
    // Check for built-in functions
    const upperFormula = formula.toUpperCase();
    
    // Handle SUM function
    if (upperFormula.startsWith('SUM(')) {
      return this.handleSumFunction(formula, currentRow, currentCol);
    }
    
    // Handle AVERAGE function
    if (upperFormula.startsWith('AVERAGE(')) {
      return this.handleAverageFunction(formula, currentRow, currentCol);
    }
    
    // Handle COUNT function
    if (upperFormula.startsWith('COUNT(')) {
      return this.handleCountFunction(formula, currentRow, currentCol);
    }
    
    // Handle MAX function
    if (upperFormula.startsWith('MAX(')) {
      return this.handleMaxFunction(formula, currentRow, currentCol);
    }
    
    // Handle MIN function
    if (upperFormula.startsWith('MIN(')) {
      return this.handleMinFunction(formula, currentRow, currentCol);
    }
    
    // Handle IF function
    if (upperFormula.startsWith('IF(')) {
      return this.handleIfFunction(formula, currentRow, currentCol);
    }
    
    // Handle CONCATENATE function
    if (upperFormula.startsWith('CONCATENATE(')) {
      return this.handleConcatenateFunction(formula, currentRow, currentCol);
    }
    
    // Handle cell references and basic expressions
    return this.evaluateExpression(formula, currentRow, currentCol);
  }

  /**
   * Evaluate a basic expression with cell references
   * @param {string} expression - The expression to evaluate
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {*} - Evaluated result
   */
  evaluateExpression(expression, currentRow, currentCol) {
    // Replace cell references with their values
    const cellRefRegex = /([A-Z]+[0-9]+)(?![A-Za-z0-9])/g;
    
    expression = expression.replace(cellRefRegex, (match) => {
      const cellValue = this.getCellValue(match, currentRow, currentCol);
      
      // Handle error values
      if (typeof cellValue === 'string' && cellValue.startsWith('#')) {
        throw new Error(cellValue);
      }
      
      // If it's a number, return it directly
      if (typeof cellValue === 'number') {
        return cellValue.toString();
      }
      
      // If it's a string that's not a number, wrap in quotes
      if (isNaN(cellValue)) {
        return `"${cellValue}"`;
      }
      
      return cellValue;
    });
    
    // Safely evaluate the expression
    try {
      // Use Function instead of eval for better security
      const result = new Function(`return ${expression}`)();
      
      // Format the result
      if (typeof result === 'number') {
        // Round to 10 decimal places to avoid floating point issues
        return Math.round(result * 1e10) / 1e10;
      }
      
      return result;
    } catch (error) {
      console.error('Expression evaluation error:', error);
      return this.errorValue;
    }
  }

  /**
   * Get the value of a cell by its reference
   * @param {string} cellRef - Cell reference (e.g., A1)
   * @param {number} currentRow - Current row for relative references
   * @param {number} currentCol - Current column for relative references
   * @returns {*} - Cell value
   */
  getCellValue(cellRef, currentRow, currentCol) {
    const coords = this.parseCellReference(cellRef);
    
    if (!coords) {
      return this.refErrorValue;
    }
    
    const { row, col } = coords;
    
    // Check if cell is in bounds
    const sheet = this.spreadsheet.getCurrentSheet();
    if (row < 0 || row >= sheet.data.length || col < 0 || col >= sheet.data[0].length) {
      return this.refErrorValue;
    }
    
    // Get cell value
    const cellValue = sheet.data[row][col];
    
    // If the cell contains a formula, we need to avoid circular references
    if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
      // Check for circular reference
      const key = `${row},${col}`;
      const currentKey = `${currentRow},${currentCol}`;
      
      if (this.cellDependencies.has(key) && this.cellDependencies.get(key).includes(currentKey)) {
        return this.errorValue; // Circular reference
      }
      
      // Use display data if available (already calculated)
      if (sheet.displayData && sheet.displayData[row] && sheet.displayData[row][col] !== undefined) {
        return sheet.displayData[row][col];
      }
      
      // Otherwise, need to calculate it
      return this.errorValue; // Can't calculate now, would need a separate pass
    }
    
    return cellValue;
  }

  /**
   * Handle SUM function
   * @param {string} formula - The SUM formula
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {number} - Sum of values
   */
  handleSumFunction(formula, currentRow, currentCol) {
    const args = this.extractFunctionArgs(formula, 'SUM');
    let sum = 0;
    
    for (const arg of args) {
      // Check if it's a cell range
      if (arg.includes(':')) {
        const rangeValues = this.getCellRangeValues(arg, currentRow, currentCol);
        
        for (const value of rangeValues) {
          if (typeof value === 'number') {
            sum += value;
          } else if (typeof value === 'string' && !isNaN(value)) {
            sum += parseFloat(value);
          }
        }
      } else {
        // Single cell or value
        let value;
        
        if (/^[A-Z]+[0-9]+$/.test(arg)) {
          // It's a cell reference
          value = this.getCellValue(arg, currentRow, currentCol);
        } else {
          // It's a direct value
          value = isNaN(arg) ? 0 : parseFloat(arg);
        }
        
        if (typeof value === 'number') {
          sum += value;
        } else if (typeof value === 'string' && !isNaN(value)) {
          sum += parseFloat(value);
        }
      }
    }
    
    return sum;
  }

  /**
   * Handle AVERAGE function
   * @param {string} formula - The AVERAGE formula
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {number} - Average of values
   */
  handleAverageFunction(formula, currentRow, currentCol) {
    const args = this.extractFunctionArgs(formula, 'AVERAGE');
    let sum = 0;
    let count = 0;
    
    for (const arg of args) {
      // Check if it's a cell range
      if (arg.includes(':')) {
        const rangeValues = this.getCellRangeValues(arg, currentRow, currentCol);
        
        for (const value of rangeValues) {
          if (typeof value === 'number') {
            sum += value;
            count++;
          } else if (typeof value === 'string' && !isNaN(value)) {
            sum += parseFloat(value);
            count++;
          }
        }
      } else {
        // Single cell or value
        let value;
        
        if (/^[A-Z]+[0-9]+$/.test(arg)) {
          // It's a cell reference
          value = this.getCellValue(arg, currentRow, currentCol);
        } else {
          // It's a direct value
          value = isNaN(arg) ? null : parseFloat(arg);
        }
        
        if (typeof value === 'number') {
          sum += value;
          count++;
        } else if (typeof value === 'string' && !isNaN(value)) {
          sum += parseFloat(value);
          count++;
        }
      }
    }
    
    return count === 0 ? 0 : sum / count;
  }

  /**
   * Handle COUNT function
   * @param {string} formula - The COUNT formula
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {number} - Count of numeric values
   */
  handleCountFunction(formula, currentRow, currentCol) {
    const args = this.extractFunctionArgs(formula, 'COUNT');
    let count = 0;
    
    for (const arg of args) {
      // Check if it's a cell range
      if (arg.includes(':')) {
        const rangeValues = this.getCellRangeValues(arg, currentRow, currentCol);
        
        for (const value of rangeValues) {
          if (typeof value === 'number' || (typeof value === 'string' && !isNaN(value))) {
            count++;
          }
        }
      } else {
        // Single cell or value
        let value;
        
        if (/^[A-Z]+[0-9]+$/.test(arg)) {
          // It's a cell reference
          value = this.getCellValue(arg, currentRow, currentCol);
        } else {
          // It's a direct value
          value = arg;
        }
        
        if (typeof value === 'number' || (typeof value === 'string' && !isNaN(value))) {
          count++;
        }
      }
    }
    
    return count;
  }

  /**
   * Handle MAX function
   * @param {string} formula - The MAX formula
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {number} - Maximum value
   */
  handleMaxFunction(formula, currentRow, currentCol) {
    const args = this.extractFunctionArgs(formula, 'MAX');
    let max = -Infinity;
    let foundValue = false;
    
    for (const arg of args) {
      // Check if it's a cell range
      if (arg.includes(':')) {
        const rangeValues = this.getCellRangeValues(arg, currentRow, currentCol);
        
        for (const value of rangeValues) {
          if (typeof value === 'number') {
            max = Math.max(max, value);
            foundValue = true;
          } else if (typeof value === 'string' && !isNaN(value)) {
            max = Math.max(max, parseFloat(value));
            foundValue = true;
          }
        }
      } else {
        // Single cell or value
        let value;
        
        if (/^[A-Z]+[0-9]+$/.test(arg)) {
          // It's a cell reference
          value = this.getCellValue(arg, currentRow, currentCol);
        } else {
          // It's a direct value
          value = isNaN(arg) ? null : parseFloat(arg);
        }
        
        if (typeof value === 'number') {
          max = Math.max(max, value);
          foundValue = true;
        } else if (typeof value === 'string' && !isNaN(value)) {
          max = Math.max(max, parseFloat(value));
          foundValue = true;
        }
      }
    }
    
    return foundValue ? max : 0;
  }

  /**
   * Handle MIN function
   * @param {string} formula - The MIN formula
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {number} - Minimum value
   */
  handleMinFunction(formula, currentRow, currentCol) {
    const args = this.extractFunctionArgs(formula, 'MIN');
    let min = Infinity;
    let foundValue = false;
    
    for (const arg of args) {
      // Check if it's a cell range
      if (arg.includes(':')) {
        const rangeValues = this.getCellRangeValues(arg, currentRow, currentCol);
        
        for (const value of rangeValues) {
          if (typeof value === 'number') {
            min = Math.min(min, value);
            foundValue = true;
          } else if (typeof value === 'string' && !isNaN(value)) {
            min = Math.min(min, parseFloat(value));
            foundValue = true;
          }
        }
      } else {
        // Single cell or value
        let value;
        
        if (/^[A-Z]+[0-9]+$/.test(arg)) {
          // It's a cell reference
          value = this.getCellValue(arg, currentRow, currentCol);
        } else {
          // It's a direct value
          value = isNaN(arg) ? null : parseFloat(arg);
        }
        
        if (typeof value === 'number') {
          min = Math.min(min, value);
          foundValue = true;
        } else if (typeof value === 'string' && !isNaN(value)) {
          min = Math.min(min, parseFloat(value));
          foundValue = true;
        }
      }
    }
    
    return foundValue ? min : 0;
  }

  /**
   * Handle IF function
   * @param {string} formula - The IF formula
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {*} - Result based on condition
   */
  handleIfFunction(formula, currentRow, currentCol) {
    // Extract arguments
    const argsString = formula.substring(3, formula.length - 1);
    const args = this.splitFunctionArgs(argsString);
    
    if (args.length !== 3) {
      return this.errorValue;
    }
    
    const [condition, trueValue, falseValue] = args;
    
    // Evaluate the condition
    let conditionResult;
    try {
      conditionResult = this.evaluateExpression(condition, currentRow, currentCol);
    } catch (error) {
      return this.errorValue;
    }
    
    // Determine which value to return
    if (conditionResult) {
      return this.evaluateArg(trueValue, currentRow, currentCol);
    } else {
      return this.evaluateArg(falseValue, currentRow, currentCol);
    }
  }

  /**
   * Handle CONCATENATE function
   * @param {string} formula - The CONCATENATE formula
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {string} - Concatenated string
   */
  handleConcatenateFunction(formula, currentRow, currentCol) {
    const args = this.extractFunctionArgs(formula, 'CONCATENATE');
    let result = '';
    
    for (const arg of args) {
      // Check if it's a cell reference
      if (/^[A-Z]+[0-9]+$/.test(arg)) {
        const value = this.getCellValue(arg, currentRow, currentCol);
        result += value !== null && value !== undefined ? value : '';
      } else if (arg.startsWith('"') && arg.endsWith('"')) {
        // It's a string literal
        result += arg.substring(1, arg.length - 1);
      } else {
        // Try to evaluate as an expression
        try {
          const value = this.evaluateExpression(arg, currentRow, currentCol);
          result += value !== null && value !== undefined ? value : '';
        } catch (error) {
          result += arg;
        }
      }
    }
    
    return result;
  }

  /**
   * Extract function arguments from a formula
   * @param {string} formula - The formula
   * @param {string} funcName - Function name
   * @returns {string[]} - Array of arguments
   */
  extractFunctionArgs(formula, funcName) {
    const argsStart = formula.indexOf('(', funcName.length) + 1;
    const argsEnd = this.findMatchingClosingBracket(formula, argsStart - 1);
    const argsString = formula.substring(argsStart, argsEnd);
    
    return this.splitFunctionArgs(argsString);
  }

  /**
   * Split function arguments, respecting nested functions and quotes
   * @param {string} argsString - Arguments string
   * @returns {string[]} - Array of arguments
   */
  splitFunctionArgs(argsString) {
    const args = [];
    let currentArg = '';
    let inQuotes = false;
    let bracketDepth = 0;
    
    for (let i = 0; i < argsString.length; i++) {
      const char = argsString[i];
      
      if (char === '"' && argsString[i - 1] !== '\\') {
        inQuotes = !inQuotes;
        currentArg += char;
      } else if (char === '(' && !inQuotes) {
        bracketDepth++;
        currentArg += char;
      } else if (char === ')' && !inQuotes) {
        bracketDepth--;
        currentArg += char;
      } else if (char === ',' && !inQuotes && bracketDepth === 0) {
        args.push(currentArg.trim());
        currentArg = '';
      } else {
        currentArg += char;
      }
    }
    
    if (currentArg) {
      args.push(currentArg.trim());
    }
    
    return args;
  }

  /**
   * Find the matching closing bracket
   * @param {string} formula - The formula
   * @param {number} openingIndex - Index of opening bracket
   * @returns {number} - Index of matching closing bracket
   */
  findMatchingClosingBracket(formula, openingIndex) {
    let depth = 1;
    for (let i = openingIndex + 1; i < formula.length; i++) {
      if (formula[i] === '(') {
        depth++;
      } else if (formula[i] === ')') {
        depth--;
        if (depth === 0) {
          return i;
        }
      }
    }
    return formula.length;
  }

  /**
   * Evaluate a single argument
   * @param {string} arg - Argument to evaluate
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {*} - Evaluated value
   */
  evaluateArg(arg, currentRow, currentCol) {
    // Check if it's a cell reference
    if (/^[A-Z]+[0-9]+$/.test(arg)) {
      return this.getCellValue(arg, currentRow, currentCol);
    }
    
    // Check if it's a string literal
    if (arg.startsWith('"') && arg.endsWith('"')) {
      return arg.substring(1, arg.length - 1);
    }
    
    // Try to evaluate as an expression
    try {
      return this.evaluateExpression(arg, currentRow, currentCol);
    } catch (error) {
      return arg;
    }
  }

  /**
   * Get all values in a cell range
   * @param {string} range - Cell range (e.g., "A1:B5")
   * @param {number} currentRow - Current row index
   * @param {number} currentCol - Current column index
   * @returns {Array} - Array of cell values
   */
  getCellRangeValues(range, currentRow, currentCol) {
    const [startRef, endRef] = range.split(':');
    const startCoords = this.parseCellReference(startRef);
    const endCoords = this.parseCellReference(endRef);
    
    if (!startCoords || !endCoords) {
      return [];
    }
    
    const values = [];
    const sheet = this.spreadsheet.getCurrentSheet();
    
    for (let row = startCoords.row; row <= endCoords.row; row++) {
      for (let col = startCoords.col; col <= endCoords.col; col++) {
        if (row >= 0 && row < sheet.data.length && col >= 0 && col < sheet.data[0].length) {
          values.push(this.getCellValue(`${this.getColumnLabel(col)}${row + 1}`, currentRow, currentCol));
        }
      }
    }
    
    return values;
  }

  /**
   * Parse cell reference to row and column indices
   * @param {string} cellRef - Cell reference (e.g., "A1")
   * @returns {object|null} - Row and column indices or null if invalid
   */
  parseCellReference(cellRef) {
    const match = cellRef.match(/^([A-Z]+)([0-9]+)$/);
    if (!match) return null;
    
    const colStr = match[1];
    const rowNum = parseInt(match[2], 10);
    
    // Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, ...)
    let colIndex = 0;
    for (let i = 0; i < colStr.length; i++) {
      colIndex = colIndex * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    colIndex--; // Adjust to 0-based index
    
    // Convert row number to index (1-based to 0-based)
    const rowIndex = rowNum - 1;
    
    return { row: rowIndex, col: colIndex };
  }

  /**
   * Get column label from index
   * @param {number} index - Column index
   * @returns {string} - Column label (e.g., "A", "B", ..., "AA", etc.)
   */
  getColumnLabel(index) {
    let columnLabel = '';
    let i = index;
    
    while (i >= 0) {
      columnLabel = String.fromCharCode(65 + (i % 26)) + columnLabel;
      i = Math.floor(i / 26) - 1;
    }
    
    return columnLabel;
  }

  /**
   * Track cell dependencies for recalculation
   * @param {string} formula - The formula
   * @param {number} row - Current row index
   * @param {number} col - Current column index
   */
  trackDependencies(formula, row, col) {
    const cellKey = `${row},${col}`;
    
    // Extract all cell references
    const cellRefs = this.extractCellReferences(formula);
    
    // Add dependencies
    cellRefs.forEach(ref => {
      const coords = this.parseCellReference(ref);
      if (coords) {
        const depKey = `${coords.row},${coords.col}`;
        
        if (!this.cellDependencies.has(depKey)) {
          this.cellDependencies.set(depKey, []);
        }
        
        const deps = this.cellDependencies.get(depKey);
        if (!deps.includes(cellKey)) {
          deps.push(cellKey);
        }
      }
    });
  }

  /**
   * Extract all cell references from a formula
   * @param {string} formula - The formula
   * @returns {string[]} - Array of cell references
   */
  extractCellReferences(formula) {
    const cellRefRegex = /([A-Z]+[0-9]+)(?![A-Za-z0-9])/g;
    const cellRefs = [];
    let match;
    
    while ((match = cellRefRegex.exec(formula)) !== null) {
      cellRefs.push(match[1]);
    }
    
    // Also handle cell ranges
    const cellRangeRegex = /([A-Z]+[0-9]+):([A-Z]+[0-9]+)/g;
    
    while ((match = cellRangeRegex.exec(formula)) !== null) {
      const [, startRef, endRef] = match;
      const startCoords = this.parseCellReference(startRef);
      const endCoords = this.parseCellReference(endRef);
      
      if (startCoords && endCoords) {
        for (let row = startCoords.row; row <= endCoords.row; row++) {
          for (let col = startCoords.col; col <= endCoords.col; col++) {
            cellRefs.push(`${this.getColumnLabel(col)}${row + 1}`);
          }
        }
      }
    }
    
    return [...new Set(cellRefs)]; // Remove duplicates
  }

  /**
   * Get cells that need to be recalculated when a cell changes
   * @param {number} row - Row of changed cell
   * @param {number} col - Column of changed cell
   * @returns {string[]} - Array of cell keys that need recalculation
   */
  getCellsToRecalculate(row, col) {
    const cellKey = `${row},${col}`;
    const result = [];
    
    if (this.cellDependencies.has(cellKey)) {
      const directDependents = this.cellDependencies.get(cellKey);
      
      // Recursive function to get all dependents
      const getAllDependents = (key) => {
        if (!result.includes(key)) {
          result.push(key);
          
          if (this.cellDependencies.has(key)) {
            this.cellDependencies.get(key).forEach(depKey => {
              getAllDependents(depKey);
            });
          }
        }
      };
      
      directDependents.forEach(depKey => {
        getAllDependents(depKey);
      });
    }
    
    return result;
  }

  /**
   * Clear the formula cache
   */
  clearCache() {
    this.formulaCache.clear();
  }
}

// Add FormulaEngine to the window object
window.FormulaEngine = FormulaEngine;