/**
 * Excel Lite Spreadsheet
 * Core spreadsheet functionality
 */
class Spreadsheet {
    constructor(options = {}) {
        this.sheets = [];
        this.currentSheetIndex = 0;
        this.selectedCell = { row: 0, col: 0 };
        this.undoStack = [];
        this.redoStack = [];
        this.clipboard = null;
        this.isFormulaMode = false;
        this.formulaEngine = null;
        this.conditionalFormatManager = null;
        this.chartManager = null;
        
        // Handle options
        this.rows = options.rows || 20;
        this.cols = options.cols || 15;
        this.container = options.container || document.getElementById('spreadsheet-container');
        
        this.init();
    }

    init() {
        this.addSheet();
        this.render();
        this.setupEventListeners();
    }

    addSheet() {
        const newSheet = {
            name: `Sheet ${this.sheets.length + 1}`,
            data: this.createEmptyData(),
            styles: {},  // For cell styling (font, color, etc.)
            columnWidths: {},  // For custom column widths
            rowHeights: {},    // For custom row heights
            mergedCells: [],   // For tracking merged cells
            id: Date.now()     // Unique identifier for the sheet
        };
        this.sheets.push(newSheet);
        this.currentSheetIndex = this.sheets.length - 1;
        this.updateStatusBar();
    }

    createEmptyData(rows = 20, cols = 15) {
        const data = [];
        for (let i = 0; i < rows; i++) {
            const row = Array(cols).fill('');
            data.push(row);
        }
        return data;
    }

    getCurrentSheet() {
        return this.sheets[this.currentSheetIndex];
    }

    /**
     * Set the value of a cell
     * @param {number} row - Row index
     * @param {number} col - Column index
     * @param {*} value - Cell value
     */
    setCellValue(row, col, value) {
        const sheet = this.getCurrentSheet();
        
        // Store the value in the data model
        sheet.data[row][col] = value;
        
        // If it's a formula, evaluate it
        if (typeof value === 'string' && value.startsWith('=')) {
            // Initialize display data if not exists
            if (!sheet.displayData) {
                sheet.displayData = Array(sheet.data.length).fill().map(() => Array(sheet.data[0].length).fill(null));
            }
            
            // Evaluate the formula
            const result = this.evaluateFormula(value, row, col);
            
            // Store the result in display data
            sheet.displayData[row][col] = result;
            
            // Update cells that depend on this cell
            if (this.formulaEngine) {
                const cellsToUpdate = this.formulaEngine.getCellsToRecalculate(row, col);
                
                cellsToUpdate.forEach(cellKey => {
                    const [depRow, depCol] = cellKey.split(',').map(Number);
                    const depValue = sheet.data[depRow][depCol];
                    
                    if (typeof depValue === 'string' && depValue.startsWith('=')) {
                        const depResult = this.evaluateFormula(depValue, depRow, depCol);
                        sheet.displayData[depRow][depCol] = depResult;
                        this.updateCellDisplay(depRow, depCol);
                    }
                });
            }
        } else if (sheet.displayData && sheet.displayData[row] && sheet.displayData[row][col] !== undefined) {
            // If not a formula but display data exists, remove it
            sheet.displayData[row][col] = null;
        }
        
        // Update the cell display
        this.updateCellDisplay(row, col);
        
        // Apply conditional formatting
        if (this.conditionalFormatManager) {
            this.conditionalFormatManager.applyRules();
        }
        
        // Update the formula bar
        this.updateFormulaBar();
        
        // Save changes
        this.saveSpreadsheetData();
    }

    calculateFormulas() {
        const sheet = this.getCurrentSheet();
        
        // Create a display data structure that will contain evaluated values
        const displayData = JSON.parse(JSON.stringify(sheet.data));
        
        // Process each cell to evaluate formulas
        for (let i = 0; i < sheet.data.length; i++) {
            for (let j = 0; j < sheet.data[i].length; j++) {
                const cellValue = sheet.data[i][j];
                if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
                    try {
                        const result = this.evaluateFormula(cellValue, i, j);
                        if (result !== null) {
                            displayData[i][j] = result;
                        }
                    } catch (e) {
                        displayData[i][j] = '#ERROR';
                        console.error(`Error evaluating formula at [${i},${j}]:`, e);
                    }
                }
            }
        }
        
        // Store the display data for rendering
        sheet.displayData = displayData;
    }

    render() {
        this.renderHeaders();
        this.renderSheetData();
        this.renderSheetTabs();
        this.calculateFormulas();
    }

    renderHeaders() {
        const sheet = this.getCurrentSheet();
        const headerRow = document.getElementById('header-row');
        headerRow.innerHTML = '<th class="corner-header"></th>';
        
        // Add column headers (A, B, C, etc.)
        for (let i = 0; i < sheet.data[0].length; i++) {
            const colHeader = document.createElement('th');
            colHeader.textContent = this.getColumnLabel(i);
            colHeader.className = 'column-header';
            
            // Apply custom width if defined
            if (sheet.columnWidths[i]) {
                colHeader.style.width = sheet.columnWidths[i] + 'px';
            }
            
            // Add modern header styling with hover effect
            colHeader.innerHTML = `
                <div class="header-content">
                    <span>${this.getColumnLabel(i)}</span>
                    <div class="header-actions">
                        <button class="header-sort-btn" title="Sort" data-col="${i}">
                            <i class="fas fa-sort"></i>
                        </button>
                        <button class="header-filter-btn" title="Filter" data-col="${i}">
                            <i class="fas fa-filter"></i>
                        </button>
                    </div>
                </div>
            `;
            
            // Add resize handle with improved styling
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'resize-handle';
            resizeHandle.innerHTML = '<div class="resize-handle-line"></div>';
            resizeHandle.addEventListener('mousedown', this.handleColumnResize.bind(this, i));
            colHeader.appendChild(resizeHandle);
            
            headerRow.appendChild(colHeader);
        }
    }

    renderSheetData() {
        const sheet = this.getCurrentSheet();
        const tbody = document.getElementById('data-rows');
        tbody.innerHTML = '';
        
        sheet.data.forEach((rowData, rowIndex) => {
            const tr = document.createElement('tr');
            
            // Add row header (1, 2, 3, etc.)
            const rowHeader = document.createElement('th');
            rowHeader.className = 'row-header';
            rowHeader.textContent = rowIndex + 1;
            
            // Apply custom height if defined
            if (sheet.rowHeights[rowIndex]) {
                tr.style.height = sheet.rowHeights[rowIndex] + 'px';
            }
            
            tr.appendChild(rowHeader);
            
            // Add cells
            rowData.forEach((cellData, colIndex) => {
                const td = document.createElement('td');
                td.dataset.row = rowIndex;
                td.dataset.col = colIndex;
                
                // Check if this cell is part of a merged cell
                const mergedCell = this.findMergedCell(rowIndex, colIndex);
                if (mergedCell) {
                    if (mergedCell.row === rowIndex && mergedCell.col === colIndex) {
                        // This is the top-left cell of the merged region
                        td.rowSpan = mergedCell.rowSpan;
                        td.colSpan = mergedCell.colSpan;
                        this.renderCellContent(td, rowIndex, colIndex, sheet);
                    } else {
                        // This is a merged cell that should be hidden
                        return;
                    }
                } else {
                    this.renderCellContent(td, rowIndex, colIndex, sheet);
                }
                
                // Apply styling if defined
                if (sheet.styles[`${rowIndex},${colIndex}`]) {
                    Object.assign(td.style, sheet.styles[`${rowIndex},${colIndex}`]);
                }
                
                tr.appendChild(td);
            });
            
            tbody.appendChild(tr);
        });

        // Select active cell
        this.selectCell(this.selectedCell.row, this.selectedCell.col);
    }

    renderCellContent(td, rowIndex, colIndex, sheet) {
        const rawValue = sheet.data[rowIndex][colIndex];
        const displayValue = sheet.displayData ? sheet.displayData[rowIndex][colIndex] : rawValue;
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = displayValue;
        
        // Add events
        input.addEventListener('focus', () => {
            this.selectCell(rowIndex, colIndex);
            if (typeof rawValue === 'string' && rawValue.startsWith('=')) {
                input.value = rawValue; // Show formula when focused
            }
        });
        
        input.addEventListener('blur', () => {
            if (input.value !== rawValue) {
                this.setCellValue(rowIndex, colIndex, input.value);
            } else if (typeof rawValue === 'string' && rawValue.startsWith('=')) {
                input.value = displayValue; // Show calculated value when blurred
            }
        });
        
        input.addEventListener('keydown', (e) => {
            this.handleCellKeydown(e, rowIndex, colIndex);
        });
        
        td.appendChild(input);
    }

    renderSheetTabs() {
        const tabsContainer = document.getElementById('sheet-tabs');
        tabsContainer.innerHTML = '';
        
        this.sheets.forEach((sheet, index) => {
            const tab = document.createElement('div');
            tab.className = 'sheet-tab';
            if (index === this.currentSheetIndex) {
                tab.classList.add('active');
            }
            
            tab.innerHTML = `
                <span>${sheet.name}</span>
                <i class="fas fa-times sheet-close"></i>
            `;
            
            tab.querySelector('span').addEventListener('click', () => {
                this.switchToSheet(index);
            });
            
            tab.querySelector('span').addEventListener('dblclick', () => {
                this.renameSheet(index);
            });
            
            tab.querySelector('.sheet-close').addEventListener('click', (e) => {
                e.stopPropagation();
                this.deleteSheet(index);
            });
            
            tabsContainer.appendChild(tab);
        });
        
        const addButton = document.createElement('button');
        addButton.id = 'add-sheet';
        addButton.className = 'button-icon';
        addButton.innerHTML = '<i class="fas fa-plus"></i>';
        addButton.addEventListener('click', () => {
            this.addSheet();
            this.render();
        });
        
        tabsContainer.appendChild(addButton);
    }

    switchToSheet(index) {
        if (index >= 0 && index < this.sheets.length) {
            this.currentSheetIndex = index;
            this.render();
            this.updateStatusBar();
        }
    }

    renameSheet(index) {
        const newName = prompt('Enter new sheet name:', this.sheets[index].name);
        if (newName && newName.trim() !== '') {
            this.sheets[index].name = newName.trim();
            this.renderSheetTabs();
        }
    }

    deleteSheet(index) {
        if (this.sheets.length <= 1) {
            alert('Cannot delete the only sheet.');
            return;
        }
        
        if (confirm(`Are you sure you want to delete "${this.sheets[index].name}"?`)) {
            this.sheets.splice(index, 1);
            if (this.currentSheetIndex >= index && this.currentSheetIndex > 0) {
                this.currentSheetIndex--;
            }
            this.render();
            this.updateStatusBar();
        }
    }

    selectCell(row, col) {
        const sheet = this.getCurrentSheet();
        
        // Validate row and col are within bounds
        if (row < 0 || row >= sheet.data.length || col < 0 || col >= sheet.data[0].length) {
            return;
        }
        
        // Remove previous selection
        const prevSelected = document.querySelector('.selected');
        if (prevSelected) {
            prevSelected.classList.remove('selected');
        }
        
        // Add new selection
        const cell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
        if (cell) {
            cell.classList.add('selected');
            
            // Focus the input
            const input = cell.querySelector('input');
            if (input) {
                input.focus();
            }
        }
        
        this.selectedCell = { row, col };
        this.updateFormulaBar(row, col);
        this.updateActiveCellDisplay(row, col);
    }

    updateFormulaBar(row, col) {
        const sheet = this.getCurrentSheet();
        const formulaInput = document.getElementById('formula-input');
        const cellValue = sheet.data[row][col];
        
        formulaInput.value = cellValue || '';
    }

    updateActiveCellDisplay(row, col) {
        const activeCellDisplay = document.getElementById('active-cell');
        activeCellDisplay.textContent = `${this.getColumnLabel(col)}${row + 1}`;
    }

    updateStatusBar() {
        const sheet = this.getCurrentSheet();
        const statusInfo = document.getElementById('sheet-info');
        statusInfo.textContent = `${sheet.name} | ${sheet.data.length} rows × ${sheet.data[0].length} columns`;
    }

    /**
     * Set the formula engine
     * @param {FormulaEngine} engine - Formula engine instance
     */
    setFormulaEngine(engine) {
        this.formulaEngine = engine;
    }

    /**
     * Set the conditional format manager
     * @param {ConditionalFormatManager} manager - Conditional format manager instance
     */
    setConditionalFormatManager(manager) {
        this.conditionalFormatManager = manager;
    }

    /**
     * Set the chart manager
     * @param {ChartManager} manager - Chart manager instance
     */
    setChartManager(manager) {
        this.chartManager = manager;
    }

    /**
     * Evaluate a formula
     * @param {string} formula - Formula to evaluate
     * @param {number} row - Row index
     * @param {number} col - Column index
     * @returns {*} - Evaluated result
     */
    evaluateFormula(formula, row, col) {
        if (!this.formulaEngine) {
            console.error('Formula engine not initialized');
            return '#ERROR!';
        }
        
        return this.formulaEngine.evaluate(formula, row, col);
    }

    /**
     * Recalculate all formulas in the current sheet
     */
    recalculateFormulas() {
        const sheet = this.getCurrentSheet();
        
        // Initialize display data if not exists
        if (!sheet.displayData) {
            sheet.displayData = Array(sheet.data.length).fill().map(() => Array(sheet.data[0].length).fill(null));
        }
        
        // Find all cells with formulas
        for (let row = 0; row < sheet.data.length; row++) {
            for (let col = 0; col < sheet.data[0].length; col++) {
                const value = sheet.data[row][col];
                
                if (typeof value === 'string' && value.startsWith('=')) {
                    // Evaluate formula
                    const result = this.evaluateFormula(value, row, col);
                    
                    // Store result in display data
                    sheet.displayData[row][col] = result;
                    
                    // Update cell in UI
                    this.updateCellDisplay(row, col);
                }
            }
        }
        
        // Apply conditional formatting
        if (this.conditionalFormatManager) {
            this.conditionalFormatManager.applyRules();
        }
    }

    showStatusMessage(message, type = 'info') {
        const statusEl = document.getElementById('status-message');
        if (statusEl) {
            statusEl.textContent = message;
            statusEl.className = type;
            
            // Clear message after 3 seconds
            setTimeout(() => {
                statusEl.textContent = '';
                statusEl.className = '';
            }, 3000);
        }
    }

    handleCellKeydown(e, row, col) {
        const sheet = this.getCurrentSheet();
        
        // Handle navigation with arrow keys
        if (e.key === 'ArrowUp' && row > 0) {
            e.preventDefault();
            this.selectCell(row - 1, col);
        } else if (e.key === 'ArrowDown' && row < sheet.data.length - 1) {
            e.preventDefault();
            this.selectCell(row + 1, col);
        } else if (e.key === 'ArrowLeft' && col > 0) {
            if (e.target.selectionStart === 0) {
                e.preventDefault();
                this.selectCell(row, col - 1);
            }
        } else if (e.key === 'ArrowRight' && col < sheet.data[0].length - 1) {
            if (e.target.selectionStart === e.target.value.length) {
                e.preventDefault();
                this.selectCell(row, col + 1);
            }
        } else if (e.key === 'Tab') {
            e.preventDefault();
            if (e.shiftKey && col > 0) {
                this.selectCell(row, col - 1);
            } else if (!e.shiftKey && col < sheet.data[0].length - 1) {
                this.selectCell(row, col + 1);
            }
        } else if (e.key === 'Enter') {
            e.preventDefault();
            if (e.shiftKey && row > 0) {
                this.selectCell(row - 1, col);
            } else if (!e.shiftKey && row < sheet.data.length - 1) {
                this.selectCell(row + 1, col);
            }
        }
    }

    handleColumnResize(colIndex, e) {
        e.preventDefault();
        e.stopPropagation();
        
        const sheet = this.getCurrentSheet();
        const startX = e.clientX;
        const currentWidth = sheet.columnWidths[colIndex] || 100;
        
        const onMouseMove = (moveEvent) => {
            const deltaX = moveEvent.clientX - startX;
            const newWidth = Math.max(50, currentWidth + deltaX);
            sheet.columnWidths[colIndex] = newWidth;
            this.renderHeaders();
        };
        
        const onMouseUp = () => {
            document.removeEventListener('mousemove', onMouseMove);
            document.removeEventListener('mouseup', onMouseUp);
        };
        
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
    }

    findMergedCell(row, col) {
        const sheet = this.getCurrentSheet();
        return sheet.mergedCells.find(cell => 
            row >= cell.row && row < cell.row + cell.rowSpan &&
            col >= cell.col && col < cell.col + cell.colSpan
        );
    }

    mergeCells(startRow, startCol, endRow, endCol) {
        if (startRow > endRow) [startRow, endRow] = [endRow, startRow];
        if (startCol > endCol) [startCol, endCol] = [endCol, startCol];
        
        const sheet = this.getCurrentSheet();
        
        // Save state for undo
        this.saveStateForUndo();
        
        // Add merged cell information
        sheet.mergedCells.push({
            row: startRow,
            col: startCol,
            rowSpan: endRow - startRow + 1,
            colSpan: endCol - startCol + 1
        });
        
        // Copy the content from the top-left cell to all cells in the merged region
        const content = sheet.data[startRow][startCol];
        for (let i = startRow; i <= endRow; i++) {
            for (let j = startCol; j <= endCol; j++) {
                if (i !== startRow || j !== startCol) {
                    sheet.data[i][j] = '';
                }
            }
        }
        
        this.render();
    }

    unmergeCells(row, col) {
        const sheet = this.getCurrentSheet();
        const mergedCellIndex = sheet.mergedCells.findIndex(cell => 
            row >= cell.row && row < cell.row + cell.rowSpan &&
            col >= cell.col && col < cell.col + cell.colSpan
        );
        
        if (mergedCellIndex !== -1) {
            // Save state for undo
            this.saveStateForUndo();
            
            // Remove the merged cell
            sheet.mergedCells.splice(mergedCellIndex, 1);
            this.render();
        }
    }

    addRow(index) {
        const sheet = this.getCurrentSheet();
        
        // Save state for undo
        this.saveStateForUndo();
        
        // Create new empty row
        const newRow = Array(sheet.data[0].length).fill('');
        
        // Insert at specified index or append
        if (index !== undefined) {
            sheet.data.splice(index, 0, newRow);
        } else {
            sheet.data.push(newRow);
        }
        
        // Update merged cells
        sheet.mergedCells.forEach(cell => {
            if (index !== undefined && cell.row >= index) {
                cell.row++;
            }
        });
        
        this.render();
        this.updateStatusBar();
        this.showStatusMessage('Row added');
    }

    addColumn(index) {
        const sheet = this.getCurrentSheet();
        
        // Save state for undo
        this.saveStateForUndo();
        
        // Insert new column
        sheet.data.forEach(row => {
            if (index !== undefined) {
                row.splice(index, 0, '');
            } else {
                row.push('');
            }
        });
        
        // Update merged cells
        sheet.mergedCells.forEach(cell => {
            if (index !== undefined && cell.col >= index) {
                cell.col++;
            }
        });
        
        this.render();
        this.updateStatusBar();
        this.showStatusMessage('Column added');
    }

    deleteRow(index) {
        const sheet = this.getCurrentSheet();
        
        if (sheet.data.length <= 1) {
            this.showStatusMessage('Cannot delete the only row', 'error');
            return;
        }
        
        // Save state for undo
        this.saveStateForUndo();
        
        // Delete row
        sheet.data.splice(index, 1);
        
        // Update or remove affected merged cells
        for (let i = sheet.mergedCells.length - 1; i >= 0; i--) {
            const cell = sheet.mergedCells[i];
            if (cell.row <= index && index < cell.row + cell.rowSpan) {
                // This merged cell includes the deleted row
                if (cell.rowSpan > 1) {
                    cell.rowSpan--;
                } else {
                    // Remove the merged cell if it only had one row
                    sheet.mergedCells.splice(i, 1);
                }
            } else if (cell.row > index) {
                // This merged cell is below the deleted row
                cell.row--;
            }
        }
        
        // Adjust selected cell if needed
        if (this.selectedCell.row >= sheet.data.length) {
            this.selectedCell.row = sheet.data.length - 1;
        }
        
        this.render();
        this.updateStatusBar();
        this.showStatusMessage('Row deleted');
    }

    deleteColumn(index) {
        const sheet = this.getCurrentSheet();
        
        if (sheet.data[0].length <= 1) {
            this.showStatusMessage('Cannot delete the only column', 'error');
            return;
        }
        
        // Save state for undo
        this.saveStateForUndo();
        
        // Delete column
        sheet.data.forEach(row => {
            row.splice(index, 1);
        });
        
        // Update or remove affected merged cells
        for (let i = sheet.mergedCells.length - 1; i >= 0; i--) {
            const cell = sheet.mergedCells[i];
            if (cell.col <= index && index < cell.col + cell.colSpan) {
                // This merged cell includes the deleted column
                if (cell.colSpan > 1) {
                    cell.colSpan--;
                } else {
                    // Remove the merged cell if it only had one column
                    sheet.mergedCells.splice(i, 1);
                }
            } else if (cell.col > index) {
                // This merged cell is to the right of the deleted column
                cell.col--;
            }
        }
        
        // Adjust selected cell if needed
        if (this.selectedCell.col >= sheet.data[0].length) {
            this.selectedCell.col = sheet.data[0].length - 1;
        }
        
        this.render();
        this.updateStatusBar();
        this.showStatusMessage('Column deleted');
    }

    saveStateForUndo() {
        // Save current state to undo stack
        const currentState = JSON.stringify(this.sheets);
        this.undoStack.push(currentState);
        
        // Clear redo stack when a new action is performed
        this.redoStack = [];
        
        // Limit undo stack size
        if (this.undoStack.length > 20) {
            this.undoStack.shift();
        }
    }

    undo() {
        if (this.undoStack.length === 0) {
            this.showStatusMessage('Nothing to undo', 'info');
            return;
        }
        
        // Save current state to redo stack
        const currentState = JSON.stringify(this.sheets);
        this.redoStack.push(currentState);
        
        // Restore previous state
        const previousState = this.undoStack.pop();
        this.sheets = JSON.parse(previousState);
        
        this.render();
        this.updateStatusBar();
        this.showStatusMessage('Undo successful');
    }

    redo() {
        if (this.redoStack.length === 0) {
            this.showStatusMessage('Nothing to redo', 'info');
            return;
        }
        
        // Save current state to undo stack
        const currentState = JSON.stringify(this.sheets);
        this.undoStack.push(currentState);
        
        // Restore next state
        const nextState = this.redoStack.pop();
        this.sheets = JSON.parse(nextState);
        
        this.render();
        this.updateStatusBar();
        this.showStatusMessage('Redo successful');
    }

    copy() {
        const sheet = this.getCurrentSheet();
        const { row, col } = this.selectedCell;
        
        this.clipboard = {
            type: 'cell',
            content: sheet.data[row][col],
            style: sheet.styles[`${row},${col}`] || {}
        };
        
        this.showStatusMessage('Copied cell');
    }

    cut() {
        this.copy();
        
        // Save state for undo
        this.saveStateForUndo();
        
        const sheet = this.getCurrentSheet();
        const { row, col } = this.selectedCell;
        
        // Clear the cell
        sheet.data[row][col] = '';
        
        this.render();
        this.showStatusMessage('Cut cell');
    }

    paste() {
        if (!this.clipboard) {
            this.showStatusMessage('Nothing to paste', 'info');
            return;
        }
        
        // Save state for undo
        this.saveStateForUndo();
        
        const sheet = this.getCurrentSheet();
        const { row, col } = this.selectedCell;
        
        if (this.clipboard.type === 'cell') {
            // Paste content
            sheet.data[row][col] = this.clipboard.content;
            
            // Paste style if any
            if (Object.keys(this.clipboard.style).length > 0) {
                sheet.styles[`${row},${col}`] = { ...this.clipboard.style };
            }
        }
        
        this.render();
        this.showStatusMessage('Pasted cell');
    }

    importCSV(csvContent) {
        try {
            // Save state for undo
            this.saveStateForUndo();
            
            // Parse CSV
            const rows = csvContent.split('\n');
            const data = rows.map(row => row.split(',').map(cell => cell.trim()));
            
            // Create new sheet with the imported data
            const newSheet = {
                name: `Imported ${this.sheets.length + 1}`,
                data: data,
                styles: {},
                columnWidths: {},
                rowHeights: {},
                mergedCells: [],
                id: Date.now()
            };
            
            this.sheets.push(newSheet);
            this.currentSheetIndex = this.sheets.length - 1;
            
            this.render();
            this.updateStatusBar();
            this.showStatusMessage('CSV imported successfully');
        } catch (error) {
            console.error('Error importing CSV:', error);
            this.showStatusMessage('Error importing CSV', 'error');
        }
    }

    exportCSV() {
        const sheet = this.getCurrentSheet();
        
        // Convert data to CSV format
        const csvContent = sheet.data.map(row => row.join(',')).join('\n');
        
        // Create a blob and download
        const blob = new Blob([csvContent], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = `${sheet.name.replace(/\s+/g, '_')}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        this.showStatusMessage('CSV exported successfully');
    }

    setupEventListeners() {
        // Formula bar input handling
        const formulaInput = document.getElementById('formula-input');
        formulaInput.addEventListener('input', () => {
            const { row, col } = this.selectedCell;
            const sheet = this.getCurrentSheet();
            sheet.data[row][col] = formulaInput.value;
        });
        
        formulaInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const { row, col } = this.selectedCell;
                this.setCellValue(row, col, formulaInput.value);
                this.selectCell(row + 1, col); // Move to next row
            }
        });
        
        // Button event listeners
        document.getElementById('add-row').addEventListener('click', () => {
            this.addRow(this.selectedCell.row + 1);
        });
        
        document.getElementById('add-column').addEventListener('click', () => {
            this.addColumn(this.selectedCell.col + 1);
        });
        
        document.getElementById('delete-row').addEventListener('click', () => {
            this.deleteRow(this.selectedCell.row);
        });
        
        document.getElementById('delete-column').addEventListener('click', () => {
            this.deleteColumn(this.selectedCell.col);
        });
        
        document.getElementById('undo').addEventListener('click', () => {
            this.undo();
        });
        
        document.getElementById('redo').addEventListener('click', () => {
            this.redo();
        });
        
        // Import/Export
        document.getElementById('import-csv').addEventListener('click', () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.csv';
            input.onchange = (e) => {
                const file = e.target.files[0];
                if (file) {
                    const reader = new FileReader();
                    reader.onload = (event) => {
                        this.importCSV(event.target.result);
                    };
                    reader.readAsText(file);
                }
            };
            input.click();
        });
        
        document.getElementById('export-csv').addEventListener('click', () => {
            this.exportCSV();
        });
        
        // Dark mode toggle
        document.getElementById('dark-mode-toggle').addEventListener('click', () => {
            document.body.classList.toggle('dark-mode');
            document.getElementById('dark-mode-toggle').classList.toggle('active');
        });
        
        // Context menu
        document.addEventListener('contextmenu', (e) => {
            const cell = e.target.closest('td');
            if (cell) {
                e.preventDefault();
                const row = parseInt(cell.dataset.row);
                const col = parseInt(cell.dataset.col);
                this.selectCell(row, col);
                
                const contextMenu = document.getElementById('context-menu');
                contextMenu.style.left = `${e.pageX}px`;
                contextMenu.style.top = `${e.pageY}px`;
                contextMenu.classList.add('show');
            }
        });
        
        document.addEventListener('click', () => {
            document.getElementById('context-menu').classList.remove('show');
        });
        
        // Handle context menu actions
        document.getElementById('context-menu').addEventListener('click', (e) => {
            const action = e.target.closest('.context-menu-item').dataset.action;
            switch (action) {
                case 'cut':
                    this.cut();
                    break;
                case 'copy':
                    this.copy();
                    break;
                case 'paste':
                    this.paste();
                    break;
                case 'insert-row':
                    this.addRow(this.selectedCell.row);
                    break;
                case 'insert-column':
                    this.addColumn(this.selectedCell.col);
                    break;
                case 'delete-row':
                    this.deleteRow(this.selectedCell.row);
                    break;
                case 'delete-column':
                    this.deleteColumn(this.selectedCell.col);
                    break;
            }
        });
        
        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            // Ctrl+Z: Undo
            if (e.ctrlKey && e.key === 'z') {
                e.preventDefault();
                this.undo();
            }
            // Ctrl+Y: Redo
            else if (e.ctrlKey && e.key === 'y') {
                e.preventDefault();
                this.redo();
            }
            // Ctrl+C: Copy
            else if (e.ctrlKey && e.key === 'c') {
                if (!e.target.matches('input')) {
                    e.preventDefault();
                    this.copy();
                }
            }
            // Ctrl+X: Cut
            else if (e.ctrlKey && e.key === 'x') {
                if (!e.target.matches('input')) {
                    e.preventDefault();
                    this.cut();
                }
            }
            // Ctrl+V: Paste
            else if (e.ctrlKey && e.key === 'v') {
                if (!e.target.matches('input')) {
                    e.preventDefault();
                    this.paste();
                }
            }
        });
    }

    getColumnLabel(index) {
        let columnLabel = '';
        let i = index;
        
        while (i >= 0) {
            columnLabel = String.fromCharCode(65 + (i % 26)) + columnLabel;
            i = Math.floor(i / 26) - 1;
        }
        
        return columnLabel;
    }
    
    saveSpreadsheetData() {
        // Implement data saving logic here (e.g., to server or local storage)
        console.log('Spreadsheet data saved');
    }
    
    updateCellDisplay(row, col) {
        const sheet = this.getCurrentSheet();
        const cell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
        if (cell) {
            const value = sheet.displayData ? sheet.displayData[row][col] : null;
            const input = cell.querySelector('input');
            if (input) {
                input.value = value !== null && value !== undefined ? value : sheet.data[row][col];
            }
        }
    }
}

// Initialize the spreadsheet
const spreadsheet = new Spreadsheet();