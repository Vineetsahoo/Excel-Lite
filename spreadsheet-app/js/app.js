/**
 * Excel Lite - Main Application
 * Initialize and connect all components
 */
document.addEventListener('DOMContentLoaded', () => {
    // Remove loading indicator code
    // showLoadingIndicator();
    
    // Initialize spreadsheet
    const spreadsheet = new Spreadsheet({
        rows: 100,
        cols: 26,
        container: document.getElementById('spreadsheet-container')
    });
    
    // Initialize theme manager
    const themeManager = new ThemeManager();
    window.themeManager = themeManager;
    
    // Initialize formula engine and connect to spreadsheet
    const formulaEngine = new FormulaEngine(spreadsheet);
    spreadsheet.setFormulaEngine(formulaEngine);
    
    // Initialize conditional formatting manager
    const conditionalFormatManager = new ConditionalFormatManager(spreadsheet);
    spreadsheet.setConditionalFormatManager(conditionalFormatManager);
    
    // Initialize chart manager
    const chartManager = new ChartManager(spreadsheet);
    spreadsheet.setChartManager(chartManager);
    
    // Initialize file manager
    const fileManager = new FileManager(spreadsheet);
    
    // Add to window for debugging
    window.spreadsheet = spreadsheet;
    window.fileManager = fileManager;
    
    // Set up event listeners
    setupEventListeners(spreadsheet, fileManager, themeManager);
    
    // Remove loading indicator code
    // hideLoadingIndicator();
    
    // Show welcome dialog for first-time users
    if (!localStorage.getItem('excel-lite-welcomed')) {
        showWelcomeDialog();
        localStorage.setItem('excel-lite-welcomed', 'true');
    }
    
    // Add scroll shadow effects to spreadsheet container
    addScrollShadowEffects();
});

/**
 * Add shadow effects when scrolling the spreadsheet container
 * This gives a more modern look and better visual cues
 */
function addScrollShadowEffects() {
    const spreadsheetContainer = document.getElementById('spreadsheet-container');
    
    if (!spreadsheetContainer) return;
    
    // Add shadow effects on scroll
    spreadsheetContainer.addEventListener('scroll', function() {
        // Add shadow classes based on scroll position
        const hasHorizontalScroll = this.scrollWidth > this.clientWidth;
        const hasVerticalScroll = this.scrollHeight > this.clientHeight;
        
        // Add/remove shadow classes for top
        if (this.scrollTop > 10) {
            this.classList.add('shadow-top');
        } else {
            this.classList.remove('shadow-top');
        }
        
        // Add/remove shadow classes for bottom
        if (hasVerticalScroll && (this.scrollHeight - this.scrollTop - this.clientHeight) > 10) {
            this.classList.add('shadow-bottom');
        } else {
            this.classList.remove('shadow-bottom');
        }
        
        // Add/remove shadow classes for left
        if (this.scrollLeft > 10) {
            this.classList.add('shadow-left');
        } else {
            this.classList.remove('shadow-left');
        }
        
        // Add/remove shadow classes for right
        if (hasHorizontalScroll && (this.scrollWidth - this.scrollLeft - this.clientWidth) > 10) {
            this.classList.add('shadow-right');
        } else {
            this.classList.remove('shadow-right');
        }
    });
    
    // Trigger initial scroll event to set shadows properly
    spreadsheetContainer.dispatchEvent(new Event('scroll'));
}

/**
 * Excel Lite - Main Application
 * Initialize and connect all components
 */
document.addEventListener('DOMContentLoaded', () => {
    // Add loading indicator with modern animation
    showLoadingIndicator();
    
    // Initialize spreadsheet with responsive sizing
    const spreadsheet = new Spreadsheet({
        rows: 100,
        cols: 26,
        container: document.getElementById('spreadsheet-container')
    });
    
    // Initialize theme manager with enhanced UI
    const themeManager = new ThemeManager();
    window.themeManager = themeManager;
    
    // Apply default theme with smooth transition
    document.body.classList.add('theme-transition');
    setTimeout(() => {
        document.body.classList.remove('theme-transition');
    }, 1000);
    
    // Initialize formula engine with modern syntax highlighting
    const formulaEngine = new FormulaEngine(spreadsheet);
    spreadsheet.setFormulaEngine(formulaEngine);
    
    // Initialize conditional formatting manager with enhanced visuals
    const conditionalFormatManager = new ConditionalFormatManager(spreadsheet);
    spreadsheet.setConditionalFormatManager(conditionalFormatManager);
    
    // Initialize chart manager with modern chart styles
    const chartManager = new ChartManager(spreadsheet);
    spreadsheet.setChartManager(chartManager);
    
    // Initialize file manager with progress indicators
    const fileManager = new FileManager(spreadsheet);
    
    // Add to window for debugging
    window.spreadsheet = spreadsheet;
    window.fileManager = fileManager;
    
    // Set up event listeners with improved feedback
    setupEventListeners(spreadsheet, fileManager, themeManager);
    
    // Hide loading indicator with fade-out animation
    setTimeout(() => {
        hideLoadingIndicator();
    }, 500);
    
    // Show welcome dialog for first-time users with modern design
    if (!localStorage.getItem('excel-lite-welcomed')) {
        showModernWelcomeDialog();
        localStorage.setItem('excel-lite-welcomed', 'true');
    }
    
    // Add scroll shadow effects to spreadsheet container
    addScrollShadowEffects();
});

// Helper function to show a visually appealing loading indicator
function showLoadingIndicator() {
    const loadingIndicator = document.createElement('div');
    loadingIndicator.className = 'app-loading';
    loadingIndicator.id = 'app-loading';
    loadingIndicator.innerHTML = `
        <div class="loading-content">
            <div class="spinner-container">
                <div class="spinner"></div>
            </div>
            <div class="loading-text">
                <h3>Loading Excel Lite</h3>
                <p>Preparing your spreadsheet experience...</p>
            </div>
        </div>
    `;
    document.body.appendChild(loadingIndicator);
}

// Helper function to hide the loading indicator with animation
function hideLoadingIndicator() {
    const loadingIndicator = document.getElementById('app-loading');
    if (loadingIndicator) {
        loadingIndicator.classList.add('fade-out');
        setTimeout(() => {
            loadingIndicator.remove();
        }, 500);
    }
}

/**
 * Set up application event listeners
 */
function setupEventListeners(spreadsheet, fileManager, themeManager) {
    // Import/Export buttons
    document.getElementById('import-csv')?.addEventListener('click', () => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.csv';
        input.onchange = e => {
            const file = e.target.files[0];
            if (file) {
                fileManager.importCSV(file)
                    .then(() => {
                        showStatusMessage('CSV imported successfully', 'success');
                    })
                    .catch(error => {
                        showStatusMessage('Error importing CSV: ' + error.message, 'error');
                    });
            }
        };
        input.click();
    });
    
    document.getElementById('export-csv')?.addEventListener('click', () => {
        fileManager.exportCSV();
    });
    
    document.getElementById('export-excel')?.addEventListener('click', () => {
        fileManager.exportExcel();
    });
    
    // Formula bar
    const formulaInput = document.getElementById('formula-input');
    if (formulaInput) {
        formulaInput.addEventListener('keydown', e => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const formula = formulaInput.value;
                const { row, col } = spreadsheet.selectedCell;
                spreadsheet.setCellValue(row, col, formula);
                spreadsheet.focusCell(row, col + 1);
            }
        });
    }
    
    // Cell formatting buttons
    document.getElementById('format-bold')?.addEventListener('click', () => {
        spreadsheet.toggleCellFormat('fontWeight', 'bold', 'normal');
    });
    
    document.getElementById('format-italic')?.addEventListener('click', () => {
        spreadsheet.toggleCellFormat('fontStyle', 'italic', 'normal');
    });
    
    document.getElementById('format-underline')?.addEventListener('click', () => {
        spreadsheet.toggleCellFormat('textDecoration', 'underline', 'none');
    });
    
    // Cell alignment buttons
    document.getElementById('align-left')?.addEventListener('click', () => {
        spreadsheet.setCellFormat('textAlign', 'left');
    });
    
    document.getElementById('align-center')?.addEventListener('click', () => {
        spreadsheet.setCellFormat('textAlign', 'center');
    });
    
    document.getElementById('align-right')?.addEventListener('click', () => {
        spreadsheet.setCellFormat('textAlign', 'right');
    });
    
    // Color pickers
    document.querySelectorAll('.color-option').forEach(option => {
        option.addEventListener('click', () => {
            const property = option.closest('.color-dropdown').dataset.property;
            const color = option.dataset.color || option.style.backgroundColor;
            spreadsheet.setCellFormat(property, color);
        });
    });
    
    // Conditional formatting button
    document.getElementById('conditional-format')?.addEventListener('click', () => {
        spreadsheet.conditionalFormatManager.showFormatDialog();
    });
    
    // Dark mode toggle - now handled by ThemeManager
    
    // Add sheet button
    document.getElementById('add-sheet')?.addEventListener('click', () => {
        spreadsheet.addSheet();
    });
    
    // Handle sheet tab clicks
    document.getElementById('sheet-tabs')?.addEventListener('click', e => {
        const tab = e.target.closest('.sheet-tab');
        if (tab) {
            const sheetIndex = parseInt(tab.dataset.index);
            if (!isNaN(sheetIndex)) {
                spreadsheet.switchToSheet(sheetIndex);
            }
            
            // Handle close button
            if (e.target.classList.contains('sheet-close')) {
                spreadsheet.deleteSheet(sheetIndex);
                e.stopPropagation();
            }
        }
    });
    
    // Handle double-click on sheet tab for renaming
    document.getElementById('sheet-tabs')?.addEventListener('dblclick', e => {
        const tab = e.target.closest('.sheet-tab');
        if (tab && e.target.tagName !== 'I') {
            const sheetIndex = parseInt(tab.dataset.index);
            const newName = prompt('Enter new sheet name:', spreadsheet.sheets[sheetIndex].name);
            if (newName) {
                spreadsheet.renameSheet(sheetIndex, newName);
            }
        }
    });
    
    // Initialize theme selector in the header actions
    const headerActions = document.querySelector('.header-actions');
    if (headerActions) {
        themeManager.createThemeSelector(headerActions);
    }
}

/**
 * Show status message
 * @param {string} message - Message to display
 * @param {string} type - Message type (success, error, info)
 */
function showStatusMessage(message, type = 'info') {
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

/**
 * Show welcome dialog for first-time users
 */
function showModernWelcomeDialog() {
    const dialog = document.createElement('div');
    dialog.className = 'welcome-dialog';
    
    dialog.innerHTML = `
        <div class="welcome-content">
            <h2>Welcome to Excel Lite!</h2>
            <p>Excel Lite is a powerful web-based spreadsheet application with Excel-like features. Here are some quick tips to get started:</p>
            <ul>
                <li>Click on cells to edit their content</li>
                <li>Use formulas starting with = (like =SUM(A1:A5))</li>
                <li>Use the toolbar for formatting options</li>
                <li>Create charts and use conditional formatting</li>
                <li>Import and export data as needed</li>
                <li>Switch between dark and light modes using the toggle</li>
            </ul>
            <button id="welcome-close">Get Started</button>
        </div>
    `;
    
    document.body.appendChild(dialog);
    
    document.getElementById('welcome-close').addEventListener('click', () => {
        dialog.style.opacity = '0';
        setTimeout(() => dialog.remove(), 300);
    });
}

// This file is the entry point for the JavaScript functionality.
// It initializes the spreadsheet and handles global application events.

// Remove the duplicate initialization - comment out or delete this section
/*
document.addEventListener('DOMContentLoaded', () => {
    // Initialize the application
    initApp();
    
    // Register service worker for offline capability
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('./service-worker.js')
            .then(reg => console.log('Service Worker registered'))
            .catch(err => console.log('Service Worker registration failed:', err));
    }
});
*/

// Only keep one initialization function
function initApp() {
    // The main spreadsheet instance is created in spreadsheet.js
    
    // Add global keyboard shortcuts
    setupGlobalShortcuts();
    
    // Setup theme toggle
    setupThemeToggle();
    
    // Setup formula suggestions
    setupFormulaSuggestions();
    
    // Setup ribbon tabs
    setupRibbonTabs();
    
    // Setup toolbar functionality
    setupToolbarInteractions();
    
    // Setup tooltips
    setupTooltips();
    
    // Initialize charts and conditional formatting
    initializeAdvancedFeatures();
    
    // Welcome message
    showWelcomeMessage();
    
    // Setup find/replace panel
    setupFindReplacePanel();
}

function setupGlobalShortcuts() {
    document.addEventListener('keydown', (e) => {
        // Ctrl+S: Save (export)
        if (e.ctrlKey && e.key === 's') {
            e.preventDefault();
            document.getElementById('export-csv').click();
        }
        
        // Ctrl+O: Open (import)
        if (e.ctrlKey && e.key === 'o') {
            e.preventDefault();
            document.getElementById('import-csv').click();
        }
        
        // Ctrl+F: Find
        if (e.ctrlKey && e.key === 'f') {
            e.preventDefault();
            showFindDialog();
        }
        
        // Ctrl+H: Replace
        if (e.ctrlKey && e.key === 'h') {
            e.preventDefault();
            showReplaceDialog();
        }
        
        // Esc: Close any open dialogs
        if (e.key === 'Escape') {
            closeAllDialogs();
        }
    });
}

function setupThemeToggle() {
    const themeToggle = document.getElementById('dark-mode-toggle');
    
    // Check for saved theme preference
    const savedTheme = localStorage.getItem('spreadsheet-theme');
    if (savedTheme === 'dark') {
        document.body.classList.add('dark-mode');
        themeToggle.classList.add('active');
    }
    
    // Toggle theme on click
    themeToggle.addEventListener('click', () => {
        document.body.classList.toggle('dark-mode');
        themeToggle.classList.toggle('active');
        
        // Save preference
        const isDark = document.body.classList.contains('dark-mode');
        localStorage.setItem('spreadsheet-theme', isDark ? 'dark' : 'light');
    });
}

function setupFormulaSuggestions() {
    const formulaInput = document.getElementById('formula-input');
    const suggestionsList = document.createElement('div');
    suggestionsList.className = 'formula-suggestions';
    document.querySelector('.formula-bar').appendChild(suggestionsList);
    
    const formulaFunctions = [
        { name: 'SUM', description: 'Adds values', syntax: 'SUM(number1, [number2], ...)' },
        { name: 'AVERAGE', description: 'Calculates average', syntax: 'AVERAGE(number1, [number2], ...)' },
        { name: 'COUNT', description: 'Counts non-empty cells', syntax: 'COUNT(value1, [value2], ...)' },
        { name: 'MAX', description: 'Finds maximum value', syntax: 'MAX(number1, [number2], ...)' },
        { name: 'MIN', description: 'Finds minimum value', syntax: 'MIN(number1, [number2], ...)' },
        { name: 'IF', description: 'Conditional logic', syntax: 'IF(condition, value_if_true, value_if_false)' },
        { name: 'CONCATENATE', description: 'Joins text values', syntax: 'CONCATENATE(text1, [text2], ...)' }
    ];
    
    formulaInput.addEventListener('input', () => {
        const value = formulaInput.value;
        
        // Show suggestions when input starts with =
        if (value.startsWith('=')) {
            const searchTerm = value.slice(1).toUpperCase();
            const matches = formulaFunctions.filter(fn => 
                fn.name.startsWith(searchTerm) && searchTerm.length > 0
            );
            
            if (matches.length > 0 && searchTerm.length > 0) {
                suggestionsList.innerHTML = '';
                matches.forEach(match => {
                    const item = document.createElement('div');
                    item.className = 'suggestion-item';
                    item.innerHTML = `
                        <strong>${match.name}</strong>
                        <span>${match.description}</span>
                        <small>${match.syntax}</small>
                    `;
                    item.addEventListener('click', () => {
                        formulaInput.value = `=${match.name}()`;
                        formulaInput.focus();
                        // Position cursor inside parentheses
                        const cursorPos = formulaInput.value.length - 1;
                        formulaInput.setSelectionRange(cursorPos, cursorPos);
                        suggestionsList.innerHTML = '';
                    });
                    suggestionsList.appendChild(item);
                });
                suggestionsList.style.display = 'block';
            } else {
                suggestionsList.style.display = 'none';
            }
        } else {
            suggestionsList.style.display = 'none';
        }
    });
    
    // Hide suggestions when clicking outside
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.formula-bar')) {
            suggestionsList.style.display = 'none';
        }
    });
}

function showWelcomeMessage() {
    // Check if this is the first visit
    if (!localStorage.getItem('spreadsheet-visited')) {
        const welcomeDialog = document.createElement('div');
        welcomeDialog.className = 'welcome-dialog';
        welcomeDialog.innerHTML = `
            <div class="welcome-content">
                <h2>Welcome to Excel Lite!</h2>
                <p>This modern spreadsheet application allows you to:</p>
                <ul>
                    <li>Create and manage multiple sheets</li>
                    <li>Use formulas like SUM, AVERAGE, COUNT</li>
                    <li>Import and export CSV files</li>
                    <li>Customize your experience with dark mode</li>
                </ul>
                <p>Get started by clicking on any cell and entering data.</p>
                <button id="welcome-close" class="button">Get Started</button>
            </div>
        `;
        document.body.appendChild(welcomeDialog);
        
        document.getElementById('welcome-close').addEventListener('click', () => {
            welcomeDialog.remove();
            localStorage.setItem('spreadsheet-visited', 'true');
        });
    }
}

function showFindDialog() {
    closeAllDialogs();
    
    const findReplacePanel = document.getElementById('find-replace-panel');
    findReplacePanel.style.display = 'block';
    
    // Switch to Find tab
    const findTab = findReplacePanel.querySelector('[data-tab="find"]');
    findTab.click();
    
    // Focus on input
    document.getElementById('find-input').focus();
}

function showReplaceDialog() {
    closeAllDialogs();
    
    const findReplacePanel = document.getElementById('find-replace-panel');
    findReplacePanel.style.display = 'block';
    
    // Switch to Replace tab
    const replaceTab = findReplacePanel.querySelector('[data-tab="replace"]');
    replaceTab.click();
    
    // Focus on input
    document.getElementById('replace-find-input').focus();
}

function setupFindReplacePanel() {
    const panel = document.getElementById('find-replace-panel');
    const tabs = panel.querySelectorAll('.find-replace-tab');
    const tabContents = panel.querySelectorAll('.tab-content');
    const closeButton = panel.querySelector('.close-button');
    
    // Tab switching
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            // Remove active class from all tabs
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            
            // Show corresponding content
            const tabId = tab.dataset.tab;
            tabContents.forEach(content => {
                content.style.display = content.dataset.content === tabId ? 'block' : 'none';
            });
        });
    });
    
    // Close button
    closeButton.addEventListener('click', () => {
        panel.style.display = 'none';
    });
    
    // Find panel buttons
    document.getElementById('find-next-btn').addEventListener('click', () => {
        const text = document.getElementById('find-input').value;
        const matchCase = document.getElementById('find-match-case').checked;
        const wholeCell = document.getElementById('find-whole-cell').checked;
        
        findInSpreadsheet(text, matchCase, wholeCell);
    });
    
    document.getElementById('find-all-btn').addEventListener('click', () => {
        const text = document.getElementById('find-input').value;
        const matchCase = document.getElementById('find-match-case').checked;
        const wholeCell = document.getElementById('find-whole-cell').checked;
        
        findAllInSpreadsheet(text, matchCase, wholeCell);
    });
    
    // Replace panel buttons
    document.getElementById('replace-btn').addEventListener('click', () => {
        const findText = document.getElementById('replace-find-input').value;
        const replaceText = document.getElementById('replace-with-input').value;
        const matchCase = document.getElementById('replace-match-case').checked;
        const wholeCell = document.getElementById('replace-whole-cell').checked;
        
        replaceInSpreadsheet(findText, replaceText, matchCase, wholeCell);
    });
    
    document.getElementById('replace-all-btn').addEventListener('click', () => {
        const findText = document.getElementById('replace-find-input').value;
        const replaceText = document.getElementById('replace-with-input').value;
        const matchCase = document.getElementById('replace-match-case').checked;
        const wholeCell = document.getElementById('replace-whole-cell').checked;
        
        replaceAllInSpreadsheet(findText, replaceText, matchCase, wholeCell);
    });
}

function closeAllDialogs() {
    const dialogs = document.querySelectorAll('.dialog');
    dialogs.forEach(dialog => dialog.remove());
}

// These functions will be implemented to interact with the spreadsheet instance
function findInSpreadsheet(searchText, matchCase, wholeCell) {
    // This would call methods on the spreadsheet instance
    if (window.spreadsheet) {
        // Implement search functionality
        console.log(`Finding: ${searchText}, Match case: ${matchCase}, Whole cell: ${wholeCell}`);
        // Example implementation would be:
        // spreadsheet.findNext(searchText, matchCase, wholeCell);
    }
}

function findAllInSpreadsheet(searchText, matchCase, wholeCell) {
    if (window.spreadsheet) {
        console.log(`Finding all: ${searchText}, Match case: ${matchCase}, Whole cell: ${wholeCell}`);
        // Example implementation would be:
        // spreadsheet.findAll(searchText, matchCase, wholeCell);
    }
}

function replaceInSpreadsheet(findText, replaceText, matchCase, wholeCell) {
    if (window.spreadsheet) {
        console.log(`Replacing: ${findText} with ${replaceText}`);
        // Example implementation would be:
        // spreadsheet.replaceNext(findText, replaceText, matchCase, wholeCell);
    }
}

function replaceAllInSpreadsheet(findText, replaceText, matchCase, wholeCell) {
    if (window.spreadsheet) {
        console.log(`Replacing all: ${findText} with ${replaceText}`);
        // Example implementation would be:
        // spreadsheet.replaceAll(findText, replaceText, matchCase, wholeCell);
    }
}

function initializeAdvancedFeatures() {
    // Wait for spreadsheet to be initialized
    setTimeout(() => {
        if (window.spreadsheet) {
            // Initialize Chart Manager
            window.chartManager = new ChartManager(window.spreadsheet);
            
            // Initialize Conditional Format Manager
            window.conditionalFormatManager = new ConditionalFormatManager(window.spreadsheet);
            
            // Initialize File Manager
            window.fileManager = new FileManager(window.spreadsheet);
            
            console.log('Advanced features initialized');
        }
    }, 500);
}

// Adding missing functions
function setupRibbonTabs() {
    const ribbonTabs = document.querySelectorAll('.ribbon-tab');
    
    ribbonTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            // Remove active class from all tabs
            ribbonTabs.forEach(t => t.classList.remove('active'));
            
            // Add active class to clicked tab
            tab.classList.add('active');
            
            // Here you would typically show the corresponding ribbon panel
            // For now, we'll just show a message
            console.log(`Switched to ${tab.textContent} tab`);
        });
    });
}

function setupToolbarInteractions() {
    // Text formatting buttons
    document.getElementById('bold-button')?.addEventListener('click', () => {
        applyFormatting('fontWeight', 'bold');
    });
    
    document.getElementById('italic-button')?.addEventListener('click', () => {
        applyFormatting('fontStyle', 'italic');
    });
    
    document.getElementById('underline-button')?.addEventListener('click', () => {
        applyFormatting('textDecoration', 'underline');
    });
    
    // Alignment buttons
    document.getElementById('align-left')?.addEventListener('click', () => {
        applyFormatting('textAlign', 'left');
    });
    
    document.getElementById('align-center')?.addEventListener('click', () => {
        applyFormatting('textAlign', 'center');
    });
    
    document.getElementById('align-right')?.addEventListener('click', () => {
        applyFormatting('textAlign', 'right');
    });
    
    // Merge cells button
    document.getElementById('merge-cells')?.addEventListener('click', () => {
        if (window.spreadsheet) {
            // Implement merge functionality
            const selection = getSelectedRange();
            if (selection && selection.startRow !== selection.endRow || selection.startCol !== selection.endCol) {
                window.spreadsheet.mergeCells(selection.startRow, selection.startCol, selection.endRow, selection.endCol);
            } else {
                showTooltip('Select multiple cells to merge', 'warning');
            }
        }
    });
    
    // Setup font dropdown
    const fontOptions = document.querySelectorAll('.font-option');
    fontOptions.forEach(option => {
        option.addEventListener('click', () => {
            const fontFamily = option.textContent;
            applyFormatting('fontFamily', fontFamily);
            // Update the dropdown display
            const fontDropdown = document.querySelector('.toolbar-dropdown span');
            if (fontDropdown) {
                fontDropdown.textContent = fontFamily;
            }
        });
    });
    
    // Setup font size dropdown
    const fontSizeOptions = document.querySelectorAll('.font-size-option');
    fontSizeOptions.forEach(option => {
        option.addEventListener('click', () => {
            const fontSize = option.textContent + 'px';
            applyFormatting('fontSize', fontSize);
            // Update the dropdown display
            const fontSizeDropdown = document.querySelector('.toolbar-dropdown:nth-child(2) span');
            if (fontSizeDropdown) {
                fontSizeDropdown.textContent = option.textContent;
            }
        });
    });
    
    // Setup color pickers
    const textColorOptions = document.querySelectorAll('.text-color-dropdown .color-option');
    textColorOptions.forEach(option => {
        option.addEventListener('click', () => {
            const color = option.style.backgroundColor;
            applyFormatting('color', color);
        });
    });
    
    const fillColorOptions = document.querySelectorAll('.fill-color-dropdown .color-option');
    fillColorOptions.forEach(option => {
        option.addEventListener('click', () => {
            const color = option.style.backgroundColor;
            applyFormatting('backgroundColor', color);
        });
    });
}

function setupTooltips() {
    const tooltipElements = document.querySelectorAll('[title]');
    const tooltip = document.getElementById('tooltip');
    
    tooltipElements.forEach(element => {
        element.addEventListener('mouseenter', (e) => {
            const title = e.target.getAttribute('title');
            e.target.removeAttribute('title'); // Remove to prevent default tooltip
            
            tooltip.textContent = title;
            tooltip.classList.add('show');
            
            // Position the tooltip
            const rect = e.target.getBoundingClientRect();
            tooltip.style.left = rect.left + (rect.width / 2) - (tooltip.offsetWidth / 2) + 'px';
            tooltip.style.top = rect.bottom + 5 + 'px';
        });
        
        element.addEventListener('mouseleave', (e) => {
            e.target.setAttribute('title', tooltip.textContent); // Restore title
            tooltip.classList.remove('show');
        });
    });
}

function applyFormatting(property, value) {
    if (window.spreadsheet) {
        const { row, col } = window.spreadsheet.selectedCell;
        const sheet = window.spreadsheet.getCurrentSheet();
        
        // Save state for undo
        window.spreadsheet.saveStateForUndo();
        
        // Initialize styles object if it doesn't exist
        if (!sheet.styles[`${row},${col}`]) {
            sheet.styles[`${row},${col}`] = {};
        }
        
        // Toggle property if it's already set
        if (sheet.styles[`${row},${col}`][property] === value) {
            delete sheet.styles[`${row},${col}`][property];
        } else {
            sheet.styles[`${row},${col}`][property] = value;
        }
        
        // Cleanup empty style objects
        if (Object.keys(sheet.styles[`${row},${col}`]).length === 0) {
            delete sheet.styles[`${row},${col}`];
        }
        
        // Render the changes
        window.spreadsheet.render();
    }
}

function getSelectedRange() {
    // For demonstration, we'll just return a simple range
    // In a real implementation, this would get the actual selected range from the spreadsheet
    if (window.spreadsheet) {
        const { row, col } = window.spreadsheet.selectedCell;
        return {
            startRow: row,
            startCol: col,
            endRow: row + 1,
            endCol: col + 1
        };
    }
    return null;
}

function showTooltip(message, type = 'info') {
    const tooltip = document.getElementById('tooltip');
    tooltip.textContent = message;
    tooltip.className = 'tooltip show ' + type;
    
    // Position the tooltip in the center of the viewport
    tooltip.style.left = '50%';
    tooltip.style.top = '10%';
    tooltip.style.transform = 'translateX(-50%)';
    
    // Hide after 3 seconds
    setTimeout(() => {
        tooltip.classList.remove('show');
    }, 3000);
}

// Add additional styling for dialogs
const dialogStyles = document.createElement('style');
dialogStyles.textContent = `
    .dialog {
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: white;
        border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
        z-index: 1000;
        min-width: 400px;
        overflow: hidden;
    }
    
    .dark-mode .dialog {
        background: #2d2d2d;
        color: #e0e0e0;
    }
    
    .dialog-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 12px 16px;
        background: var(--primary-color);
        color: white;
    }
    
    .dialog-header h3 {
        margin: 0;
        font-weight: 500;
    }
    
    .dialog-close {
        background: transparent;
        border: none;
        color: white;
        cursor: pointer;
        font-size: 16px;
    }
    
    .dialog-content {
        padding: 16px;
    }
    
    .form-group {
        margin-bottom: 16px;
    }
    
    .form-group label {
        display: block;
        margin-bottom: 6px;
    }
    
    .form-group input {
        width: 100%;
        padding: 8px 10px;
        border: 1px solid #ddd;
        border-radius: 4px;
        background-color: white;
        color: var(--dark-color);
    }
    
    .dark-mode .form-group input {
        background-color: #1e1e1e;
        border-color: #444;
        color: #e0e0e0;
    }
    
    .form-options {
        display: flex;
        gap: 16px;
        margin-bottom: 16px;
    }
    
    .dialog-buttons {
        display: flex;
        justify-content: flex-end;
        gap: 8px;
    }
    
    .welcome-dialog {
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.5);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 2000;
    }
    
    .welcome-content {
        background: white;
        padding: 24px;
        border-radius: 8px;
        max-width: 500px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.2);
    }
    
    .dark-mode .welcome-content {
        background: #2d2d2d;
        color: #e0e0e0;
    }
    
    .welcome-content h2 {
        color: var(--primary-color);
        margin-top: 0;
    }
    
    .welcome-content ul {
        margin-bottom: 20px;
    }
    
    .formula-suggestions {
        position: absolute;
        top: 100%;
        left: 0;
        width: 100%;
        background: white;
        border: 1px solid var(--border-color);
        border-top: none;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        z-index: 100;
        max-height: 200px;
        overflow-y: auto;
        display: none;
    }
    
    .dark-mode .formula-suggestions {
        background: #2d2d2d;
        border-color: #444;
    }
    
    .suggestion-item {
        padding: 8px 12px;
        cursor: pointer;
        border-bottom: 1px solid var(--border-color);
        display: flex;
        flex-direction: column;
    }
    
    .suggestion-item:hover {
        background-color: rgba(0,0,0,0.05);
    }
    
    .dark-mode .suggestion-item:hover {
        background-color: rgba(255,255,255,0.1);
    }
    
    .suggestion-item strong {
        color: var(--primary-color);
    }
    
    .suggestion-item small {
        color: #888;
        font-size: 0.8em;
    }
    
    .dark-mode .suggestion-item small {
        color: #aaa;
    }
`;

document.head.appendChild(dialogStyles);

// Add additional modern styles
const modernStyles = document.createElement('style');
modernStyles.textContent = `
    /* Animation for clicking effects */
    .button-click-effect {
        position: relative;
        overflow: hidden;
    }
    
    .button-click-effect::after {
        content: '';
        position: absolute;
        top: 50%;
        left: 50%;
        width: 5px;
        height: 5px;
        background: rgba(255, 255, 255, 0.5);
        opacity: 0;
        border-radius: 100%;
        transform: scale(1, 1) translate(-50%);
        transform-origin: 50% 50%;
    }
    
    .button-click-effect:active::after {
        animation: ripple 0.6s ease-out;
    }
    
    @keyframes ripple {
        0% {
            transform: scale(0, 0);
            opacity: 0.5;
        }
        20% {
            transform: scale(25, 25);
            opacity: 0.3;
        }
        100% {
            opacity: 0;
            transform: scale(40, 40);
        }
    }
    
    /* Excel-like custom scrollbar for formula suggestions */
    .formula-suggestions::-webkit-scrollbar {
        width: 8px;
    }
    
    .formula-suggestions::-webkit-scrollbar-track {
        background: #f1f1f1;
        border-radius: 4px;
    }
    
    .formula-suggestions::-webkit-scrollbar-thumb {
        background: #c1c1c1;
        border-radius: 4px;
    }
    
    .formula-suggestions::-webkit-scrollbar-thumb:hover {
        background: #a8a8a8;
    }
    
    /* Modern loading indicator */
    .app-loading {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(255, 255, 255, 0.8);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 9999;
    }
    
    .spinner {
        width: 40px;
        height: 40px;
        border: 4px solid rgba(33, 115, 70, 0.1);
        border-radius: 50%;
        border-top: 4px solid var(--primary-color);
        animation: spin 1s linear infinite;
    }
    
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
`;

document.head.appendChild(modernStyles);