/**
 * Excel Lite Conditional Format Manager
 * Handles conditional formatting for spreadsheet cells
 */
class ConditionalFormatManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.rules = [];
  }
  
  /**
   * Add a formatting rule
   * @param {Object} rule - The formatting rule to add
   */
  addRule(rule) {
    this.rules.push(rule);
    this.applyRules();
  }
  
  /**
   * Remove a formatting rule
   * @param {number} index - Index of the rule to remove
   */
  removeRule(index) {
    if (index >= 0 && index < this.rules.length) {
      this.rules.splice(index, 1);
      this.applyRules();
    }
  }
  
  /**
   * Apply all formatting rules to the spreadsheet
   */
  applyRules() {
    if (!this.spreadsheet) {
      console.error('Spreadsheet not initialized');
      return;
    }
    
    const sheet = this.spreadsheet.getCurrentSheet();
    
    // Clear previous conditional formatting
    this.clearFormatting();
    
    // Apply each rule
    this.rules.forEach(rule => {
      this.applySingleRule(rule, sheet);
    });
    
    // Refresh the spreadsheet display
    this.spreadsheet.render();
  }
  
  /**
   * Apply a single formatting rule
   * @param {Object} rule - The rule to apply
   * @param {Object} sheet - The current sheet
   */
  applySingleRule(rule, sheet) {
    // Implementation would depend on rule types
    console.log('Applying rule:', rule);
  }
  
  /**
   * Clear all conditional formatting
   */
  clearFormatting() {
    // Implementation would clear formatting from cells
    console.log('Clearing conditional formatting');
  }
  
  /**
   * Show the conditional formatting dialog
   */
  showFormatDialog() {
    const dialog = document.getElementById('conditional-format-dialog');
    if (dialog) {
      dialog.style.display = 'block';
      
      // Set up event listeners for the dialog
      this.setupDialogEvents();
    }
  }
  
  /**
   * Set up event listeners for the conditional formatting dialog
   */
  setupDialogEvents() {
    const dialog = document.getElementById('conditional-format-dialog');
    const ruleTypeSelect = document.getElementById('cf-rule-type');
    const applyBtn = document.getElementById('cf-apply-btn');
    const cancelBtn = document.getElementById('cf-cancel-btn');
    const closeBtn = dialog.querySelector('.dialog-close');
    
    // Show/hide appropriate inputs based on rule type
    ruleTypeSelect?.addEventListener('change', () => {
      this.updateDialogInputs(ruleTypeSelect.value);
    });
    
    // Apply button
    applyBtn?.addEventListener('click', () => {
      this.createRuleFromDialog();
      dialog.style.display = 'none';
    });
    
    // Cancel button
    cancelBtn?.addEventListener('click', () => {
      dialog.style.display = 'none';
    });
    
    // Close button
    closeBtn?.addEventListener('click', () => {
      dialog.style.display = 'none';
    });
  }
  
  /**
   * Update dialog inputs based on selected rule type
   * @param {string} ruleType - Selected rule type
   */
  updateDialogInputs(ruleType) {
    const singleValueInput = document.getElementById('cf-single-value');
    const betweenValues = document.querySelector('.cf-between-values');
    const colorScale = document.querySelector('.cf-color-scale');
    
    // Hide all inputs first
    singleValueInput.style.display = 'none';
    betweenValues.style.display = 'none';
    colorScale.style.display = 'none';
    
    // Show appropriate inputs
    switch (ruleType) {
      case 'between':
        betweenValues.style.display = 'block';
        break;
      case 'colorScale':
        colorScale.style.display = 'block';
        break;
      default:
        singleValueInput.style.display = 'block';
    }
  }
  
  /**
   * Create a rule from dialog inputs
   */
  createRuleFromDialog() {
    const ruleType = document.getElementById('cf-rule-type').value;
    
    // Get values based on rule type
    let value, minValue, maxValue;
    
    if (ruleType === 'between') {
      minValue = document.getElementById('cf-min-value').value;
      maxValue = document.getElementById('cf-max-value').value;
    } else if (ruleType === 'colorScale') {
      minValue = document.getElementById('cf-min-color').value;
      maxValue = document.getElementById('cf-max-color').value;
    } else {
      value = document.getElementById('cf-value').value;
    }
    
    // Get formatting options
    const bgColor = document.getElementById('cf-bg-color').value;
    const textColor = document.getElementById('cf-text-color').value;
    const bold = document.getElementById('cf-bold').checked;
    const italic = document.getElementById('cf-italic').checked;
    
    // Create the rule
    const rule = {
      type: ruleType,
      value,
      minValue,
      maxValue,
      format: {
        backgroundColor: bgColor,
        color: textColor,
        fontWeight: bold ? 'bold' : 'normal',
        fontStyle: italic ? 'italic' : 'normal'
      }
    };
    
    // Add the rule
    this.addRule(rule);
  }
}

// Add ConditionalFormatManager to the window object
window.ConditionalFormatManager = ConditionalFormatManager;
