/**
 * Excel Lite File Manager
 * Handles import and export functionality
 */
class FileManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  /**
   * Import CSV file
   * @param {File} file - CSV file to import
   * @returns {Promise} - Promise that resolves when import is complete
   */
  importCSV(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const csvContent = e.target.result;
          
          // Parse CSV content
          const rows = csvContent.split('\n');
          const data = rows.map(row => row.split(',').map(cell => cell.trim()));
          
          // Create new sheet with imported data
          if (this.spreadsheet) {
            this.spreadsheet.importCSV(csvContent);
            resolve();
          } else {
            reject(new Error('Spreadsheet not initialized'));
          }
        } catch (error) {
          reject(error);
        }
      };
      
      reader.onerror = (error) => {
        reject(error);
      };
      
      reader.readAsText(file);
    });
  }

  /**
   * Export as CSV
   */
  exportCSV() {
    if (!this.spreadsheet) {
      console.error('Spreadsheet not initialized');
      return;
    }

    try {
      this.spreadsheet.exportCSV();
    } catch (error) {
      console.error('Error exporting CSV:', error);
    }
  }

  /**
   * Export as Excel
   * (Basic implementation - in a real app, this would use a library like xlsx.js)
   */
  exportExcel() {
    console.log('Excel export not implemented yet');
    alert('Excel export not implemented yet. Try CSV export instead.');
  }
}

// Add FileManager to the window object
window.FileManager = FileManager;
