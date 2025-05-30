# Excel Lite: Modern Spreadsheet Application

Excel Lite is a comprehensive web-based spreadsheet application that replicates the core functionality and user interface of Microsoft Excel. Built with modern web technologies, it provides a familiar environment for data manipulation, calculation, and visualization.


## 🚀 Features

- **Excel-like UI**: Ribbon interface, sheets, cell formatting, and familiar navigation
- **Formula Support**: Implements common Excel formulas (SUM, AVERAGE, COUNT, etc.)
- **Multi-sheet Management**: Create, rename, delete, and navigate between sheets
- **Cell Formatting**: Apply text formatting, colors, borders, and alignment
- **Import/Export**: Support for CSV import and export
- **Data Visualization**: Create and customize charts from your data
- **Conditional Formatting**: Highlight cells based on their values
- **Dark Mode**: Toggle between light and dark themes for comfortable viewing
- **Responsive Design**: Works on desktop and tablet devices
- **Offline Support**: Progressive Web App capabilities for offline use
- **Keyboard Shortcuts**: Familiar Excel shortcuts for improved productivity

## 🛠️ Technical Implementation

Excel Lite is built with vanilla JavaScript, HTML, and CSS, focusing on modern standards and best practices:

- **Component-based Architecture**: Modular code organization for maintainability
- **Custom Formula Engine**: JavaScript-based formula parser and evaluator
- **Cell Reference System**: A1-style cell referencing with dependency tracking
- **Theming System**: CSS variables and theme management for consistent styling
- **Service Worker**: Offline capabilities and caching
- **Local Storage**: Persistent storage for user preferences and data

## 🔧 Installation & Development

### Prerequisites

- Modern web browser (Chrome, Firefox, Edge, Safari)
- Node.js and npm (for development)

### Setup

1. Clone the repository:
   ```
   git clone https://github.com/Vineetsahoo/Excel-Lite.git
   cd excel-lite
   ```

2. Open `index.html` in your browser or use a local server:
   ```
   npx serve
   ```

### Development

The project follows a modular structure:

- `index.html` - Main application entry point
- `css/styles.css` - Styles for the application
- `js/` - JavaScript modules:
  - `app.js` - Application initialization
  - `spreadsheet.js` - Core spreadsheet functionality
  - `formulas.js` - Formula parsing and evaluation
  - `conditional-formatting.js` - Cell formatting based on values
  - `charts.js` - Data visualization capabilities
  - `import-export.js` - File operations
  - `utils.js` - Utility functions
  - `themes.js` - Theme management

## 🤝 Contributing
Contributions are welcome! Feel free to fork the repository and submit pull requests.

1. Fork the repository
2. Create your feature branch: `git checkout -b feature/amazing-feature`
3. Commit your changes: `git commit -m 'Add some amazing feature'`
4. Push to the branch: `git push origin feature/amazing-feature`
5. Open a Pull Request

## 🌟 Usage Examples

### Basic Cell Editing
Click on any cell and start typing to enter data. Press Enter to confirm or navigate to another cell.

### Using Formulas
Start with an equals sign (=) followed by your formula, e.g., `=SUM(A1:A10)` or `=AVERAGE(B1:B5)`.

Supported formulas include:
- SUM
- AVERAGE
- COUNT
- MAX
- MIN
- IF
- CONCATENATE

### Formatting Cells
Select a cell or range, then use the toolbar to apply formatting such as:
- Font style (bold, italic, underline)
- Text color and background color
- Text alignment
- Number formats

### Working with Sheets
Use the sheet tabs at the bottom to:
- Switch between sheets
- Create new sheets (+ button)
- Delete sheets (× button)
- Rename sheets (double-click tab)

## 📄 License
This project is licensed under the MIT License - see the LICENSE file for details.

## 👏 Acknowledgments
- Inspired by the functionality and design of Microsoft Excel
- Uses Font Awesome for icons
- Special thanks to all contributors who have helped improve this project