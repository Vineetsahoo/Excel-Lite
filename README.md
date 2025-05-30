# 📊 Excel Lite – A Web-Based Spreadsheet Application

> *"A nimble reimagining of the spreadsheet, forged in the browser, faithful to function, and light in form."*

**Excel Lite** is a modern, responsive, and offline-capable spreadsheet web application that emulates key features of Microsoft Excel — all within your browser, with zero installation. With robust formula support, multi-sheet functionality, CSV import/export, charting, conditional formatting, and theme toggling, it serves as a powerful tool for data entry, analysis, and visualization.

---

## ✨ Features

- 🧮 **Formula Engine**: Supports Excel-style formulas like `SUM`, `AVERAGE`, `COUNT`, and more.
- 📊 **Charting**: Visualize data dynamically using built-in chart rendering.
- 🎨 **Conditional Formatting**: Automatically style cells based on their values.
- 📁 **CSV Import/Export**: Easily load and save data using standard CSV format.
- 📄 **Multiple Sheets**: Create and manage several sheets in a single workspace.
- 🌗 **Theming**: Toggle between dark and light mode for eye comfort.
- ⚙️ **Offline Capability (PWA)**: Works even without internet once loaded.
- 📱 **Responsive UI**: Optimized for desktops, tablets, and mobile devices.

---

## 🛠️ Built With

- **HTML5** – Structure of the application
- **CSS3** – Theming, responsive layout, and UI styling
- **JavaScript (ES6)** – Functional logic and user interaction
- **Service Workers + manifest.json** – Offline support via Progressive Web App architecture
- **Canvas API / Charting** – For visual data representation

---

## 🚀 Getting Started

### Prerequisites

- A modern web browser (Chrome, Firefox, Edge, Safari)
- No additional dependencies — it's a pure frontend application

### Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/Vineetsahoo/Excel-Lite.git
   ```

2. **Navigate into the Project Directory**
   ```bash
   cd Excel-Lite/spreadsheet-app
   ```

3. **Launch the App**
   - Simply open `index.html` in your browser.
   
   Or, to enable full PWA functionality, you may serve it via a local server (recommended):
   ```bash
   npx serve .
   ```

---

## 📂 Project Structure

```
Excel-Lite/
├── README.md
└── spreadsheet-app/
    ├── index.html                  # Main HTML file
    ├── manifest.json               # PWA manifest for install/offline
    ├── service-worker.js           # Enables offline use
    ├── css/
    │   └── styles.css              # Main styles and UI components
    ├── data/
    │   └── sample.csv              # Sample data for import testing
    ├── js/
    │   ├── app.js                  # App initialization logic
    │   ├── charts.js               # Charting engine
    │   ├── conditional-formatting.js  # Conditional formatting logic
    │   ├── formulas.js             # Core formula evaluation engine
    │   ├── import-export.js        # Import/export functionality
    │   ├── spreadsheet.js          # Spreadsheet logic and cell/grid handling
    │   ├── themes.js               # Dark/light theme switching
    │   └── utils.js                # Helper and utility functions
    └── icons/
        └── icon-72x72.png          # PWA icons (likely more sizes)
```

---

## 📌 Usage

1. Launch the app in your browser (`index.html`).

2. Click on cells to enter data.

3. Use the toolbar or shortcuts to:
   - Apply formulas (`=SUM(A1:A5)`)
   - Change themes
   - Apply formatting
   - Import/export CSV files
   - Create charts
   - Highlight cells conditionally

4. Work across multiple sheets using the tab interface.

5. Enjoy offline use via PWA support (you can even install it on mobile/desktop!).

---

## 💡 Key Functionality

### Formula Support
- Basic arithmetic: `+`, `-`, `*`, `/`
- Common functions: `SUM()`, `AVERAGE()`, `COUNT()`, `MAX()`, `MIN()`
- Cell references: `A1`, `B2:B10`

### Chart Types
- Bar charts
- Line charts
- Pie charts
- Dynamic data visualization

### Import/Export
- CSV file import
- CSV file export
- Data preservation across sessions

---

## 🤝 Contributing

Contributions are welcome!

1. Fork the repository
2. Create a new branch (`git checkout -b feature/your-feature`)
3. Commit your changes (`git commit -m 'Add some feature'`)
4. Push to the branch (`git push origin feature/your-feature`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the MIT License.

---

## 🙏 Acknowledgments

- Inspired by the elegance and utility of Microsoft Excel.
- Thanks to the open-source community and PWA pioneers for their tools and guidance.

---

## 📞 Contact

For any questions or suggestions, feel free to reach out:

- GitHub: [@Vineetsahoo](https://github.com/Vineetsahoo)
- Project Link: [Excel-Lite](https://github.com/Vineetsahoo/Excel-Lite)

---

## 🔧 Technical Notes

### PWA Features
- Installable on desktop and mobile devices
- Offline functionality via service workers
- App-like experience with manifest.json

### Browser Compatibility
- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

---

*Made with 💻 and ☕ for spreadsheet enthusiasts who value simplicity and power.*
