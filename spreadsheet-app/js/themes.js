/**
 * Excel Lite Theme Manager
 * Handles theme-related functionality including dark mode and UI customization
 */
class ThemeManager {
  constructor() {
    this.themes = {
      light: {
        name: 'Light',
        colors: {
          primary: '#217346', // Excel green
          secondary: '#2b579a', // Excel blue
          accent: '#ed7d31', // Excel orange
          background: '#ffffff',
          text: '#212529',
          border: '#e0e0e0', // Lighter border for modern look
          hover: '#f5f5f5',
          cellBackground: '#ffffff',
          headerBackground: '#f8f8f8', // Lighter header for modern look
          toolbarBackground: '#ffffff'
        }
      },
      dark: {
        name: 'Dark',
        colors: {
          primary: '#26865b', // Brighter green for dark mode
          secondary: '#4472c4', // Brighter blue for dark mode
          accent: '#f4b183', // Lighter orange for dark mode
          background: '#1e1e1e',
          text: '#f0f0f0',
          border: '#444444',
          hover: '#333333',
          cellBackground: '#252525',
          headerBackground: '#2a2a2a',
          toolbarBackground: '#2d2d2d'
        }
      },
      excel365: {
        name: 'Excel 365',
        colors: {
          primary: '#217346', // Excel green
          secondary: '#2b579a', // Excel blue
          accent: '#ed7d31', // Excel orange
          background: '#ffffff',
          text: '#212529',
          border: '#d4d4d4',
          hover: '#f3f3f3',
          cellBackground: '#ffffff',
          headerBackground: '#f3f2f1',
          toolbarBackground: '#ffffff'
        }
      },
      classic: {
        name: 'Classic Excel',
        colors: {
          primary: '#009933', // Classic Excel green
          secondary: '#0070c0', // Classic Excel blue
          accent: '#ff9900', // Classic Excel orange
          background: '#ffffff',
          text: '#000000',
          border: '#c0c0c0',
          hover: '#e6e6e6',
          cellBackground: '#ffffff',
          headerBackground: '#e6e6e6',
          toolbarBackground: '#f0f0f0'
        }
      },
      // Updated modern themes
      modern: {
        name: 'Modern',
        colors: {
          primary: '#00a884', // Modern teal
          secondary: '#0078d4', // Modern blue
          accent: '#ff6f61', // Coral accent
          background: '#ffffff',
          text: '#202020',
          border: '#e8e8e8', // Lighter border for modern look
          hover: '#f7f7f7',
          cellBackground: '#ffffff',
          headerBackground: '#fafafa',
          toolbarBackground: '#f8f8f8'
        }
      },
      midnight: {
        name: 'Midnight',
        colors: {
          primary: '#4cc2ff', // Bright blue
          secondary: '#9d65ff', // Purple
          accent: '#ff7eb6', // Pink
          background: '#0f172a', // Dark blue-black
          text: '#e2e8f0',
          border: '#334155',
          hover: '#1e293b',
          cellBackground: '#141e33',
          headerBackground: '#1e293b',
          toolbarBackground: '#172032'
        }
      },
      solarized: {
        name: 'Solarized',
        colors: {
          primary: '#268bd2', // Solarized blue
          secondary: '#859900', // Solarized green
          accent: '#cb4b16', // Solarized orange
          background: '#fdf6e3', // Solarized light background
          text: '#657b83', // Solarized text
          border: '#eee8d5',
          hover: '#eee8d5',
          cellBackground: '#fdf6e3',
          headerBackground: '#eee8d5',
          toolbarBackground: '#fdf6e3'
        }
      },
      highContrast: {
        name: 'High Contrast',
        colors: {
          primary: '#ffff00', // Yellow
          secondary: '#00ffff', // Cyan
          accent: '#ff00ff', // Magenta
          background: '#000000', // Black
          text: '#ffffff', // White
          border: '#ffffff',
          hover: '#333333',
          cellBackground: '#000000',
          headerBackground: '#000000',
          toolbarBackground: '#000000'
        }
      },
      // Add new contemporary themes
      fluent: {
        name: 'Fluent Design',
        colors: {
          primary: '#0078d7', // Microsoft blue
          secondary: '#107c10', // Microsoft green
          accent: '#ffb900', // Microsoft yellow
          background: '#fafafa',
          text: '#323130',
          border: '#edebe9',
          hover: '#f3f2f1',
          cellBackground: '#ffffff',
          headerBackground: '#f5f5f5',
          toolbarBackground: '#fafafa'
        }
      },
      material: {
        name: 'Material',
        colors: {
          primary: '#6200ee', // Material purple
          secondary: '#03dac6', // Material teal
          accent: '#ff4081', // Material pink
          background: '#ffffff',
          text: '#121212',
          border: '#e0e0e0',
          hover: '#f5f5f5',
          cellBackground: '#ffffff',
          headerBackground: '#f5f5f5',
          toolbarBackground: '#fafafa'
        }
      },
      nord: {
        name: 'Nord',
        colors: {
          primary: '#5e81ac', // Nord blue
          secondary: '#81a1c1', // Nord light blue
          accent: '#ebcb8b', // Nord yellow
          background: '#eceff4',
          text: '#2e3440',
          border: '#d8dee9',
          hover: '#e5e9f0',
          cellBackground: '#ffffff',
          headerBackground: '#e5e9f0',
          toolbarBackground: '#eceff4'
        }
      },
      nordDark: {
        name: 'Nord Dark',
        colors: {
          primary: '#88c0d0', // Nord cyan
          secondary: '#81a1c1', // Nord light blue
          accent: '#ebcb8b', // Nord yellow
          background: '#2e3440',
          text: '#eceff4',
          border: '#3b4252',
          hover: '#434c5e',
          cellBackground: '#3b4252',
          headerBackground: '#434c5e',
          toolbarBackground: '#3b4252'
        }
      }
    };
    
    this.currentTheme = 'light';
    this.init();
  }

  /**
   * Initialize the theme manager
   */
  init() {
    // Load custom themes first, then saved preferences
    this.loadCustomThemes();
    
    // Load saved theme preference
    this.loadSavedTheme();
    
    // Apply the current theme
    this.applyTheme(this.currentTheme);
    
    // Set up event listeners
    this.setupEventListeners();
    
    // Log initialization
    console.log('Theme Manager initialized with theme:', this.currentTheme);
  }

  /**
   * Load saved theme preference from localStorage
   */
  loadSavedTheme() {
    const savedTheme = localStorage.getItem('spreadsheet-theme');
    if (savedTheme && this.themes[savedTheme]) {
      this.currentTheme = savedTheme;
    }
  }

  /**
   * Apply a theme to the application
   * @param {string} themeName - Name of the theme to apply
   */
  applyTheme(themeName) {
    if (!this.themes[themeName]) {
      console.error(`Theme "${themeName}" not found`);
      return;
    }
    
    // Save the current theme
    this.currentTheme = themeName;
    
    // Save to localStorage
    localStorage.setItem('spreadsheet-theme', themeName);
    
    // Apply dark mode class if needed
    if (themeName === 'dark' || themeName === 'midnight' || themeName === 'highContrast') {
      document.body.classList.add('dark-mode');
      document.getElementById('dark-mode-toggle')?.classList.add('active');
    } else {
      document.body.classList.remove('dark-mode');
      document.getElementById('dark-mode-toggle')?.classList.remove('active');
    }
    
    // Apply special high-contrast class if needed
    if (themeName === 'highContrast') {
      document.body.classList.add('high-contrast');
    } else {
      document.body.classList.remove('high-contrast');
    }
    
    // Apply theme colors to CSS variables
    const theme = this.themes[themeName];
    const root = document.documentElement;
    
    // Apply all theme colors to CSS variables
    Object.entries(theme.colors).forEach(([key, value]) => {
      root.style.setProperty(`--${key}-color`, value);
    });
    
    // Set primary variables for backward compatibility
    root.style.setProperty('--primary-color', theme.colors.primary);
    root.style.setProperty('--secondary-color', theme.colors.secondary);
    root.style.setProperty('--accent-color', theme.colors.accent);
    root.style.setProperty('--light-color', theme.colors.background);
    root.style.setProperty('--dark-color', theme.colors.text);
    root.style.setProperty('--border-color', theme.colors.border);
    root.style.setProperty('--hover-color', theme.colors.hover);
    
    // Update toolbar if it exists
    this.updateToolbarThemeIndicator(themeName);
    
    // Dispatch a theme change event
    window.dispatchEvent(new CustomEvent('themechange', { 
      detail: { 
        theme: themeName,
        colors: theme.colors
      } 
    }));
    
    console.log(`Theme applied: ${themeName}`);
  }

  /**
   * Update theme indicator in toolbar if it exists
   */
  updateToolbarThemeIndicator(themeName) {
    const themeIndicator = document.querySelector('.current-theme-indicator');
    if (themeIndicator) {
      const theme = this.themes[themeName];
      themeIndicator.textContent = theme.name;
      
      const themeColorDot = themeIndicator.querySelector('.theme-color-dot');
      if (themeColorDot) {
        themeColorDot.style.backgroundColor = theme.colors.primary;
      }
    }
  }

  /**
   * Toggle between light and dark themes
   */
  toggleDarkMode() {
    const newTheme = this.currentTheme === 'dark' ? 'light' : 'dark';
    this.applyTheme(newTheme);
    
    // Add animation effect to the body
    document.body.classList.add('theme-transition');
    setTimeout(() => {
      document.body.classList.remove('theme-transition');
    }, 1000);
  }

  /**
   * Set up event listeners for theme-related actions
   */
  setupEventListeners() {
    // Set up the dark mode toggle
    const darkModeToggle = document.getElementById('dark-mode-toggle');
    if (darkModeToggle) {
      darkModeToggle.addEventListener('click', () => {
        this.toggleDarkMode();
      });
    }
    
    // Listen for theme selection events
    document.addEventListener('click', (e) => {
      const themeSelector = e.target.closest('[data-theme]');
      if (themeSelector) {
        const themeName = themeSelector.dataset.theme;
        if (this.themes[themeName]) {
          this.applyTheme(themeName);
          
          // Add ripple effect
          const ripple = document.createElement('span');
          ripple.className = 'theme-selector-ripple';
          themeSelector.appendChild(ripple);
          
          setTimeout(() => {
            ripple.remove();
          }, 600);
        }
      }
    });
    
    // Handle system preference changes
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    
    // Initial check
    if (mediaQuery.matches && !localStorage.getItem('spreadsheet-theme')) {
      this.applyTheme('dark');
    }
    
    // Listen for changes
    mediaQuery.addEventListener('change', (e) => {
      // Only apply automatically if the user hasn't set a preference
      if (!localStorage.getItem('spreadsheet-theme')) {
        this.applyTheme(e.matches ? 'dark' : 'light');
      }
    });
    
    // Listen for keyboard shortcuts
    document.addEventListener('keydown', (e) => {
      // Alt+Shift+T - Toggle theme selector
      if (e.altKey && e.shiftKey && e.key === 'T') {
        e.preventDefault();
        this.toggleThemeSelector();
      }
      
      // Alt+Shift+D - Toggle dark mode
      if (e.altKey && e.shiftKey && e.key === 'D') {
        e.preventDefault();
        this.toggleDarkMode();
      }
    });
  }

  /**
   * Toggle the theme selector dialog
   */
  toggleThemeSelector() {
    let themeDialog = document.getElementById('theme-selector-dialog');
    
    if (themeDialog) {
      // Close if already open
      themeDialog.remove();
      return;
    }
    
    // Create theme selector dialog
    themeDialog = document.createElement('div');
    themeDialog.id = 'theme-selector-dialog';
    themeDialog.className = 'theme-selector-dialog';
    
    const dialogContent = document.createElement('div');
    dialogContent.className = 'theme-selector-content';
    
    // Add header
    const header = document.createElement('div');
    header.className = 'theme-selector-header';
    header.innerHTML = `
      <h3>Select Theme</h3>
      <button class="theme-dialog-close"><i class="fas fa-times"></i></button>
    `;
    dialogContent.appendChild(header);
    
    // Add theme grid
    const themeGrid = document.createElement('div');
    themeGrid.className = 'theme-grid';
    
    // Add themes to grid
    Object.keys(this.themes).forEach(themeId => {
      const theme = this.themes[themeId];
      const themeCard = document.createElement('div');
      themeCard.className = 'theme-card';
      themeCard.dataset.theme = themeId;
      
      if (themeId === this.currentTheme) {
        themeCard.classList.add('active');
      }
      
      themeCard.innerHTML = `
        <div class="theme-preview" style="
          background-color: ${theme.colors.background};
          color: ${theme.colors.text};
          border: 1px solid ${theme.colors.border};
        ">
          <div class="preview-header" style="background-color: ${theme.colors.primary}"></div>
          <div class="preview-toolbar" style="background-color: ${theme.colors.toolbarBackground}"></div>
          <div class="preview-content">
            <div class="preview-cell" style="background-color: ${theme.colors.cellBackground}"></div>
            <div class="preview-cell" style="background-color: ${theme.colors.cellBackground}"></div>
            <div class="preview-cell highlighted" style="background-color: ${theme.colors.accent}"></div>
          </div>
        </div>
        <div class="theme-name">${theme.name}</div>
      `;
      
      themeGrid.appendChild(themeCard);
    });
    
    dialogContent.appendChild(themeGrid);
    
    // Add custom theme button
    const customizeButton = document.createElement('button');
    customizeButton.className = 'customize-theme-button';
    customizeButton.innerHTML = '<i class="fas fa-palette"></i> Customize Theme';
    customizeButton.addEventListener('click', () => {
      this.showCustomThemeDialog();
      themeDialog.remove();
    });
    
    dialogContent.appendChild(customizeButton);
    
    themeDialog.appendChild(dialogContent);
    document.body.appendChild(themeDialog);
    
    // Handle close button
    themeDialog.querySelector('.theme-dialog-close').addEventListener('click', () => {
      themeDialog.remove();
    });
    
    // Close when clicking outside
    themeDialog.addEventListener('click', (e) => {
      if (e.target === themeDialog) {
        themeDialog.remove();
      }
    });
    
    // Add animation
    setTimeout(() => {
      themeDialog.classList.add('show');
    }, 10);
  }

  /**
   * Show dialog to create a custom theme
   */
  showCustomThemeDialog() {
    const dialog = document.createElement('div');
    dialog.className = 'theme-selector-dialog custom-theme-dialog';
    dialog.id = 'custom-theme-dialog';
    
    const content = document.createElement('div');
    content.className = 'theme-selector-content';
    
    // Add header
    const header = document.createElement('div');
    header.className = 'theme-selector-header';
    header.innerHTML = `
      <h3>Create Custom Theme</h3>
      <button class="theme-dialog-close"><i class="fas fa-times"></i></button>
    `;
    content.appendChild(header);
    
    // Add form
    const form = document.createElement('div');
    form.className = 'custom-theme-form';
    
    form.innerHTML = `
      <div class="form-group">
        <label for="custom-theme-name">Theme Name</label>
        <input type="text" id="custom-theme-name" placeholder="My Custom Theme">
      </div>
      
      <div class="color-pickers">
        <div class="form-group">
          <label for="primary-color">Primary Color</label>
          <input type="color" id="primary-color" value="#217346">
        </div>
        
        <div class="form-group">
          <label for="secondary-color">Secondary Color</label>
          <input type="color" id="secondary-color" value="#2b579a">
        </div>
        
        <div class="form-group">
          <label for="accent-color">Accent Color</label>
          <input type="color" id="accent-color" value="#ed7d31">
        </div>
        
        <div class="form-group">
          <label for="background-color">Background</label>
          <input type="color" id="background-color" value="#ffffff">
        </div>
        
        <div class="form-group">
          <label for="text-color">Text</label>
          <input type="color" id="text-color" value="#212529">
        </div>
      </div>
      
      <div class="preview-section">
        <h4>Live Preview</h4>
        <div class="theme-preview" id="live-preview">
          <div class="preview-header"></div>
          <div class="preview-toolbar"></div>
          <div class="preview-content">
            <div class="preview-cell"></div>
            <div class="preview-cell"></div>
            <div class="preview-cell highlighted"></div>
          </div>
        </div>
      </div>
      
      <div class="dialog-buttons">
        <button id="save-custom-theme" class="button">Save Theme</button>
        <button id="cancel-custom-theme" class="button button-secondary">Cancel</button>
      </div>
    `;
    
    content.appendChild(form);
    dialog.appendChild(content);
    document.body.appendChild(dialog);
    
    // Initialize color values
    const primaryColor = document.getElementById('primary-color');
    const secondaryColor = document.getElementById('secondary-color');
    const accentColor = document.getElementById('accent-color');
    const backgroundColor = document.getElementById('background-color');
    const textColor = document.getElementById('text-color');
    
    // Handle live preview
    const updatePreview = () => {
      const preview = document.getElementById('live-preview');
      
      if (preview) {
        const previewHeader = preview.querySelector('.preview-header');
        const previewToolbar = preview.querySelector('.preview-toolbar');
        const previewCells = preview.querySelectorAll('.preview-cell');
        const previewHighlighted = preview.querySelector('.preview-cell.highlighted');
        
        preview.style.backgroundColor = backgroundColor.value;
        preview.style.color = textColor.value;
        preview.style.borderColor = this.adjustColor(backgroundColor.value, -20);
        
        previewHeader.style.backgroundColor = primaryColor.value;
        previewToolbar.style.backgroundColor = this.adjustColor(backgroundColor.value, -5);
        
        previewCells.forEach(cell => {
          cell.style.backgroundColor = backgroundColor.value;
          cell.style.borderColor = this.adjustColor(backgroundColor.value, -15);
        });
        
        previewHighlighted.style.backgroundColor = accentColor.value;
      }
    };
    
    // Attach event listeners for live preview
    [primaryColor, secondaryColor, accentColor, backgroundColor, textColor].forEach(input => {
      input.addEventListener('input', updatePreview);
    });
    
    // Initial preview update
    updatePreview();
    
    // Handle close button
    dialog.querySelector('.theme-dialog-close').addEventListener('click', () => {
      dialog.remove();
    });
    
    // Handle cancel button
    document.getElementById('cancel-custom-theme').addEventListener('click', () => {
      dialog.remove();
    });
    
    // Handle save button
    document.getElementById('save-custom-theme').addEventListener('click', () => {
      const name = document.getElementById('custom-theme-name').value.trim() || 'Custom Theme';
      
      const customColors = {
        primary: primaryColor.value,
        secondary: secondaryColor.value,
        accent: accentColor.value,
        background: backgroundColor.value,
        text: textColor.value,
        border: this.adjustColor(backgroundColor.value, -20),
        hover: this.adjustColor(backgroundColor.value, -5),
        cellBackground: backgroundColor.value,
        headerBackground: this.adjustColor(backgroundColor.value, -5),
        toolbarBackground: this.adjustColor(backgroundColor.value, -2)
      };
      
      if (this.addCustomTheme(name, customColors)) {
        this.applyTheme(name.toLowerCase().replace(/\s+/g, '-'));
        dialog.remove();
        
        // Show success message
        this.showNotification('Custom theme created and applied!');
      } else {
        // Show error message
        this.showNotification('Theme name already exists. Please choose a different name.', 'error');
      }
    });
    
    // Add animation
    setTimeout(() => {
      dialog.classList.add('show');
    }, 10);
  }

  /**
   * Adjust a color's lightness
   * @param {string} color - Hex color
   * @param {number} percent - Percent to adjust (-100 to 100)
   * @returns {string} - Adjusted hex color
   */
  adjustColor(color, percent) {
    let R = parseInt(color.substring(1, 3), 16);
    let G = parseInt(color.substring(3, 5), 16);
    let B = parseInt(color.substring(5, 7), 16);

    R = parseInt(R * (100 + percent) / 100);
    G = parseInt(G * (100 + percent) / 100);
    B = parseInt(B * (100 + percent) / 100);

    R = (R < 255) ? R : 255;
    G = (G < 255) ? G : 255;
    B = (B < 255) ? B : 255;

    R = (R > 0) ? R : 0;
    G = (G > 0) ? G : 0;
    B = (B > 0) ? B : 0;

    const RR = ((R.toString(16).length === 1) ? '0' + R.toString(16) : R.toString(16));
    const GG = ((G.toString(16).length === 1) ? '0' + G.toString(16) : G.toString(16));
    const BB = ((B.toString(16).length === 1) ? '0' + B.toString(16) : B.toString(16));

    return '#' + RR + GG + BB;
  }

  /**
   * Show notification message
   * @param {string} message - Message to display
   * @param {string} type - Message type (success, error)
   */
  showNotification(message, type = 'success') {
    let notification = document.createElement('div');
    notification.className = `theme-notification ${type}`;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    // Show with animation
    setTimeout(() => {
      notification.classList.add('show');
    }, 10);
    
    // Remove after delay
    setTimeout(() => {
      notification.classList.remove('show');
      setTimeout(() => {
        notification.remove();
      }, 300);
    }, 3000);
  }

  /**
   * Create a theme selection dropdown in the specified container
   * @param {HTMLElement} container - Container element for the theme dropdown
   */
  createThemeSelector(container) {
    if (!container) return;
    
    const dropdown = document.createElement('div');
    dropdown.className = 'theme-dropdown';
    
    const button = document.createElement('button');
    button.className = 'theme-button';
    
    // Create theme indicator
    const currentTheme = this.themes[this.currentTheme];
    button.innerHTML = `
      <i class="fas fa-palette"></i>
      <span class="current-theme-indicator">
        ${currentTheme.name}
        <span class="theme-color-dot" style="background-color: ${currentTheme.colors.primary}"></span>
      </span>
      <i class="fas fa-caret-down"></i>
    `;
    
    const content = document.createElement('div');
    content.className = 'theme-dropdown-content';
    
    // Add theme options
    Object.keys(this.themes).forEach(themeName => {
      const theme = this.themes[themeName];
      const option = document.createElement('div');
      option.className = 'theme-option';
      option.dataset.theme = themeName;
      
      if (themeName === this.currentTheme) {
        option.classList.add('active');
      }
      
      const colorSwatch = document.createElement('span');
      colorSwatch.className = 'theme-swatch';
      colorSwatch.style.backgroundColor = theme.colors.primary;
      
      const label = document.createElement('span');
      label.textContent = theme.name;
      
      option.appendChild(colorSwatch);
      option.appendChild(label);
      content.appendChild(option);
    });
    
    // Add customize option
    const customizeOption = document.createElement('div');
    customizeOption.className = 'theme-option customize-option';
    customizeOption.innerHTML = `
      <i class="fas fa-sliders-h"></i>
      <span>Customize Themes...</span>
    `;
    customizeOption.addEventListener('click', (e) => {
      e.stopPropagation();
      this.toggleThemeSelector();
      
      // Close the dropdown
      const dropdownContent = e.target.closest('.theme-dropdown-content');
      if (dropdownContent) {
        dropdownContent.style.display = 'none';
        setTimeout(() => {
          dropdownContent.style.display = '';
        }, 100);
      }
    });
    
    content.appendChild(document.createElement('hr'));
    content.appendChild(customizeOption);
    
    dropdown.appendChild(button);
    dropdown.appendChild(content);
    container.appendChild(dropdown);
    
    // Add keyboard shortcut hint
    const shortcutHint = document.createElement('div');
    shortcutHint.className = 'shortcut-hint';
    shortcutHint.textContent = 'Alt+Shift+T';
    dropdown.appendChild(shortcutHint);
  }

  /**
   * Get the current theme
   * @returns {object} - Current theme object
   */
  getCurrentTheme() {
    return this.themes[this.currentTheme];
  }

  /**
   * Add a new custom theme
   * @param {string} name - Theme name
   * @param {object} colors - Theme colors
   * @returns {boolean} - Success status
   */
  addCustomTheme(name, colors) {
    const themeId = name.toLowerCase().replace(/\s+/g, '-');
    
    if (this.themes[themeId]) {
      return false; // Theme already exists
    }
    
    this.themes[themeId] = {
      name: name,
      colors: {
        primary: colors.primary || '#217346',
        secondary: colors.secondary || '#2b579a',
        accent: colors.accent || '#ed7d31',
        background: colors.background || '#ffffff',
        text: colors.text || '#212529',
        border: colors.border || '#d4d4d4',
        hover: colors.hover || '#f3f3f3',
        cellBackground: colors.cellBackground || '#ffffff',
        headerBackground: colors.headerBackground || '#f3f2f1',
        toolbarBackground: colors.toolbarBackground || '#ffffff'
      }
    };
    
    // Save custom themes to localStorage
    this.saveCustomThemes();
    
    return true;
  }

  /**
   * Save custom themes to localStorage
   */
  saveCustomThemes() {
    const customThemes = {};
    
    // Filter out built-in themes
    Object.keys(this.themes).forEach(themeId => {
      if (!['light', 'dark', 'excel365', 'classic', 'modern', 'midnight', 'solarized', 'highContrast'].includes(themeId)) {
        customThemes[themeId] = this.themes[themeId];
      }
    });
    
    localStorage.setItem('spreadsheet-custom-themes', JSON.stringify(customThemes));
  }

  /**
   * Load custom themes from localStorage
   */
  loadCustomThemes() {
    const savedThemes = localStorage.getItem('spreadsheet-custom-themes');
    
    if (savedThemes) {
      try {
        const customThemes = JSON.parse(savedThemes);
        
        Object.keys(customThemes).forEach(themeId => {
          this.themes[themeId] = customThemes[themeId];
        });
      } catch (error) {
        console.error('Error loading custom themes:', error);
      }
    }
  }

  /**
   * Export all themes to JSON
   * @returns {string} - JSON string of all themes
   */
  exportThemes() {
    return JSON.stringify(this.themes, null, 2);
  }

  /**
   * Import themes from JSON
   * @param {string} json - JSON string of themes
   * @returns {boolean} - Success status
   */
  importThemes(json) {
    try {
      const importedThemes = JSON.parse(json);
      
      // Validate themes structure
      for (const [id, theme] of Object.entries(importedThemes)) {
        if (!theme.name || !theme.colors) {
          throw new Error(`Invalid theme structure for "${id}"`);
        }
      }
      
      // Save builtin themes
      const builtinThemeIds = ['light', 'dark', 'excel365', 'classic', 'modern', 'midnight', 'solarized', 'highContrast'];
      const builtinThemes = {};
      
      builtinThemeIds.forEach(id => {
        if (this.themes[id]) {
          builtinThemes[id] = this.themes[id];
        }
      });
      
      // Replace themes with imported ones, but keep builtin themes
      this.themes = { ...importedThemes, ...builtinThemes };
      
      // Save custom themes
      this.saveCustomThemes();
      
      // Reapply current theme
      this.applyTheme(this.currentTheme);
      
      return true;
    } catch (error) {
      console.error('Error importing themes:', error);
      return false;
    }
  }
}

// Add ThemeManager to the window object
window.ThemeManager = ThemeManager;
