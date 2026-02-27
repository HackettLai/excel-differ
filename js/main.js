// main.js
// Entry point of the Excel Differ application
// Orchestrates all modules and manages the overall application flow

import FileHandler from './fileHandler.js';
import ExcelParser from './excelParser.js';
import DiffEngine from './diffEngine.js';
import DiffViewer from './diffViewer.js';
import Copyright from './copyright.js';

/**
 * ExcelDiffer Class
 * Main application controller that coordinates all components
 */
class ExcelDiffer {
  /**
   * Constructor
   * Initializes all module instances and application state
   */
  constructor() {
    // Initialize all module instances
    this.fileHandler = new FileHandler(); // Handles file uploads and UI
    this.excelParser = new ExcelParser(); // Parses Excel files
    this.diffEngine = new DiffEngine(); // Compares Excel data
    this.diffViewer = new DiffViewer(); // Renders diff results
    this.copyright = new Copyright(); // Handles copyright display

    // Application state
    this.fileA = null; // Stores uploaded File A object
    this.fileB = null; // Stores uploaded File B object
    this.dataA = null; // Parsed data from File A
    this.dataB = null; // Parsed data from File B
    this.diffResults = null; // Comparison results
  }

  /**
   * init()
   * Initializes the application on page load
   * Called by DOMContentLoaded event listener
   */
  init() {
    // Initialize FileHandler with callback for file uploads
    this.fileHandler.init(this.handleFileUpload.bind(this));

    // Initialize Copyright module
    this.copyright.init();

    // Set up event listeners for UI interactions
    this.setupEventListeners();
  }

  /**
   * setupEventListeners()
   * Attaches event listeners to UI elements
   * Handles button clicks for main application flow
   */
  setupEventListeners() {
    // Get DOM elements
    const startCompareBtn = document.getElementById('startCompareBtn');
    const backBtn = document.getElementById('backBtn');
    const compareSheetBtn = document.getElementById('compareSheetBtn');
    const prevBtn = document.getElementById('prevChangeBtn');
    const nextBtn = document.getElementById('nextChangeBtn');

    // "Start Comparing" button - triggers file comparison
    if (startCompareBtn) {
      startCompareBtn.addEventListener('click', () => this.compareFiles());
    }

    // "Back" button - returns to upload view
    if (backBtn) {
      backBtn.addEventListener('click', () => this.reset());
    }

    // "Compare" button - compares selected sheets
    if (compareSheetBtn) {
      compareSheetBtn.addEventListener('click', () => this.handleCompareSheets());
    }
  }

  /**
   * handleFileUpload(file, type)
   * Callback function triggered when user uploads a file
   *
   * @param {File} file - The uploaded file object
   * @param {string} type - Either 'A' or 'B' indicating which file slot
   */
  handleFileUpload(file, type) {
    // Store file in appropriate slot
    if (type === 'A') {
      this.fileA = file;
    } else if (type === 'B') {
      this.fileB = file;
    }

    // Enable "Start Comparing" button only when both files are uploaded
    const startCompareBtn = document.getElementById('startCompareBtn');
    if (this.fileA && this.fileB && startCompareBtn) {
      startCompareBtn.disabled = false;
    }
  }

  /**
   * compareFiles()
   * Main comparison workflow that parses and compares both Excel files.
   * Handles file type detection and adapts the comparison flow accordingly.
   *
   * Process:
   * 1. Shows loading overlay
   * 2. Validates both files exist and are File objects
   * 3. Parses File A using excelParser
   * 4. Parses File B using excelParser
   * 5. Determines file types and comparison mode
   * 6. For Excel files: compares all sheets automatically
   * 7. For mixed file types: initializes UI then auto-triggers first sheet comparison
   * 8. Switches UI from upload view to diff view
   *
   * @throws {Error} If files are missing or invalid
   */
  async compareFiles() {
    const loadingOverlay = document.getElementById('loadingOverlay');

    try {
      // Show loading overlay
      if (loadingOverlay) loadingOverlay.style.display = 'flex';

      // Validate both files exist
      if (!this.fileA || !this.fileB) {
        throw new Error('Please select both files');
      }

      // Validate files are File objects
      if (!(this.fileA instanceof File)) {
        throw new Error('File A is not a valid File object');
      }

      if (!(this.fileB instanceof File)) {
        throw new Error('File B is not a valid File object');
      }

      // Parse File A
      console.log('📂 Parsing File A:', this.fileA.name);
      this.dataA = await this.excelParser.parse(this.fileA);

      // Parse File B
      console.log('📂 Parsing File B:', this.fileB.name);
      this.dataB = await this.excelParser.parse(this.fileB);

      // Determine file types and select comparison mode
      const bothExcel = this.fileA.name.endsWith('.xlsx') && this.fileB.name.endsWith('.xlsx');

      if (bothExcel) {
        // Excel vs Excel: automatically compare all sheets
        console.log('📊 Mode: Excel vs Excel (auto-compare all sheets)');

        this.diffResults = this.diffEngine.compare(this.dataA, this.dataB);

        // Create fresh DiffViewer instance
        this.diffViewer = new DiffViewer();
        this.diffViewer.init(this.dataA, this.dataB, this.diffResults);
      } else {
        // Mixed file types: initialize UI then automatically compare first sheet
        console.log('📊 Mode: Mixed (CSV/Excel) - auto-compare first sheet');

        // Create fresh DiffViewer instance and initialize UI
        this.diffViewer = new DiffViewer();
        this.diffViewer.init(this.dataA, this.dataB, null);

        // Wait for UI to render, then automatically trigger comparison
        setTimeout(() => {
          const compareBtn = document.getElementById('compareSheetBtn');
          if (compareBtn) {
            console.log('Automatically triggering comparison...');
            compareBtn.click();
          } else {
            console.error('Compare button not found');
          }
        }, 150);
      }

      // Switch to diff view
      this.hideUploadSection();
      this.showDiffSection();
    } catch (error) {
      // Error handling
      console.error('❌ Error comparing files:', error);
      alert('Error comparing files: ' + error.message);
    } finally {
      // Always hide loading overlay
      if (loadingOverlay) loadingOverlay.style.display = 'none';
    }
  }

  /**
   * handleCompareSheets()
   * Triggers comparison of currently selected sheets
   * Called when user clicks "Compare" button in diff view
   */
  handleCompareSheets() {
    if (this.diffViewer) {
      this.diffViewer.compareSelectedSheets();
    }
  }

  /**
   * reset()
   * Resets application to initial state
   * Clears all data and returns to upload view
   * Called when user clicks "Back" button
   */
  reset() {
    // Clear all stored data
    this.fileA = null;
    this.fileB = null;
    this.dataA = null;
    this.dataB = null;
    this.diffResults = null;

    // Reset DiffViewer state by clearing all stored data and rendered UI
    if (this.diffViewer) {
      this.diffViewer.dataA = null;
      this.diffViewer.dataB = null;
      this.diffViewer.diffResults = null;
      this.diffViewer.changedCells = [];
      this.diffViewer.currentChangeIndex = -1;

      // Clear all dropdown selectors
      this.clearDiffViewDropdowns();

      // Clear the comparison results table
      const tableContainer = document.getElementById('unifiedTableContainer');
      if (tableContainer) {
        tableContainer.innerHTML = '';
      }
    }

    // Reset FileHandler (clears file inputs)
    this.fileHandler.reset();

    // Switch back to upload view
    this.hideDiffSection();
    this.showUploadSection();

    // Disable "Start Comparing" button
    const startCompareBtn = document.getElementById('startCompareBtn');
    if (startCompareBtn) startCompareBtn.disabled = true;
  }

  /**
   * clearDiffViewDropdowns()
   * Resets all dropdown selectors in the diff view to their initial state
   */
  clearDiffViewDropdowns() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');
    const keyColumnSelect = document.getElementById('keyColumnSelect');

    if (sheetSelectA) sheetSelectA.innerHTML = '';
    if (sheetSelectB) sheetSelectB.innerHTML = '';
    if (headerRowA) headerRowA.innerHTML = '';
    if (headerRowB) headerRowB.innerHTML = '';
    if (keyColumnSelect) keyColumnSelect.innerHTML = '<option value="">-- Select Key Column --</option>';
  }

  /**
   * hideUploadSection()
   * Hides the file upload UI section
   * Sets display to 'none' for upload section and header
   */
  hideUploadSection() {
    const uploadSection = document.getElementById('uploadSection');
    if (uploadSection) uploadSection.style.display = 'none';

    const header = document.getElementById('header');
    if (header) header.style.display = 'none';
  }

  /**
   * showUploadSection()
   * Shows the file upload UI section
   * Sets display to 'flex' for upload section and 'block' for header
   */
  showUploadSection() {
    const uploadSection = document.getElementById('uploadSection');
    if (uploadSection) uploadSection.style.display = 'flex';

    const header = document.getElementById('header');
    if (header) header.style.display = 'block';
  }

  /**
   * hideDiffSection()
   * Hides the diff comparison UI section
   * Sets display to 'none' for diff section
   */
  hideDiffSection() {
    const diffSection = document.getElementById('diffSection');
    if (diffSection) diffSection.style.display = 'none';
  }

  /**
   * showDiffSection()
   * Shows the diff comparison UI section
   * Sets display to 'flex' for diff section
   */
  showDiffSection() {
    const diffSection = document.getElementById('diffSection');
    if (diffSection) diffSection.style.display = 'flex';
  }
}

/**
 * DOMContentLoaded Event Listener
 * Initializes the application when DOM is fully loaded
 * Creates ExcelDiffer instance and exposes it globally
 */
document.addEventListener('DOMContentLoaded', () => {
  // Create and initialize application instance
  const app = new ExcelDiffer();
  app.init();

  // Expose to global scope for debugging/testing
  window.excelDiffer = app;
});
