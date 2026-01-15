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
   * Main comparison workflow - parses and compares both Excel files
   * Async function that handles the entire comparison process
   *
   * Process:
   * 1. Shows loading overlay
   * 2. Validates both files exist and are File objects
   * 3. Parses File A using excelParser
   * 4. Parses File B using excelParser
   * 5. Compares parsed data using diffEngine
   * 6. Initializes DiffViewer with results
   * 7. Switches UI from upload view to diff view
   *
   * @throws {Error} If files are missing or invalid
   */
  async compareFiles() {
    const loadingOverlay = document.getElementById('loadingOverlay');

    try {
      // Show loading overlay
      if (loadingOverlay) loadingOverlay.style.display = 'flex';

      // Debug logging
      console.log('File A:', this.fileA);
      console.log('File B:', this.fileB);

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
      console.log('Parsing File A...');
      this.dataA = await this.excelParser.parse(this.fileA);
      console.log('File A parsed:', this.dataA);

      // Parse File B
      console.log('Parsing File B...');
      this.dataB = await this.excelParser.parse(this.fileB);
      console.log('File B parsed:', this.dataB);

      // Compare the two files
      console.log('Comparing files...');
      this.diffResults = this.diffEngine.compare(this.dataA, this.dataB);
      console.log('Diff results:', this.diffResults);

      // Initialize DiffViewer with comparison results
      this.diffViewer.init(this.dataA, this.dataB, this.diffResults);

      // Switch to diff view
      this.hideUploadSection();
      this.showDiffSection();
    } catch (error) {
      // Error handling
      console.error('Error comparing files:', error);
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
