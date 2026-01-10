// main.js - Main application flow controller (Safari compatible)

const App = {
  /**
   * Initialize the application
   * Sets up file handlers and event bindings
   */
  init() {
    console.log('📊 Excel Differ initializing...');

    // Check if SheetJS library is loaded
    if (typeof XLSX === 'undefined') {
      alert('Error: SheetJS library not loaded!');
      return;
    }

    // Initialize file handler
    FileHandler.init();

    // Bind events
    this.bindEvents();

    console.log('✅ Initialization complete');
  },

  /**
   * Bind event listeners to UI elements
   */
  bindEvents() {
    // Compare button
    document.getElementById('compareBtn').addEventListener('click', () => {
      this.startComparison();
    });

    // Back to upload button
    document.getElementById('backBtn').addEventListener('click', () => {
      this.backToUpload();
    });

    // Back to summary button
    document.getElementById('backToSummaryBtn').addEventListener('click', () => {
      this.backToSummary();
    });
  },

  /**
   * Wait for browser to complete rendering
   * This fixes Safari rendering issues by ensuring UI updates are complete
   * @returns {Promise<void>} Resolves when rendering is complete
   */
  waitForRender() {
    return new Promise((resolve) => {
      // Use setTimeout to ensure UI update is complete
      setTimeout(() => {
        // Use requestAnimationFrame to ensure next frame is rendered
        requestAnimationFrame(() => {
          resolve();
        });
      }, 0);
    });
  },

  /**
   * Start the file comparison process
   * Reads files, parses Excel data, performs diff, and displays results
   */
  async startComparison() {
    try {
      // Show loading overlay
      this.showLoading(true);

      // Wait for loading overlay to render (Safari fix)
      await this.waitForRender();

      // Read files
      console.log('📖 Reading files...');
      const fileDataA = await FileHandler.getFileData('A');
      const fileDataB = await FileHandler.getFileData('B');

      // Wait again to ensure UI is stable
      await this.waitForRender();

      // Parse Excel files
      console.log('🔍 Parsing Excel...');
      const parsedA = ExcelParser.parse(fileDataA.arrayBuffer, fileDataA.file.name);
      const parsedB = ExcelParser.parse(fileDataB.arrayBuffer, fileDataB.file.name);

      console.log('File A:', parsedA);
      console.log('File B:', parsedB);

      // Wait for UI update after parsing
      await this.waitForRender();

      // Perform diff comparison
      console.log('⚡ Performing diff comparison...');
      const diffResult = DiffEngine.compare(parsedA, parsedB);

      console.log('Diff result:', diffResult);

      // Hide loading overlay
      this.showLoading(false);

      // Wait for loading overlay to hide
      await this.waitForRender();

      // Display summary view
      SummaryView.show(diffResult, parsedA, parsedB);
    } catch (error) {
      this.showLoading(false);
      console.error('Error during comparison:', error);
      alert(`Error: ${error.message}`);
    }
  },

  /**
   * Return to the file upload page
   * Hides summary and diff views, resets file selection
   */
  backToUpload() {
    // Hide summary
    SummaryView.hide();

    // Hide diff viewer
    DiffViewer.hide();

    // Show upload section
    document.getElementById('uploadSection').style.display = 'block';

    // Reset files
    FileHandler.reset();
  },

  /**
   * Return to the summary page from diff view
   */
  backToSummary() {
    // Hide diff viewer
    DiffViewer.hide();

    // Show summary section
    document.getElementById('summarySection').style.display = 'block';
  },

  /**
   * Show or hide the loading overlay
   * @param {boolean} show - True to show, false to hide
   */
  showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    overlay.style.display = show ? 'flex' : 'none';
  },
};

// Initialize app when DOM is fully loaded
document.addEventListener('DOMContentLoaded', () => {
  App.init();
});

// Prevent accidental page close when files are loaded
window.addEventListener('beforeunload', (e) => {
  if (FileHandler.fileA || FileHandler.fileB) {
    e.preventDefault();
    e.returnValue = 'Are you sure you want to leave? Unsaved comparison results will be lost.';
  }
});
