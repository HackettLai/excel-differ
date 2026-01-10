// fileHandler.js - Handle file upload and reading

const FileHandler = {
  fileA: null,
  fileB: null,

  /**
   * Initialize file upload listeners and drag-and-drop support
   */
  init() {
    const fileInputA = document.getElementById('fileA');
    const fileInputB = document.getElementById('fileB');
    const compareBtn = document.getElementById('compareBtn');

    fileInputA.addEventListener('change', (e) => {
      this.handleFileSelect(e, 'A');
    });

    fileInputB.addEventListener('change', (e) => {
      this.handleFileSelect(e, 'B');
    });

    // Set up drag-and-drop support
    this.setupDragAndDrop('uploadA', fileInputA);
    this.setupDragAndDrop('uploadB', fileInputB);
  },

  /**
   * Handle file selection event
   * @param {Event} event - File input change event
   * @param {string} fileType - 'A' or 'B' to identify which file slot
   */
  handleFileSelect(event, fileType) {
    const file = event.target.files[0];

    if (!file) return;

    // Validate file type
    if (!this.isValidFile(file)) {
      alert('Please select a valid Excel file (.xlsx, .xls, .csv)');
      return;
    }

    // Validate file size (limit 50MB)
    if (file.size > 50 * 1024 * 1024) {
      alert('File too large! Please select a file smaller than 50MB');
      return;
    }

    // Save file reference
    if (fileType === 'A') {
      this.fileA = file;
      document.getElementById('fileNameA').textContent = file.name;
    } else {
      this.fileB = file;
      document.getElementById('fileNameB').textContent = file.name;
    }

    // Check if compare button should be enabled
    this.checkCompareButton();
  },

  /**
   * Validate if file has a valid Excel extension
   * @param {File} file - File object to validate
   * @returns {boolean} True if file extension is valid
   */
  isValidFile(file) {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const fileName = file.name.toLowerCase();
    return validExtensions.some((ext) => fileName.endsWith(ext));
  },

  /**
   * Enable or disable the compare button based on file selection
   * Compare button is enabled only when both files are selected
   */
  checkCompareButton() {
    const compareBtn = document.getElementById('compareBtn');

    if (this.fileA && this.fileB) {
      compareBtn.disabled = false;
    } else {
      compareBtn.disabled = true;
    }
  },

  /**
   * Set up drag-and-drop functionality for file upload area
   * @param {string} uploadAreaId - ID of the upload area element
   * @param {HTMLInputElement} fileInput - Associated file input element
   */
  setupDragAndDrop(uploadAreaId, fileInput) {
    const uploadArea = document.getElementById(uploadAreaId);

    // Prevent default drag-and-drop behavior
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach((eventName) => {
      uploadArea.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
      });
    });

    // Highlight drop area when dragging over
    ['dragenter', 'dragover'].forEach((eventName) => {
      uploadArea.addEventListener(eventName, () => {
        uploadArea.classList.add('drag-over');
      });
    });

    // Remove highlight when dragging leaves or file is dropped
    ['dragleave', 'drop'].forEach((eventName) => {
      uploadArea.addEventListener(eventName, () => {
        uploadArea.classList.remove('drag-over');
      });
    });

    // Handle file drop
    uploadArea.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;
      if (files.length > 0) {
        // Assign dropped files to file input and trigger change event
        fileInput.files = files;
        fileInput.dispatchEvent(new Event('change'));
      }
    });
  },

  /**
   * Read a file as ArrayBuffer
   * @param {File} file - File object to read
   * @returns {Promise<ArrayBuffer>} Promise that resolves with file data
   */
  readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        resolve(e.target.result);
      };

      reader.onerror = (e) => {
        reject(new Error('Failed to read file'));
      };

      reader.readAsArrayBuffer(file);
    });
  },

  /**
   * Get file data for comparison
   * @param {string} fileType - 'A' or 'B' to identify which file
   * @returns {Promise<Object>} Object containing file and its ArrayBuffer data
   */
  async getFileData(fileType) {
    const file = fileType === 'A' ? this.fileA : this.fileB;

    if (!file) {
      throw new Error(`File ${fileType} not selected`);
    }

    return {
      file: file,
      arrayBuffer: await this.readFile(file),
    };
  },

  /**
   * Reset file selection and UI
   * Clears both file slots and disables compare button
   */
  reset() {
    this.fileA = null;
    this.fileB = null;

    document.getElementById('fileA').value = '';
    document.getElementById('fileB').value = '';
    document.getElementById('fileNameA').textContent = 'No file selected';
    document.getElementById('fileNameB').textContent = 'No file selected';

    this.checkCompareButton();
  },
};
