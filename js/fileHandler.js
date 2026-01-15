// fileHandler.js
// Handles file upload functionality including click-to-upload and drag-and-drop
// Manages file selection UI state and validation

/**
 * FileHandler Class
 * Manages file upload interactions for both File A and File B
 * Supports both click-to-upload and drag-and-drop functionality
 */
class FileHandler {
  /**
   * Constructor
   * Initializes file handler state
   */
  constructor() {
    // DOM element references
    this.fileInputA = null; // Hidden file input for File A
    this.fileInputB = null; // Hidden file input for File B
    this.uploadBoxA = null; // Upload box UI element for File A
    this.uploadBoxB = null; // Upload box UI element for File B

    // File storage
    this.fileA = null; // Stores File A object
    this.fileB = null; // Stores File B object

    // Callback function
    this.onFileUploadCallback = null; // Callback to notify parent when file is uploaded
  }

  /**
   * init(onFileUpload)
   * Initializes the FileHandler with DOM elements and callback
   *
   * @param {Function} onFileUpload - Callback function(file, type) to be called when file is selected
   */
  init(onFileUpload) {
    // Store callback function
    this.onFileUploadCallback = onFileUpload;

    // Get DOM element references
    this.fileInputA = document.getElementById('fileInputA');
    this.fileInputB = document.getElementById('fileInputB');
    this.uploadBoxA = document.getElementById('uploadBoxA');
    this.uploadBoxB = document.getElementById('uploadBoxB');

    // Set up event listeners
    this.setupEventListeners();
  }

  /**
   * setupEventListeners()
   * Attaches event listeners to file input elements and upload boxes
   * Handles click-to-upload, drag-and-drop, and file removal
   */
  setupEventListeners() {
    // ========== File Input A ==========
    if (this.uploadBoxA && this.fileInputA) {
      // Click on upload box triggers hidden file input
      this.uploadBoxA.addEventListener('click', () => {
        this.fileInputA.click();
      });

      // Handle file selection from file input
      this.fileInputA.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
          this.handleFileSelect(file, 'A');
        }
      });

      // Enable drag-and-drop for upload box A
      this.setupDragAndDrop(this.uploadBoxA, 'A');
    }

    // ========== File Input B ==========
    if (this.uploadBoxB && this.fileInputB) {
      // Click on upload box triggers hidden file input
      this.uploadBoxB.addEventListener('click', () => {
        this.fileInputB.click();
      });

      // Handle file selection from file input
      this.fileInputB.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
          this.handleFileSelect(file, 'B');
        }
      });

      // Enable drag-and-drop for upload box B
      this.setupDragAndDrop(this.uploadBoxB, 'B');
    }

    // ========== Remove File A Button ==========
    const removeFileA = document.getElementById('removeFileA');
    if (removeFileA) {
      removeFileA.addEventListener('click', (e) => {
        e.stopPropagation(); // Prevent triggering upload box click
        this.removeFile('A');
      });
    }

    // ========== Remove File B Button ==========
    const removeFileB = document.getElementById('removeFileB');
    if (removeFileB) {
      removeFileB.addEventListener('click', (e) => {
        e.stopPropagation(); // Prevent triggering upload box click
        this.removeFile('B');
      });
    }
  }

  /**
   * setupDragAndDrop(uploadBox, type)
   * Configures drag-and-drop functionality for an upload box
   *
   * @param {HTMLElement} uploadBox - The upload box element to enable drag-and-drop
   * @param {string} type - Either 'A' or 'B' indicating which file slot
   */
  setupDragAndDrop(uploadBox, type) {
    if (!uploadBox) return;

    // Prevent default browser behavior for drag events
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach((eventName) => {
      uploadBox.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
      });
    });

    // Add 'active' class when dragging over upload box (visual feedback)
    ['dragenter', 'dragover'].forEach((eventName) => {
      uploadBox.addEventListener(eventName, () => {
        uploadBox.classList.add('active');
      });
    });

    // Remove 'active' class when drag leaves or file is dropped
    ['dragleave', 'drop'].forEach((eventName) => {
      uploadBox.addEventListener(eventName, () => {
        uploadBox.classList.remove('active');
      });
    });

    // Handle file drop
    uploadBox.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;

      // Process first file if any files were dropped
      if (files.length > 0) {
        const file = files[0];
        this.handleFileSelect(file, type);
      }
    });
  }

  /**
   * handleFileSelect(file, type)
   * Processes a selected file, validates it, and stores it
   *
   * @param {File} file - The file object to be processed
   * @param {string} type - Either 'A' or 'B' indicating which file slot
   */
  handleFileSelect(file, type) {
    // Debug logging
    console.log('File selected:', file.name, 'Type:', type, 'File object:', file);

    // Validate file exists
    if (!file) {
      console.error('No file provided');
      return;
    }

    // Validate file extension (must be .xlsx or .xls)
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      alert('Please select a valid Excel file (.xlsx or .xls)');
      return;
    }

    // Store file in appropriate slot
    if (type === 'A') {
      this.fileA = file;
      this.updateFileInfo('A', file.name);
    } else if (type === 'B') {
      this.fileB = file;
      this.updateFileInfo('B', file.name);
    }

    // Notify parent application via callback
    if (this.onFileUploadCallback) {
      this.onFileUploadCallback(file, type); // Pass actual File object
    }
  }

  /**
   * updateFileInfo(type, fileName)
   * Updates the UI to show the selected file name
   *
   * @param {string} type - Either 'A' or 'B' indicating which file slot
   * @param {string} fileName - Name of the selected file
   */
  updateFileInfo(type, fileName) {
    const fileInfo = document.getElementById(`fileInfo${type}`);
    const fileNameEl = document.getElementById(`fileName${type}`);
    const uploadBox = document.getElementById(`uploadBox${type}`);

    if (fileInfo && fileNameEl && uploadBox) {
      // Display file name
      fileNameEl.textContent = fileName;

      // Show file info section
      fileInfo.style.display = 'flex';

      // Add visual indicator that file is selected
      uploadBox.classList.add('file-selected');
    }
  }

  /**
   * removeFile(type)
   * Removes a selected file and resets UI state
   *
   * @param {string} type - Either 'A' or 'B' indicating which file slot
   */
  removeFile(type) {
    // Clear file storage
    if (type === 'A') {
      this.fileA = null;
      if (this.fileInputA) this.fileInputA.value = ''; // Reset file input
    } else if (type === 'B') {
      this.fileB = null;
      if (this.fileInputB) this.fileInputB.value = ''; // Reset file input
    }

    // Update UI
    const fileInfo = document.getElementById(`fileInfo${type}`);
    const uploadBox = document.getElementById(`uploadBox${type}`);

    if (fileInfo && uploadBox) {
      // Hide file info section
      fileInfo.style.display = 'none';

      // Remove visual indicator
      uploadBox.classList.remove('file-selected');
    }

    // Disable "Start Comparing" button if either file is missing
    const startCompareBtn = document.getElementById('startCompareBtn');
    if (startCompareBtn && (!this.fileA || !this.fileB)) {
      startCompareBtn.disabled = true;
    }

    // Notify parent application via callback
    if (this.onFileUploadCallback) {
      this.onFileUploadCallback(null, type);
    }
  }

  /**
   * reset()
   * Resets both file slots to initial state
   * Removes both File A and File B
   */
  reset() {
    this.removeFile('A');
    this.removeFile('B');
  }

  /**
   * getFile(type)
   * Retrieves the currently stored file for a given type
   *
   * @param {string} type - Either 'A' or 'B' indicating which file slot
   * @returns {File|null} The stored file object or null if no file
   */
  getFile(type) {
    return type === 'A' ? this.fileA : this.fileB;
  }
}

// Export for use in other modules
export default FileHandler;
