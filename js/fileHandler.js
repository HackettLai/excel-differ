// fileHandler.js

class FileHandler {
  constructor() {
    this.fileInputA = null;
    this.fileInputB = null;
    this.uploadBoxA = null;
    this.uploadBoxB = null;
    this.fileA = null;
    this.fileB = null;
    this.onFileUploadCallback = null;
  }

  init(onFileUpload) {
    this.onFileUploadCallback = onFileUpload;

    this.fileInputA = document.getElementById('fileInputA');
    this.fileInputB = document.getElementById('fileInputB');
    this.uploadBoxA = document.getElementById('uploadBoxA');
    this.uploadBoxB = document.getElementById('uploadBoxB');

    this.setupEventListeners();
  }

  setupEventListeners() {
    // File input A
    if (this.uploadBoxA && this.fileInputA) {
      this.uploadBoxA.addEventListener('click', () => {
        this.fileInputA.click();
      });

      this.fileInputA.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
          this.handleFileSelect(file, 'A');
        }
      });

      // ✅ 加入 Drag & Drop for Box A
      this.setupDragAndDrop(this.uploadBoxA, 'A');
    }

    // File input B
    if (this.uploadBoxB && this.fileInputB) {
      this.uploadBoxB.addEventListener('click', () => {
        this.fileInputB.click();
      });

      this.fileInputB.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
          this.handleFileSelect(file, 'B');
        }
      });

      // ✅ 加入 Drag & Drop for Box B
      this.setupDragAndDrop(this.uploadBoxB, 'B');
    }

    // Remove file A
    const removeFileA = document.getElementById('removeFileA');
    if (removeFileA) {
      removeFileA.addEventListener('click', (e) => {
        e.stopPropagation();
        this.removeFile('A');
      });
    }

    // Remove file B
    const removeFileB = document.getElementById('removeFileB');
    if (removeFileB) {
      removeFileB.addEventListener('click', (e) => {
        e.stopPropagation();
        this.removeFile('B');
      });
    }
  }

  /**
   * ✅ 設定 Drag & Drop 功能
   */
  setupDragAndDrop(uploadBox, type) {
    if (!uploadBox) return;

    // 防止預設行為
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach((eventName) => {
      uploadBox.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
      });
    });

    // Drag enter/over - 高亮顯示
    ['dragenter', 'dragover'].forEach((eventName) => {
      uploadBox.addEventListener(eventName, () => {
        uploadBox.classList.add('active');
      });
    });

    // Drag leave - 移除高亮
    ['dragleave', 'drop'].forEach((eventName) => {
      uploadBox.addEventListener(eventName, () => {
        uploadBox.classList.remove('active');
      });
    });

    // Drop - 處理檔案
    uploadBox.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;

      if (files.length > 0) {
        const file = files[0];
        this.handleFileSelect(file, type);
      }
    });
  }

  handleFileSelect(file, type) {
    console.log('File selected:', file.name, 'Type:', type, 'File object:', file); // ✅ Debug log

    if (!file) {
      console.error('No file provided');
      return;
    }

    // Validate file type
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      alert('Please select a valid Excel file (.xlsx or .xls)');
      return;
    }

    // Store file
    if (type === 'A') {
      this.fileA = file;
      this.updateFileInfo('A', file.name);
    } else if (type === 'B') {
      this.fileB = file;
      this.updateFileInfo('B', file.name);
    }

    // Callback to main app
    if (this.onFileUploadCallback) {
      this.onFileUploadCallback(file, type); // ✅ 傳遞真正的 File 物件
    }
  }

  updateFileInfo(type, fileName) {
    const fileInfo = document.getElementById(`fileInfo${type}`);
    const fileNameEl = document.getElementById(`fileName${type}`);
    const uploadBox = document.getElementById(`uploadBox${type}`);

    if (fileInfo && fileNameEl && uploadBox) {
      fileNameEl.textContent = fileName;
      fileInfo.style.display = 'flex';
      uploadBox.classList.add('file-selected');
    }
  }

  removeFile(type) {
    if (type === 'A') {
      this.fileA = null;
      if (this.fileInputA) this.fileInputA.value = '';
    } else if (type === 'B') {
      this.fileB = null;
      if (this.fileInputB) this.fileInputB.value = '';
    }

    const fileInfo = document.getElementById(`fileInfo${type}`);
    const uploadBox = document.getElementById(`uploadBox${type}`);

    if (fileInfo && uploadBox) {
      fileInfo.style.display = 'none';
      uploadBox.classList.remove('file-selected');
    }

    // Disable compare button if either file is missing
    const startCompareBtn = document.getElementById('startCompareBtn');
    if (startCompareBtn && (!this.fileA || !this.fileB)) {
      startCompareBtn.disabled = true;
    }

    // Notify main app
    if (this.onFileUploadCallback) {
      this.onFileUploadCallback(null, type);
    }
  }

  reset() {
    this.removeFile('A');
    this.removeFile('B');
  }

  // ✅ 新增：取得目前儲存的檔案
  getFile(type) {
    return type === 'A' ? this.fileA : this.fileB;
  }
}

export default FileHandler;
