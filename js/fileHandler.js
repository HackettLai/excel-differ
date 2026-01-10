// fileHandler.js - 處理文件上傳和讀取

const FileHandler = {
    fileA: null,
    fileB: null,
    
    // 初始化文件上傳監聽器
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
        
        // 拖拽支持
        this.setupDragAndDrop('uploadA', fileInputA);
        this.setupDragAndDrop('uploadB', fileInputB);
    },
    
    // 處理文件選擇
    handleFileSelect(event, fileType) {
        const file = event.target.files[0];
        
        if (!file) return;
        
        // 驗證文件類型
        if (!this.isValidFile(file)) {
            alert('Please select a valid Excel file (.xlsx, .xls, .csv)');
            return;
        }
        
        // 驗證文件大小 (限制 50MB)
        if (file.size > 50 * 1024 * 1024) {
            alert('File too large! Please select a file smaller than 50MB');
            return;
        }
        
        // 保存文件引用
        if (fileType === 'A') {
            this.fileA = file;
            document.getElementById('fileNameA').textContent = file.name;
        } else {
            this.fileB = file;
            document.getElementById('fileNameB').textContent = file.name;
        }
        
        // 檢查是否可以啟用比較按鈕
        this.checkCompareButton();
    },
    
    // 驗證文件類型
    isValidFile(file) {
        const validExtensions = ['.xlsx', '.xls', '.csv'];
        const fileName = file.name.toLowerCase();
        return validExtensions.some(ext => fileName.endsWith(ext));
    },
    
    // 檢查並啟用/禁用比較按鈕
    checkCompareButton() {
        const compareBtn = document.getElementById('compareBtn');
        
        if (this.fileA && this.fileB) {
            compareBtn.disabled = false;
        } else {
            compareBtn.disabled = true;
        }
    },
    
    // 設置拖放功能
    setupDragAndDrop(uploadAreaId, fileInput) {
        const uploadArea = document.getElementById(uploadAreaId);
        
        // 防止默認拖放行為
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
            });
        });
        
        // 高亮拖放區域
        ['dragenter', 'dragover'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => {
                uploadArea.classList.add('drag-over');
            });
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => {
                uploadArea.classList.remove('drag-over');
            });
        });
        
        // 處理文件放下
        uploadArea.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                // 創建一個新的 FileList 並觸發 change 事件
                fileInput.files = files;
                fileInput.dispatchEvent(new Event('change'));
            }
        });
    },
    
    // 讀取文件為 ArrayBuffer
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
    
    // 獲取文件數據
    async getFileData(fileType) {
        const file = fileType === 'A' ? this.fileA : this.fileB;
        
        if (!file) {
            throw new Error(`File ${fileType} not selected`);
        }
        
        return {
            file: file,
            arrayBuffer: await this.readFile(file)
        };
    },
    
    // 重置文件選擇
    reset() {
        this.fileA = null;
        this.fileB = null;
        
        document.getElementById('fileA').value = '';
        document.getElementById('fileB').value = '';
        document.getElementById('fileNameA').textContent = 'No file selected';
        document.getElementById('fileNameB').textContent = 'No file selected';
        
        this.checkCompareButton();
    }
};