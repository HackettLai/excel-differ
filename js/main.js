// main.js

import FileHandler from './fileHandler.js';
import ExcelParser from './excelParser.js';
import DiffEngine from './diffEngine.js';
import DiffViewer from './diffViewer.js';
import Copyright from './copyright.js';

class ExcelDiffer {
    constructor() {
        this.fileHandler = new FileHandler();
        this.excelParser = new ExcelParser();
        this.diffEngine = new DiffEngine();
        this.diffViewer = new DiffViewer();
        this.copyright = new Copyright();

        this.fileA = null;
        this.fileB = null;
        this.dataA = null;
        this.dataB = null;
        this.diffResults = null;
    }

    init() {
        this.fileHandler.init(this.handleFileUpload.bind(this));
        this.copyright.init();
        this.setupEventListeners();
    }

    setupEventListeners() {
        const compareBtn = document.getElementById('compareBtn');
        const backBtn = document.getElementById('backBtn');
        const compareSheetBtn = document.getElementById('compareSheetBtn');
        const prevBtn = document.getElementById('prevChangeBtn');
        const nextBtn = document.getElementById('nextChangeBtn');

        if (compareBtn) {
            compareBtn.addEventListener('click', () => this.compareFiles());
        }

        if (backBtn) {
            backBtn.addEventListener('click', () => this.reset());
        }

        if (compareSheetBtn) {
            compareSheetBtn.addEventListener('click', () => this.handleCompareSheets());
        }

        if (prevBtn) {
            prevBtn.addEventListener('click', () => {
                if (this.diffViewer) {
                    this.diffViewer.navigateToChange(-1);
                }
            });
        }

        if (nextBtn) {
            nextBtn.addEventListener('click', () => {
                if (this.diffViewer) {
                    this.diffViewer.navigateToChange(1);
                }
            });
        }
    }

    handleFileUpload(file, type) {
        if (type === 'A') {
            this.fileA = file;
        } else if (type === 'B') {
            this.fileB = file;
        }

        const compareBtn = document.getElementById('compareBtn');
        if (this.fileA && this.fileB && compareBtn) {
            compareBtn.disabled = false;
        }
    }

    async compareFiles() {
        const loadingOverlay = document.getElementById('loadingOverlay');

        try {
            if (loadingOverlay) loadingOverlay.style.display = 'flex';

            console.log('File A:', this.fileA);
            console.log('File B:', this.fileB);

            if (!this.fileA || !this.fileB) {
                throw new Error('Please select both files');
            }

            if (!(this.fileA instanceof File)) {
                throw new Error('File A is not a valid File object');
            }

            if (!(this.fileB instanceof File)) {
                throw new Error('File B is not a valid File object');
            }

            console.log('Parsing File A...');
            this.dataA = await this.excelParser.parse(this.fileA);
            console.log('File A parsed:', this.dataA);

            console.log('Parsing File B...');
            this.dataB = await this.excelParser.parse(this.fileB);
            console.log('File B parsed:', this.dataB);

            console.log('Comparing files...');
            this.diffResults = this.diffEngine.compare(this.dataA, this.dataB);
            console.log('Diff results:', this.diffResults);

            // 直接進入 Diff View
            this.diffViewer.init(this.dataA, this.dataB, this.diffResults);

            this.hideUploadSection();
            this.showDiffSection();
        } catch (error) {
            console.error('Error comparing files:', error);
            alert('Error comparing files: ' + error.message);
        } finally {
            if (loadingOverlay) loadingOverlay.style.display = 'none';
        }
    }

    handleCompareSheets() {
        if (this.diffViewer) {
            this.diffViewer.compareSelectedSheets();
        }
    }

    reset() {
        this.fileA = null;
        this.fileB = null;
        this.dataA = null;
        this.dataB = null;
        this.diffResults = null;

        this.fileHandler.reset();
        this.hideDiffSection();
        this.showUploadSection();

        const compareBtn = document.getElementById('compareBtn');
        if (compareBtn) compareBtn.disabled = true;
    }

    hideUploadSection() {
        const uploadSection = document.getElementById('uploadSection');
        if (uploadSection) uploadSection.style.display = 'none';
    }

    showUploadSection() {
        const uploadSection = document.getElementById('uploadSection');
        if (uploadSection) uploadSection.style.display = 'block';
    }

    hideDiffSection() {
        const diffSection = document.getElementById('diffSection');
        if (diffSection) diffSection.style.display = 'none';
    }

    showDiffSection() {
        const diffSection = document.getElementById('diffSection');
        if (diffSection) diffSection.style.display = 'block';
    }
}

document.addEventListener('DOMContentLoaded', () => {
    const app = new ExcelDiffer();
    app.init();

    window.excelDiffer = app;
});