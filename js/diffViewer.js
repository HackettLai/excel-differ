// diffViewer.js

import DiffEngine from './diffEngine.js';

class DiffViewer {
  constructor() {
    this.dataA = null;
    this.dataB = null;
    this.diffResults = null;
    this.selectedSheetA = null;
    this.selectedSheetB = null;
    this.changedCells = [];
    this.currentChangeIndex = -1;
    this._allColumns = [];
    this.diffEngine = new DiffEngine(); // ✨ 使用 DiffEngine 實例
  }

  /**
   * Initialize diff viewer with data
   * @param {Object} dataA - File A workbook data
   * @param {Object} dataB - File B workbook data
   * @param {Object} diffResults - Diff results from DiffEngine
   */
  init(dataA, dataB, diffResults) {
    console.log('Initializing DiffViewer');

    this.dataA = dataA;
    this.dataB = dataB;
    this.diffResults = diffResults;

    // Populate sheet dropdowns
    this.populateSheetDropdowns();

    // Auto-select sheets (same name if exists, otherwise first sheet)
    this.autoSelectSheets();

    console.log('DiffViewer initialized');
  }

  /**
   * Populate sheet selection dropdowns
   */
  populateSheetDropdowns() {
    const selectA = document.getElementById('sheetSelectA');
    const selectB = document.getElementById('sheetSelectB');

    if (!selectA || !selectB) {
      console.error('Sheet select elements not found');
      return;
    }

    // Clear existing options
    selectA.innerHTML = '';
    selectB.innerHTML = '';

    // Populate File A sheets
    if (this.dataA && this.dataA.sheetNames) {
      this.dataA.sheetNames.forEach((sheetName) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        selectA.appendChild(option);
      });
    }

    // Populate File B sheets
    if (this.dataB && this.dataB.sheetNames) {
      this.dataB.sheetNames.forEach((sheetName) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        selectB.appendChild(option);
      });
    }

    console.log('Sheet dropdowns populated');
  }

  /**
   * Auto-select sheets based on matching names
   */
  autoSelectSheets() {
    const selectA = document.getElementById('sheetSelectA');
    const selectB = document.getElementById('sheetSelectB');

    if (!selectA || !selectB) return;

    // Get first sheet from each file
    const firstSheetA = this.dataA?.sheetNames?.[0];
    const firstSheetB = this.dataB?.sheetNames?.[0];

    if (!firstSheetA || !firstSheetB) {
      console.warn('No sheets found in files');
      return;
    }

    // Check if there's a matching sheet name
    let matchFound = false;

    for (const sheetA of this.dataA.sheetNames) {
      if (this.dataB.sheetNames.includes(sheetA)) {
        // Found matching sheet name
        selectA.value = sheetA;
        selectB.value = sheetA;
        this.selectedSheetA = sheetA;
        this.selectedSheetB = sheetA;
        matchFound = true;
        console.log('Auto-selected matching sheets:', sheetA);

        // Auto-compare
        this.compareSelectedSheets();
        break;
      }
    }

    if (!matchFound) {
      // No match, just select first sheet of each
      selectA.value = firstSheetA;
      selectB.value = firstSheetB;
      this.selectedSheetA = firstSheetA;
      this.selectedSheetB = firstSheetB;
      console.log('No matching sheets, selected first sheets');

      // Don't auto-compare if names don't match
    }
  }

  /**
   * Compare selected sheets (triggered by Compare button)
   */
  compareSelectedSheets() {
    const selectA = document.getElementById('sheetSelectA');
    const selectB = document.getElementById('sheetSelectB');

    if (!selectA || !selectB) {
      console.error('Sheet select elements not found');
      return;
    }

    this.selectedSheetA = selectA.value;
    this.selectedSheetB = selectB.value;

    console.log('Comparing sheets:', this.selectedSheetA, 'vs', this.selectedSheetB);

    // Get sheet data
    const sheetDataA = this.dataA.sheets[this.selectedSheetA];
    const sheetDataB = this.dataB.sheets[this.selectedSheetB];

    if (!sheetDataA || !sheetDataB) {
      console.error('Sheet data not found');
      this.showNoDataMessage();
      return;
    }

    // ✨ Use DiffEngine's compareSheets method
    const singleSheetDiff = this.diffEngine.compareSheets(sheetDataA, sheetDataB);

    console.log('Single sheet diff result:', singleSheetDiff);

    // Render the diff table
    this.renderUnifiedTable(singleSheetDiff);
  }

  /**
   * Render unified diff table
   */
  renderUnifiedTable(sheetDiff) {
    const container = document.getElementById('unifiedTableContainer');
    if (!container) {
      console.error('unifiedTableContainer not found');
      return;
    }

    if (!sheetDiff || !sheetDiff.rowDiff) {
      this.showNoDataMessage();
      this.hideChangeNavigation();
      return;
    }

    // Cache all columns
    this._allColumns = this.getAllColumns(sheetDiff.rowDiff);

    // Create table
    const table = document.createElement('table');
    table.className = 'unified-table diff-table';

    // Create header
    const thead = this.createUnifiedHeader();
    table.appendChild(thead);

    // Create body
    const tbody = this.createUnifiedBody(sheetDiff.rowDiff);
    table.appendChild(tbody);

    container.innerHTML = '';
    container.appendChild(table);

    // Collect changed cells
    this.collectChangedCells();
  }

  /**
   * Show "no data" message
   */
  showNoDataMessage() {
    const container = document.getElementById('unifiedTableContainer');
    if (container) {
      container.innerHTML = '<p style="padding: 20px; text-align: center; color: #718096;">No data available for comparison</p>';
    }
  }

  /**
   * Create unified table header
   */
  createUnifiedHeader() {
    const thead = document.createElement('thead');
    const tr = document.createElement('tr');

    // Old Index column
    const thOld = document.createElement('th');
    thOld.className = 'index-col';
    thOld.textContent = 'Old';
    tr.appendChild(thOld);

    // New Index column
    const thNew = document.createElement('th');
    thNew.className = 'index-col';
    thNew.textContent = 'New';
    tr.appendChild(thNew);

    // Column headers
    this._allColumns.forEach((col) => {
      const th = document.createElement('th');
      th.textContent = col;
      tr.appendChild(th);
    });

    thead.appendChild(tr);
    return thead;
  }

  /**
   * Create unified table body
   */
  createUnifiedBody(rowDiff) {
    const tbody = document.createElement('tbody');

    rowDiff.forEach((rowData, idx) => {
      const tr = this.createUnifiedRow(rowData, idx);
      tbody.appendChild(tr);
    });

    return tbody;
  }

  /**
   * Create a single unified row
   */
  createUnifiedRow(rowData, rowIndex) {
    const tr = document.createElement('tr');
    tr.className = `row-${rowData.type}`;
    tr.dataset.rowIndex = rowIndex;

    // Old Index cell
    const tdOld = document.createElement('td');
    tdOld.className = 'index-cell old-idx';
    tdOld.textContent = rowData.oldIndex !== null ? rowData.oldIndex : '-';
    tr.appendChild(tdOld);

    // New Index cell
    const tdNew = document.createElement('td');
    tdNew.className = 'index-cell new-idx';
    tdNew.textContent = rowData.newIndex !== null ? rowData.newIndex : '-';
    tr.appendChild(tdNew);

    // Data cells
    this._allColumns.forEach((colName) => {
      const cellData = rowData.cells[colName];
      const td = this.createUnifiedCell(cellData, rowData.type, rowIndex, colName);
      tr.appendChild(td);
    });

    return tr;
  }

  /**
   * Create a single unified cell
   */
  createUnifiedCell(cellData, rowType, rowIndex, colName) {
    const td = document.createElement('td');
    td.dataset.row = rowIndex;
    td.dataset.col = colName;

    if (!cellData) {
      td.className = `cell-${rowType}`;
      td.textContent = '';
      return td;
    }

    if (cellData.changed) {
      td.className = 'cell-modified';

      const oldSpan = document.createElement('span');
      oldSpan.className = 'old-value';
      oldSpan.textContent = this.formatCellValue(cellData.oldValue);

      const separator = document.createElement('span');
      separator.className = 'value-separator';
      separator.textContent = '→';

      const newSpan = document.createElement('span');
      newSpan.className = 'new-value';
      newSpan.textContent = this.formatCellValue(cellData.newValue);

      td.appendChild(oldSpan);
      td.appendChild(separator);
      td.appendChild(newSpan);

      td.dataset.hasChange = 'true';
    } else {
      td.className = `cell-${rowType}`;
      td.textContent = this.formatCellValue(cellData.value);
    }

    return td;
  }

  /**
   * Get all column names from row diff data
   */
  getAllColumns(rowDiff) {
    const colSet = new Set();

    rowDiff.forEach((row) => {
      if (row.cells) {
        Object.keys(row.cells).forEach((col) => {
          colSet.add(col);
        });
      }
    });

    // ✅ 按照 A, B, C... 的順序排列
    return Array.from(colSet).sort((a, b) => {
      return this.columnToIndex(a) - this.columnToIndex(b);
    });
  }

  /**
   * ✅ 新增：將欄位名稱轉換為數字索引
   */
  columnToIndex(col) {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.charCodeAt(i) - 64);
    }
    return index;
  }

  /**
   * Format cell value for display
   */
  formatCellValue(value) {
    if (value === null || value === undefined || value === '') {
      return '';
    }
    return String(value);
  }

  /**
   * Collect all changed cells for navigation
   */
  collectChangedCells() {
    this.changedCells = [];
    const cells = document.querySelectorAll('[data-has-change="true"]');

    cells.forEach((cell) => {
      this.changedCells.push({
        element: cell,
        row: cell.dataset.row,
        col: cell.dataset.col,
      });
    });

    this.currentChangeIndex = -1;

    if (this.changedCells.length > 0) {
      this.showChangeNavigation();
      this.updateChangeCounter();
    } else {
      this.hideChangeNavigation();
    }

    console.log(`Found ${this.changedCells.length} changed cells`);
  }

  /**
   * Navigate to next/previous change
   */
  navigateToChange(direction) {
    if (this.changedCells.length === 0) return;

    if (this.currentChangeIndex >= 0) {
      const current = this.changedCells[this.currentChangeIndex];
      if (current && current.element) {
        current.element.style.outline = '';
      }
    }

    this.currentChangeIndex += direction;

    if (this.currentChangeIndex >= this.changedCells.length) {
      this.currentChangeIndex = 0;
    } else if (this.currentChangeIndex < 0) {
      this.currentChangeIndex = this.changedCells.length - 1;
    }

    const target = this.changedCells[this.currentChangeIndex];
    if (target && target.element) {
      target.element.style.outline = '3px solid #667eea';
      target.element.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    this.updateChangeCounter();
  }

  /**
   * Update change counter display
   */
  updateChangeCounter() {
    const counter = document.getElementById('changeCounter');
    if (counter) {
      const current = this.currentChangeIndex + 1;
      const total = this.changedCells.length;
      counter.textContent = `${current} / ${total}`;
    }

    const prevBtn = document.getElementById('prevChangeBtn');
    const nextBtn = document.getElementById('nextChangeBtn');

    if (prevBtn) prevBtn.disabled = this.changedCells.length === 0;
    if (nextBtn) nextBtn.disabled = this.changedCells.length === 0;
  }

  /**
   * Show change navigation
   */
  showChangeNavigation() {
    const nav = document.getElementById('changeNavigation');
    if (nav) nav.style.display = 'flex';
  }

  /**
   * Hide change navigation
   */
  hideChangeNavigation() {
    const nav = document.getElementById('changeNavigation');
    if (nav) nav.style.display = 'none';
  }
}

export default DiffViewer;
