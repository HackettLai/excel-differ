/**
 * diffViewer.js
 * Excel Diff Viewer with Interleaved Row Sorting
 *
 * Responsibilities:
 * - Render unified diff table with side-by-side comparison
 * - Handle user interactions (sheet selection, key column selection, navigation)
 * - Support both key-based matching and position-based matching
 * - Display row changes (added/deleted/matched) with proper visual styling
 * - Implement interleaved sorting: deleted rows inserted before next new row position
 */

import DiffEngine from './diffEngine.js';

/**
 * DiffViewer Class
 * Main UI controller for rendering Excel comparison results
 * Handles table generation, change navigation, and user interactions
 */
class DiffViewer {
  constructor() {
    this.dataA = null; // Parsed data from first Excel file
    this.dataB = null; // Parsed data from second Excel file
    this.diffResults = null; // Comparison results from DiffEngine
    this.changedCells = []; // Array of all changed cells for navigation
    this.currentChangeIndex = -1; // Current position in change navigation
  }

  /**
   * init(dataA, dataB, diffResults)
   * Initialize the diff viewer with parsed Excel data and comparison results
   * Sets up dropdowns, auto-selects sheets, and prepares UI for comparison
   *
   * @param {Object} dataA - Parsed data from first Excel file
   * @param {Object} dataB - Parsed data from second Excel file
   * @param {Object} diffResults - Comparison results from DiffEngine
   */
  init(dataA, dataB, diffResults) {
    this.dataA = dataA;
    this.dataB = dataB;
    this.diffResults = diffResults;

    console.log('🔧 Initializing DiffViewer...');

    // Populate sheet selection dropdowns
    this.populateSheetDropdowns();

    // Auto-select first sheet in both files
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');

    if (sheetSelectA && this.dataA.sheetNames?.[0]) {
      sheetSelectA.value = this.dataA.sheetNames[0];
      console.log(`✅ Auto-selected Sheet A: ${this.dataA.sheetNames[0]}`);
    }

    if (sheetSelectB && this.dataB.sheetNames?.[0]) {
      sheetSelectB.value = this.dataB.sheetNames[0];
      console.log(`✅ Auto-selected Sheet B: ${this.dataB.sheetNames[0]}`);
    }

    // Populate header row and key column dropdowns
    this.populateHeaderRowDropdowns();
    this.populateKeyColumnDropdown();

    // Setup event listeners
    this.setupSheetChangeListeners();
    this.setupChangeNavigation();

    console.log('✅ DiffViewer initialized');
  }

  /**
   * populateHeaderRowDropdowns()
   * Populate header row selection dropdowns for both files
   * Allows user to specify which row contains column headers (default: row 1)
   */
  populateHeaderRowDropdowns() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');

    if (!headerRowA || !headerRowB) {
      console.error('❌ Header row dropdown elements not found');
      return;
    }

    const selectedSheetA = sheetSelectA?.value || this.dataA?.sheetNames?.[0];
    const selectedSheetB = sheetSelectB?.value || this.dataB?.sheetNames?.[0];

    const rowCountA = this.dataA?.sheets[selectedSheetA]?.rowCount || 50;
    const rowCountB = this.dataB?.sheets[selectedSheetB]?.rowCount || 50;

    // Populate dropdown A (max 50 rows)
    headerRowA.innerHTML = '';
    const maxRowsA = Math.min(rowCountA, 50);
    for (let i = 1; i <= maxRowsA; i++) {
      const option = document.createElement('option');
      option.value = i;
      option.textContent = `Row ${i}`;
      headerRowA.appendChild(option);
    }

    // Populate dropdown B (max 50 rows)
    headerRowB.innerHTML = '';
    const maxRowsB = Math.min(rowCountB, 50);
    for (let i = 1; i <= maxRowsB; i++) {
      const option = document.createElement('option');
      option.value = i;
      option.textContent = `Row ${i}`;
      headerRowB.appendChild(option);
    }

    // Default to row 1
    headerRowA.value = '1';
    headerRowB.value = '1';
  }

  /**
   * populateKeyColumnDropdown()
   * Populate key column selection dropdown with common columns from both sheets
   * Auto-selects columns named "email", "code", or "id" if available
   * Includes "(Use Row Position)" option for position-based matching
   */
  populateKeyColumnDropdown() {
    const keyColumnSelect = document.getElementById('keyColumnSelect');
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');

    if (!keyColumnSelect) {
      console.error('❌ Key column dropdown not found');
      return;
    }

    const selectedSheetA = sheetSelectA?.value || this.dataA?.sheetNames?.[0];
    const selectedSheetB = sheetSelectB?.value || this.dataB?.sheetNames?.[0];
    const selectedHeaderA = parseInt(headerRowA?.value) || 1;
    const selectedHeaderB = parseInt(headerRowB?.value) || 1;

    const sheetA = this.dataA?.sheets[selectedSheetA]?.data || [];
    const sheetB = this.dataB?.sheets[selectedSheetB]?.data || [];

    if (sheetA.length === 0 || sheetB.length === 0) {
      console.warn('⚠️ One or both sheets are empty');
      keyColumnSelect.innerHTML = '<option value="">-- No data available --</option>';
      return;
    }

    const headerIndexA = selectedHeaderA - 1;
    const headerIndexB = selectedHeaderB - 1;

    if (headerIndexA >= sheetA.length || headerIndexB >= sheetB.length) {
      console.warn('⚠️ Invalid header row index');
      keyColumnSelect.innerHTML = '<option value="">-- Invalid header row --</option>';
      return;
    }

    // Extract headers from both sheets
    const headersA = sheetA[headerIndexA] ? Object.values(sheetA[headerIndexA]).filter((h) => h && String(h).trim()) : [];
    const headersB = sheetB[headerIndexB] ? Object.values(sheetB[headerIndexB]).filter((h) => h && String(h).trim()) : [];

    // Find common columns between both sheets
    const commonColumns = this.findCommonColumns(headersA, headersB);

    // Build dropdown options
    keyColumnSelect.innerHTML = '';

    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = '── Select Key Column ──';
    keyColumnSelect.appendChild(defaultOption);

    const rowPositionOption = document.createElement('option');
    rowPositionOption.value = '__ROW_POSITION__';
    rowPositionOption.textContent = '(Use Row Position)';
    keyColumnSelect.appendChild(rowPositionOption);

    const separator = document.createElement('option');
    separator.disabled = true;
    separator.textContent = '────────────────────────';
    keyColumnSelect.appendChild(separator);

    // If no common columns, default to row position matching
    if (commonColumns.length === 0) {
      keyColumnSelect.value = '__ROW_POSITION__';
    } else {
      let autoSelected = false;

      // Add common columns to dropdown
      commonColumns.forEach((col) => {
        if (!col.name) return;

        const option = document.createElement('option');
        option.value = col.colIndex;
        option.textContent = col.name;

        // Auto-select columns named "email", "code", or "id"
        const lowerName = col.name.toLowerCase();
        if (!autoSelected && (lowerName.includes('email') || lowerName.includes('code') || lowerName.includes('id'))) {
          option.selected = true;
          autoSelected = true;
        }

        keyColumnSelect.appendChild(option);
      });

      // If no auto-selection, default to first common column
      if (!autoSelected && commonColumns.length > 0) {
        keyColumnSelect.value = commonColumns[0].colIndex;
      }
    }
  }

  /**
   * findCommonColumns(headersA, headersB)
   * Find columns that exist in both sheets with matching names (case-insensitive)
   * Returns array of common columns with their names and column letters
   *
   * @param {Array<string>} headersA - Headers from first sheet
   * @param {Array<string>} headersB - Headers from second sheet
   * @returns {Array<Object>} Array of common columns with { name, colIndex, indexA, indexB }
   */
  findCommonColumns(headersA, headersB) {
    const commonColumns = [];

    headersA.forEach((headerA, indexA) => {
      if (!headerA) return;
      const trimmedA = String(headerA).trim();
      if (!trimmedA) return;

      const normalizedA = trimmedA.toLowerCase();

      headersB.forEach((headerB, indexB) => {
        if (!headerB) return;
        const trimmedB = String(headerB).trim();
        if (!trimmedB) return;

        const normalizedB = trimmedB.toLowerCase();

        // Case-insensitive match
        if (normalizedA === normalizedB) {
          const alreadyExists = commonColumns.some((col) => col.name === trimmedA);

          if (!alreadyExists) {
            commonColumns.push({
              name: trimmedA,
              colIndex: this.getColumnLetter(indexA),
              indexA: indexA,
              indexB: indexB,
            });
          }
        }
      });
    });

    return commonColumns;
  }

  /**
   * getColumnLetter(index)
   * Convert column index (0-based) to Excel column letter (A, B, C, ..., Z, AA, AB, ...)
   *
   * @param {number} index - Zero-based column index
   * @returns {string} Excel column letter
   */
  getColumnLetter(index) {
    let letter = '';
    while (index >= 0) {
      letter = String.fromCharCode((index % 26) + 65) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }

  /**
   * setupSheetChangeListeners()
   * Setup event listeners for sheet and header row selection changes
   * Updates key column dropdown when selections change
   */
  setupSheetChangeListeners() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');

    if (sheetSelectA) {
      sheetSelectA.addEventListener('change', () => {
        this.updateHeaderRowDropdown('A');
        this.populateKeyColumnDropdown();
      });
    }

    if (sheetSelectB) {
      sheetSelectB.addEventListener('change', () => {
        this.updateHeaderRowDropdown('B');
        this.populateKeyColumnDropdown();
      });
    }

    if (headerRowA) {
      headerRowA.addEventListener('change', () => {
        this.populateKeyColumnDropdown();
      });
    }

    if (headerRowB) {
      headerRowB.addEventListener('change', () => {
        this.populateKeyColumnDropdown();
      });
    }
  }

  /**
   * updateHeaderRowDropdown(type)
   * Update header row dropdown when sheet selection changes
   * Rebuilds dropdown options based on selected sheet's row count
   *
   * @param {string} type - 'A' or 'B' to specify which file
   */
  updateHeaderRowDropdown(type) {
    const sheetSelect = document.getElementById(`sheetSelect${type}`);
    const headerRowSelect = document.getElementById(`headerRow${type}`);
    const data = type === 'A' ? this.dataA : this.dataB;

    if (!sheetSelect || !headerRowSelect || !data) return;

    const selectedSheet = sheetSelect.value;
    const rowCount = data.sheets[selectedSheet]?.rowCount || 20;

    headerRowSelect.innerHTML = '';
    const maxRows = Math.min(rowCount, 50);
    for (let i = 1; i <= maxRows; i++) {
      const option = document.createElement('option');
      option.value = i;
      option.textContent = `Row ${i}`;
      headerRowSelect.appendChild(option);
    }

    headerRowSelect.value = '1';
  }

  /**
   * findMatchingSheet()
   * Find sheets with matching names between both files
   * Used for auto-selection if sheet names match
   *
   * @returns {Object|null} { sheetA, sheetB } or null if no match found
   */
  findMatchingSheet() {
    if (!this.dataA.sheetNames || !this.dataB.sheetNames) return null;

    for (let sheetA of this.dataA.sheetNames) {
      if (this.dataB.sheetNames.includes(sheetA)) {
        return { sheetA, sheetB: sheetA };
      }
    }

    return null;
  }

  /**
   * populateSheetDropdowns()
   * Populate sheet selection dropdowns for both files
   * Shows all available sheets from both Excel files
   */
  populateSheetDropdowns() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');

    if (!sheetSelectA || !sheetSelectB) {
      console.error('Sheet dropdown elements not found');
      return;
    }

    sheetSelectA.innerHTML = '';
    sheetSelectB.innerHTML = '';

    if (this.dataA && this.dataA.sheetNames) {
      this.dataA.sheetNames.forEach((sheetName) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelectA.appendChild(option);
      });
    }

    if (this.dataB && this.dataB.sheetNames) {
      this.dataB.sheetNames.forEach((sheetName) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelectB.appendChild(option);
      });
    }
  }

  /**
   * compareSelectedSheets()
   * Compare currently selected sheets with specified header rows and key column
   * Triggers either key-based matching or position-based matching
   * Called when user clicks "Compare" button
   */
  compareSelectedSheets() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');
    const keyColumnSelect = document.getElementById('keyColumnSelect');

    if (!sheetSelectA || !sheetSelectB || !headerRowA || !headerRowB || !keyColumnSelect) {
      console.error('Selection elements not found');
      return;
    }

    const selectedSheetA = sheetSelectA.value;
    const selectedSheetB = sheetSelectB.value;
    const selectedHeaderA = parseInt(headerRowA.value) || 1;
    const selectedHeaderB = parseInt(headerRowB.value) || 1;
    const selectedKeyColumn = keyColumnSelect.value;

    if (!selectedSheetA || !selectedSheetB) {
      console.error('⚠️ Please select sheets to compare');
      return;
    }

    if (!selectedKeyColumn) {
      console.error('⚠️ Please select a Key Column or "Use Row Position"');
      return;
    }

    let sheetA = this.dataA.sheets[selectedSheetA]?.data || [];
    let sheetB = this.dataB.sheets[selectedSheetB]?.data || [];

    // Adjust data based on header row selection
    sheetA = this.adjustDataForHeaderRow(sheetA, selectedHeaderA);
    sheetB = this.adjustDataForHeaderRow(sheetB, selectedHeaderB);

    if (sheetA.length === 0 || sheetB.length === 0) {
      console.error('⚠️ Selected sheet is empty');
      return;
    }

    if (selectedKeyColumn === '__ROW_POSITION__') {
      // Position-based matching
      console.log('🔢 Using row position matching');
      this.compareByRowPosition(sheetA, sheetB, selectedHeaderA, selectedHeaderB);
    } else {
      // Key-based matching
      console.log(`🔑 Using key column: ${selectedKeyColumn}`);

      const diffEngine = new DiffEngine();
      const singleSheetDiff = diffEngine.compareSheets(sheetA, sheetB, selectedKeyColumn, selectedHeaderA, selectedHeaderB);

      singleSheetDiff.sheetName = `${selectedSheetA} vs ${selectedSheetB}`;
      singleSheetDiff.headerRowA = selectedHeaderA;
      singleSheetDiff.headerRowB = selectedHeaderB;

      this.renderUnifiedTable(singleSheetDiff, selectedKeyColumn);
    }
  }

  /**
   * adjustDataForHeaderRow(data, headerRow)
   * Slice data array to start from the specified header row
   * Ensures header row is treated as row 0 in the returned array
   *
   * @param {Array<Object>} data - Sheet data
   * @param {number} headerRow - Header row number (1-based)
   * @returns {Array<Object>} Adjusted data starting from header row
   */
  adjustDataForHeaderRow(data, headerRow) {
    if (!data || data.length === 0) return data;

    const headerIndex = headerRow - 1;

    if (headerIndex < 0 || headerIndex >= data.length) {
      console.warn(`Invalid header row ${headerRow}, using row 1`);
      return data;
    }

    return data.slice(headerIndex);
  }

  /**
   * renderUnifiedTable(sheetDiff, keyColumn)
   * Render unified diff table showing side-by-side comparison
   * Displays added/deleted/modified rows with visual highlighting
   * Uses interleaved sorting: deleted rows inserted before next new row position
   *
   * @param {Object} sheetDiff - Comparison results from DiffEngine
   * @param {string} keyColumn - Key column used for matching
   */
  renderUnifiedTable(sheetDiff, keyColumn = 'A') {
    const container = document.getElementById('unifiedTableContainer');
    if (!container) {
      console.error('unifiedTableContainer element not found');
      return;
    }

    container.innerHTML = '';

    const table = document.createElement('table');
    table.className = 'unified-table diff-table';

    const thead = this.buildUnifiedHeader(sheetDiff);
    table.appendChild(thead);

    const tbody = this.buildUnifiedBody(sheetDiff, keyColumn);
    table.appendChild(tbody);

    const wrapper = document.createElement('div');
    wrapper.className = 'table-wrapper';
    wrapper.appendChild(table);

    container.appendChild(wrapper);

    this.collectChangedCells();
    this.setupCellClickNavigation();
  }

  /**
   * setupCellClickNavigation()
   * Setup click listeners on changed cells for navigation
   * Allows user to click on changed cells to navigate to them
   */
  setupCellClickNavigation() {
    const table = document.querySelector('#unifiedTableContainer .diff-table');
    if (!table) return;

    const tbody = table.querySelector('tbody');
    if (!tbody) return;

    tbody.addEventListener('click', (e) => {
      const clickedCell = e.target.closest('td');
      if (!clickedCell) return;

      const isChangedCell = clickedCell.classList.contains('cell-modified') || clickedCell.classList.contains('cell-added') || clickedCell.classList.contains('cell-deleted');

      if (!isChangedCell) return;

      const clickedRow = clickedCell.closest('tr');

      for (let i = 0; i < this.changedCells.length; i++) {
        const { row, cell } = this.changedCells[i];

        if (row === clickedRow && cell === clickedCell) {
          this.currentChangeIndex = i;
          this.updateNavigationUI();
          this.scrollToChange();
          return;
        }
      }
    });
  }

  /**
   * getUnifiedColumns(sheetDiff)
   * Build unified column list merging columns from both sheets
   * Handles added/deleted columns and blank columns
   * Returns columns in new file's order, with deleted columns appended at end
   *
   * @param {Object} sheetDiff - Comparison results
   * @returns {Array<Object>} Array of unified columns with { header, oldCol, newCol, type }
   */
  getUnifiedColumns(sheetDiff) {
    const oldHeaders = sheetDiff.oldData[0] || {};
    const newHeaders = sheetDiff.newData[0] || {};

    const headerMap = new Map();

    const normalizeHeader = (header) => {
      return String(header || '')
        .trim()
        .toLowerCase();
    };

    // Process old file headers
    Object.keys(oldHeaders).forEach((col) => {
      const content = String(oldHeaders[col] || '').trim();

      if (content) {
        const normalized = normalizeHeader(content);

        headerMap.set(normalized, {
          originalA: content,
          originalB: null,
          oldCol: col,
          newCol: null,
        });
      } else {
        headerMap.set(`__empty_old_${col}`, {
          originalA: '',
          originalB: null,
          oldCol: col,
          newCol: null,
        });
      }
    });

    // Process new file headers
    Object.keys(newHeaders).forEach((col) => {
      const content = String(newHeaders[col] || '').trim();

      if (content) {
        const normalized = normalizeHeader(content);

        if (headerMap.has(normalized)) {
          const existing = headerMap.get(normalized);
          existing.originalB = content;
          existing.newCol = col;
        } else {
          headerMap.set(normalized, {
            originalA: null,
            originalB: content,
            oldCol: null,
            newCol: col,
          });
        }
      } else {
        const key = `__empty_new_${col}`;

        const oldEmptyKey = `__empty_old_${col}`;
        if (headerMap.has(oldEmptyKey)) {
          headerMap.get(oldEmptyKey).newCol = col;
          headerMap.set(`__empty_both_${col}`, headerMap.get(oldEmptyKey));
          headerMap.delete(oldEmptyKey);
        } else {
          headerMap.set(key, {
            originalA: null,
            originalB: '',
            oldCol: null,
            newCol: col,
          });
        }
      }
    });

    const result = [];
    const processedHeaders = new Set();

    // First pass: Add columns in new file's order
    Object.keys(newHeaders).forEach((newCol) => {
      for (let [normalized, mapping] of headerMap) {
        if (mapping.newCol === newCol && !processedHeaders.has(normalized)) {
          processedHeaders.add(normalized);

          const displayName = mapping.originalB || mapping.originalA || '(Blank Column)';

          result.push({
            header: displayName.startsWith('__empty_') ? '(Blank Column)' : displayName,
            oldCol: mapping.oldCol,
            newCol: mapping.newCol,
            type: mapping.oldCol && mapping.newCol ? 'normal' : mapping.oldCol ? 'deleted' : 'added',
          });
          break;
        }
      }
    });

    // Second pass: Add deleted columns (only in old file)
    for (let [normalized, mapping] of headerMap) {
      if (!processedHeaders.has(normalized)) {
        const displayName = mapping.originalA || mapping.originalB || '(Blank Column)';

        result.push({
          header: displayName.startsWith('__empty_') ? '(Blank Column)' : displayName,
          oldCol: mapping.oldCol,
          newCol: mapping.newCol,
          type: 'deleted',
        });
      }
    }

    return result;
  }

  /**
   * buildUnifiedHeader(sheetDiff)
   * Build table header with two rows:
   * - Row 1: Column letters (with +/- prefix for added/deleted columns)
   * - Row 2: Column names
   *
   * @param {Object} sheetDiff - Comparison results
   * @returns {HTMLElement} Table header element
   */
  buildUnifiedHeader(sheetDiff) {
    const thead = document.createElement('thead');
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);

    // Row 1: Index columns + Column letters
    const tr1 = document.createElement('tr');

    const th1Old = document.createElement('th');
    th1Old.className = 'index-col';
    th1Old.textContent = 'Old';
    th1Old.rowSpan = 2;
    tr1.appendChild(th1Old);

    const th1New = document.createElement('th');
    th1New.className = 'index-col';
    th1New.textContent = 'New';
    th1New.rowSpan = 2;
    tr1.appendChild(th1New);

    unifiedColumns.forEach((col) => {
      const th = document.createElement('th');

      let colLabel = col.newCol || col.oldCol;

      if (col.type === 'added') {
        th.className = 'col-added';
        colLabel = `+${col.newCol}`;
      } else if (col.type === 'deleted') {
        th.className = 'col-deleted';
        colLabel = `−${col.oldCol}`;
      }

      th.textContent = colLabel;
      tr1.appendChild(th);
    });

    // Row 2: Column names
    const tr2 = document.createElement('tr');

    unifiedColumns.forEach((col) => {
      const th = document.createElement('th');

      if (col.type === 'added') {
        th.className = 'col-added';
      } else if (col.type === 'deleted') {
        th.className = 'col-deleted';
      }

      th.textContent = col.header;
      tr2.appendChild(th);
    });

    thead.appendChild(tr1);
    thead.appendChild(tr2);
    return thead;
  }

  /**
   * buildUnifiedBody(sheetDiff, keyColumn)
   * Build table body with all rows (matched/added/deleted)
   * Uses interleaved sorting: deleted rows appear before next new row position
   * Applies visual styling to changed/added/deleted cells and rows
   *
   * @param {Object} sheetDiff - Comparison results from DiffEngine
   * @param {string} keyColumn - Key column used for matching
   * @returns {HTMLElement} Table body element
   */
  buildUnifiedBody(sheetDiff, keyColumn = 'A') {
    const tbody = document.createElement('tbody');
    const allRows = this.getAllRows(sheetDiff, keyColumn);
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);
    const cellChanges = this.buildCellChangeMap(sheetDiff.differences);
    const rowChanges = this.buildRowChangeMap(sheetDiff.rowChanges);

    console.log('🔍 Row Changes Map:');
    rowChanges.forEach((change, key) => {
      console.log(`   Key: "${key}" → Type: ${change.type}`);
    });

    allRows.forEach((rowInfo) => {
      const tr = document.createElement('tr');

      console.log(`🔑 Processing row: key="${rowInfo.key}", originalKey="${rowInfo.originalKey}", oldIndex=${rowInfo.oldIndex}, newIndex=${rowInfo.newIndex}`);

      // Determine row type by presence of old/new index
      let rowType = null;

      if (rowInfo.oldIndex !== null && rowInfo.newIndex !== null) {
        rowType = 'matched';
      } else if (rowInfo.oldIndex !== null && rowInfo.newIndex === null) {
        rowType = 'deleted';
      } else if (rowInfo.oldIndex === null && rowInfo.newIndex !== null) {
        rowType = 'added';
      }

      console.log(`   ✅ Row type: ${rowType}`);

      // Apply row-level styling
      if (rowType === 'added') {
        tr.className = 'row-added';
      } else if (rowType === 'deleted') {
        tr.className = 'row-deleted';
      }

      // Old Index column
      const tdOldIdx = document.createElement('td');
      tdOldIdx.className = 'index-cell old-idx';
      tdOldIdx.textContent = rowInfo.oldIndex !== null ? rowInfo.oldIndex : '-';
      tr.appendChild(tdOldIdx);

      // New Index column
      const tdNewIdx = document.createElement('td');
      tdNewIdx.className = 'index-cell new-idx';
      tdNewIdx.textContent = rowInfo.newIndex !== null ? rowInfo.newIndex : '-';
      tr.appendChild(tdNewIdx);

      // Render data cells
      unifiedColumns.forEach((col) => {
        const td = document.createElement('td');

        const oldValue = col.oldCol ? rowInfo.oldRow?.[col.oldCol] : null;
        const newValue = col.newCol ? rowInfo.newRow?.[col.newCol] : null;

        const cellKey = `${rowInfo.newIndex || rowInfo.oldIndex}-${col.header}`;
        const cellDiff = cellChanges.get(cellKey);

        // 🆕 Build tooltip data
        const tooltipData = {
          header: col.header,
          oldCell: col.oldCol && rowInfo.oldIndex ? `${col.oldCol}${rowInfo.oldIndex}` : null,
          newCell: col.newCol && rowInfo.newIndex ? `${col.newCol}${rowInfo.newIndex}` : null,
        };

        if (cellDiff) {
          // Cell value changed
          td.className = 'cell-modified';
          td.innerHTML = `
      <div class="cell-value-change">
        <span class="old-value">${this.formatValue(cellDiff.oldValue)}</span>
        <span class="value-separator">→</span>
        <span class="new-value">${this.formatValue(cellDiff.newValue)}</span>
      </div>
    `;
        } else if (col.type === 'added') {
          // Column was added
          td.className = 'cell-added';
          td.innerHTML = this.formatValue(newValue);
        } else if (col.type === 'deleted') {
          // Column was deleted
          td.className = 'cell-deleted';
          td.innerHTML = this.formatValue(oldValue);
        } else if (rowType === 'deleted') {
          // Row was deleted
          td.className = 'cell-deleted';
          td.innerHTML = this.formatValue(oldValue);
        } else if (rowType === 'added') {
          // Row was added
          td.className = 'cell-added';
          td.innerHTML = this.formatValue(newValue);
        } else {
          // Cell unchanged
          td.className = 'cell-unchanged';
          td.innerHTML = this.formatValue(newValue || oldValue);
        }

        // 🆕 Attach tooltip on hover
        this.attachCellTooltip(td, tooltipData);

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    return tbody;
  }

  /**
   * getAllRows(sheetDiff, keyColumn)
   * INTERLEAVED SORTING
   *
   * Build complete row list with proper matching and sorting:
   * 1. Match rows by key column value
   * 2. Handle blank key columns (match by position + content)
   * 3. Sort with interleaved logic:
   *    - Matched/Added rows: Follow new file order
   *    - Deleted rows: Insert before next new row position based on old row position
   *
   * Sorting Algorithm:
   * - Matched/Added → sortKey = newIndex
   * - Deleted → sortKey = (next new row's old position) - 0.5
   *   This ensures deleted rows appear immediately before the new row at the same position
   *
   * @param {Object} sheetDiff - Comparison results with oldData and newData
   * @param {string} keyColumn - Column letter used for matching (e.g. 'A', 'B')
   * @returns {Array<Object>} Sorted array of row objects with { key, oldRow, oldIndex, newRow, newIndex }
   */
  getAllRows(sheetDiff, keyColumn = 'A') {
    const rowMap = new Map();
    const headerRowA = sheetDiff.headerRowA || 1;
    const headerRowB = sheetDiff.headerRowB || 1;

    const oldDataRows = sheetDiff.oldData.slice(1); // Exclude header row
    const newDataRows = sheetDiff.newData.slice(1); // Exclude header row

    // ====================================
    // Phase 1: Process Old File Rows
    // ====================================
    oldDataRows.forEach((row, index) => {
      const keyValue = String(row[keyColumn] || '').trim();
      const excelRowNumber = headerRowA + 1 + index; // Convert to Excel row number

      if (keyValue) {
        // Non-blank key: Use key value as map key
        rowMap.set(keyValue, {
          key: keyValue,
          originalKey: keyValue,
          oldRow: row,
          oldIndex: excelRowNumber,
          newRow: null,
          newIndex: null,
        });
      } else {
        // Blank key: Use position-based unique key
        const uniqueKey = `__blank_old_${excelRowNumber}`;
        rowMap.set(uniqueKey, {
          key: `(empty)`,
          originalKey: null,
          oldRow: row,
          oldIndex: excelRowNumber,
          newRow: null,
          newIndex: null,
          isBlankKey: true,
        });
      }
    });

    // ====================================
    // Phase 2: Process New File Rows
    // ====================================
    newDataRows.forEach((row, index) => {
      const keyValue = String(row[keyColumn] || '').trim();
      const excelRowNumber = headerRowB + 1 + index;

      if (keyValue) {
        // Non-blank key: Match with existing row or create new
        if (rowMap.has(keyValue)) {
          const existing = rowMap.get(keyValue);
          existing.newRow = row;
          existing.newIndex = excelRowNumber;
        } else {
          rowMap.set(keyValue, {
            key: keyValue,
            originalKey: keyValue,
            oldRow: null,
            oldIndex: null,
            newRow: row,
            newIndex: excelRowNumber,
          });
        }
      } else {
        // Blank key: Try to match by position + content
        const positionKey = `__blank_old_${excelRowNumber}`;

        if (rowMap.has(positionKey)) {
          const existing = rowMap.get(positionKey);
          const allColumnsMatch = this.areRowsIdentical(existing.oldRow, row);

          if (allColumnsMatch) {
            // Content matches: Treat as matched row
            existing.newRow = row;
            existing.newIndex = excelRowNumber;
          } else {
            // Content differs: Treat as separate added row
            const newKey = `__blank_new_${excelRowNumber}`;
            rowMap.set(newKey, {
              key: `(empty)`,
              originalKey: null,
              oldRow: null,
              oldIndex: null,
              newRow: row,
              newIndex: excelRowNumber,
              isBlankKey: true,
            });
          }
        } else {
          // No matching position in old file: Treat as added row
          const newKey = `__blank_new_${excelRowNumber}`;
          rowMap.set(newKey, {
            key: `(empty)`,
            originalKey: null,
            oldRow: null,
            oldIndex: null,
            newRow: row,
            newIndex: excelRowNumber,
            isBlankKey: true,
          });
        }
      }
    });

    const rows = Array.from(rowMap.values());

    console.log('📊 Total rows before sorting:', rows.length);
    console.log('📊 Deleted rows:', rows.filter((r) => r.oldIndex && !r.newIndex).length);
    console.log('📊 Added rows:', rows.filter((r) => !r.oldIndex && r.newIndex).length);
    console.log('📊 Matched rows:', rows.filter((r) => r.oldIndex && r.newIndex).length);

    // ====================================
    // Phase 3: Interleaved Sorting
    // ====================================
    rows.sort((a, b) => {
      const getType = (row) => {
        if (row.oldIndex !== null && row.newIndex !== null) return 'matched';
        if (row.oldIndex === null && row.newIndex !== null) return 'added';
        if (row.oldIndex !== null && row.newIndex === null) return 'deleted';
        return 'unknown';
      };

      const typeA = getType(a);
      const typeB = getType(b);

      // Calculate display position for sorting
      const getDisplayPosition = (row, type) => {
        if (type === 'matched' || type === 'added') {
          // Matched/Added: Use newIndex directly (scaled up)
          return row.newIndex * 1000;
        } else if (type === 'deleted') {
          // Deleted: Find next new row after this old position
          const newIndexes = rows
            .filter((r) => r.newIndex !== null)
            .map((r) => r.newIndex)
            .sort((x, y) => x - y);

          // Find next new row whose old position is after this deleted row
          const nextNewIndex = newIndexes.find((idx) => {
            const matchedRow = rows.find((r) => r.newIndex === idx && r.oldIndex !== null);
            if (matchedRow && matchedRow.oldIndex > row.oldIndex) {
              return true;
            }
            return false;
          });

          if (nextNewIndex !== undefined) {
            // Insert before next new row
            return nextNewIndex * 1000 - 500 + (row.oldIndex % 1000);
          } else {
            // No next new row: Append at end
            return (Math.max(...newIndexes, 0) + 1) * 1000 + row.oldIndex;
          }
        }
        return Infinity;
      };

      const posA = getDisplayPosition(a, typeA);
      const posB = getDisplayPosition(b, typeB);

      if (posA !== posB) {
        return posA - posB;
      }

      // Tie-breaker: deleted first, then matched, then added
      const typePriority = {
        deleted: 0,
        matched: 1,
        added: 2,
        unknown: 3,
      };

      return typePriority[typeA] - typePriority[typeB];
    });

    console.log('✅ Rows sorted with interleaved deleted rows');

    // Debug: Show final order
    console.log('🔍 Final row order:');
    rows.forEach((row, idx) => {
      const type = row.oldIndex && row.newIndex ? 'matched' : !row.oldIndex && row.newIndex ? 'added' : 'deleted';
      console.log(`  ${idx + 1}. [${type}] Key=${row.key}, Old=${row.oldIndex}, New=${row.newIndex}`);
    });

    return rows;
  }

  /**
   * areRowsIdentical(row1, row2)
   * Check if two rows have identical content across all columns
   * Used for matching blank-key rows by content
   *
   * @param {Object} row1 - First row data
   * @param {Object} row2 - Second row data
   * @returns {boolean} True if all cell values match
   */
  areRowsIdentical(row1, row2) {
    if (!row1 || !row2) return false;

    const allKeys = new Set([...Object.keys(row1), ...Object.keys(row2)]);

    for (const key of allKeys) {
      const val1 = this.normalizeValue(row1[key]);
      const val2 = this.normalizeValue(row2[key]);

      if (val1 !== val2) {
        return false;
      }
    }

    return true;
  }

  /**
   * normalizeValue(cell)
   * Normalize cell value for comparison
   * Removes invisible characters and trims whitespace
   *
   * @param {any} cell - Cell value
   * @returns {string} Normalized string value
   */
  normalizeValue(cell) {
    if (cell === null || cell === undefined) {
      return '';
    }

    let value = String(cell);
    value = value.replace(/[\u200B-\u200D\uFEFF]/g, ''); // Remove zero-width chars
    value = value.trim();

    return value;
  }

  /**
   * buildRowChangeMap(rowChanges)
   * Build Map for quick lookup of row changes by row key
   *
   * @param {Array<Object>} rowChanges - Array of row changes from DiffEngine
   * @returns {Map} Map of rowKey -> change object
   */
  buildRowChangeMap(rowChanges) {
    const map = new Map();
    rowChanges.forEach((change) => {
      map.set(change.rowKey, change);
    });
    return map;
  }

  /**
   * buildCellChangeMap(differences)
   * Build Map for quick lookup of cell changes by row-header key
   *
   * @param {Array<Object>} differences - Array of cell differences from DiffEngine
   * @returns {Map} Map of "rowIndex-headerName" -> diff object
   */
  buildCellChangeMap(differences) {
    const map = new Map();
    differences.forEach((diff) => {
      const key = `${diff.row}-${diff.header}`;
      map.set(key, diff);
    });
    return map;
  }

  /**
   * formatValue(value)
   * Format cell value for display
   * Shows "Blank" for null/undefined/empty values
   *
   * @param {any} value - Cell value
   * @returns {string} Formatted HTML string
   */
  formatValue(value) {
    if (value === null || value === undefined || value === '') {
      return '<em class="empty-cell">Blank</em>';
    }
    return String(value);
  }

  /**
   * setupChangeNavigation()
   * Setup keyboard shortcuts and button handlers for change navigation
   * Keyboard: 'N' = next, 'P' = previous
   */
  setupChangeNavigation() {
    const prevBtn = document.getElementById('prevChangeBtn');
    const nextBtn = document.getElementById('nextChangeBtn');

    if (prevBtn) {
      prevBtn.onclick = () => this.navigateToChange('prev');
    }

    if (nextBtn) {
      nextBtn.onclick = () => this.navigateToChange('next');
    }

    this.setupKeyboardShortcuts();
  }

  /**
   * setupKeyboardShortcuts()
   * Setup keyboard event listeners for change navigation
   * P = previous change, N = next change
   */
  setupKeyboardShortcuts() {
    if (this.keyboardHandler) {
      document.removeEventListener('keydown', this.keyboardHandler);
    }

    this.keyboardHandler = (e) => {
      const isTyping = ['INPUT', 'TEXTAREA', 'SELECT'].includes(e.target.tagName);
      if (isTyping) return;

      const key = e.key.toLowerCase();

      if (key === 'p') {
        e.preventDefault();
        this.navigateToChange('prev');
      } else if (key === 'n') {
        e.preventDefault();
        this.navigateToChange('next');
      }
    };

    document.addEventListener('keydown', this.keyboardHandler);
  }

  /**
   * collectChangedCells()
   * Collect all changed cells (modified/added/deleted) for navigation
   * Stores references to DOM elements for highlighting
   */
  collectChangedCells() {
    this.changedCells = [];
    this.currentChangeIndex = -1;

    const table = document.querySelector('#unifiedTableContainer .diff-table');
    if (!table) {
      this.updateNavigationUI();
      return;
    }

    const tbody = table.querySelector('tbody');
    if (!tbody) {
      this.updateNavigationUI();
      return;
    }

    const rows = tbody.querySelectorAll('tr');
    rows.forEach((row, rowIndex) => {
      const cells = row.querySelectorAll('td.cell-modified, td.cell-added, td.cell-deleted');
      cells.forEach((cell) => {
        this.changedCells.push({ row, cell });
      });
    });

    this.updateNavigationUI();
  }

  /**
   * navigateToChange(direction)
   * Navigate to next or previous changed cell
   *
   * @param {string} direction - 'next' or 'prev'
   */
  navigateToChange(direction) {
    if (this.changedCells.length === 0) return;

    if (direction === 'next') {
      this.currentChangeIndex = (this.currentChangeIndex + 1) % this.changedCells.length;
    } else if (direction === 'prev') {
      this.currentChangeIndex = (this.currentChangeIndex - 1 + this.changedCells.length) % this.changedCells.length;
    }

    this.updateNavigationUI();
    this.scrollToChange();
  }

  /**
   * scrollToChange()
   * Scroll to currently selected changed cell and highlight it temporarily
   */
  scrollToChange() {
    if (this.currentChangeIndex < 0 || this.currentChangeIndex >= this.changedCells.length) return;

    const { cell } = this.changedCells[this.currentChangeIndex];

    // Remove previous highlights
    document.querySelectorAll('.cell-highlighted').forEach((c) => {
      c.classList.remove('cell-highlighted');
    });

    // Highlight current cell
    cell.classList.add('cell-highlighted');

    // Scroll to cell
    cell.scrollIntoView({ behavior: 'smooth', block: 'center' });

    // Remove highlight after 2 seconds
    setTimeout(() => {
      cell.classList.remove('cell-highlighted');
    }, 2000);
  }

  /**
   * updateNavigationUI()
   * Update change counter display and button states
   */
  updateNavigationUI() {
    const counter = document.getElementById('changeCounter');
    const prevBtn = document.getElementById('prevChangeBtn');
    const nextBtn = document.getElementById('nextChangeBtn');

    if (!counter) return;

    const total = this.changedCells.length;
    const current = this.currentChangeIndex >= 0 ? this.currentChangeIndex + 1 : 0;

    counter.textContent = `${current} / ${total}`;

    if (prevBtn && nextBtn) {
      prevBtn.disabled = total === 0;
      nextBtn.disabled = total === 0;
    }
  }

  /**
   * compareByRowPosition(sheetA, sheetB, headerRowA, headerRowB)
   * Compare sheets by row position (no key column matching)
   * Matches row N in file A with row N in file B
   *
   * @param {Array<Object>} sheetA - Data from first sheet
   * @param {Array<Object>} sheetB - Data from second sheet
   * @param {number} headerRowA - Header row number for sheet A
   * @param {number} headerRowB - Header row number for sheet B
   */
  compareByRowPosition(sheetA, sheetB, headerRowA, headerRowB) {
    console.log('🔢 Comparing by row position...');

    const unifiedColumns = this.getUnifiedColumns({
      oldData: sheetA,
      newData: sheetB,
    });

    this.changedCells = [];

    const dataRowsA = sheetA.slice(1);
    const dataRowsB = sheetB.slice(1);

    const maxRows = Math.max(dataRowsA.length, dataRowsB.length);

    const differences = [];
    const rowChanges = [];

    // Compare rows by position
    for (let i = 0; i < maxRows; i++) {
      const rowA = dataRowsA[i] || {};
      const rowB = dataRowsB[i] || {};

      const excelRowA = headerRowA + 1 + i;
      const excelRowB = headerRowB + 1 + i;

      const rowExistsInA = i < dataRowsA.length;
      const rowExistsInB = i < dataRowsB.length;

      if (!rowExistsInA) {
        // Row only in B: Added
        rowChanges.push({
          rowKey: `position-${i}`,
          type: 'added',
          newRowIndex: excelRowB,
          row: rowB,
        });
      } else if (!rowExistsInB) {
        // Row only in A: Deleted
        rowChanges.push({
          rowKey: `position-${i}`,
          type: 'deleted',
          oldRowIndex: excelRowA,
          row: rowA,
        });
      } else {
        // Row in both: Compare cells
        unifiedColumns.forEach(({ header, oldCol, newCol }) => {
          const valueA = oldCol ? (rowA[oldCol] ?? '') : '';
          const valueB = newCol ? (rowB[newCol] ?? '') : '';

          const normalizedA = String(valueA).trim();
          const normalizedB = String(valueB).trim();

          if (normalizedA !== normalizedB) {
            differences.push({
              row: excelRowB,
              header: header,
              oldCol: oldCol,
              newCol: newCol,
              oldValue: valueA,
              newValue: valueB,
            });
          }
        });
      }
    }

    const singleSheetDiff = {
      sheetName: 'Position-based Comparison',
      differences: differences,
      rowChanges: rowChanges,
      columnChanges: [],
      oldData: sheetA,
      newData: sheetB,
      headerRowA: headerRowA,
      headerRowB: headerRowB,
    };

    this.renderUnifiedTableByPosition(singleSheetDiff, headerRowA, headerRowB);
  }

  /**
   * renderUnifiedTableByPosition(sheetDiff, headerRowA, headerRowB)
   * Render unified table for position-based comparison
   * Similar to renderUnifiedTable but uses position-based matching
   *
   * @param {Object} sheetDiff - Comparison results
   * @param {number} headerRowA - Header row number for sheet A
   * @param {number} headerRowB - Header row number for sheet B
   */
  renderUnifiedTableByPosition(sheetDiff, headerRowA, headerRowB) {
    const container = document.getElementById('unifiedTableContainer');
    if (!container) {
      console.error('unifiedTableContainer element not found');
      return;
    }

    container.innerHTML = '';

    const table = document.createElement('table');
    table.className = 'unified-table diff-table';

    const thead = this.buildUnifiedHeader(sheetDiff);
    table.appendChild(thead);

    const tbody = this.buildUnifiedBodyByPosition(sheetDiff, headerRowA, headerRowB);
    table.appendChild(tbody);

    const wrapper = document.createElement('div');
    wrapper.className = 'table-wrapper';
    wrapper.appendChild(table);

    container.appendChild(wrapper);

    this.collectChangedCells();
    this.setupCellClickNavigation();
  }

  /**
   * buildUnifiedBodyByPosition(sheetDiff, headerRowA, headerRowB)
   * Build table body for position-based comparison
   * Rows are matched by position (row 1 vs row 1, row 2 vs row 2, etc.)
   *
   * @param {Object} sheetDiff - Comparison results
   * @param {number} headerRowA - Header row number for sheet A
   * @param {number} headerRowB - Header row number for sheet B
   * @returns {HTMLElement} Table body element
   */
  buildUnifiedBodyByPosition(sheetDiff, headerRowA, headerRowB) {
    const tbody = document.createElement('tbody');
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);
    const cellChanges = this.buildCellChangeMap(sheetDiff.differences);
    const rowChanges = this.buildRowChangeMap(sheetDiff.rowChanges);

    const dataRowsA = sheetDiff.oldData.slice(1);
    const dataRowsB = sheetDiff.newData.slice(1);
    const maxRows = Math.max(dataRowsA.length, dataRowsB.length);

    for (let i = 0; i < maxRows; i++) {
      const tr = document.createElement('tr');

      const rowA = dataRowsA[i];
      const rowB = dataRowsB[i];

      const excelRowA = headerRowA + 1 + i;
      const excelRowB = headerRowB + 1 + i;

      const rowExistsInA = i < dataRowsA.length;
      const rowExistsInB = i < dataRowsB.length;

      const rowChange = rowChanges.get(`position-${i}`);

      // Apply row-level styling
      if (rowChange?.type === 'added') {
        tr.className = 'row-added';
      } else if (rowChange?.type === 'deleted') {
        tr.className = 'row-deleted';
      }

      // Old Index column
      const tdOldIdx = document.createElement('td');
      tdOldIdx.className = 'index-cell old-idx';
      tdOldIdx.textContent = rowExistsInA ? excelRowA : '-';
      tr.appendChild(tdOldIdx);

      // New Index column
      const tdNewIdx = document.createElement('td');
      tdNewIdx.className = 'index-cell new-idx';
      tdNewIdx.textContent = rowExistsInB ? excelRowB : '-';
      tr.appendChild(tdNewIdx);

      // Render data cells
      unifiedColumns.forEach((col) => {
        const td = document.createElement('td');

        const oldValue = col.oldCol && rowA ? (rowA[col.oldCol] ?? null) : null;
        const newValue = col.newCol && rowB ? (rowB[col.newCol] ?? null) : null;

        const cellKey = `${excelRowB}-${col.header}`;
        const cellDiff = cellChanges.get(cellKey);

        if (cellDiff) {
          // Cell value changed
          td.className = 'cell-modified';
          td.innerHTML = `
            <div class="cell-value-change">
              <span class="old-value">${this.formatValue(cellDiff.oldValue)}</span>
              <span class="value-separator">→</span>
              <span class="new-value">${this.formatValue(cellDiff.newValue)}</span>
            </div>
          `;
        } else if (col.type === 'added') {
          td.className = 'cell-added';
          td.innerHTML = this.formatValue(newValue);
        } else if (col.type === 'deleted') {
          td.className = 'cell-deleted';
          td.innerHTML = this.formatValue(oldValue);
        } else if (rowChange?.type === 'deleted') {
          td.className = 'cell-deleted';
          td.innerHTML = this.formatValue(oldValue);
        } else if (rowChange?.type === 'added') {
          td.className = 'cell-added';
          td.innerHTML = this.formatValue(newValue);
        } else {
          td.className = 'cell-unchanged';
          td.innerHTML = this.formatValue(newValue || oldValue);
        }
        // 🆕 Build tooltip data
        const tooltipData = {
          header: col.header,
          oldCell: col.oldCol && rowExistsInA ? `${col.oldCol}${excelRowA}` : null,
          newCell: col.newCol && rowExistsInB ? `${col.newCol}${excelRowB}` : null,
        };

        // 🆕 Attach tooltip on hover
        this.attachCellTooltip(td, tooltipData);

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    }

    return tbody;
  }

  /**
   * attachCellTooltip(td, tooltipData)
   * Attach hover/tap tooltip showing cell position info
   * Desktop: Hover to show, auto-updates on mouse move
   * Mobile: Tap to show, auto-closes on scroll
   *
   * @param {HTMLElement} td - Table cell element
   * @param {Object} tooltipData - { header, oldCell, newCell }
   */
  attachCellTooltip(td, tooltipData) {
    let tooltip = null;
    let isMobile = 'ontouchstart' in window;
    let scrollListener = null;
    let outsideClickHandler = null;

    // Helper: Create tooltip
    const createTooltip = () => {
      if (tooltip) return; // Already exists

      tooltip = document.createElement('div');
      tooltip.className = 'cell-tooltip';

      // Build tooltip content
      let content = `<div class="cell-tooltip-label">${tooltipData.header}</div>`;

      if (tooltipData.oldCell) {
        content += `<div class="cell-tooltip-row"><span>Old:</span><span>${tooltipData.oldCell}</span></div>`;
      }

      if (tooltipData.newCell) {
        content += `<div class="cell-tooltip-row"><span>New:</span><span>${tooltipData.newCell}</span></div>`;
      }

      tooltip.innerHTML = content;
      document.body.appendChild(tooltip);

      // Position tooltip
      positionTooltip(td);

      // Only attach scroll listener on mobile
      if (isMobile) {
        attachScrollListener();
      }
    };

    // Helper: Position tooltip
    const positionTooltip = (target) => {
      if (!tooltip) return;

      const rect = target.getBoundingClientRect();
      const scrollY = window.scrollY || window.pageYOffset;
      const scrollX = window.scrollX || window.pageXOffset;

      // Position above cell (centered)
      tooltip.style.position = 'absolute';
      tooltip.style.left = `${rect.left + scrollX + rect.width / 2}px`;
      tooltip.style.top = `${rect.top + scrollY - 10}px`;
      tooltip.style.transform = 'translate(-50%, -100%)';
    };

    // Helper: Remove tooltip
    const removeTooltip = () => {
      if (tooltip) {
        tooltip.remove();
        tooltip = null;
      }

      // Remove scroll listener
      removeScrollListener();

      // Remove outside click handler
      removeOutsideClickHandler();
    };

    // Helper: Attach scroll listener (mobile only)
    const attachScrollListener = () => {
      if (scrollListener) return; // Already attached

      scrollListener = () => {
        removeTooltip();
      };

      // Listen to both scroll and touchmove
      window.addEventListener('scroll', scrollListener, { passive: true });
      document.addEventListener('touchmove', scrollListener, { passive: true });

      // Also check table wrapper scroll
      const tableWrapper = td.closest('.table-wrapper');
      if (tableWrapper) {
        tableWrapper.addEventListener('scroll', scrollListener, { passive: true });
      }
    };

    // Helper: Remove scroll listener
    const removeScrollListener = () => {
      if (!scrollListener) return;

      window.removeEventListener('scroll', scrollListener);
      document.removeEventListener('touchmove', scrollListener);

      const tableWrapper = td.closest('.table-wrapper');
      if (tableWrapper) {
        tableWrapper.removeEventListener('scroll', scrollListener);
      }

      scrollListener = null;
    };

    // Helper: Attach outside click handler (mobile only)
    const attachOutsideClickHandler = () => {
      if (outsideClickHandler) return; // Already attached

      outsideClickHandler = (e) => {
        // Check if click is outside both tooltip and cell
        if (tooltip && !td.contains(e.target) && !tooltip.contains(e.target)) {
          removeTooltip();
        }
      };

      // Delay to avoid immediate trigger
      setTimeout(() => {
        document.addEventListener('click', outsideClickHandler, { capture: true });
      }, 100);
    };

    // Helper: Remove outside click handler
    const removeOutsideClickHandler = () => {
      if (!outsideClickHandler) return;

      document.removeEventListener('click', outsideClickHandler, { capture: true });
      outsideClickHandler = null;
    };

    // ========================================
    // Event Handlers
    // ========================================

    if (isMobile) {
      // ========== MOBILE: Tap to toggle tooltip ==========
      td.addEventListener('click', (e) => {
        e.stopPropagation(); // Prevent event bubbling

        if (tooltip) {
          // Tooltip already visible: Hide it
          removeTooltip();
        } else {
          // Remove any other visible tooltips first
          document.querySelectorAll('.cell-tooltip').forEach((t) => t.remove());

          // Show tooltip for this cell
          createTooltip();

          // Attach outside click handler
          attachOutsideClickHandler();
        }
      });
    } else {
      // ========== DESKTOP: Hover to show tooltip ==========
      td.addEventListener('mouseenter', () => {
        createTooltip();
      });

      td.addEventListener('mousemove', () => {
        positionTooltip(td);
      });

      td.addEventListener('mouseleave', () => {
        removeTooltip();
      });
    }
  }
}

export default DiffViewer;
