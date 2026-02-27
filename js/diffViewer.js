/**
 * diffViewer.js
 * Renders Excel comparison results in a unified table view
 * Displays side-by-side comparison with old/new indices and change highlighting
 * Supports change navigation with keyboard shortcuts and mouse clicks
 */

import DiffEngine from './diffEngine.js';

/**
 * DiffViewer Class
 * Manages the visual presentation of Excel diff results
 * Provides interactive navigation through changes
 */
class DiffViewer {
  /**
   * Constructor
   * Initializes the viewer state
   */
  constructor() {
    this.dataA = null; // Parsed data from File A
    this.dataB = null; // Parsed data from File B
    this.diffResults = null; // Complete diff results
    this.changedCells = []; // Array of changed cell elements for navigation
    this.currentChangeIndex = -1; // Index of currently highlighted change
  }

  /**
   * init(dataA, dataB, diffResults)
   * Initializes the viewer with comparison data
   * Populates sheet dropdowns and auto-selects matching sheets
   *
   * @param {Object} dataA - Parsed data from File A
   * @param {Object} dataB - Parsed data from File B
   * @param {Object} diffResults - Complete diff results
   */

  init(dataA, dataB, diffResults) {
    this.dataA = dataA;
    this.dataB = dataB;
    this.diffResults = diffResults;

    console.log('🔧 Initializing DiffViewer...');

    // STEP 1: Populate sheet dropdowns
    this.populateSheetDropdowns();

    // STEP 2: Find matching sheets
    const matchedSheet = this.findMatchingSheet();

    if (matchedSheet) {
      // console.log(`✅ Found matching sheets: ${matchedSheet.sheetA} vs ${matchedSheet.sheetB}`);

      // STEP 3: Set sheet values
      document.getElementById('sheetSelectA').value = matchedSheet.sheetA;
      document.getElementById('sheetSelectB').value = matchedSheet.sheetB;

      // STEP 4: Populate header row dropdowns
      this.populateHeaderRowDropdowns();

      // STEP 5: Populate key column dropdown
      this.populateKeyColumnDropdown();

      // STEP 6: Compare sheets
      this.compareSelectedSheets();
    } else {
      // No matching sheet names found, but still populate dropdowns
      console.warn('No matching sheet names found');

      // Auto-select first sheet from each file
      const firstSheetA = this.dataA.sheetNames?.[0];
      const firstSheetB = this.dataB.sheetNames?.[0];

      if (firstSheetA) {
        document.getElementById('sheetSelectA').value = firstSheetA;
      }

      if (firstSheetB) {
        document.getElementById('sheetSelectB').value = firstSheetB;
      }

      // Populate header row dropdowns
      this.populateHeaderRowDropdowns();

      // Populate key column dropdown
      this.populateKeyColumnDropdown();

      // Note: Don't auto-compare yet, let user verify selections
    }

    this.setupChangeNavigation();
    console.log('DiffViewer initialized');
  }

  /**
   * populateHeaderRowDropdowns()
   * Populates header row dropdowns with available rows
   * Detects row count from currently selected sheets
   */
  populateHeaderRowDropdowns() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');

    if (!headerRowA || !headerRowB) {
      console.error('Header row dropdown elements not found');
      return;
    }

    // Get row counts from the currently selected sheets
    const selectedSheetA = sheetSelectA?.value;
    const selectedSheetB = sheetSelectB?.value;

    // Only proceed if both sheets have been selected
    if (!selectedSheetA || !selectedSheetB) {
      console.warn('⚠️ Sheets not selected yet, skipping header row population');
      return;
    }

    const rowCountA = this.dataA?.sheets[selectedSheetA]?.rowCount || 20;
    const rowCountB = this.dataB?.sheets[selectedSheetB]?.rowCount || 20;

    // Populate File A header row dropdown
    headerRowA.innerHTML = '';
    const maxRowsA = Math.min(rowCountA, 50);
    for (let i = 1; i <= maxRowsA; i++) {
      const option = document.createElement('option');
      option.value = i;
      option.textContent = `Row ${i}`;
      headerRowA.appendChild(option);
    }

    // Populate File B header row dropdown
    headerRowB.innerHTML = '';
    const maxRowsB = Math.min(rowCountB, 50);
    for (let i = 1; i <= maxRowsB; i++) {
      const option = document.createElement('option');
      option.value = i;
      option.textContent = `Row ${i}`;
      headerRowB.appendChild(option);
    }

    // Set default header row to 1
    headerRowA.value = '1';
    headerRowB.value = '1';

    // console.log(`📋 Header row dropdowns populated: A=${maxRowsA} rows, B=${maxRowsB} rows`);

    // Listen for sheet changes
    this.setupSheetChangeListeners();
  }

  /**
   * populateKeyColumnDropdown()
   * Populates key column dropdown with header columns from both files
   * Uses the selected header rows to extract column names
   */
  populateKeyColumnDropdown() {
    const keyColumnSelect = document.getElementById('keyColumnSelect');
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');

    if (!keyColumnSelect || !sheetSelectA || !sheetSelectB || !headerRowA || !headerRowB) {
      console.error('Key column dropdown elements not found');
      return;
    }

    const selectedSheetA = sheetSelectA.value;
    const selectedSheetB = sheetSelectB.value;
    const selectedHeaderA = parseInt(headerRowA.value) || 1;
    const selectedHeaderB = parseInt(headerRowB.value) || 1;

    // console.log(`🔍 Populating key column dropdown:`);
    // console.log(`   Sheet A: ${selectedSheetA}, Header Row: ${selectedHeaderA}`);
    // console.log(`   Sheet B: ${selectedSheetB}, Header Row: ${selectedHeaderB}`);

    if (!selectedSheetA || !selectedSheetB) {
      console.warn('⚠️ No sheets selected yet');
      keyColumnSelect.innerHTML = '<option value="">-- Select Key Column --</option>';
      return;
    }

    // Get sheet data
    const sheetA = this.dataA?.sheets[selectedSheetA]?.data || [];
    const sheetB = this.dataB?.sheets[selectedSheetB]?.data || [];

    if (sheetA.length === 0 || sheetB.length === 0) {
      console.warn('⚠️ One or both sheets are empty');
      keyColumnSelect.innerHTML = '<option value="">-- No data available --</option>';
      return;
    }

    // Get header rows
    const headerIndexA = selectedHeaderA - 1;
    const headerIndexB = selectedHeaderB - 1;

    if (headerIndexA >= sheetA.length || headerIndexB >= sheetB.length) {
      console.warn('⚠️ Invalid header row index');
      keyColumnSelect.innerHTML = '<option value="">-- Invalid header row --</option>';
      return;
    }

    // Extract headers
    const headersA = sheetA[headerIndexA] ? Object.values(sheetA[headerIndexA]) : [];
    const headersB = sheetB[headerIndexB] ? Object.values(sheetB[headerIndexB]) : [];

    // console.log('🔍 Headers A:', headersA);
    // console.log('🔍 Headers B:', headersB);

    // Find common columns
    const commonColumns = this.findCommonColumns(headersA, headersB);

    // console.log('✅ Common columns:', commonColumns);

    // ✅ Clear and populate dropdown
    keyColumnSelect.innerHTML = '';

    // ✅ Add default "Select Key Column" option
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = '── Select Key Column ──';
    keyColumnSelect.appendChild(defaultOption);

    // ✅ Add "(Use Row Position)" option
    const rowPositionOption = document.createElement('option');
    rowPositionOption.value = '__ROW_POSITION__'; // Special internal value
    rowPositionOption.textContent = '(Use Row Position)';
    keyColumnSelect.appendChild(rowPositionOption);

    // ✅ Add separator (optional, for visual clarity)
    const separator = document.createElement('option');
    separator.disabled = true;
    separator.textContent = '────────────────────────';
    keyColumnSelect.appendChild(separator);

    // ✅ Add common columns
    if (commonColumns.length === 0) {
      // ✅ No common columns found → Auto-select "Use Row Position"
      console.warn('⚠️ No common columns found, defaulting to Row Position');
      keyColumnSelect.value = '__ROW_POSITION__';
    } else {
      // ✅ Add all common columns
      let autoSelected = false;

      commonColumns.forEach((col) => {
        const option = document.createElement('option');
        option.value = col.colIndex;
        option.textContent = col.name;

        // ✅ Auto-select if looks like a key column
        const lowerName = col.name.toLowerCase();
        if (!autoSelected && (lowerName.includes('id') || lowerName.includes('email') || lowerName.includes('code') || lowerName.includes('題目'))) {
          option.selected = true;
          autoSelected = true;
          // console.log(`✅ Auto-selected key column: "${col.name}"`);
        }

        keyColumnSelect.appendChild(option);
      });

      // ✅ If no auto-select happened, use first column
      if (!autoSelected && commonColumns.length > 0) {
        keyColumnSelect.value = commonColumns[0].colIndex;
        // console.log(`✅ Auto-selected first column: "${commonColumns[0].name}"`);
      }
    }

    // console.log(`✅ Key column dropdown populated with ${commonColumns.length} columns`);
  }

  /**
   * findCommonColumns(headersA, headersB)
   * Finds columns that exist in both files
   *
   * @param {Array<string>} headersA - Headers from File A
   * @param {Array<string>} headersB - Headers from File B
   * @returns {Array<Object>} Common columns with {name, colIndex}
   */
  findCommonColumns(headersA, headersB) {
    const commonColumns = [];

    headersA.forEach((headerA, indexA) => {
      const normalizedA = String(headerA).trim().toLowerCase();

      headersB.forEach((headerB, indexB) => {
        const normalizedB = String(headerB).trim().toLowerCase();

        // Check if headers match (case-insensitive)
        if (normalizedA === normalizedB && normalizedA !== '') {
          // Check if not already added (avoid duplicates)
          const alreadyExists = commonColumns.some((col) => col.name === headerA);

          if (!alreadyExists) {
            commonColumns.push({
              name: headerA, // Original header name
              colIndex: this.getColumnLetter(indexA), // Convert to A, B, C, etc.
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
   * Converts column index to Excel-style letter (0 -> A, 1 -> B, etc.)
   *
   * @param {number} index - Column index (0-based)
   * @returns {string} Column letter (A, B, C, ..., Z, AA, AB, etc.)
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
   * Updates header row and key column dropdowns when user changes selection
   */
  setupSheetChangeListeners() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');

    // Listen for sheet changes
    if (sheetSelectA) {
      sheetSelectA.addEventListener('change', () => {
        this.updateHeaderRowDropdown('A');
        this.populateKeyColumnDropdown(); // NEW
      });
    }

    if (sheetSelectB) {
      sheetSelectB.addEventListener('change', () => {
        this.updateHeaderRowDropdown('B');
        this.populateKeyColumnDropdown(); // NEW
      });
    }

    // Listen for header row changes
    if (headerRowA) {
      headerRowA.addEventListener('change', () => {
        this.populateKeyColumnDropdown(); // NEW
      });
    }

    if (headerRowB) {
      headerRowB.addEventListener('change', () => {
        this.populateKeyColumnDropdown(); // NEW
      });
    }
  }

  /**
   * updateHeaderRowDropdown(type)
   * Updates header row dropdown when sheet selection changes
   *
   * @param {string} type - 'A' or 'B'
   */
  updateHeaderRowDropdown(type) {
    const sheetSelect = document.getElementById(`sheetSelect${type}`);
    const headerRowSelect = document.getElementById(`headerRow${type}`);
    const data = type === 'A' ? this.dataA : this.dataB;

    if (!sheetSelect || !headerRowSelect || !data) return;

    const selectedSheet = sheetSelect.value;
    const rowCount = data.sheets[selectedSheet]?.rowCount || 20;

    // Repopulate dropdown
    headerRowSelect.innerHTML = '';
    const maxRows = Math.min(rowCount, 50);
    for (let i = 1; i <= maxRows; i++) {
      const option = document.createElement('option');
      option.value = i;
      option.textContent = `Row ${i}`;
      headerRowSelect.appendChild(option);
    }

    // Reset to row 1
    headerRowSelect.value = '1';

    // console.log(`📋 Updated header row dropdown for File ${type}: ${maxRows} rows`);
  }

  /**
   * findMatchingSheet()
   * Finds first pair of sheets with identical names across both files
   *
   * @returns {Object|null} Object with {sheetA, sheetB} or null if no match
   */
  findMatchingSheet() {
    if (!this.dataA.sheetNames || !this.dataB.sheetNames) return null;

    // Find first sheet name that exists in both files
    for (let sheetA of this.dataA.sheetNames) {
      if (this.dataB.sheetNames.includes(sheetA)) {
        return { sheetA, sheetB: sheetA };
      }
    }

    return null;
  }

  /**
   * populateSheetDropdowns()
   * Fills both sheet selection dropdowns with available sheet names
   */
  populateSheetDropdowns() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');

    if (!sheetSelectA || !sheetSelectB) {
      console.error('Sheet dropdown elements not found');
      return;
    }

    // Clear existing options
    sheetSelectA.innerHTML = '';
    sheetSelectB.innerHTML = '';

    // Populate dropdown A with File A's sheets
    if (this.dataA && this.dataA.sheetNames) {
      this.dataA.sheetNames.forEach((sheetName) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelectA.appendChild(option);
      });
    }

    // Populate dropdown B with File B's sheets
    if (this.dataB && this.dataB.sheetNames) {
      this.dataB.sheetNames.forEach((sheetName) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelectB.appendChild(option);
      });
    }

    // console.log('✅ Sheet dropdowns populated');
  }

  /**
   * compareSelectedSheets()
   * Compares the currently selected sheets from both dropdowns
   * Uses selected header rows and key column for comparison
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

    // Validate selections
    if (!selectedSheetA || !selectedSheetB) {
      alert('Please select sheets to compare');
      return;
    }

    if (!selectedKeyColumn) {
      alert('⚠️ Please select a Key Column or "Use Row Position"');
      return;
    }

    // Get sheet data
    let sheetA = this.dataA.sheets[selectedSheetA]?.data || [];
    let sheetB = this.dataB.sheets[selectedSheetB]?.data || [];

    // Adjust data based on selected header row
    sheetA = this.adjustDataForHeaderRow(sheetA, selectedHeaderA);
    sheetB = this.adjustDataForHeaderRow(sheetB, selectedHeaderB);

    // Validate sheet data
    if (sheetA.length === 0 || sheetB.length === 0) {
      alert('Selected sheet is empty');
      return;
    }

    // ✅ Check if user wants row position matching
    if (selectedKeyColumn === '__ROW_POSITION__') {
      console.log('🔢 Using row position matching (no key column)');
      this.compareByRowPosition(sheetA, sheetB, selectedHeaderA, selectedHeaderB);
    } else {
      console.log(`🔑 Using key column: ${selectedKeyColumn}`);

      // Perform comparison with key column
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
   * Adjusts sheet data to treat specified row as header
   * Removes all rows before header row
   *
   * @param {Array<Object>} data - Original sheet data
   * @param {number} headerRow - Selected header row (1-based)
   * @returns {Array<Object>} Adjusted data with header row as first row
   */
  adjustDataForHeaderRow(data, headerRow) {
    if (!data || data.length === 0) return data;

    // Convert to 0-based index
    const headerIndex = headerRow - 1;

    // Validate index
    if (headerIndex < 0 || headerIndex >= data.length) {
      console.warn(`Invalid header row ${headerRow}, using row 1`);
      return data;
    }

    // Return data starting from header row
    // Header row becomes row 0, data rows start from row 1
    return data.slice(headerIndex);
  }

  /**
   * renderUnifiedTable(sheetDiff, keyColumn)
   * Renders a unified comparison table with old/new indices
   * Shows all rows and columns with change highlighting
   *
   * @param {Object} sheetDiff - Single sheet comparison result
   * @param {string} keyColumn - Column letter used as key
   */
  renderUnifiedTable(sheetDiff, keyColumn = 'A') {
    const container = document.getElementById('unifiedTableContainer');
    if (!container) {
      console.error('unifiedTableContainer element not found');
      return;
    }

    // Clear existing content
    container.innerHTML = '';

    // Create table element
    const table = document.createElement('table');
    table.className = 'unified-table diff-table';

    // Build table header (two-row header with column letters and content)
    const thead = this.buildUnifiedHeader(sheetDiff);
    table.appendChild(thead);

    // Build table body (data rows) - Pass keyColumn
    const tbody = this.buildUnifiedBody(sheetDiff, keyColumn);
    table.appendChild(tbody);

    // Wrap table in scrollable container
    const wrapper = document.createElement('div');
    wrapper.className = 'table-wrapper';
    wrapper.appendChild(table);

    container.appendChild(wrapper);

    // Collect all changed cells for navigation
    this.collectChangedCells();

    // Set up click-to-navigate on cells
    this.setupCellClickNavigation();
  }

  /**
   * setupCellClickNavigation()
   * Enables clicking on changed cells to navigate to them
   * Updates navigation state and scrolls to clicked change
   */
  setupCellClickNavigation() {
    const table = document.querySelector('#unifiedTableContainer .diff-table');
    if (!table) return;

    const tbody = table.querySelector('tbody');
    if (!tbody) return;

    // Add click listener to tbody
    tbody.addEventListener('click', (e) => {
      // Check if clicked element is a cell
      const clickedCell = e.target.closest('td');
      if (!clickedCell) return;

      // Check if it's a changed cell
      const isChangedCell = clickedCell.classList.contains('cell-modified') || clickedCell.classList.contains('cell-added') || clickedCell.classList.contains('cell-deleted');

      if (!isChangedCell) {
        // console.log('⚠️ Clicked cell is not a changed cell');
        return;
      }

      // Find this cell in changedCells array
      const clickedRow = clickedCell.closest('tr');

      for (let i = 0; i < this.changedCells.length; i++) {
        const { row, cell } = this.changedCells[i];

        if (row === clickedRow && cell === clickedCell) {
          // console.log(`✅ Clicked on change #${i + 1}`);
          this.currentChangeIndex = i;
          this.updateNavigationUI();
          this.scrollToChange();
          return;
        }
      }

      console.log('⚠️ Could not find corresponding change');
    });
  }

  /**
   * getUnifiedColumns(sheetDiff)
   * Creates a unified column list by merging columns from both files
   * Matches columns by header content, not position
   * Handles column reordering, additions, and deletions
   *
   * @param {Object} sheetDiff - Sheet comparison result
   * @returns {Array<Object>} Array of unified column objects
   *
   * Each column object:
   * {
   *   header: string,     // Header content or '(Blank Column)'
   *   oldCol: string,     // Column letter in File A (null if added)
   *   newCol: string,     // Column letter in File B (null if deleted)
   *   type: string        // 'normal' | 'added' | 'deleted'
   * }
   */
  getUnifiedColumns(sheetDiff) {
    const oldHeaders = sheetDiff.oldData[0] || {};
    const newHeaders = sheetDiff.newData[0] || {};

    const headerMap = new Map(); // normalized header → { originalA, originalB, oldCol, newCol }

    // ✅ Helper: Normalize header for comparison
    const normalizeHeader = (header) => {
      return String(header || '')
        .trim()
        .toLowerCase();
    };

    // Step 1: Process headers from File A
    Object.keys(oldHeaders).forEach((col) => {
      const content = String(oldHeaders[col] || '').trim();

      if (content) {
        const normalized = normalizeHeader(content); // ✅ Normalize!

        headerMap.set(normalized, {
          originalA: content, // Keep original for display
          originalB: null,
          oldCol: col,
          newCol: null,
        });
      } else {
        // Empty header - use column letter as key
        headerMap.set(`__empty_old_${col}`, {
          originalA: '',
          originalB: null,
          oldCol: col,
          newCol: null,
        });
      }
    });

    // Step 2: Process headers from File B
    Object.keys(newHeaders).forEach((col) => {
      const content = String(newHeaders[col] || '').trim();

      if (content) {
        const normalized = normalizeHeader(content); // ✅ Normalize!

        if (headerMap.has(normalized)) {
          // ✅ Found matching header (case-insensitive)!
          const existing = headerMap.get(normalized);
          existing.originalB = content; // Keep original for display
          existing.newCol = col;
        } else {
          // Header unique to File B
          headerMap.set(normalized, {
            originalA: null,
            originalB: content,
            oldCol: null,
            newCol: col,
          });
        }
      } else {
        // Empty column in File B
        const key = `__empty_new_${col}`;

        // Check if File A also has empty column at same position
        const oldEmptyKey = `__empty_old_${col}`;
        if (headerMap.has(oldEmptyKey)) {
          headerMap.get(oldEmptyKey).newCol = col;
          // Rename key to indicate both files have it
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

    // Step 3: Convert to array (ordered by File B column order)
    const result = [];
    const processedHeaders = new Set();

    // First, add columns in File B order
    Object.keys(newHeaders).forEach((newCol) => {
      for (let [normalized, mapping] of headerMap) {
        if (mapping.newCol === newCol && !processedHeaders.has(normalized)) {
          processedHeaders.add(normalized);

          // ✅ Use originalB for display (prefer File B's casing)
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

    // Then, add File A exclusive columns (deleted columns)
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

    // console.log('📋 Unified Columns (case-insensitive):', result);
    return result;
  }

  /**
   * buildUnifiedHeader(sheetDiff)
   * Builds a two-row table header
   * Row 1: Column letters (A, +B, C, −D, etc.) with +/− indicating added/deleted
   * Row 2: Header content
   *
   * @param {Object} sheetDiff - Sheet comparison result
   * @returns {HTMLElement} thead element with two rows
   */
  buildUnifiedHeader(sheetDiff) {
    const thead = document.createElement('thead');
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);

    // Row 1: Column letters
    const tr1 = document.createElement('tr');

    // Old Index column header
    const th1Old = document.createElement('th');
    th1Old.className = 'index-col';
    th1Old.textContent = 'Old';
    th1Old.rowSpan = 2; // Spans both header rows
    tr1.appendChild(th1Old);

    // New Index column header
    const th1New = document.createElement('th');
    th1New.className = 'index-col';
    th1New.textContent = 'New';
    th1New.rowSpan = 2; // Spans both header rows
    tr1.appendChild(th1New);

    // Data column headers (A, +B, C, −D, etc.)
    unifiedColumns.forEach((col) => {
      const th = document.createElement('th');

      let colLabel = col.newCol || col.oldCol; // Prefer newCol

      if (col.type === 'added') {
        th.className = 'col-added';
        colLabel = `+${col.newCol}`; // Prefix with +
      } else if (col.type === 'deleted') {
        th.className = 'col-deleted';
        colLabel = `−${col.oldCol}`; // Prefix with −
      }

      th.textContent = colLabel;
      tr1.appendChild(th);
    });

    // Row 2: Header content
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
   * Builds the table body with all rows from both files
   * Highlights changed, added, and deleted cells/rows
   *
   * @param {Object} sheetDiff - Sheet comparison result
   * @param {string} keyColumn - Column letter used as key
   * @returns {HTMLElement} tbody element with data rows
   *
   * Cell highlighting:
   * - cell-modified: Cell value changed
   * - cell-added: Cell in added column or added row
   * - cell-deleted: Cell in deleted column or deleted row
   * - cell-unchanged: No changes
   *
   * Row highlighting:
   * - row-added: Row exists only in File B
   * - row-deleted: Row exists only in File A
   */
  buildUnifiedBody(sheetDiff, keyColumn = 'A') {
    const tbody = document.createElement('tbody');
    const allRows = this.getAllRows(sheetDiff, keyColumn); // Pass keyColumn
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);
    const cellChanges = this.buildCellChangeMap(sheetDiff.differences);
    const rowChanges = this.buildRowChangeMap(sheetDiff.rowChanges);

    allRows.forEach((rowInfo) => {
      const tr = document.createElement('tr');

      const rowChange = rowChanges.get(rowInfo.key);

      // Apply row-level highlighting
      if (rowChange?.type === 'added') {
        tr.className = 'row-added';
      } else if (rowChange?.type === 'deleted') {
        tr.className = 'row-deleted';
      }

      // Old Index cell
      const tdOldIdx = document.createElement('td');
      tdOldIdx.className = 'index-cell old-idx';
      tdOldIdx.textContent = rowInfo.oldIndex !== null ? rowInfo.oldIndex : '-';
      tr.appendChild(tdOldIdx);

      // New Index cell
      const tdNewIdx = document.createElement('td');
      tdNewIdx.className = 'index-cell new-idx';
      tdNewIdx.textContent = rowInfo.newIndex !== null ? rowInfo.newIndex : '-';
      tr.appendChild(tdNewIdx);

      // Data cells
      unifiedColumns.forEach((col) => {
        const td = document.createElement('td');

        const oldValue = col.oldCol ? rowInfo.oldRow?.[col.oldCol] : null;
        const newValue = col.newCol ? rowInfo.newRow?.[col.newCol] : null;

        // Build cell key using header content (not column letter)
        const cellKey = `${rowInfo.oldIndex || rowInfo.newIndex}-${col.header}`;
        const cellDiff = cellChanges.get(cellKey);

        if (cellDiff) {
          // Cell value was modified
          td.className = 'cell-modified';
          td.innerHTML = `
          <div class="cell-value-change">
            <span class="old-value">${this.formatValue(cellDiff.oldValue)}</span>
            <span class="value-separator">→</span>
            <span class="new-value">${this.formatValue(cellDiff.newValue)}</span>
          </div>
        `;
        } else if (col.type === 'added') {
          // Cell in added column
          td.className = 'cell-added';
          td.innerHTML = this.formatValue(newValue);
        } else if (col.type === 'deleted') {
          // Cell in deleted column
          td.className = 'cell-deleted';
          td.innerHTML = this.formatValue(oldValue);
        } else if (rowChange?.type === 'deleted') {
          // Cell in deleted row
          td.className = 'cell-deleted';
          td.innerHTML = this.formatValue(oldValue);
        } else if (rowChange?.type === 'added') {
          // Cell in added row
          td.className = 'cell-added';
          td.innerHTML = this.formatValue(newValue);
        } else {
          // Unchanged cell
          td.className = 'cell-unchanged';
          td.innerHTML = this.formatValue(newValue || oldValue);
        }

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    return tbody;
  }

  /**
   * getAllRows(sheetDiff, keyColumn)
   * Merges all rows from both files into a unified list
   * Uses specified key column for row matching
   * Sorted by row number (ascending), with deleted rows before added rows at same position
   *
   * @param {Object} sheetDiff - Sheet comparison result
   * @param {string} keyColumn - Column letter used as key (e.g. 'A', 'B', 'C')
   * @returns {Array<Object>} Sorted array of unified row objects
   */
  getAllRows(sheetDiff, keyColumn = 'A') {
    const rowMap = new Map();

    const headerRowA = sheetDiff.headerRowA || 1;
    const headerRowB = sheetDiff.headerRowB || 1;

    const oldDataRows = sheetDiff.oldData.slice(1);
    const newDataRows = sheetDiff.newData.slice(1);

    // Process Old File
    oldDataRows.forEach((row, index) => {
      const keyValue = String(row[keyColumn] || '').trim();
      const excelRowNumber = headerRowA + 1 + index;

      if (keyValue) {
        rowMap.set(keyValue, {
          key: keyValue,
          oldRow: row,
          oldIndex: excelRowNumber,
          newRow: null,
          newIndex: null,
        });
      } else {
        const uniqueKey = `__blank_old_${excelRowNumber}`;
        rowMap.set(uniqueKey, {
          key: `(empty)`,
          oldRow: row,
          oldIndex: excelRowNumber,
          newRow: null,
          newIndex: null,
          isBlankKey: true,
        });
      }
    });

    // Process New File
    newDataRows.forEach((row, index) => {
      const keyValue = String(row[keyColumn] || '').trim();
      const excelRowNumber = headerRowB + 1 + index;

      if (keyValue) {
        if (rowMap.has(keyValue)) {
          const existing = rowMap.get(keyValue);
          existing.newRow = row;
          existing.newIndex = excelRowNumber;
        } else {
          rowMap.set(keyValue, {
            key: keyValue,
            oldRow: null,
            oldIndex: null,
            newRow: row,
            newIndex: excelRowNumber,
          });
        }
      } else {
        const positionKey = `__blank_old_${excelRowNumber}`;

        if (rowMap.has(positionKey)) {
          const existing = rowMap.get(positionKey);
          const allColumnsMatch = this.areRowsIdentical(existing.oldRow, row);

          if (allColumnsMatch) {
            existing.newRow = row;
            existing.newIndex = excelRowNumber;
          } else {
            const newKey = `__blank_new_${excelRowNumber}`;
            rowMap.set(newKey, {
              key: `(empty)`,
              oldRow: null,
              oldIndex: null,
              newRow: row,
              newIndex: excelRowNumber,
              isBlankKey: true,
            });
          }
        } else {
          const newKey = `__blank_new_${excelRowNumber}`;
          rowMap.set(newKey, {
            key: `(empty)`,
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

    // ✅ IMPROVED SORTING LOGIC - Prioritize New Index
    rows.sort((a, b) => {
      // ✅ PRIMARY SORT: By New Index (File B order)
      // This ensures new rows appear in their correct position
      const aNew = a.newIndex !== null ? a.newIndex : Infinity;
      const bNew = b.newIndex !== null ? b.newIndex : Infinity;

      if (aNew !== bNew) {
        return aNew - bNew;
      }

      // ✅ SECONDARY SORT: By Old Index (if both don't have New Index)
      // This handles deleted rows
      const aOld = a.oldIndex !== null ? a.oldIndex : Infinity;
      const bOld = b.oldIndex !== null ? b.oldIndex : Infinity;

      if (aOld !== bOld) {
        return aOld - bOld;
      }

      // ✅ TERTIARY SORT: By row type (Matched < Deleted < Added)
      const getRowType = (row) => {
        if (row.oldIndex !== null && row.newIndex !== null) return 0; // Matched
        if (row.oldIndex !== null && row.newIndex === null) return 1; // Deleted
        if (row.oldIndex === null && row.newIndex !== null) return 2; // Added
        return 3;
      };

      const aType = getRowType(a);
      const bType = getRowType(b);

      return aType - bType;
    });

    return rows;
  }

  /**
   * Check if two rows are identical across all columns
   */
  areRowsIdentical(row1, row2) {
    if (!row1 || !row2) return false;

    // Get all unique column keys
    const allKeys = new Set([...Object.keys(row1), ...Object.keys(row2)]);

    // Compare each column
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
   * Normalize cell value for comparison
   */
  normalizeValue(cell) {
    if (cell === null || cell === undefined) {
      return '';
    }

    let value = String(cell);
    value = value.replace(/[\u200B-\u200D\uFEFF]/g, ''); // Remove invisible chars
    value = value.trim();

    return value;
  }

  /**
   * buildRowChangeMap(rowChanges)
   * Creates a Map for quick row change lookup
   *
   * @param {Array<Object>} rowChanges - Array of row change objects
   * @returns {Map} Map of rowKey → change object
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
   * Creates a Map for quick cell difference lookup
   * Uses "rowIndex-headerContent" as key
   *
   * @param {Array<Object>} differences - Array of cell difference objects
   * @returns {Map} Map of cellKey → difference object
   */
  buildCellChangeMap(differences) {
    const map = new Map();
    differences.forEach((diff) => {
      // Use "rowIndex-headerContent" as key (not column letter)
      const key = `${diff.row}-${diff.header}`;
      map.set(key, diff);
    });
    return map;
  }

  /**
   * formatValue(value)
   * Formats cell value for display
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
   * Sets up navigation controls for stepping through changes
   * Binds Previous/Next buttons and keyboard shortcuts
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

    // Set up keyboard shortcuts (P = Previous, N = Next)
    this.setupKeyboardShortcuts();
  }

  /**
   * setupKeyboardShortcuts()
   * Enables keyboard navigation through changes
   * P key = Previous change
   * N key = Next change
   */
  setupKeyboardShortcuts() {
    // Remove old listener to avoid duplicates
    if (this.keyboardHandler) {
      document.removeEventListener('keydown', this.keyboardHandler);
    }

    // Create new listener
    this.keyboardHandler = (e) => {
      // Don't trigger shortcuts if user is typing in input/textarea
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

    // Bind to document
    document.addEventListener('keydown', this.keyboardHandler);

    // console.log('⌨️ Keyboard shortcuts enabled: P = Previous, N = Next');
  }

  /**
   * collectChangedCells()
   * Collects all changed cells in the table for navigation
   * Stores references to changed cell elements
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

    // Find all changed cells
    const rows = tbody.querySelectorAll('tr');
    rows.forEach((row, rowIndex) => {
      const cells = row.querySelectorAll('td.cell-modified, td.cell-added, td.cell-deleted');
      // const cells = row.querySelectorAll('td.cell-modified');
      cells.forEach((cell) => {
        this.changedCells.push({ row, cell });
      });
    });

    // console.log(`📍 Collected ${this.changedCells.length} changes`);
    this.updateNavigationUI();
  }

  /**
   * navigateToChange(direction)
   * Navigates to the previous or next change
   * Wraps around at start/end of change list
   *
   * @param {string} direction - 'next' or 'prev'
   */
  navigateToChange(direction) {
    if (this.changedCells.length === 0) return;

    if (direction === 'next') {
      // Move to next change (wrap around to start)
      this.currentChangeIndex = (this.currentChangeIndex + 1) % this.changedCells.length;
    } else if (direction === 'prev') {
      // Move to previous change (wrap around to end)
      this.currentChangeIndex = (this.currentChangeIndex - 1 + this.changedCells.length) % this.changedCells.length;
    }

    this.updateNavigationUI();
    this.scrollToChange();
  }

  /**
   * scrollToChange()
   * Scrolls to the currently selected change and highlights it
   * Auto-removes highlight after 2 seconds
   */
  scrollToChange() {
    if (this.currentChangeIndex < 0 || this.currentChangeIndex >= this.changedCells.length) return;

    const { cell } = this.changedCells[this.currentChangeIndex];

    // Remove previous highlights
    document.querySelectorAll('.cell-highlighted').forEach((c) => {
      c.classList.remove('cell-highlighted');
    });

    // Add highlight to current cell
    cell.classList.add('cell-highlighted');

    // Scroll cell into view (smooth animation, centered)
    cell.scrollIntoView({ behavior: 'smooth', block: 'center' });

    // Remove highlight after 2 seconds
    setTimeout(() => {
      cell.classList.remove('cell-highlighted');
    }, 2000);
  }

  /**
   * updateNavigationUI()
   * Updates the change counter and navigation button states
   * Shows current change number and total changes (e.g., "5 / 23")
   */
  updateNavigationUI() {
    const counter = document.getElementById('changeCounter');
    const prevBtn = document.getElementById('prevChangeBtn');
    const nextBtn = document.getElementById('nextChangeBtn');

    if (!counter) return;

    const total = this.changedCells.length;
    const current = this.currentChangeIndex >= 0 ? this.currentChangeIndex + 1 : 0;

    // Update counter text (e.g., "5 / 23")
    counter.textContent = `${current} / ${total}`;

    // Disable buttons if no changes
    if (prevBtn && nextBtn) {
      prevBtn.disabled = total === 0;
      nextBtn.disabled = total === 0;
    }
  }

  /**
   * compareByRowPosition(sheetA, sheetB, headerRowA, headerRowB)
   * Compares sheets by matching rows at the same position
   * This approach matches rows by their row number rather than by key column
   *
   * @param {Array<Object>} sheetA - Sheet A data
   * @param {Array<Object>} sheetB - Sheet B data
   * @param {number} headerRowA - Header row index for sheet A (1-based)
   * @param {number} headerRowB - Header row index for sheet B (1-based)
   */
  compareByRowPosition(sheetA, sheetB, headerRowA, headerRowB) {
    console.log('🔢 Comparing by row position...');

    // Get headers
    const headersA = sheetA[0] || {};
    const headersB = sheetB[0] || {};

    // Get unified columns (case-insensitive matching)
    const unifiedColumns = this.getUnifiedColumns({
      oldData: sheetA,
      newData: sheetB,
    });

    // console.log('📋 Unified Columns:', unifiedColumns);

    // Clear previous results
    this.changedCells = [];

    // Get data rows (skip header row)
    const dataRowsA = sheetA.slice(1);
    const dataRowsB = sheetB.slice(1);

    const maxRows = Math.max(dataRowsA.length, dataRowsB.length);

    // console.log(`📊 Comparing ${maxRows} rows (position-based)...`);

    const differences = [];
    const rowChanges = [];

    // Compare rows at same position
    for (let i = 0; i < maxRows; i++) {
      const rowA = dataRowsA[i] || {};
      const rowB = dataRowsB[i] || {};

      const excelRowA = headerRowA + 1 + i; // Excel row number (1-based)
      const excelRowB = headerRowB + 1 + i; // Excel row number (1-based)

      // Check if row exists in both files
      const rowExistsInA = i < dataRowsA.length;
      const rowExistsInB = i < dataRowsB.length;

      if (!rowExistsInA) {
        // Row added in File B
        rowChanges.push({
          rowKey: `row-${i}`,
          type: 'added',
          newRowIndex: excelRowB,
          row: rowB,
        });
      } else if (!rowExistsInB) {
        // Row deleted from File A
        rowChanges.push({
          rowKey: `row-${i}`,
          type: 'deleted',
          oldRowIndex: excelRowA,
          row: rowA,
        });
      } else {
        // Compare all columns
        unifiedColumns.forEach(({ header, oldCol, newCol }) => {
          const valueA = oldCol ? (rowA[oldCol] ?? '') : '';
          const valueB = newCol ? (rowB[newCol] ?? '') : '';

          // Normalize values for comparison
          const normalizedA = String(valueA).trim();
          const normalizedB = String(valueB).trim();

          if (normalizedA !== normalizedB) {
            // Cell has changed
            differences.push({
              row: excelRowB, // Use File B row number
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

    // console.log(`🔍 Found ${differences.length} cell differences (position-based)`);
    // console.log(`🔍 Found ${rowChanges.length} row changes`);

    // Build sheet diff object
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

    // Render the unified table for position-based comparison
    this.renderUnifiedTableByPosition(singleSheetDiff, headerRowA, headerRowB);
  }

  /**
   * renderUnifiedTableByPosition(sheetDiff, headerRowA, headerRowB)
   * Renders a table for position-based comparison of sheets
   * Displays rows matched by position rather than by key column
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

    // Build header
    const thead = this.buildUnifiedHeader(sheetDiff);
    table.appendChild(thead);

    // ✅ Build body with position-based rows
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
   * Builds the table body for position-based comparison
   * Matches rows by their position in each sheet rather than by a key column
   */
  buildUnifiedBodyByPosition(sheetDiff, headerRowA, headerRowB) {
    const tbody = document.createElement('tbody');
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);
    const cellChanges = this.buildCellChangeMap(sheetDiff.differences);
    const rowChanges = this.buildRowChangeMap(sheetDiff.rowChanges);

    const dataRowsA = sheetDiff.oldData.slice(1);
    const dataRowsB = sheetDiff.newData.slice(1);
    const maxRows = Math.max(dataRowsA.length, dataRowsB.length);

    // ✅ Build rows by position
    for (let i = 0; i < maxRows; i++) {
      const tr = document.createElement('tr');

      const rowA = dataRowsA[i];
      const rowB = dataRowsB[i];

      const excelRowA = headerRowA + 1 + i;
      const excelRowB = headerRowB + 1 + i;

      const rowExistsInA = i < dataRowsA.length;
      const rowExistsInB = i < dataRowsB.length;

      // Check if row was added/deleted
      const rowChange = rowChanges.get(`position-${i}`);

      if (rowChange?.type === 'added') {
        tr.className = 'row-added';
      } else if (rowChange?.type === 'deleted') {
        tr.className = 'row-deleted';
      }

      // Old Index cell
      const tdOldIdx = document.createElement('td');
      tdOldIdx.className = 'index-cell old-idx';
      tdOldIdx.textContent = rowExistsInA ? excelRowA : '-';
      tr.appendChild(tdOldIdx);

      // New Index cell
      const tdNewIdx = document.createElement('td');
      tdNewIdx.className = 'index-cell new-idx';
      tdNewIdx.textContent = rowExistsInB ? excelRowB : '-';
      tr.appendChild(tdNewIdx);

      // Data cells
      unifiedColumns.forEach((col) => {
        const td = document.createElement('td');

        const oldValue = col.oldCol && rowA ? (rowA[col.oldCol] ?? null) : null;
        const newValue = col.newCol && rowB ? (rowB[col.newCol] ?? null) : null;

        // Build cell key
        const cellKey = `${excelRowB}-${col.header}`;
        const cellDiff = cellChanges.get(cellKey);

        if (cellDiff) {
          // Cell value was modified
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

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    }

    return tbody;
  }
}

// Export for use in other modules
export default DiffViewer;
