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
    this.dataA = null;              // Parsed data from File A
    this.dataB = null;              // Parsed data from File B
    this.diffResults = null;        // Complete diff results
    this.changedCells = [];         // Array of changed cell elements for navigation
    this.currentChangeIndex = -1;   // Index of currently highlighted change
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

    // Populate sheet selection dropdowns
    this.populateSheetDropdowns();

    // Try to find matching sheet names between files
    const matchedSheet = this.findMatchingSheet();

    if (matchedSheet) {
      // Auto-select matching sheets
      document.getElementById('sheetSelectA').value = matchedSheet.sheetA;
      document.getElementById('sheetSelectB').value = matchedSheet.sheetB;
      
      // Auto-compare matched sheets
      this.compareSelectedSheets();
    } else {
      console.log('⚠️ No matching sheets found, waiting for user to manually compare');
    }

    // Set up change navigation controls
    this.setupChangeNavigation();
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

    console.log('✅ Sheet dropdowns populated');
  }

  /**
   * compareSelectedSheets()
   * Compares the currently selected sheets from both dropdowns
   * Triggered when user clicks "Compare" button
   */
  compareSelectedSheets() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');

    if (!sheetSelectA || !sheetSelectB) {
      console.error('Sheet dropdown elements not found');
      return;
    }

    const selectedSheetA = sheetSelectA.value;
    const selectedSheetB = sheetSelectB.value;

    // Validate selections
    if (!selectedSheetA || !selectedSheetB) {
      alert('Please select sheets to compare');
      return;
    }

    console.log(`Comparing sheets: ${selectedSheetA} vs ${selectedSheetB}`);

    // Get sheet data
    const sheetA = this.dataA.sheets[selectedSheetA]?.data || [];
    const sheetB = this.dataB.sheets[selectedSheetB]?.data || [];

    // Validate sheet data
    if (sheetA.length === 0 || sheetB.length === 0) {
      alert('Selected sheet is empty');
      return;
    }

    // Perform comparison
    const diffEngine = new DiffEngine();
    const singleSheetDiff = diffEngine.compareSheets(sheetA, sheetB);
    singleSheetDiff.sheetName = `${selectedSheetA} vs ${selectedSheetB}`;

    // Render results
    this.renderUnifiedTable(singleSheetDiff);
  }

  /**
   * renderUnifiedTable(sheetDiff)
   * Renders a unified comparison table with old/new indices
   * Shows all rows and columns with change highlighting
   * 
   * @param {Object} sheetDiff - Single sheet comparison result
   */
  renderUnifiedTable(sheetDiff) {
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

    // Build table body (data rows)
    const tbody = this.buildUnifiedBody(sheetDiff);
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
      const isChangedCell = clickedCell.classList.contains('cell-modified') || 
                           clickedCell.classList.contains('cell-added') || 
                           clickedCell.classList.contains('cell-deleted');

      if (!isChangedCell) {
        console.log('⚠️ Clicked cell is not a changed cell');
        return;
      }

      // Find this cell in changedCells array
      const clickedRow = clickedCell.closest('tr');

      for (let i = 0; i < this.changedCells.length; i++) {
        const { row, cell } = this.changedCells[i];

        if (row === clickedRow && cell === clickedCell) {
          console.log(`✅ Clicked on change #${i + 1}`);
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

    const headerMap = new Map(); // header content → { oldCol, newCol }

    // Step 1: Process headers from File A
    Object.keys(oldHeaders).forEach((col) => {
      const content = String(oldHeaders[col] || '').trim();
      if (content) {
        headerMap.set(content, { oldCol: col, newCol: null });
      } else {
        // Empty header - use column letter as key
        headerMap.set(`__empty_old_${col}`, { oldCol: col, newCol: null });
      }
    });

    // Step 2: Process headers from File B
    Object.keys(newHeaders).forEach((col) => {
      const content = String(newHeaders[col] || '').trim();

      if (content) {
        if (headerMap.has(content)) {
          // Found matching header
          headerMap.get(content).newCol = col;
        } else {
          // Header unique to File B
          headerMap.set(content, { oldCol: null, newCol: col });
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
          headerMap.set(key, { oldCol: null, newCol: col });
        }
      }
    });

    // Step 3: Convert to array (ordered by File B column order)
    const result = [];
    const processedHeaders = new Set();

    // First, add columns in File B order
    Object.keys(newHeaders).forEach((newCol) => {
      for (let [header, mapping] of headerMap) {
        if (mapping.newCol === newCol && !processedHeaders.has(header)) {
          processedHeaders.add(header);
          result.push({
            header: header.startsWith('__empty_') ? '(Blank Column)' : header,
            oldCol: mapping.oldCol,
            newCol: mapping.newCol,
            type: mapping.oldCol && mapping.newCol ? 'normal' : mapping.oldCol ? 'deleted' : 'added',
          });
          break;
        }
      }
    });

    // Then, add File A exclusive columns (deleted columns)
    for (let [header, mapping] of headerMap) {
      if (!processedHeaders.has(header)) {
        result.push({
          header: header.startsWith('__empty_') ? '(Blank Column)' : header,
          oldCol: mapping.oldCol,
          newCol: mapping.newCol,
          type: 'deleted',
        });
      }
    }

    console.log('📋 Unified Columns:', result);
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
   * buildUnifiedBody(sheetDiff)
   * Builds the table body with all rows from both files
   * Highlights changed, added, and deleted cells/rows
   * 
   * @param {Object} sheetDiff - Sheet comparison result
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
  buildUnifiedBody(sheetDiff) {
    const tbody = document.createElement('tbody');
    const allRows = this.getAllRows(sheetDiff);
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
   * getAllRows(sheetDiff)
   * Merges all rows from both files into a unified list
   * Uses column A as the row key for matching
   * 
   * @param {Object} sheetDiff - Sheet comparison result
   * @returns {Array<Object>} Array of unified row objects
   * 
   * Each row object:
   * {
   *   key: string,          // Row identifier (column A value or fallback)
   *   oldRow: Object,       // Row data from File A (null if added)
   *   oldIndex: number,     // Row number in File A (null if added)
   *   newRow: Object,       // Row data from File B (null if deleted)
   *   newIndex: number      // Row number in File B (null if deleted)
   * }
   */
  getAllRows(sheetDiff) {
    const rowMap = new Map();

    // Process rows from File A (skip header row)
    sheetDiff.oldData.slice(1).forEach((row, index) => {
      const key = String(row.A || '').trim() || `old-${index}`;
      rowMap.set(key, {
        key: key,
        oldRow: row,
        oldIndex: index + 2, // +2: +1 for 1-based, +1 for skipping header
        newRow: null,
        newIndex: null,
      });
    });

    // Process rows from File B (skip header row)
    sheetDiff.newData.slice(1).forEach((row, index) => {
      const key = String(row.A || '').trim() || `new-${index}`;

      if (rowMap.has(key)) {
        // Row exists in both files - update existing entry
        const existing = rowMap.get(key);
        existing.newRow = row;
        existing.newIndex = index + 2;
      } else {
        // Row only in File B - create new entry
        rowMap.set(key, {
          key: key,
          oldRow: null,
          oldIndex: null,
          newRow: row,
          newIndex: index + 2,
        });
      }
    });

    return Array.from(rowMap.values());
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

    console.log('⌨️ Keyboard shortcuts enabled: P = Previous, N = Next');
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

    console.log(`📍 Collected ${this.changedCells.length} changes`);
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
}

// Export for use in other modules
export default DiffViewer;