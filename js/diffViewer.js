// diffViewer.js - Diff Viewer Module (FIXED)

const DiffViewer = {
  // Core data
  currentDiffResult: null,
  currentFileA: null,
  currentFileB: null,
  currentSheet: null,
  currentStatus: null,

  // Scroll synchronization state
  isSyncingVertical: false,
  syncEnabled: true,
  wrapperA: null,
  wrapperB: null,
  HEADER_HEIGHT: 38,
  verticalSyncTimeoutId: null,

  // Change navigation state
  navigationTimeoutId: null,
  changedCells: [],
  currentChangeIndex: -1,
  prevBtn: null,
  nextBtn: null,
  counterDisplay: null,

  // Tooltip state
  tooltipElement: null,
  currentTooltipCell: null,

  /**
   * Initialize the sync scroll toggle button
   * Sets up click event listener for the sync button
   */
  initSyncButton() {
    const syncBtn = document.getElementById('syncToggle');
    if (!syncBtn) return;

    syncBtn.addEventListener('click', () => {
      this.toggleSync();
    });
  },

  /**
   * Initialize the change navigation system
   * Sets up Previous/Next buttons and keyboard shortcuts (P/N)
   */
  initChangeNavigation() {
    this.prevBtn = document.getElementById('prevChangeBtn');
    this.nextBtn = document.getElementById('nextChangeBtn');
    this.counterDisplay = document.getElementById('changeCounter');

    if (!this.prevBtn || !this.nextBtn || !this.counterDisplay) {
      console.warn('❌ Change navigation elements not found');
      return;
    }

    // Bind click events to navigation buttons
    this.prevBtn.addEventListener('click', () => {
      this.navigateToChange('prev');
    });

    this.nextBtn.addEventListener('click', () => {
      this.navigateToChange('next');
    });

    // Set up keyboard shortcuts for navigation
    // P = Previous, N = Next (only when diff section is visible)
    document.addEventListener('keydown', (e) => {
      if (document.getElementById('diffSection').style.display !== 'block') return;

      if (e.key === 'n' || e.key === 'N') {
        e.preventDefault();
        this.navigateToChange('next');
      } else if (e.key === 'p' || e.key === 'P') {
        e.preventDefault();
        this.navigateToChange('prev');
      }
    });

    console.log('✅ Change navigation initialized');
  },

  /**
   * Collect all changed cells from both comparison tables
   * Builds an ordered array of cell references for navigation
   * Each cell stores its position and references to both table A and B cells
   */
  collectChangedCells() {
    this.changedCells = [];
    this.currentChangeIndex = -1;

    const tableA = document.getElementById('tableA');
    const tableB = document.getElementById('tableB');

    if (!tableA || !tableB) {
      this.updateNavigationUI();
      return;
    }

    // Use a Set to prevent duplicate entries for the same cell position
    const cellSet = new Set();

    /**
     * Helper function to collect changed cells from a single table
     * @param {HTMLElement} table - The table element to scan
     * @param {string} side - Either 'A' or 'B' to identify which table
     */
    const collectFromTable = (table, side) => {
      // Find all cells with change classes (modified, added, or removed)
      const cells = table.querySelectorAll('td.cell-modified, td.cell-added, td.cell-removed');

      cells.forEach((cell) => {
        const row = cell.parentElement;
        const rowIndex = Array.from(row.parentElement.children).indexOf(row);
        const colIndex = Array.from(row.children).indexOf(cell) - 1; // -1 to exclude row header

        if (colIndex >= 0) {
          const key = `${rowIndex}-${colIndex}`;

          // Only add if we haven't seen this position before
          if (!cellSet.has(key)) {
            cellSet.add(key);
            this.changedCells.push({
              row: rowIndex,
              col: colIndex,
              cellA: side === 'A' ? cell : this.getCellAt(tableA, rowIndex, colIndex),
              cellB: side === 'B' ? cell : this.getCellAt(tableB, rowIndex, colIndex),
              key: key,
            });
          }
        }
      });
    };

    // Collect from both tables
    collectFromTable(tableA, 'A');
    collectFromTable(tableB, 'B');

    // Sort changes by position: first by row, then by column
    this.changedCells.sort((a, b) => {
      if (a.row !== b.row) return a.row - b.row;
      return a.col - b.col;
    });

    console.log(`📍 Collected ${this.changedCells.length} changed cells`);
    this.updateNavigationUI();
  },

  /**
   * Get a specific cell element from a table by row and column index
   * @param {HTMLElement} table - The table element to search
   * @param {number} rowIndex - Zero-based row index
   * @param {number} colIndex - Zero-based column index (excluding row header)
   * @returns {HTMLElement|null} The cell element, or null if not found
   */
  getCellAt(table, rowIndex, colIndex) {
    const tbody = table.querySelector('tbody');
    if (!tbody) return null;

    const rows = tbody.querySelectorAll('tr');
    if (rowIndex >= rows.length) return null;

    const row = rows[rowIndex];
    const cells = row.querySelectorAll('td');

    // +1 because the first cell is the row header
    if (colIndex + 1 >= cells.length) return null;

    return cells[colIndex + 1];
  },

  /**
   * Initialize the custom tooltip system
   * Sets up the tooltip element and mouse tracking for positioning
   */
  initTooltip() {
    this.tooltipElement = document.getElementById('customTooltip');
    if (!this.tooltipElement) {
      console.warn('❌ Tooltip element not found');
    } else {
      console.log('✅ Tooltip element found!');

      // Track mouse movement to update tooltip position in real-time
      document.addEventListener('mousemove', (e) => {
        if (this.currentTooltipCell && this.tooltipElement.classList.contains('visible')) {
          this.updateTooltipPosition(e.clientX, e.clientY);
        }
      });
    }
  },

  /**
   * Navigate to the next or previous change in the list
   * Uses debouncing to handle rapid clicking smoothly
   * @param {string} direction - Either 'next' or 'prev'
   */
  navigateToChange(direction) {
    if (this.changedCells.length === 0) {
      console.log('⚠️ No changes to navigate');
      return;
    }

    // Cancel any pending navigation to prevent scrolling conflicts
    if (this.navigationTimeoutId) {
      clearTimeout(this.navigationTimeoutId);
      this.navigationTimeoutId = null;
    }

    // Calculate the new index (wraps around at boundaries)
    if (direction === 'next') {
      this.currentChangeIndex = (this.currentChangeIndex + 1) % this.changedCells.length;
    } else if (direction === 'prev') {
      this.currentChangeIndex = (this.currentChangeIndex - 1 + this.changedCells.length) % this.changedCells.length;
    }

    // Update the counter immediately for responsive feedback
    this.updateNavigationUI();

    // Delay the scroll action slightly to prevent conflicts when clicking rapidly
    this.navigationTimeoutId = setTimeout(() => {
      const change = this.changedCells[this.currentChangeIndex];
      this.scrollToChange(change);
      this.navigationTimeoutId = null;
    }, 50);
  },

  /**
   * Scroll both tables to display a specific change
   * Centers the changed cells in the viewport and highlights them temporarily
   * @param {Object} change - Change object containing cellA and cellB references
   */
  scrollToChange(change) {
    if (!change) return;

    // Clear any existing highlights from previous navigation
    document.querySelectorAll('.cell-highlighted').forEach((cell) => {
      cell.classList.remove('cell-highlighted');
    });

    // Temporarily disable scroll synchronization to avoid interference
    this.syncEnabled = false;
    this.isSyncingVertical = true;

    // Clear any pending sync timeout
    if (this.verticalSyncTimeoutId) {
      clearTimeout(this.verticalSyncTimeoutId);
      this.verticalSyncTimeoutId = null;
    }

    // Scroll table A to center the changed cell
    if (change.cellA && this.wrapperA) {
      const targetTop = change.cellA.offsetTop - this.wrapperA.clientHeight / 2 + change.cellA.clientHeight / 2;
      const targetLeft = change.cellA.offsetLeft - this.wrapperA.clientWidth / 2 + change.cellA.clientWidth / 2;

      this.wrapperA.scrollTo({
        top: Math.max(0, targetTop),
        left: Math.max(0, targetLeft),
        behavior: 'auto', // Instant scroll for faster navigation
      });

      change.cellA.classList.add('cell-highlighted');
    }

    // Scroll table B to center the changed cell
    if (change.cellB && this.wrapperB) {
      const targetTop = change.cellB.offsetTop - this.wrapperB.clientHeight / 2 + change.cellB.clientHeight / 2;
      const targetLeft = change.cellB.offsetLeft - this.wrapperB.clientWidth / 2 + change.cellB.clientWidth / 2;

      this.wrapperB.scrollTo({
        top: Math.max(0, targetTop),
        left: Math.max(0, targetLeft),
        behavior: 'auto',
      });

      change.cellB.classList.add('cell-highlighted');
    }

    // Re-enable scroll synchronization after scrolling completes
    setTimeout(() => {
      this.isSyncingVertical = false;

      if (!this.syncEnabled) {
        this.syncEnabled = true;
        const syncBtn = document.getElementById('syncToggle');
        if (syncBtn) {
          syncBtn.classList.add('active');
          syncBtn.querySelector('.sync-text').textContent = 'Sync Scroll';
        }
        console.log('✅ Sync scroll force enabled after navigation');
      }
    }, 200);

    // Remove highlight effect after animation completes
    setTimeout(() => {
      if (change.cellA) change.cellA.classList.remove('cell-highlighted');
      if (change.cellB) change.cellB.classList.remove('cell-highlighted');
    }, 1500);

    console.log(`🎯 Navigated to change ${this.currentChangeIndex + 1}/${this.changedCells.length} at (${change.row}, ${change.col})`);
  },

  /**
   * Find the nearest change to a given cell position
   * Uses Manhattan distance (sum of row and column differences)
   * @param {number} row - Zero-based row index
   * @param {number} col - Zero-based column index
   * @returns {number} Index of the nearest change, or -1 if no changes exist
   */
  findNearestChangeIndex(row, col) {
    if (this.changedCells.length === 0) return -1;

    let nearestIndex = 0;
    let minDistance = Infinity;

    this.changedCells.forEach((change, index) => {
      // Calculate Manhattan distance: |row1 - row2| + |col1 - col2|
      const distance = Math.abs(change.row - row) + Math.abs(change.col - col);

      if (distance < minDistance) {
        minDistance = distance;
        nearestIndex = index;
      }
    });

    return nearestIndex;
  },

  /**
   * Update the navigation UI elements (counter and button states)
   * Displays current position like "5 / 23" and disables buttons when appropriate
   */
  updateNavigationUI() {
    if (!this.counterDisplay || !this.prevBtn || !this.nextBtn) return;

    const total = this.changedCells.length;
    const current = this.currentChangeIndex >= 0 ? this.currentChangeIndex + 1 : 0;

    this.counterDisplay.textContent = `${current} / ${total}`;

    // Disable navigation buttons if there are no changes
    if (total === 0) {
      this.prevBtn.disabled = true;
      this.nextBtn.disabled = true;
    } else {
      this.prevBtn.disabled = false;
      this.nextBtn.disabled = false;
    }
  },

  /**
   * Update tooltip position to follow the mouse cursor
   * Ensures the tooltip stays within the viewport boundaries
   * @param {number} x - Mouse X coordinate
   * @param {number} y - Mouse Y coordinate
   */
  updateTooltipPosition(x, y) {
    if (!this.tooltipElement) return;

    const tooltipRect = this.tooltipElement.getBoundingClientRect();
    const padding = 10;

    let left = x + padding;
    let top = y + padding;

    // Flip tooltip to the left if it would overflow the right edge
    if (left + tooltipRect.width > window.innerWidth) {
      left = x - tooltipRect.width - padding;
    }

    // Flip tooltip upward if it would overflow the bottom edge
    if (top + tooltipRect.height > window.innerHeight) {
      top = y - tooltipRect.height - padding;
    }

    this.tooltipElement.style.left = `${left}px`;
    this.tooltipElement.style.top = `${top}px`;
  },

  /**
   * Display the tooltip with specified text at given coordinates
   * @param {string} text - Text to display in the tooltip
   * @param {number} x - Initial X position
   * @param {number} y - Initial Y position
   */
  showTooltip(text, x, y) {
    if (!this.tooltipElement || !text) return;

    this.tooltipElement.textContent = text;
    this.tooltipElement.classList.add('visible');
    this.updateTooltipPosition(x, y);
  },

  /**
   * Hide the tooltip and clear the current cell reference
   */
  hideTooltip() {
    if (!this.tooltipElement) return;
    this.tooltipElement.classList.remove('visible');
    this.currentTooltipCell = null;
  },

  /**
   * Toggle synchronized scrolling on or off
   * When enabled, both tables scroll together
   * When disabled, tables can be scrolled independently
   */
  toggleSync() {
    this.syncEnabled = !this.syncEnabled;
    const syncBtn = document.getElementById('syncToggle');

    if (this.syncEnabled) {
      syncBtn.classList.add('active');
      syncBtn.querySelector('.sync-text').textContent = 'Sync Scroll';

      // Realign the scroll positions when re-enabling sync
      requestAnimationFrame(() => {
        this.realignScroll();
      });
    } else {
      syncBtn.classList.remove('active');
      syncBtn.querySelector('.sync-text').textContent = 'Disabled';
      this.isSyncingVertical = false;
    }
  },

  /**
   * Get the row number of the first fully visible row in a wrapper
   * Accounts for the sticky header height
   * @param {HTMLElement} wrapper - The table wrapper element
   * @returns {number|null} Row number (1-based), or null if not found
   */
  getFirstVisibleRowNumber(wrapper) {
    if (!wrapper) return null;

    const table = wrapper.querySelector('table');
    if (!table) return null;

    const tbody = table.querySelector('tbody');
    if (!tbody) return null;

    const rows = tbody.querySelectorAll('tr');
    if (rows.length === 0) return null;

    const scrollTop = wrapper.scrollTop + this.HEADER_HEIGHT;

    // Find the first row whose bottom edge is below the current scroll position
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowTop = row.offsetTop;
      const rowBottom = rowTop + row.offsetHeight;

      if (rowBottom > scrollTop) {
        const rowHeader = row.querySelector('.row-header');
        if (rowHeader) {
          return parseInt(rowHeader.textContent);
        }
      }
    }

    return null;
  },

  /**
   * Scroll a wrapper to display a specific row number
   * @param {HTMLElement} wrapper - The table wrapper to scroll
   * @param {number} rowNumber - The row number to scroll to (1-based)
   * @returns {boolean} True if successful, false if row not found
   */
  scrollToRowNumber(wrapper, rowNumber) {
    if (!wrapper) return false;

    const table = wrapper.querySelector('table');
    if (!table) return false;

    const tbody = table.querySelector('tbody');
    if (!tbody) return false;

    const rows = tbody.querySelectorAll('tr');

    // Find the row with the matching row number
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowHeader = row.querySelector('.row-header');

      if (rowHeader && parseInt(rowHeader.textContent) === rowNumber) {
        const targetScrollTop = Math.max(0, row.offsetTop - this.HEADER_HEIGHT);
        wrapper.scrollTop = targetScrollTop;
        return true;
      }
    }

    return false;
  },

  /**
   * Realign the scroll positions of both wrappers
   * Syncs to the earlier row and averages the horizontal scroll positions
   * Used when re-enabling synchronized scrolling
   */
  realignScroll() {
    if (!this.wrapperA || !this.wrapperB) return;

    const rowNumberA = this.getFirstVisibleRowNumber(this.wrapperA);
    const rowNumberB = this.getFirstVisibleRowNumber(this.wrapperB);

    if (!rowNumberA && !rowNumberB) {
      return;
    }

    // Sync to the earlier (smaller) row number to avoid scrolling past content
    const targetRowNumber = Math.min(rowNumberA || Infinity, rowNumberB || Infinity);

    this.isSyncingVertical = true;

    // Scroll both wrappers to the target row
    this.scrollToRowNumber(this.wrapperA, targetRowNumber);
    this.scrollToRowNumber(this.wrapperB, targetRowNumber);

    // Average the horizontal scroll positions for consistency
    const avgScrollLeft = (this.wrapperA.scrollLeft + this.wrapperB.scrollLeft) / 2;
    this.wrapperA.scrollLeft = avgScrollLeft;
    this.wrapperB.scrollLeft = avgScrollLeft;

    // Reset the syncing flag after a short delay
    if (this.verticalSyncTimeoutId) {
      clearTimeout(this.verticalSyncTimeoutId);
    }
    this.verticalSyncTimeoutId = setTimeout(() => {
      this.isSyncingVertical = false;
      this.verticalSyncTimeoutId = null;
    }, 150);
  },

  /**
   * Synchronize vertical scroll position from source to target wrapper
   * Matches the first visible row between tables
   * @param {HTMLElement} sourceWrapper - The wrapper being scrolled by the user
   * @param {HTMLElement} targetWrapper - The wrapper to sync to
   */
  syncVertical(sourceWrapper, targetWrapper) {
    const sourceRowNumber = this.getFirstVisibleRowNumber(sourceWrapper);

    if (!sourceRowNumber) return;

    this.scrollToRowNumber(targetWrapper, sourceRowNumber);
  },

  /**
   * Synchronize horizontal scroll position from source to target wrapper
   * Simply copies the scrollLeft value
   * @param {HTMLElement} sourceWrapper - The wrapper being scrolled
   * @param {HTMLElement} targetWrapper - The wrapper to sync to
   */
  syncHorizontal(sourceWrapper, targetWrapper) {
    targetWrapper.scrollLeft = sourceWrapper.scrollLeft;
  },

  /**
   * Display the diff viewer for a specific sheet
   * Hides the summary view and shows side-by-side comparison tables
   * @param {Object} diffResult - The complete diff comparison result
   * @param {Object} fileA - Parsed data from file A
   * @param {Object} fileB - Parsed data from file B
   * @param {string} sheetName - Name of the sheet to display
   * @param {string} status - Sheet status: 'modified', 'added', 'removed', or 'renamed'
   * @param {string|null} viewSide - Optional: 'A' or 'B' for single-sided view
   */
  show(diffResult, fileA, fileB, sheetName, status = 'modified', viewSide = null) {
    this.currentDiffResult = diffResult;
    this.currentFileA = fileA;
    this.currentFileB = fileB;
    this.currentSheet = sheetName;
    this.currentStatus = status;

    // Hide the summary view
    SummaryView.hide();

    // Show the diff comparison section
    document.getElementById('diffSection').style.display = 'block';

    // Update the sheet title display
    document.getElementById('currentSheetTitle').textContent = `Sheet: ${sheetName}`;

    // Update the file name labels
    document.getElementById('fileAName').textContent = diffResult.fileA;
    document.getElementById('fileBName').textContent = diffResult.fileB;

    // Initialize the sheet navigation bar
    NavBar.init(diffResult, fileA, fileB, sheetName, (name, status, side) => {
      this.show(diffResult, fileA, fileB, name, status, side);
    });

    // Render the comparison tables
    this.renderTables(sheetName, status, viewSide);
  },

  /**
   * Set up synchronized scrolling between both table wrappers
   * Clones the wrappers to remove old event listeners and adds new ones
   * Handles both vertical (row-based) and horizontal scroll synchronization
   */
  setupSyncScroll() {
    this.wrapperA = document.querySelector('#diffPaneA .table-wrapper');
    this.wrapperB = document.querySelector('#diffPaneB .table-wrapper');

    if (!this.wrapperA || !this.wrapperB) {
      console.warn('❌ Scroll containers not found');
      return;
    }

    /**
     * Check if a wrapper is scrollable and add appropriate CSS classes
     * Used to show/hide scroll shadow overlays
     */
    const checkScrollable = (wrapper, name) => {
      const pane = wrapper.closest('.diff-pane');
      const canScroll = wrapper.scrollHeight > wrapper.clientHeight || wrapper.scrollWidth > wrapper.clientWidth;

      if (canScroll) {
        wrapper.classList.add('scrollable');
        pane.classList.add('has-scrollable');
      } else {
        wrapper.classList.remove('scrollable');
        pane.classList.remove('has-scrollable');
      }
    };

    checkScrollable(this.wrapperA, 'wrapperA');
    checkScrollable(this.wrapperB, 'wrapperB');

    // Clone wrappers to remove all old event listeners
    const newWrapperA = this.wrapperA.cloneNode(true);
    const newWrapperB = this.wrapperB.cloneNode(true);

    this.wrapperA.parentNode.replaceChild(newWrapperA, this.wrapperA);
    this.wrapperB.parentNode.replaceChild(newWrapperB, this.wrapperB);

    this.wrapperA = newWrapperA;
    this.wrapperB = newWrapperB;

    // Add scroll event listener to wrapper A
    this.wrapperA.addEventListener(
      'scroll',
      () => {
        if (!this.syncEnabled) return;

        // Sync horizontal scroll immediately (smooth, no delay needed)
        this.syncHorizontal(this.wrapperA, this.wrapperB);

        // Skip vertical sync if already syncing to prevent infinite loops
        if (this.isSyncingVertical) return;

        // Perform vertical sync (row-based)
        this.isSyncingVertical = true;
        this.syncVertical(this.wrapperA, this.wrapperB);

        // Reset the syncing flag after a delay
        if (this.verticalSyncTimeoutId) {
          clearTimeout(this.verticalSyncTimeoutId);
        }
        this.verticalSyncTimeoutId = setTimeout(() => {
          this.isSyncingVertical = false;
          this.verticalSyncTimeoutId = null;
        }, 100);
      },
      { passive: true } // Improve scroll performance
    );

    // Add scroll event listener to wrapper B (same logic as A)
    this.wrapperB.addEventListener(
      'scroll',
      () => {
        if (!this.syncEnabled) return;

        this.syncHorizontal(this.wrapperB, this.wrapperA);

        if (this.isSyncingVertical) return;

        this.isSyncingVertical = true;
        this.syncVertical(this.wrapperB, this.wrapperA);

        if (this.verticalSyncTimeoutId) {
          clearTimeout(this.verticalSyncTimeoutId);
        }
        this.verticalSyncTimeoutId = setTimeout(() => {
          this.isSyncingVertical = false;
          this.verticalSyncTimeoutId = null;
        }, 100);
      },
      { passive: true }
    );

    console.log('✅ Synchronized scrolling set up!');
  },

  /**
   * Render both comparison tables based on sheet status
   * Handles different scenarios: added sheets, removed sheets, renamed sheets, and standard comparisons
   * @param {string} sheetName - Name of the sheet to render
   * @param {string} status - Sheet status: 'added', 'removed', 'renamed', or 'modified'
   * @param {string|null} viewSide - Optional: which side to view for single-sided sheets
   */
  renderTables(sheetName, status, viewSide) {
    const tableA = document.getElementById('tableA');
    const tableB = document.getElementById('tableB');

    if (status === 'added') {
      // Sheet only exists in File B (newly added sheet)
      this.renderEmptyTable(tableA, 'Sheet does not exist in File A');
      this.renderTable(tableB, sheetName, this.currentFileB, null, 'added');
    } else if (status === 'removed') {
      // Sheet only exists in File A (deleted sheet)
      this.renderTable(tableA, sheetName, this.currentFileA, null, 'removed');
      this.renderEmptyTable(tableB, 'Sheet does not exist in File B');
    } else if (status === 'renamed') {
      // Sheet was renamed between files
      const rename = this.currentDiffResult.sheetChanges.renamed.find((r) => r.to === sheetName);

      if (rename) {
        // Get cell-level differences using the new name as key
        const cellDiff = this.currentDiffResult.cellDiffs[rename.to];

        console.log('📊 Renamed sheet cell diff:', cellDiff);

        // Render using old name for table A, new name for table B
        this.renderTable(tableA, rename.from, this.currentFileA, cellDiff, 'comparison');
        this.renderTable(tableB, rename.to, this.currentFileB, cellDiff, 'comparison');
      }
    } else {
      // Standard comparison (modified or unchanged sheets)
      const cellDiff = this.currentDiffResult.cellDiffs[sheetName];
      this.renderTable(tableA, sheetName, this.currentFileA, cellDiff, 'comparison');
      this.renderTable(tableB, sheetName, this.currentFileB, cellDiff, 'comparison');
    }

    // Set up scroll synchronization and event bindings after DOM updates
    setTimeout(() => {
      this.setupSyncScroll();
      this.rebindTooltipEvents();
      this.collectChangedCells();
    }, 100);
  },

  /**
   * Rebind tooltip and click events to all changed cells
   * Must be called after table re-rendering to ensure events work
   * Also adds click handlers to update the current change index
   */
  rebindTooltipEvents() {
    const self = this;

    // Remove old event listeners by cloning and replacing elements
    document.querySelectorAll('td[data-tooltip]').forEach((td) => {
      td.replaceWith(td.cloneNode(true));
    });

    // Rebind fresh event listeners to all cells with tooltips
    document.querySelectorAll('td[data-tooltip]').forEach((td) => {
      // Show tooltip on mouse enter
      td.addEventListener('mouseenter', function (e) {
        const text = this.dataset.tooltip;
        if (text) {
          console.log('✅ mouseenter triggered:', text);
          self.currentTooltipCell = this;
          self.showTooltip(text, e.clientX, e.clientY);
        }
      });

      // Hide tooltip on mouse leave
      td.addEventListener('mouseleave', function () {
        console.log('✅ mouseleave triggered');
        self.hideTooltip();
      });

      // Update tooltip position as mouse moves
      td.addEventListener('mousemove', function (e) {
        if (self.tooltipElement && self.tooltipElement.classList.contains('visible')) {
          self.updateTooltipPosition(e.clientX, e.clientY);
        }
      });

      // Add click event to update the current change index
      // When user clicks a changed cell, find the nearest change and update navigation
      td.addEventListener('click', function () {
        const row = this.parentElement;
        const rowIndex = Array.from(row.parentElement.children).indexOf(row);
        const colIndex = Array.from(row.children).indexOf(this) - 1; // -1 to skip row header

        if (colIndex >= 0) {
          const nearestIndex = self.findNearestChangeIndex(rowIndex, colIndex);
          if (nearestIndex >= 0) {
            self.currentChangeIndex = nearestIndex;
            self.updateNavigationUI();
            console.log(`🎯 Clicked cell (${rowIndex}, ${colIndex}), nearest change: ${nearestIndex + 1}/${self.changedCells.length}`);
          }
        }
      });
    });

    console.log('✅ Tooltip events rebound,', document.querySelectorAll('td[data-tooltip]').length, 'cells total');
  },

  /**
   * Render an empty table with a centered message
   * Used when a sheet doesn't exist in one of the files
   * @param {HTMLElement} tableElement - The table element to render into
   * @param {string} message - The message to display
   */
  renderEmptyTable(tableElement, message) {
    tableElement.innerHTML = `
            <tr>
                <td style="text-align: center; padding: 40px; color: #999; font-style: italic;">
                    ${message}
                </td>
            </tr>
        `;
  },

  /**
   * Render a single comparison table with cell-level diff highlighting
   * Creates a complete HTML table with header row, row numbers, and colored cells
   * @param {HTMLElement} tableElement - The table element to render into
   * @param {string} sheetName - Name of the sheet to render
   * @param {Object} parsedFile - Parsed Excel file object
   * @param {Object|null} cellDiff - Cell difference data (null for single-sided views)
   * @param {string} mode - Render mode: 'comparison', 'added', or 'removed'
   */
  renderTable(tableElement, sheetName, parsedFile, cellDiff, mode) {
    const sheet = parsedFile.sheets.find((s) => s.name === sheetName);

    if (!sheet) {
      this.renderEmptyTable(tableElement, 'This Sheet cannot be found.');
      return;
    }

    const data = ExcelParser.normalizeData(sheet.data);

    // Create a lookup map for quick cell diff queries
    let diffMap = null;
    if (cellDiff && mode === 'comparison') {
      diffMap = DiffEngine.createCellDiffMap(cellDiff.changes);
      console.log(`🗺️ DiffMap created for ${sheetName}:`, diffMap.size, 'changes');
    }

    tableElement.innerHTML = '';

    // Create table header with column letters (A, B, C, ...)
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    // Empty cell for the top-left corner
    const emptyTh = document.createElement('th');
    emptyTh.textContent = '';
    headerRow.appendChild(emptyTh);

    // Column headers (A, B, C, ...)
    const maxCols = data.length > 0 ? data[0].length : 0;
    for (let c = 0; c < maxCols; c++) {
      const th = document.createElement('th');
      th.textContent = ExcelParser.getColumnName(c);
      headerRow.appendChild(th);
    }

    thead.appendChild(headerRow);
    tableElement.appendChild(thead);

    // Create table body with data cells
    const tbody = document.createElement('tbody');

    data.forEach((row, rowIndex) => {
      const tr = document.createElement('tr');

      // Add row number header (1, 2, 3, ...)
      const rowHeaderTd = document.createElement('td');
      rowHeaderTd.className = 'row-header';
      rowHeaderTd.textContent = rowIndex + 1;
      tr.appendChild(rowHeaderTd);

      // Add data cells for each column
      row.forEach((cellValue, colIndex) => {
        const td = document.createElement('td');

        td.textContent = cellValue || '';

        // Mark empty cells with special styling
        if (cellValue === '' || cellValue === null || cellValue === undefined) {
          td.classList.add('cell-empty');
          td.textContent = '(Empty)';
        }

        // Apply diff highlighting if in comparison mode
        if (diffMap) {
          const diff = DiffEngine.getCellDiff(rowIndex, colIndex, diffMap);

          if (diff) {
            // Add appropriate CSS class based on change type
            td.classList.add(`cell-${diff.type}`);

            // Create tooltip text showing old and new values
            let tooltipText = '';
            if (diff.type === 'modified') {
              tooltipText = `Old Value: ${diff.oldValue}\nNew Value: ${diff.newValue}`;
            } else if (diff.type === 'added') {
              tooltipText = `Added: ${diff.newValue}`;
            } else if (diff.type === 'removed') {
              tooltipText = `Removed: ${diff.oldValue}`;
            }

            // Store tooltip text in data attribute (events bound later)
            td.dataset.tooltip = tooltipText;
          } else {
            td.classList.add('cell-unchanged');
          }
        } else if (mode === 'added') {
          // All cells are marked as added (for new sheets)
          td.classList.add('cell-added');
        } else if (mode === 'removed') {
          // All cells are marked as removed (for deleted sheets)
          td.classList.add('cell-removed');
        } else {
          // No highlighting needed
          td.classList.add('cell-unchanged');
        }

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    tableElement.appendChild(tbody);
  },

  /**
   * Hide the diff viewer and tooltip
   * Returns to the summary view
   */
  hide() {
    document.getElementById('diffSection').style.display = 'none';
    this.hideTooltip();
  },
};

// Initialize all diff viewer systems when the page loads
document.addEventListener('DOMContentLoaded', () => {
  DiffViewer.initSyncButton();
  DiffViewer.initTooltip();
  DiffViewer.initChangeNavigation();
});
