// diffViewer.js - Diff Viewer Module

const DiffViewer = {
  currentDiffResult: null,
  currentFileA: null,
  currentFileB: null,
  currentSheet: null,
  currentStatus: null,
  isSyncingVertical: false,
  syncEnabled: true,
  wrapperA: null,
  wrapperB: null,
  HEADER_HEIGHT: 38,
  verticalSyncTimeoutId: null,
  tooltipElement: null,
  currentTooltipCell: null,

  /**
   * Initialize the sync scroll toggle button
   */
  initSyncButton() {
    const syncBtn = document.getElementById('syncToggle');
    if (!syncBtn) return;

    syncBtn.addEventListener('click', () => {
      this.toggleSync();
    });
  },

  /**
   * Initialize the tooltip element and event listeners
   */
  initTooltip() {
    this.tooltipElement = document.getElementById('customTooltip');
    if (!this.tooltipElement) {
      console.warn('❌ Tooltip element not found');
    } else {
      console.log('✅ Tooltip element found!');

      // Update tooltip position on mouse move
      document.addEventListener('mousemove', (e) => {
        if (this.currentTooltipCell && this.tooltipElement.classList.contains('visible')) {
          this.updateTooltipPosition(e.clientX, e.clientY);
        }
      });
    }
  },

  /**
   * Update tooltip position based on mouse coordinates
   * Ensures tooltip stays within viewport boundaries
   */
  updateTooltipPosition(x, y) {
    if (!this.tooltipElement) return;

    const tooltipRect = this.tooltipElement.getBoundingClientRect();
    const padding = 10;

    let left = x + padding;
    let top = y + padding;

    // Prevent tooltip from going off-screen horizontally
    if (left + tooltipRect.width > window.innerWidth) {
      left = x - tooltipRect.width - padding;
    }

    // Prevent tooltip from going off-screen vertically
    if (top + tooltipRect.height > window.innerHeight) {
      top = y - tooltipRect.height - padding;
    }

    this.tooltipElement.style.left = `${left}px`;
    this.tooltipElement.style.top = `${top}px`;
  },

  /**
   * Show tooltip with given text at specified coordinates
   */
  showTooltip(text, x, y) {
    if (!this.tooltipElement || !text) return;

    this.tooltipElement.textContent = text;
    this.tooltipElement.classList.add('visible');
    this.updateTooltipPosition(x, y);
  },

  /**
   * Hide the tooltip
   */
  hideTooltip() {
    if (!this.tooltipElement) return;
    this.tooltipElement.classList.remove('visible');
    this.currentTooltipCell = null;
  },

  /**
   * Toggle synchronized scrolling on/off
   */
  toggleSync() {
    this.syncEnabled = !this.syncEnabled;
    const syncBtn = document.getElementById('syncToggle');

    if (this.syncEnabled) {
      syncBtn.classList.add('active');
      syncBtn.querySelector('.sync-text').textContent = 'Sync Scroll';

      // Realign scroll positions when sync is enabled
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
   * Get the row number of the first visible row in a wrapper
   * Accounts for header height offset
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

    // Find the first row whose bottom edge is below the scroll position
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
   * Scroll wrapper to display the specified row number
   * Returns true if successful, false otherwise
   */
  scrollToRowNumber(wrapper, rowNumber) {
    if (!wrapper) return false;

    const table = wrapper.querySelector('table');
    if (!table) return false;

    const tbody = table.querySelector('tbody');
    if (!tbody) return false;

    const rows = tbody.querySelectorAll('tr');

    // Find the row with matching row number
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
   * Realign scroll positions of both wrappers
   * Syncs to the earlier row and averages horizontal scroll
   */
  realignScroll() {
    if (!this.wrapperA || !this.wrapperB) return;

    const rowNumberA = this.getFirstVisibleRowNumber(this.wrapperA);
    const rowNumberB = this.getFirstVisibleRowNumber(this.wrapperB);

    if (!rowNumberA && !rowNumberB) {
      return;
    }

    // Sync to the earlier (smaller) row number
    const targetRowNumber = Math.min(rowNumberA || Infinity, rowNumberB || Infinity);

    this.isSyncingVertical = true;

    // Scroll both wrappers to the target row
    this.scrollToRowNumber(this.wrapperA, targetRowNumber);
    this.scrollToRowNumber(this.wrapperB, targetRowNumber);

    // Average the horizontal scroll positions
    const avgScrollLeft = (this.wrapperA.scrollLeft + this.wrapperB.scrollLeft) / 2;
    this.wrapperA.scrollLeft = avgScrollLeft;
    this.wrapperB.scrollLeft = avgScrollLeft;

    // Reset syncing flag after a delay
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
   */
  syncVertical(sourceWrapper, targetWrapper) {
    const sourceRowNumber = this.getFirstVisibleRowNumber(sourceWrapper);

    if (!sourceRowNumber) return;

    this.scrollToRowNumber(targetWrapper, sourceRowNumber);
  },

  /**
   * Synchronize horizontal scroll position from source to target wrapper
   */
  syncHorizontal(sourceWrapper, targetWrapper) {
    targetWrapper.scrollLeft = sourceWrapper.scrollLeft;
  },

  /**
   * Display the diff viewer with the specified sheet and status
   * @param {Object} diffResult - The diff comparison result
   * @param {Object} fileA - Parsed file A
   * @param {Object} fileB - Parsed file B
   * @param {string} sheetName - Name of the sheet to display
   * @param {string} status - Status: 'modified', 'added', 'removed', or 'renamed'
   * @param {string|null} viewSide - Optional side to view
   */
  show(diffResult, fileA, fileB, sheetName, status = 'modified', viewSide = null) {
    this.currentDiffResult = diffResult;
    this.currentFileA = fileA;
    this.currentFileB = fileB;
    this.currentSheet = sheetName;
    this.currentStatus = status;

    // Hide summary view
    SummaryView.hide();

    // Show diff section
    document.getElementById('diffSection').style.display = 'block';

    // Update sheet title
    document.getElementById('currentSheetTitle').textContent = `Sheet: ${sheetName}`;

    // Update file names
    document.getElementById('fileAName').textContent = diffResult.fileA;
    document.getElementById('fileBName').textContent = diffResult.fileB;

    // Initialize navigation bar
    NavBar.init(diffResult, fileA, fileB, sheetName, (name, status, side) => {
      this.show(diffResult, fileA, fileB, name, status, side);
    });

    // Render comparison tables
    this.renderTables(sheetName, status, viewSide);
  },

  /**
   * Set up synchronized scrolling between both table wrappers
   * Handles both vertical and horizontal scroll synchronization
   */
  setupSyncScroll() {
    this.wrapperA = document.querySelector('#diffPaneA .table-wrapper');
    this.wrapperB = document.querySelector('#diffPaneB .table-wrapper');

    if (!this.wrapperA || !this.wrapperB) {
      console.warn('❌ Scroll containers not found');
      return;
    }

    /**
     * Check if wrapper is scrollable and add appropriate classes
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

    // Clone nodes to remove old event listeners
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

        // Sync horizontal scroll immediately
        this.syncHorizontal(this.wrapperA, this.wrapperB);

        // Skip vertical sync if already syncing
        if (this.isSyncingVertical) return;

        // Perform vertical sync
        this.isSyncingVertical = true;
        this.syncVertical(this.wrapperA, this.wrapperB);

        // Reset syncing flag after delay
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

    // Add scroll event listener to wrapper B
    this.wrapperB.addEventListener(
      'scroll',
      () => {
        if (!this.syncEnabled) return;

        // Sync horizontal scroll immediately
        this.syncHorizontal(this.wrapperB, this.wrapperA);

        // Skip vertical sync if already syncing
        if (this.isSyncingVertical) return;

        // Perform vertical sync
        this.isSyncingVertical = true;
        this.syncVertical(this.wrapperB, this.wrapperA);

        // Reset syncing flag after delay
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

    console.log('✅ Synchronized scrolling set up (separate vertical/horizontal)!');
  },

  /**
   * Render both comparison tables based on sheet status
   * @param {string} sheetName - Name of the sheet
   * @param {string} status - Sheet status: 'added', 'removed', 'renamed', or 'modified'
   * @param {string|null} viewSide - Optional side to view
   */
  renderTables(sheetName, status, viewSide) {
    const tableA = document.getElementById('tableA');
    const tableB = document.getElementById('tableB');

    if (status === 'added') {
      // Sheet only exists in File B
      this.renderEmptyTable(tableA, 'Sheet does not exist in File A');
      this.renderTable(tableB, sheetName, this.currentFileB, null, 'added');
    } else if (status === 'removed') {
      // Sheet only exists in File A
      this.renderTable(tableA, sheetName, this.currentFileA, null, 'removed');
      this.renderEmptyTable(tableB, 'Sheet does not exist in File B');
    } else if (status === 'renamed') {
      // Sheet was renamed between files
      const rename = this.currentDiffResult.sheetChanges.renamed.find((r) => r.to === sheetName);

      if (rename) {
        const cellDiff = this.currentDiffResult.cellDiffs[sheetName];
        this.renderTable(tableA, rename.from, this.currentFileA, cellDiff, 'comparison');
        this.renderTable(tableB, rename.to, this.currentFileB, cellDiff, 'comparison');
      }
    } else {
      // Standard comparison (modified or unchanged)
      const cellDiff = this.currentDiffResult.cellDiffs[sheetName];
      this.renderTable(tableA, sheetName, this.currentFileA, cellDiff, 'comparison');
      this.renderTable(tableB, sheetName, this.currentFileB, cellDiff, 'comparison');
    }

    // Set up sync scroll and rebind tooltip events after render
    setTimeout(() => {
      this.setupSyncScroll();
      this.rebindTooltipEvents();
    }, 100);
  },

  /**
   * Rebind tooltip events to all cells with data-tooltip attribute
   * Called after table re-rendering to ensure tooltips work
   */
  rebindTooltipEvents() {
    const self = this;

    // Remove old event listeners by cloning elements
    document.querySelectorAll('td[data-tooltip]').forEach((td) => {
      td.replaceWith(td.cloneNode(true));
    });

    // Rebind new event listeners
    document.querySelectorAll('td[data-tooltip]').forEach((td) => {
      td.addEventListener('mouseenter', function (e) {
        const text = this.dataset.tooltip;
        if (text) {
          console.log('✅ mouseenter triggered:', text);
          self.currentTooltipCell = this;
          self.showTooltip(text, e.clientX, e.clientY);
        }
      });

      td.addEventListener('mouseleave', function () {
        console.log('✅ mouseleave triggered');
        self.hideTooltip();
      });
    });

    console.log('✅ Tooltip events rebound,', document.querySelectorAll('td[data-tooltip]').length, 'cells total');
  },

  /**
   * Render an empty table with a message
   * Used when a sheet doesn't exist in one of the files
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
   * Render a single table with cell diff highlighting
   * @param {HTMLElement} tableElement - The table element to render into
   * @param {string} sheetName - Name of the sheet
   * @param {Object} parsedFile - Parsed Excel file object
   * @param {Object|null} cellDiff - Cell difference data
   * @param {string} mode - Render mode: 'comparison', 'added', or 'removed'
   */
  renderTable(tableElement, sheetName, parsedFile, cellDiff, mode) {
    const sheet = parsedFile.sheets.find((s) => s.name === sheetName);

    if (!sheet) {
      this.renderEmptyTable(tableElement, 'This Sheet cannot be found.');
      return;
    }

    const data = ExcelParser.normalizeData(sheet.data);

    // Create diff map for cell highlighting
    let diffMap = null;
    if (cellDiff && mode === 'comparison') {
      diffMap = DiffEngine.createCellDiffMap(cellDiff.changes);
    }

    tableElement.innerHTML = '';

    // Create table header with column letters (A, B, C, ...)
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    const emptyTh = document.createElement('th');
    emptyTh.textContent = '';
    headerRow.appendChild(emptyTh);

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

      // Add row number header
      const rowHeaderTd = document.createElement('td');
      rowHeaderTd.className = 'row-header';
      rowHeaderTd.textContent = rowIndex + 1;
      tr.appendChild(rowHeaderTd);

      // Add data cells
      row.forEach((cellValue, colIndex) => {
        const td = document.createElement('td');

        td.textContent = cellValue || '';

        // Mark empty cells
        if (cellValue === '' || cellValue === null || cellValue === undefined) {
          td.classList.add('cell-empty');
          td.textContent = '(Empty)';
        }

        // Apply diff highlighting if in comparison mode
        if (diffMap) {
          const diff = DiffEngine.getCellDiff(rowIndex, colIndex, diffMap);

          if (diff) {
            td.classList.add(`cell-${diff.type}`);

            // Create tooltip text based on diff type
            let tooltipText = '';
            if (diff.type === 'modified') {
              tooltipText = `Old Value: ${diff.oldValue}\nNew Value: ${diff.newValue}`;
            } else if (diff.type === 'added') {
              tooltipText = `Added: ${diff.newValue}`;
            } else if (diff.type === 'removed') {
              tooltipText = `Removed: ${diff.oldValue}`;
            }

            // Set tooltip data attribute (event binding happens in rebindTooltipEvents)
            td.dataset.tooltip = tooltipText;
          } else {
            td.classList.add('cell-unchanged');
          }
        } else if (mode === 'added') {
          // All cells are added
          td.classList.add('cell-added');
        } else if (mode === 'removed') {
          // All cells are removed
          td.classList.add('cell-removed');
        } else {
          // No diff highlighting
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
   */
  hide() {
    document.getElementById('diffSection').style.display = 'none';
    this.hideTooltip();
  },
};

// Initialize sync button and tooltip when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
  DiffViewer.initSyncButton();
  DiffViewer.initTooltip();
});
