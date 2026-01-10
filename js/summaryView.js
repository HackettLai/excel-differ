// summaryView.js - Summary view renderer

const SummaryView = {
  currentDiffResult: null,
  parsedFileA: null,
  parsedFileB: null,

  /**
   * Display the summary view
   * Shows overall comparison results and list of sheet changes
   * @param {Object} diffResult - Diff comparison result
   * @param {Object} fileA - Parsed file A
   * @param {Object} fileB - Parsed file B
   */
  show(diffResult, fileA, fileB) {
    this.currentDiffResult = diffResult;
    this.parsedFileA = fileA;
    this.parsedFileB = fileB;

    // Hide upload section
    document.getElementById('uploadSection').style.display = 'none';

    // Show summary section
    document.getElementById('summarySection').style.display = 'block';

    // Populate file information
    document.getElementById('summaryFileA').textContent = `${diffResult.fileA} (${fileA.sheetCount} sheets)`;
    document.getElementById('summaryFileB').textContent = `${diffResult.fileB} (${fileB.sheetCount} sheets)`;

    // Render sheet changes list
    this.renderSheetChanges();
  },

  /**
   * Render the list of sheet changes
   * Groups sheets by status: common, renamed, added, removed
   */
  renderSheetChanges() {
    const container = document.getElementById('sheetChangesList');
    container.innerHTML = '';

    const { sheetChanges, cellDiffs } = this.currentDiffResult;

    // Collect all sheets with their status
    const allSheets = [];

    // 1. Common sheets (present in both files)
    sheetChanges.common.forEach((sheetName) => {
      const diff = cellDiffs[sheetName];
      allSheets.push({
        name: sheetName,
        status: diff.changes.length > 0 ? 'modified' : 'unchanged',
        changeCount: diff.totalChanges,
        canView: true,
      });
    });

    // 2. Renamed sheets
    sheetChanges.renamed.forEach((rename) => {
      allSheets.push({
        name: `${rename.from} → ${rename.to}`,
        originalName: rename.from,
        newName: rename.to,
        status: 'renamed',
        confidence: Math.round(rename.confidence * 100),
        canView: true,
      });
    });

    // 3. Added sheets (only in file B)
    sheetChanges.added.forEach((sheetName) => {
      allSheets.push({
        name: sheetName,
        status: 'added',
        canView: true,
        viewSide: 'B',
      });
    });

    // 4. Removed sheets (only in file A)
    sheetChanges.removed.forEach((sheetName) => {
      allSheets.push({
        name: sheetName,
        status: 'removed',
        canView: true,
        viewSide: 'A',
      });
    });

    // Render each sheet item
    allSheets.forEach((sheet) => {
      const item = this.createSheetItem(sheet);
      container.appendChild(item);
    });

    // If no changes found
    if (allSheets.length === 0) {
      container.innerHTML = '<p style="text-align: center; color: #666;">The two files are identical</p>';
    }
  },

  /**
   * Create a sheet item element for the summary list
   * @param {Object} sheet - Sheet information object
   * @param {string} sheet.name - Sheet name or display name
   * @param {string} sheet.status - Sheet status: 'unchanged', 'modified', 'added', 'removed', 'renamed'
   * @param {number} [sheet.changeCount] - Number of changes (for modified sheets)
   * @param {number} [sheet.confidence] - Similarity confidence (for renamed sheets)
   * @param {boolean} sheet.canView - Whether the sheet can be viewed
   * @param {string} [sheet.viewSide] - Which side to view ('A' or 'B')
   * @returns {HTMLElement} Sheet item element
   */
  createSheetItem(sheet) {
    const item = document.createElement('div');
    item.className = 'sheet-item';

    // Sheet name section
    const nameDiv = document.createElement('div');
    nameDiv.className = 'sheet-item-name';

    // Add status icon
    const icon = this.getStatusIcon(sheet.status);
    nameDiv.innerHTML = `<span>${icon}</span><strong>${sheet.name}</strong>`;

    // Status label
    const statusDiv = document.createElement('div');
    statusDiv.className = `sheet-status status-${sheet.status}`;
    statusDiv.textContent = this.getStatusText(sheet);

    item.appendChild(nameDiv);
    item.appendChild(statusDiv);

    // Add click event if sheet can be viewed
    if (sheet.canView) {
      item.style.cursor = 'pointer';
      item.addEventListener('click', () => {
        this.viewSheetDiff(sheet);
      });
    }

    return item;
  },

  /**
   * Get status icon emoji for a given status
   * @param {string} status - Sheet status
   * @returns {string} Emoji icon
   */
  getStatusIcon(status) {
    const icons = {
      unchanged: '✅',
      modified: '✏️',
      added: '➕',
      removed: '❌',
      renamed: '🔄',
    };
    return icons[status] || '•';
  },

  /**
   * Get human-readable status text
   * @param {Object} sheet - Sheet information object
   * @returns {string} Status description text
   */
  getStatusText(sheet) {
    switch (sheet.status) {
      case 'unchanged':
        return 'Unchanged';
      case 'modified':
        return `${sheet.changeCount} changes`;
      case 'added':
        return 'New Sheet';
      case 'removed':
        return 'Sheet Removed';
      case 'renamed':
        return `Renamed (${sheet.confidence}% Similar)`;
      default:
        return '';
    }
  },

  /**
   * View detailed diff for a specific sheet
   * Opens the diff viewer for the selected sheet
   * @param {Object} sheet - Sheet to view
   */
  viewSheetDiff(sheet) {
    // Prepare sheet name
    let sheetToView = sheet.name;

    // If renamed, use the new name
    if (sheet.status === 'renamed') {
      sheetToView = sheet.newName;
    }

    // Switch to diff view
    DiffViewer.show(this.currentDiffResult, this.parsedFileA, this.parsedFileB, sheetToView, sheet.status, sheet.viewSide);
  },

  /**
   * Hide the summary view
   */
  hide() {
    document.getElementById('summarySection').style.display = 'none';
  },
};
