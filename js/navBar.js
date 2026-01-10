// navBar.js - Sheet navigation bar

const NavBar = {
  currentSheet: null,
  diffResult: null,
  parsedFileA: null,
  parsedFileB: null,
  onSheetChange: null,

  /**
   * Initialize the navigation bar
   * @param {Object} diffResult - Diff comparison result
   * @param {Object} fileA - Parsed file A
   * @param {Object} fileB - Parsed file B
   * @param {string} currentSheet - Currently selected sheet name
   * @param {Function} onSheetChangeCallback - Callback when sheet is changed
   */
  init(diffResult, fileA, fileB, currentSheet, onSheetChangeCallback) {
    this.diffResult = diffResult;
    this.parsedFileA = fileA;
    this.parsedFileB = fileB;
    this.currentSheet = currentSheet;
    this.onSheetChange = onSheetChangeCallback;

    this.render();
  },

  /**
   * Render the navigation bar with all sheet tabs
   * Groups sheets by status: common, renamed, added, removed
   */
  render() {
    const navBar = document.getElementById('sheetNavBar');
    navBar.innerHTML = '';

    const { sheetChanges, cellDiffs } = this.diffResult;

    // Collect all visible sheet tabs
    const tabs = [];

    // 1. Common sheets (includes modified and unchanged)
    sheetChanges.common.forEach((sheetName) => {
      const diff = cellDiffs[sheetName];
      tabs.push({
        name: sheetName,
        displayName: sheetName,
        status: diff.changes.length > 0 ? 'modified' : 'unchanged',
        available: true,
      });
    });

    // 2. Renamed sheets
    sheetChanges.renamed.forEach((rename) => {
      tabs.push({
        name: rename.to,
        displayName: `${rename.from} → ${rename.to}`,
        status: 'renamed',
        available: true,
      });
    });

    // 3. Added sheets (only in file B)
    sheetChanges.added.forEach((sheetName) => {
      tabs.push({
        name: sheetName,
        displayName: sheetName,
        status: 'added',
        available: true,
        side: 'B',
      });
    });

    // 4. Removed sheets (only in file A)
    sheetChanges.removed.forEach((sheetName) => {
      tabs.push({
        name: sheetName,
        displayName: sheetName,
        status: 'removed',
        available: true,
        side: 'A',
      });
    });

    // Render each tab
    tabs.forEach((tab) => {
      const tabElement = this.createTab(tab);
      navBar.appendChild(tabElement);
    });
  },

  /**
   * Create a single navigation tab element
   * @param {Object} tab - Tab configuration object
   * @param {string} tab.name - Sheet name
   * @param {string} tab.displayName - Display name for the tab
   * @param {string} tab.status - Sheet status: 'modified', 'added', 'removed', 'unchanged', 'renamed'
   * @param {string} [tab.side] - Optional side indicator ('A' or 'B')
   * @returns {HTMLElement} Tab element
   */
  createTab(tab) {
    const tabDiv = document.createElement('div');
    tabDiv.className = 'sheet-tab';

    // Add status class
    if (tab.status === 'modified') {
      tabDiv.classList.add('has-changes');
    } else if (tab.status === 'added') {
      tabDiv.classList.add('added');
    } else if (tab.status === 'removed') {
      tabDiv.classList.add('removed');
    }

    // Add active class if this is the current sheet
    if (tab.name === this.currentSheet) {
      tabDiv.classList.add('active');
    }

    // Set tab text
    tabDiv.textContent = tab.displayName;

    // Add click event listener
    tabDiv.addEventListener('click', () => {
      if (this.onSheetChange) {
        this.onSheetChange(tab.name, tab.status, tab.side);
      }
    });

    return tabDiv;
  },

  /**
   * Update the currently active sheet
   * Highlights the selected tab and removes highlight from others
   * @param {string} sheetName - Name of the sheet to set as active
   */
  setActiveSheet(sheetName) {
    this.currentSheet = sheetName;

    // Update active state for all tabs
    const tabs = document.querySelectorAll('.sheet-tab');
    tabs.forEach((tab) => {
      tab.classList.remove('active');
      // Check if tab name matches (handles both simple names and renamed format)
      if (tab.textContent.includes(sheetName) || tab.textContent === sheetName) {
        tab.classList.add('active');
      }
    });
  },
};
