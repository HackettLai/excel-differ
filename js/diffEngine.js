// diffEngine.js - Diff algorithm core engine (FIXED)

const DiffEngine = {
  /**
   * Main comparison function
   * Compares two parsed Excel files and returns detailed diff results
   * @param {Object} parsedFileA - Parsed file A data
   * @param {Object} parsedFileB - Parsed file B data
   * @returns {Object} Comprehensive diff result with sheet changes and cell diffs
   */
  compare(parsedFileA, parsedFileB) {
    const result = {
      fileA: parsedFileA.fileName,
      fileB: parsedFileB.fileName,
      sheetChanges: this.compareSheets(parsedFileA.sheets, parsedFileB.sheets),
      cellDiffs: {},
      statistics: {
        totalSheets: 0,
        addedSheets: 0,
        removedSheets: 0,
        modifiedSheets: 0,
        unchangedSheets: 0,
      },
    };

    // Compare the contents of each common sheet
    result.sheetChanges.common.forEach((sheetName) => {
      const sheetA = parsedFileA.sheets.find((s) => s.name === sheetName);
      const sheetB = parsedFileB.sheets.find((s) => s.name === sheetName);

      const cellDiff = this.compareCells(sheetA, sheetB);
      result.cellDiffs[sheetName] = cellDiff;
    });

    // Also compare contents for renamed sheets!
    result.sheetChanges.renamed.forEach((rename) => {
      const sheetA = parsedFileA.sheets.find((s) => s.name === rename.from);
      const sheetB = parsedFileB.sheets.find((s) => s.name === rename.to);

      const cellDiff = this.compareCells(sheetA, sheetB);
      result.cellDiffs[rename.to] = cellDiff; // Use new name as key
    });

    // Calculate statistics
    result.statistics.totalSheets = parsedFileA.sheets.length + result.sheetChanges.added.length;
    result.statistics.addedSheets = result.sheetChanges.added.length;
    result.statistics.removedSheets = result.sheetChanges.removed.length;
    result.statistics.renamedSheets = result.sheetChanges.renamed.length;

    // Calculate modified and unchanged sheet counts
    result.sheetChanges.common.forEach((sheetName) => {
      const diff = result.cellDiffs[sheetName];
      if (diff.changes.length > 0) {
        result.statistics.modifiedSheets++;
      } else {
        result.statistics.unchangedSheets++;
      }
    });

    return result;
  },

  /**
   * Compare sheet-level changes between two files
   * Identifies added, removed, common, and renamed sheets
   * @param {Array} sheetsA - Sheets from file A
   * @param {Array} sheetsB - Sheets from file B
   * @returns {Object} Sheet comparison result
   */
  compareSheets(sheetsA, sheetsB) {
    const namesA = sheetsA.map((s) => s.name);
    const namesB = sheetsB.map((s) => s.name);

    const setA = new Set(namesA);
    const setB = new Set(namesB);

    const added = namesB.filter((name) => !setA.has(name));
    const removed = namesA.filter((name) => !setB.has(name));
    const common = namesA.filter((name) => setB.has(name));

    // Detect possible renames
    const renamed = this.detectRenames(removed, added, sheetsA, sheetsB);

    // Remove identified renamed sheets from added and removed lists
    const renamedOldNames = renamed.map((r) => r.from);
    const renamedNewNames = renamed.map((r) => r.to);

    const finalAdded = added.filter((name) => !renamedNewNames.includes(name));
    const finalRemoved = removed.filter((name) => !renamedOldNames.includes(name));

    return {
      added: finalAdded,
      removed: finalRemoved,
      common: common,
      renamed: renamed,
    };
  },

  /**
   * Detect sheet renaming by comparing content similarity
   * Uses similarity threshold to identify likely renames
   * @param {Array} removedNames - Sheet names removed from file A
   * @param {Array} addedNames - Sheet names added in file B
   * @param {Array} sheetsA - Sheets from file A
   * @param {Array} sheetsB - Sheets from file B
   * @returns {Array} Array of rename objects with from/to/confidence
   */
  detectRenames(removedNames, addedNames, sheetsA, sheetsB) {
    const renames = [];
    const threshold = 0.8; //  Lowered from 0.85 to 0.80

    removedNames.forEach((oldName) => {
      const oldSheet = sheetsA.find((s) => s.name === oldName);

      addedNames.forEach((newName) => {
        const newSheet = sheetsB.find((s) => s.name === newName);

        const similarity = this.calculateSheetSimilarity(oldSheet, newSheet);

        console.log(`🔍 Comparing "${oldName}" vs "${newName}": ${(similarity * 100).toFixed(2)}%`);

        if (similarity >= threshold) {
          renames.push({
            from: oldName,
            to: newName,
            confidence: similarity,
          });
        }
      });
    });

    // Only keep the highest similarity match for each sheet
    const finalRenames = [];
    const usedOld = new Set();
    const usedNew = new Set();

    // Sort by similarity (highest first)
    renames.sort((a, b) => b.confidence - a.confidence);

    renames.forEach((rename) => {
      if (!usedOld.has(rename.from) && !usedNew.has(rename.to)) {
        finalRenames.push(rename);
        usedOld.add(rename.from);
        usedNew.add(rename.to);
      }
    });

    return finalRenames;
  },

  /**
   * Calculate similarity between two sheets based on content
   * Compares dimensions and cell values in first 10 rows
   * @param {Object} sheetA - Sheet from file A
   * @param {Object} sheetB - Sheet from file B
   * @returns {number} Similarity score between 0 and 1
   */
  calculateSheetSimilarity(sheetA, sheetB) {
    // If dimension differences are too large, similarity is 0
    const rowDiff = Math.abs(sheetA.rowCount - sheetB.rowCount);
    const colDiff = Math.abs(sheetA.colCount - sheetB.colCount);

    if (rowDiff > 10 || colDiff > 5) {
      return 0;
    }

    // Compare first 10 rows or actual row count (whichever is smaller)
    const maxRow = Math.min(10, sheetA.data.length, sheetB.data.length);
    let matches = 0;
    let total = 0;

    for (let r = 0; r < maxRow; r++) {
      const rowA = sheetA.data[r] || [];
      const rowB = sheetB.data[r] || [];
      const maxCol = Math.max(rowA.length, rowB.length);

      for (let c = 0; c < maxCol; c++) {
        total++;
        const valA = rowA[c];
        const valB = rowB[c];

        if (ExcelParser.areValuesEqual(valA, valB)) {
          matches++;
        }
      }
    }

    return total > 0 ? matches / total : 0;
  },

  /**
   * Compare cell-level content between two sheets
   * Identifies added, removed, and modified cells
   * @param {Object} sheetA - Sheet from file A
   * @param {Object} sheetB - Sheet from file B
   * @returns {Object} Cell diff result with changes and statistics
   */
  compareCells(sheetA, sheetB) {
    const changes = [];

    // Normalize data for easier comparison
    const dataA = ExcelParser.normalizeData(sheetA.data);
    const dataB = ExcelParser.normalizeData(sheetB.data);

    const maxRow = Math.max(dataA.length, dataB.length);
    const maxCol = Math.max(dataA.length > 0 ? dataA[0].length : 0, dataB.length > 0 ? dataB[0].length : 0);

    // Cell-by-cell comparison
    for (let r = 0; r < maxRow; r++) {
      const rowA = dataA[r] || [];
      const rowB = dataB[r] || [];

      for (let c = 0; c < maxCol; c++) {
        const valA = rowA[c];
        const valB = rowB[c];

        if (!ExcelParser.areValuesEqual(valA, valB)) {
          const isEmpty_A = valA === null || valA === undefined || valA === '';
          const isEmpty_B = valB === null || valB === undefined || valB === '';

          let changeType;
          if (isEmpty_A && !isEmpty_B) {
            changeType = 'added';
          } else if (!isEmpty_A && isEmpty_B) {
            changeType = 'removed';
          } else {
            changeType = 'modified';
          }

          changes.push({
            row: r,
            col: c,
            type: changeType,
            oldValue: valA,
            newValue: valB,
            cellRef: this.getCellReference(r, c),
          });
        }
      }
    }

    return {
      changes: changes,
      totalChanges: changes.length,
      additions: changes.filter((c) => c.type === 'added').length,
      deletions: changes.filter((c) => c.type === 'removed').length,
      modifications: changes.filter((c) => c.type === 'modified').length,
    };
  },

  /**
   * Get cell reference in Excel notation (e.g., A1, B2, AA10)
   * @param {number} row - Zero-based row index
   * @param {number} col - Zero-based column index
   * @returns {string} Cell reference (e.g., 'A1')
   */
  getCellReference(row, col) {
    return ExcelParser.getColumnName(col) + (row + 1);
  },

  /**
   * Get the status of a sheet
   * @param {string} sheetName - Name of the sheet
   * @param {Object} sheetChanges - Sheet changes object
   * @param {Object} cellDiffs - Cell diffs object
   * @returns {string} Status: 'added', 'removed', 'renamed', 'modified', or 'unchanged'
   */
  getSheetStatus(sheetName, sheetChanges, cellDiffs) {
    if (sheetChanges.added.includes(sheetName)) {
      return 'added';
    }

    if (sheetChanges.removed.includes(sheetName)) {
      return 'removed';
    }

    const renamed = sheetChanges.renamed.find((r) => r.from === sheetName || r.to === sheetName);
    if (renamed) {
      return 'renamed';
    }

    if (cellDiffs[sheetName]) {
      const diff = cellDiffs[sheetName];
      if (diff.changes.length > 0) {
        return 'modified';
      }
    }

    return 'unchanged';
  },

  /**
   * Create a cell diff map for quick lookups
   * Maps "row-col" keys to change objects
   * @param {Array} changes - Array of cell changes
   * @returns {Map} Map of cell coordinates to change objects
   */
  createCellDiffMap(changes) {
    const map = new Map();

    changes.forEach((change) => {
      const key = `${change.row}-${change.col}`;
      map.set(key, change);
    });

    return map;
  },

  /**
   * Get diff information for a specific cell
   * @param {number} row - Zero-based row index
   * @param {number} col - Zero-based column index
   * @param {Map} diffMap - Cell diff map
   * @returns {Object|undefined} Change object if cell has changes, undefined otherwise
   */
  getCellDiff(row, col, diffMap) {
    const key = `${row}-${col}`;
    return diffMap.get(key);
  },
};
