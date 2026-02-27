/**
 * diffEngine.js
 * Excel comparison engine - core logic for comparing two Excel workbooks
 * Detects differences in sheets, rows, columns, and cell values
 */

/**
 * DiffEngine Class
 * Handles all comparison logic between two parsed Excel workbooks
 * Generates comprehensive diff results including added/deleted/modified sheets and cells
 */
/**
 * diffEngine.js - CORRECTED VERSION
 */

class DiffEngine {
  constructor() {
    this.results = {
      summary: {
        totalSheets: 0,
        modifiedSheets: 0,
        addedSheets: 0,
        deletedSheets: 0,
      },
      sheets: [],
    };
  }

  compare(dataA, dataB) {
    console.log('🔍 Starting comparison...', { dataA, dataB });

    if (!dataA || !dataB) {
      console.error('❌ Missing comparison data');
      return this.results;
    }

    const sheetsA = new Set(dataA.sheetNames || []);
    const sheetsB = new Set(dataB.sheetNames || []);
    const allSheets = new Set([...sheetsA, ...sheetsB]);

    allSheets.forEach((sheetName) => {
      const inA = sheetsA.has(sheetName);
      const inB = sheetsB.has(sheetName);

      let sheetResult = {
        sheetName: sheetName,
        status: 'unchanged',
        differences: [],
        rowChanges: [],
        columnChanges: [],
        oldData: [],
        newData: [],
      };

      if (inA && inB) {
        const sheetA = dataA.sheets[sheetName]?.data || [];
        const sheetB = dataB.sheets[sheetName]?.data || [];

        sheetResult = this.compareSheets(sheetA, sheetB);
        sheetResult.sheetName = sheetName;

        if (sheetResult.differences.length > 0 || sheetResult.rowChanges.length > 0 || sheetResult.columnChanges.length > 0) {
          sheetResult.status = 'modified';
          this.results.summary.modifiedSheets++;
        }
      } else if (inB && !inA) {
        sheetResult.status = 'added';
        sheetResult.newData = dataB.sheets[sheetName]?.data || [];
        this.results.summary.addedSheets++;
      } else if (inA && !inB) {
        sheetResult.status = 'deleted';
        sheetResult.oldData = dataA.sheets[sheetName]?.data || [];
        this.results.summary.deletedSheets++;
      }

      this.results.sheets.push(sheetResult);
    });

    this.results.summary.totalSheets = allSheets.size;

    console.log('✅ Comparison complete', this.results);
    return this.results;
  }

  /**
   * compareSheets(sheetA, sheetB, keyColumn, headerRowA, headerRowB)
   * Compares two sheets and identifies all differences
   * Detects column changes, row changes, and cell differences
   *
   * @param {Array<Object>} sheetA - Data from first sheet
   * @param {Array<Object>} sheetB - Data from second sheet
   * @param {string} keyColumn - Column to use for row matching
   * @param {number} headerRowA - Header row index for sheet A (1-based)
   * @param {number} headerRowB - Header row index for sheet B (1-based)
   * @returns {Object} Comparison results including differences, row changes, and column changes
   */
  compareSheets(sheetA, sheetB, keyColumn = null, headerRowA = 1, headerRowB = 1) {
    if (!sheetA || !sheetB) {
      return {
        differences: [],
        rowChanges: [],
        columnChanges: [],
        oldData: sheetA || [],
        newData: sheetB || [],
      };
    }

    const headerA = sheetA[0] || {};
    const headerB = sheetB[0] || {};

    const columnChanges = this.detectColumnChanges(sheetA, sheetB);

    const rowChanges = keyColumn ? this.detectRowChangesByKey(sheetA.slice(1), sheetB.slice(1), keyColumn, headerRowA, headerRowB) : this.detectRowChanges(sheetA, sheetB, headerRowA, headerRowB);

    const differences = keyColumn ? this.compareCellsByKey(sheetA.slice(1), sheetB.slice(1), keyColumn, headerA, headerB, headerRowA, headerRowB) : this.compareCells(sheetA, sheetB, headerRowA, headerRowB);

    return {
      differences: differences,
      rowChanges: rowChanges,
      columnChanges: columnChanges,
      oldData: sheetA,
      newData: sheetB,
    };
  }

  /**
   * detectRowChangesByKey(rowsA, rowsB, keyColumn, headerRowA, headerRowB)
   * Detects added and deleted rows using a specified key column for matching
   * Rows are matched by the value in the key column, not by position
   *
   * @param {Array<Object>} rowsA - Data rows from sheet A
   * @param {Array<Object>} rowsB - Data rows from sheet B
   * @param {string} keyColumn - Column letter to use for row matching
   * @param {number} headerRowA - Header row index for sheet A (1-based)
   * @param {number} headerRowB - Header row index for sheet B (1-based)
   * @returns {Array<Object>} Array of row changes
   */
  detectRowChangesByKey(rowsA, rowsB, keyColumn, headerRowA = 1, headerRowB = 1) {
    const changes = [];
    const mapA = new Map();
    const mapB = new Map();

    // Build map of File A rows using key column values
    rowsA.forEach((row, index) => {
      const keyValue = row[keyColumn];
      if (keyValue !== undefined && keyValue !== null && keyValue !== '') {
        mapA.set(String(keyValue).trim(), {
          row,
          index: headerRowA + 1 + index, // Excel row number calculation
        });
      }
    });

    // Build map of File B rows using key column values
    rowsB.forEach((row, index) => {
      const keyValue = row[keyColumn];
      if (keyValue !== undefined && keyValue !== null && keyValue !== '') {
        mapB.set(String(keyValue).trim(), {
          row,
          index: headerRowB + 1 + index, // Excel row number calculation
        });
      }
    });

    // Find added rows
    mapB.forEach((data, key) => {
      if (!mapA.has(key)) {
        changes.push({
          rowKey: key,
          type: 'added',
          newRowIndex: data.index,
          row: data.row,
        });
      }
    });

    // Find deleted rows
    mapA.forEach((data, key) => {
      if (!mapB.has(key)) {
        changes.push({
          rowKey: key,
          type: 'deleted',
          oldRowIndex: data.index,
          row: data.row,
        });
      }
    });

    console.log(`🔑 Detected ${changes.length} row changes using key column ${keyColumn}`);
    return changes;
  }

  /**
   * compareCellsByKey(rowsA, rowsB, keyColumn, headerA, headerB, headerRowA, headerRowB)
   * Compares cell values between matched rows using key column
   * Skips the key column itself and compares all other columns
   *
   * @param {Array<Object>} rowsA - Data rows from sheet A
   * @param {Array<Object>} rowsB - Data rows from sheet B
   * @param {string} keyColumn - Column letter used for row matching
   * @param {Object} headerA - Header row from sheet A
   * @param {Object} headerB - Header row from sheet B
   * @param {number} headerRowA - Header row index for sheet A (1-based)
   * @param {number} headerRowB - Header row index for sheet B (1-based)
   * @returns {Array<Object>} Array of cell differences
   */
  compareCellsByKey(rowsA, rowsB, keyColumn, headerA, headerB, headerRowA = 1, headerRowB = 1) {
    const differences = [];

    // Build header mapping
    const headerToOldCol = new Map();
    const headerToNewCol = new Map();

    Object.keys(headerA).forEach((col) => {
      const content = String(headerA[col] || '').trim();
      if (content) {
        headerToOldCol.set(content, col);
      }
    });

    Object.keys(headerB).forEach((col) => {
      const content = String(headerB[col] || '').trim();
      if (content) {
        headerToNewCol.set(content, col);
      }
    });

    // Normalize key column name for consistent comparison
    const normalizedKeyColumn = String(keyColumn || '')
      .trim()
      .toLowerCase();

    // Create maps using key column values
    const mapA = new Map();
    const mapB = new Map();

    rowsA.forEach((row, index) => {
      const keyValue = row[keyColumn];
      if (keyValue !== undefined && keyValue !== null && keyValue !== '') {
        mapA.set(String(keyValue).trim(), {
          row,
          index: headerRowA + 1 + index,
        });
      }
    });

    rowsB.forEach((row, index) => {
      const keyValue = row[keyColumn];
      if (keyValue !== undefined && keyValue !== null && keyValue !== '') {
        mapB.set(String(keyValue).trim(), {
          row,
          index: headerRowB + 1 + index,
        });
      }
    });

    // Compare cells for matching rows
    mapA.forEach((oldRowData, key) => {
      if (mapB.has(key)) {
        const newRowData = mapB.get(key);
        const oldRow = oldRowData.row;
        const newRow = newRowData.row;

        headerToOldCol.forEach((oldCol, headerContent) => {
          // Skip the key column from comparison
          const normalizedHeader = String(headerContent || '')
            .trim()
            .toLowerCase();
          if (normalizedHeader === normalizedKeyColumn) {
            console.log(`⏭️  Skipping key column "${headerContent}" in comparison`);
            return; // Skip this iteration
          }

          const newCol = headerToNewCol.get(headerContent);

          if (newCol) {
            const oldVal = oldRow[oldCol];
            const newVal = newRow[newCol];

            if (oldVal !== newVal) {
              differences.push({
                row: oldRowData.index,
                header: headerContent,
                oldCol: oldCol,
                newCol: newCol,
                oldValue: oldVal,
                newValue: newVal,
              });
            }
          }
        });
      }
    });

    console.log(`🔑 Found ${differences.length} cell differences using key column ${keyColumn}`);
    return differences;
  }

  /**
   * detectRowChanges(oldData, newData, headerRowA, headerRowB)
   * Detects added and deleted rows when no key column is specified
   * Matches rows by position or first column value
   *
   * @param {Array<Object>} oldData - Data from sheet A
   * @param {Array<Object>} newData - Data from sheet B
   * @param {number} headerRowA - Header row index for sheet A (1-based)
   * @param {number} headerRowB - Header row index for sheet B (1-based)
   * @returns {Array<Object>} Array of row changes
   */
  detectRowChanges(oldData, newData, headerRowA = 1, headerRowB = 1) {
    const changes = [];

    if (!oldData || !newData) return changes;

    const oldRows = oldData.slice(1);
    const newRows = newData.slice(1);

    const oldRowMap = new Map();
    const newRowMap = new Map();

    oldRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `old-${index}`;
      oldRowMap.set(key, {
        row,
        index: headerRowA + 1 + index, // ✅ CORRECT
      });
    });

    newRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `new-${index}`;
      newRowMap.set(key, {
        row,
        index: headerRowB + 1 + index, // ✅ CORRECT
      });
    });

    newRowMap.forEach((data, key) => {
      if (!oldRowMap.has(key)) {
        changes.push({
          rowKey: key,
          type: 'added',
          newRowIndex: data.index,
          row: data.row,
        });
      }
    });

    oldRowMap.forEach((data, key) => {
      if (!newRowMap.has(key)) {
        changes.push({
          rowKey: key,
          type: 'deleted',
          oldRowIndex: data.index,
          row: data.row,
        });
      }
    });

    return changes;
  }

  /**
   * compareCells(oldData, newData, headerRowA, headerRowB)
   * Compares cell values between matched rows when no key column is specified
   * Matches rows by position or first column value
   *
   * @param {Array<Object>} oldData - Data from sheet A
   * @param {Array<Object>} newData - Data from sheet B
   * @param {number} headerRowA - Header row index for sheet A (1-based)
   * @param {number} headerRowB - Header row index for sheet B (1-based)
   * @returns {Array<Object>} Array of cell differences
   */
  compareCells(oldData, newData, headerRowA = 1, headerRowB = 1) {
    const differences = [];

    if (!oldData || !newData) return differences;

    const oldHeaders = oldData[0] || {};
    const newHeaders = newData[0] || {};

    const headerToOldCol = new Map();
    const headerToNewCol = new Map();

    Object.keys(oldHeaders).forEach((col) => {
      const content = String(oldHeaders[col] || '').trim();
      if (content) {
        headerToOldCol.set(content, col);
      }
    });

    Object.keys(newHeaders).forEach((col) => {
      const content = String(newHeaders[col] || '').trim();
      if (content) {
        headerToNewCol.set(content, col);
      }
    });

    const oldRows = oldData.slice(1);
    const newRows = newData.slice(1);

    const oldRowMap = new Map();
    const newRowMap = new Map();

    oldRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `old-${index}`;
      oldRowMap.set(key, {
        row,
        index: headerRowA + 1 + index, // ✅ CORRECT
      });
    });

    newRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `new-${index}`;
      newRowMap.set(key, {
        row,
        index: headerRowB + 1 + index, // ✅ CORRECT
      });
    });

    oldRowMap.forEach((oldRowData, key) => {
      if (newRowMap.has(key)) {
        const newRowData = newRowMap.get(key);
        const oldRow = oldRowData.row;
        const newRow = newRowData.row;

        headerToOldCol.forEach((oldCol, headerContent) => {
          const newCol = headerToNewCol.get(headerContent);

          if (newCol) {
            const oldVal = oldRow[oldCol];
            const newVal = newRow[newCol];

            if (oldVal !== newVal) {
              differences.push({
                row: oldRowData.index, // ✅ Now correct
                header: headerContent,
                oldCol: oldCol,
                newCol: newCol,
                oldValue: oldVal,
                newValue: newVal,
              });
            }
          }
        });
      }
    });

    return differences;
  }

  /**
   * detectColumnChanges(oldData, newData)
   * Detects added and deleted columns between two sheets
   *
   * @param {Array<Object>} oldData - Data from sheet A
   * @param {Array<Object>} newData - Data from sheet B
   * @returns {Array<Object>} Array of column changes
   */
  detectColumnChanges(oldData, newData) {
    if (!oldData || !newData || oldData.length === 0 || newData.length === 0) {
      return [];
    }

    const oldHeaders = oldData[0];
    const newHeaders = newData[0];

    const oldHeaderSet = new Set();
    const newHeaderSet = new Set();

    Object.values(oldHeaders).forEach((val) => {
      const content = String(val || '').trim();
      if (content) oldHeaderSet.add(content);
    });

    Object.values(newHeaders).forEach((val) => {
      const content = String(val || '').trim();
      if (content) newHeaderSet.add(content);
    });

    const changes = [];
    const oldCols = Object.keys(oldHeaders);
    const newCols = Object.keys(newHeaders);

    newCols.forEach((col) => {
      const newContent = String(newHeaders[col] || '').trim();

      if (newContent) {
        if (!oldHeaderSet.has(newContent)) {
          changes.push({
            column: col,
            type: 'added',
            header: newContent,
          });
        }
      } else {
        if (!oldCols.includes(col)) {
          changes.push({
            column: col,
            type: 'added',
            header: '(Blank Column)',
          });
        }
      }
    });

    oldCols.forEach((col) => {
      const oldContent = String(oldHeaders[col] || '').trim();

      if (oldContent) {
        if (!newHeaderSet.has(oldContent)) {
          changes.push({
            column: col,
            type: 'deleted',
            header: oldContent,
          });
        }
      } else {
        if (!newCols.includes(col)) {
          changes.push({
            column: col,
            type: 'deleted',
            header: '(Blank Column)',
          });
        }
      }
    });

    return changes;
  }
}

export default DiffEngine;
