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
   * ✅ FIXED: Accept headerRowA and headerRowB
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

    const rowChanges = keyColumn
      ? this.detectRowChangesByKey(sheetA.slice(1), sheetB.slice(1), keyColumn, headerRowA, headerRowB)
      : this.detectRowChanges(sheetA, sheetB, headerRowA, headerRowB);

    const differences = keyColumn
      ? this.compareCellsByKey(sheetA.slice(1), sheetB.slice(1), keyColumn, headerA, headerB, headerRowA, headerRowB)
      : this.compareCells(sheetA, sheetB, headerRowA, headerRowB);

    return {
      differences: differences,
      rowChanges: rowChanges,
      columnChanges: columnChanges,
      oldData: sheetA,
      newData: sheetB,
    };
  }

  /**
   * ✅ FIXED VERSION - 刪除舊版，只保留呢個
   */
  detectRowChangesByKey(rowsA, rowsB, keyColumn, headerRowA = 1, headerRowB = 1) {
    const changes = [];
    const mapA = new Map();
    const mapB = new Map();

    // ✅ CORRECT: Use headerRowA + 1 + index
    rowsA.forEach((row, index) => {
      const keyValue = row[keyColumn];
      if (keyValue !== undefined && keyValue !== null && keyValue !== '') {
        mapA.set(String(keyValue).trim(), {
          row,
          index: headerRowA + 1 + index, // ✅ Excel row = header + 1 + data index
        });
      }
    });

    // ✅ CORRECT: Use headerRowB + 1 + index
    rowsB.forEach((row, index) => {
      const keyValue = row[keyColumn];
      if (keyValue !== undefined && keyValue !== null && keyValue !== '') {
        mapB.set(String(keyValue).trim(), {
          row,
          index: headerRowB + 1 + index, // ✅ Excel row = header + 1 + data index
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
   * ✅ FIXED VERSION - 刪除舊版，只保留呢個
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

    // Create maps using key column values
    const mapA = new Map();
    const mapB = new Map();

    // ✅ CORRECT: Use headerRowA + 1 + index
    rowsA.forEach((row, index) => {
      const keyValue = row[keyColumn];
      if (keyValue !== undefined && keyValue !== null && keyValue !== '') {
        mapA.set(String(keyValue).trim(), {
          row,
          index: headerRowA + 1 + index, // ✅ Excel row number
        });
      }
    });

    // ✅ CORRECT: Use headerRowB + 1 + index
    rowsB.forEach((row, index) => {
      const keyValue = row[keyColumn];
      if (keyValue !== undefined && keyValue !== null && keyValue !== '') {
        mapB.set(String(keyValue).trim(), {
          row,
          index: headerRowB + 1 + index, // ✅ Excel row number
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
          const newCol = headerToNewCol.get(headerContent);

          if (newCol) {
            const oldVal = oldRow[oldCol];
            const newVal = newRow[newCol];

            if (oldVal !== newVal) {
              differences.push({
                row: oldRowData.index, // ✅ Now correct Excel row number
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
   * ✅ FIXED: detectRowChanges (non-key mode)
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
   * ✅ FIXED: compareCells (non-key mode)
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
   * detectColumnChanges - 呢個唔需要改
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