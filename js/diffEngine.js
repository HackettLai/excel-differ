// diffEngine.js

class DiffEngine {
  constructor() {
    this.keyColumns = null;
  }

  /**
   * Compare two workbooks
   */
  compare(workbookA, workbookB) {
    console.log('Starting workbook comparison');

    const result = {
      sheets: {},
      summary: {
        totalSheets: 0,
        modifiedSheets: 0,
        addedSheets: 0,
        removedSheets: 0,
      },
    };

    const allSheetNames = new Set([
      ...workbookA.sheetNames,
      ...workbookB.sheetNames,
    ]);

    allSheetNames.forEach((sheetName) => {
      const sheetA = workbookA.sheets[sheetName];
      const sheetB = workbookB.sheets[sheetName];

      if (!sheetA) {
        result.sheets[sheetName] = {
          status: 'added',
          rowDiff: [],
        };
        result.summary.addedSheets++;
      } else if (!sheetB) {
        result.sheets[sheetName] = {
          status: 'removed',
          rowDiff: [],
        };
        result.summary.removedSheets++;
      } else {
        result.sheets[sheetName] = this.compareSheets(sheetA, sheetB);
        if (result.sheets[sheetName].hasChanges) {
          result.summary.modifiedSheets++;
        }
      }

      result.summary.totalSheets++;
    });

    console.log('Comparison summary:', result.summary);
    return result;
  }

  /**
   * Compare two sheets with intelligent row matching
   */
  compareSheets(sheetA, sheetB) {
    console.log('Comparing sheets with intelligent matching');

  // ✨ 轉換為陣列
  const rowsA = this.normalizeSheet(sheetA);
  const rowsB = this.normalizeSheet(sheetB);

  // ✅ Debug: 檢查資料結構
  console.log('📊 Sheet A structure:', {
    totalRows: rowsA.length,
    firstRow: rowsA[0],
    firstRowKeys: rowsA[0] ? Object.keys(rowsA[0]) : [],
    secondRow: rowsA[1]
  });

  console.log('📊 Sheet B structure:', {
    totalRows: rowsB.length,
    firstRow: rowsB[0],
    firstRowKeys: rowsB[0] ? Object.keys(rowsB[0]) : [],
    secondRow: rowsB[1]
  });

  console.log('Rows A:', rowsA.length, 'Rows B:', rowsB.length);

  // Find best key columns
  this.keyColumns = this.findBestKeyColumns(rowsA, rowsB);
  console.log('Selected key columns:', this.keyColumns);

    // Detect column changes
    const columnDiff = this.compareColumns(rowsA, rowsB);

    // ✨ 使用新的智能匹配算法
    const rowDiff = this.intelligentMatchRows(rowsA, rowsB);

    const hasChanges = rowDiff.some((row) => row.type !== 'unchanged');

    return {
      rowDiff,
      columnDiff,
      hasChanges,
      keyColumns: this.keyColumns,
    };
  }

  /**
   * ✨ 智能匹配算法：結合 index 和 key
   */
  intelligentMatchRows(rowsA, rowsB) {
    const result = [];
    const maxLen = Math.max(rowsA.length, rowsB.length);

    // 如果有好的 key columns，建立 key map
    let useKeyMatching = false;
    let keyMapA = new Map();
    let keyMapB = new Map();

    if (this.keyColumns && this.keyColumns.length > 0) {
      // 檢查 key 的唯一性
      const keysA = rowsA.map((row, idx) => ({
        key: this.getRowKey(row, this.keyColumns),
        index: idx,
        row
      }));
      
      const keysB = rowsB.map((row, idx) => ({
        key: this.getRowKey(row, this.keyColumns),
        index: idx,
        row
      }));

      const uniqueKeysA = new Set(keysA.map(k => k.key));
      const uniqueKeysB = new Set(keysB.map(k => k.key));

      // 如果唯一性 > 90%，使用 key matching
      const uniquenessA = uniqueKeysA.size / keysA.length;
      const uniquenessB = uniqueKeysB.size / keysB.length;

      if (uniquenessA > 0.9 && uniquenessB > 0.9) {
        useKeyMatching = true;
        console.log('Using key-based matching');

        keysA.forEach(item => {
          if (!keyMapA.has(item.key)) {
            keyMapA.set(item.key, []);
          }
          keyMapA.get(item.key).push(item);
        });

        keysB.forEach(item => {
          if (!keyMapB.has(item.key)) {
            keyMapB.set(item.key, []);
          }
          keyMapB.get(item.key).push(item);
        });
      }
    }

    if (useKeyMatching) {
      // ✨ Key-based matching
      const allKeys = new Set([...keyMapA.keys(), ...keyMapB.keys()]);
      const processedA = new Set();
      const processedB = new Set();

      allKeys.forEach(key => {
        const itemsA = keyMapA.get(key) || [];
        const itemsB = keyMapB.get(key) || [];

        const maxItems = Math.max(itemsA.length, itemsB.length);

        for (let i = 0; i < maxItems; i++) {
          const itemA = itemsA[i];
          const itemB = itemsB[i];

          if (itemA && itemB) {
            // Both exist - compare
            const cellDiff = this.compareCells(itemA.row, itemB.row);
            const hasChanges = Object.values(cellDiff).some(cell => cell.changed);

            result.push({
              type: hasChanges ? 'modified' : 'unchanged',
              oldIndex: itemA.index + 1,
              newIndex: itemB.index + 1,
              cells: cellDiff,
              key: key,
            });

            processedA.add(itemA.index);
            processedB.add(itemB.index);

          } else if (itemA) {
            // Removed
            result.push({
              type: 'removed',
              oldIndex: itemA.index + 1,
              newIndex: null,
              cells: this.createCellsFromRow(itemA.row, 'removed'),
              key: key,
            });
            processedA.add(itemA.index);

          } else if (itemB) {
            // Added
            result.push({
              type: 'added',
              oldIndex: null,
              newIndex: itemB.index + 1,
              cells: this.createCellsFromRow(itemB.row, 'added'),
              key: key,
            });
            processedB.add(itemB.index);
          }
        }
      });

      // Sort by index
      result.sort((a, b) => {
        if (a.oldIndex !== null && b.oldIndex !== null) {
          return a.oldIndex - b.oldIndex;
        }
        if (a.newIndex !== null && b.newIndex !== null) {
          return a.newIndex - b.newIndex;
        }
        return 0;
      });

    } else {
      // ✨ Index-based matching (fallback)
      console.log('Using index-based matching');

      for (let i = 0; i < maxLen; i++) {
        const rowA = rowsA[i];
        const rowB = rowsB[i];

        if (rowA && rowB) {
          // Both rows exist - compare
          const cellDiff = this.compareCells(rowA, rowB);
          const hasChanges = Object.values(cellDiff).some(cell => cell.changed);

          result.push({
            type: hasChanges ? 'modified' : 'unchanged',
            oldIndex: i + 1,
            newIndex: i + 1,
            cells: cellDiff,
          });

        } else if (rowA && !rowB) {
          // Row removed
          result.push({
            type: 'removed',
            oldIndex: i + 1,
            newIndex: null,
            cells: this.createCellsFromRow(rowA, 'removed'),
          });

        } else if (!rowA && rowB) {
          // Row added
          result.push({
            type: 'added',
            oldIndex: null,
            newIndex: i + 1,
            cells: this.createCellsFromRow(rowB, 'added'),
          });
        }
      }
    }

    return result;
  }

  /**
   * ✨ 生成行的 key
   */
  getRowKey(row, keyColumns) {
    if (!keyColumns || keyColumns.length === 0) {
      return '';
    }
    return keyColumns.map(col => (row[col] ?? '')).join('|');
  }

  /**
   * ✨ 將 sheet 轉換為標準陣列格式
   */
/**
 * ✅ 修正：正確處理 sheet 的資料結構
 */
normalizeSheet(sheet) {
  if (!sheet) {
    console.warn('normalizeSheet: sheet is null or undefined');
    return [];
  }
  
  // ✅ 如果傳入的是 sheet object (有 data 屬性)
  if (sheet.data && Array.isArray(sheet.data)) {
    console.log('✅ Found sheet.data array, length:', sheet.data.length);
    return sheet.data.filter(row => row && typeof row === 'object');
  }
  
  // 如果已經是陣列，直接返回
  if (Array.isArray(sheet)) {
    console.log('✅ Sheet is already an array, length:', sheet.length);
    return sheet.filter(row => row && typeof row === 'object');
  }
  
  // 如果是物件但沒有 data 屬性，嘗試轉換
  if (typeof sheet === 'object') {
    console.log('⚠️ Sheet is object without data property');
    return Object.values(sheet).filter(row => row && typeof row === 'object');
  }
  
  console.warn('❌ Unable to normalize sheet:', sheet);
  return [];
}

  /**
   * Find best key columns for row matching
   */
  findBestKeyColumns(rowsA, rowsB) {
    if (rowsA.length === 0 || rowsB.length === 0) {
      console.warn('Empty sheets provided');
      return [];
    }

    // Get all column names (from first row)
    const firstRowA = rowsA[0];
    const firstRowB = rowsB[0];

    if (!firstRowA || !firstRowB) {
      console.warn('No valid first row');
      return [];
    }

    const columnsA = Object.keys(firstRowA);
    const columnsB = Object.keys(firstRowB);
    const commonColumns = columnsA.filter((col) => columnsB.includes(col));

    if (commonColumns.length === 0) {
      console.warn('No common columns found');
      return [];
    }

    // Score each column based on uniqueness
    const columnScores = commonColumns.map((col) => {
      const uniquenessA = this.calculateUniqueness(rowsA, col);
      const uniquenessB = this.calculateUniqueness(rowsB, col);
      const avgUniqueness = (uniquenessA + uniquenessB) / 2;

      return {
        column: col,
        score: avgUniqueness,
      };
    });

    // Sort by score (higher is better)
    columnScores.sort((a, b) => b.score - a.score);

    console.log('Column scores:', columnScores.slice(0, 5));

    // If top column has high uniqueness (> 90%), use it alone
    if (columnScores[0].score > 0.9) {
      return [columnScores[0].column];
    }

    // Otherwise, try combination of top 2-3 columns
    const topColumns = columnScores.slice(0, Math.min(3, columnScores.length)).map((c) => c.column);
    const combinedUniqueness = this.calculateCombinedUniqueness(
      rowsA,
      rowsB,
      topColumns
    );

    console.log('Combined uniqueness:', combinedUniqueness);

    if (combinedUniqueness > 0.95) {
      return topColumns;
    }

    // ✨ 如果沒有好的 key，返回空陣列（使用 index matching）
    return [];
  }

  /**
   * Calculate uniqueness of a column (0-1, higher is better)
   */
  calculateUniqueness(rows, column) {
    const values = rows
      .map((row) => row[column])
      .filter((v) => v != null && v !== '');
    
    if (values.length === 0) return 0;
    
    const uniqueValues = new Set(values);
    return uniqueValues.size / values.length;
  }

  /**
   * Calculate combined uniqueness of multiple columns
   */
  calculateCombinedUniqueness(rowsA, rowsB, columns) {
    const keysA = rowsA
      .map((row) => columns.map((col) => row[col] ?? '').join('|'))
      .filter(key => key.trim() !== '' && !key.split('|').every(v => v === ''));
    
    const keysB = rowsB
      .map((row) => columns.map((col) => row[col] ?? '').join('|'))
      .filter(key => key.trim() !== '' && !key.split('|').every(v => v === ''));

    if (keysA.length === 0 || keysB.length === 0) return 0;

    const uniqueA = new Set(keysA).size / keysA.length;
    const uniqueB = new Set(keysB).size / keysB.length;

    return (uniqueA + uniqueB) / 2;
  }

  /**
   * Compare columns between two sheets
   */
  compareColumns(rowsA, rowsB) {
    const firstRowA = rowsA.find(row => row && typeof row === 'object');
    const firstRowB = rowsB.find(row => row && typeof row === 'object');

    const columnsA = new Set(firstRowA ? Object.keys(firstRowA) : []);
    const columnsB = new Set(firstRowB ? Object.keys(firstRowB) : []);

    const added = Array.from(columnsB).filter((col) => !columnsA.has(col));
    const removed = Array.from(columnsA).filter((col) => !columnsB.has(col));
    const common = Array.from(columnsA).filter((col) => columnsB.has(col));

    return { added, removed, common };
  }

  /**
   * Compare cells in two rows
   */
  compareCells(rowA, rowB) {
    const allKeys = new Set([...Object.keys(rowA), ...Object.keys(rowB)]);
    const cells = {};

    allKeys.forEach((key) => {
      const valA = rowA[key];
      const valB = rowB[key];

      if (valA !== valB) {
        cells[key] = {
          changed: true,
          oldValue: valA,
          newValue: valB,
        };
      } else {
        cells[key] = {
          changed: false,
          value: valA,
        };
      }
    });

    return cells;
  }

  /**
   * Create cells object from a single row
   */
  createCellsFromRow(row, type) {
    const cells = {};
    Object.keys(row).forEach((key) => {
      cells[key] = {
        value: row[key],
        changed: false,
      };
    });
    return cells;
  }
}

export default DiffEngine;