/**
 * diffEngine.js
 * Excel 比對核心引擎
 */

class DiffEngine {
  constructor() {
    this.results = {
      summary: {
        totalSheets: 0,
        modifiedSheets: 0,
        addedSheets: 0,
        deletedSheets: 0
      },
      sheets: []
    };
  }

  /**
   * ✅ 主要比對入口
   */
  compare(dataA, dataB) {
    console.log('🔍 開始比對...', { dataA, dataB });

    if (!dataA || !dataB) {
      console.error('❌ 缺少比對資料');
      return this.results;
    }

    const sheetsA = new Set(dataA.sheetNames || []);
    const sheetsB = new Set(dataB.sheetNames || []);

    // ✅ 找出所有 sheet 名稱
    const allSheets = new Set([...sheetsA, ...sheetsB]);

    allSheets.forEach(sheetName => {
      const inA = sheetsA.has(sheetName);
      const inB = sheetsB.has(sheetName);

      let sheetResult = {
        sheetName: sheetName,
        status: 'unchanged',
        differences: [],
        rowChanges: [],
        columnChanges: [],
        oldData: [],
        newData: []
      };

      if (inA && inB) {
        // 兩個檔案都有呢個 sheet，進行比對
        const sheetA = dataA.sheets[sheetName]?.data || [];
        const sheetB = dataB.sheets[sheetName]?.data || [];

        sheetResult = this.compareSheets(sheetA, sheetB);
        sheetResult.sheetName = sheetName;

        if (sheetResult.differences.length > 0 || 
            sheetResult.rowChanges.length > 0 || 
            sheetResult.columnChanges.length > 0) {
          sheetResult.status = 'modified';
          this.results.summary.modifiedSheets++;
        }
      } else if (inB && !inA) {
        // File B 新增嘅 sheet
        sheetResult.status = 'added';
        sheetResult.newData = dataB.sheets[sheetName]?.data || [];
        this.results.summary.addedSheets++;
      } else if (inA && !inB) {
        // File A 有但 File B 冇（被刪除）
        sheetResult.status = 'deleted';
        sheetResult.oldData = dataA.sheets[sheetName]?.data || [];
        this.results.summary.deletedSheets++;
      }

      this.results.sheets.push(sheetResult);
    });

    this.results.summary.totalSheets = allSheets.size;

    console.log('✅ 比對完成', this.results);
    return this.results;
  }

  /**
   * ✅ 比對單個 Sheet
   */
  compareSheets(oldData, newData) {
    const result = {
      differences: [],
      rowChanges: [],
      columnChanges: [],
      oldData: oldData,
      newData: newData
    };

    if (!oldData || !newData || oldData.length === 0 || newData.length === 0) {
      return result;
    }

    // ✅ 偵測欄位變更（只標記新增/刪除）
    result.columnChanges = this.detectColumnChanges(oldData, newData);

    // ✅ 偵測行變更
    result.rowChanges = this.detectRowChanges(oldData, newData);

    // ✅ 比對儲存格內容
    result.differences = this.compareCells(oldData, newData);

    return result;
  }

  /**
   * 🔥 只偵測新增/刪除欄位，忽略移位
   */
  detectColumnChanges(oldData, newData) {
    if (!oldData || !newData || oldData.length === 0 || newData.length === 0) {
      return [];
    }

    const oldHeaders = oldData[0];  // 第1行
    const newHeaders = newData[0];  // 第1行

    // ✅ 建立 Header 內容 Set（忽略 null/empty）
    const oldHeaderSet = new Set();
    const newHeaderSet = new Set();

    Object.values(oldHeaders).forEach(val => {
      const content = String(val || '').trim();
      if (content) oldHeaderSet.add(content);
    });

    Object.values(newHeaders).forEach(val => {
      const content = String(val || '').trim();
      if (content) newHeaderSet.add(content);
    });

    const changes = [];

    // ✅ 找出新增的欄位（File B 有但 File A 冇的 column letter）
    const oldCols = Object.keys(oldHeaders);
    const newCols = Object.keys(newHeaders);

    newCols.forEach(col => {
      const newContent = String(newHeaders[col] || '').trim();
      
      // 只標記有內容的新欄位
      if (newContent) {
        // 檢查呢個 header 係咪真係新增（File A 完全冇呢個 header）
        if (!oldHeaderSet.has(newContent)) {
          changes.push({
            column: col,
            type: 'added',
            header: newContent
          });
        }
      } else {
        // 檢查呢個空欄係咪真係新增（File A 冇呢個 column letter）
        if (!oldCols.includes(col)) {
          changes.push({
            column: col,
            type: 'added',
            header: '(Blank Column)'
          });
        }
      }
    });

    // ✅ 找出刪除的欄位（File A 有但 File B 冇的 header）
    oldCols.forEach(col => {
      const oldContent = String(oldHeaders[col] || '').trim();
      
      if (oldContent) {
        // 檢查呢個 header 係咪真係刪除（File B 完全冇呢個 header）
        if (!newHeaderSet.has(oldContent)) {
          changes.push({
            column: col,
            type: 'deleted',
            header: oldContent
          });
        }
      } else {
        // 檢查呢個空欄係咪真係刪除（File B 冇呢個 column letter）
        if (!newCols.includes(col)) {
          changes.push({
            column: col,
            type: 'deleted',
            header: '(Blank Column)'
          });
        }
      }
    });

    console.log('📊 Column Changes:', changes);
    return changes;
  }

  /**
   * ✅ 偵測行變更（用第1欄做 key）
   */
  detectRowChanges(oldData, newData) {
    const changes = [];

    if (!oldData || !newData) return changes;

    // 跳過第1行（header）
    const oldRows = oldData.slice(1);
    const newRows = newData.slice(1);

    const oldRowMap = new Map();
    const newRowMap = new Map();

    // 用 A 欄做 key
    oldRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `old-${index}`;
      oldRowMap.set(key, { row, index: index + 2 });
    });

    newRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `new-${index}`;
      newRowMap.set(key, { row, index: index + 2 });
    });

    // 找出新增的行
    newRowMap.forEach((data, key) => {
      if (!oldRowMap.has(key)) {
        changes.push({
          rowKey: key,
          type: 'added',
          newRowIndex: data.index,
          row: data.row
        });
      }
    });

    // 找出刪除的行
    oldRowMap.forEach((data, key) => {
      if (!newRowMap.has(key)) {
        changes.push({
          rowKey: key,
          type: 'deleted',
          oldRowIndex: data.index,
          row: data.row
        });
      }
    });

    return changes;
  }

  /**
   * ✅ 比對儲存格內容
   */
  compareCells(oldData, newData) {
  const differences = [];

  if (!oldData || !newData) return differences;

  // ✅ 建立 header mapping（header content → column letters）
  const oldHeaders = oldData[0] || {};
  const newHeaders = newData[0] || {};
  
  const headerToOldCol = new Map();  // "Email Address" → "G"
  const headerToNewCol = new Map();  // "Email Address" → "H"

  Object.keys(oldHeaders).forEach(col => {
    const content = String(oldHeaders[col] || '').trim();
    if (content) {
      headerToOldCol.set(content, col);
    }
  });

  Object.keys(newHeaders).forEach(col => {
    const content = String(newHeaders[col] || '').trim();
    if (content) {
      headerToNewCol.set(content, col);
    }
  });

  // ✅ 建立 row mapping
  const oldRows = oldData.slice(1);
  const newRows = newData.slice(1);

  const oldRowMap = new Map();
  const newRowMap = new Map();

  oldRows.forEach((row, index) => {
    const key = String(row.A || '').trim() || `old-${index}`;
    oldRowMap.set(key, { row, index: index + 2 });
  });

  newRows.forEach((row, index) => {
    const key = String(row.A || '').trim() || `new-${index}`;
    newRowMap.set(key, { row, index: index + 2 });
  });

  // ✅ 比對相同 rowKey 的儲存格（用 header mapping）
  oldRowMap.forEach((oldRowData, key) => {
    if (newRowMap.has(key)) {
      const newRowData = newRowMap.get(key);
      const oldRow = oldRowData.row;
      const newRow = newRowData.row;

      // 🔥 用 header content 做 key，唔係 column letter
      headerToOldCol.forEach((oldCol, headerContent) => {
        const newCol = headerToNewCol.get(headerContent);
        
        if (newCol) {
          // ✅ 同一個 header，比對對應嘅 column
          const oldVal = oldRow[oldCol];
          const newVal = newRow[newCol];

          if (oldVal !== newVal) {
            differences.push({
              row: oldRowData.index,
              header: headerContent,    // ✅ 用 header content 做 key
              oldCol: oldCol,
              newCol: newCol,
              oldValue: oldVal,
              newValue: newVal
            });
          }
        }
      });
    }
  });

  return differences;
}
}

export default DiffEngine;