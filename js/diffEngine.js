// diffEngine.js - Diff 算法核心引擎

const DiffEngine = {
    
    // 主要的 diff 函數
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
                unchangedSheets: 0
            }
        };
        
        // 比較每個 sheet 的內容
        result.sheetChanges.common.forEach(sheetName => {
            const sheetA = parsedFileA.sheets.find(s => s.name === sheetName);
            const sheetB = parsedFileB.sheets.find(s => s.name === sheetName);
            
            const cellDiff = this.compareCells(sheetA, sheetB);
            result.cellDiffs[sheetName] = cellDiff;
        });
        
        // 計算統計數據
        result.statistics.totalSheets = parsedFileA.sheets.length + result.sheetChanges.added.length;
        result.statistics.addedSheets = result.sheetChanges.added.length;
        result.statistics.removedSheets = result.sheetChanges.removed.length;
        result.statistics.renamedSheets = result.sheetChanges.renamed.length;
        
        // 計算修改和未修改的 sheet 數量
        result.sheetChanges.common.forEach(sheetName => {
            const diff = result.cellDiffs[sheetName];
            if (diff.changes.length > 0) {
                result.statistics.modifiedSheets++;
            } else {
                result.statistics.unchangedSheets++;
            }
        });
        
        return result;
    },
    
    // 比較 sheet 層級
    compareSheets(sheetsA, sheetsB) {
        const namesA = sheetsA.map(s => s.name);
        const namesB = sheetsB.map(s => s.name);
        
        const setA = new Set(namesA);
        const setB = new Set(namesB);
        
        const added = namesB.filter(name => !setA.has(name));
        const removed = namesA.filter(name => !setB.has(name));
        const common = namesA.filter(name => setB.has(name));
        
        // 檢測可能的重命名
        const renamed = this.detectRenames(removed, added, sheetsA, sheetsB);
        
        // 從 added 和 removed 中移除已識別為重命名的 sheet
        const renamedOldNames = renamed.map(r => r.from);
        const renamedNewNames = renamed.map(r => r.to);
        
        const finalAdded = added.filter(name => !renamedNewNames.includes(name));
        const finalRemoved = removed.filter(name => !renamedOldNames.includes(name));
        
        return {
            added: finalAdded,
            removed: finalRemoved,
            common: common,
            renamed: renamed
        };
    },
    
    // 檢測 sheet 重命名
    detectRenames(removedNames, addedNames, sheetsA, sheetsB) {
        const renames = [];
        const threshold = 0.85; // 相似度閾值
        
        removedNames.forEach(oldName => {
            const oldSheet = sheetsA.find(s => s.name === oldName);
            
            addedNames.forEach(newName => {
                const newSheet = sheetsB.find(s => s.name === newName);
                
                const similarity = this.calculateSheetSimilarity(oldSheet, newSheet);
                
                if (similarity >= threshold) {
                    renames.push({
                        from: oldName,
                        to: newName,
                        confidence: similarity
                    });
                }
            });
        });
        
        // 只保留每個 sheet 最高相似度的匹配
        const finalRenames = [];
        const usedOld = new Set();
        const usedNew = new Set();
        
        // 按相似度排序
        renames.sort((a, b) => b.confidence - a.confidence);
        
        renames.forEach(rename => {
            if (!usedOld.has(rename.from) && !usedNew.has(rename.to)) {
                finalRenames.push(rename);
                usedOld.add(rename.from);
                usedNew.add(rename.to);
            }
        });
        
        return finalRenames;
    },
    
    // 計算兩個 sheet 的相似度
    calculateSheetSimilarity(sheetA, sheetB) {
        // 如果維度差異太大，相似度為 0
        const rowDiff = Math.abs(sheetA.rowCount - sheetB.rowCount);
        const colDiff = Math.abs(sheetA.colCount - sheetB.colCount);
        
        if (rowDiff > 10 || colDiff > 5) {
            return 0;
        }
        
        // 比較前 10 行或實際行數（取較小值）
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
    
    // 比較兩個 sheet 的 cell 內容
    compareCells(sheetA, sheetB) {
        const changes = [];
        
        // 標準化數據
        const dataA = ExcelParser.normalizeData(sheetA.data);
        const dataB = ExcelParser.normalizeData(sheetB.data);
        
        const maxRow = Math.max(dataA.length, dataB.length);
        const maxCol = Math.max(
            dataA.length > 0 ? dataA[0].length : 0,
            dataB.length > 0 ? dataB[0].length : 0
        );
        
        // 逐 cell 比較
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
                        cellRef: this.getCellReference(r, c)
                    });
                }
            }
        }
        
        return {
            changes: changes,
            totalChanges: changes.length,
            additions: changes.filter(c => c.type === 'added').length,
            deletions: changes.filter(c => c.type === 'removed').length,
            modifications: changes.filter(c => c.type === 'modified').length
        };
    },
    
    // 獲取 cell 引用 (例如: A1, B2)
    getCellReference(row, col) {
        return ExcelParser.getColumnName(col) + (row + 1);
    },
    
    // 獲取 sheet 的狀態
    getSheetStatus(sheetName, sheetChanges, cellDiffs) {
        if (sheetChanges.added.includes(sheetName)) {
            return 'added';
        }
        
        if (sheetChanges.removed.includes(sheetName)) {
            return 'removed';
        }
        
        const renamed = sheetChanges.renamed.find(r => r.from === sheetName || r.to === sheetName);
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
    
    // 創建 cell 的 diff map（用於快速查找）
    createCellDiffMap(changes) {
        const map = new Map();
        
        changes.forEach(change => {
            const key = `${change.row}-${change.col}`;
            map.set(key, change);
        });
        
        return map;
    },
    
    // 獲取指定 cell 的 diff 信息
    getCellDiff(row, col, diffMap) {
        const key = `${row}-${col}`;
        return diffMap.get(key);
    }
};