// summaryView.js - Summary 視圖渲染

const SummaryView = {
    currentDiffResult: null,
    parsedFileA: null,
    parsedFileB: null,
    
    // 顯示 summary 視圖
    show(diffResult, fileA, fileB) {
        this.currentDiffResult = diffResult;
        this.parsedFileA = fileA;
        this.parsedFileB = fileB;
        
        // 隱藏上傳區域
        document.getElementById('uploadSection').style.display = 'none';
        
        // 顯示 summary 區域
        document.getElementById('summarySection').style.display = 'block';
        
        // 填充文件信息
        document.getElementById('summaryFileA').textContent = 
            `${diffResult.fileA} (${fileA.sheetCount} sheets)`;
        document.getElementById('summaryFileB').textContent = 
            `${diffResult.fileB} (${fileB.sheetCount} sheets)`;
        
        // 渲染 sheet 變更列表
        this.renderSheetChanges();
    },
    
    // 渲染 sheet 變更列表
    renderSheetChanges() {
        const container = document.getElementById('sheetChangesList');
        container.innerHTML = '';
        
        const { sheetChanges, cellDiffs } = this.currentDiffResult;
        
        // 收集所有 sheet（包含狀態）
        const allSheets = [];
        
        // 1. Common sheets
        sheetChanges.common.forEach(sheetName => {
            const diff = cellDiffs[sheetName];
            allSheets.push({
                name: sheetName,
                status: diff.changes.length > 0 ? 'modified' : 'unchanged',
                changeCount: diff.totalChanges,
                canView: true
            });
        });
        
        // 2. Renamed sheets
        sheetChanges.renamed.forEach(rename => {
            allSheets.push({
                name: `${rename.from} → ${rename.to}`,
                originalName: rename.from,
                newName: rename.to,
                status: 'renamed',
                confidence: Math.round(rename.confidence * 100),
                canView: true
            });
        });
        
        // 3. Added sheets
        sheetChanges.added.forEach(sheetName => {
            allSheets.push({
                name: sheetName,
                status: 'added',
                canView: true,
                viewSide: 'B'
            });
        });
        
        // 4. Removed sheets
        sheetChanges.removed.forEach(sheetName => {
            allSheets.push({
                name: sheetName,
                status: 'removed',
                canView: true,
                viewSide: 'A'
            });
        });
        
        // 渲染每個 sheet item
        allSheets.forEach(sheet => {
            const item = this.createSheetItem(sheet);
            container.appendChild(item);
        });
        
        // 如果沒有變更
        if (allSheets.length === 0) {
            container.innerHTML = '<p style="text-align: center; color: #666;">兩個文件完全相同</p>';
        }
    },
    
    // 創建 sheet item 元素
    createSheetItem(sheet) {
        const item = document.createElement('div');
        item.className = 'sheet-item';
        
        // Sheet 名稱
        const nameDiv = document.createElement('div');
        nameDiv.className = 'sheet-item-name';
        
        // 添加圖標
        const icon = this.getStatusIcon(sheet.status);
        nameDiv.innerHTML = `<span>${icon}</span><strong>${sheet.name}</strong>`;
        
        // 狀態標籤
        const statusDiv = document.createElement('div');
        statusDiv.className = `sheet-status status-${sheet.status}`;
        statusDiv.textContent = this.getStatusText(sheet);
        
        item.appendChild(nameDiv);
        item.appendChild(statusDiv);
        
        // 添加點擊事件
        if (sheet.canView) {
            item.style.cursor = 'pointer';
            item.addEventListener('click', () => {
                this.viewSheetDiff(sheet);
            });
        }
        
        return item;
    },
    
    // 獲取狀態圖標
    getStatusIcon(status) {
        const icons = {
            unchanged: '✅',
            modified: '✏️',
            added: '➕',
            removed: '❌',
            renamed: '🔄'
        };
        return icons[status] || '•';
    },
    
    // 獲取狀態文字
    getStatusText(sheet) {
        switch (sheet.status) {
            case 'unchanged':
                return '無變化';
            case 'modified':
                return `${sheet.changeCount} 個變更`;
            case 'added':
                return '新增';
            case 'removed':
                return '刪除';
            case 'renamed':
                return `重命名 (${sheet.confidence}% 相似)`;
            default:
                return '';
        }
    },
    
    // 查看 sheet 的詳細差異
    viewSheetDiff(sheet) {
        // 準備數據
        let sheetToView = sheet.name;
        
        // 如果是重命名的，使用新名稱
        if (sheet.status === 'renamed') {
            sheetToView = sheet.newName;
        }
        
        // 切換到 diff 視圖
        DiffViewer.show(
            this.currentDiffResult,
            this.parsedFileA,
            this.parsedFileB,
            sheetToView,
            sheet.status,
            sheet.viewSide
        );
    },
    
    // 隱藏 summary 視圖
    hide() {
        document.getElementById('summarySection').style.display = 'none';
    }
};