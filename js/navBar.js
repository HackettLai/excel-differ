// navBar.js - Sheet 導航列

const NavBar = {
    currentSheet: null,
    diffResult: null,
    parsedFileA: null,
    parsedFileB: null,
    onSheetChange: null,
    
    // 初始化導航列
    init(diffResult, fileA, fileB, currentSheet, onSheetChangeCallback) {
        this.diffResult = diffResult;
        this.parsedFileA = fileA;
        this.parsedFileB = fileB;
        this.currentSheet = currentSheet;
        this.onSheetChange = onSheetChangeCallback;
        
        this.render();
    },
    
    // 渲染導航列
    render() {
        const navBar = document.getElementById('sheetNavBar');
        navBar.innerHTML = '';
        
        const { sheetChanges, cellDiffs } = this.diffResult;
        
        // 收集所有可見的 sheet tabs
        const tabs = [];
        
        // 1. Common sheets (包括 modified 和 unchanged)
        sheetChanges.common.forEach(sheetName => {
            const diff = cellDiffs[sheetName];
            tabs.push({
                name: sheetName,
                displayName: sheetName,
                status: diff.changes.length > 0 ? 'modified' : 'unchanged',
                available: true
            });
        });
        
        // 2. Renamed sheets
        sheetChanges.renamed.forEach(rename => {
            tabs.push({
                name: rename.to,
                displayName: `${rename.from} → ${rename.to}`,
                status: 'renamed',
                available: true
            });
        });
        
        // 3. Added sheets
        sheetChanges.added.forEach(sheetName => {
            tabs.push({
                name: sheetName,
                displayName: sheetName,
                status: 'added',
                available: true,
                side: 'B'
            });
        });
        
        // 4. Removed sheets
        sheetChanges.removed.forEach(sheetName => {
            tabs.push({
                name: sheetName,
                displayName: sheetName,
                status: 'removed',
                available: true,
                side: 'A'
            });
        });
        
        // 渲染每個 tab
        tabs.forEach(tab => {
            const tabElement = this.createTab(tab);
            navBar.appendChild(tabElement);
        });
    },
    
    // 創建單個 tab
    createTab(tab) {
        const tabDiv = document.createElement('div');
        tabDiv.className = 'sheet-tab';
        
        // 添加狀態 class
        if (tab.status === 'modified') {
            tabDiv.classList.add('has-changes');
        } else if (tab.status === 'added') {
            tabDiv.classList.add('added');
        } else if (tab.status === 'removed') {
            tabDiv.classList.add('removed');
        }
        
        // 如果是當前 sheet，添加 active class
        if (tab.name === this.currentSheet) {
            tabDiv.classList.add('active');
        }
        
        // 設置文字
        tabDiv.textContent = tab.displayName;
        
        // 添加點擊事件
        tabDiv.addEventListener('click', () => {
            if (this.onSheetChange) {
                this.onSheetChange(tab.name, tab.status, tab.side);
            }
        });
        
        return tabDiv;
    },
    
    // 更新當前選中的 sheet
    setActiveSheet(sheetName) {
        this.currentSheet = sheetName;
        
        // 更新所有 tab 的 active 狀態
        const tabs = document.querySelectorAll('.sheet-tab');
        tabs.forEach(tab => {
            tab.classList.remove('active');
            if (tab.textContent.includes(sheetName) || 
                tab.textContent === sheetName) {
                tab.classList.add('active');
            }
        });
    }
};