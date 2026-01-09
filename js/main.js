// main.js - 主流程控制

const App = {
    
    // 初始化應用
    init() {
        console.log('📊 Excel Differ 初始化...');
        
        // 檢查 SheetJS 是否加載
        if (typeof XLSX === 'undefined') {
            alert('錯誤：SheetJS 庫未加載！');
            return;
        }
        
        // 初始化文件處理器
        FileHandler.init();
        
        // 綁定事件
        this.bindEvents();
        
        console.log('✅ 初始化完成');
    },
    
    // 綁定事件
    bindEvents() {
        // 比較按鈕
        document.getElementById('compareBtn').addEventListener('click', () => {
            this.startComparison();
        });
        
        // 返回上傳按鈕
        document.getElementById('backBtn').addEventListener('click', () => {
            this.backToUpload();
        });
        
        // 返回摘要按鈕
        document.getElementById('backToSummaryBtn').addEventListener('click', () => {
            this.backToSummary();
        });
    },
    
    // 開始比較流程
    async startComparison() {
        try {
            // 顯示 loading
            this.showLoading(true);
            
            // 讀取文件
            console.log('📖 讀取文件...');
            const fileDataA = await FileHandler.getFileData('A');
            const fileDataB = await FileHandler.getFileData('B');
            
            // 解析 Excel
            console.log('🔍 解析 Excel...');
            const parsedA = ExcelParser.parse(fileDataA.arrayBuffer, fileDataA.file.name);
            const parsedB = ExcelParser.parse(fileDataB.arrayBuffer, fileDataB.file.name);
            
            console.log('File A:', parsedA);
            console.log('File B:', parsedB);
            
            // 執行 Diff
            console.log('⚡ 執行差異比較...');
            const diffResult = DiffEngine.compare(parsedA, parsedB);
            
            console.log('Diff 結果:', diffResult);
            
            // 隱藏 loading
            this.showLoading(false);
            
            // 顯示 Summary
            SummaryView.show(diffResult, parsedA, parsedB);
            
        } catch (error) {
            this.showLoading(false);
            console.error('比較過程中發生錯誤:', error);
            alert(`錯誤：${error.message}`);
        }
    },
    
    // 返回上傳頁面
    backToUpload() {
        // 隱藏 summary
        SummaryView.hide();
        
        // 隱藏 diff
        DiffViewer.hide();
        
        // 顯示上傳區域
        document.getElementById('uploadSection').style.display = 'block';
        
        // 重置文件
        FileHandler.reset();
    },
    
    // 返回摘要頁面
    backToSummary() {
        // 隱藏 diff
        DiffViewer.hide();
        
        // 顯示 summary
        document.getElementById('summarySection').style.display = 'block';
    },
    
    // 顯示/隱藏 loading
    showLoading(show) {
        const overlay = document.getElementById('loadingOverlay');
        overlay.style.display = show ? 'flex' : 'none';
    }
};

// 當 DOM 加載完成後初始化
document.addEventListener('DOMContentLoaded', () => {
    App.init();
});

// 防止意外關閉頁面時丟失數據
window.addEventListener('beforeunload', (e) => {
    if (FileHandler.fileA || FileHandler.fileB) {
        e.preventDefault();
        e.returnValue = '確定要離開？未保存的比較結果將丟失。';
    }
});