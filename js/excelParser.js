// excelParser.js - 使用 SheetJS 解析 Excel 文件

const ExcelParser = {
    
    // 解析 Excel 文件
    parse(arrayBuffer, fileName) {
        try {
            // 使用 SheetJS 讀取工作簿
            const workbook = XLSX.read(arrayBuffer, { 
                type: 'array',
                cellDates: true,
                cellNF: false,
                cellStyles: false
            });
            
            const sheets = [];
            
            // 遍歷所有 sheet
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const sheetData = this.parseSheet(worksheet, sheetName);
                sheets.push(sheetData);
            });
            
            return {
                fileName: fileName,
                sheets: sheets,
                sheetCount: sheets.length
            };
            
        } catch (error) {
            console.error('Excel 解析錯誤:', error);
            throw new Error(`無法解析 Excel 文件: ${error.message}`);
        }
    },
    
    // 解析單個 sheet
    parseSheet(worksheet, sheetName) {
        // 獲取 sheet 的範圍
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        
        const data = [];
        const maxRow = range.e.r;
        const maxCol = range.e.c;
        
        // 逐行讀取數據
        for (let R = range.s.r; R <= maxRow; R++) {
            const row = [];
            
            for (let C = range.s.c; C <= maxCol; C++) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = worksheet[cellAddress];
                
                // 獲取 cell 值
                let cellValue = '';
                if (cell) {
                    if (cell.f) {
                        // 如果有公式，優先使用計算值
                        cellValue = cell.v !== undefined ? cell.v : cell.f;
                    } else {
                        cellValue = cell.v !== undefined ? cell.v : '';
                    }
                    
                    // 處理日期格式
                    if (cell.t === 'd') {
                        cellValue = this.formatDate(cellValue);
                    }
                    
                    // 處理數字格式
                    if (cell.t === 'n' && cell.z) {
                        cellValue = this.formatNumber(cellValue, cell.z);
                    }
                }
                
                row.push(cellValue);
            }
            
            data.push(row);
        }
        
        return {
            name: sheetName,
            data: data,
            rowCount: maxRow + 1,
            colCount: maxCol + 1,
            range: worksheet['!ref'] || 'A1'
        };
    },
    
    // 格式化日期
    formatDate(date) {
        if (date instanceof Date) {
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }
        return date;
    },
    
    // 格式化數字
    formatNumber(num, format) {
        // 簡單的數字格式化
        if (typeof num === 'number') {
            // 檢查是否為百分比格式
            if (format && format.includes('%')) {
                return (num * 100).toFixed(2) + '%';
            }
            
            // 檢查是否需要千分位
            if (format && format.includes(',')) {
                return num.toLocaleString('zh-TW');
            }
            
            // 默認保留兩位小數
            if (num % 1 !== 0) {
                return num.toFixed(2);
            }
        }
        
        return num;
    },
    
    // 獲取 cell 的列名 (A, B, C, ... Z, AA, AB, ...)
    getColumnName(colIndex) {
        let columnName = '';
        let dividend = colIndex + 1;
        
        while (dividend > 0) {
            const modulo = (dividend - 1) % 26;
            columnName = String.fromCharCode(65 + modulo) + columnName;
            dividend = Math.floor((dividend - modulo) / 26);
        }
        
        return columnName;
    },
    
    // 將數據轉換為二維數組（標準化）
    normalizeData(data) {
        if (!data || data.length === 0) {
            return [];
        }
        
        // 找出最大列數
        const maxCols = Math.max(...data.map(row => row.length));
        
        // 補齊每一行的列數
        return data.map(row => {
            const normalizedRow = [...row];
            while (normalizedRow.length < maxCols) {
                normalizedRow.push('');
            }
            return normalizedRow;
        });
    },
    
    // 比較兩個值是否相等
    areValuesEqual(valueA, valueB) {
        // 處理空值
        if (valueA === null || valueA === undefined || valueA === '') {
            return valueB === null || valueB === undefined || valueB === '';
        }
        
        if (valueB === null || valueB === undefined || valueB === '') {
            return false;
        }
        
        // 轉換為字符串比較
        const strA = String(valueA).trim();
        const strB = String(valueB).trim();
        
        return strA === strB;
    }
};