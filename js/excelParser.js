// excelParser.js

class ExcelParser {
    constructor() {
        this.workbook = null;
    }

    async parse(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    const parsedData = {
                        fileName: file.name,
                        sheets: {},
                        sheetNames: workbook.SheetNames
                    };

                    // Parse each sheet
                    workbook.SheetNames.forEach(sheetName => {
                        const worksheet = workbook.Sheets[sheetName];
                        const sheetData = this.parseSheet(worksheet);

                        parsedData.sheets[sheetName] = {
                            name: sheetName,
                            data: sheetData,
                            rowCount: sheetData.length,
                            colCount: sheetData.length > 0 ? Object.keys(sheetData[0]).length : 0
                        };
                    });

                    console.log('✅ Parsed workbook:', parsedData);
                    resolve(parsedData);
                } catch (error) {
                    console.error('Error parsing Excel file:', error);
                    reject(error);
                }
            };

            reader.onerror = (error) => {
                console.error('FileReader error:', error);
                reject(error);
            };

            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * ✅ 修正：使用 A, B, C... 作為欄位名稱，保留所有行（包括第一行）
     */
    parseSheet(worksheet) {
        // 使用 header: 1 取得原始陣列（每一列都是陣列）
        const rawData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: null,
            blankrows: true
        });

        if (rawData.length === 0) {
            console.warn('Empty sheet');
            return [];
        }

        // ✅ 找出最大欄位數
        const maxCols = Math.max(...rawData.map(row => row.length));
        
        console.log('📋 Raw data rows:', rawData.length, 'Max columns:', maxCols);
        
        // ✅ 將每一列轉換為物件，使用 A, B, C... 作為 key
        const dataRows = rawData.map((row, rowIndex) => {
            const rowObj = {};
            
            for (let colIndex = 0; colIndex < maxCols; colIndex++) {
                const colName = this.getColumnName(colIndex);  // A, B, C...
                rowObj[colName] = row[colIndex] ?? null;
            }
            
            return rowObj;
        });

        console.log('✅ First data row:', dataRows[0]);
        console.log('✅ Column names:', Object.keys(dataRows[0]));
        
        return dataRows;
    }

    /**
     * ✅ 取得 Excel 欄位名稱 (A, B, C, ..., Z, AA, AB, ...)
     */
    getColumnName(index) {
        let name = '';
        index++;
        while (index > 0) {
            const mod = (index - 1) % 26;
            name = String.fromCharCode(65 + mod) + name;
            index = Math.floor((index - mod) / 26);
        }
        return name;
    }

    /**
     * Get cell reference (e.g., "A1", "B2")
     */
    static getCellReference(row, col) {
        return `${ExcelParser.getColumnName(col)}${row + 1}`;
    }

    /**
     * Get column name from index (A, B, C, ..., Z, AA, AB, ...)
     */
    static getColumnName(index) {
        let name = '';
        index++;
        while (index > 0) {
            const mod = (index - 1) % 26;
            name = String.fromCharCode(65 + mod) + name;
            index = Math.floor((index - mod) / 26);
        }
        return name;
    }

    /**
     * Validate Excel file
     */
    static isValidExcelFile(file) {
        if (!file) return false;

        const validExtensions = ['.xlsx', '.xls'];
        const validMimeTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];

        const fileName = file.name.toLowerCase();
        const hasValidExtension = validExtensions.some(ext => fileName.endsWith(ext));
        const hasValidMimeType = validMimeTypes.includes(file.type);

        return hasValidExtension || hasValidMimeType;
    }
}

export default ExcelParser;