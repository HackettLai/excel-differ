// excelParser.js
// Parses Excel files (.xlsx, .xls) into JavaScript objects using XLSX.js library
// Converts Excel sheets into structured data with column headers (A, B, C, etc.)

/**
 * ExcelParser Class
 * Handles parsing of Excel files into structured JavaScript objects
 * Uses XLSX.js library for reading Excel file formats
 */
class ExcelParser {
    /**
     * Constructor
     * Initializes the parser
     */
    constructor() {
        this.workbook = null; // Stores the current workbook being processed
    }

    /**
     * parse(file)
     * Main parsing method - converts Excel file to structured data object
     * 
     * @param {File} file - The Excel file object to parse
     * @returns {Promise<Object>} Resolves with parsed workbook data
     * 
     * Returned Object Structure:
     * {
     *   fileName: string,
     *   sheets: {
     *     [sheetName]: {
     *       name: string,
     *       data: Array<Object>,
     *       rowCount: number,
     *       colCount: number
     *     }
     *   },
     *   sheetNames: Array<string>
     * }
     */
    async parse(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            /**
             * onload - Triggered when file is successfully read
             */
            reader.onload = (e) => {
                try {
                    // Convert file data to Uint8Array for XLSX.js
                    const data = new Uint8Array(e.target.result);
                    
                    // Parse Excel file using XLSX.js
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Initialize parsed data structure
                    const parsedData = {
                        fileName: file.name,           // Original file name
                        sheets: {},                    // Object to store all sheets
                        sheetNames: workbook.SheetNames // Array of sheet names
                    };

                    // Parse each sheet in the workbook
                    workbook.SheetNames.forEach(sheetName => {
                        const worksheet = workbook.Sheets[sheetName];
                        const sheetData = this.parseSheet(worksheet);

                        // Store sheet data with metadata
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

            /**
             * onerror - Triggered when file reading fails
             */
            reader.onerror = (error) => {
                console.error('FileReader error:', error);
                reject(error);
            };

            // Start reading file as ArrayBuffer
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * parseSheet(worksheet)
     * Converts an Excel worksheet into an array of row objects
     * Uses column names A, B, C, etc. as object keys
     * Preserves all rows including the first row (no header row assumption)
     * 
     * @param {Object} worksheet - XLSX worksheet object
     * @returns {Array<Object>} Array of row objects with column keys A, B, C, etc.
     * 
     * Example output:
     * [
     *   { A: 'Name', B: 'Age', C: 'City' },
     *   { A: 'John', B: 25, C: 'NYC' },
     *   { A: 'Jane', B: 30, C: 'LA' }
     * ]
     */
    parseSheet(worksheet) {
        // Convert sheet to raw array format (each row is an array)
        // header: 1 means don't use first row as headers
        // defval: null sets default value for empty cells
        // blankrows: true includes blank rows
        const rawData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: null,
            blankrows: true
        });

        // Handle empty sheets
        if (rawData.length === 0) {
            console.warn('Empty sheet');
            return [];
        }

        // Find maximum number of columns across all rows
        const maxCols = Math.max(...rawData.map(row => row.length));
        
        // console.log('📋 Raw data rows:', rawData.length, 'Max columns:', maxCols);
        
        // Convert each row array to an object with column keys A, B, C, etc.
        const dataRows = rawData.map((row, rowIndex) => {
            const rowObj = {};
            
            // Create key-value pairs for each column
            for (let colIndex = 0; colIndex < maxCols; colIndex++) {
                const colName = this.getColumnName(colIndex);  // Get column name: A, B, C...
                rowObj[colName] = row[colIndex] ?? null;        // Use nullish coalescing for empty cells
            }
            
            return rowObj;
        });

        // Debug logging
        // console.log('✅ First data row:', dataRows[0]);
        // console.log('✅ Column names:', Object.keys(dataRows[0]));
        
        return dataRows;
    }

    /**
     * getColumnName(index)
     * Converts column index to Excel column name (A, B, C, ..., Z, AA, AB, ...)
     * 
     * @param {number} index - Zero-based column index
     * @returns {string} Excel column name
     * 
     * Examples:
     * 0 -> 'A'
     * 1 -> 'B'
     * 25 -> 'Z'
     * 26 -> 'AA'
     * 27 -> 'AB'
     */
    getColumnName(index) {
        let name = '';
        index++; // Convert to 1-based index
        
        while (index > 0) {
            const mod = (index - 1) % 26;
            name = String.fromCharCode(65 + mod) + name; // 65 is ASCII code for 'A'
            index = Math.floor((index - mod) / 26);
        }
        
        return name;
    }

    /**
     * getCellReference(row, col)
     * Static method - generates Excel cell reference (e.g., "A1", "B2", "AA10")
     * 
     * @param {number} row - Zero-based row index
     * @param {number} col - Zero-based column index
     * @returns {string} Excel cell reference
     * 
     * Examples:
     * getCellReference(0, 0) -> 'A1'
     * getCellReference(1, 2) -> 'C2'
     * getCellReference(9, 26) -> 'AA10'
     */
    static getCellReference(row, col) {
        return `${ExcelParser.getColumnName(col)}${row + 1}`;
    }

    /**
     * getColumnName(index)
     * Static version of getColumnName instance method
     * Converts column index to Excel column name (A, B, C, ..., Z, AA, AB, ...)
     * 
     * @param {number} index - Zero-based column index
     * @returns {string} Excel column name
     */
    static getColumnName(index) {
        let name = '';
        index++; // Convert to 1-based index
        
        while (index > 0) {
            const mod = (index - 1) % 26;
            name = String.fromCharCode(65 + mod) + name; // 65 is ASCII code for 'A'
            index = Math.floor((index - mod) / 26);
        }
        
        return name;
    }

    /**
     * isValidExcelFile(file)
     * Static method - validates if a file is a valid Excel file
     * Checks both file extension and MIME type
     * 
     * @param {File} file - File object to validate
     * @returns {boolean} True if valid Excel file, false otherwise
     * 
     * Valid Extensions: .xlsx, .xls
     * Valid MIME Types:
     * - application/vnd.openxmlformats-officedocument.spreadsheetml.sheet (.xlsx)
     * - application/vnd.ms-excel (.xls)
     */
    static isValidExcelFile(file) {
        if (!file) return false;

        // Supported Excel file extensions
        const validExtensions = ['.xlsx', '.xls'];
        
        // Supported MIME types
        const validMimeTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel' // .xls
        ];

        const fileName = file.name.toLowerCase();
        
        // Check if file has valid extension
        const hasValidExtension = validExtensions.some(ext => fileName.endsWith(ext));
        
        // Check if file has valid MIME type
        const hasValidMimeType = validMimeTypes.includes(file.type);

        // File is valid if it has either valid extension OR valid MIME type
        return hasValidExtension || hasValidMimeType;
    }
}

// Export for use in other modules
export default ExcelParser;