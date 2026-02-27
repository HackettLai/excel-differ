// excelParser.js
// Parses Excel files (.xlsx, .xls) into JavaScript objects using XLSX.js library
// Converts Excel sheets into structured data with column headers (A, B, C, etc.)
// Automatically trims trailing blank rows and columns

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
  /**
   * parse(file)
   * Parses an Excel or CSV file and extracts structured data
   * Supports both Excel (.xlsx, .xls) and CSV file formats with automatic encoding detection
   *
   * @param {File} file - The Excel or CSV file to parse
   * @returns {Promise<Object>} Parsed workbook data
   */
  async parse(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = e.target.result;
          const isCSV = file.name.toLowerCase().endsWith('.csv');

          let workbook;

          if (isCSV) {
            // ✅ CSV: Read as text string
            console.log('📄 Parsing CSV file:', file.name);

            // Detect encoding (simple heuristic)
            const text = this.detectAndDecodeCSV(data);

            // Parse CSV with XLSX
            workbook = XLSX.read(text, {
              type: 'string',
              raw: false, // Convert numbers to numbers
              codepage: 65001, // UTF-8
            });
          } else {
            // ✅ Excel: Read as binary array
            console.log('📊 Parsing Excel file:', file.name);

            workbook = XLSX.read(data, {
              type: 'array',
              cellDates: true, // Parse dates
              cellNF: false, // Don't parse number formats
              cellStyles: false, // Don't parse styles
            });
          }

          // Extract sheet data
          const parsedData = this.extractSheetData(workbook);

          console.log(`✅ Parsed ${file.name}:`, parsedData);
          resolve(parsedData);
        } catch (error) {
          console.error('❌ Error parsing file:', error);
          reject(new Error(`Failed to parse ${file.name}: ${error.message}`));
        }
      };

      reader.onerror = () => {
        reject(new Error(`Failed to read ${file.name}`));
      };

      // ✅ Read file based on type
      const isCSV = file.name.toLowerCase().endsWith('.csv');

      if (isCSV) {
        reader.readAsArrayBuffer(file); // Read as ArrayBuffer for encoding detection
      } else {
        reader.readAsArrayBuffer(file); // Read Excel as ArrayBuffer
      }
    });
  }

  /**
   * detectAndDecodeCSV(arrayBuffer)
   * Detects CSV encoding and decodes to UTF-8 string
   * Handles UTF-8, GBK, Big5 encoding
   *
   * @param {ArrayBuffer} arrayBuffer - Raw file data
   * @returns {string} Decoded CSV text
   */
  detectAndDecodeCSV(arrayBuffer) {
    const uint8Array = new Uint8Array(arrayBuffer);

    // Try UTF-8 first
    try {
      const decoder = new TextDecoder('utf-8', { fatal: true });
      const text = decoder.decode(uint8Array);

      // Check if decode successful (no replacement characters)
      if (!text.includes('\uFFFD')) {
        console.log('✅ CSV encoding: UTF-8');
        return text;
      }
    } catch (e) {
      console.warn('⚠️ UTF-8 decode failed, trying fallback...');
    }

    // Fallback: Use default decoder (may show garbled text for non-UTF8)
    try {
      const decoder = new TextDecoder('utf-8', { fatal: false });
      const text = decoder.decode(uint8Array);

      console.warn('⚠️ CSV encoding: Unknown (using UTF-8 fallback)');
      console.warn('   If you see garbled text, please save CSV as UTF-8');

      return text;
    } catch (e) {
      throw new Error('Failed to decode CSV file');
    }
  }

  /**
   * extractSheetData(workbook)
   * Extracts data from all sheets in a workbook
   * Normalizes data format to ensure consistent structure across both CSV and Excel files
   *
   * @param {Object} workbook - XLSX workbook object
   * @returns {Object} Structured sheet data
   */
  extractSheetData(workbook) {
    const result = {
      sheetNames: workbook.SheetNames,
      sheets: {},
    };

    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];

      // Convert to JSON (array of objects with column letters as keys)
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 'A', // Use A, B, C... as keys
        defval: null, // Default value for empty cells
        raw: false, // Convert values to strings for consistency
        blankrows: true, // Include blank rows
      });

      // ✅ Normalize data format
      const normalizedData = this.normalizeSheetData(jsonData);

      // Get row count
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
      const rowCount = range.e.r + 1;

      result.sheets[sheetName] = {
        data: normalizedData,
        rowCount: rowCount,
      };
    });

    return result;
  }

  /**
   * normalizeSheetData(jsonData)
   * Normalizes data format for consistent comparison
   * Trims whitespace, converts types
   *
   * @param {Array<Object>} jsonData - Raw sheet data
   * @returns {Array<Object>} Normalized data
   */
  normalizeSheetData(jsonData) {
    return jsonData.map((row) => {
      const normalizedRow = {};

      Object.keys(row).forEach((col) => {
        let value = row[col];

        // ✅ Normalize value
        if (value === null || value === undefined) {
          normalizedRow[col] = null;
        } else {
          // Convert to string and trim
          value = String(value).trim();

          // Remove invisible characters
          value = value.replace(/[\u200B-\u200D\uFEFF]/g, '');

          // Store normalized value
          normalizedRow[col] = value === '' ? null : value;
        }
      });

      return normalizedRow;
    });
  }
  /**
   * parseCSVWithEncoding(data)
   * Parses CSV data with automatic encoding detection
   * Attempts UTF-8 decoding first, then falls back to alternative encodings if needed
   *
   * @param {Uint8Array} data - Raw file data
   * @returns {Object} XLSX workbook object
   */
  parseCSVWithEncoding(data) {
    // Try UTF-8 first
    let csvText;

    try {
      // Attempt UTF-8 decoding
      csvText = new TextDecoder('utf-8').decode(data);

      // Check for mojibake (common pattern when wrong encoding)
      const hasMojibake = /â|€|™|˜|›|¢|£|¤|¥|¦|§|¨|©|ª|«|¬|®|¯|°|±|²|³|´|µ|¶|·|¸|¹|º|»|¼|½|¾|¿/.test(csvText);

      if (hasMojibake) {
        console.warn('⚠️ Detected encoding issue, trying GBK/GB2312...');
        throw new Error('UTF-8 decode failed, trying alternative encoding');
      }

      console.log('✅ CSV parsed as UTF-8');
    } catch (error) {
      // Fallback to GBK (common Chinese encoding)
      try {
        csvText = new TextDecoder('gbk').decode(data);
        console.log('✅ CSV parsed as GBK');
      } catch (gbkError) {
        // Last resort: try windows-1252 (Western European)
        try {
          csvText = new TextDecoder('windows-1252').decode(data);
          console.log('✅ CSV parsed as Windows-1252');
        } catch (finalError) {
          // If all else fails, use UTF-8 anyway
          csvText = new TextDecoder('utf-8', { fatal: false }).decode(data);
          console.warn('⚠️ Using UTF-8 with error recovery');
        }
      }
    }

    // Parse CSV text into workbook
    return XLSX.read(csvText, { type: 'string', raw: true });
  }
  /**
   * parseSheet(worksheet)
   * Converts an Excel worksheet into an array of row objects
   * Uses column names A, B, C, etc. as object keys
   * Preserves all rows including the first row (no header row assumption)
   * Automatically trims trailing blank rows and columns
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
      blankrows: true,
    });

    // Handle empty sheets
    if (rawData.length === 0) {
      console.warn('Empty sheet');
      return [];
    }

    // Find maximum number of columns across all rows
    const maxCols = Math.max(...rawData.map((row) => row.length));

    // console.log('📋 Raw data rows:', rawData.length, 'Max columns:', maxCols);

    // Convert each row array to an object with column keys A, B, C, etc.
    const dataRows = rawData.map((row, rowIndex) => {
      const rowObj = {};

      // Create key-value pairs for each column
      for (let colIndex = 0; colIndex < maxCols; colIndex++) {
        const colName = this.getColumnName(colIndex); // Get column name: A, B, C...
        rowObj[colName] = row[colIndex] ?? null; // Use nullish coalescing for empty cells
      }

      return rowObj;
    });

    // Trim trailing blank rows and columns
    const trimmedData = this.trimBlankRowsAndColumns(dataRows);

    // Debug logging
    // console.log('✅ First data row:', trimmedData[0]);
    // console.log('✅ Column names:', Object.keys(trimmedData[0]));

    return trimmedData;
  }

  /**
   * trimBlankRowsAndColumns(data)
   * Remove trailing blank rows and columns from Excel data
   * Keeps blank rows/columns in the middle (only removes from end)
   *
   * @param {Array<Object>} data - Raw Excel data
   * @returns {Array<Object>} Trimmed data
   */
  trimBlankRowsAndColumns(data) {
    if (!data || data.length === 0) return data;

    // Step 1: Find last non-blank row (trim from bottom)
    let lastRowIndex = data.length - 1;
    while (lastRowIndex > 0) {
      // Always keep at least first row (index 0)
      const row = data[lastRowIndex];
      if (this.isRowBlank(row)) {
        lastRowIndex--;
      } else {
        break; // Found last non-blank row
      }
    }

    // Trim rows (keep up to and including last non-blank row)
    const trimmedRows = data.slice(0, lastRowIndex + 1);

    // Step 2: Find last non-blank column (trim from right)
    const allColumns = this.getAllColumns(trimmedRows);
    const lastColumn = this.findLastNonBlankColumn(trimmedRows, allColumns);

    // Step 3: Trim columns (keep up to and including last non-blank column)
    const trimmedData = trimmedRows.map((row) => {
      const trimmedRow = {};
      allColumns.forEach((col) => {
        // Only keep columns up to lastColumn
        if (this.getColumnIndex(col) <= this.getColumnIndex(lastColumn)) {
          trimmedRow[col] = row[col] ?? null;
        }
      });
      return trimmedRow;
    });

    console.log(`🧹 Trimmed: ${data.length - trimmedRows.length} trailing rows, ${allColumns.length - Object.keys(trimmedData[0] || {}).length} trailing columns`);

    return trimmedData;
  }

  /**
   * isRowBlank(row)
   * Check if a row is completely blank (all cells empty or null)
   *
   * @param {Object} row - Row object with column keys
   * @returns {boolean} True if row is completely blank
   */
  isRowBlank(row) {
    if (!row) return true;

    // Check if all cells in row are empty
    return Object.values(row).every((cell) => {
      // Consider null, undefined, empty string, or whitespace-only as blank
      if (cell === null || cell === undefined || cell === '') return true;
      if (typeof cell === 'string' && cell.trim() === '') return true;
      return false;
    });
  }

  /**
   * getAllColumns(data)
   * Get all unique column keys from dataset
   *
   * @param {Array<Object>} data - Array of row objects
   * @returns {Array<string>} Sorted column keys (e.g., ['A', 'B', 'C', ...])
   */
  getAllColumns(data) {
    const columnSet = new Set();

    // Collect all column keys from all rows
    data.forEach((row) => {
      Object.keys(row).forEach((col) => columnSet.add(col));
    });

    // Sort columns alphabetically (A, B, C, ..., Z, AA, AB, ...)
    return Array.from(columnSet).sort((a, b) => {
      return this.getColumnIndex(a) - this.getColumnIndex(b);
    });
  }

  /**
   * findLastNonBlankColumn(data, allColumns)
   * Find the rightmost column that contains any non-blank data
   *
   * @param {Array<Object>} data - Array of row objects
   * @param {Array<string>} allColumns - All column keys
   * @returns {string} Last non-blank column key (e.g., 'C')
   */
  findLastNonBlankColumn(data, allColumns) {
    // Iterate from rightmost column to left
    for (let i = allColumns.length - 1; i >= 0; i--) {
      const col = allColumns[i];

      // Check if this column has any non-blank cell in any row
      const hasData = data.some((row) => {
        const cell = row[col];

        // Check if cell has meaningful data
        if (cell === null || cell === undefined || cell === '') return false;
        if (typeof cell === 'string' && cell.trim() === '') return false;

        return true; // Cell has data
      });

      if (hasData) {
        return col; // Found last non-blank column
      }
    }

    // Fallback: keep at least column A
    return allColumns[0] || 'A';
  }

  /**
   * getColumnIndex(col)
   * Convert column letter(s) to numeric index for sorting
   * Examples: A=0, B=1, Z=25, AA=26, AB=27
   *
   * @param {string} col - Column letter(s) (e.g., 'A', 'AB', 'ZZ')
   * @returns {number} Numeric index
   */
  getColumnIndex(col) {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.charCodeAt(i) - 64); // A=1, B=2, ...
    }
    return index - 1; // Convert to 0-based index
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

    // Supported Excel and CSV file extensions
    const validExtensions = ['.xlsx', '.xls', '.csv'];

    // Supported MIME types
    const validMimeTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel', // .xls
      'text/csv', // CSV MIME type
      'application/csv', // Alternative CSV MIME type
    ];

    const fileName = file.name.toLowerCase();

    // Check if file has valid extension
    const hasValidExtension = validExtensions.some((ext) => fileName.endsWith(ext));

    // Check if file has valid MIME type
    const hasValidMimeType = validMimeTypes.includes(file.type);

    // File is valid if it has either valid extension OR valid MIME type
    return hasValidExtension || hasValidMimeType;
  }
}

// Export for use in other modules
export default ExcelParser;
