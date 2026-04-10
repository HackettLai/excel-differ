// excelParser.js
// Parses Excel files (.xlsx, .xls) and CSV files into JavaScript objects using XLSX.js library
// Converts Excel sheets into structured data with column headers (A, B, C, etc.)
// Automatically trims leading/trailing blank rows and completely blank columns

/**
 * ExcelParser Class
 * Handles parsing of Excel and CSV files into structured JavaScript objects
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
   * After trim, the actual data length is counted as rowCount.
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
        blankrows: true, // Include blank rows initially
      });

      // Normalize AND trim data
      const processedData = this.normalizeAndTrimSheetData(jsonData);

      // Use processed data length
      const rowCount = processedData.length;

      result.sheets[sheetName] = {
        data: processedData,
        rowCount: rowCount, 
      };

      console.log(`📊 Sheet "${sheetName}": ${rowCount} rows (after trimming)`);
    });

    return result;
  }

  /**
   * normalizeAndTrimSheetData(jsonData)
   * Normalize values AND trim blank rows/columns
   * 
   * Process:
   * 1. Normalize all cell values (trim whitespace, remove invisible chars)
   * 2. Trim leading blank rows
   * 3. Trim trailing blank rows
   * 4. Trim completely blank columns
   *
   * @param {Array<Object>} jsonData - Raw sheet data
   * @returns {Array<Object>} Normalized and trimmed data
   */
  normalizeAndTrimSheetData(jsonData) {
    if (!jsonData || jsonData.length === 0) return [];

    // Step 1: Normalize all values
    const normalized = jsonData.map((row) => {
      const normalizedRow = {};

      Object.keys(row).forEach((col) => {
        let value = row[col];

        // Normalize value
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

    // Step 2: Trim leading blank rows
    let firstNonBlankIndex = 0;
    for (let i = 0; i < normalized.length; i++) {
      if (!this.isRowBlank(normalized[i])) {
        firstNonBlankIndex = i;
        break;
      }
    }

    const afterLeadingTrim = normalized.slice(firstNonBlankIndex);

    if (firstNonBlankIndex > 0) {
      console.log(`🧹 Trimmed ${firstNonBlankIndex} leading blank rows`);
    }

    // Step 3: Trim trailing blank rows
    let lastNonBlankIndex = afterLeadingTrim.length - 1;
    for (let i = afterLeadingTrim.length - 1; i >= 0; i--) {
      if (!this.isRowBlank(afterLeadingTrim[i])) {
        lastNonBlankIndex = i;
        break;
      }
    }

    const afterTrailingTrim = afterLeadingTrim.slice(0, lastNonBlankIndex + 1);

    const trailingRowsRemoved = afterLeadingTrim.length - afterTrailingTrim.length;
    if (trailingRowsRemoved > 0) {
      console.log(`🧹 Trimmed ${trailingRowsRemoved} trailing blank rows`);
    }

    // Step 4: Trim completely blank columns
    const finalData = this.trimCompletelyBlankColumns(afterTrailingTrim);

    const totalRowsRemoved = normalized.length - finalData.length;
    if (totalRowsRemoved > 0) {
      console.log(`🧹 Total: ${normalized.length} → ${finalData.length} rows`);
    }

    return finalData;
  }

  /**
   * trimCompletelyBlankColumns(data)
   * Remove completely blank columns (all rows are blank columns).
   */
  trimCompletelyBlankColumns(data) {
    if (!data || data.length === 0) return data;

    const allColumns = this.getAllColumns(data);
    const nonBlankColumns = [];

    // Check each column
    allColumns.forEach((col) => {
      const hasData = data.some((row) => {
        const cell = row[col];
        if (cell === null || cell === undefined || cell === '') return false;
        if (typeof cell === 'string' && cell.trim() === '') return false;
        return true;
      });

      if (hasData) {
        nonBlankColumns.push(col);
      }
    });

    const trimmedCount = allColumns.length - nonBlankColumns.length;
    if (trimmedCount > 0) {
      console.log(`🧹 Trimmed ${trimmedCount} completely blank columns`);
    }

    // Rebuild data with only non-blank columns
    return data.map((row) => {
      const trimmedRow = {};
      nonBlankColumns.forEach((col) => {
        trimmedRow[col] = row[col] ?? null;
      });
      return trimmedRow;
    });
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
   * Valid Extensions: .xlsx, .xls, .csv
   * Valid MIME Types:
   * - application/vnd.openxmlformats-officedocument.spreadsheetml.sheet (.xlsx)
   * - application/vnd.ms-excel (.xls)
   * - text/csv (.csv)
   * - application/csv (.csv)
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