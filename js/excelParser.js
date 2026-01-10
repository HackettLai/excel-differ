// excelParser.js - Excel file parser using SheetJS

const ExcelParser = {
  /**
   * Parse an Excel file from ArrayBuffer
   * @param {ArrayBuffer} arrayBuffer - The Excel file data
   * @param {string} fileName - Name of the file
   * @returns {Object} Parsed Excel data with all sheets
   */
  parse(arrayBuffer, fileName) {
    try {
      // Read workbook using SheetJS
      const workbook = XLSX.read(arrayBuffer, {
        type: 'array',
        cellDates: true,
        cellNF: false,
        cellStyles: false,
      });

      const sheets = [];

      // Iterate through all sheets
      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = this.parseSheet(worksheet, sheetName);
        sheets.push(sheetData);
      });

      return {
        fileName: fileName,
        sheets: sheets,
        sheetCount: sheets.length,
      };
    } catch (error) {
      console.error('Excel parsing error:', error);
      throw new Error(`Unable to parse Excel file: ${error.message}`);
    }
  },

  /**
   * Parse a single worksheet
   * @param {Object} worksheet - SheetJS worksheet object
   * @param {string} sheetName - Name of the sheet
   * @returns {Object} Parsed sheet data with metadata
   */
  parseSheet(worksheet, sheetName) {
    // Get the range of the sheet
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');

    const data = [];
    const maxRow = range.e.r;
    const maxCol = range.e.c;

    // Read data row by row
    for (let R = range.s.r; R <= maxRow; R++) {
      const row = [];

      for (let C = range.s.c; C <= maxCol; C++) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = worksheet[cellAddress];

        // Get cell value
        let cellValue = '';
        if (cell) {
          if (cell.f) {
            // If cell has a formula, use calculated value or formula
            cellValue = cell.v !== undefined ? cell.v : cell.f;
          } else {
            cellValue = cell.v !== undefined ? cell.v : '';
          }

          // Handle date format
          if (cell.t === 'd') {
            cellValue = this.formatDate(cellValue);
          }

          // Handle number format
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
      range: worksheet['!ref'] || 'A1',
    };
  },

  /**
   * Format a date value to YYYY-MM-DD string
   * @param {Date|*} date - Date object or value
   * @returns {string} Formatted date string
   */
  formatDate(date) {
    if (date instanceof Date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    return date;
  },

  /**
   * Format a number according to its format string
   * @param {number} num - The number to format
   * @param {string} format - Excel format string
   * @returns {string|number} Formatted number
   */
  formatNumber(num, format) {
    // Simple number formatting
    if (typeof num === 'number') {
      // Check if it's a percentage format
      if (format && format.includes('%')) {
        return (num * 100).toFixed(2) + '%';
      }

      // Check if thousands separator is needed
      if (format && format.includes(',')) {
        return num.toLocaleString('zh-TW');
      }

      // Default to 2 decimal places for non-integers
      if (num % 1 !== 0) {
        return num.toFixed(2);
      }
    }

    return num;
  },

  /**
   * Get column name from column index (A, B, C, ... Z, AA, AB, ...)
   * @param {number} colIndex - Zero-based column index
   * @returns {string} Column name (e.g., 'A', 'B', 'AA')
   */
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

  /**
   * Normalize data to a 2D array with consistent column count
   * Pads shorter rows with empty strings
   * @param {Array<Array>} data - Raw 2D array data
   * @returns {Array<Array>} Normalized 2D array
   */
  normalizeData(data) {
    if (!data || data.length === 0) {
      return [];
    }

    // Find maximum column count
    const maxCols = Math.max(...data.map((row) => row.length));

    // Pad each row to match the maximum column count
    return data.map((row) => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxCols) {
        normalizedRow.push('');
      }
      return normalizedRow;
    });
  },

  /**
   * Compare two cell values for equality
   * Handles null, undefined, and empty string as equivalent
   * @param {*} valueA - First value to compare
   * @param {*} valueB - Second value to compare
   * @returns {boolean} True if values are equal
   */
  areValuesEqual(valueA, valueB) {
    // Handle empty/null values
    if (valueA === null || valueA === undefined || valueA === '') {
      return valueB === null || valueB === undefined || valueB === '';
    }

    if (valueB === null || valueB === undefined || valueB === '') {
      return false;
    }

    // Convert to strings and compare
    const strA = String(valueA).trim();
    const strB = String(valueB).trim();

    return strA === strB;
  },
};
