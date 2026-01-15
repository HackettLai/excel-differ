/**
 * diffEngine.js
 * Excel comparison engine - core logic for comparing two Excel workbooks
 * Detects differences in sheets, rows, columns, and cell values
 */

/**
 * DiffEngine Class
 * Handles all comparison logic between two parsed Excel workbooks
 * Generates comprehensive diff results including added/deleted/modified sheets and cells
 */
class DiffEngine {
  /**
   * Constructor
   * Initializes the diff results structure
   */
  constructor() {
    // Results object stores all comparison data
    this.results = {
      summary: {
        totalSheets: 0,      // Total number of unique sheets across both files
        modifiedSheets: 0,   // Number of sheets with changes
        addedSheets: 0,      // Number of sheets only in File B (new)
        deletedSheets: 0     // Number of sheets only in File A (removed)
      },
      sheets: []             // Array of per-sheet comparison results
    };
  }

  /**
   * compare(dataA, dataB)
   * Main comparison entry point - compares two parsed Excel workbooks
   * 
   * @param {Object} dataA - Parsed data from File A (old file)
   * @param {Object} dataB - Parsed data from File B (new file)
   * @returns {Object} Complete diff results with summary and per-sheet details
   * 
   * Process:
   * 1. Validates input data
   * 2. Finds all unique sheet names across both files
   * 3. For each sheet:
   *    - If exists in both: compare content
   *    - If only in B: mark as added
   *    - If only in A: mark as deleted
   * 4. Updates summary statistics
   */
  compare(dataA, dataB) {
    console.log('🔍 Starting comparison...', { dataA, dataB });

    // Validate input data
    if (!dataA || !dataB) {
      console.error('❌ Missing comparison data');
      return this.results;
    }

    // Get all sheet names from both files
    const sheetsA = new Set(dataA.sheetNames || []);
    const sheetsB = new Set(dataB.sheetNames || []);

    // Combine all unique sheet names
    const allSheets = new Set([...sheetsA, ...sheetsB]);

    // Process each sheet
    allSheets.forEach(sheetName => {
      const inA = sheetsA.has(sheetName);  // Sheet exists in File A
      const inB = sheetsB.has(sheetName);  // Sheet exists in File B

      // Initialize sheet result object
      let sheetResult = {
        sheetName: sheetName,
        status: 'unchanged',     // Status: 'unchanged' | 'modified' | 'added' | 'deleted'
        differences: [],         // Cell-level differences
        rowChanges: [],          // Added/deleted rows
        columnChanges: [],       // Added/deleted columns
        oldData: [],             // Original data from File A
        newData: []              // New data from File B
      };

      if (inA && inB) {
        // Sheet exists in both files - perform detailed comparison
        const sheetA = dataA.sheets[sheetName]?.data || [];
        const sheetB = dataB.sheets[sheetName]?.data || [];

        sheetResult = this.compareSheets(sheetA, sheetB);
        sheetResult.sheetName = sheetName;

        // Mark as modified if any changes detected
        if (sheetResult.differences.length > 0 || 
            sheetResult.rowChanges.length > 0 || 
            sheetResult.columnChanges.length > 0) {
          sheetResult.status = 'modified';
          this.results.summary.modifiedSheets++;
        }
      } else if (inB && !inA) {
        // Sheet only exists in File B - it's a new sheet
        sheetResult.status = 'added';
        sheetResult.newData = dataB.sheets[sheetName]?.data || [];
        this.results.summary.addedSheets++;
      } else if (inA && !inB) {
        // Sheet only exists in File A - it was deleted
        sheetResult.status = 'deleted';
        sheetResult.oldData = dataA.sheets[sheetName]?.data || [];
        this.results.summary.deletedSheets++;
      }

      // Add sheet result to results array
      this.results.sheets.push(sheetResult);
    });

    // Update total sheets count
    this.results.summary.totalSheets = allSheets.size;

    console.log('✅ Comparison complete', this.results);
    return this.results;
  }

  /**
   * compareSheets(oldData, newData)
   * Compares two individual sheets in detail
   * 
   * @param {Array<Object>} oldData - Sheet data from File A
   * @param {Array<Object>} newData - Sheet data from File B
   * @returns {Object} Sheet comparison result with differences, row/column changes
   * 
   * Process:
   * 1. Detects column changes (added/deleted columns)
   * 2. Detects row changes (added/deleted rows)
   * 3. Compares cell content for matching rows
   */
  compareSheets(oldData, newData) {
    const result = {
      differences: [],     // Array of cell-level differences
      rowChanges: [],      // Array of added/deleted rows
      columnChanges: [],   // Array of added/deleted columns
      oldData: oldData,    // Original sheet data
      newData: newData     // New sheet data
    };

    // Validate input data
    if (!oldData || !newData || oldData.length === 0 || newData.length === 0) {
      return result;
    }

    // Detect column changes (added/deleted columns)
    result.columnChanges = this.detectColumnChanges(oldData, newData);

    // Detect row changes (added/deleted rows)
    result.rowChanges = this.detectRowChanges(oldData, newData);

    // Compare cell content
    result.differences = this.compareCells(oldData, newData);

    return result;
  }

  /**
   * detectColumnChanges(oldData, newData)
   * Detects added and deleted columns based on header content
   * Ignores column position changes (only detects true additions/deletions)
   * 
   * @param {Array<Object>} oldData - Sheet data from File A
   * @param {Array<Object>} newData - Sheet data from File B
   * @returns {Array<Object>} Array of column change objects
   * 
   * Each change object:
   * {
   *   column: string,    // Column letter (e.g., 'A', 'B', 'AA')
   *   type: string,      // 'added' or 'deleted'
   *   header: string     // Header content or '(Blank Column)'
   * }
   * 
   * Logic:
   * - Uses first row as headers
   * - Compares header CONTENT, not column position
   * - A column is "added" if its header exists in File B but not in File A
   * - A column is "deleted" if its header exists in File A but not in File B
   */
  detectColumnChanges(oldData, newData) {
    if (!oldData || !newData || oldData.length === 0 || newData.length === 0) {
      return [];
    }

    const oldHeaders = oldData[0];  // First row of File A
    const newHeaders = newData[0];  // First row of File B

    // Build sets of header content (ignore null/empty)
    const oldHeaderSet = new Set();
    const newHeaderSet = new Set();

    Object.values(oldHeaders).forEach(val => {
      const content = String(val || '').trim();
      if (content) oldHeaderSet.add(content);
    });

    Object.values(newHeaders).forEach(val => {
      const content = String(val || '').trim();
      if (content) newHeaderSet.add(content);
    });

    const changes = [];

    // Get column letters
    const oldCols = Object.keys(oldHeaders);
    const newCols = Object.keys(newHeaders);

    // Find added columns (exist in File B but not in File A)
    newCols.forEach(col => {
      const newContent = String(newHeaders[col] || '').trim();
      
      if (newContent) {
        // Check if this header content is truly new (doesn't exist in File A)
        if (!oldHeaderSet.has(newContent)) {
          changes.push({
            column: col,
            type: 'added',
            header: newContent
          });
        }
      } else {
        // Check if this blank column is new (column letter doesn't exist in File A)
        if (!oldCols.includes(col)) {
          changes.push({
            column: col,
            type: 'added',
            header: '(Blank Column)'
          });
        }
      }
    });

    // Find deleted columns (exist in File A but not in File B)
    oldCols.forEach(col => {
      const oldContent = String(oldHeaders[col] || '').trim();
      
      if (oldContent) {
        // Check if this header content was deleted (doesn't exist in File B)
        if (!newHeaderSet.has(oldContent)) {
          changes.push({
            column: col,
            type: 'deleted',
            header: oldContent
          });
        }
      } else {
        // Check if this blank column was deleted (column letter doesn't exist in File B)
        if (!newCols.includes(col)) {
          changes.push({
            column: col,
            type: 'deleted',
            header: '(Blank Column)'
          });
        }
      }
    });

    console.log('📊 Column Changes:', changes);
    return changes;
  }

  /**
   * detectRowChanges(oldData, newData)
   * Detects added and deleted rows using column A as the row key
   * 
   * @param {Array<Object>} oldData - Sheet data from File A
   * @param {Array<Object>} newData - Sheet data from File B
   * @returns {Array<Object>} Array of row change objects
   * 
   * Each change object:
   * {
   *   rowKey: string,        // Value from column A (row identifier)
   *   type: string,          // 'added' or 'deleted'
   *   oldRowIndex: number,   // Original row number (for deleted rows)
   *   newRowIndex: number,   // New row number (for added rows)
   *   row: Object            // Full row data
   * }
   * 
   * Logic:
   * - Skips first row (header row)
   * - Uses column A value as unique row identifier
   * - If column A is empty, uses fallback key "old-{index}" or "new-{index}"
   * - A row is "added" if its key exists in File B but not in File A
   * - A row is "deleted" if its key exists in File A but not in File B
   */
  detectRowChanges(oldData, newData) {
    const changes = [];

    if (!oldData || !newData) return changes;

    // Skip first row (header row)
    const oldRows = oldData.slice(1);
    const newRows = newData.slice(1);

    const oldRowMap = new Map();
    const newRowMap = new Map();

    // Build row map for File A (using column A as key)
    oldRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `old-${index}`;
      oldRowMap.set(key, { row, index: index + 2 }); // +2 because: +1 for 1-based, +1 for header
    });

    // Build row map for File B (using column A as key)
    newRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `new-${index}`;
      newRowMap.set(key, { row, index: index + 2 }); // +2 because: +1 for 1-based, +1 for header
    });

    // Find added rows (exist in File B but not in File A)
    newRowMap.forEach((data, key) => {
      if (!oldRowMap.has(key)) {
        changes.push({
          rowKey: key,
          type: 'added',
          newRowIndex: data.index,
          row: data.row
        });
      }
    });

    // Find deleted rows (exist in File A but not in File B)
    oldRowMap.forEach((data, key) => {
      if (!newRowMap.has(key)) {
        changes.push({
          rowKey: key,
          type: 'deleted',
          oldRowIndex: data.index,
          row: data.row
        });
      }
    });

    return changes;
  }

  /**
   * compareCells(oldData, newData)
   * Compares individual cell values for matching rows
   * Uses header content mapping to handle column reordering
   * 
   * @param {Array<Object>} oldData - Sheet data from File A
   * @param {Array<Object>} newData - Sheet data from File B
   * @returns {Array<Object>} Array of cell difference objects
   * 
   * Each difference object:
   * {
   *   row: number,         // Row number
   *   header: string,      // Header content (column identifier)
   *   oldCol: string,      // Old column letter
   *   newCol: string,      // New column letter
   *   oldValue: any,       // Original cell value
   *   newValue: any        // New cell value
   * }
   * 
   * Logic:
   * 1. Builds header content → column letter mapping for both files
   * 2. Builds row key → row data mapping using column A
   * 3. For matching rows (same row key):
   *    - For each header in File A, find corresponding column in File B
   *    - Compare cell values
   *    - Record differences
   * 
   * This approach handles column reordering correctly:
   * - If "Email Address" moves from column G to column H, cells are still matched correctly
   */
  compareCells(oldData, newData) {
    const differences = [];

    if (!oldData || !newData) return differences;

    // Build header mapping (header content → column letters)
    const oldHeaders = oldData[0] || {};
    const newHeaders = newData[0] || {};
    
    const headerToOldCol = new Map();  // "Email Address" → "G"
    const headerToNewCol = new Map();  // "Email Address" → "H"

    // Map header content to column letters in File A
    Object.keys(oldHeaders).forEach(col => {
      const content = String(oldHeaders[col] || '').trim();
      if (content) {
        headerToOldCol.set(content, col);
      }
    });

    // Map header content to column letters in File B
    Object.keys(newHeaders).forEach(col => {
      const content = String(newHeaders[col] || '').trim();
      if (content) {
        headerToNewCol.set(content, col);
      }
    });

    // Build row mapping (skip header row)
    const oldRows = oldData.slice(1);
    const newRows = newData.slice(1);

    const oldRowMap = new Map();
    const newRowMap = new Map();

    // Map row keys to row data for File A
    oldRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `old-${index}`;
      oldRowMap.set(key, { row, index: index + 2 });
    });

    // Map row keys to row data for File B
    newRows.forEach((row, index) => {
      const key = String(row.A || '').trim() || `new-${index}`;
      newRowMap.set(key, { row, index: index + 2 });
    });

    // Compare cells for matching rows (using header content mapping)
    oldRowMap.forEach((oldRowData, key) => {
      if (newRowMap.has(key)) {
        const newRowData = newRowMap.get(key);
        const oldRow = oldRowData.row;
        const newRow = newRowData.row;

        // For each header in File A, compare with corresponding column in File B
        headerToOldCol.forEach((oldCol, headerContent) => {
          const newCol = headerToNewCol.get(headerContent);
          
          if (newCol) {
            // Same header exists in both files - compare values
            const oldVal = oldRow[oldCol];
            const newVal = newRow[newCol];

            // Record difference if values don't match
            if (oldVal !== newVal) {
              differences.push({
                row: oldRowData.index,
                header: headerContent,    // Use header content as identifier
                oldCol: oldCol,
                newCol: newCol,
                oldValue: oldVal,
                newValue: newVal
              });
            }
          }
        });
      }
    });

    return differences;
  }
}

// Export for use in other modules
export default DiffEngine;