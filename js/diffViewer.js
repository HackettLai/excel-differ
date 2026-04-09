/**
 * diffViewer.js
 * ✅ FINAL CORRECTED VERSION
 */

import DiffEngine from './diffEngine.js';

class DiffViewer {
  constructor() {
    this.dataA = null;
    this.dataB = null;
    this.diffResults = null;
    this.changedCells = [];
    this.currentChangeIndex = -1;
  }

  init(dataA, dataB, diffResults) {
    this.dataA = dataA;
    this.dataB = dataB;
    this.diffResults = diffResults;

    console.log('🔧 Initializing DiffViewer...');

    this.populateSheetDropdowns();

    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');

    if (sheetSelectA && this.dataA.sheetNames?.[0]) {
      sheetSelectA.value = this.dataA.sheetNames[0];
      console.log(`✅ Auto-selected Sheet A: ${this.dataA.sheetNames[0]}`);
    }

    if (sheetSelectB && this.dataB.sheetNames?.[0]) {
      sheetSelectB.value = this.dataB.sheetNames[0];
      console.log(`✅ Auto-selected Sheet B: ${this.dataB.sheetNames[0]}`);
    }

    this.populateHeaderRowDropdowns();
    this.populateKeyColumnDropdown();
    this.setupSheetChangeListeners();
    this.setupChangeNavigation();

    console.log('✅ DiffViewer initialized');
  }

  populateHeaderRowDropdowns() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');

    if (!headerRowA || !headerRowB) {
      console.error('❌ Header row dropdown elements not found');
      return;
    }

    const selectedSheetA = sheetSelectA?.value || this.dataA?.sheetNames?.[0];
    const selectedSheetB = sheetSelectB?.value || this.dataB?.sheetNames?.[0];

    const rowCountA = this.dataA?.sheets[selectedSheetA]?.rowCount || 50;
    const rowCountB = this.dataB?.sheets[selectedSheetB]?.rowCount || 50;

    headerRowA.innerHTML = '';
    const maxRowsA = Math.min(rowCountA, 50);
    for (let i = 1; i <= maxRowsA; i++) {
      const option = document.createElement('option');
      option.value = i;
      option.textContent = `Row ${i}`;
      headerRowA.appendChild(option);
    }

    headerRowB.innerHTML = '';
    const maxRowsB = Math.min(rowCountB, 50);
    for (let i = 1; i <= maxRowsB; i++) {
      const option = document.createElement('option');
      option.value = i;
      option.textContent = `Row ${i}`;
      headerRowB.appendChild(option);
    }

    headerRowA.value = '1';
    headerRowB.value = '1';
  }

  populateKeyColumnDropdown() {
    const keyColumnSelect = document.getElementById('keyColumnSelect');
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');

    if (!keyColumnSelect) {
      console.error('❌ Key column dropdown not found');
      return;
    }

    const selectedSheetA = sheetSelectA?.value || this.dataA?.sheetNames?.[0];
    const selectedSheetB = sheetSelectB?.value || this.dataB?.sheetNames?.[0];
    const selectedHeaderA = parseInt(headerRowA?.value) || 1;
    const selectedHeaderB = parseInt(headerRowB?.value) || 1;

    const sheetA = this.dataA?.sheets[selectedSheetA]?.data || [];
    const sheetB = this.dataB?.sheets[selectedSheetB]?.data || [];

    if (sheetA.length === 0 || sheetB.length === 0) {
      console.warn('⚠️ One or both sheets are empty');
      keyColumnSelect.innerHTML = '<option value="">-- No data available --</option>';
      return;
    }

    const headerIndexA = selectedHeaderA - 1;
    const headerIndexB = selectedHeaderB - 1;

    if (headerIndexA >= sheetA.length || headerIndexB >= sheetB.length) {
      console.warn('⚠️ Invalid header row index');
      keyColumnSelect.innerHTML = '<option value="">-- Invalid header row --</option>';
      return;
    }

    const headersA = sheetA[headerIndexA] ? Object.values(sheetA[headerIndexA]).filter((h) => h && String(h).trim()) : [];
    const headersB = sheetB[headerIndexB] ? Object.values(sheetB[headerIndexB]).filter((h) => h && String(h).trim()) : [];

    const commonColumns = this.findCommonColumns(headersA, headersB);

    keyColumnSelect.innerHTML = '';

    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = '── Select Key Column ──';
    keyColumnSelect.appendChild(defaultOption);

    const rowPositionOption = document.createElement('option');
    rowPositionOption.value = '__ROW_POSITION__';
    rowPositionOption.textContent = '(Use Row Position)';
    keyColumnSelect.appendChild(rowPositionOption);

    const separator = document.createElement('option');
    separator.disabled = true;
    separator.textContent = '────────────────────────';
    keyColumnSelect.appendChild(separator);

    if (commonColumns.length === 0) {
      keyColumnSelect.value = '__ROW_POSITION__';
    } else {
      let autoSelected = false;

      commonColumns.forEach((col) => {
        if (!col.name) return;

        const option = document.createElement('option');
        option.value = col.colIndex;
        option.textContent = col.name;

        const lowerName = col.name.toLowerCase();
        if (!autoSelected && (lowerName.includes('email') || lowerName.includes('code') || lowerName.includes('id'))) {
          option.selected = true;
          autoSelected = true;
        }

        keyColumnSelect.appendChild(option);
      });

      if (!autoSelected && commonColumns.length > 0) {
        keyColumnSelect.value = commonColumns[0].colIndex;
      }
    }
  }

  findCommonColumns(headersA, headersB) {
    const commonColumns = [];

    headersA.forEach((headerA, indexA) => {
      if (!headerA) return;
      const trimmedA = String(headerA).trim();
      if (!trimmedA) return;

      const normalizedA = trimmedA.toLowerCase();

      headersB.forEach((headerB, indexB) => {
        if (!headerB) return;
        const trimmedB = String(headerB).trim();
        if (!trimmedB) return;

        const normalizedB = trimmedB.toLowerCase();

        if (normalizedA === normalizedB) {
          const alreadyExists = commonColumns.some((col) => col.name === trimmedA);

          if (!alreadyExists) {
            commonColumns.push({
              name: trimmedA,
              colIndex: this.getColumnLetter(indexA),
              indexA: indexA,
              indexB: indexB,
            });
          }
        }
      });
    });

    return commonColumns;
  }

  getColumnLetter(index) {
    let letter = '';
    while (index >= 0) {
      letter = String.fromCharCode((index % 26) + 65) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }

  setupSheetChangeListeners() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');

    if (sheetSelectA) {
      sheetSelectA.addEventListener('change', () => {
        this.updateHeaderRowDropdown('A');
        this.populateKeyColumnDropdown();
      });
    }

    if (sheetSelectB) {
      sheetSelectB.addEventListener('change', () => {
        this.updateHeaderRowDropdown('B');
        this.populateKeyColumnDropdown();
      });
    }

    if (headerRowA) {
      headerRowA.addEventListener('change', () => {
        this.populateKeyColumnDropdown();
      });
    }

    if (headerRowB) {
      headerRowB.addEventListener('change', () => {
        this.populateKeyColumnDropdown();
      });
    }
  }

  updateHeaderRowDropdown(type) {
    const sheetSelect = document.getElementById(`sheetSelect${type}`);
    const headerRowSelect = document.getElementById(`headerRow${type}`);
    const data = type === 'A' ? this.dataA : this.dataB;

    if (!sheetSelect || !headerRowSelect || !data) return;

    const selectedSheet = sheetSelect.value;
    const rowCount = data.sheets[selectedSheet]?.rowCount || 20;

    headerRowSelect.innerHTML = '';
    const maxRows = Math.min(rowCount, 50);
    for (let i = 1; i <= maxRows; i++) {
      const option = document.createElement('option');
      option.value = i;
      option.textContent = `Row ${i}`;
      headerRowSelect.appendChild(option);
    }

    headerRowSelect.value = '1';
  }

  findMatchingSheet() {
    if (!this.dataA.sheetNames || !this.dataB.sheetNames) return null;

    for (let sheetA of this.dataA.sheetNames) {
      if (this.dataB.sheetNames.includes(sheetA)) {
        return { sheetA, sheetB: sheetA };
      }
    }

    return null;
  }

  populateSheetDropdowns() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');

    if (!sheetSelectA || !sheetSelectB) {
      console.error('Sheet dropdown elements not found');
      return;
    }

    sheetSelectA.innerHTML = '';
    sheetSelectB.innerHTML = '';

    if (this.dataA && this.dataA.sheetNames) {
      this.dataA.sheetNames.forEach((sheetName) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelectA.appendChild(option);
      });
    }

    if (this.dataB && this.dataB.sheetNames) {
      this.dataB.sheetNames.forEach((sheetName) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelectB.appendChild(option);
      });
    }
  }

  compareSelectedSheets() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');
    const headerRowA = document.getElementById('headerRowA');
    const headerRowB = document.getElementById('headerRowB');
    const keyColumnSelect = document.getElementById('keyColumnSelect');

    if (!sheetSelectA || !sheetSelectB || !headerRowA || !headerRowB || !keyColumnSelect) {
      console.error('Selection elements not found');
      return;
    }

    const selectedSheetA = sheetSelectA.value;
    const selectedSheetB = sheetSelectB.value;
    const selectedHeaderA = parseInt(headerRowA.value) || 1;
    const selectedHeaderB = parseInt(headerRowB.value) || 1;
    const selectedKeyColumn = keyColumnSelect.value;

    if (!selectedSheetA || !selectedSheetB) {
      console.error('⚠️ Please select sheets to compare');
      return;
    }

    if (!selectedKeyColumn) {
      console.error('⚠️ Please select a Key Column or "Use Row Position"');
      return;
    }

    let sheetA = this.dataA.sheets[selectedSheetA]?.data || [];
    let sheetB = this.dataB.sheets[selectedSheetB]?.data || [];

    sheetA = this.adjustDataForHeaderRow(sheetA, selectedHeaderA);
    sheetB = this.adjustDataForHeaderRow(sheetB, selectedHeaderB);

    if (sheetA.length === 0 || sheetB.length === 0) {
      console.error('⚠️ Selected sheet is empty');
      return;
    }

    if (selectedKeyColumn === '__ROW_POSITION__') {
      console.log('🔢 Using row position matching');
      this.compareByRowPosition(sheetA, sheetB, selectedHeaderA, selectedHeaderB);
    } else {
      console.log(`🔑 Using key column: ${selectedKeyColumn}`);

      const diffEngine = new DiffEngine();
      const singleSheetDiff = diffEngine.compareSheets(sheetA, sheetB, selectedKeyColumn, selectedHeaderA, selectedHeaderB);

      singleSheetDiff.sheetName = `${selectedSheetA} vs ${selectedSheetB}`;
      singleSheetDiff.headerRowA = selectedHeaderA;
      singleSheetDiff.headerRowB = selectedHeaderB;

      this.renderUnifiedTable(singleSheetDiff, selectedKeyColumn);
    }
  }

  adjustDataForHeaderRow(data, headerRow) {
    if (!data || data.length === 0) return data;

    const headerIndex = headerRow - 1;

    if (headerIndex < 0 || headerIndex >= data.length) {
      console.warn(`Invalid header row ${headerRow}, using row 1`);
      return data;
    }

    return data.slice(headerIndex);
  }

  renderUnifiedTable(sheetDiff, keyColumn = 'A') {
    const container = document.getElementById('unifiedTableContainer');
    if (!container) {
      console.error('unifiedTableContainer element not found');
      return;
    }

    container.innerHTML = '';

    const table = document.createElement('table');
    table.className = 'unified-table diff-table';

    const thead = this.buildUnifiedHeader(sheetDiff);
    table.appendChild(thead);

    const tbody = this.buildUnifiedBody(sheetDiff, keyColumn);
    table.appendChild(tbody);

    const wrapper = document.createElement('div');
    wrapper.className = 'table-wrapper';
    wrapper.appendChild(table);

    container.appendChild(wrapper);

    this.collectChangedCells();
    this.setupCellClickNavigation();
  }

  setupCellClickNavigation() {
    const table = document.querySelector('#unifiedTableContainer .diff-table');
    if (!table) return;

    const tbody = table.querySelector('tbody');
    if (!tbody) return;

    tbody.addEventListener('click', (e) => {
      const clickedCell = e.target.closest('td');
      if (!clickedCell) return;

      const isChangedCell = clickedCell.classList.contains('cell-modified') || clickedCell.classList.contains('cell-added') || clickedCell.classList.contains('cell-deleted');

      if (!isChangedCell) return;

      const clickedRow = clickedCell.closest('tr');

      for (let i = 0; i < this.changedCells.length; i++) {
        const { row, cell } = this.changedCells[i];

        if (row === clickedRow && cell === clickedCell) {
          this.currentChangeIndex = i;
          this.updateNavigationUI();
          this.scrollToChange();
          return;
        }
      }
    });
  }

  getUnifiedColumns(sheetDiff) {
    const oldHeaders = sheetDiff.oldData[0] || {};
    const newHeaders = sheetDiff.newData[0] || {};

    const headerMap = new Map();

    const normalizeHeader = (header) => {
      return String(header || '')
        .trim()
        .toLowerCase();
    };

    Object.keys(oldHeaders).forEach((col) => {
      const content = String(oldHeaders[col] || '').trim();

      if (content) {
        const normalized = normalizeHeader(content);

        headerMap.set(normalized, {
          originalA: content,
          originalB: null,
          oldCol: col,
          newCol: null,
        });
      } else {
        headerMap.set(`__empty_old_${col}`, {
          originalA: '',
          originalB: null,
          oldCol: col,
          newCol: null,
        });
      }
    });

    Object.keys(newHeaders).forEach((col) => {
      const content = String(newHeaders[col] || '').trim();

      if (content) {
        const normalized = normalizeHeader(content);

        if (headerMap.has(normalized)) {
          const existing = headerMap.get(normalized);
          existing.originalB = content;
          existing.newCol = col;
        } else {
          headerMap.set(normalized, {
            originalA: null,
            originalB: content,
            oldCol: null,
            newCol: col,
          });
        }
      } else {
        const key = `__empty_new_${col}`;

        const oldEmptyKey = `__empty_old_${col}`;
        if (headerMap.has(oldEmptyKey)) {
          headerMap.get(oldEmptyKey).newCol = col;
          headerMap.set(`__empty_both_${col}`, headerMap.get(oldEmptyKey));
          headerMap.delete(oldEmptyKey);
        } else {
          headerMap.set(key, {
            originalA: null,
            originalB: '',
            oldCol: null,
            newCol: col,
          });
        }
      }
    });

    const result = [];
    const processedHeaders = new Set();

    Object.keys(newHeaders).forEach((newCol) => {
      for (let [normalized, mapping] of headerMap) {
        if (mapping.newCol === newCol && !processedHeaders.has(normalized)) {
          processedHeaders.add(normalized);

          const displayName = mapping.originalB || mapping.originalA || '(Blank Column)';

          result.push({
            header: displayName.startsWith('__empty_') ? '(Blank Column)' : displayName,
            oldCol: mapping.oldCol,
            newCol: mapping.newCol,
            type: mapping.oldCol && mapping.newCol ? 'normal' : mapping.oldCol ? 'deleted' : 'added',
          });
          break;
        }
      }
    });

    for (let [normalized, mapping] of headerMap) {
      if (!processedHeaders.has(normalized)) {
        const displayName = mapping.originalA || mapping.originalB || '(Blank Column)';

        result.push({
          header: displayName.startsWith('__empty_') ? '(Blank Column)' : displayName,
          oldCol: mapping.oldCol,
          newCol: mapping.newCol,
          type: 'deleted',
        });
      }
    }

    return result;
  }

  buildUnifiedHeader(sheetDiff) {
    const thead = document.createElement('thead');
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);

    const tr1 = document.createElement('tr');

    const th1Old = document.createElement('th');
    th1Old.className = 'index-col';
    th1Old.textContent = 'Old';
    th1Old.rowSpan = 2;
    tr1.appendChild(th1Old);

    const th1New = document.createElement('th');
    th1New.className = 'index-col';
    th1New.textContent = 'New';
    th1New.rowSpan = 2;
    tr1.appendChild(th1New);

    unifiedColumns.forEach((col) => {
      const th = document.createElement('th');

      let colLabel = col.newCol || col.oldCol;

      if (col.type === 'added') {
        th.className = 'col-added';
        colLabel = `+${col.newCol}`;
      } else if (col.type === 'deleted') {
        th.className = 'col-deleted';
        colLabel = `−${col.oldCol}`;
      }

      th.textContent = colLabel;
      tr1.appendChild(th);
    });

    const tr2 = document.createElement('tr');

    unifiedColumns.forEach((col) => {
      const th = document.createElement('th');

      if (col.type === 'added') {
        th.className = 'col-added';
      } else if (col.type === 'deleted') {
        th.className = 'col-deleted';
      }

      th.textContent = col.header;
      tr2.appendChild(th);
    });

    thead.appendChild(tr1);
    thead.appendChild(tr2);
    return thead;
  }

 /**
 * buildUnifiedBody() - ✅ 用 originalKey 版
 */
/**
 * buildUnifiedBody() - ✅ 最終修正版 (v5)
 */
buildUnifiedBody(sheetDiff, keyColumn = 'A') {
    const tbody = document.createElement('tbody');
    const allRows = this.getAllRows(sheetDiff, keyColumn);
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);
    const cellChanges = this.buildCellChangeMap(sheetDiff.differences);
    const rowChanges = this.buildRowChangeMap(sheetDiff.rowChanges);

    console.log('🔍 Row Changes Map:');
    rowChanges.forEach((change, key) => {
        console.log(`   Key: "${key}" → Type: ${change.type}`);
    });

    allRows.forEach((rowInfo) => {
        const tr = document.createElement('tr');

        console.log(`🔑 Processing row: key="${rowInfo.key}", originalKey="${rowInfo.originalKey}", oldIndex=${rowInfo.oldIndex}, newIndex=${rowInfo.newIndex}`);

        // ✅ 方法 1: 直接用 oldIndex/newIndex 判斷
        let rowType = null;
        
        if (rowInfo.oldIndex !== null && rowInfo.newIndex !== null) {
            rowType = 'matched';
        } else if (rowInfo.oldIndex !== null && rowInfo.newIndex === null) {
            rowType = 'deleted';
        } else if (rowInfo.oldIndex === null && rowInfo.newIndex !== null) {
            rowType = 'added';
        }

        console.log(`   ✅ Row type: ${rowType}`);

        // Apply row-level styling
        if (rowType === 'added') {
            tr.className = 'row-added';
        } else if (rowType === 'deleted') {
            tr.className = 'row-deleted';
        }

        // Old Index
        const tdOldIdx = document.createElement('td');
        tdOldIdx.className = 'index-cell old-idx';
        tdOldIdx.textContent = rowInfo.oldIndex !== null ? rowInfo.oldIndex : '-';
        tr.appendChild(tdOldIdx);

        // New Index
        const tdNewIdx = document.createElement('td');
        tdNewIdx.className = 'index-cell new-idx';
        tdNewIdx.textContent = rowInfo.newIndex !== null ? rowInfo.newIndex : '-';
        tr.appendChild(tdNewIdx);

        // Render cells
        unifiedColumns.forEach((col) => {
            const td = document.createElement('td');

            const oldValue = col.oldCol ? rowInfo.oldRow?.[col.oldCol] : null;
            const newValue = col.newCol ? rowInfo.newRow?.[col.newCol] : null;

            const cellKey = `${rowInfo.newIndex || rowInfo.oldIndex}-${col.header}`;
            const cellDiff = cellChanges.get(cellKey);

            if (cellDiff) {
                // Cell was modified
                td.className = 'cell-modified';
                td.innerHTML = `
                    <div class="cell-value-change">
                        <span class="old-value">${this.formatValue(cellDiff.oldValue)}</span>
                        <span class="value-separator">→</span>
                        <span class="new-value">${this.formatValue(cellDiff.newValue)}</span>
                    </div>
                `;
            } else if (col.type === 'added') {
                // Column was added
                td.className = 'cell-added';
                td.innerHTML = this.formatValue(newValue);
            } else if (col.type === 'deleted') {
                // Column was deleted
                td.className = 'cell-deleted';
                td.innerHTML = this.formatValue(oldValue);
            } else if (rowType === 'deleted') {
                // Row was deleted
                td.className = 'cell-deleted';
                td.innerHTML = this.formatValue(oldValue);
            } else if (rowType === 'added') {
                // Row was added
                td.className = 'cell-added';
                td.innerHTML = this.formatValue(newValue);
            } else {
                // Cell unchanged
                td.className = 'cell-unchanged';
                td.innerHTML = this.formatValue(newValue || oldValue);
            }

            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    return tbody;
}

  /**
   * getAllRows() - ✅ 最終修正版
   */
/**
 * getAllRows() - ✅ Interleaved Sorting (v6 - FINAL)
 */
getAllRows(sheetDiff, keyColumn = 'A') {
    const rowMap = new Map();
    const headerRowA = sheetDiff.headerRowA || 1;
    const headerRowB = sheetDiff.headerRowB || 1;

    const oldDataRows = sheetDiff.oldData.slice(1);
    const newDataRows = sheetDiff.newData.slice(1);

    // Process Old File
    oldDataRows.forEach((row, index) => {
        const keyValue = String(row[keyColumn] || '').trim();
        const excelRowNumber = headerRowA + 1 + index;

        if (keyValue) {
            rowMap.set(keyValue, {
                key: keyValue,
                originalKey: keyValue,
                oldRow: row,
                oldIndex: excelRowNumber,
                newRow: null,
                newIndex: null,
            });
        } else {
            const uniqueKey = `__blank_old_${excelRowNumber}`;
            rowMap.set(uniqueKey, {
                key: `(empty)`,
                originalKey: null,
                oldRow: row,
                oldIndex: excelRowNumber,
                newRow: null,
                newIndex: null,
                isBlankKey: true,
            });
        }
    });

    // Process New File
    newDataRows.forEach((row, index) => {
        const keyValue = String(row[keyColumn] || '').trim();
        const excelRowNumber = headerRowB + 1 + index;

        if (keyValue) {
            if (rowMap.has(keyValue)) {
                const existing = rowMap.get(keyValue);
                existing.newRow = row;
                existing.newIndex = excelRowNumber;
            } else {
                rowMap.set(keyValue, {
                    key: keyValue,
                    originalKey: keyValue,
                    oldRow: null,
                    oldIndex: null,
                    newRow: row,
                    newIndex: excelRowNumber,
                });
            }
        } else {
            const positionKey = `__blank_old_${excelRowNumber}`;

            if (rowMap.has(positionKey)) {
                const existing = rowMap.get(positionKey);
                const allColumnsMatch = this.areRowsIdentical(existing.oldRow, row);

                if (allColumnsMatch) {
                    existing.newRow = row;
                    existing.newIndex = excelRowNumber;
                } else {
                    const newKey = `__blank_new_${excelRowNumber}`;
                    rowMap.set(newKey, {
                        key: `(empty)`,
                        originalKey: null,
                        oldRow: null,
                        oldIndex: null,
                        newRow: row,
                        newIndex: excelRowNumber,
                        isBlankKey: true,
                    });
                }
            } else {
                const newKey = `__blank_new_${excelRowNumber}`;
                rowMap.set(newKey, {
                    key: `(empty)`,
                    originalKey: null,
                    oldRow: null,
                    oldIndex: null,
                    newRow: row,
                    newIndex: excelRowNumber,
                    isBlankKey: true,
                });
            }
        }
    });

    const rows = Array.from(rowMap.values());

    console.log('📊 Total rows before sorting:', rows.length);
    console.log('📊 Deleted rows:', rows.filter((r) => r.oldIndex && !r.newIndex).length);
    console.log('📊 Added rows:', rows.filter((r) => !r.oldIndex && r.newIndex).length);
    console.log('📊 Matched rows:', rows.filter((r) => r.oldIndex && r.newIndex).length);

    // ✅ FINAL: Interleaved Sorting (v7 - 真·最終版)
rows.sort((a, b) => {
    const getType = (row) => {
        if (row.oldIndex !== null && row.newIndex !== null) return 'matched';
        if (row.oldIndex === null && row.newIndex !== null) return 'added';
        if (row.oldIndex !== null && row.newIndex === null) return 'deleted';
        return 'unknown';
    };

    const typeA = getType(a);
    const typeB = getType(b);

    // ✅ Step 1: Group deleted rows by their position relative to new file
    const getDisplayPosition = (row, type) => {
        if (type === 'matched' || type === 'added') {
            // Matched/Added: Use newIndex directly
            return row.newIndex * 1000; // Scale up to leave room for deleted rows
        } else if (type === 'deleted') {
            // Deleted: Find the next new index after this old index
            // Insert deleted rows BEFORE the next new row
            
            // Find all matched/added rows
            const newIndexes = rows
                .filter(r => r.newIndex !== null)
                .map(r => r.newIndex)
                .sort((x, y) => x - y);
            
            // Find the next new index that comes after this deleted row's old position
            const nextNewIndex = newIndexes.find(idx => {
                // Find which old row this new index corresponds to
                const matchedRow = rows.find(r => r.newIndex === idx && r.oldIndex !== null);
                if (matchedRow && matchedRow.oldIndex > row.oldIndex) {
                    return true;
                }
                return false;
            });
            
            if (nextNewIndex !== undefined) {
                // Insert before the next new row
                return nextNewIndex * 1000 - 500 + (row.oldIndex % 1000);
            } else {
                // No next new row, append at the end
                return (Math.max(...newIndexes, 0) + 1) * 1000 + row.oldIndex;
            }
        }
        return Infinity;
    };

    const posA = getDisplayPosition(a, typeA);
    const posB = getDisplayPosition(b, typeB);

    if (posA !== posB) {
        return posA - posB;
    }

    // TIE-BREAKER
    const typePriority = {
        deleted: 0,
        matched: 1,
        added: 2,
        unknown: 3,
    };

    return typePriority[typeA] - typePriority[typeB];
});

    console.log('✅ Rows sorted with interleaved deleted rows');
    
    // Debug: Show final order
    console.log('🔍 Final row order:');
    rows.forEach((row, idx) => {
        const type = row.oldIndex && row.newIndex ? 'matched' : 
                     !row.oldIndex && row.newIndex ? 'added' : 'deleted';
        console.log(`  ${idx + 1}. [${type}] Key=${row.key}, Old=${row.oldIndex}, New=${row.newIndex}`);
    });

    return rows;
}

  areRowsIdentical(row1, row2) {
    if (!row1 || !row2) return false;

    const allKeys = new Set([...Object.keys(row1), ...Object.keys(row2)]);

    for (const key of allKeys) {
      const val1 = this.normalizeValue(row1[key]);
      const val2 = this.normalizeValue(row2[key]);

      if (val1 !== val2) {
        return false;
      }
    }

    return true;
  }

  normalizeValue(cell) {
    if (cell === null || cell === undefined) {
      return '';
    }

    let value = String(cell);
    value = value.replace(/[\u200B-\u200D\uFEFF]/g, '');
    value = value.trim();

    return value;
  }

  buildRowChangeMap(rowChanges) {
    const map = new Map();
    rowChanges.forEach((change) => {
      map.set(change.rowKey, change); // ✅ rowKey 應該係 ID value
    });
    return map;
  }

  buildCellChangeMap(differences) {
    const map = new Map();
    differences.forEach((diff) => {
      const key = `${diff.row}-${diff.header}`;
      map.set(key, diff);
    });
    return map;
  }

  formatValue(value) {
    if (value === null || value === undefined || value === '') {
      return '<em class="empty-cell">Blank</em>';
    }
    return String(value);
  }

  setupChangeNavigation() {
    const prevBtn = document.getElementById('prevChangeBtn');
    const nextBtn = document.getElementById('nextChangeBtn');

    if (prevBtn) {
      prevBtn.onclick = () => this.navigateToChange('prev');
    }

    if (nextBtn) {
      nextBtn.onclick = () => this.navigateToChange('next');
    }

    this.setupKeyboardShortcuts();
  }

  setupKeyboardShortcuts() {
    if (this.keyboardHandler) {
      document.removeEventListener('keydown', this.keyboardHandler);
    }

    this.keyboardHandler = (e) => {
      const isTyping = ['INPUT', 'TEXTAREA', 'SELECT'].includes(e.target.tagName);
      if (isTyping) return;

      const key = e.key.toLowerCase();

      if (key === 'p') {
        e.preventDefault();
        this.navigateToChange('prev');
      } else if (key === 'n') {
        e.preventDefault();
        this.navigateToChange('next');
      }
    };

    document.addEventListener('keydown', this.keyboardHandler);
  }

  collectChangedCells() {
    this.changedCells = [];
    this.currentChangeIndex = -1;

    const table = document.querySelector('#unifiedTableContainer .diff-table');
    if (!table) {
      this.updateNavigationUI();
      return;
    }

    const tbody = table.querySelector('tbody');
    if (!tbody) {
      this.updateNavigationUI();
      return;
    }

    const rows = tbody.querySelectorAll('tr');
    rows.forEach((row, rowIndex) => {
      const cells = row.querySelectorAll('td.cell-modified, td.cell-added, td.cell-deleted');
      cells.forEach((cell) => {
        this.changedCells.push({ row, cell });
      });
    });

    this.updateNavigationUI();
  }

  navigateToChange(direction) {
    if (this.changedCells.length === 0) return;

    if (direction === 'next') {
      this.currentChangeIndex = (this.currentChangeIndex + 1) % this.changedCells.length;
    } else if (direction === 'prev') {
      this.currentChangeIndex = (this.currentChangeIndex - 1 + this.changedCells.length) % this.changedCells.length;
    }

    this.updateNavigationUI();
    this.scrollToChange();
  }

  scrollToChange() {
    if (this.currentChangeIndex < 0 || this.currentChangeIndex >= this.changedCells.length) return;

    const { cell } = this.changedCells[this.currentChangeIndex];

    document.querySelectorAll('.cell-highlighted').forEach((c) => {
      c.classList.remove('cell-highlighted');
    });

    cell.classList.add('cell-highlighted');

    cell.scrollIntoView({ behavior: 'smooth', block: 'center' });

    setTimeout(() => {
      cell.classList.remove('cell-highlighted');
    }, 2000);
  }

  updateNavigationUI() {
    const counter = document.getElementById('changeCounter');
    const prevBtn = document.getElementById('prevChangeBtn');
    const nextBtn = document.getElementById('nextChangeBtn');

    if (!counter) return;

    const total = this.changedCells.length;
    const current = this.currentChangeIndex >= 0 ? this.currentChangeIndex + 1 : 0;

    counter.textContent = `${current} / ${total}`;

    if (prevBtn && nextBtn) {
      prevBtn.disabled = total === 0;
      nextBtn.disabled = total === 0;
    }
  }

  compareByRowPosition(sheetA, sheetB, headerRowA, headerRowB) {
    console.log('🔢 Comparing by row position...');

    const unifiedColumns = this.getUnifiedColumns({
      oldData: sheetA,
      newData: sheetB,
    });

    this.changedCells = [];

    const dataRowsA = sheetA.slice(1);
    const dataRowsB = sheetB.slice(1);

    const maxRows = Math.max(dataRowsA.length, dataRowsB.length);

    const differences = [];
    const rowChanges = [];

    for (let i = 0; i < maxRows; i++) {
      const rowA = dataRowsA[i] || {};
      const rowB = dataRowsB[i] || {};

      const excelRowA = headerRowA + 1 + i;
      const excelRowB = headerRowB + 1 + i;

      const rowExistsInA = i < dataRowsA.length;
      const rowExistsInB = i < dataRowsB.length;

      if (!rowExistsInA) {
        rowChanges.push({
          rowKey: `position-${i}`,
          type: 'added',
          newRowIndex: excelRowB,
          row: rowB,
        });
      } else if (!rowExistsInB) {
        rowChanges.push({
          rowKey: `position-${i}`,
          type: 'deleted',
          oldRowIndex: excelRowA,
          row: rowA,
        });
      } else {
        unifiedColumns.forEach(({ header, oldCol, newCol }) => {
          const valueA = oldCol ? (rowA[oldCol] ?? '') : '';
          const valueB = newCol ? (rowB[newCol] ?? '') : '';

          const normalizedA = String(valueA).trim();
          const normalizedB = String(valueB).trim();

          if (normalizedA !== normalizedB) {
            differences.push({
              row: excelRowB,
              header: header,
              oldCol: oldCol,
              newCol: newCol,
              oldValue: valueA,
              newValue: valueB,
            });
          }
        });
      }
    }

    const singleSheetDiff = {
      sheetName: 'Position-based Comparison',
      differences: differences,
      rowChanges: rowChanges,
      columnChanges: [],
      oldData: sheetA,
      newData: sheetB,
      headerRowA: headerRowA,
      headerRowB: headerRowB,
    };

    this.renderUnifiedTableByPosition(singleSheetDiff, headerRowA, headerRowB);
  }

  renderUnifiedTableByPosition(sheetDiff, headerRowA, headerRowB) {
    const container = document.getElementById('unifiedTableContainer');
    if (!container) {
      console.error('unifiedTableContainer element not found');
      return;
    }

    container.innerHTML = '';

    const table = document.createElement('table');
    table.className = 'unified-table diff-table';

    const thead = this.buildUnifiedHeader(sheetDiff);
    table.appendChild(thead);

    const tbody = this.buildUnifiedBodyByPosition(sheetDiff, headerRowA, headerRowB);
    table.appendChild(tbody);

    const wrapper = document.createElement('div');
    wrapper.className = 'table-wrapper';
    wrapper.appendChild(table);

    container.appendChild(wrapper);

    this.collectChangedCells();
    this.setupCellClickNavigation();
  }

  buildUnifiedBodyByPosition(sheetDiff, headerRowA, headerRowB) {
    const tbody = document.createElement('tbody');
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);
    const cellChanges = this.buildCellChangeMap(sheetDiff.differences);
    const rowChanges = this.buildRowChangeMap(sheetDiff.rowChanges);

    const dataRowsA = sheetDiff.oldData.slice(1);
    const dataRowsB = sheetDiff.newData.slice(1);
    const maxRows = Math.max(dataRowsA.length, dataRowsB.length);

    for (let i = 0; i < maxRows; i++) {
      const tr = document.createElement('tr');

      const rowA = dataRowsA[i];
      const rowB = dataRowsB[i];

      const excelRowA = headerRowA + 1 + i;
      const excelRowB = headerRowB + 1 + i;

      const rowExistsInA = i < dataRowsA.length;
      const rowExistsInB = i < dataRowsB.length;

      const rowChange = rowChanges.get(`position-${i}`);

      if (rowChange?.type === 'added') {
        tr.className = 'row-added';
      } else if (rowChange?.type === 'deleted') {
        tr.className = 'row-deleted';
      }

      const tdOldIdx = document.createElement('td');
      tdOldIdx.className = 'index-cell old-idx';
      tdOldIdx.textContent = rowExistsInA ? excelRowA : '-';
      tr.appendChild(tdOldIdx);

      const tdNewIdx = document.createElement('td');
      tdNewIdx.className = 'index-cell new-idx';
      tdNewIdx.textContent = rowExistsInB ? excelRowB : '-';
      tr.appendChild(tdNewIdx);

      unifiedColumns.forEach((col) => {
        const td = document.createElement('td');

        const oldValue = col.oldCol && rowA ? (rowA[col.oldCol] ?? null) : null;
        const newValue = col.newCol && rowB ? (rowB[col.newCol] ?? null) : null;

        const cellKey = `${excelRowB}-${col.header}`;
        const cellDiff = cellChanges.get(cellKey);

        if (cellDiff) {
          td.className = 'cell-modified';
          td.innerHTML = `
            <div class="cell-value-change">
              <span class="old-value">${this.formatValue(cellDiff.oldValue)}</span>
              <span class="value-separator">→</span>
              <span class="new-value">${this.formatValue(cellDiff.newValue)}</span>
            </div>
          `;
        } else if (col.type === 'added') {
          td.className = 'cell-added';
          td.innerHTML = this.formatValue(newValue);
        } else if (col.type === 'deleted') {
          td.className = 'cell-deleted';
          td.innerHTML = this.formatValue(oldValue);
        } else if (rowChange?.type === 'deleted') {
          td.className = 'cell-deleted';
          td.innerHTML = this.formatValue(oldValue);
        } else if (rowChange?.type === 'added') {
          td.className = 'cell-added';
          td.innerHTML = this.formatValue(newValue);
        } else {
          td.className = 'cell-unchanged';
          td.innerHTML = this.formatValue(newValue || oldValue);
        }

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    }

    return tbody;
  }
}

export default DiffViewer;