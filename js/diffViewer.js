/**
 * diffViewer.js
 * 顯示 Excel 比對結果 - UNIFIED TABLE with Old/New Index
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

  /**
   * ✅ 初始化：填入 dropdown + 自動選中同名 sheet
   */
  init(dataA, dataB, diffResults) {
    this.dataA = dataA;
    this.dataB = dataB;
    this.diffResults = diffResults;

    this.populateSheetDropdowns();

    const matchedSheet = this.findMatchingSheet();

    if (matchedSheet) {
      document.getElementById('sheetSelectA').value = matchedSheet.sheetA;
      document.getElementById('sheetSelectB').value = matchedSheet.sheetB;
      this.compareSelectedSheets();
    } else {
      console.log('⚠️ 冇同名 sheet，等用戶手動按 Compare');
    }

    this.setupChangeNavigation();
  }

  /**
   * ✅ 尋找同名 sheet
   */
  findMatchingSheet() {
    if (!this.dataA.sheetNames || !this.dataB.sheetNames) return null;

    for (let sheetA of this.dataA.sheetNames) {
      if (this.dataB.sheetNames.includes(sheetA)) {
        return { sheetA, sheetB: sheetA };
      }
    }

    return null;
  }

  /**
   * ✅ 填入 sheet names 到兩個 dropdown
   */
  populateSheetDropdowns() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');

    if (!sheetSelectA || !sheetSelectB) {
      console.error('找不到 sheet dropdown');
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

    console.log('✅ Sheet dropdowns 已填入');
  }

  /**
   * ✅ 當用戶點擊 "Compare" 按鈕時，比對選定的 sheets
   */
  compareSelectedSheets() {
    const sheetSelectA = document.getElementById('sheetSelectA');
    const sheetSelectB = document.getElementById('sheetSelectB');

    if (!sheetSelectA || !sheetSelectB) {
      console.error('找不到 sheet dropdown');
      return;
    }

    const selectedSheetA = sheetSelectA.value;
    const selectedSheetB = sheetSelectB.value;

    if (!selectedSheetA || !selectedSheetB) {
      alert('請選擇要比對的 Sheet');
      return;
    }

    console.log(`比對 Sheet: ${selectedSheetA} vs ${selectedSheetB}`);

    const sheetA = this.dataA.sheets[selectedSheetA]?.data || [];
    const sheetB = this.dataB.sheets[selectedSheetB]?.data || [];

    if (sheetA.length === 0 || sheetB.length === 0) {
      alert('選定的 Sheet 為空');
      return;
    }

    const diffEngine = new DiffEngine();
    const singleSheetDiff = diffEngine.compareSheets(sheetA, sheetB);
    singleSheetDiff.sheetName = `${selectedSheetA} vs ${selectedSheetB}`;

    this.renderUnifiedTable(singleSheetDiff);
  }

  /**
   * ✅ 渲染 Unified Table（有 Old/New Index）
   */
  renderUnifiedTable(sheetDiff) {
    const container = document.getElementById('unifiedTableContainer');
    if (!container) {
      console.error('找不到 unifiedTableContainer 容器');
      return;
    }

    container.innerHTML = '';

    const table = document.createElement('table');
    table.className = 'unified-table diff-table';

    // ✅ 建立 header（兩層）
    const thead = this.buildUnifiedHeader(sheetDiff);
    table.appendChild(thead);

    // ✅ 建立 body
    const tbody = this.buildUnifiedBody(sheetDiff);
    table.appendChild(tbody);

    const wrapper = document.createElement('div');
    wrapper.className = 'table-wrapper';
    wrapper.appendChild(table);

    container.appendChild(wrapper);

    this.collectChangedCells();

    // ✅ 新增：綁定 cell click event
    this.setupCellClickNavigation();
  }

  /**
   * ✅ 新增：綁定 cell click event，點擊後跳到最近嘅 change
   */
  setupCellClickNavigation() {
    const table = document.querySelector('#unifiedTableContainer .diff-table');
    if (!table) return;

    const tbody = table.querySelector('tbody');
    if (!tbody) return;

    tbody.addEventListener('click', (e) => {
      // 檢查係咪點擊咗 td
      const clickedCell = e.target.closest('td');
      if (!clickedCell) return;

      // 檢查係咪 changed cell
      const isChangedCell = clickedCell.classList.contains('cell-modified') || clickedCell.classList.contains('cell-added') || clickedCell.classList.contains('cell-deleted');

      if (!isChangedCell) {
        console.log('⚠️ 點擊的不是 changed cell');
        return;
      }

      // 搵出呢個 cell 喺 changedCells 入面嘅 index
      const clickedRow = clickedCell.closest('tr');

      for (let i = 0; i < this.changedCells.length; i++) {
        const { row, cell } = this.changedCells[i];

        if (row === clickedRow && cell === clickedCell) {
          console.log(`✅ 點擊咗 change #${i + 1}`);
          this.currentChangeIndex = i;
          this.updateNavigationUI();
          this.scrollToChange();
          return;
        }
      }

      console.log('⚠️ 搵唔到對應嘅 change');
    });
  }

  /**
   * 🔥 修正版：建立 Unified Column List（根據 Header 內容合併）
   */
  getUnifiedColumns(sheetDiff) {
    const oldHeaders = sheetDiff.oldData[0] || {};
    const newHeaders = sheetDiff.newData[0] || {};

    const headerMap = new Map(); // header content → { oldCol, newCol }

    // 1️⃣ 先處理 File A 嘅 headers
    Object.keys(oldHeaders).forEach((col) => {
      const content = String(oldHeaders[col] || '').trim();
      if (content) {
        headerMap.set(content, { oldCol: col, newCol: null });
      } else {
        // 空 header，用 column letter 做 key
        headerMap.set(`__empty_old_${col}`, { oldCol: col, newCol: null });
      }
    });

    // 2️⃣ 再處理 File B 嘅 headers
    Object.keys(newHeaders).forEach((col) => {
      const content = String(newHeaders[col] || '').trim();

      if (content) {
        if (headerMap.has(content)) {
          // ✅ 搵到同名 header
          headerMap.get(content).newCol = col;
        } else {
          // ✅ File B 獨有嘅 header
          headerMap.set(content, { oldCol: null, newCol: col });
        }
      } else {
        // ✅ File B 嘅空欄
        const key = `__empty_new_${col}`;

        // 檢查係咪 File A 都有同位置嘅空欄
        const oldEmptyKey = `__empty_old_${col}`;
        if (headerMap.has(oldEmptyKey)) {
          headerMap.get(oldEmptyKey).newCol = col;
          // 改返個 key
          headerMap.set(`__empty_both_${col}`, headerMap.get(oldEmptyKey));
          headerMap.delete(oldEmptyKey);
        } else {
          headerMap.set(key, { oldCol: null, newCol: col });
        }
      }
    });

    // 3️⃣ 轉換成 array（按照 File B 嘅 column order）
    const result = [];
    const processedHeaders = new Set();

    // 先按 File B 嘅順序
    Object.keys(newHeaders).forEach((newCol) => {
      for (let [header, mapping] of headerMap) {
        if (mapping.newCol === newCol && !processedHeaders.has(header)) {
          processedHeaders.add(header);
          result.push({
            header: header.startsWith('__empty_') ? '(Blank Column)' : header,
            oldCol: mapping.oldCol,
            newCol: mapping.newCol,
            type: mapping.oldCol && mapping.newCol ? 'normal' : mapping.oldCol ? 'deleted' : 'added',
          });
          break;
        }
      }
    });

    // 再加入 File A 獨有嘅（deleted columns）
    for (let [header, mapping] of headerMap) {
      if (!processedHeaders.has(header)) {
        result.push({
          header: header.startsWith('__empty_') ? '(Blank Column)' : header,
          oldCol: mapping.oldCol,
          newCol: mapping.newCol,
          type: 'deleted',
        });
      }
    }

    console.log('📋 Unified Columns:', result);
    return result;
  }

  /**
   * 🔥 修正版：建立兩層 Header
   */
  buildUnifiedHeader(sheetDiff) {
    const thead = document.createElement('thead');
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);

    // ✅ 第1層：欄位名稱 (A, +B, C, -D...)
    const tr1 = document.createElement('tr');

    // Old Index
    const th1Old = document.createElement('th');
    th1Old.className = 'index-col';
    th1Old.textContent = 'Old';
    th1Old.rowSpan = 2;
    tr1.appendChild(th1Old);

    // New Index
    const th1New = document.createElement('th');
    th1New.className = 'index-col';
    th1New.textContent = 'New';
    th1New.rowSpan = 2;
    tr1.appendChild(th1New);

    // 資料欄（A, +B, C, -D...）
    unifiedColumns.forEach((col) => {
      const th = document.createElement('th');

      let colLabel = col.newCol || col.oldCol; // 優先用 newCol

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

    // ✅ 第2層：欄位內容（Header 內容）
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
   * 🔥 修正版：建立 Body（用 unifiedColumns）
   */
  /**
   * 🔥 修正版：建立 Body
   */
  buildUnifiedBody(sheetDiff) {
    const tbody = document.createElement('tbody');
    const allRows = this.getAllRows(sheetDiff);
    const unifiedColumns = this.getUnifiedColumns(sheetDiff);
    const cellChanges = this.buildCellChangeMap(sheetDiff.differences);
    const rowChanges = this.buildRowChangeMap(sheetDiff.rowChanges);

    allRows.forEach((rowInfo) => {
      const tr = document.createElement('tr');

      const rowChange = rowChanges.get(rowInfo.key);

      if (rowChange?.type === 'added') {
        tr.className = 'row-added';
      } else if (rowChange?.type === 'deleted') {
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

      // 資料欄
      unifiedColumns.forEach((col) => {
        const td = document.createElement('td');

        const oldValue = col.oldCol ? rowInfo.oldRow?.[col.oldCol] : null;
        const newValue = col.newCol ? rowInfo.newRow?.[col.newCol] : null;

        // 🔥 用 header content 做 key match
        const cellKey = `${rowInfo.oldIndex || rowInfo.newIndex}-${col.header}`;
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
    });

    return tbody;
  }

  /**
   * ✅ 取得所有行（Union of A & B，用 A 欄做 key）
   */
  getAllRows(sheetDiff) {
    const rowMap = new Map();

    sheetDiff.oldData.slice(1).forEach((row, index) => {
      const key = String(row.A || '').trim() || `old-${index}`;
      rowMap.set(key, {
        key: key,
        oldRow: row,
        oldIndex: index + 2,
        newRow: null,
        newIndex: null,
      });
    });

    sheetDiff.newData.slice(1).forEach((row, index) => {
      const key = String(row.A || '').trim() || `new-${index}`;

      if (rowMap.has(key)) {
        const existing = rowMap.get(key);
        existing.newRow = row;
        existing.newIndex = index + 2;
      } else {
        rowMap.set(key, {
          key: key,
          oldRow: null,
          oldIndex: null,
          newRow: row,
          newIndex: index + 2,
        });
      }
    });

    return Array.from(rowMap.values());
  }

  /**
   * ✅ 建立 Row Change Map
   */
  buildRowChangeMap(rowChanges) {
    const map = new Map();
    rowChanges.forEach((change) => {
      map.set(change.rowKey, change);
    });
    return map;
  }

  /**
   * ✅ 建立 Cell Change Map
   */
  buildCellChangeMap(differences) {
    const map = new Map();
    differences.forEach((diff) => {
      // ✅ 用 "rowIndex-headerContent" 做 key
      const key = `${diff.row}-${diff.header}`;
      map.set(key, diff);
    });
    return map;
  }

  /**
   * ✅ 格式化值
   */
  formatValue(value) {
    if (value === null || value === undefined || value === '') {
      return '<em class="empty-cell">Blank</em>';
    }
    return String(value);
  }

  /**
   * ✅ 綁定 Change Navigation 按鈕
   */
  setupChangeNavigation() {
    const prevBtn = document.getElementById('prevChangeBtn');
    const nextBtn = document.getElementById('nextChangeBtn');

    if (prevBtn) {
      prevBtn.onclick = () => this.navigateToChange('prev');
    }

    if (nextBtn) {
      nextBtn.onclick = () => this.navigateToChange('next');
    }

    // ✅ 加入 Keyboard Shortcuts
    this.setupKeyboardShortcuts();
  }

  /**
   * ✅ 新增：設定鍵盤快捷鍵
   */
  setupKeyboardShortcuts() {
    // 移除舊 listener（避免重複綁定）
    if (this.keyboardHandler) {
      document.removeEventListener('keydown', this.keyboardHandler);
    }

    // 建立新 listener
    this.keyboardHandler = (e) => {
      // 如果用戶正在輸入（input/textarea），唔觸發快捷鍵
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

    // 綁定到 document
    document.addEventListener('keydown', this.keyboardHandler);

    console.log('⌨️ Keyboard shortcuts enabled: P = Previous, N = Next');
  }

  /**
   * ✅ 收集所有 changed cells
   */
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

    console.log(`📍 收集到 ${this.changedCells.length} 個變更`);
    this.updateNavigationUI();
  }

  /**
   * ✅ Navigate to change
   */
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

  /**
   * ✅ 滾動到當前變更
   */
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

  /**
   * ✅ 更新 navigation UI
   */
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
}

export default DiffViewer;
