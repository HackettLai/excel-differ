// diffViewer.js - Side-by-side 表格差異顯示

const DiffViewer = {
  currentDiffResult: null,
  currentFileA: null,
  currentFileB: null,
  currentSheet: null,
  currentStatus: null,
  isSyncing: false,
  syncEnabled: true,
  wrapperA: null,
  wrapperB: null,
  HEADER_HEIGHT: 38, // 👈 新增：sticky header 高度

  initSyncButton() {
    const syncBtn = document.getElementById('syncToggle');
    if (!syncBtn) return;

    syncBtn.addEventListener('click', () => {
      this.toggleSync();
    });
  },

  toggleSync() {
    this.syncEnabled = !this.syncEnabled;
    const syncBtn = document.getElementById('syncToggle');

    if (this.syncEnabled) {
      syncBtn.classList.add('active');
      syncBtn.querySelector('.sync-text').textContent = '同步滾動';
      this.realignScroll();
    } else {
      syncBtn.classList.remove('active');
      syncBtn.querySelector('.sync-text').textContent = '已關閉';
    }
  },

  // 👇 修改：考慮 header 高度
  getFirstVisibleRowNumber(wrapper) {
    if (!wrapper) return null;
    
    const table = wrapper.querySelector('table');
    if (!table) return null;

    const tbody = table.querySelector('tbody');
    if (!tbody) return null;

    const rows = tbody.querySelectorAll('tr');
    if (rows.length === 0) return null;

    // 👇 加上 header 高度的偏移
    const scrollTop = wrapper.scrollTop + this.HEADER_HEIGHT;
    
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowTop = row.offsetTop;
      const rowBottom = rowTop + row.offsetHeight;
      
      // 👇 當 row 底部超過可視區域頂部（含 header）
      if (rowBottom > scrollTop) {
        const rowHeader = row.querySelector('.row-header');
        if (rowHeader) {
          return parseInt(rowHeader.textContent);
        }
      }
    }
    
    return null;
  },

  // 👇 修改：滾動時減去 header 高度
  scrollToRowNumber(wrapper, rowNumber, smooth = false) {
    if (!wrapper) return false;
    
    const table = wrapper.querySelector('table');
    if (!table) return false;

    const tbody = table.querySelector('tbody');
    if (!tbody) return false;

    const rows = tbody.querySelectorAll('tr');
    
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowHeader = row.querySelector('.row-header');
      
      if (rowHeader && parseInt(rowHeader.textContent) === rowNumber) {
        // 👇 減去 header 高度，讓 row 出現在 header 下方
        const targetScrollTop = Math.max(0, row.offsetTop - this.HEADER_HEIGHT);
        
        if (smooth) {
          wrapper.scrollTo({
            top: targetScrollTop,
            behavior: 'smooth'
          });
        } else {
          wrapper.scrollTop = targetScrollTop;
        }
        return true;
      }
    }
    
    return false;
  },

  realignScroll() {
    if (!this.wrapperA || !this.wrapperB) return;

    const rowNumberA = this.getFirstVisibleRowNumber(this.wrapperA);
    const rowNumberB = this.getFirstVisibleRowNumber(this.wrapperB);
    
    if (!rowNumberA && !rowNumberB) return;
    
    const targetRowNumber = Math.min(
      rowNumberA || Infinity,
      rowNumberB || Infinity
    );
    
    this.isSyncing = true;
    
    this.scrollToRowNumber(this.wrapperA, targetRowNumber, true);
    this.scrollToRowNumber(this.wrapperB, targetRowNumber, true);
    
    const avgScrollLeft = (this.wrapperA.scrollLeft + this.wrapperB.scrollLeft) / 2;
    
    this.wrapperA.scrollTo({
      left: avgScrollLeft,
      behavior: 'smooth'
    });
    this.wrapperB.scrollTo({
      left: avgScrollLeft,
      behavior: 'smooth'
    });
    
    setTimeout(() => {
      this.isSyncing = false;
    }, 500);
  },

  syncScrollByRowNumber(sourceWrapper, targetWrapper) {
    const sourceRowNumber = this.getFirstVisibleRowNumber(sourceWrapper);
    
    if (!sourceRowNumber) return;
    
    const found = this.scrollToRowNumber(targetWrapper, sourceRowNumber, false);
    
    if (!found) return;
    
    targetWrapper.scrollLeft = sourceWrapper.scrollLeft;
  },

  show(diffResult, fileA, fileB, sheetName, status = 'modified', viewSide = null) {
    this.currentDiffResult = diffResult;
    this.currentFileA = fileA;
    this.currentFileB = fileB;
    this.currentSheet = sheetName;
    this.currentStatus = status;

    SummaryView.hide();

    document.getElementById('diffSection').style.display = 'block';

    document.getElementById('currentSheetTitle').textContent = `Sheet: ${sheetName}`;

    document.getElementById('fileAName').textContent = diffResult.fileA;
    document.getElementById('fileBName').textContent = diffResult.fileB;

    NavBar.init(diffResult, fileA, fileB, sheetName, (name, status, side) => {
      this.show(diffResult, fileA, fileB, name, status, side);
    });

    this.renderTables(sheetName, status, viewSide);
  },

  setupSyncScroll() {
    this.wrapperA = document.querySelector('#diffPaneA .table-wrapper');
    this.wrapperB = document.querySelector('#diffPaneB .table-wrapper');

    if (!this.wrapperA || !this.wrapperB) {
      console.warn('❌ 找不到滾動容器');
      return;
    }

    const checkScrollable = (wrapper, name) => {
      const pane = wrapper.closest('.diff-pane');
      const canScroll = wrapper.scrollHeight > wrapper.clientHeight || wrapper.scrollWidth > wrapper.clientWidth;

      if (canScroll) {
        wrapper.classList.add('scrollable');
        pane.classList.add('has-scrollable');
      } else {
        wrapper.classList.remove('scrollable');
        pane.classList.remove('has-scrollable');
      }
    };

    checkScrollable(this.wrapperA, 'wrapperA');
    checkScrollable(this.wrapperB, 'wrapperB');

    const newWrapperA = this.wrapperA.cloneNode(true);
    const newWrapperB = this.wrapperB.cloneNode(true);
    
    this.wrapperA.parentNode.replaceChild(newWrapperA, this.wrapperA);
    this.wrapperB.parentNode.replaceChild(newWrapperB, this.wrapperB);
    
    this.wrapperA = newWrapperA;
    this.wrapperB = newWrapperB;

    let rafA = null;
    let rafB = null;

    this.wrapperA.addEventListener('scroll', () => {
      if (this.isSyncing || !this.syncEnabled) return;

      if (rafA) cancelAnimationFrame(rafA);

      rafA = requestAnimationFrame(() => {
        this.isSyncing = true;
        this.syncScrollByRowNumber(this.wrapperA, this.wrapperB);
        
        requestAnimationFrame(() => {
          this.isSyncing = false;
        });
      });
    }, { passive: true });

    this.wrapperB.addEventListener('scroll', () => {
      if (this.isSyncing || !this.syncEnabled) return;

      if (rafB) cancelAnimationFrame(rafB);

      rafB = requestAnimationFrame(() => {
        this.isSyncing = true;
        this.syncScrollByRowNumber(this.wrapperB, this.wrapperA);
        
        requestAnimationFrame(() => {
          this.isSyncing = false;
        });
      });
    }, { passive: true });

    console.log('✅ 同步滾動已設置！');
  },

  renderTables(sheetName, status, viewSide) {
    const tableA = document.getElementById('tableA');
    const tableB = document.getElementById('tableB');

    if (status === 'added') {
      this.renderEmptyTable(tableA, 'Sheet 在 File A 中不存在');
      this.renderTable(tableB, sheetName, this.currentFileB, null, 'added');
    } else if (status === 'removed') {
      this.renderTable(tableA, sheetName, this.currentFileA, null, 'removed');
      this.renderEmptyTable(tableB, 'Sheet 在 File B 中不存在');
    } else if (status === 'renamed') {
      const rename = this.currentDiffResult.sheetChanges.renamed.find((r) => r.to === sheetName);

      if (rename) {
        const cellDiff = this.currentDiffResult.cellDiffs[sheetName];
        this.renderTable(tableA, rename.from, this.currentFileA, cellDiff, 'comparison');
        this.renderTable(tableB, rename.to, this.currentFileB, cellDiff, 'comparison');
      }
    } else {
      const cellDiff = this.currentDiffResult.cellDiffs[sheetName];
      this.renderTable(tableA, sheetName, this.currentFileA, cellDiff, 'comparison');
      this.renderTable(tableB, sheetName, this.currentFileB, cellDiff, 'comparison');
    }

    setTimeout(() => {
      this.setupSyncScroll();
    }, 100);
  },

  renderEmptyTable(tableElement, message) {
    tableElement.innerHTML = `
            <tr>
                <td style="text-align: center; padding: 40px; color: #999; font-style: italic;">
                    ${message}
                </td>
            </tr>
        `;
  },

  renderTable(tableElement, sheetName, parsedFile, cellDiff, mode) {
    const sheet = parsedFile.sheets.find((s) => s.name === sheetName);

    if (!sheet) {
      this.renderEmptyTable(tableElement, '找不到此 Sheet');
      return;
    }

    const data = ExcelParser.normalizeData(sheet.data);

    let diffMap = null;
    if (cellDiff && mode === 'comparison') {
      diffMap = DiffEngine.createCellDiffMap(cellDiff.changes);
    }

    tableElement.innerHTML = '';

    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    const emptyTh = document.createElement('th');
    emptyTh.textContent = '';
    headerRow.appendChild(emptyTh);

    const maxCols = data.length > 0 ? data[0].length : 0;
    for (let c = 0; c < maxCols; c++) {
      const th = document.createElement('th');
      th.textContent = ExcelParser.getColumnName(c);
      headerRow.appendChild(th);
    }

    thead.appendChild(headerRow);
    tableElement.appendChild(thead);

    const tbody = document.createElement('tbody');

    data.forEach((row, rowIndex) => {
      const tr = document.createElement('tr');

      const rowHeaderTd = document.createElement('td');
      rowHeaderTd.className = 'row-header';
      rowHeaderTd.textContent = rowIndex + 1;
      tr.appendChild(rowHeaderTd);

      row.forEach((cellValue, colIndex) => {
        const td = document.createElement('td');

        td.textContent = cellValue || '';

        if (cellValue === '' || cellValue === null || cellValue === undefined) {
          td.classList.add('cell-empty');
          td.textContent = '(空)';
        }

        if (diffMap) {
          const diff = DiffEngine.getCellDiff(rowIndex, colIndex, diffMap);

          if (diff) {
            td.classList.add(`cell-${diff.type}`);

            let tooltipText = '';
            if (diff.type === 'modified') {
              tooltipText = `舊值: ${diff.oldValue}\n新值: ${diff.newValue}`;
            } else if (diff.type === 'added') {
              tooltipText = `新增: ${diff.newValue}`;
            } else if (diff.type === 'removed') {
              tooltipText = `刪除: ${diff.oldValue}`;
            }
            td.title = tooltipText;
          } else {
            td.classList.add('cell-unchanged');
          }
        } else if (mode === 'added') {
          td.classList.add('cell-added');
        } else if (mode === 'removed') {
          td.classList.add('cell-removed');
        } else {
          td.classList.add('cell-unchanged');
        }

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    tableElement.appendChild(tbody);
  },

  hide() {
    document.getElementById('diffSection').style.display = 'none';
  },
};

document.addEventListener('DOMContentLoaded', () => {
  DiffViewer.initSyncButton();
});