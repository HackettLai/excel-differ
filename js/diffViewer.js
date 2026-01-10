// diffViewer.js - 差異檢視器模組

const DiffViewer = {
  currentDiffResult: null,
  currentFileA: null,
  currentFileB: null,
  currentSheet: null,
  currentStatus: null,
  isSyncingVertical: false,
  syncEnabled: true,
  wrapperA: null,
  wrapperB: null,
  HEADER_HEIGHT: 38,
  verticalSyncTimeoutId: null,
  tooltipElement: null,
  currentTooltipCell: null,

  initSyncButton() {
    const syncBtn = document.getElementById('syncToggle');
    if (!syncBtn) return;

    syncBtn.addEventListener('click', () => {
      this.toggleSync();
    });
  },

  initTooltip() {
    this.tooltipElement = document.getElementById('customTooltip');
    if (!this.tooltipElement) {
      console.warn('❌ 找不到 tooltip 元素');
    } else {
      console.log('✅ Tooltip 元素找到了！');

      document.addEventListener('mousemove', (e) => {
        if (this.currentTooltipCell && this.tooltipElement.classList.contains('visible')) {
          this.updateTooltipPosition(e.clientX, e.clientY);
        }
      });
    }
  },

  updateTooltipPosition(x, y) {
    if (!this.tooltipElement) return;

    const tooltipRect = this.tooltipElement.getBoundingClientRect();
    const padding = 10;

    let left = x + padding;
    let top = y + padding;

    if (left + tooltipRect.width > window.innerWidth) {
      left = x - tooltipRect.width - padding;
    }

    if (top + tooltipRect.height > window.innerHeight) {
      top = y - tooltipRect.height - padding;
    }

    this.tooltipElement.style.left = `${left}px`;
    this.tooltipElement.style.top = `${top}px`;
  },

  showTooltip(text, x, y) {
    if (!this.tooltipElement || !text) return;

    this.tooltipElement.textContent = text;
    this.tooltipElement.classList.add('visible');
    this.updateTooltipPosition(x, y);
  },

  hideTooltip() {
    if (!this.tooltipElement) return;
    this.tooltipElement.classList.remove('visible');
    this.currentTooltipCell = null;
  },

  toggleSync() {
    this.syncEnabled = !this.syncEnabled;
    const syncBtn = document.getElementById('syncToggle');

    if (this.syncEnabled) {
      syncBtn.classList.add('active');
      syncBtn.querySelector('.sync-text').textContent = 'Sync Scroll';

      requestAnimationFrame(() => {
        this.realignScroll();
      });
    } else {
      syncBtn.classList.remove('active');
      syncBtn.querySelector('.sync-text').textContent = 'Disabled';
      this.isSyncingVertical = false;
    }
  },

  getFirstVisibleRowNumber(wrapper) {
    if (!wrapper) return null;

    const table = wrapper.querySelector('table');
    if (!table) return null;

    const tbody = table.querySelector('tbody');
    if (!tbody) return null;

    const rows = tbody.querySelectorAll('tr');
    if (rows.length === 0) return null;

    const scrollTop = wrapper.scrollTop + this.HEADER_HEIGHT;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowTop = row.offsetTop;
      const rowBottom = rowTop + row.offsetHeight;

      if (rowBottom > scrollTop) {
        const rowHeader = row.querySelector('.row-header');
        if (rowHeader) {
          return parseInt(rowHeader.textContent);
        }
      }
    }

    return null;
  },

  scrollToRowNumber(wrapper, rowNumber) {
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
        const targetScrollTop = Math.max(0, row.offsetTop - this.HEADER_HEIGHT);
        wrapper.scrollTop = targetScrollTop;
        return true;
      }
    }

    return false;
  },

  realignScroll() {
    if (!this.wrapperA || !this.wrapperB) return;

    const rowNumberA = this.getFirstVisibleRowNumber(this.wrapperA);
    const rowNumberB = this.getFirstVisibleRowNumber(this.wrapperB);

    if (!rowNumberA && !rowNumberB) {
      return;
    }

    const targetRowNumber = Math.min(rowNumberA || Infinity, rowNumberB || Infinity);

    this.isSyncingVertical = true;

    this.scrollToRowNumber(this.wrapperA, targetRowNumber);
    this.scrollToRowNumber(this.wrapperB, targetRowNumber);

    const avgScrollLeft = (this.wrapperA.scrollLeft + this.wrapperB.scrollLeft) / 2;
    this.wrapperA.scrollLeft = avgScrollLeft;
    this.wrapperB.scrollLeft = avgScrollLeft;

    if (this.verticalSyncTimeoutId) {
      clearTimeout(this.verticalSyncTimeoutId);
    }
    this.verticalSyncTimeoutId = setTimeout(() => {
      this.isSyncingVertical = false;
      this.verticalSyncTimeoutId = null;
    }, 150);
  },

  syncVertical(sourceWrapper, targetWrapper) {
    const sourceRowNumber = this.getFirstVisibleRowNumber(sourceWrapper);

    if (!sourceRowNumber) return;

    this.scrollToRowNumber(targetWrapper, sourceRowNumber);
  },

  syncHorizontal(sourceWrapper, targetWrapper) {
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

    // 👇 移除 cloneNode，直接重新綁定事件
    // 先移除舊的事件監聽器（如果有的話）
    const newWrapperA = this.wrapperA.cloneNode(true);
    const newWrapperB = this.wrapperB.cloneNode(true);

    this.wrapperA.parentNode.replaceChild(newWrapperA, this.wrapperA);
    this.wrapperB.parentNode.replaceChild(newWrapperB, this.wrapperB);

    this.wrapperA = newWrapperA;
    this.wrapperB = newWrapperB;

    this.wrapperA.addEventListener(
      'scroll',
      () => {
        if (!this.syncEnabled) return;

        this.syncHorizontal(this.wrapperA, this.wrapperB);

        if (this.isSyncingVertical) return;

        this.isSyncingVertical = true;
        this.syncVertical(this.wrapperA, this.wrapperB);

        if (this.verticalSyncTimeoutId) {
          clearTimeout(this.verticalSyncTimeoutId);
        }
        this.verticalSyncTimeoutId = setTimeout(() => {
          this.isSyncingVertical = false;
          this.verticalSyncTimeoutId = null;
        }, 100);
      },
      { passive: true }
    );

    this.wrapperB.addEventListener(
      'scroll',
      () => {
        if (!this.syncEnabled) return;

        this.syncHorizontal(this.wrapperB, this.wrapperA);

        if (this.isSyncingVertical) return;

        this.isSyncingVertical = true;
        this.syncVertical(this.wrapperB, this.wrapperA);

        if (this.verticalSyncTimeoutId) {
          clearTimeout(this.verticalSyncTimeoutId);
        }
        this.verticalSyncTimeoutId = setTimeout(() => {
          this.isSyncingVertical = false;
          this.verticalSyncTimeoutId = null;
        }, 100);
      },
      { passive: true }
    );

    console.log('✅ 同步滾動已設置（垂直/水平分離版）！');
  },

  renderTables(sheetName, status, viewSide) {
    const tableA = document.getElementById('tableA');
    const tableB = document.getElementById('tableB');

    if (status === 'added') {
      this.renderEmptyTable(tableA, 'Sheet does not exist in File A');
      this.renderTable(tableB, sheetName, this.currentFileB, null, 'added');
    } else if (status === 'removed') {
      this.renderTable(tableA, sheetName, this.currentFileA, null, 'removed');
      this.renderEmptyTable(tableB, 'Sheet does not exist in File B');
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

    // 👇 修改：先渲染表格，再設置同步滾動
    setTimeout(() => {
      this.setupSyncScroll();
      // 👇 在設置同步滾動後，重新綁定 tooltip 事件
      this.rebindTooltipEvents();
    }, 100);
  },

  // 👇 新增：重新綁定 tooltip 事件
  rebindTooltipEvents() {
    const self = this;

    // 找到所有有 data-tooltip 的儲存格
    document.querySelectorAll('td[data-tooltip]').forEach((td) => {
      // 移除舊的事件（避免重複綁定）
      td.replaceWith(td.cloneNode(true));
    });

    // 重新綁定事件
    document.querySelectorAll('td[data-tooltip]').forEach((td) => {
      td.addEventListener('mouseenter', function (e) {
        const text = this.dataset.tooltip;
        if (text) {
          console.log('✅ mouseenter 觸發:', text);
          self.currentTooltipCell = this;
          self.showTooltip(text, e.clientX, e.clientY);
        }
      });

      td.addEventListener('mouseleave', function () {
        console.log('✅ mouseleave 觸發');
        self.hideTooltip();
      });
    });

    console.log('✅ Tooltip 事件已重新綁定，共', document.querySelectorAll('td[data-tooltip]').length, '個儲存格');
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
      this.renderEmptyTable(tableElement, 'This Sheet cannot be found.');
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
          td.textContent = '(Empty)';
        }

        if (diffMap) {
          const diff = DiffEngine.getCellDiff(rowIndex, colIndex, diffMap);

          if (diff) {
            td.classList.add(`cell-${diff.type}`);

            let tooltipText = '';
            if (diff.type === 'modified') {
              tooltipText = `Old Value: ${diff.oldValue}\nNew Value: ${diff.newValue}`;
            } else if (diff.type === 'added') {
              tooltipText = `Added: ${diff.newValue}`;
            } else if (diff.type === 'removed') {
              tooltipText = `Removed: ${diff.oldValue}`;
            }

            // 👇 只設置 data-tooltip，事件在 rebindTooltipEvents() 中統一綁定
            td.dataset.tooltip = tooltipText;
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
    this.hideTooltip();
  },
};

document.addEventListener('DOMContentLoaded', () => {
  DiffViewer.initSyncButton();
  DiffViewer.initTooltip();
});
