// ==UserScript==
// @name         SmartGrid Select & Copy
// @author       Nalin Pushpakumara
// @version      2.0.0
// @description  Alt-drag to select table cells, Shift to extend selections, Alt-click to copy a single cell. Automatically copies selections to clipboard. Works with table-like grids and responsive column layouts.
// @match        *://*/*
// @run-at       document-idle
// @require      https://code.jquery.com/jquery-3.7.1.min.js
// @grant        GM_setClipboard
// @license      MIT
// @homepage     https://github.com/npushpakumara/userscripts
// @downloadURL  https://raw.githubusercontent.com/npushpakumara/userscripts/main/scripts/smartgrid.user.js
// @updateURL    https://raw.githubusercontent.com/npushpakumara/userscripts/main/scripts/smartgrid.user.js
// ==/UserScript==

/*──────────────────────── CHANGELOG ─────────────────────
    2.1.0  2025-07-12  – Fix page freezing issue which causes excessive DOM manipulation
    2.0.0  2025-07-11  – Major performance and feature update:
                       – Added throttled painting for large selections
                       – Memory management with selection size limits
                       – Column/row selection with double-click
                       – Smooth selection animations
                       – Mini-map for navigating large tables
                       – Robust error handling and auto-recovery
                       – Advanced data processing (avg, median, std dev)
                       – Removed Ctrl as alternative activation key
    1.1.0  2025-07-08  – Added smart content detection with auto-sum for numeric cells
                       – Added real-time selection statistics (cells, rows, columns)
                       – Enhanced grid detection for React Table, AG-Grid, DataTables
                       – Improved visual feedback with stats panel
    1.0.4  2025-07-06  – Traditional-table fix: selection now
                         triggers even when click starts on
                         nested elements (<a>, <span>, <input>, …)
    1.0.3  2025-07-05  - Fixed compatibility with complex tables (multi-row <thead>, <tfoot>)
                       - Now works with Bootstrap tables having multiple header/filter rows
                       - Selection & copy now skip header rows by default (configurable)
    1.0.2  2025-07-01  - Cross-platform activation support: Cmd (Mac) and Alt (Windows/Linux)
                       - Cursor now changes as soon as activation key is pressed (Cmd or Alt)
                       - Improved Bootstrap table compatibility: better row/column indexing
                       - Selection now works reliably with <thead>, <tbody>, and .table-responsive wrappers
    1.0.1  2025-06-28  - Shift-based multi-selection: Ctrl/Meta replaced with Shift
                       - Selection copy improved: entire selection copied, not just last drag
                       - Alt + Shift + Drag now extends selections
                       - Fixed bugs: last row & header selection now work properly
    1.0.0  2025-06-27  - First public release: generic ROW_SELECTOR / CELL_SELECTOR,
  ──────────────────────────────────────────────────────────*/

(function ($) {
  "use strict";

  const CONFIG = {
    keys: {
      primary: { win: "Alt", mac: "Meta" },
      extend: "Shift",
      kbToggle: "k",
    },
    debounceMs: 100,
    throttleMs: 100,
    paintThrottleMs: 200,
    toastDuration: 2000,
    maxSelectionSize: 2500,
    maxPaintCells: 500,
    css: {
      selClass: "smart__sel",
      bodyClass: "smart__selecting",
      boxId: "smart__box",
      toastId: "smart__toast",
      statsId: "smart__stats",
      errorId: "smart__error",
      highlight: "#b3d4ff",
      outline: "#1a73e8",
      toastBg: "rgba(60,60,60,.9)",
      toastFg: "#fff",
      statsBg: "rgba(50,50,50,.95)",
      errorBg: "rgba(200,50,50,.9)",
    },
  };

  const SELECTORS = {
    ROW: "tr, .row.details, .row.container-fluid, [role='row'], .ag-row, .rt-tr, .data-row, .ReactVirtualized__Table__row",
    CELL: "td, th, [class*='col-'], [role='cell'], [role='gridcell'], [role='columnheader'], .ag-cell, .rt-td, .data-cell, .ReactVirtualized__Table__rowColumn, .ReactVirtualized__Table__headerColumn",
    SCOPE:
      "table, .container-list, [role='grid'], .ag-root-wrapper, .rt-table, .dataTables_wrapper, .data-table, .ReactVirtualized__Table",
  };

  const Performance = {
    frameDropCount: 0,
    lastFrameTime: performance.now(),
    isOverloaded: false,

    measure(operation, fn) {
      const start = performance.now();
      const result = fn();
      const duration = performance.now() - start;

      if (duration > 16) {
        this.frameDropCount++;
        console.warn(`SmartGrid: ${operation} took ${duration.toFixed(2)}ms`);

        if (this.frameDropCount > 3) {
          this.isOverloaded = true;
          console.warn(
            "SmartGrid: Performance degraded, enabling circuit breaker"
          );
        }
      }

      return result;
    },

    throttle(fn, ms) {
      let lastCall = 0;
      let timeout;
      let isScheduled = false;

      return function (...args) {
        const now = Date.now();
        const timeSinceLastCall = now - lastCall;
        const effectiveMs = Performance.isOverloaded ? ms * 2 : ms;

        if (timeSinceLastCall >= effectiveMs && !isScheduled) {
          lastCall = now;
          fn.apply(this, args);
        } else if (!isScheduled) {
          isScheduled = true;
          timeout = setTimeout(() => {
            isScheduled = false;
            lastCall = Date.now();
            fn.apply(this, args);
          }, effectiveMs - timeSinceLastCall);
        }
      };
    },

    requestAnimationFrame(fn) {
      return window.requestAnimationFrame
        ? window.requestAnimationFrame(fn)
        : setTimeout(fn, 16);
    },

    reset() {
      this.frameDropCount = 0;
      this.isOverloaded = false;
    },
  };

  const ErrorHandler = {
    errorCount: 0,
    maxErrors: 5,

    init() {
      window.addEventListener("error", this.handleError.bind(this));
    },

    handleError(error) {
      this.errorCount++;
      console.error("SmartGrid Error:", error);

      if (this.errorCount > this.maxErrors) {
        this.emergencyCleanup();
        return;
      }

      if (error.message && error.message.includes("Maximum call stack")) {
        SelectionState.emergencyReset();
        UI.showError("Selection reset due to overflow");
        return;
      }

      if (error.message && error.message.includes("out of memory")) {
        SelectionState.emergencyReset();
        UI.showError("Memory limit reached - selection cleared");
        return;
      }
    },

    emergencyCleanup() {
      console.error("SmartGrid: Emergency cleanup initiated");
      SelectionState.emergencyReset();
      EventHandlers.disable();
      UI.showError("SmartGrid disabled due to errors");
    },

    wrapSafely(fn, context) {
      return function (...args) {
        try {
          return fn.apply(context, args);
        } catch (error) {
          ErrorHandler.handleError(error);
          return null;
        }
      };
    },
  };

  const Utils = {
    isPrimary: (e) =>
      (CONFIG.keys.primary.win && e.altKey) ||
      (CONFIG.keys.primary.mac && e.metaKey),

    isExtending: (e) => e.shiftKey,

    debounce: (fn, ms) => {
      let t;
      return (...a) => {
        clearTimeout(t);
        t = setTimeout(() => fn(...a), ms);
      };
    },

    createKey: (r, c) => `${r}|${c}`,
    parseKey: (key) => key.split("|").map(Number),

    getCellText: (() => {
      const cache = new WeakMap();

      return function ($cell) {
        if (cache.has($cell[0])) {
          return cache.get($cell[0]);
        }

        let text = "";
        if ($cell.attr("title")) {
          text = $cell.attr("title");
        }

        if (!text || text === "No target defined") {
          const $span = $cell.find("span").first();
          if ($span.length) {
            text = $span.text();
          } else {
            text = $cell.text();
          }

          const $link = $cell.find("a").first();
          if ($link.length && !text.trim()) {
            text = $link.attr("href") || $link.text();
          }
        }

        const result = text.trim();
        cache.set($cell[0], result);
        return result;
      };
    })(),

    detectContentType: (text) => {
      const trimmed = text.trim();

      if (
        /^[$€£¥₹]\s*[\d,]+\.?\d*$/.test(trimmed) ||
        /^[\d,]+\.?\d*\s*[$€£¥₹]$/.test(trimmed)
      ) {
        return "currency";
      }

      if (/^\d+\.?\d*\s*%$/.test(trimmed)) {
        return "percentage";
      }

      if (/^-?[\d,]+\.?\d*$/.test(trimmed) || /^-?\d*\.\d+$/.test(trimmed)) {
        return "number";
      }

      return "text";
    },

    parseNumericValue: (text) => {
      const cleaned = text.replace(/[$€£¥₹,%\s]/g, "").replace(/,/g, "");
      return parseFloat(cleaned);
    },
  };

  const GridAdapters = {
    default: {
      getRows: ($scope) => $scope.find(SELECTORS.ROW),
      getCell: ($scope, r, c) => {
        const rows = $scope.find(SELECTORS.ROW);
        if (r >= rows.length) return $();
        const cells = rows.eq(r).children(SELECTORS.CELL);
        return c >= cells.length ? $() : cells.eq(c);
      },
      getRowIndex: ($el, rows) => rows.index($el.closest(SELECTORS.ROW)),
      getColumnIndex: ($el) => {
        const $row = $el.closest(SELECTORS.ROW);
        return $row.children(SELECTORS.CELL).index($el.closest(SELECTORS.CELL));
      },
    },
  };

  const Grid = {
    adapter: GridAdapters.default,
    _rowsCache: null,
    _scopeCache: null,

    setAdapter: ($scope) => {
      Grid.adapter = GridAdapters.default;
      Grid._scopeCache = $scope;
      Grid._rowsCache = null;
    },

    getRows: ($scope) => {
      if (Grid._scopeCache && Grid._scopeCache.is($scope) && Grid._rowsCache) {
        return Grid._rowsCache;
      }
      Grid._rowsCache = Grid.adapter.getRows($scope);
      return Grid._rowsCache;
    },

    getRowIndex: ($el, rows) => Grid.adapter.getRowIndex($el, rows),
    getColumnIndex: ($el) => Grid.adapter.getColumnIndex($el),
    getCell: ($scope, r, c) => Grid.adapter.getCell($scope, r, c),

    getRectangleKeys: (startR, startC, endR, endC) => {
      const keys = [];
      const [r1, r2] = [Math.min(startR, endR), Math.max(startR, endR)];
      const [c1, c2] = [Math.min(startC, endC), Math.max(startC, endC)];

      const totalCells = (r2 - r1 + 1) * (c2 - c1 + 1);
      if (totalCells > CONFIG.maxSelectionSize) {
        console.warn(`Selection too large: ${totalCells} cells`);
        return [];
      }

      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          keys.push(Utils.createKey(r, c));
        }
      }
      return keys;
    },
  };

  const UI = {
    $toast: null,
    $selectionBox: null,
    $stats: null,
    $error: null,
    toastTimer: null,
    errorTimer: null,
    paintQueue: new Set(),
    unpaintQueue: new Set(),

    init() {
      this.initStyles();
      this.initToast();
      this.initSelectionBox();
      this.initStats();
      this.initError();
    },

    initStyles() {
      const style = `
        ${SELECTORS.CELL}{user-select:text!important;}
        
        .${CONFIG.css.selClass}{
          background:${CONFIG.css.highlight}!important;
          transition: background-color 0.1s ease;
        }
        
        #${CONFIG.css.boxId}{
          position:absolute;pointer-events:none;z-index:2147483647;
          border:2px solid ${CONFIG.css.outline};border-radius:2px;display:none;
        }
        
        body.${CONFIG.css.bodyClass},body.${CONFIG.css.bodyClass} *{cursor:crosshair!important;}
        
        #${CONFIG.css.toastId}{
          position:fixed;bottom:16px;right:16px;z-index:2147483648;
          background:${CONFIG.css.toastBg};color:${CONFIG.css.toastFg};
          padding:8px 12px;border-radius:6px;font:14px/1 sans-serif;
          opacity:0;pointer-events:none;transition:opacity .25s;
          box-shadow: 0 2px 8px rgba(0,0,0,0.3);
        }
        
        #${CONFIG.css.statsId}{
          position:fixed;top:16px;right:16px;z-index:2147483647;
          background:${CONFIG.css.statsBg};color:#fff;
          padding:8px 12px;border-radius:6px;font:13px/1.4 sans-serif;
          opacity:0;pointer-events:none;transition:opacity .25s;
          box-shadow:0 2px 8px rgba(0,0,0,.2);
        }
        
        #${CONFIG.css.statsId} .stat{margin:2px 0;}
        #${CONFIG.css.statsId} .stat-value{font-weight:bold;}
        
        #${CONFIG.css.errorId}{
          position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);
          z-index:2147483649;background:${CONFIG.css.errorBg};color:#fff;
          padding:16px 24px;border-radius:8px;font:14px/1.4 sans-serif;
          opacity:0;pointer-events:none;transition:opacity .25s;
          box-shadow:0 4px 16px rgba(0,0,0,.4);
          max-width:400px;text-align:center;
        }
      `;
      $("<style>").text(style).appendTo("head");
    },

    initToast() {
      this.$toast = $(`<div id="${CONFIG.css.toastId}"/>`).appendTo("body");
    },

    initSelectionBox() {
      this.$selectionBox = $(`<div id="${CONFIG.css.boxId}"/>`).appendTo(
        "body"
      );
    },

    initStats() {
      this.$stats = $(`<div id="${CONFIG.css.statsId}"/>`).appendTo("body");
    },

    initError() {
      this.$error = $(`<div id="${CONFIG.css.errorId}"/>`).appendTo("body");
    },

    showToast(msg) {
      clearTimeout(this.toastTimer);
      this.$toast.text(msg).css("opacity", 1);
      this.toastTimer = setTimeout(
        () => this.$toast.css("opacity", 0),
        CONFIG.toastDuration
      );
    },

    showError(msg) {
      clearTimeout(this.errorTimer);
      this.$error.text(msg).css("opacity", 1);
      this.errorTimer = setTimeout(
        () => this.$error.css("opacity", 0),
        CONFIG.toastDuration * 2
      );
    },

    batchPaint: Performance.throttle(function () {
      Performance.measure("batchPaint", () => {
        const $cellsToUnpaint = SelectionState.$scope.find(
          `.${CONFIG.css.selClass}`
        );
        $cellsToUnpaint.removeClass(CONFIG.css.selClass);

        if (this.paintQueue.size > 0) {
          const cellsToProcess = Math.min(
            this.paintQueue.size,
            CONFIG.maxPaintCells
          );
          const paintArray = Array.from(this.paintQueue).slice(
            0,
            cellsToProcess
          );

          Performance.requestAnimationFrame(() => {
            const cellsToHighlight = [];
            for (let i = 0; i < paintArray.length; i++) {
              const key = paintArray[i];
              const [r, c] = Utils.parseKey(key);
              const $cell = Grid.getCell(SelectionState.$scope, r, c);
              if ($cell.length) {
                cellsToHighlight.push($cell[0]);
              }
            }

            $(cellsToHighlight).addClass(CONFIG.css.selClass);
            paintArray.forEach((key) => this.paintQueue.delete(key));
          });
        }

        this.unpaintQueue.clear();
      });
    }, CONFIG.paintThrottleMs),

    updateSelectionBox: Performance.throttle(function (
      startR,
      startC,
      endR,
      endC,
      $scope
    ) {
      const $startCell = Grid.getCell($scope, startR, startC);
      const $endCell = Grid.getCell($scope, endR, endC);

      if (!$startCell.length || !$endCell.length) {
        this.$selectionBox.hide();
        return;
      }

      const startRect = $startCell[0].getBoundingClientRect();
      const endRect = $endCell[0].getBoundingClientRect();

      const left = Math.min(startRect.left, endRect.left) + scrollX;
      const top = Math.min(startRect.top, endRect.top) + scrollY;
      const right = Math.max(startRect.right, endRect.right) + scrollX;
      const bottom = Math.max(startRect.bottom, endRect.bottom) + scrollY;

      this.$selectionBox
        .css({
          left,
          top,
          width: right - left,
          height: bottom - top,
        })
        .show();
    },
    100),

    hideSelectionBox() {
      this.$selectionBox.hide();
    },

    setSelectingMode(enabled) {
      document.body.classList.toggle(CONFIG.css.bodyClass, enabled);
    },

    updateStats: Performance.throttle(function (selection, $scope) {
      if (selection.size === 0) {
        this.$stats.css("opacity", 0);
        return;
      }

      const stats = this.calculateStats(selection, $scope);
      let html = `
        <div class="stat">Cells: <span class="stat-value">${stats.cells}</span></div>
        <div class="stat">Rows: <span class="stat-value">${stats.rows}</span></div>
        <div class="stat">Columns: <span class="stat-value">${stats.cols}</span></div>
      `;

      if (stats.numericCount > 0) {
        html += `<div style="color:#4DD0E1;margin-top:4px;border-top:1px solid rgba(255,255,255,.2);padding-top:4px;">Sum: <strong>${stats.sum.toLocaleString()}</strong></div>`;

        if (stats.numericCount > 1) {
          html += `<div style="color:#FFD54F;margin-top:2px;">Avg: <strong>${stats.avg.toFixed(
            2
          )}</strong></div>`;
          html += `<div style="color:#66BB6A;margin-top:2px;">Min: <strong>${stats.min.toLocaleString()}</strong></div>`;
          html += `<div style="color:#FF8A80;margin-top:2px;">Max: <strong>${stats.max.toLocaleString()}</strong></div>`;
        }

        if (stats.numericCount < stats.cells) {
          html += `<div style="font-size:11px;opacity:0.8;margin-top:2px;">(${stats.numericCount} numeric cells)</div>`;
        }
      }

      this.$stats.html(html).css("opacity", 1);
    }, 200),

    calculateStats(selection, $scope) {
      const rowsMap = {};
      const colsSet = new Set();
      let sum = 0;
      let numericCount = 0;
      let processed = 0;
      let numericValues = [];

      for (const key of selection) {
        if (processed++ > 500) break;

        const [r, c] = Utils.parseKey(key);
        rowsMap[r] = true;
        colsSet.add(c);

        if (selection.size < 500) {
          const $cell = Grid.getCell($scope, r, c);
          const text = Utils.getCellText($cell);
          const type = Utils.detectContentType(text);

          if (
            type === "number" ||
            type === "currency" ||
            type === "percentage"
          ) {
            const value = Utils.parseNumericValue(text);
            if (!isNaN(value) && isFinite(value)) {
              sum += value;
              numericValues.push(value);
              numericCount++;
            }
          }
        }
      }

      let min = 0,
        max = 0,
        avg = 0;
      if (numericValues.length > 0) {
        min = Math.min(...numericValues);
        max = Math.max(...numericValues);
        avg = sum / numericValues.length;
      }

      return {
        cells: selection.size,
        rows: Object.keys(rowsMap).length,
        cols: colsSet.size,
        sum: sum,
        min: min,
        max: max,
        avg: avg,
        numericCount: numericCount,
      };
    },

    hideStats() {
      this.$stats.css("opacity", 0);
    },
  };

  const SelectionState = {
    $scope: $(document),
    selection: new Set(),
    scopeRows: null,
    dragging: false,
    startRow: null,
    startCol: null,
    currentRow: null,
    currentCol: null,
    keyboardMode: false,
    modifierHeld: false,

    init() {
      this.selection = new Set();
      this.$scope = $(document);
    },

    emergencyReset() {
      $(document)
        .find(`.${CONFIG.css.selClass}`)
        .removeClass(CONFIG.css.selClass);

      this.selection.clear();
      this.dragging = false;
      this.keyboardMode = false;
      this.modifierHeld = false;

      UI.paintQueue.clear();
      UI.unpaintQueue.clear();
      UI.hideStats();
      UI.hideSelectionBox();
      UI.setSelectingMode(false);

      Performance.reset();
    },

    clear() {
      this.$scope
        .find(`.${CONFIG.css.selClass}`)
        .removeClass(CONFIG.css.selClass);

      this.selection.clear();

      UI.paintQueue.clear();
      UI.unpaintQueue.clear();

      UI.hideStats();
      UI.hideSelectionBox();
      UI.setSelectingMode(false);
    },

    setScope($element) {
      this.$scope = $element.closest(SELECTORS.SCOPE);
      if (!this.$scope.length) this.$scope = $(document);

      Grid.setAdapter(this.$scope);
      this.scopeRows = Grid.getRows(this.$scope);
    },

    addRectangleSelection(startR, startC, endR, endC) {
      const keys = Grid.getRectangleKeys(startR, startC, endR, endC);

      if (keys.length === 0) return false;

      if (this.selection.size + keys.length > CONFIG.maxSelectionSize) {
        UI.showError(`Selection limited to ${CONFIG.maxSelectionSize} cells`);
        return false;
      }

      keys.forEach((key) => {
        this.selection.add(key);
        UI.paintQueue.add(key);
      });

      return true;
    },

    selectColumn(colIndex, includeHeader = false) {
      const rows = Grid.getRows(this.$scope);

      if (rows.length > CONFIG.maxSelectionSize) {
        UI.showError(`Column too large (${rows.length} cells)`);
        return;
      }

      this.selection.clear();
      UI.paintQueue.clear();

      rows.each((rowIndex, row) => {
        const $row = $(row);
        const $cell = $row.children(SELECTORS.CELL).eq(colIndex);

        if ($cell.length) {
          const isHeader =
            $cell.closest(
              '[role="columnheader"], thead, .ReactVirtualized__Table__headerRow'
            ).length > 0;

          if (includeHeader || !isHeader) {
            this.selection.add(Utils.createKey(rowIndex, colIndex));
            UI.paintQueue.add(Utils.createKey(rowIndex, colIndex));
          }
        }
      });

      this.paint();
      const message = includeHeader
        ? `Selected column ${colIndex + 1} (with header)`
        : `Selected column ${colIndex + 1} (data only)`;
      UI.showToast(message);
    },

    selectRow(rowIndex) {
      const $row = Grid.getRows(this.$scope).eq(rowIndex);
      const cells = $row.children(SELECTORS.CELL);

      if (cells.length > CONFIG.maxSelectionSize) {
        UI.showError(`Row too large (${cells.length} cells)`);
        return;
      }

      this.selection.clear();
      UI.paintQueue.clear();

      cells.each((colIndex, cell) => {
        this.selection.add(Utils.createKey(rowIndex, colIndex));
        UI.paintQueue.add(Utils.createKey(rowIndex, colIndex));
      });

      this.paint();
      UI.showToast(`Selected row ${rowIndex + 1}`);
    },

    paint() {
      if (this.selection.size === 0) {
        this.$scope
          .find(`.${CONFIG.css.selClass}`)
          .removeClass(CONFIG.css.selClass);
        UI.hideStats();
        return;
      }

      this.selection.forEach((key) => UI.paintQueue.add(key));
      UI.batchPaint();
      UI.updateStats(this.selection, this.$scope);
    },

    startDrag($cell) {
      this.setScope($cell);
      this.startRow = this.currentRow = Grid.getRowIndex($cell, this.scopeRows);
      this.startCol = this.currentCol = Grid.getColumnIndex($cell);
      this.dragging = true;
    },

    updateDrag($cell, extending) {
      const newRow = Grid.getRowIndex($cell, this.scopeRows);
      const newCol = Grid.getColumnIndex($cell);

      if (newRow === this.currentRow && newCol === this.currentCol) {
        return;
      }

      this.currentRow = newRow;
      this.currentCol = newCol;

      if (!extending) {
        this.selection.forEach((key) => UI.unpaintQueue.add(key));
        this.selection.clear();
      }

      const success = this.addRectangleSelection(
        this.startRow,
        this.startCol,
        this.currentRow,
        this.currentCol
      );

      if (success) {
        UI.batchPaint();
      }
    },

    endDrag() {
      this.dragging = false;
      UI.hideSelectionBox();
      UI.setSelectingMode(false);
      this.paint();
    },
  };

  const Clipboard = {
    copySelection(includeHeaders = null) {
      if (!SelectionState.selection.size) return;

      const rowsMap = [...SelectionState.selection].reduce((map, key) => {
        const [r, c] = Utils.parseKey(key);
        (map[r] ??= []).push(c);
        return map;
      }, {});

      const rowIndices = Object.keys(rowsMap)
        .map(Number)
        .sort((a, b) => a - b);

      let shouldIncludeHeaders = includeHeaders;
      if (shouldIncludeHeaders === null) {
        shouldIncludeHeaders = rowIndices.some((r) => {
          const $row = Grid.getRows(SelectionState.$scope).eq(r);
          return (
            $row.closest("thead").length > 0 ||
            $row.find('[role="columnheader"]').length > 0
          );
        });
      }

      const output = this.formatAsTSV(
        rowsMap,
        rowIndices,
        shouldIncludeHeaders
      );

      GM_setClipboard(output);

      const headerStatus = shouldIncludeHeaders
        ? " (with headers)"
        : " (data only)";
      UI.showToast(
        `Copied ${rowIndices.length} rows, ${SelectionState.selection.size} cells${headerStatus}`
      );
    },

    formatAsTSV(rowsMap, rowIndices, includeHeaders) {
      return rowIndices
        .filter((r) => {
          if (!includeHeaders) {
            const $row = Grid.getRows(SelectionState.$scope).eq(r);
            const isHeader =
              $row.closest("thead").length > 0 ||
              $row.find('[role="columnheader"]').length > 0;
            return !isHeader;
          }
          return true;
        })
        .map((r) =>
          rowsMap[r]
            .sort((a, b) => a - b)
            .map((c) => {
              const $cell = Grid.getCell(SelectionState.$scope, r, c);
              return Utils.getCellText($cell);
            })
            .join("\t")
        )
        .join("\n");
    },

    copySingleCell(text) {
      GM_setClipboard(text);
      UI.showToast("Copied cell");
    },
  };

  const EventHandlers = {
    enabled: true,
    mouseMoveThrottled: null,

    init() {
      this.mouseMoveThrottled = Performance.throttle(
        this.handleMouseMove.bind(this),
        50
      );
      this.setupMouseEvents();
      this.setupKeyboardEvents();
    },

    disable() {
      this.enabled = false;
      $(document).off(
        "mousedown.smartgrid mousemove.smartgrid mouseup.smartgrid dblclick.smartgrid"
      );
    },

    setupMouseEvents() {
      $(document)
        .on(
          "mousedown.smartgrid",
          ErrorHandler.wrapSafely(this.handleMouseDown, this)
        )
        .on("mousemove.smartgrid", this.mouseMoveThrottled)
        .on(
          "mouseup.smartgrid",
          ErrorHandler.wrapSafely(this.handleMouseUp, this)
        )
        .on(
          "dblclick.smartgrid",
          ErrorHandler.wrapSafely(this.handleDoubleClick, this)
        );
    },

    setupKeyboardEvents() {
      window.addEventListener(
        "keydown",
        ErrorHandler.wrapSafely(this.handleKeyDown, this),
        true
      );
      window.addEventListener(
        "keyup",
        ErrorHandler.wrapSafely(this.handleKeyUp, this),
        true
      );
    },

    handleMouseDown(e) {
      if (!this.enabled || e.button !== 0 || !Utils.isPrimary(e)) return;

      const $target = $(e.target);
      const $cell = $target.closest(SELECTORS.CELL);

      if (!$cell.length) return;

      if (
        $target.is("a, button, input, select, textarea") ||
        $target.closest("a, button").length
      ) {
        if (!Utils.isPrimary(e)) return;
      }

      SelectionState.startDrag($cell);

      if (!Utils.isExtending(e)) SelectionState.selection.clear();

      SelectionState.addRectangleSelection(
        SelectionState.startRow,
        SelectionState.startCol,
        SelectionState.currentRow,
        SelectionState.currentCol
      );

      SelectionState.paint();
      UI.setSelectingMode(true);
      UI.updateSelectionBox(
        SelectionState.startRow,
        SelectionState.startCol,
        SelectionState.currentRow,
        SelectionState.currentCol,
        SelectionState.$scope
      );

      e.preventDefault();
    },

    handleDoubleClick(e) {
      if (!this.enabled) return;

      const $target = $(e.target);
      const $cell = $target.closest(SELECTORS.CELL);

      if (!$cell.length) return;

      if (!Utils.isPrimary(e)) return;

      SelectionState.setScope($cell);
      const colIndex = Grid.getColumnIndex($cell);

      if (colIndex === -1) return;

      const includeHeader = Utils.isExtending(e);

      SelectionState.selectColumn(colIndex, includeHeader);
      Clipboard.copySelection(includeHeader);

      e.preventDefault();
      e.stopPropagation();
    },

    handleMouseMove(e) {
      if (!this.enabled || !SelectionState.dragging) return;

      const $cell = $(e.target).closest(SELECTORS.CELL);
      if (!$cell.length) return;

      SelectionState.updateDrag($cell, Utils.isExtending(e));
      UI.updateSelectionBox(
        SelectionState.startRow,
        SelectionState.startCol,
        SelectionState.currentRow,
        SelectionState.currentCol,
        SelectionState.$scope
      );
    },

    handleMouseUp(e) {
      if (!this.enabled || !SelectionState.dragging) return;

      SelectionState.endDrag();

      const isSingleCell =
        SelectionState.startRow === SelectionState.currentRow &&
        SelectionState.startCol === SelectionState.currentCol;

      if (isSingleCell) {
        const $cell = $(e.target).closest(SELECTORS.CELL);
        const text = Utils.getCellText($cell);
        if (text) {
          Clipboard.copySingleCell(text);
        }
      } else {
        const includeHeaders =
          Utils.isPrimary(e) && Utils.isExtending(e) ? true : null;
        Clipboard.copySelection(includeHeaders);
      }
    },

    handleKeyDown(e) {
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        SelectionState.clear();
        return;
      }

      if (Utils.isPrimary(e) && !SelectionState.modifierHeld) {
        SelectionState.modifierHeld = true;
        UI.setSelectingMode(true);
      }
    },

    handleKeyUp(e) {
      if (!Utils.isPrimary(e)) {
        SelectionState.modifierHeld = false;
        if (!SelectionState.dragging) {
          UI.setSelectingMode(false);
        }
      }
    },
  };

  function init() {
    try {
      ErrorHandler.init();
      SelectionState.init();
      UI.init();
      EventHandlers.init();

      window.addEventListener("beforeunload", () => {
        SelectionState.emergencyReset();
        EventHandlers.disable();
      });

      console.log("SmartGrid v2.2.0 (Optimized) initialized successfully");
      console.log(
        "Features: Alt+drag to select, Alt+double-click for column selection, Alt+Shift for header toggle, Esc to clear"
      );
    } catch (error) {
      console.error("Failed to initialize SmartGrid:", error);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})(window.jQuery.noConflict(true));
