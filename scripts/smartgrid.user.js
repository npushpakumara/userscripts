// ==UserScript==
// @name         SmartGrid Select & Copy
// @author       Nalin Pushpakumara
// @version      1.1.0
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
    debounceMs: 250,
    toastDuration: 2000,
    css: {
      selClass: "smart__sel",
      bodyClass: "smart__selecting",
      boxId: "smart__box",
      toastId: "smart__toast",
      statsId: "smart__stats",
      highlight: "#b3d4ff",
      outline: "#1a73e8",
      toastBg: "rgba(60,60,60,.9)",
      toastFg: "#fff",
      statsBg: "rgba(50,50,50,.9)",
    },
  };

  const SELECTORS = {
    // Enhanced selectors for various grid types including ReactVirtualized
    ROW: "tr, .row.details, .row.container-fluid, [role='row'], .ag-row, .rt-tr, .data-row, .ReactVirtualized__Table__row",
    CELL: "td, th, [class*='col-'], [role='cell'], [role='gridcell'], [role='columnheader'], .ag-cell, .rt-td, .data-cell, .ReactVirtualized__Table__rowColumn, .ReactVirtualized__Table__headerColumn",
    SCOPE:
      "table, .container-list, [role='grid'], .ag-root-wrapper, .rt-table, .dataTables_wrapper, .data-table, .ReactVirtualized__Table",
  };

  const GridAdapters = {
    default: {
      getRows: ($scope) => $scope.find(SELECTORS.ROW),
      getCell: ($scope, r, c) =>
        $scope.find(SELECTORS.ROW).eq(r).children(SELECTORS.CELL).eq(c),
      getRowIndex: ($el, rows) => rows.index($el.closest(SELECTORS.ROW)),
      getColumnIndex: ($el) =>
        $el
          .closest(SELECTORS.ROW)
          .children(SELECTORS.CELL)
          .index($el.closest(SELECTORS.CELL)),
    },

    agGrid: {
      getRows: ($scope) => $scope.find(".ag-row"),
      getCell: ($scope, r, c) =>
        $scope.find(".ag-row").eq(r).find(".ag-cell").eq(c),
      getRowIndex: ($el, rows) => rows.index($el.closest(".ag-row")),
      getColumnIndex: ($el) =>
        $el.closest(".ag-row").find(".ag-cell").index($el.closest(".ag-cell")),
    },

    reactTable: {
      getRows: ($scope) => $scope.find(".rt-tr"),
      getCell: ($scope, r, c) =>
        $scope.find(".rt-tr").eq(r).find(".rt-td").eq(c),
      getRowIndex: ($el, rows) => rows.index($el.closest(".rt-tr")),
      getColumnIndex: ($el) =>
        $el.closest(".rt-tr").find(".rt-td").index($el.closest(".rt-td")),
    },

    reactVirtualized: {
      getRows: ($scope) => {
        const headerRow = $scope.find(".ReactVirtualized__Table__headerRow");
        const dataRows = $scope.find(".ReactVirtualized__Table__row");
        return headerRow.add(dataRows);
      },
      getCell: ($scope, r, c) => {
        const rows = GridAdapters.reactVirtualized.getRows($scope);
        const $row = rows.eq(r);

        // Check if it's the header row
        if ($row.hasClass("ReactVirtualized__Table__headerRow")) {
          return $row.find(".ReactVirtualized__Table__headerColumn").eq(c);
        } else {
          return $row.find(".ReactVirtualized__Table__rowColumn").eq(c);
        }
      },
      getRowIndex: ($el, rows) => {
        const $row = $el.closest(
          ".ReactVirtualized__Table__row, .ReactVirtualized__Table__headerRow"
        );
        return rows.index($row);
      },
      getColumnIndex: ($el) => {
        const $row = $el.closest(
          ".ReactVirtualized__Table__row, .ReactVirtualized__Table__headerRow"
        );
        const $cell = $el.closest(
          ".ReactVirtualized__Table__rowColumn, .ReactVirtualized__Table__headerColumn"
        );

        if ($row.hasClass("ReactVirtualized__Table__headerRow")) {
          return $row
            .find(".ReactVirtualized__Table__headerColumn")
            .index($cell);
        } else {
          return $row.find(".ReactVirtualized__Table__rowColumn").index($cell);
        }
      },
    },
  };

  const Utils = {
    isPrimary: (e) => {
      const altPressed = e.altKey;

      const primaryPressed =
        (CONFIG.keys.primary.win && altPressed) ||
        (CONFIG.keys.primary.mac && e.metaKey);

      return primaryPressed;
    },

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

      if (/^-?[\d,]+\.?\d*$/.test(trimmed)) {
        return "number";
      }

      if (
        /^\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}$/.test(trimmed) ||
        /^\d{4}-\d{2}-\d{2}$/.test(trimmed)
      ) {
        return "date";
      }

      if (/^https?:\/\//.test(trimmed) || /^www\./.test(trimmed)) {
        return "url";
      }

      if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) {
        return "email";
      }

      return "text";
    },

    parseNumericValue: (text) => {
      const cleaned = text.replace(/[$€£¥₹,%\s]/g, "").replace(/,/g, "");
      return parseFloat(cleaned);
    },

    getCellText: ($cell) => {
      let text = "";

      if ($cell.attr("title")) {
        text = $cell.attr("title");
      } else {
        text = $cell.text();

        const $link = $cell.find("a").first();
        if ($link.length && !text.trim()) {
          text = $link.attr("href") || $link.text();
        }
      }

      return text.trim();
    },
  };

  const Grid = {
    adapter: GridAdapters.default,

    detectGridType: ($scope) => {
      if (
        $scope.find(".ReactVirtualized__Table").length ||
        $scope.hasClass("ReactVirtualized__Table")
      ) {
        console.log("SmartGrid: Detected ReactVirtualized table");
        return GridAdapters.reactVirtualized;
      }
      if (
        $scope.find(".ag-root-wrapper").length ||
        $scope.closest(".ag-root-wrapper").length
      ) {
        console.log("SmartGrid: Detected AG-Grid");
        return GridAdapters.agGrid;
      }
      if (
        $scope.find(".rt-table").length ||
        $scope.closest(".rt-table").length
      ) {
        console.log("SmartGrid: Detected React Table");
        return GridAdapters.reactTable;
      }
      if (
        $scope.find(".dataTables_wrapper").length ||
        $scope.closest(".dataTables_wrapper").length
      ) {
        console.log("SmartGrid: Detected DataTables");
        return GridAdapters.default;
      }
      console.log("SmartGrid: Using default table adapter");
      return GridAdapters.default;
    },

    setAdapter: ($scope) => {
      Grid.adapter = Grid.detectGridType($scope);
    },

    getRows: ($scope) => Grid.adapter.getRows($scope),

    getRowIndex: ($el, rows) => Grid.adapter.getRowIndex($el, rows),

    getColumnIndex: ($el) => Grid.adapter.getColumnIndex($el),

    getCell: ($scope, r, c) => Grid.adapter.getCell($scope, r, c),

    getRectangleKeys: (startR, startC, endR, endC) => {
      const keys = [];
      const [r1, r2] = [startR, endR].sort((a, b) => a - b);
      const [c1, c2] = [startC, endC].sort((a, b) => a - b);
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
    toastTimer: null,

    init() {
      this.initStyles();
      this.initToast();
      this.initSelectionBox();
      this.initStats();
    },

    initStyles() {
      const style = `
        ${SELECTORS.CELL}{user-select:text!important;}
        .${CONFIG.css.selClass}{background:${CONFIG.css.highlight}!important;}
        #${CONFIG.css.boxId}{
          position:absolute;pointer-events:none;z-index:2147483647;
          border:2px solid ${CONFIG.css.outline};border-radius:2px;display:none;}
        body.${CONFIG.css.bodyClass},body.${CONFIG.css.bodyClass} *{cursor:crosshair!important;}
        #${CONFIG.css.toastId}{
          position:fixed;bottom:16px;right:16px;z-index:2147483648;
          background:${CONFIG.css.toastBg};color:${CONFIG.css.toastFg};
          padding:8px 12px;border-radius:6px;font:14px/1 sans-serif;
          opacity:0;pointer-events:none;transition:opacity .25s;}
        #${CONFIG.css.statsId}{
          position:fixed;top:16px;right:16px;z-index:2147483647;
          background:${CONFIG.css.statsBg};color:#fff;
          padding:8px 12px;border-radius:6px;font:13px/1.4 sans-serif;
          opacity:0;pointer-events:none;transition:opacity .25s;
          box-shadow:0 2px 8px rgba(0,0,0,.2);}
        #${CONFIG.css.statsId} .stat{margin:2px 0;}
        #${CONFIG.css.statsId} .sum{color:#4CAF50;font-weight:bold;margin-top:4px;
          padding-top:4px;border-top:1px solid rgba(255,255,255,.2);}
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

    showToast(msg) {
      clearTimeout(this.toastTimer);
      this.$toast.text(msg).css("opacity", 1);
      this.toastTimer = setTimeout(
        () => this.$toast.css("opacity", 0),
        CONFIG.toastDuration
      );
    },

    updateSelectionBox(startR, startC, endR, endC, $scope) {
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

    hideSelectionBox() {
      this.$selectionBox.hide();
    },

    setSelectingMode(enabled) {
      document.body.classList.toggle(CONFIG.css.bodyClass, enabled);
    },

    updateStats(selection, $scope) {
      if (selection.size === 0) {
        this.$stats.css("opacity", 0);
        return;
      }

      const stats = this.calculateStats(selection, $scope);
      let html = `
        <div class="stat">Cells: <strong>${stats.cells}</strong></div>
        <div class="stat">Rows: <strong>${stats.rows}</strong></div>
        <div class="stat">Columns: <strong>${stats.cols}</strong></div>
      `;

      if (stats.numericCount > 0) {
        html += `<div class="sum">Sum: ${stats.sum.toLocaleString()}</div>`;
        if (stats.numericCount < stats.cells) {
          html += `<div class="stat" style="font-size:11px">(${stats.numericCount} numeric cells)</div>`;
        }
      }

      this.$stats.html(html).css("opacity", 1);
    },

    calculateStats(selection, $scope) {
      const rowsMap = [...selection].reduce((map, key) => {
        const [r, c] = Utils.parseKey(key);
        (map[r] ??= new Set()).add(c);
        return map;
      }, {});

      const allCols = new Set();
      let sum = 0;
      let numericCount = 0;

      selection.forEach((key) => {
        const [r, c] = Utils.parseKey(key);
        allCols.add(c);

        const $cell = Grid.getCell($scope, r, c);
        const text = Utils.getCellText($cell);
        const type = Utils.detectContentType(text);

        if (type === "number" || type === "currency" || type === "percentage") {
          const value = Utils.parseNumericValue(text);
          if (!isNaN(value)) {
            sum += value;
            numericCount++;
          }
        }
      });

      return {
        cells: selection.size,
        rows: Object.keys(rowsMap).length,
        cols: allCols.size,
        sum: sum,
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

    clear() {
      this.selection.clear();
      this.paint();
      UI.hideStats();
    },

    setScope($element) {
      this.$scope = $element.closest(SELECTORS.SCOPE);
      if (!this.$scope.length) this.$scope = $(document);

      Grid.setAdapter(this.$scope);

      this.scopeRows = Grid.getRows(this.$scope);
    },

    addRectangleSelection(startR, startC, endR, endC) {
      const keys = Grid.getRectangleKeys(startR, startC, endR, endC);
      keys.forEach((key) => this.selection.add(key));
    },

    paint() {
      this.$scope
        .find(`.${CONFIG.css.selClass}`)
        .removeClass(CONFIG.css.selClass);
      this.selection.forEach((key) => {
        const [r, c] = Utils.parseKey(key);
        Grid.getCell(this.$scope, r, c).addClass(CONFIG.css.selClass);
      });

      UI.updateStats(this.selection, this.$scope);
    },

    startDrag($cell) {
      this.setScope($cell);
      this.startRow = this.currentRow = Grid.getRowIndex($cell, this.scopeRows);
      this.startCol = this.currentCol = Grid.getColumnIndex($cell);
      this.dragging = true;
    },

    updateDrag($cell, extending) {
      this.currentRow = Grid.getRowIndex($cell, this.scopeRows);
      this.currentCol = Grid.getColumnIndex($cell);

      if (!extending) this.selection.clear();
      this.addRectangleSelection(
        this.startRow,
        this.startCol,
        this.currentRow,
        this.currentCol
      );
      this.paint();
    },

    endDrag() {
      this.dragging = false;
      UI.hideSelectionBox();
      UI.setSelectingMode(false);
      this.paint();
    },
  };

  const Clipboard = {
    copySelection(format = "tsv") {
      if (!SelectionState.selection.size) return;

      const rowsMap = [...SelectionState.selection].reduce((map, key) => {
        const [r, c] = Utils.parseKey(key);
        (map[r] ??= []).push(c);
        return map;
      }, {});

      const rowIndices = Object.keys(rowsMap)
        .map(Number)
        .sort((a, b) => a - b);

      const output =
        format === "markdown"
          ? this.formatAsMarkdown(rowsMap, rowIndices)
          : this.formatAsTSV(rowsMap, rowIndices);

      GM_setClipboard(output);
      this.showCopyToast(rowIndices.length, rowsMap, format);
    },

    formatAsMarkdown(rowsMap, rowIndices) {
      return rowIndices
        .map((r, i) => {
          const line =
            "| " +
            rowsMap[r]
              .sort((a, b) => a - b)
              .map((c) => {
                const $cell = Grid.getCell(SelectionState.$scope, r, c);
                return Utils.getCellText($cell);
              })
              .join(" | ") +
            " |";

          if (i === 0) {
            return (
              line + "\n| " + rowsMap[r].map(() => "---").join(" | ") + " |"
            );
          }
          return line;
        })
        .join("\n");
    },

    formatAsTSV(rowsMap, rowIndices) {
      return rowIndices
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

    showCopyToast(rowCount, rowsMap, format) {
      const colCount = Math.max(
        ...Object.values(rowsMap).map((arr) => arr.length)
      );
      const stats = UI.calculateStats(
        SelectionState.selection,
        SelectionState.$scope
      );

      let message =
        format === "markdown"
          ? `Copied ${rowCount}×${colCount} as Markdown`
          : `Copied ${rowCount}×${colCount} cells`;

      if (stats.numericCount > 0) {
        message += ` (Sum: ${stats.sum.toLocaleString()})`;
      }

      UI.showToast(message);
    },

    copySingleCell(text) {
      GM_setClipboard(text);
      const type = Utils.detectContentType(text);
      const message =
        type !== "text" ? `Copied ${type}: ${text}` : "Copied cell";
      UI.showToast(message);
    },
  };

  const TextUnlock = {
    init() {
      this.unlockDocument();
      this.setupMutationObserver();
    },

    unlockDocument() {
      this.unlock($(document));
    },

    unlock($element) {
      $element.find(SELECTORS.CELL).each(function () {
        this.onselectstart = this.onmousedown = null;
      });
    },

    setupMutationObserver() {
      const unlockBatch = Utils.debounce(
        (nodes) => nodes.forEach((node) => this.unlock($(node))),
        CONFIG.debounceMs
      );

      const observer = new MutationObserver((mutations) =>
        unlockBatch(mutations.flatMap((record) => [...record.addedNodes]))
      );

      observer.observe(document.body, { childList: true, subtree: true });

      document.addEventListener("visibilitychange", () => {
        document.hidden
          ? observer.disconnect()
          : observer.observe(document.body, { childList: true, subtree: true });
      });
    },
  };

  const EventHandlers = {
    init() {
      this.setupMouseEvents();
      this.setupKeyboardEvents();
    },

    setupMouseEvents() {
      $(document)
        .on("mousedown", this.handleMouseDown.bind(this))
        .on("mousemove", this.handleMouseMove.bind(this))
        .on("mouseup", this.handleMouseUp.bind(this));
    },

    setupKeyboardEvents() {
      window.addEventListener("keydown", this.handleKeyDown.bind(this), true);
      window.addEventListener("keyup", this.handleKeyUp.bind(this), true);
    },

    handleMouseDown(e) {
      if (e.button !== 0 || !Utils.isPrimary(e)) return;

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

    handleMouseMove(e) {
      if (!SelectionState.dragging) return;

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

      $cell[0].scrollIntoView({ block: "nearest", inline: "nearest" });
    },

    handleMouseUp(e) {
      if (!SelectionState.dragging) return;

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
        Clipboard.copySelection();
      }
    },

    handleKeyDown(e) {
      if (e.key === "Escape") {
        this.handleEscapeKey();
        return;
      }

      if (Utils.isPrimary(e) && e.key.toLowerCase() === CONFIG.keys.kbToggle) {
        this.handleKeyboardToggle(e);
        return;
      }

      if (Utils.isPrimary(e) && e.key.toLowerCase() === "m") {
        e.preventDefault();
        Clipboard.copySelection("markdown");
        return;
      }

      if (SelectionState.keyboardMode) {
        this.handleKeyboardNavigation(e);
        return;
      }

      if (Utils.isPrimary(e) && !SelectionState.modifierHeld) {
        SelectionState.modifierHeld = true;
        UI.setSelectingMode(true);
        UI.showToast("Selection mode active (Alt + drag)");
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

    handleEscapeKey() {
      if (SelectionState.keyboardMode) {
        SelectionState.keyboardMode = false;
        SelectionState.paint();
        return;
      }
      SelectionState.clear();
    },

    handleKeyboardToggle(e) {
      e.preventDefault();
      SelectionState.keyboardMode = !SelectionState.keyboardMode;

      if (SelectionState.keyboardMode) {
        this.initKeyboardMode();
      } else {
        SelectionState.clear();
      }
    },

    initKeyboardMode() {
      const $focused = $(document.activeElement).closest(SELECTORS.CELL);
      const $cell = $focused.length
        ? $focused
        : $(SELECTORS.CELL).filter(":visible").first();

      if (!$cell.length) {
        SelectionState.keyboardMode = false;
        return;
      }

      SelectionState.setScope($cell);
      SelectionState.startRow = SelectionState.currentRow = Grid.getRowIndex(
        $cell,
        SelectionState.scopeRows
      );
      SelectionState.startCol = SelectionState.currentCol =
        Grid.getColumnIndex($cell);
      SelectionState.selection.clear();
      SelectionState.selection.add(
        Utils.createKey(SelectionState.startRow, SelectionState.startCol)
      );
      SelectionState.paint();
    },

    handleKeyboardNavigation(e) {
      const arrows = {
        ArrowUp: -1,
        ArrowDown: 1,
        ArrowLeft: -1,
        ArrowRight: 1,
      };

      if (e.key in arrows) {
        e.preventDefault();

        if (e.key === "ArrowUp" || e.key === "ArrowDown") {
          SelectionState.currentRow = Math.max(
            0,
            Math.min(
              SelectionState.scopeRows.length - 1,
              SelectionState.currentRow + arrows[e.key]
            )
          );
        } else {
          SelectionState.currentCol = Math.max(
            0,
            SelectionState.currentCol + arrows[e.key]
          );
        }

        SelectionState.selection.clear();
        SelectionState.addRectangleSelection(
          SelectionState.startRow,
          SelectionState.startCol,
          SelectionState.currentRow,
          SelectionState.currentCol
        );
        SelectionState.paint();

        Grid.getCell(
          SelectionState.$scope,
          SelectionState.currentRow,
          SelectionState.currentCol
        )[0].scrollIntoView({ block: "nearest" });
        return;
      }

      if (e.key === " " || e.key === "Enter") {
        e.preventDefault();
        SelectionState.keyboardMode = false;
        Clipboard.copySelection();
        SelectionState.paint();
      }
    },
  };

  function init() {
    SelectionState.init();
    UI.init();
    TextUnlock.init();
    EventHandlers.init();
  }
  init();
})(window.jQuery.noConflict(true));
