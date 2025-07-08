// ==UserScript==
// @name         SmartGrid Select & Copy
// @author       Nalin Pushpakumara
// @version      1.0.2
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

// >>  Add a new dated bullet every time you increment @version.  <<
// >>  Keep older entries or truncate to last N versions—your call. <<

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
      highlight: "#b3d4ff",
      outline: "#1a73e8",
      toastBg: "rgba(60,60,60,.9)",
      toastFg: "#fff",
    },
  };

  const isPrimary = (e) =>
    (CONFIG.keys.primary.win && e.altKey) ||
    (CONFIG.keys.primary.mac && e.metaKey);
  const isExtending = (e) => e.shiftKey;

  const ROW_SEL = "tr, .row.details, .row.container-fluid";
  const CELL_SEL = "td, th, [class*='col-']";

  const style = `
    ${CELL_SEL}{user-select:text!important;}
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
  `;
  $("<style>").text(style).appendTo("head");

  const $toast = $(`<div id="${CONFIG.css.toastId}"/>`).appendTo("body");
  let toastTimer;
  function showToast(msg) {
    clearTimeout(toastTimer);
    $toast.text(msg).css("opacity", 1);
    toastTimer = setTimeout(
      () => $toast.css("opacity", 0),
      CONFIG.toastDuration
    );
  }

  const unlock = ($r) =>
    $r.find(CELL_SEL).each(function () {
      this.onselectstart = this.onmousedown = null;
    });

  const debounce = (fn, ms) => {
    let t;
    return (...a) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...a), ms);
    };
  };
  const unlockBatch = debounce(
    (nodes) => nodes.forEach((n) => unlock($(n))),
    CONFIG.debounceMs
  );

  unlock($(document));
  const mo = new MutationObserver((m) =>
    unlockBatch(m.flatMap((r) => [...r.addedNodes]))
  );
  mo.observe(document.body, { childList: true, subtree: true });

  document.addEventListener("visibilitychange", () => {
    document.hidden
      ? mo.disconnect()
      : mo.observe(document.body, { childList: true, subtree: true });
  });

  const getRows = ($t) => $t.find(ROW_SEL);
  const rowIdx = ($el, rows) => rows.index($el.closest(ROW_SEL));
  const colIdx = ($el) =>
    $el.closest(ROW_SEL).children(CELL_SEL).index($el.closest(CELL_SEL));
  const getCell = ($scope, r, c) =>
    $scope.find(ROW_SEL).eq(r).children(CELL_SEL).eq(c);
  const key = (r, c) => `${r}|${c}`;

  let $scope = $(document);
  const sel = new Set();
  const $box = $(`<div id="${CONFIG.css.boxId}"/>`).appendTo("body");
  let dragging = false,
    startR,
    startC,
    curR,
    curC,
    scopeRows,
    modifierHeld = false,
    kbMode = false;

  const paint = () => {
    $scope.find(`.${CONFIG.css.selClass}`).removeClass(CONFIG.css.selClass);
    sel.forEach((k) => {
      const [r, c] = k.split("|").map(Number);
      getCell($scope, r, c).addClass(CONFIG.css.selClass);
    });
  };

  const rectKeys = () => {
    const ks = [];
    const [r1, r2] = [startR, curR].sort((a, b) => a - b);
    const [c1, c2] = [startC, curC].sort((a, b) => a - b);
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) ks.push(key(r, c));
    return ks;
  };

  const updateBox = () => {
    const $a = getCell($scope, startR, startC);
    const $b = getCell($scope, curR, curC);
    if (!$a.length || !$b.length) {
      $box.hide();
      return;
    }
    const ra = $a[0].getBoundingClientRect(),
      rb = $b[0].getBoundingClientRect();
    const left = Math.min(ra.left, rb.left) + scrollX;
    const top = Math.min(ra.top, rb.top) + scrollY;
    const right = Math.max(ra.right, rb.right) + scrollX;
    const bot = Math.max(ra.bottom, rb.bottom) + scrollY;
    $box.css({ left, top, width: right - left, height: bot - top });
  };

  function copySelection(format = "tsv") {
    if (!sel.size) return;
    const rowsMap = [...sel].reduce((m, k) => {
      const [r, c] = k.split("|");
      (m[r] ??= []).push(+c);
      return m;
    }, {});
    const rowIdxs = Object.keys(rowsMap)
      .map(Number)
      .sort((a, b) => a - b);

    let out;
    if (format === "markdown") {
      out = rowIdxs
        .map((r, i) => {
          const line =
            "| " +
            rowsMap[r]
              .sort((a, b) => a - b)
              .map((c) => getCell($scope, r, c).text().trim())
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
    } else {
      out = rowIdxs
        .map((r) =>
          rowsMap[r]
            .sort((a, b) => a - b)
            .map((c) => getCell($scope, r, c).text().trim())
            .join("\t")
        )
        .join("\n");
    }
    GM_setClipboard(out);
    const rCnt = rowIdxs.length,
      cCnt = Math.max(...Object.values(rowsMap).map((a) => a.length));
    showToast(
      format === "markdown"
        ? `Copied ${rCnt}×${cCnt} as Markdown`
        : `Copied ${rCnt}×${cCnt} cells`
    );
  }

  $(document)
    .on("mousedown", (e) => {
      if (e.button !== 0 || !isPrimary(e)) return;
      const $cell = $(e.target).closest(CELL_SEL);
      if (!$cell.length) return;

      $scope = $cell.closest("table, .container-list");
      if (!$scope.length) $scope = $(document);
      scopeRows = getRows($scope);
      startR = curR = rowIdx($cell, scopeRows);
      startC = curC = colIdx($cell);

      if (!isExtending(e)) sel.clear();
      rectKeys().forEach((k) => sel.add(k));

      dragging = true;
      paint();
      $box.show();
      document.body.classList.add(CONFIG.css.bodyClass);
      updateBox();
      e.preventDefault();
    })
    .on("mousemove", (e) => {
      if (!dragging) return;
      const $cell = $(e.target).closest(CELL_SEL);
      if (!$cell.length) return;
      curR = rowIdx($cell, scopeRows);
      curC = colIdx($cell);
      if (!isExtending(e)) sel.clear();
      rectKeys().forEach((k) => sel.add(k));
      paint();
      updateBox();
      $cell[0].scrollIntoView({ block: "nearest", inline: "nearest" }); // edge scroll
    })
    .on("mouseup", (e) => {
      if (!dragging) return;
      dragging = false;
      $box.hide();
      document.body.classList.remove(CONFIG.css.bodyClass);
      paint();

      const single = startR === curR && startC === curC;
      const txt = $(e.target).closest(CELL_SEL).text().trim();
      if (single && txt) {
        GM_setClipboard(txt);
        showToast("Copied cell");
      } else {
        copySelection();
      }
    });

  window.addEventListener(
    "keydown",
    (e) => {
      if (e.key === "Escape") {
        if (kbMode) {
          kbMode = false;
          paint();
          return;
        }
        sel.clear();
        paint();
        return;
      }

      if (isPrimary(e) && e.key.toLowerCase() === CONFIG.keys.kbToggle) {
        e.preventDefault();
        kbMode = !kbMode;
        if (kbMode) {
          const $focused = $(document.activeElement).closest(CELL_SEL);
          const $cell = $focused.length
            ? $focused
            : $(CELL_SEL).filter(":visible").first();
          if (!$cell.length) {
            kbMode = false;
            return;
          }
          $scope = $cell.closest("table, .container-list");
          scopeRows = getRows($scope);
          startR = curR = rowIdx($cell, scopeRows);
          startC = curC = colIdx($cell);
          sel.clear();
          sel.add(key(startR, startC));
          paint();
        } else {
          sel.clear();
          paint();
        }
        return;
      }

      if (isPrimary(e) && e.key.toLowerCase() === "m") {
        e.preventDefault();
        copySelection("markdown");
        return;
      }

      if (kbMode) {
        const AR = { ArrowUp: -1, ArrowDown: 1, ArrowLeft: -1, ArrowRight: 1 };
        if (e.key in AR) {
          e.preventDefault();
          if (e.key === "ArrowUp" || e.key === "ArrowDown") {
            curR = Math.max(
              0,
              Math.min(scopeRows.length - 1, curR + AR[e.key])
            );
          } else {
            curC = Math.max(0, curC + AR[e.key]);
          }
          sel.clear();
          rectKeys().forEach((k) => sel.add(k));
          paint();
          getCell($scope, curR, curC)[0].scrollIntoView({ block: "nearest" });
          return;
        }
        if (e.key === " " || e.key === "Enter") {
          e.preventDefault();
          kbMode = false;
          copySelection();
          paint();
          return;
        }
      }

      if (isPrimary(e) && !modifierHeld) {
        modifierHeld = true;
        document.body.classList.add(CONFIG.css.bodyClass);
      }
    },
    true
  );

  window.addEventListener(
    "keyup",
    (e) => {
      if (!isPrimary(e)) {
        modifierHeld = false;
        if (!dragging) document.body.classList.remove(CONFIG.css.bodyClass);
      }
    },
    true
  );
})(window.jQuery.noConflict(true));
