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

// ───── CHANGE LOG ──────────────────────────────────────────────────────────
// 1.0.2  2025-07-01  - Cross-platform activation support: Cmd (Mac) and Alt (Windows/Linux)
//                    - Cursor now changes as soon as activation key is pressed (Cmd or Alt)
//                    - Improved Bootstrap table compatibility: better row/column indexing
//                    - Selection now works reliably with <thead>, <tbody>, and .table-responsive wrappers
// 1.0.1  2025-06-28  - Shift-based multi-selection: Ctrl/Meta replaced with Shift
//                    - Selection copy improved: entire selection copied, not just last drag
//                    - Alt + Shift + Drag now extends selections
//                    - Fixed bugs: last row & header selection now work properly
// 1.0.0  2025-06-27  - First public release: generic ROW_SELECTOR / CELL_SELECTOR,
//                    - Initial table implementation with single-block selection.
// ───────────────────────────────────────────────────────────────────────────
//
// >>  Add a new dated bullet every time you increment @version.  <<
// >>  Keep older entries or truncate to last N versions—your call. <<

(function ($) {
  "use strict";

  const USE_ALT = true;
  const USE_META = true;

  const MODIFIER_KEY = {
    primary: (e) => (USE_ALT && e.altKey) || (USE_META && e.metaKey),
    shift: (e) => e.shiftKey,
  };

  const ROW_SELECTOR = "thead tr, tbody tr, tfoot tr, .row";
  const CELL_SELECTOR = "td, th, .row";
  const css = `
    ${CELL_SELECTOR}{
      user-select:text!important;
    }
    .smart__sel{
      background:#b3d4ff!important;
    }
    #smart__box{
      position:absolute;
      pointer-events:none;
      z-index:2147483647;
      border:2px solid #1a73e8;
      border-radius:2px;
      display:none;
    }
    body.smart__selecting,body.smart__selecting * {
      cursor: crosshair!important;
      }`;
  $("<style>").text(css).appendTo("head");

  const unlock = ($root) =>
    $root.find("td, th").each(function () {
      this.onselectstart = this.onmousedown = null;
    });
  unlock($(document));
  new MutationObserver((m) =>
    m.forEach((r) => $(r.addedNodes).each((_, n) => unlock($(n))))
  ).observe(document.body, { childList: true, subtree: true });

  const rIdx = ($c) => {
    const $row = $c.closest("tr");
    const $table = $row.closest("table");
    return $table.find("tr").index($row);
  };

  const cIdx = ($c) => {
    const $cell = $c.closest("td, th");
    const $row = $cell.closest("tr");
    return $row.find("td, th").index($cell);
  };

  const getCell = (r, c) => {
    const $rows = $scope.find("tr");
    const $row = $rows.eq(r);
    return $row.find("td, th").eq(c);
  };

  const key = (r, c) => `${r}|${c}`;

  const sel = new Set();
  const $box = $('<div id="smart__box">').appendTo("body");
  let dragging = false,
    $scope = $(document);
  let startR = 0,
    startC = 0,
    curR = 0,
    curC = 0;

  const paint = () => {
    $(CELL_SELECTOR, $scope).removeClass("smart__sel");
    sel.forEach((k) => {
      const [r, c] = k.split("|").map(Number);
      const $cel = getCell(r, c);
      $cel.addClass("smart__sel");
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
    const $start = $(CELL_SELECTOR, $scope)
      .filter((_, x) => rIdx($(x)) === startR && cIdx($(x)) === startC)
      .first();
    const $cur = $(CELL_SELECTOR, $scope)
      .filter((_, x) => rIdx($(x)) === curR && cIdx($(x)) === curC)
      .first();
    if (!$start.length || !$cur.length) {
      $box.hide();
      return;
    }
    const r1 = $start[0].getBoundingClientRect();
    const r2 = $cur[0].getBoundingClientRect();
    const left = Math.min(r1.left, r2.left) + scrollX;
    const top = Math.min(r1.top, r2.top) + scrollY;
    const right = Math.max(r1.right, r2.right) + scrollX;
    const bottom = Math.max(r1.bottom, r2.bottom) + scrollY;
    $box.css({ left, top, width: right - left, height: bottom - top });
  };

  const copySel = () => {
    if (!sel.size) return;
    const rows = [...sel].reduce((m, k) => {
      const [r, c] = k.split("|");
      (m[r] ??= []).push(+c);
      return m;
    }, {});
    const tsv = Object.keys(rows)
      .sort((a, b) => a - b)
      .map((r) =>
        rows[r]
          .sort((a, b) => a - b)
          .map((c) => {
            const $row = ROW_SELECTOR
              ? $scope.find(ROW_SELECTOR).eq(r)
              : $scope;
            const $cel = ROW_SELECTOR
              ? $row.children(CELL_SELECTOR).eq(c)
              : $row.children(CELL_SELECTOR).eq(r);
            return $cel.text().trim();
          })
          .join("\t")
      )
      .join("\n");
    GM_setClipboard(tsv);
  };

  const isAddMode = (e) => MODIFIER_KEY.primary(e) && MODIFIER_KEY.shift(e);

  $(document)
    .on("mousedown", CELL_SELECTOR, function (e) {
      const wantsSelect = MODIFIER_KEY.primary(e);
      if (e.button !== 0 || !wantsSelect) return;
      if (!isModifierHeld) $("body").addClass("smart__selecting");
      $scope = $(this).closest("table");
      startR = curR = rIdx($(this));
      startC = curC = cIdx($(this));
      if (!isAddMode(e)) sel.clear();
      rectKeys().forEach((k) => sel.add(k));
      rectKeys().forEach((k) => sel.add(k));
      paint();
      $box.show();
      updateBox();
      dragging = true;
      e.preventDefault();
    })

    .on("mousemove", function (e) {
      if (!dragging) return;
      const $cell = $(e.target).closest(CELL_SELECTOR);
      if (!$cell.length) return;
      curR = rIdx($cell);
      curC = cIdx($cell);
      if (!isAddMode(e)) sel.clear();
      rectKeys().forEach((k) => sel.add(k));
      paint();
      updateBox();
    })

    .on("mouseup", function (e) {
      if (!dragging) return;
      dragging = false;
      $("body").removeClass("smart__selecting");
      $box.hide();

      paint();

      const isClickOnly = startR === curR && startC === curC;
      const txt = $(e.target).closest(CELL_SELECTOR).text().trim();

      if (isClickOnly && txt) {
        GM_setClipboard(txt);
      } else {
        copySel();
      }
    });

  window.addEventListener(
    "keydown",
    (e) => {
      if (e.key === "Escape" || e.code === "Escape") {
        sel.clear();
        paint();
      }
    },
    true
  );

  let isModifierHeld = false;

  window.addEventListener("keydown", (e) => {
    if (MODIFIER_KEY.primary(e)) {
      isModifierHeld = true;
      document.body.classList.add("smart__selecting");
    }
  });

  window.addEventListener("keyup", (e) => {
    if (!MODIFIER_KEY.primary(e)) {
      isModifierHeld = false;
      if (!dragging) {
        document.body.classList.remove("smart__selecting");
      }
    }
  });
})(window.jQuery.noConflict(true));
