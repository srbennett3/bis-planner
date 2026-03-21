// ============================================================
// BIS Planner — Apps Script
//
// Copy this entire file to: Extensions > Apps Script
// Then save (Ctrl+S). The script runs automatically on cell edits.
// Optional: run addBISPlannerToolsMenu() once from the editor (Force refresh + Show last timing log).
//
// Created by Steven Bennett 2026
// ============================================================

var BIS_PLANNER_SHEET = "BIS Planner";
var CURRENT_EQUIP_SHEET = "Current Equipment";

// First data row on each sheet (must match generate_bis_planner.py export_xlsx intro rows)
var BIS_FIRST_DATA_ROW = 4;
var CE_FIRST_DATA_ROW = 4;

// BIS Planner columns (1-indexed)
var GG_INTEREST = 1;  // A
var GG_SPEC     = 2;  // B
var GG_GEARTYPE = 3;  // C
var GG_NAME     = 4;  // D

// Current Equipment columns (1-indexed)
var CE_GEARTYPE = 1;  // A
var CE_ITEMNAME = 2;  // B

// BIS Planner: Comparison (rich text) = K, CmpRaw (formula) = N — hidden Equip/Δ cols are M–AH
var CMP_DISP_COL = 11;
var CMP_RAW_COL = 14;
/** Muted color when CmpRaw has no +/- stat segments (e.g. "Current Equipment Not Specified") */
var CMP_NOTICE_COLOR = "#B06000";
// Same-row mirror K→N without "=N4" (avoids rare parse issues); offset = CmpRaw − Comparison
var CMP_RAW_R1C1_OFFSET = CMP_RAW_COL - CMP_DISP_COL;

/** After CE / Interest edits: one short wait before reading CmpRaw (N). */
var EDIT_RECALC_WAIT_MS = 150;
/** Before each spec on open / force refresh (formulas still settling). */
var OPEN_RECALC_WAIT_MS = 550;
/** Between two open passes so CmpRaw can finish recalculating before the second read. */
var OPEN_SECOND_PASS_SLEEP_MS = 220;

/** Set false after profiling. When true: Executions → View logs, or BIS Planner tools → Show last timing log. */
var BIS_TIMING_DEBUG = true;
var BIS_TIMING_PROP_KEY = "BIS_LAST_TIMING";

/**
 * @return {?{start:number, last:number, rows:Array<string>}}
 */
function bisTimingCreate_() {
  if (!BIS_TIMING_DEBUG) return null;
  return { start: Date.now(), last: Date.now(), rows: [] };
}

function bisTimingStep_(ctx, name) {
  if (!ctx) return;
  var now = Date.now();
  var stepMs = now - ctx.last;
  var totalMs = now - ctx.start;
  ctx.rows.push(stepMs + " ms step | " + totalMs + " ms total — " + name);
  ctx.last = now;
}

function bisTimingFinish_(ctx, title) {
  if (!ctx || ctx.rows.length === 0) return;
  var total = Date.now() - ctx.start;
  var header = title + " — total " + total + " ms";
  var body = header + "\n" + ctx.rows.join("\n");
  Logger.log(body);
  try {
    var max = 9000;
    PropertiesService.getDocumentProperties().setProperty(
      BIS_TIMING_PROP_KEY,
      body.length > max ? body.substring(0, max) + "\n…(truncated)" : body
    );
  } catch (eProp) {}
}

function normalizeItemName(v) {
  if (v == null || v === "") return "";
  return String(v).trim();
}

function titleCaseWords(s) {
  var parts = String(s).split(/\s+/);
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var w = parts[i];
    if (!w) continue;
    out.push(w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
  }
  return out.join(" ");
}

/** Matches generate_bis_planner.py: f"--- {spec_name.upper()} ---" */
function specDisplayNameFromSectionHeader(sectionLine) {
  var m = String(sectionLine).trim().match(/^---\s*(.+?)\s*---\s*$/);
  if (!m) return null;
  var inner = m[1].trim();
  if (!inner) return null;
  return titleCaseWords(inner);
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  if (sheetName === BIS_PLANNER_SHEET) {
    handleBISPlannerEdit(e, sheet);
  } else if (sheetName === CURRENT_EQUIP_SHEET) {
    handleCurrentEquipEdit(e, sheet);
  }
}

function onOpen() {
  var bp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIS_PLANNER_SHEET);
  if (!bp) return;
  refreshComparisonAllSpecs_(bp);
}

/** Distinct spec names (column B) for scoped refresh — avoids touching every row on each edit. */
function uniqueSpecFiltersFromPlanner_(bpSheet) {
  var lastRow = bpSheet.getLastRow();
  if (lastRow < BIS_FIRST_DATA_ROW) return [];
  var n = lastRow - BIS_FIRST_DATA_ROW + 1;
  var vals = bpSheet.getRange(BIS_FIRST_DATA_ROW, GG_SPEC, n, 1).getValues();
  var seen = {};
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var s = normalizeItemName(vals[i][0]);
    if (!s || seen[s]) continue;
    seen[s] = true;
    out.push(s);
  }
  return out;
}

/**
 * Full planner refresh: one spec at a time. Two light passes per spec (open only) catch rows
 * where CmpRaw was still blank on the first read; edit paths use a single pass.
 */
function refreshComparisonAllSpecs_(bp) {
  var specs = uniqueSpecFiltersFromPlanner_(bp);
  if (specs.length === 0) {
    Utilities.sleep(OPEN_RECALC_WAIT_MS);
    refreshComparisonRichText(bp, null, null, null, null);
    Utilities.sleep(OPEN_SECOND_PASS_SLEEP_MS);
    refreshComparisonRichText(bp, null, null, null, null);
    return;
  }
  for (var si = 0; si < specs.length; si++) {
    Utilities.sleep(OPEN_RECALC_WAIT_MS);
    refreshComparisonRichText(bp, specs[si], null, null, null);
    Utilities.sleep(OPEN_SECOND_PASS_SLEEP_MS);
    refreshComparisonRichText(bp, specs[si], null, null, null);
  }
}

// ============================================================
// BIS Planner edits (Interest column)
// ============================================================

function handleBISPlannerEdit(e, bpSheet) {
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row < BIS_FIRST_DATA_ROW) return;

  if (col === GG_INTEREST) {
    var tCtx = bisTimingCreate_();
    SpreadsheetApp.flush();
    bisTimingStep_(tCtx, "after first flush");
    var newValue = e.range.getValue();
    var meta = bpSheet.getRange(row, GG_SPEC, 1, GG_NAME - GG_SPEC + 1).getValues()[0];
    var spec = meta[0];
    var gearType = meta[1];
    var itemName = normalizeItemName(meta[2]);
    bisTimingStep_(tCtx, "after read spec / gear / name (one range B:D)");

    if (!normalizeItemName(spec) || !normalizeItemName(gearType)) {
      bisTimingFinish_(tCtx, "BIS Interest edit — early exit (missing spec or gear)");
      return;
    }

    var plannerABCReuse = null;
    if (newValue === "Equipped") {
      e.range.setFontWeight("bold");
      var lastRowEq = bpSheet.getLastRow();
      var nEq = lastRowEq - BIS_FIRST_DATA_ROW + 1;
      plannerABCReuse =
        nEq > 0
          ? bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, nEq, GG_GEARTYPE).getValues()
          : [];
      bisTimingStep_(tCtx, "after read planner A:C (shared: clear + refresh filter)");
      clearOtherEquipped(bpSheet, row, spec, gearType, plannerABCReuse);
      bisTimingStep_(tCtx, "after clearOtherEquipped");
      setCEItem(spec, gearType, itemName);
      bisTimingStep_(tCtx, "after setCEItem");
    } else {
      e.range.setFontWeight("normal");
      if (e.oldValue === "Equipped") {
        clearCEItemIfMatch(spec, gearType, itemName);
      }
      bisTimingStep_(tCtx, "after unequip branch");
    }
    SpreadsheetApp.flush();
    bisTimingStep_(tCtx, "after second flush");
    Utilities.sleep(EDIT_RECALC_WAIT_MS);
    bisTimingStep_(tCtx, "after Utilities.sleep(EDIT_RECALC_WAIT_MS)");
    refreshComparisonRichText(bpSheet, spec, gearType, tCtx, plannerABCReuse);
    bisTimingFinish_(tCtx, "BIS Interest edit → refreshComparisonRichText");
    return;
  }

  if (col <= GG_NAME) {
    var sg = bpSheet.getRange(row, GG_SPEC, 1, 2).getValues()[0];
    var specForRow = sg[0];
    var gearForRow = sg[1];
    SpreadsheetApp.flush();
    refreshComparisonRichText(bpSheet, specForRow, gearForRow, null, null);
  }
}

/**
 * @param {?Array<Array<*>>} plannerABCOpt - from getValues A:C (same height as data rows); if null, reads sheet.
 */
function clearOtherEquipped(bpSheet, currentRow, spec, gearType, plannerABCOpt) {
  var data;
  if (plannerABCOpt != null && plannerABCOpt.length > 0) {
    data = plannerABCOpt;
  } else {
    var lastRow = bpSheet.getLastRow();
    var numRows = lastRow - BIS_FIRST_DATA_ROW + 1;
    if (numRows <= 0) return;
    data = bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, numRows, GG_GEARTYPE).getValues();
  }
  var specN = normalizeItemName(spec);
  var gearN = normalizeItemName(gearType);
  var toClear = [];

  for (var i = 0; i < data.length; i++) {
    var r = i + BIS_FIRST_DATA_ROW;
    if (r === currentRow) continue;
    if (data[i][GG_INTEREST - 1] === "Equipped" &&
        normalizeItemName(data[i][GG_SPEC - 1]) === specN &&
        normalizeItemName(data[i][GG_GEARTYPE - 1]) === gearN) {
      toClear.push("A" + r);
    }
  }

  if (toClear.length === 0) return;
  var rl = bpSheet.getRangeList(toClear);
  rl.setValue("");
  rl.setFontWeight("normal");
}

// ============================================================
// Current Equipment edits (Item Name column)
// ============================================================

function handleCurrentEquipEdit(e, ceSheet) {
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (col !== CE_ITEMNAME || row < CE_FIRST_DATA_ROW) return;

  var gearType = normalizeItemName(ceSheet.getRange(row, CE_GEARTYPE).getValue());
  if (!gearType || gearType.indexOf("---") === 0) return;

  var spec = findSpecForCERow(ceSheet, row);
  if (!spec) return;

  var newItemName = normalizeItemName(e.range.getValue());
  var oldItemName = normalizeItemName(e.oldValue);

  var bpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIS_PLANNER_SHEET);
  if (!bpSheet) return;

  var tCtx = bisTimingCreate_();

  if (oldItemName) {
    clearInterestForItem(bpSheet, spec, gearType, oldItemName);
    bisTimingStep_(tCtx, "after clearInterestForItem");
  }

  if (newItemName) {
    setInterestForItem(bpSheet, spec, gearType, newItemName);
    bisTimingStep_(tCtx, "after setInterestForItem");
  }

  SpreadsheetApp.flush();
  bisTimingStep_(tCtx, "after flush");
  Utilities.sleep(EDIT_RECALC_WAIT_MS);
  bisTimingStep_(tCtx, "after sleep");
  refreshComparisonRichText(bpSheet, spec, gearType, tCtx, null);
  bisTimingFinish_(tCtx, "Current Equipment edit → refreshComparisonRichText");
}

function findSpecForCERow(ceSheet, targetRow) {
  for (var r = targetRow; r >= 1; r--) {
    var val = ceSheet.getRange(r, CE_GEARTYPE).getValue();
    var s = String(val).trim();
    if (s.indexOf("---") === 0) {
      return specDisplayNameFromSectionHeader(s);
    }
  }
  return null;
}

// ============================================================
// Helpers: sync between sheets
// ============================================================

function setCEItem(spec, gearType, itemName) {
  var ceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CURRENT_EQUIP_SHEET);
  if (!ceSheet) return;

  var ceRow = findCERow(ceSheet, spec, gearType);
  if (ceRow) {
    ceSheet.getRange(ceRow, CE_ITEMNAME).setValue(normalizeItemName(itemName));
  }
}

function clearCEItemIfMatch(spec, gearType, itemName) {
  var ceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CURRENT_EQUIP_SHEET);
  if (!ceSheet) return;

  var ceRow = findCERow(ceSheet, spec, gearType);
  if (ceRow) {
    var current = normalizeItemName(ceSheet.getRange(ceRow, CE_ITEMNAME).getValue());
    if (current === normalizeItemName(itemName)) {
      ceSheet.getRange(ceRow, CE_ITEMNAME).setValue("");
    }
  }
}

function findCERow(ceSheet, spec, gearType) {
  var lastRow = ceSheet.getLastRow();
  var data = ceSheet.getRange(1, CE_GEARTYPE, lastRow, 1).getValues();
  var inSection = false;

  for (var i = 0; i < data.length; i++) {
    var val = String(data[i][0]);
    if (val.indexOf("---") === 0) {
      inSection = val.indexOf(spec.toUpperCase()) !== -1;
      continue;
    }
    if (inSection && normalizeItemName(val) === normalizeItemName(gearType)) {
      return i + 1;
    }
  }
  return null;
}

function setInterestForItem(bpSheet, spec, gearType, itemName) {
  var want = normalizeItemName(itemName);
  if (!want) return;

  var lastRow = bpSheet.getLastRow();
  var numRows = lastRow - BIS_FIRST_DATA_ROW + 1;
  if (numRows <= 0) return;
  var data = bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, numRows, GG_NAME).getValues();

  for (var i = 0; i < data.length; i++) {
    var r = i + BIS_FIRST_DATA_ROW;
    if (normalizeItemName(data[i][GG_SPEC - 1]) === normalizeItemName(spec) &&
        normalizeItemName(data[i][GG_GEARTYPE - 1]) === normalizeItemName(gearType)) {
      if (normalizeItemName(data[i][GG_NAME - 1]) === want) {
        if (data[i][GG_INTEREST - 1] !== "Equipped") {
          bpSheet.getRange(r, GG_INTEREST).setValue("Equipped");
          bpSheet.getRange(r, GG_INTEREST).setFontWeight("bold");
        }
      } else if (data[i][GG_INTEREST - 1] === "Equipped") {
        bpSheet.getRange(r, GG_INTEREST).setValue("");
        bpSheet.getRange(r, GG_INTEREST).setFontWeight("normal");
      }
    }
  }
}

function clearInterestForItem(bpSheet, spec, gearType, itemName) {
  var want = normalizeItemName(itemName);
  if (!want) return;

  var lastRow = bpSheet.getLastRow();
  var numRows = lastRow - BIS_FIRST_DATA_ROW + 1;
  if (numRows <= 0) return;
  var data = bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, numRows, GG_NAME).getValues();

  for (var i = 0; i < data.length; i++) {
    var r = i + BIS_FIRST_DATA_ROW;
    if (normalizeItemName(data[i][GG_SPEC - 1]) === normalizeItemName(spec) &&
        normalizeItemName(data[i][GG_GEARTYPE - 1]) === normalizeItemName(gearType) &&
        normalizeItemName(data[i][GG_NAME - 1]) === want &&
        data[i][GG_INTEREST - 1] === "Equipped") {
      bpSheet.getRange(r, GG_INTEREST).setValue("");
      bpSheet.getRange(r, GG_INTEREST).setFontWeight("normal");
    }
  }
}

// ============================================================
// Comparison column: Rich Text (Google Sheets only; CmpRaw = N, Comparison = K)
// ============================================================

/** 1-based column index to A1 letter(s) (e.g. 14 → "N"). */
function colLetterFromNum_(n) {
  var s = "";
  var x = n;
  while (x > 0) {
    var m = (x - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    x = ((x - 1) / 26) | 0;
  }
  return s;
}

/** True if K already mirrors CmpRaw on this row (=RC[offset] or =Nrow). */
function cellIsCmpRawMirrorFormula_(formula, row) {
  if (formula == null || formula === "") return false;
  var f = String(formula).replace(/^\s+|\s+$/g, "");
  if (new RegExp("RC\\s*\\[\\s*" + CMP_RAW_R1C1_OFFSET + "\\s*\\]").test(f)) return true;
  var L = colLetterFromNum_(CMP_RAW_COL);
  var re = new RegExp("^=\\s*\\$?" + L + "\\s*\\$?" + row + "\\s*$", "i");
  return re.test(f);
}

/**
 * Writes queued Comparison (K) updates in batches (clear + setRichTextValues / setFormulas) to cut round-trips.
 * @param {Array<{r:number, kind:string, rich?:GoogleAppsScript.Spreadsheet.RichTextValue, needClearBefore:boolean}>} writes
 */
function flushComparisonWrites_(bpSheet, writes, cmpRawColLetter, timingCtx) {
  if (!writes || writes.length === 0) {
    bisTimingStep_(timingCtx, "flushComparisonWrites_: 0 writes (skip)");
    return;
  }
  bisTimingStep_(timingCtx, "flushComparisonWrites_: start batched K writes");

  var segments = [];
  var cur = [writes[0]];
  for (var wi = 1; wi < writes.length; wi++) {
    if (writes[wi].r === writes[wi - 1].r + 1) {
      cur.push(writes[wi]);
    } else {
      segments.push(cur);
      cur = [writes[wi]];
    }
  }
  segments.push(cur);

  for (var si = 0; si < segments.length; si++) {
    var seg = segments[si];
    var j = 0;
    while (j < seg.length) {
      if (seg[j].kind === "rich") {
        var runEnd = j;
        while (
          runEnd + 1 < seg.length &&
          seg[runEnd + 1].kind === "rich" &&
          seg[runEnd + 1].r === seg[runEnd].r + 1
        ) {
          runEnd++;
        }
        var sub = j;
        while (sub <= runEnd) {
          var needClear = seg[sub].needClearBefore;
          var subEnd = sub;
          while (
            subEnd < runEnd &&
            seg[subEnd + 1].kind === "rich" &&
            seg[subEnd + 1].r === seg[subEnd].r + 1 &&
            seg[subEnd + 1].needClearBefore === needClear
          ) {
            subEnd++;
          }
          var r0 = seg[sub].r;
          var h = seg[subEnd].r - r0 + 1;
          if (needClear) {
            bpSheet.getRange(r0, CMP_DISP_COL, h, 1).clearContent();
          }
          var matrix = [];
          for (var t = sub; t <= subEnd; t++) {
            matrix.push([seg[t].rich]);
          }
          bpSheet.getRange(r0, CMP_DISP_COL, h, 1).setRichTextValues(matrix);
          sub = subEnd + 1;
        }
        j = runEnd + 1;
      } else {
        var mEnd = j;
        while (
          mEnd + 1 < seg.length &&
          seg[mEnd + 1].kind === "mirror" &&
          seg[mEnd + 1].r === seg[mEnd].r + 1
        ) {
          mEnd++;
        }
        var r0m = seg[j].r;
        var hm = seg[mEnd].r - r0m + 1;
        bpSheet.getRange(r0m, CMP_DISP_COL, hm, 1).clearContent();
        var fmatrix = [];
        for (var tm = j; tm <= mEnd; tm++) {
          fmatrix.push(["=" + cmpRawColLetter + seg[tm].r]);
        }
        bpSheet.getRange(r0m, CMP_DISP_COL, hm, 1).setFormulas(fmatrix);
        j = mEnd + 1;
      }
    }
  }
  bisTimingStep_(timingCtx, "flushComparisonWrites_: finished");
}

/** Sorted 0-based data-row indices → [[start, end], ...] inclusive contiguous runs. */
function bisIndicesToRuns_(indices) {
  if (!indices || indices.length === 0) return [];
  var runs = [];
  var a = indices[0];
  var b = indices[0];
  for (var u = 1; u < indices.length; u++) {
    var x = indices[u];
    if (x === b + 1) b = x;
    else {
      runs.push([a, b]);
      a = b = x;
    }
  }
  runs.push([a, b]);
  return runs;
}

/**
 * Applies green/red Rich Text to Comparison (K) from CmpRaw display (N).
 * Mirror fallback uses batched =N{row} (same link as R1C1 =RC[offset]).
 * @param {?string} specFilter - If set, only rows whose Spec (B) matches.
 * @param {?string} gearTypeFilter - With specFilter, only rows whose Gear type (C) matches (same slot
 *   as a CE or Interest edit). Omit or null to refresh every row in the spec (e.g. onOpen).
 * @param {?{start:number, last:number, rows:Array<string>}} timingCtx - optional; pass null from onOpen / menu.
 * @param {?Array<Array<*>>} plannerABCReuse - optional A:C getValues snapshot (rows aligned to BIS_FIRST_DATA_ROW);
 *   when spec+gear scoped, skips a second B+C read (Interest equip path passes same snapshot as clearOtherEquipped).
 */
function refreshComparisonRichText(bpSheet, specFilter, gearTypeFilter, timingCtx, plannerABCReuse) {
  var specFilterNorm =
    specFilter != null && String(specFilter).replace(/^\s+|\s+$/g, "") !== ""
      ? normalizeItemName(specFilter)
      : null;
  var gearFilterNorm =
    gearTypeFilter != null && String(gearTypeFilter).replace(/^\s+|\s+$/g, "") !== ""
      ? normalizeItemName(gearTypeFilter)
      : null;

  var lastRow = bpSheet.getLastRow();
  if (lastRow < BIS_FIRST_DATA_ROW) return;

  bisTimingStep_(timingCtx, "refreshComparisonRichText: enter");

  // Sheet.getRange(row, col, numRows, numColumns) — 3rd/4th are counts, not end row/col.
  var cmpNumRows = Math.max(0, lastRow - BIS_FIRST_DATA_ROW + 1);
  var cmpRawLetter = colLetterFromNum_(CMP_RAW_COL);
  var offNInBlock = CMP_RAW_COL - GG_SPEC;
  var indices = [];
  var blockDisplay = null;
  var kFormulas = null;
  var rowNK = null;

  if (!specFilterNorm && !gearFilterNorm) {
    blockDisplay = bpSheet
      .getRange(BIS_FIRST_DATA_ROW, GG_SPEC, cmpNumRows, offNInBlock + 1)
      .getDisplayValues();
    kFormulas = bpSheet.getRange(BIS_FIRST_DATA_ROW, CMP_DISP_COL, cmpNumRows, 1).getFormulas();
    bisTimingStep_(timingCtx, "refresh: after reads (B:N display + K formulas)");
    for (var ai = 0; ai < cmpNumRows; ai++) indices.push(ai);
  } else if (specFilterNorm && gearFilterNorm) {
    var useAbcReuse =
      plannerABCReuse != null &&
      plannerABCReuse.length === cmpNumRows &&
      cmpNumRows > 0;
    var bc2 = null;
    if (!useAbcReuse) {
      bc2 = bpSheet.getRange(BIS_FIRST_DATA_ROW, GG_SPEC, cmpNumRows, 2).getValues();
      bisTimingStep_(timingCtx, "refresh: after read B+C (scoped)");
    } else {
      bisTimingStep_(timingCtx, "refresh: scoped B+C from reused A:C (no read)");
    }
    for (var bi = 0; bi < cmpNumRows; bi++) {
      var sCell = useAbcReuse
        ? plannerABCReuse[bi][GG_SPEC - 1]
        : bc2[bi][0];
      var gCell = useAbcReuse
        ? plannerABCReuse[bi][GG_GEARTYPE - 1]
        : bc2[bi][1];
      if (normalizeItemName(sCell) !== specFilterNorm) continue;
      if (normalizeItemName(gCell) !== gearFilterNorm) continue;
      indices.push(bi);
    }
    if (indices.length === 0) {
      bisTimingStep_(timingCtx, "refresh: no matching rows (skip N/K)");
      bisTimingStep_(timingCtx, "refreshComparisonRichText: return");
      return;
    }
    rowNK = {};
    var runsG = bisIndicesToRuns_(indices);
    for (var gi = 0; gi < runsG.length; gi++) {
      var sg = runsG[gi][0];
      var pg = runsG[gi][1];
      var hg = pg - sg + 1;
      var r0g = BIS_FIRST_DATA_ROW + sg;
      var dG = bpSheet.getRange(r0g, CMP_RAW_COL, hg, 1).getDisplayValues();
      var kG = bpSheet.getRange(r0g, CMP_DISP_COL, hg, 1).getFormulas();
      for (var jg = 0; jg < hg; jg++) {
        rowNK[sg + jg] = { d: dG[jg][0], k: kG[jg][0] };
      }
    }
    bisTimingStep_(timingCtx, "refresh: after scoped N+K reads (" + runsG.length + " run(s))");
  } else {
    var bOnly = bpSheet.getRange(BIS_FIRST_DATA_ROW, GG_SPEC, cmpNumRows, 1).getValues();
    bisTimingStep_(timingCtx, "refresh: after read B (spec-scoped)");
    for (var ci = 0; ci < bOnly.length; ci++) {
      if (normalizeItemName(bOnly[ci][0]) !== specFilterNorm) continue;
      indices.push(ci);
    }
    if (indices.length === 0) {
      bisTimingStep_(timingCtx, "refresh: no matching rows (skip N/K)");
      bisTimingStep_(timingCtx, "refreshComparisonRichText: return");
      return;
    }
    rowNK = {};
    var runsS = bisIndicesToRuns_(indices);
    for (var si2 = 0; si2 < runsS.length; si2++) {
      var ss = runsS[si2][0];
      var ps = runsS[si2][1];
      var hs = ps - ss + 1;
      var r0s = BIS_FIRST_DATA_ROW + ss;
      var dS = bpSheet.getRange(r0s, CMP_RAW_COL, hs, 1).getDisplayValues();
      var kS = bpSheet.getRange(r0s, CMP_DISP_COL, hs, 1).getFormulas();
      for (var js = 0; js < hs; js++) {
        rowNK[ss + js] = { d: dS[js][0], k: kS[js][0] };
      }
    }
    bisTimingStep_(timingCtx, "refresh: after scoped N+K reads (" + runsS.length + " run(s))");
  }

  var writes = [];

  for (var ki = 0; ki < indices.length; ki++) {
    var i = indices[ki];
    var r = i + BIS_FIRST_DATA_ROW;
    var rawDisplay = blockDisplay != null ? blockDisplay[i][offNInBlock] : rowNK[i].d;
    var kFOrig = blockDisplay != null ? kFormulas[i][0] : rowNK[i].k;
    var bogus = kFOrig && /^=\s*[+\-]\d/.test(kFOrig);
    var kF = bogus ? "" : kFOrig;
    var dispStr = rawDisplay == null ? "" : String(rawDisplay);
    // Do not skip when N display is empty: leaving K unchanged preserves stale rich text (e.g. after
    // equipping). Empty CmpRaw → build returns null → mirror =RC[] so K shows blank until N fills, then
    // the sheet updates the mirror live; onOpen / second pass still reapplies colored rich text.
    // CmpRaw display errors (#VALUE!, #N/A, etc.) — mirror CmpRaw, no rich text
    if (dispStr.length > 0 && dispStr.charAt(0) === "#") {
      if (!cellIsCmpRawMirrorFormula_(kF, r)) {
        writes.push({ r: r, kind: "mirror", needClearBefore: true });
      }
      continue;
    }
    try {
      var rich = buildComparisonRichTextValue(dispStr);
      if (rich) {
        var hadFormula = kF && String(kF).replace(/^\s+|\s+$/g, "") !== "";
        writes.push({
          r: r,
          kind: "rich",
          rich: rich,
          needClearBefore: bogus || hadFormula,
        });
      } else {
        if (!cellIsCmpRawMirrorFormula_(kF, r)) {
          writes.push({ r: r, kind: "mirror", needClearBefore: true });
        }
      }
    } catch (err) {
      if (!cellIsCmpRawMirrorFormula_(kF, r)) {
        writes.push({ r: r, kind: "mirror", needClearBefore: true });
      }
    }
  }

  bisTimingStep_(timingCtx, "refresh: after CPU (filter + buildComparison queue)");
  flushComparisonWrites_(bpSheet, writes, cmpRawLetter, timingCtx);
  bisTimingStep_(timingCtx, "refreshComparisonRichText: return");
}

/**
 * Sheets sometimes formats negatives with Unicode minus (U+2212). That breaks /^[+-]/ parsing
 * and can confuse RichTextBuilder ranges; normalize before splitting and styling.
 */
function normalizeComparisonDisplay(s) {
  if (s == null) return "";
  var t = String(s);
  t = t.replace(/\u2212/g, "-"); // minus sign
  t = t.replace(/\uFE63/g, "-"); // small hyphen-minus
  t = t.replace(/\uFF0D/g, "-"); // fullwidth hyphen-minus
  t = t.replace(/\uFF0B/g, "+"); // fullwidth plus
  t = t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  t = t.trim();
  // Bad import / copy-paste can leave "=" in the displayed string; would become "=+400…" as a bogus formula
  while (t.length > 0 && t.charAt(0) === "=") {
    t = t.slice(1).trim();
  }
  return t;
}

function buildComparisonRichTextValue(display) {
  if (display == null) return null;
  var text = normalizeComparisonDisplay(display);
  if (text === "") return null;

  // Stats vs attributes: CmpRaw uses two newlines; fall back to one for older sheets
  var attSep = text.indexOf("\n\n");
  var line1;
  var line2;
  if (attSep !== -1) {
    line1 = text.substring(0, attSep);
    line2 = text.substring(attSep + 2);
  } else {
    var nl = text.indexOf("\n");
    line1 = nl === -1 ? text : text.substring(0, nl);
    line2 = nl === -1 ? "" : text.substring(nl + 1);
  }

  // CmpRaw uses ", " between diff segments; merge fragments that lack a leading +/- so commas inside stat text don't split runs.
  var rawParts = line1.split(", ");
  var segs = [];
  for (var pi = 0; pi < rawParts.length; pi++) {
    var p = rawParts[pi].trim();
    if (p.length === 0) continue;
    if (/^[+-]/.test(p)) {
      segs.push(p);
    } else if (segs.length > 0) {
      segs[segs.length - 1] += ", " + p;
    } else {
      segs.push(p);
    }
  }

  var out = "";
  var runs = [];
  var si;
  for (si = 0; si < segs.length; si++) {
    if (si > 0) {
      out += ", ";
    }
    var segStart = out.length;
    out += segs[si];
    var s = segs[si];
    var c = null;
    if (s.charAt(0) === "+") c = "#0d652d";
    else if (s.charAt(0) === "-") c = "#c5221f";
    if (c) {
      runs.push({ start: segStart, end: out.length, color: c });
    }
  }

  var line2Start = 0;
  var line1End = out.length;
  if (line2) {
    if (out) out += "\n\n";
    line2Start = out.length;
    out += line2;
  }

  if (out === "") return null;

  // Sheets treats cell values starting with + / - / = as formulas (setRichTextValue can still hit that path).
  // Invisible prefix keeps the cell as text/rich text so "+400 Armor" does not become "=+400 Armor…" / #ERROR.
  var pad = "";
  if (/^[+\-]/.test(out)) {
    pad = "\u200B";
  }
  var fullText = pad + out;
  var po = pad.length;
  var fullLen = fullText.length;

  var builder = SpreadsheetApp.newRichTextValue().setText(fullText);
  for (var j = 0; j < runs.length; j++) {
    var rr = runs[j];
    var rs = Math.max(0, Math.min(rr.start + po, fullLen));
    var re = Math.max(0, Math.min(rr.end + po, fullLen));
    if (rs < re) {
      builder.setTextStyle(
        rs,
        re,
        SpreadsheetApp.newTextStyle().setForegroundColor(rr.color).build()
      );
    }
  }
  // No +/- stat runs (e.g. "Current Equipment Not Specified") — color the stats line so it is not plain black.
  if (runs.length === 0 && line1End > 0) {
    var n1s = po;
    var n1e = Math.min(po + line1End, fullLen);
    if (n1s < n1e) {
      builder.setTextStyle(
        n1s,
        n1e,
        SpreadsheetApp.newTextStyle().setForegroundColor(CMP_NOTICE_COLOR).build()
      );
    }
  }
  if (line2) {
    var l2s = Math.max(0, Math.min(line2Start + po, fullLen));
    var l2e = fullLen;
    if (l2s < l2e) {
      builder.setTextStyle(
        l2s,
        l2e,
        SpreadsheetApp.newTextStyle().setForegroundColor("#0d652d").build()
      );
    }
  }
  return builder.build();
}

// ============================================================
// Menu (run addBISPlannerToolsMenu once from the script editor)
// ============================================================

function addBISPlannerToolsMenu() {
  SpreadsheetApp.getUi()
    .createMenu("BIS Planner tools")
    .addItem("Force refresh comparison colors", "forceRefreshComparisonColors")
    .addItem("Show last timing log", "showLastBISTimingLog_")
    .addToUi();
}

function showLastBISTimingLog_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(BIS_TIMING_PROP_KEY);
  if (!raw) {
    try {
      SpreadsheetApp.getUi().alert(
        "No timing saved yet. Set BIS_TIMING_DEBUG = true in the script, then edit Interest or CE."
      );
    } catch (e0) {}
    return;
  }
  try {
    var esc = raw.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    var html =
      '<p style="font-family:monospace;font-size:11px;white-space:pre-wrap;margin:0">' + esc + "</p>";
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(html).setWidth(520).setHeight(420),
      "Last BIS timing"
    );
  } catch (e1) {
    Logger.log(raw);
  }
}

/** Same logic as onOpen; refreshes spec-by-spec so it can finish within the execution limit. */
function forceRefreshComparisonColors() {
  var bp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIS_PLANNER_SHEET);
  if (!bp) {
    try {
      SpreadsheetApp.getUi().alert('No sheet named "' + BIS_PLANNER_SHEET + '".');
    } catch (e0) {}
    return;
  }
  refreshComparisonAllSpecs_(bp);
  try {
    SpreadsheetApp.getUi().alert("Comparison column refresh finished.");
  } catch (e1) {}
}
