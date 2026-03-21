// ============================================================
// BIS Planner — Apps Script
//
// Copy this entire file to: Extensions > Apps Script
// Then save (Ctrl+S). The script runs automatically on cell edits.
// Optional: run addBISPlannerToolsMenu() once from the editor (Force refresh, Rebuild equipped index, timing log).
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
var GG_ACQ      = 6;  // F — Acquisition Type (heroic dungeon row detection; match _get_row_color)
var GG_DUNGEON  = 8;  // H
var GG_DIFFICULTY = 9; // I
var GG_EQUIP    = 13; // M — hidden; must match generate_bis_planner.py SHEET_FIELDS ("Equip")
var GG_EQUIP2   = 14; // N — hidden ("Equip2")
// Current Equipment column A uses "Ring 1" / "Ring 2" / "Trinket 1" / "Trinket 2"; planner column C stays Ring/Trinket.

// Current Equipment columns (1-indexed)
var CE_GEARTYPE = 1;  // A
var CE_ITEMNAME = 2;  // B

// BIS Planner: Comparison (rich text) = K; hidden M/N = Equip / Equip2; O = CmpRaw; then Δ columns
var CMP_DISP_COL = 11;
var CMP_RAW_COL = 15;
/** Muted color when CmpRaw has no +/- stat segments (e.g. "Current Equipment Not Specified") */
var CMP_NOTICE_COLOR = "#B06000";
/** Heroic dungeon rows (dark fill + white text in export): lighter Comparison colors for contrast */
var CMP_DELTA_POS_DARK_ROW = "#A5D6A7";
var CMP_DELTA_NEG_DARK_ROW = "#FFAB91";
var CMP_ATTR_LINE_DARK_ROW = "#C8E6C9";
var CMP_NOTICE_DARK_ROW = "#FFE082";
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
/** JSON map: equippedSlotKey_(spec, gear) → row number (1-based). Rebuilt on onOpen; reconciles clearOtherEquipped. */
var BIS_EQUIPPED_INDEX_PROP = "BIS_EQUIPPED_INDEX_JSON_V1";

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

function interestIsEquippedState_(v) {
  var s = normalizeItemName(v);
  return s === "Equipped" || s === "Equipped 1" || s === "Equipped 2";
}

function gearIsRingOrTrinket_(gearType) {
  var g = normalizeItemName(gearType);
  return g === "Ring" || g === "Trinket";
}

/** 1 = slot 1 (Equipped / Equipped 1), 2 = Equipped 2 — only for Ring/Trinket planner rows. */
function ringTrinketSlotBucket_(interest) {
  return normalizeItemName(interest) === "Equipped 2" ? 2 : 1;
}

/** CE column A label for setCEItem / findCERow (Ring 1, Trinket 2, or Head, …). */
function ceSlotLabelForPlannerInterest_(gearType, interest) {
  var g = normalizeItemName(gearType);
  if (g === "Ring") {
    return ringTrinketSlotBucket_(interest) === 2 ? "Ring 2" : "Ring 1";
  }
  if (g === "Trinket") {
    return ringTrinketSlotBucket_(interest) === 2 ? "Trinket 2" : "Trinket 1";
  }
  return g;
}

/**
 * Current Equipment A value → planner Gear Type (C) + Interest to set when syncing from CE.
 * @return {{plannerGear:string, interest:string}}
 */
function plannerGearAndInterestFromCE_(ceGearLabel) {
  var a = normalizeItemName(ceGearLabel);
  if (a === "Ring 1") return { plannerGear: "Ring", interest: "Equipped 1" };
  if (a === "Ring 2") return { plannerGear: "Ring", interest: "Equipped 2" };
  if (a === "Trinket 1") return { plannerGear: "Trinket", interest: "Equipped 1" };
  if (a === "Trinket 2") return { plannerGear: "Trinket", interest: "Equipped 2" };
  return { plannerGear: a, interest: "Equipped" };
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
  applyInterestDropdownsByGearType_(bp);
  rebuildEquippedIndexFromPlanner_(bp);
  refreshComparisonAllSpecs_(bp);
}

/** Ring/Trinket: no plain Equipped. Other slots: no Equipped 1/2. Matches generate_bis_planner.py lists. */
var INTEREST_LIST_STANDARD = ["Pass", "Consider", "Need", "Equipped"];
var INTEREST_LIST_RING_TRINKET = ["Pass", "Consider", "Need", "Equipped 1", "Equipped 2"];

function interestValidationRuleForGear_(gearType) {
  var g = normalizeItemName(gearType);
  var list = g === "Ring" || g === "Trinket" ? INTEREST_LIST_RING_TRINKET : INTEREST_LIST_STANDARD;
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(true)
    .build();
}

/** Contiguous row numbers (1-based) → [[start,end], …] inclusive. */
function bisMergeContiguousRowNumbers_(rows) {
  if (!rows || rows.length === 0) return [];
  var sorted = rows.slice().sort(function (a, b) {
    return a - b;
  });
  var runs = [];
  var a = sorted[0];
  var b = sorted[0];
  for (var k = 1; k < sorted.length; k++) {
    var x = sorted[k];
    if (x === b + 1) b = x;
    else {
      runs.push([a, b]);
      a = b = x;
    }
  }
  runs.push([a, b]);
  return runs;
}

function bisApplyDataValidationToRowRuns_(sheet, col, rowNumbers, rule) {
  var runs = bisMergeContiguousRowNumbers_(rowNumbers);
  for (var ri = 0; ri < runs.length; ri++) {
    var r0 = runs[ri][0];
    var r1 = runs[ri][1];
    var h = r1 - r0 + 1;
    sheet.getRange(r0, col, h, 1).setDataValidation(rule);
  }
}

/**
 * Google Sheets: Interest (A) list depends on Gear type (C). Excel export uses the same split via two DV rules.
 * allowInvalid true so legacy cells (e.g. Equipped on a ring row) stay editable until changed.
 */
function applyInterestDropdownsByGearType_(bp) {
  var lastRow = bp.getLastRow();
  if (lastRow < BIS_FIRST_DATA_ROW) return;
  var n = lastRow - BIS_FIRST_DATA_ROW + 1;
  bp.getRange(BIS_FIRST_DATA_ROW, GG_INTEREST, n, 1).clearDataValidations();
  var gears = bp.getRange(BIS_FIRST_DATA_ROW, GG_GEARTYPE, n, 1).getValues();
  var rtRows = [];
  var stdRows = [];
  for (var i = 0; i < n; i++) {
    var g = normalizeItemName(gears[i][0]);
    var r = BIS_FIRST_DATA_ROW + i;
    if (g === "Ring" || g === "Trinket") rtRows.push(r);
    else stdRows.push(r);
  }
  var rtRule = interestValidationRuleForGear_("Ring");
  var stdRule = interestValidationRuleForGear_("Head");
  bisApplyDataValidationToRowRuns_(bp, GG_INTEREST, rtRows, rtRule);
  bisApplyDataValidationToRowRuns_(bp, GG_INTEREST, stdRows, stdRule);
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

    var plannerBCReuse = null;
    var nv = newValue;
    if (!gearIsRingOrTrinket_(gearType) && (nv === "Equipped 1" || nv === "Equipped 2")) {
      e.range.setValue("Equipped");
      nv = "Equipped";
    }
    if (interestIsEquippedState_(nv)) {
      e.range.setFontWeight("bold");
      var ceLabel = ceSlotLabelForPlannerInterest_(gearType, nv);
      var lastRowEq = bpSheet.getLastRow();
      var nEq = lastRowEq - BIS_FIRST_DATA_ROW + 1;
      var eqMap = getEquippedIndexMap_();
      var eqKey = equippedSlotKey_(spec, gearType, nv);
      var prevEqRow = eqMap[eqKey];

      if (nEq <= 0) {
        plannerBCReuse = [];
        eqMap[eqKey] = row;
        saveEquippedIndexMap_(eqMap);
        bisTimingStep_(tCtx, "after equip: no planner data rows");
      } else if (prevEqRow == null || prevEqRow === "") {
        var abcCold = bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, nEq, GG_GEARTYPE).getValues();
        bisTimingStep_(tCtx, "after read planner A:C (cold index: single read for clear + refresh)");
        clearOtherEquipped(bpSheet, row, spec, gearType, nv, abcCold);
        bisTimingStep_(tCtx, "after clearOtherEquipped (cold path)");
        plannerBCReuse = [];
        for (var pci = 0; pci < abcCold.length; pci++) {
          plannerBCReuse.push([
            abcCold[pci][GG_SPEC - 1],
            abcCold[pci][GG_GEARTYPE - 1],
          ]);
        }
        eqMap[eqKey] = row;
        saveEquippedIndexMap_(eqMap);
      } else {
        plannerBCReuse = bpSheet.getRange(BIS_FIRST_DATA_ROW, GG_SPEC, nEq, 2).getValues();
        bisTimingStep_(tCtx, "after read planner B:C (warm index: clear via index + refresh)");
        clearOtherEquippedUsingIndex_(bpSheet, row, spec, gearType, nv);
        bisTimingStep_(tCtx, "after clearOtherEquippedUsingIndex_");
      }
      setCEItem(spec, ceLabel, itemName);
      bisTimingStep_(tCtx, "after setCEItem");
    } else {
      e.range.setFontWeight("normal");
      if (interestIsEquippedState_(e.oldValue)) {
        var oldCe = ceSlotLabelForPlannerInterest_(gearType, e.oldValue);
        clearCEItemIfMatch(spec, oldCe, itemName);
        clearEquippedIndexIfRow_(row, spec, gearType, e.oldValue);
      }
      bisTimingStep_(tCtx, "after unequip branch");
    }
    SpreadsheetApp.flush();
    bisTimingStep_(tCtx, "after second flush");
    Utilities.sleep(EDIT_RECALC_WAIT_MS);
    bisTimingStep_(tCtx, "after Utilities.sleep(EDIT_RECALC_WAIT_MS)");
    refreshComparisonRichText(bpSheet, spec, gearType, tCtx, plannerBCReuse);
    bisTimingFinish_(tCtx, "BIS Interest edit → refreshComparisonRichText");
    return;
  }

  if (col <= GG_NAME) {
    if (col === GG_GEARTYPE) {
      bpSheet.getRange(row, GG_INTEREST).setDataValidation(interestValidationRuleForGear_(e.range.getValue()));
    }
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
function clearOtherEquipped(bpSheet, currentRow, spec, gearType, interestValue, plannerABCOpt) {
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
  var rt = gearIsRingOrTrinket_(gearType);
  var bucket = rt ? ringTrinketSlotBucket_(interestValue) : 0;
  var toClear = [];

  for (var i = 0; i < data.length; i++) {
    var r = i + BIS_FIRST_DATA_ROW;
    if (r === currentRow) continue;
    var iv = data[i][GG_INTEREST - 1];
    if (!interestIsEquippedState_(iv)) continue;
    if (normalizeItemName(data[i][GG_SPEC - 1]) !== specN) continue;
    if (normalizeItemName(data[i][GG_GEARTYPE - 1]) !== gearN) continue;
    if (rt && ringTrinketSlotBucket_(iv) !== bucket) continue;
    toClear.push("A" + r);
  }

  if (toClear.length === 0) return;
  var rl = bpSheet.getRangeList(toClear);
  rl.setValue("");
  rl.setFontWeight("normal");
}

/** Index key: Ring/Trinket include slot 1 vs 2; other gear stays spec+gear. */
function equippedSlotKey_(spec, gearType, interest) {
  var s = normalizeItemName(spec);
  var g = normalizeItemName(gearType);
  if (gearIsRingOrTrinket_(g)) {
    var slot = ringTrinketSlotBucket_(interest || "Equipped");
    return s + "\x1f" + g + "\x1f" + String(slot);
  }
  return s + "\x1f" + g;
}

function getEquippedIndexMap_() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(BIS_EQUIPPED_INDEX_PROP);
    if (!raw) return {};
    var o = JSON.parse(raw);
    return o && typeof o === "object" ? o : {};
  } catch (e) {
    return {};
  }
}

function saveEquippedIndexMap_(map) {
  try {
    PropertiesService.getDocumentProperties().setProperty(
      BIS_EQUIPPED_INDEX_PROP,
      JSON.stringify(map)
    );
  } catch (e2) {}
}

/** Full scan A:C — source of truth for equipped row per (spec, slot). Call onOpen / repair menu. */
function rebuildEquippedIndexFromPlanner_(bpSheet) {
  var lastRow = bpSheet.getLastRow();
  var n = lastRow - BIS_FIRST_DATA_ROW + 1;
  if (n <= 0) {
    saveEquippedIndexMap_({});
    return;
  }
  var data = bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, n, GG_GEARTYPE).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var intr = data[i][GG_INTEREST - 1];
    if (!interestIsEquippedState_(intr)) continue;
    var key = equippedSlotKey_(
      data[i][GG_SPEC - 1],
      data[i][GG_GEARTYPE - 1],
      intr
    );
    map[key] = i + BIS_FIRST_DATA_ROW;
  }
  saveEquippedIndexMap_(map);
}

/**
 * Clears other Equipped for same spec+slot using DocumentProperties index; falls back to full scan
 * when index is cold or stale row. Updates index to currentRow.
 */
function clearOtherEquippedUsingIndex_(bpSheet, currentRow, spec, gearType, interestValue) {
  var key = equippedSlotKey_(spec, gearType, interestValue);
  var map = getEquippedIndexMap_();
  var prev = map[key];
  var specN = normalizeItemName(spec);
  var gearN = normalizeItemName(gearType);

  if (prev == null || prev === "") {
    clearOtherEquipped(bpSheet, currentRow, spec, gearType, interestValue, null);
  } else if (prev !== currentRow) {
    var v = bpSheet.getRange(prev, 1, 1, GG_GEARTYPE).getValues()[0];
    if (
      interestIsEquippedState_(v[GG_INTEREST - 1]) &&
      normalizeItemName(v[GG_SPEC - 1]) === specN &&
      normalizeItemName(v[GG_GEARTYPE - 1]) === gearN &&
      (!gearIsRingOrTrinket_(gearType) ||
        ringTrinketSlotBucket_(v[GG_INTEREST - 1]) === ringTrinketSlotBucket_(interestValue))
    ) {
      var rl = bpSheet.getRangeList(["A" + prev]);
      rl.setValue("");
      rl.setFontWeight("normal");
    } else {
      clearOtherEquipped(bpSheet, currentRow, spec, gearType, interestValue, null);
    }
  }
  map[key] = currentRow;
  saveEquippedIndexMap_(map);
}

/** After unequipping this row, drop index entry if it pointed here. */
function clearEquippedIndexIfRow_(row, spec, gearType, interest) {
  var key = equippedSlotKey_(spec, gearType, interest);
  var map = getEquippedIndexMap_();
  var prev = map[key];
  if (prev === row) {
    delete map[key];
    saveEquippedIndexMap_(map);
  }
}

/** Sync index from an in-memory planner block (columns through Name). */
function syncEquippedIndexForSpecGearFromPlannerData_(spec, gearType, data) {
  var map = getEquippedIndexMap_();
  var specN = normalizeItemName(spec);
  var gearN = normalizeItemName(gearType);

  if (gearIsRingOrTrinket_(gearType)) {
    var k1 = equippedSlotKey_(spec, gearType, "Equipped 1");
    var k2 = equippedSlotKey_(spec, gearType, "Equipped 2");
    var f1 = null;
    var f2 = null;
    for (var i = 0; i < data.length; i++) {
      if (normalizeItemName(data[i][GG_SPEC - 1]) !== specN) continue;
      if (normalizeItemName(data[i][GG_GEARTYPE - 1]) !== gearN) continue;
      var iv = data[i][GG_INTEREST - 1];
      if (!interestIsEquippedState_(iv)) continue;
      var r = i + BIS_FIRST_DATA_ROW;
      if (ringTrinketSlotBucket_(iv) === 2) f2 = r;
      else f1 = r;
    }
    if (f1) map[k1] = f1;
    else delete map[k1];
    if (f2) map[k2] = f2;
    else delete map[k2];
  } else {
    var key = equippedSlotKey_(spec, gearType, "Equipped");
    var found = null;
    for (var j = 0; j < data.length; j++) {
      if (!interestIsEquippedState_(data[j][GG_INTEREST - 1])) continue;
      if (normalizeItemName(data[j][GG_SPEC - 1]) !== specN) continue;
      if (normalizeItemName(data[j][GG_GEARTYPE - 1]) !== gearN) continue;
      found = j + BIS_FIRST_DATA_ROW;
      break;
    }
    if (found) map[key] = found;
    else delete map[key];
  }
  saveEquippedIndexMap_(map);
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

  var plannerGear = plannerGearAndInterestFromCE_(gearType).plannerGear;

  SpreadsheetApp.flush();
  bisTimingStep_(tCtx, "after flush");
  Utilities.sleep(EDIT_RECALC_WAIT_MS);
  bisTimingStep_(tCtx, "after sleep");
  refreshComparisonRichText(bpSheet, spec, plannerGear, tCtx, null);
  bisTimingFinish_(tCtx, "Current Equipment edit → refreshComparisonRichText");
}

function findSpecForCERow(ceSheet, targetRow) {
  if (targetRow < 1) return null;
  var vals = ceSheet.getRange(1, CE_GEARTYPE, targetRow, 1).getValues();
  for (var idx = targetRow - 1; idx >= 0; idx--) {
    var s = String(vals[idx][0]).trim();
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

/**
 * @param {?Array<Array<*>>} gearColValuesOpt - optional column A values from row 1..lastRow (same as getValues).
 */
function findCERow(ceSheet, spec, gearType, gearColValuesOpt) {
  var data = gearColValuesOpt;
  if (data == null) {
    var lastRow = ceSheet.getLastRow();
    if (lastRow < 1) return null;
    data = ceSheet.getRange(1, CE_GEARTYPE, lastRow, 1).getValues();
  }
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

/** @param {string} ceGearLabel - Current Equipment column A (e.g. Ring 1, Head). */
function setInterestForItem(bpSheet, spec, ceGearLabel, itemName) {
  var want = normalizeItemName(itemName);
  if (!want) return;

  var pi = plannerGearAndInterestFromCE_(ceGearLabel);
  var pGear = pi.plannerGear;
  var wantInterest = pi.interest;

  var lastRow = bpSheet.getLastRow();
  var numRows = lastRow - BIS_FIRST_DATA_ROW + 1;
  if (numRows <= 0) return;
  var data = bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, numRows, GG_NAME).getValues();

  for (var i = 0; i < data.length; i++) {
    var r = i + BIS_FIRST_DATA_ROW;
    if (normalizeItemName(data[i][GG_SPEC - 1]) !== normalizeItemName(spec)) continue;
    if (normalizeItemName(data[i][GG_GEARTYPE - 1]) !== normalizeItemName(pGear)) continue;
    var curI = data[i][GG_INTEREST - 1];
    if (normalizeItemName(data[i][GG_NAME - 1]) === want) {
      if (normalizeItemName(curI) !== normalizeItemName(wantInterest)) {
        bpSheet.getRange(r, GG_INTEREST).setValue(wantInterest);
        bpSheet.getRange(r, GG_INTEREST).setFontWeight("bold");
      }
    } else if (
      interestIsEquippedState_(curI) &&
      (!gearIsRingOrTrinket_(pGear) ||
        ringTrinketSlotBucket_(curI) === ringTrinketSlotBucket_(wantInterest))
    ) {
      bpSheet.getRange(r, GG_INTEREST).setValue("");
      bpSheet.getRange(r, GG_INTEREST).setFontWeight("normal");
    }
  }
  syncEquippedIndexForSpecGearFromPlannerData_(spec, pGear, data);
}

/** @param {string} ceGearLabel - Current Equipment column A */
function clearInterestForItem(bpSheet, spec, ceGearLabel, itemName) {
  var want = normalizeItemName(itemName);
  if (!want) return;

  var pi = plannerGearAndInterestFromCE_(ceGearLabel);
  var pGear = pi.plannerGear;
  var wantInterest = pi.interest;

  var lastRow = bpSheet.getLastRow();
  var numRows = lastRow - BIS_FIRST_DATA_ROW + 1;
  if (numRows <= 0) return;
  var data = bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, numRows, GG_NAME).getValues();

  for (var i = 0; i < data.length; i++) {
    var r = i + BIS_FIRST_DATA_ROW;
    if (normalizeItemName(data[i][GG_SPEC - 1]) !== normalizeItemName(spec)) continue;
    if (normalizeItemName(data[i][GG_GEARTYPE - 1]) !== normalizeItemName(pGear)) continue;
    if (normalizeItemName(data[i][GG_NAME - 1]) !== want) continue;
    if (normalizeItemName(data[i][GG_INTEREST - 1]) !== normalizeItemName(wantInterest)) continue;
    bpSheet.getRange(r, GG_INTEREST).setValue("");
    bpSheet.getRange(r, GG_INTEREST).setFontWeight("normal");
  }
  syncEquippedIndexForSpecGearFromPlannerData_(spec, pGear, data);
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
 * Applies green/red Rich Text to Comparison (K) from CmpRaw display (O).
 * Mirror fallback uses batched =N{row} (same link as R1C1 =RC[offset]).
 * @param {?string} specFilter - If set, only rows whose Spec (B) matches.
 * @param {?string} gearTypeFilter - With specFilter, only rows whose Gear type (C) matches (same slot
 *   as a CE or Interest edit). Omit or null to refresh every row in the spec (e.g. onOpen).
 * @param {?{start:number, last:number, rows:Array<string>}} timingCtx - optional; pass null from onOpen / menu.
 * @param {?Array<Array<*>>} plannerABCReuse - optional B:C (2 cols) or A:C (3 cols) getValues snapshot, rows aligned
 *   to BIS_FIRST_DATA_ROW; when spec+gear scoped, skips a second B+C read (equip path passes B:C only).
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
  /** F:I per planner data row — Acquisition, Quest, Dungeon, Difficulty (only when block B:O not loaded). */
  var metaAcqDungeonDiff = null;

  if (!specFilterNorm && !gearFilterNorm) {
    blockDisplay = bpSheet
      .getRange(BIS_FIRST_DATA_ROW, GG_SPEC, cmpNumRows, offNInBlock + 1)
      .getDisplayValues();
    kFormulas = bpSheet.getRange(BIS_FIRST_DATA_ROW, CMP_DISP_COL, cmpNumRows, 1).getFormulas();
    bisTimingStep_(timingCtx, "refresh: after reads (B:O display + K formulas)");
    for (var ai = 0; ai < cmpNumRows; ai++) indices.push(ai);
  } else if (specFilterNorm && gearFilterNorm) {
    var usePlannerReuse =
      plannerABCReuse != null &&
      plannerABCReuse.length === cmpNumRows &&
      cmpNumRows > 0;
    var reuseW = usePlannerReuse ? plannerABCReuse[0].length : 0;
    if (reuseW !== 2 && reuseW !== 3) usePlannerReuse = false;
    var bc2 = null;
    if (!usePlannerReuse) {
      bc2 = bpSheet.getRange(BIS_FIRST_DATA_ROW, GG_SPEC, cmpNumRows, 2).getValues();
      bisTimingStep_(timingCtx, "refresh: after read B+C (scoped)");
    } else {
      bisTimingStep_(timingCtx, "refresh: scoped B+C from reused snapshot (no read)");
    }
    var sIx = reuseW === 2 ? 0 : GG_SPEC - 1;
    var gIx = reuseW === 2 ? 1 : GG_GEARTYPE - 1;
    for (var bi = 0; bi < cmpNumRows; bi++) {
      var sCell = usePlannerReuse ? plannerABCReuse[bi][sIx] : bc2[bi][0];
      var gCell = usePlannerReuse ? plannerABCReuse[bi][gIx] : bc2[bi][1];
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
      var nameG = bpSheet.getRange(r0g, GG_NAME, hg, 1).getDisplayValues();
      var equipG = bpSheet.getRange(r0g, GG_EQUIP, hg, 1).getDisplayValues();
      var equip2G = bpSheet.getRange(r0g, GG_EQUIP2, hg, 1).getDisplayValues();
      var gearG = bpSheet.getRange(r0g, GG_GEARTYPE, hg, 1).getDisplayValues();
      for (var jg = 0; jg < hg; jg++) {
        rowNK[sg + jg] = {
          d: dG[jg][0],
          k: kG[jg][0],
          n: nameG[jg][0],
          e: equipG[jg][0],
          e2: equip2G[jg][0],
          g: gearG[jg][0],
        };
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
      var nameS = bpSheet.getRange(r0s, GG_NAME, hs, 1).getDisplayValues();
      var equipS = bpSheet.getRange(r0s, GG_EQUIP, hs, 1).getDisplayValues();
      var equip2S = bpSheet.getRange(r0s, GG_EQUIP2, hs, 1).getDisplayValues();
      var gearS = bpSheet.getRange(r0s, GG_GEARTYPE, hs, 1).getDisplayValues();
      for (var js = 0; js < hs; js++) {
        rowNK[ss + js] = {
          d: dS[js][0],
          k: kS[js][0],
          n: nameS[js][0],
          e: equipS[js][0],
          e2: equip2S[js][0],
          g: gearS[js][0],
        };
      }
    }
    bisTimingStep_(timingCtx, "refresh: after scoped N+K reads (" + runsS.length + " run(s))");
  }

  if (blockDisplay == null && cmpNumRows > 0) {
    var metaW = GG_DIFFICULTY - GG_ACQ + 1;
    metaAcqDungeonDiff = bpSheet
      .getRange(BIS_FIRST_DATA_ROW, GG_ACQ, cmpNumRows, metaW)
      .getDisplayValues();
    bisTimingStep_(timingCtx, "refresh: after read F:I (heroic row detection)");
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
    var nameIx = GG_NAME - GG_SPEC;
    var equipIx = GG_EQUIP - GG_SPEC;
    var equip2Ix = GG_EQUIP2 - GG_SPEC;
    var gearIx = GG_GEARTYPE - GG_SPEC;
    var nmN;
    var gearRow;
    var eqN;
    var eq2N;
    if (blockDisplay != null) {
      nmN = blockDisplay[i][nameIx];
      gearRow = blockDisplay[i][gearIx];
      eqN = blockDisplay[i][equipIx];
      eq2N = blockDisplay[i][equip2Ix];
    } else {
      nmN = rowNK[i].n;
      gearRow = rowNK[i].g;
      eqN = rowNK[i].e;
      eq2N = rowNK[i].e2;
    }
    if (normalizeItemName(nmN) !== "") {
      var nmNorm = normalizeItemName(nmN);
      if (gearIsRingOrTrinket_(gearRow)) {
        var e1 = normalizeItemName(eqN);
        var e2 = normalizeItemName(eq2N);
        if (nmNorm === e1 && e2 !== "" && nmNorm === e2) {
          dispStr = "";
        }
      } else if (nmNorm === normalizeItemName(eqN)) {
        dispStr = "";
      }
    }
    var heroicDarkRow = plannerRowIsHeroicDungeonDarkFill_(
      blockDisplay != null
        ? [
            blockDisplay[i][GG_ACQ - GG_SPEC],
            blockDisplay[i][GG_DUNGEON - GG_SPEC],
            blockDisplay[i][GG_DIFFICULTY - GG_SPEC]
          ]
        : [
            metaAcqDungeonDiff[i][0],
            metaAcqDungeonDiff[i][2],
            metaAcqDungeonDiff[i][3]
          ]
    );
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
      var rich = buildComparisonRichTextValue(dispStr, heroicDarkRow);
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
  t = t.replace(/ Use:/g, "\nUse:");
  t = t.replace(/ Equip:/g, "\nEquip:");
  return t;
}

/** ItemDB "Attributes" / tooltip prose (trinkets, relics, wands, etc.). */
function cmpRawLooksLikeEquippedAttributes_(s) {
  var t = (s == null ? "" : String(s)).replace(/^\s+|\s+$/g, "");
  if (!t) return false;
  // RegExp() avoids "/…)/i" regex-literal parsing issues in some Apps Script runtimes.
  return new RegExp("^(Use:|Equip:|Proc:|Chance\\s+on\\s+hit:)", "i").test(t);
}

/**
 * CmpRaw should show "Currently Equipped:" before equipped-item attribute text. New workbooks get that
 * from the N-column formula; this covers older formulas and edge splits (e.g. relic/ranged Equip: lines).
 */
function applyCurrentlyEquippedAttrLabel_(line1, line2) {
  var L1 = line1 == null ? "" : String(line1);
  var L2 = line2 == null ? "" : String(line2);
  var t1 = L1.replace(/^\s+|\s+$/g, "");
  var t2 = L2.replace(/^\s+|\s+$/g, "");
  if (
    /^Currently Equipped:/i.test(t1) ||
    /^Currently Equipped:/i.test(t2) ||
    /^Equipped slot 1:/i.test(t1) ||
    /^Equipped slot 1:/i.test(t2) ||
    /^Equipped slot 2(:|$)/i.test(t1) ||
    /^Equipped slot 2(:|$)/i.test(t2)
  ) {
    return { line1: L1, line2: L2 };
  }
  if (t2 !== "" && cmpRawLooksLikeEquippedAttributes_(t2)) {
    return { line1: L1, line2: "Currently Equipped:\n" + L2 };
  }
  if (t2 === "" && t1 !== "" && cmpRawLooksLikeEquippedAttributes_(t1)) {
    return { line1: "", line2: "Currently Equipped:\n" + L1 };
  }
  return { line1: L1, line2: L2 };
}

/**
 * Matches generate_bis_planner.py _get_row_color / white-text heroic dungeon rows.
 * @param {Array<*>} triple — [Acquisition Type, Dungeon, Difficulty]
 */
function plannerRowIsHeroicDungeonDarkFill_(triple) {
  if (!triple || triple.length < 3) return false;
  if (normalizeItemName(triple[2]) !== "Heroic") return false;
  if (normalizeItemName(triple[0]) !== "Dungeon Drop") return false;
  var dung = normalizeItemName(triple[1]);
  if (!dung || dung.indexOf("---") === 0) return false;
  return true;
}

/** @return {boolean} */
function isDualCmpRawHeaderLine_(trimmed) {
  return (
    /^Slot 1 Comparison:$/.test(trimmed) ||
    /^Slot 2 comparison$/.test(trimmed) ||
    /^Equipped slot [12]$/.test(trimmed)
  );
}

/**
 * Ring/Trinket CmpRaw: blocks after "Slot 1 Comparison:" / "Slot 2 comparison" (legacy: Equipped slot 1/2).
 * @return {Array<{header:string, bodyLines:Array<string>}>}
 */
function dualCmpRawBlocksFromText_(text) {
  var lines = String(text).split("\n");
  var blocks = [];
  var i = 0;
  while (i < lines.length) {
    var trimmed = lines[i].replace(/^\s+|\s+$/g, "");
    if (isDualCmpRawHeaderLine_(trimmed)) {
      var header = lines[i];
      i++;
      var bodyStart = i;
      while (i < lines.length) {
        var t2 = lines[i].replace(/^\s+|\s+$/g, "");
        if (isDualCmpRawHeaderLine_(t2)) break;
        i++;
      }
      blocks.push({ header: header, bodyLines: lines.slice(bodyStart, i) });
    } else {
      i++;
    }
  }
  return blocks;
}

/**
 * Split slot body (stats vs tail) for Ring/Trinket dual CmpRaw.
 * Do not call applyCurrentlyEquippedAttrLabel_: the sheet formula already adds Equipped slot 1/2 labels;
 * injecting "Currently Equipped:" here duplicated text under Equipped slot 2 when the tail starts with Use:/Equip:.
 */
function labelSlotBodyForCmpRaw_(bodyLines) {
  if (!bodyLines || bodyLines.length === 0) {
    return { line1: "", line2: "" };
  }
  var k = -1;
  for (var i = 0; i < bodyLines.length; i++) {
    var bl = bodyLines[i].replace(/^\s+|\s+$/g, "");
    if (
      /^Currently Equipped:/i.test(bl) ||
      /^Equipped slot 1:/i.test(bl) ||
      /^Equipped slot 2(:|$)/i.test(bl) ||
      cmpRawLooksLikeEquippedAttributes_(bodyLines[i])
    ) {
      k = i;
      break;
    }
  }
  var line1 = k < 0 ? bodyLines.join("\n") : bodyLines.slice(0, k).join("\n");
  var line2 = k < 0 ? "" : bodyLines.slice(k).join("\n");
  return { line1: line1, line2: line2 };
}

/** Append comma-merged +/- stat segments to out; pushes color runs (positions relative to current out). */
function appendMergedStatSegmentsTo_(out, runs, line1, colPos, colNeg) {
  var L1 = line1 == null ? "" : String(line1);
  var rawParts = L1.split(", ");
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
  for (var si = 0; si < segs.length; si++) {
    if (si > 0) {
      out.s += ", ";
    }
    var segStart = out.s.length;
    out.s += segs[si];
    var s = segs[si];
    var c = null;
    if (s.charAt(0) === "+") c = colPos;
    else if (s.charAt(0) === "-") c = colNeg;
    if (c) {
      runs.push({ start: segStart, end: out.s.length, color: c });
    }
  }
}

/**
 * @param {Array<{header:string, bodyLines:Array<string>}>} blocks
 * @return {?GoogleAppsScript.Spreadsheet.RichTextValue}
 */
function buildDualComparisonRichTextValue_(blocks, colPos, colNeg, colAttr, colNotice) {
  var out = { s: "" };
  var runs = [];
  for (var bi = 0; bi < blocks.length; bi++) {
    if (bi > 0) {
      out.s += "\n";
    }
    var blk = blocks[bi];
    var hStart = out.s.length;
    var hdr = blk.header;
    out.s += hdr + "\n";
    runs.push({ start: hStart, end: hStart + hdr.length, color: colNotice });

    var labeled = labelSlotBodyForCmpRaw_(blk.bodyLines);
    appendMergedStatSegmentsTo_(out, runs, labeled.line1, colPos, colNeg);
    if (labeled.line2) {
      var t1 = (labeled.line1 == null ? "" : String(labeled.line1)).replace(/^\s+|\s+$/g, "");
      if (t1 !== "") {
        out.s += "\n";
      }
      var l2s = out.s.length;
      out.s += labeled.line2;
      runs.push({ start: l2s, end: out.s.length, color: colAttr });
    }
  }

  var full = out.s;
  if (full === "") return null;

  var pad = "";
  if (/^[+\-]/.test(full)) {
    pad = "\u200B";
  }
  var fullText = pad + full;
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
  return builder.build();
}

/**
 * @param {string} display — CmpRaw display string
 * @param {boolean=} darkFillRow — heroic dungeon row (dark fill in exported sheet); lighter green/red
 */
function buildComparisonRichTextValue(display, darkFillRow) {
  if (display == null) return null;
  var text = normalizeComparisonDisplay(display);
  if (text === "") return null;

  var colPos = darkFillRow ? CMP_DELTA_POS_DARK_ROW : "#0d652d";
  var colNeg = darkFillRow ? CMP_DELTA_NEG_DARK_ROW : "#c5221f";
  var colAttr = darkFillRow ? CMP_ATTR_LINE_DARK_ROW : "#0d652d";
  var colNotice = darkFillRow ? CMP_NOTICE_DARK_ROW : CMP_NOTICE_COLOR;

  var dualBlocks = null;
  if (
    text.indexOf("Slot 1 Comparison:") >= 0 ||
    text.indexOf("Slot 2 comparison") >= 0 ||
    text.indexOf("Equipped slot 1") >= 0 ||
    text.indexOf("Equipped slot 2") >= 0
  ) {
    dualBlocks = dualCmpRawBlocksFromText_(text);
  }
  if (dualBlocks != null && dualBlocks.length > 0) {
    return buildDualComparisonRichTextValue_(dualBlocks, colPos, colNeg, colAttr, colNotice);
  }

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

  var labeled = applyCurrentlyEquippedAttrLabel_(line1, line2);
  line1 = labeled.line1;
  line2 = labeled.line2;

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
    if (s.charAt(0) === "+") c = colPos;
    else if (s.charAt(0) === "-") c = colNeg;
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
        SpreadsheetApp.newTextStyle().setForegroundColor(colNotice).build()
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
        SpreadsheetApp.newTextStyle().setForegroundColor(colAttr).build()
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
    .addItem("Rebuild equipped index (repair)", "rebuildEquippedIndexMenu_")
    .addItem("Show last timing log", "showLastBISTimingLog_")
    .addToUi();
}

function rebuildEquippedIndexMenu_() {
  var bp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIS_PLANNER_SHEET);
  if (!bp) {
    try {
      SpreadsheetApp.getUi().alert('No sheet named "' + BIS_PLANNER_SHEET + '".');
    } catch (e0) {}
    return;
  }
  rebuildEquippedIndexFromPlanner_(bp);
  try {
    SpreadsheetApp.getUi().alert("Equipped index rebuilt from column A (Interest).");
  } catch (e1) {}
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
  rebuildEquippedIndexFromPlanner_(bp);
  refreshComparisonAllSpecs_(bp);
  try {
    SpreadsheetApp.getUi().alert("Comparison column refresh finished.");
  } catch (e1) {}
}
