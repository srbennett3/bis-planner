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
var GG_EQUIP    = 14; // N — hidden; must match generate_bis_planner.py SHEET_FIELDS ("Equip")
var GG_EQUIP2   = 15; // O — hidden ("Equip2")
// Current Equipment column A uses "Ring 1" / "Ring 2" / "Trinket 1" / "Trinket 2"; planner column C stays Ring/Trinket.

// Current Equipment columns (1-indexed)
var CE_GEARTYPE = 1;  // A
var CE_ITEMNAME = 2;  // B
// Layout must match generate_bis_planner.py (CE_GEM_HEADERS … CE_META_WEIGHT_COL).
var CE_GEM_FIRST_COL = 3;
/** Last column of ideal gem block (C–E); must match generate_bis_planner NUM_CE_GEM_COLS. */
var CE_GEM_LAST_COL = 5;
var CE_IDEAL_GEM_COL = 3;
var CE_IDEAL_META_COL = 4;
var CE_TOTAL_GEM_STATS_COL = 5;
var CE_STAT_FIRST_COL = 6;
var CE_STAT_LAST_COL = 27;
var CE_GEAR_SCORE_COL = 28;
var CE_PAWN_SNAPSHOT_FIRST_COL = 29;
var CE_CUSTOM_STASH_FIRST_COL = 51;
var CE_META_WEIGHT_COL = 73;
var CE_WEIGHT_ROW_LABEL = "Weights";
var CE_WEIGHT_MODE_PAWN = "Pawn Default";
var CE_WEIGHT_MODE_CUSTOM = "Custom";

/** ItemDB VLOOKUP: first STAT_COLUMNS column index (Armor). */
var ITEMDB_FIRST_STAT_COL = 7;
var ITEMDB_SOCKR_COL = 3;
var ITEMDB_SOCKY_COL = 4;
var ITEMDB_SOCKB_COL = 5;
var ITEMDB_SOCKM_COL = 6;

var GEMDB_SHEET = "GemDB";
var ITEMDB_SHEET = "ItemDB";

/** Pawn stat keys aggregated into each planner stat column (same order as STAT_COLUMNS in Python). */
var CE_PAWN_KEYS_FOR_STAT = [
  ["Armor"],
  [],
  ["Strength"],
  ["Agility"],
  ["Stamina"],
  ["Intellect"],
  ["Spirit"],
  ["Healing"],
  ["SpellDamage"],
  ["Ap"],
  ["Mp5"],
  ["DefenseRating"],
  ["DodgeRating"],
  ["ParryRating"],
  ["BlockRating"],
  ["BlockValue"],
  ["HitRating"],
  ["CritRating"],
  ["SpellHitRating"],
  ["SpellCritRating"],
  ["HasteRating", "SpellHasteRating"],
  ["ResilienceRating"],
];

/** Non-meta GemDB categories: best weighted gem fills all colored sockets (Pawn-style). */
var CE_GEM_NON_META_CATS = ["red", "yellow", "blue", "orange", "purple", "green"];

var CE_GEM_LOCKED_BG = "#E0E0E0";

// BIS Planner: K = Stat Comparison; L = % Upgrade; P = CmpRaw; Q/R = gear scores (Apps Script); hidden N–R
var GG_UPGRADE_COL = 12;
var CMP_DISP_COL = 11;
var CMP_RAW_COL = 16;
var GG_GEAR_SCORE_COL = 17;
var GG_GEAR_SCORE_PLUS_GEMS_COL = 18;
/** Muted color when CmpRaw has no +/- stat segments (other plain notices). */
var CMP_NOTICE_COLOR = "#B06000";
/** Exact CmpRaw line for empty CE — shown in black (not CMP_NOTICE_COLOR). */
var CMP_RAW_EQUIP_NOT_SPECIFIED = "Current Equipment Not Specified";
/** Heroic dungeon rows (dark fill + white text in export): lighter Comparison colors for contrast */
var CMP_DELTA_POS_DARK_ROW = "#A5D6A7";
var CMP_DELTA_NEG_DARK_ROW = "#FFAB91";
var CMP_ATTR_LINE_DARK_ROW = "#C8E6C9";
var CMP_NOTICE_DARK_ROW = "#FFE082";
// Same-row mirror K→P without "=P4" (avoids rare parse issues); offset = CmpRaw − Stat Comparison
var CMP_RAW_R1C1_OFFSET = CMP_RAW_COL - CMP_DISP_COL;

/** After CE / Interest edits: one short wait before reading CmpRaw (P). */
var EDIT_RECALC_WAIT_MS = 150;
/** Before each spec on open / force refresh (formulas still settling). */
var OPEN_RECALC_WAIT_MS = 550;
/** Between two open passes so CmpRaw can finish recalculating before the second read. */
var OPEN_SECOND_PASS_SLEEP_MS = 220;

/**
 * Prevents nested refreshComparisonRichText (e.g. installable onEdit firing while script writes K/L/Q/R),
 * which can stack until timeout or appear to hang.
 */
var bisRefreshReentryDepth_ = 0;

/**
 * First dropdown entry: selecting it clears Interest (unequip). Must be a non-empty string for Sheets.
 */
var BIS_INTEREST_CLEAR = "----";

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

/** Case-insensitive item name equality after trim (both must be non-empty). */
function itemsNameMatch_(a, b) {
  var x = normalizeItemName(a);
  var y = normalizeItemName(b);
  return x !== "" && y !== "" && x.toLowerCase() === y.toLowerCase();
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

/**
 * Simple onOpen triggers are limited to ~30s. Do not call refreshComparisonAllSpecs_ here — it sleeps
 * twice per spec and will hang or time out. Use one short wait + one full-sheet refresh; run
 * BIS Planner tools → Force refresh comparison colors if CmpRaw was still settling.
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ce = ss.getSheetByName(CURRENT_EQUIP_SHEET);
  if (ce) {
    ceEnsureGemColumnGroup_(ce);
    ceEnsureWeightsModeValidation_(ce);
    ceRefreshDerivedOnOpen_(ce);
    ceLockGemRowsWithoutItem_(ce);
  }
  var bp = ss.getSheetByName(BIS_PLANNER_SHEET);
  if (!bp) return;
  applyInterestDropdownsByGearType_(bp);
  rebuildEquippedIndexFromPlanner_(bp);
  Utilities.sleep(OPEN_RECALC_WAIT_MS);
  refreshComparisonRichText(bp, null, null, null, null);
}

/** Ring/Trinket: no plain Equipped. First item ---- clears cell on select. Matches generate_bis_planner.py. */
var INTEREST_LIST_STANDARD = [
  BIS_INTEREST_CLEAR,
  "Pass",
  "Consider",
  "Need",
  "Equipped",
];
var INTEREST_LIST_RING_TRINKET = [
  BIS_INTEREST_CLEAR,
  "Pass",
  "Consider",
  "Need",
  "Equipped 1",
  "Equipped 2",
];

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
 * Menu / repair: two full-sheet passes with a short sleep (CmpRaw may settle between reads).
 * Do not loop per-spec — that multiplies sleeps and can exceed the 6-minute limit.
 */
function refreshComparisonAllSpecs_(bp, optTimingCtx) {
  Utilities.sleep(OPEN_RECALC_WAIT_MS);
  if (optTimingCtx) bisTimingStep_(optTimingCtx, "refreshAll: after first sleep (pass 1)");
  var ss = bp.getParent();
  var ceSh = ss.getSheetByName(CURRENT_EQUIP_SHEET);
  var sharedCaches = ceSh ? plannerRefreshCachesBuild_(ss, ceSh) : null;
  if (optTimingCtx) {
    bisTimingStep_(
      optTimingCtx,
      sharedCaches
        ? "refreshAll: shared bulk cache (reused for pass 1 + 2)"
        : "refreshAll: no CE sheet — no shared cache"
    );
  }
  var fullSheetBandSnap = {};
  refreshComparisonRichText(
    bp,
    null,
    null,
    optTimingCtx,
    null,
    sharedCaches,
    { captureFullSheetBands: fullSheetBandSnap }
  );
  if (optTimingCtx) bisTimingStep_(optTimingCtx, "refreshAll: pass 1 done");
  Utilities.sleep(OPEN_SECOND_PASS_SLEEP_MS);
  if (optTimingCtx) bisTimingStep_(optTimingCtx, "refreshAll: after second sleep (pass 2)");
  refreshComparisonRichText(
    bp,
    null,
    null,
    optTimingCtx,
    null,
    sharedCaches,
    { reuseFullSheetBands: fullSheetBandSnap }
  );
  if (optTimingCtx) bisTimingStep_(optTimingCtx, "refreshAll: pass 2 done");
}

/** True if row 3 has Gear Score / Gear Score Plus Gems (regenerated workbook). */
function plannerGearScoreColumnsReady_(bpSheet) {
  try {
    var hr = BIS_FIRST_DATA_ROW - 1;
    if (hr < 1) return false;
    var q = String(bpSheet.getRange(hr, GG_GEAR_SCORE_COL).getValue() || "")
      .replace(/^\s+|\s+$/g, "")
      .toLowerCase();
    var r = String(bpSheet.getRange(hr, GG_GEAR_SCORE_PLUS_GEMS_COL).getValue() || "")
      .replace(/^\s+|\s+$/g, "")
      .toLowerCase();
    return q === "gear score" && r.indexOf("plus gems") !== -1;
  } catch (e0) {
    return false;
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
    if (normalizeItemName(newValue) === BIS_INTEREST_CLEAR || normalizeItemName(newValue) === "None") {
      e.range.setValue("");
      newValue = "";
    }
    // Legacy sentinels from earlier experiments.
    if (normalizeItemName(newValue) === "(Clear)" || String(newValue) === " ") {
      e.range.setValue("");
      newValue = "";
    }
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
      bpSheet
        .getRange(row, GG_INTEREST)
        .setDataValidation(interestValidationRuleForGear_(e.range.getValue()));
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
// Current Equipment — weights, gems, gear score (must match generate_bis_planner.py)
// ============================================================

function ceIsWeightsRow_(ceSheet, row) {
  var a = normalizeItemName(ceSheet.getRange(row, CE_GEARTYPE).getValue());
  return a === CE_WEIGHT_ROW_LABEL;
}

function ceFindWeightsRowForSection_(ceSheet, anyRow) {
  var last = ceSheet.getLastRow();
  for (var r = anyRow; r <= last; r++) {
    var a = String(ceSheet.getRange(r, CE_GEARTYPE).getValue()).trim();
    if (a.indexOf("---") === 0) return null;
    if (a === CE_WEIGHT_ROW_LABEL) return r;
  }
  return null;
}

function ceForEachGearRowInSection_(ceSheet, weightsRow, fn) {
  var r = weightsRow - 1;
  while (r >= CE_FIRST_DATA_ROW) {
    var a = String(ceSheet.getRange(r, CE_GEARTYPE).getValue()).trim();
    if (a.indexOf("---") === 0) break;
    r--;
  }
  var first = r + 1;
  for (var i = first; i < weightsRow; i++) {
    var gt = String(ceSheet.getRange(i, CE_GEARTYPE).getValue()).trim();
    if (!gt || gt.indexOf("---") === 0 || gt === CE_WEIGHT_ROW_LABEL) continue;
    fn(i);
  }
}

function ceApplyWeightMode_(ceSheet, wRow, mode) {
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  if (mode === CE_WEIGHT_MODE_PAWN) {
    var snap = ceSheet.getRange(wRow, CE_PAWN_SNAPSHOT_FIRST_COL, 1, n).getValues()[0];
    ceSheet.getRange(wRow, CE_STAT_FIRST_COL, 1, n).setValues([snap]);
  } else {
    var stash = ceSheet.getRange(wRow, CE_CUSTOM_STASH_FIRST_COL, 1, n).getValues()[0];
    ceSheet.getRange(wRow, CE_STAT_FIRST_COL, 1, n).setValues([stash]);
  }
}

function ceCopyVisibleWeightsToStash_(ceSheet, wRow) {
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var v = ceSheet.getRange(wRow, CE_STAT_FIRST_COL, 1, n).getValues()[0];
  ceSheet.getRange(wRow, CE_CUSTOM_STASH_FIRST_COL, 1, n).setValues([v]);
}

function ceHandleWeightsEdit_(e, ceSheet, row, col) {
  if (col === CE_ITEMNAME) {
    var mode = normalizeItemName(e.range.getValue());
    if (mode === CE_WEIGHT_MODE_PAWN || mode === CE_WEIGHT_MODE_CUSTOM) {
      ceApplyWeightMode_(ceSheet, row, mode);
      ceForEachGearRowInSection_(ceSheet, row, function (gr) {
        recalculateGearScoreForRow_(ceSheet, gr);
      });
    }
    return;
  }
  if (col >= CE_STAT_FIRST_COL && col <= CE_STAT_LAST_COL) {
    ceSheet.getRange(row, CE_ITEMNAME).setValue(CE_WEIGHT_MODE_CUSTOM);
    ceCopyVisibleWeightsToStash_(ceSheet, row);
    ceForEachGearRowInSection_(ceSheet, row, function (gr) {
      recalculateGearScoreForRow_(ceSheet, gr);
    });
  }
}

function itemdbLookupSockets_(ss, itemName) {
  var z = { r: 0, y: 0, b: 0, m: 0 };
  if (!itemName) return z;
  var db = ss.getSheetByName(ITEMDB_SHEET);
  if (!db) return z;
  var lr = db.getLastRow();
  if (lr < 2) return z;
  var names = db.getRange(2, 1, lr, 1).getValues();
  var socks = db.getRange(2, ITEMDB_SOCKR_COL, lr, ITEMDB_SOCKM_COL).getValues();
  var want = normalizeItemName(itemName);
  for (var i = 0; i < names.length; i++) {
    if (itemsNameMatch_(names[i][0], itemName)) {
      z.r = Number(socks[i][0]) || 0;
      z.y = Number(socks[i][1]) || 0;
      z.b = Number(socks[i][2]) || 0;
      z.m = Number(socks[i][3]) || 0;
      return z;
    }
  }
  if (BIS_TIMING_DEBUG && want) {
    Logger.log("itemdbLookupSockets_: no ItemDB row for name (trimmed): " + want);
  }
  return z;
}

/**
 * ItemDB row for an item: socket counts + stat vector (same order as Current Equipment stat columns).
 * @return {?{socks:{r:number,y:number,b:number,m:number}, stats:Array<number>}}
 */
function itemdbLookupStatsAndSockets_(ss, itemName) {
  var want = normalizeItemName(itemName);
  if (!want) return null;
  var db = ss.getSheetByName(ITEMDB_SHEET);
  if (!db) return null;
  var lr = db.getLastRow();
  if (lr < 2) return null;
  var nStat = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var width = ITEMDB_FIRST_STAT_COL - 1 + nStat;
  var data = db.getRange(2, 1, lr, width).getValues();
  for (var i = 0; i < data.length; i++) {
    if (!itemsNameMatch_(data[i][0], itemName)) continue;
    var stats = [];
    for (var s = 0; s < nStat; s++) {
      stats.push(Number(data[i][ITEMDB_FIRST_STAT_COL - 1 + s]) || 0);
    }
    return {
      socks: {
        r: Number(data[i][ITEMDB_SOCKR_COL - 1]) || 0,
        y: Number(data[i][ITEMDB_SOCKY_COL - 1]) || 0,
        b: Number(data[i][ITEMDB_SOCKB_COL - 1]) || 0,
        m: Number(data[i][ITEMDB_SOCKM_COL - 1]) || 0,
      },
      stats: stats,
    };
  }
  return null;
}

/** One ItemDB read for refresh loops (avoids re-fetching the sheet per item). */
function itemdbReadAllStatRows_(ss) {
  var db = ss.getSheetByName(ITEMDB_SHEET);
  if (!db) return null;
  var lr = db.getLastRow();
  if (lr < 2) return null;
  var nStat = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var width = ITEMDB_FIRST_STAT_COL - 1 + nStat;
  return db.getRange(2, 1, lr, width).getValues();
}

function itemdbLookupStatsAndSocketsInData_(data, itemName) {
  if (!data || !itemName) return null;
  var nStat = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  for (var i = 0; i < data.length; i++) {
    if (!itemsNameMatch_(data[i][0], itemName)) continue;
    var stats = [];
    for (var s = 0; s < nStat; s++) {
      stats.push(Number(data[i][ITEMDB_FIRST_STAT_COL - 1 + s]) || 0);
    }
    return {
      socks: {
        r: Number(data[i][ITEMDB_SOCKR_COL - 1]) || 0,
        y: Number(data[i][ITEMDB_SOCKY_COL - 1]) || 0,
        b: Number(data[i][ITEMDB_SOCKB_COL - 1]) || 0,
        m: Number(data[i][ITEMDB_SOCKM_COL - 1]) || 0,
      },
      stats: stats,
    };
  }
  return null;
}

/** One GemDB read for refresh loops (computeIdealGemPlan scans categories in memory). */
function gemdbReadAllRows_(ss) {
  var sh = ss.getSheetByName(GEMDB_SHEET);
  if (!sh) return null;
  var lr = sh.getLastRow();
  if (lr < 2) return null;
  return sh.getRange(2, 1, lr, 4).getValues();
}

/** Weights row for a planner spec name (must match CE section title case). */
function ceFindWeightsRowForSpec_(ceSheet, spec) {
  var want = normalizeItemName(spec);
  if (!want) return null;
  var lr = ceSheet.getLastRow();
  var curSpec = null;
  for (var r = 1; r <= lr; r++) {
    var a = String(ceSheet.getRange(r, CE_GEARTYPE).getValue()).trim();
    if (a.indexOf("---") === 0) {
      var parsed = specDisplayNameFromSectionHeader(a);
      curSpec = parsed ? normalizeItemName(parsed) : null;
      continue;
    }
    if (a === CE_WEIGHT_ROW_LABEL && curSpec && curSpec === want) {
      return r;
    }
  }
  return null;
}

/** True when this planner row is the equipped item (Stat Comparison / % Upgrade stay blank). */
function plannerRowIsEquippedItemRow_(interestNorm, gearRow, itemNameRaw, eqRaw, eq2Raw) {
  if (normalizeItemName(itemNameRaw) === "") return false;
  if (gearIsRingOrTrinket_(gearRow)) {
    var e1 = normalizeItemName(eqRaw);
    var e2 = normalizeItemName(eq2Raw);
    if (interestNorm === "Equipped 1" && e1 !== "" && itemsNameMatch_(itemNameRaw, eqRaw)) return true;
    if (interestNorm === "Equipped 2" && e2 !== "" && itemsNameMatch_(itemNameRaw, eq2Raw)) return true;
    if (interestNorm === "Equipped" && e1 !== "" && itemsNameMatch_(itemNameRaw, eqRaw)) return true;
    return false;
  }
  return itemsNameMatch_(itemNameRaw, eqRaw);
}

/**
 * Baseline "Gear Score Plus Gems" for % Upgrade: same pipeline as column R, using equipped names from
 * planner columns N/O (matches Equip / Equip2 formulas). Avoids findCERow + CE Gear Score column, which can
 * disagree with what N/O show when section matching or cached CE scores drift.
 * Rings/trinkets: min of scores for the two equipped names (Pawn-style weaker-ring baseline).
 */
function plannerBaselineEquippedScore_(ss, ceSheet, spec, plannerGearType, eqN, eq2N, optCaches) {
  var g = normalizeItemName(plannerGearType);
  var specN = normalizeItemName(spec);
  var e1 = normalizeItemName(eqN);
  var e2 = normalizeItemName(eq2N);
  var sk = specN + "\0" + g + "\0" + e1 + "\0" + e2;
  if (optCaches && Object.prototype.hasOwnProperty.call(optCaches.baselineByKey, sk)) {
    return optCaches.baselineByKey[sk];
  }
  var out = 0;
  if (g === "Ring" || g === "Trinket") {
    var u1 = e1 ? plannerItemGearScore_(ss, ceSheet, spec, eqN, optCaches) : 0;
    var u2 = e2 ? plannerItemGearScore_(ss, ceSheet, spec, eq2N, optCaches) : 0;
    var d = [];
    if (u1 > 0) d.push(u1);
    if (u2 > 0) d.push(u2);
    out = d.length === 0 ? 0 : Math.min.apply(null, d);
  } else {
    out = e1 ? plannerItemGearScore_(ss, ceSheet, spec, eqN, optCaches) : 0;
  }
  if (optCaches) optCaches.baselineByKey[sk] = out;
  return out;
}

/**
 * Bulk-read CE column A once + ItemDB + GemDB; lazy weight vectors per spec; memo item + baseline scores.
 * Used only inside refreshComparisonRichTextInner_ to avoid thousands of sheet reads.
 */
function plannerRefreshCachesBuild_(ss, ceSheet) {
  var itemdbData = itemdbReadAllStatRows_(ss);
  var gemData = gemdbReadAllRows_(ss);
  var celr = ceSheet.getLastRow();
  var aCol = celr >= 1 ? ceSheet.getRange(1, CE_GEARTYPE, celr, 1).getValues() : [];
  var specToRow = {};
  var curSpec = null;
  for (var i = 0; i < aCol.length; i++) {
    var r = i + 1;
    var a = String(aCol[i][0]).trim();
    if (a.indexOf("---") === 0) {
      var parsed = specDisplayNameFromSectionHeader(a);
      curSpec = parsed ? normalizeItemName(parsed) : null;
      continue;
    }
    if (a === CE_WEIGHT_ROW_LABEL && curSpec) specToRow[curSpec] = r;
  }
  var weights = {
    specToRow: specToRow,
    wBySpec: {},
    ceSheet: ceSheet,
    get: function (spec) {
      var k = normalizeItemName(spec);
      if (!k) return null;
      var row = this.specToRow[k];
      if (!row) return null;
      if (this.wBySpec[k]) return this.wBySpec[k];
      var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
      var wVals = this.ceSheet.getRange(row, CE_STAT_FIRST_COL, 1, n).getValues()[0];
      var metaSockW = Number(this.ceSheet.getRange(row, CE_META_WEIGHT_COL).getValue()) || 0;
      var o = { wRow: row, wVals: wVals, metaSockW: metaSockW };
      this.wBySpec[k] = o;
      return o;
    },
  };
  return {
    ss: ss,
    itemdbData: itemdbData,
    gemData: gemData,
    weights: weights,
    ceGearCol: aCol,
    ceSheet: ceSheet,
    baselineByKey: {},
    itemScoreBaseByKey: {},
    itemScoreFullByKey: {},
  };
}

function plannerItemGearScoreBaseOnlyWithCaches_(caches, spec, itemName) {
  var nm = normalizeItemName(itemName);
  if (!nm) return 0;
  var ck = normalizeItemName(spec) + "\0" + nm;
  if (Object.prototype.hasOwnProperty.call(caches.itemScoreBaseByKey, ck)) {
    return caches.itemScoreBaseByKey[ck];
  }
  var w = caches.weights.get(spec);
  if (!w) {
    caches.itemScoreBaseByKey[ck] = 0;
    return 0;
  }
  var row = itemdbLookupStatsAndSocketsInData_(caches.itemdbData, itemName);
  if (!row) {
    caches.itemScoreBaseByKey[ck] = 0;
    return 0;
  }
  var base = 0;
  for (var i = 0; i < w.wVals.length; i++) {
    base += (Number(w.wVals[i]) || 0) * (Number(row.stats[i]) || 0);
  }
  var out = Math.round(base * 100) / 100;
  caches.itemScoreBaseByKey[ck] = out;
  return out;
}

function plannerItemGearScoreWithCaches_(caches, spec, itemName) {
  var nm = normalizeItemName(itemName);
  if (!nm) return 0;
  var ck = normalizeItemName(spec) + "\0" + nm;
  if (Object.prototype.hasOwnProperty.call(caches.itemScoreFullByKey, ck)) {
    return caches.itemScoreFullByKey[ck];
  }
  var w = caches.weights.get(spec);
  if (!w) {
    caches.itemScoreFullByKey[ck] = 0;
    return 0;
  }
  var row = itemdbLookupStatsAndSocketsInData_(caches.itemdbData, itemName);
  if (!row) {
    caches.itemScoreFullByKey[ck] = 0;
    return 0;
  }
  var base = 0;
  for (var i = 0; i < w.wVals.length; i++) {
    base += (Number(w.wVals[i]) || 0) * (Number(row.stats[i]) || 0);
  }
  var plan = computeIdealGemPlan_(caches.ss, w.wVals, w.metaSockW, row.socks, caches.gemData);
  var out = Math.round((base + plan.gemScore) * 100) / 100;
  caches.itemScoreFullByKey[ck] = out;
  return out;
}

function scoreGemStatsAgainstWeights_(st, wVals) {
  var tot = 0;
  for (var pk in st) {
    if (!Object.prototype.hasOwnProperty.call(st, pk)) continue;
    var wgt = ceWeightForPawnStat_(wVals, pk);
    tot += (Number(st[pk]) || 0) * wgt;
  }
  return tot;
}

function bestGemInCategory_(ss, cat, wVals, optGemData) {
  var data = optGemData;
  if (data == null) {
    var sh = ss.getSheetByName(GEMDB_SHEET);
    if (!sh) return null;
    var lr = sh.getLastRow();
    if (lr < 2) return null;
    data = sh.getRange(2, 1, lr, 4).getValues();
  }
  var bestScore = -1;
  var bestLabel = "";
  var bestStats = null;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1]) !== cat) continue;
    var rawLab = data[i][2];
    var lab = ceNormalizeGemLabel_(rawLab);
    if (!lab || lab === "----") continue;
    var js = String(data[i][3] || "");
    var st;
    try {
      st = JSON.parse(js);
    } catch (e1) {
      continue;
    }
    var sc = scoreGemStatsAgainstWeights_(st, wVals);
    if (sc > bestScore) {
      bestScore = sc;
      bestLabel = String(rawLab == null ? "" : rawLab).replace(/\u200b/g, "");
      bestStats = st;
    }
  }
  if (bestScore < 0) return null;
  return { label: bestLabel, stats: bestStats, score: bestScore };
}

/**
 * Ideal gems: best weighted gem from non-meta categories fills all colored sockets; best meta per meta socket.
 * @return {{gemScore:number, regLabel:string, regStats:?Object, metaLabel:string, metaStats:?Object, T:number, m:number}}
 */
function computeIdealGemPlan_(ss, wVals, metaSockW, socks, optGemData) {
  var T = (socks.r || 0) + (socks.y || 0) + (socks.b || 0);
  var m = socks.m || 0;
  var gemScore = 0;
  var regLabel = "";
  var regStats = null;
  var bestRegScore = -1;
  if (T > 0) {
    for (var ci = 0; ci < CE_GEM_NON_META_CATS.length; ci++) {
      var best = bestGemInCategory_(ss, CE_GEM_NON_META_CATS[ci], wVals, optGemData);
      if (best && best.score > bestRegScore) {
        bestRegScore = best.score;
        regLabel = best.label;
        regStats = best.stats;
      }
    }
    if (bestRegScore >= 0) gemScore += T * bestRegScore;
  }
  var metaLabel = "";
  var metaStats = null;
  if (m > 0) {
    gemScore += m * (Number(metaSockW) || 0);
    var bm = bestGemInCategory_(ss, "meta", wVals, optGemData);
    if (bm) {
      metaLabel = bm.label;
      metaStats = bm.stats;
      gemScore += m * bm.score;
    }
  }
  return {
    gemScore: gemScore,
    regLabel: regLabel,
    regStats: regStats,
    metaLabel: metaLabel,
    metaStats: metaStats,
    T: T,
    m: m,
  };
}

var PAWN_KEY_SHORT = {
  Armor: "Armor",
  Strength: "Str",
  Agility: "Agi",
  Stamina: "Sta",
  Intellect: "Int",
  Spirit: "Spi",
  Healing: "Healing",
  SpellDamage: "Spell Dmg",
  Ap: "AP",
  Mp5: "MP5",
  DefenseRating: "Defense",
  DodgeRating: "Dodge",
  ParryRating: "Parry",
  BlockRating: "Block Rating",
  BlockValue: "Block Value",
  HitRating: "Hit",
  CritRating: "Crit",
  SpellHitRating: "Spell Hit",
  SpellCritRating: "Spell Crit",
  HasteRating: "Haste",
  SpellHasteRating: "Haste",
  ResilienceRating: "Resilience",
  Versatility: "Versatility",
};

function aggregateGemStats_(regStats, T, metaStats, m) {
  var acc = {};
  function add(obj, mult) {
    if (!obj || mult <= 0) return;
    for (var k in obj) {
      if (!Object.prototype.hasOwnProperty.call(obj, k)) continue;
      var v = (Number(obj[k]) || 0) * mult;
      if (v === 0) continue;
      acc[k] = (acc[k] || 0) + v;
    }
  }
  add(regStats, T);
  add(metaStats, m);
  return acc;
}

function formatAggregatedGemStats_(acc) {
  var keys = [];
  for (var k in acc) {
    if (Object.prototype.hasOwnProperty.call(acc, k)) keys.push(k);
  }
  if (keys.length === 0) return "";
  keys.sort();
  var parts = [];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var v = acc[k];
    if (Math.abs(v - Math.round(v)) < 1e-6) v = Math.round(v);
    else v = Math.round(v * 100) / 100;
    var short = PAWN_KEY_SHORT[k] || k;
    parts.push((v > 0 ? "+" : "") + v + " " + short);
  }
  return parts.join(", ");
}

/** ItemDB stats × weights only (no sockets / ideal gems). */
function plannerItemGearScoreBaseOnly_(ss, ceSheet, spec, itemName, optCaches) {
  if (optCaches) return plannerItemGearScoreBaseOnlyWithCaches_(optCaches, spec, itemName);
  var nm = normalizeItemName(itemName);
  if (!nm) return 0;
  var wRow = ceFindWeightsRowForSpec_(ceSheet, spec);
  if (!wRow) return 0;
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var wVals = ceSheet.getRange(wRow, CE_STAT_FIRST_COL, 1, n).getValues()[0];
  var row = itemdbLookupStatsAndSockets_(ss, nm);
  if (!row) return 0;
  var base = 0;
  for (var i = 0; i < wVals.length; i++) {
    base += (Number(wVals[i]) || 0) * (Number(row.stats[i]) || 0);
  }
  return Math.round(base * 100) / 100;
}

function plannerItemGearScore_(ss, ceSheet, spec, itemName, optCaches) {
  if (optCaches) return plannerItemGearScoreWithCaches_(optCaches, spec, itemName);
  var nm = normalizeItemName(itemName);
  if (!nm) return 0;
  var wRow = ceFindWeightsRowForSpec_(ceSheet, spec);
  if (!wRow) return 0;
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var wVals = ceSheet.getRange(wRow, CE_STAT_FIRST_COL, 1, n).getValues()[0];
  var metaSockW = Number(ceSheet.getRange(wRow, CE_META_WEIGHT_COL).getValue()) || 0;
  var row = itemdbLookupStatsAndSockets_(ss, nm);
  if (!row) return 0;
  var base = 0;
  for (var i = 0; i < wVals.length; i++) {
    base += (Number(wVals[i]) || 0) * (Number(row.stats[i]) || 0);
  }
  var plan = computeIdealGemPlan_(ss, wVals, metaSockW, row.socks);
  return Math.round((base + plan.gemScore) * 100) / 100;
}

/**
 * % Upgrade from "Gear Score Plus Gems" vs Current Equipment baseline for that slot.
 * @param {number} itemScorePlusGems — same value written to column R
 */
function upgradePctFromScores_(itemScorePlusGems, equippedBaseline, heroicDark, hideRow) {
  var out = { text: "", color: "#000000" };
  if (hideRow) return out;
  if (equippedBaseline <= 0) return out;
  var pct = ((itemScorePlusGems - equippedBaseline) / equippedBaseline) * 100;
  var rounded = Math.round(pct * 10) / 10;
  var sign = rounded > 0 ? "+" : "";
  out.text = sign + rounded + "%";
  var posC = heroicDark ? CMP_DELTA_POS_DARK_ROW : "#0d652d";
  var negC = heroicDark ? CMP_DELTA_NEG_DARK_ROW : "#c5221f";
  out.color =
    rounded > 0 ? posC : rounded < 0 ? negC : heroicDark ? "#EEEEEE" : "#000000";
  return out;
}

function flushUpgradePctCells_(bpSheet, upgradeWrites, timingCtx) {
  if (!upgradeWrites || upgradeWrites.length === 0) {
    bisTimingStep_(timingCtx, "flushUpgradePctCells_: 0 writes (skip)");
    return;
  }
  bisTimingStep_(timingCtx, "flushUpgradePctCells_: start");
  upgradeWrites.sort(function (a, b) {
    return a.r - b.r;
  });
  var segments = [];
  var cur = [upgradeWrites[0]];
  for (var wi = 1; wi < upgradeWrites.length; wi++) {
    if (upgradeWrites[wi].r === upgradeWrites[wi - 1].r + 1) cur.push(upgradeWrites[wi]);
    else {
      segments.push(cur);
      cur = [upgradeWrites[wi]];
    }
  }
  segments.push(cur);
  for (var si = 0; si < segments.length; si++) {
    var seg = segments[si];
    var r0 = seg[0].r;
    var h = seg.length;
    var texts = [];
    var colors = [];
    for (var j = 0; j < h; j++) {
      texts.push([seg[j].text]);
      colors.push([seg[j].color]);
    }
    var upRange = bpSheet.getRange(r0, GG_UPGRADE_COL, h, 1);
    upRange.setValues(texts);
    try {
      upRange.setFontColors(colors);
    } catch (eFont) {}
  }
  bisTimingStep_(timingCtx, "flushUpgradePctCells_: done");
}

function flushPlannerGearScoreCells_(bpSheet, gearWrites, timingCtx) {
  if (!gearWrites || gearWrites.length === 0) {
    bisTimingStep_(timingCtx, "flushPlannerGearScoreCells_: 0 writes (skip)");
    return;
  }
  bisTimingStep_(timingCtx, "flushPlannerGearScoreCells_: start");
  gearWrites.sort(function (a, b) {
    return a.r - b.r;
  });
  var segments = [];
  var cur = [gearWrites[0]];
  for (var wi = 1; wi < gearWrites.length; wi++) {
    if (gearWrites[wi].r === gearWrites[wi - 1].r + 1) cur.push(gearWrites[wi]);
    else {
      segments.push(cur);
      cur = [gearWrites[wi]];
    }
  }
  segments.push(cur);
  for (var si = 0; si < segments.length; si++) {
    var seg = segments[si];
    var r0 = seg[0].r;
    var h = seg.length;
    var matrix = [];
    for (var j = 0; j < h; j++) {
      matrix.push([seg[j].base, seg[j].full]);
    }
    try {
      bpSheet.getRange(r0, GG_GEAR_SCORE_COL, h, 2).setValues(matrix);
    } catch (eGs) {}
  }
  bisTimingStep_(timingCtx, "flushPlannerGearScoreCells_: done");
}

/** Updates ideal gem display, backgrounds, and Gear Score for one CE row. */
function ceUpdateIdealGemCellsAndScore_(ceSheet, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wRow = ceFindWeightsRowForSection_(ceSheet, row);
  if (!wRow) return;
  var item = normalizeItemName(ceSheet.getRange(row, CE_ITEMNAME).getValue());
  if (!item) {
    ceSheet.getRange(row, CE_IDEAL_GEM_COL, 1, CE_TOTAL_GEM_STATS_COL - CE_IDEAL_GEM_COL + 1).setValues([["", "", ""]]);
    ceSheet.getRange(row, CE_IDEAL_GEM_COL, 1, CE_TOTAL_GEM_STATS_COL - CE_IDEAL_GEM_COL + 1).setBackground(CE_GEM_LOCKED_BG);
    ceSheet.getRange(row, CE_GEAR_SCORE_COL).setValue("");
    return;
  }
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var wVals = ceSheet.getRange(wRow, CE_STAT_FIRST_COL, 1, n).getValues()[0];
  var metaSockW = Number(ceSheet.getRange(wRow, CE_META_WEIGHT_COL).getValue()) || 0;
  var socks = itemdbLookupSockets_(ss, item);
  var T = socks.r + socks.y + socks.b;
  var m = socks.m || 0;
  var plan = computeIdealGemPlan_(ss, wVals, metaSockW, socks);
  var acc = aggregateGemStats_(plan.regStats, plan.T, plan.metaStats, plan.m);
  ceSheet.getRange(row, CE_IDEAL_GEM_COL).setValue(T > 0 ? plan.regLabel || "—" : "");
  ceSheet.getRange(row, CE_IDEAL_META_COL).setValue(m > 0 ? plan.metaLabel || "—" : "");
  ceSheet.getRange(row, CE_TOTAL_GEM_STATS_COL).setValue(formatAggregatedGemStats_(acc));
  ceSheet.getRange(row, CE_IDEAL_GEM_COL).setBackground(T > 0 ? null : CE_GEM_LOCKED_BG);
  ceSheet.getRange(row, CE_IDEAL_META_COL).setBackground(m > 0 ? null : CE_GEM_LOCKED_BG);
  ceSheet.getRange(row, CE_TOTAL_GEM_STATS_COL).setBackground(T > 0 || m > 0 ? null : CE_GEM_LOCKED_BG);
  ceSheet.getRange(row, CE_IDEAL_GEM_COL).setFontColor("#000000");
  ceSheet.getRange(row, CE_IDEAL_META_COL).setFontColor("#000000");
  ceSheet.getRange(row, CE_TOTAL_GEM_STATS_COL).setFontColor("#000000");
  var sVals = ceSheet.getRange(row, CE_STAT_FIRST_COL, 1, n).getValues()[0];
  var score = 0;
  for (var ii = 0; ii < wVals.length; ii++) {
    score += (Number(sVals[ii]) || 0) * (Number(wVals[ii]) || 0);
  }
  score += plan.gemScore;
  ceSheet.getRange(row, CE_GEAR_SCORE_COL).setValue(Math.round(score * 100) / 100);
}

/** Strip ZWSP prefix used so Sheets does not parse "+8 Str" as a formula. */
function ceNormalizeGemLabel_(v) {
  return String(v == null ? "" : v)
    .replace(/\u200b/g, "")
    .replace(/^\s+|\s+$/g, "");
}

function ceWeightForPawnStat_(wVals, pawnKey) {
  for (var i = 0; i < CE_PAWN_KEYS_FOR_STAT.length; i++) {
    var ks = CE_PAWN_KEYS_FOR_STAT[i];
    for (var j = 0; j < ks.length; j++) {
      if (ks[j] === pawnKey) return Number(wVals[i]) || 0;
    }
  }
  return 0;
}

function recalculateGearScoreForRow_(ceSheet, row) {
  ceUpdateIdealGemCellsAndScore_(ceSheet, row);
}

/** Full refresh (all gear rows). Prefer ceRefreshDerivedOnOpen_ from onOpen. */
function ceRefreshAllDerived_(ceSheet) {
  var lr = ceSheet.getLastRow();
  for (var r = CE_FIRST_DATA_ROW; r <= lr; r++) {
    if (ceIsWeightsRow_(ceSheet, r)) continue;
    var gt = String(ceSheet.getRange(r, CE_GEARTYPE).getValue()).trim();
    if (!gt || gt.indexOf("---") === 0) continue;
    recalculateGearScoreForRow_(ceSheet, r);
  }
}

/** Gear rows with no item in B: clear ideal gem cells (grey) and gear score. */
function ceLockGemRowsWithoutItem_(ceSheet) {
  var lr = ceSheet.getLastRow();
  for (var r = CE_FIRST_DATA_ROW; r <= lr; r++) {
    if (ceIsWeightsRow_(ceSheet, r)) continue;
    var gt = String(ceSheet.getRange(r, CE_GEARTYPE).getValue()).trim();
    if (!gt || gt.indexOf("---") === 0) continue;
    if (normalizeItemName(ceSheet.getRange(r, CE_ITEMNAME).getValue())) continue;
    recalculateGearScoreForRow_(ceSheet, r);
  }
}

/** onOpen: only rows with an item name — avoids O(rows×cols) validation work on empty slots. */
function ceRefreshDerivedOnOpen_(ceSheet) {
  var lr = ceSheet.getLastRow();
  for (var r = CE_FIRST_DATA_ROW; r <= lr; r++) {
    if (ceIsWeightsRow_(ceSheet, r)) continue;
    var gt = String(ceSheet.getRange(r, CE_GEARTYPE).getValue()).trim();
    if (!gt || gt.indexOf("---") === 0) continue;
    var item = normalizeItemName(ceSheet.getRange(r, CE_ITEMNAME).getValue());
    if (!item) continue;
    recalculateGearScoreForRow_(ceSheet, r);
  }
}

/**
 * One column group for ideal gem block (C–E). If grouping looks wrong after an xlsx change, delete document
 * property BIS_CE_GEM_GROUPED_V1 once and reopen so this runs again.
 */
function ceEnsureGemColumnGroup_(ceSheet) {
  try {
    var p = PropertiesService.getDocumentProperties().getProperty("BIS_CE_GEM_GROUPED_V1");
    if (p) return;
    ceSheet.getRange(1, CE_GEM_FIRST_COL, ceSheet.getMaxRows(), CE_GEM_LAST_COL).shiftColumnGroupDepth(1);
    PropertiesService.getDocumentProperties().setProperty("BIS_CE_GEM_GROUPED_V1", "1");
  } catch (eGrp) {}
}

function ceEnsureWeightsModeValidation_(ceSheet) {
  var lr = ceSheet.getLastRow();
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CE_WEIGHT_MODE_PAWN, CE_WEIGHT_MODE_CUSTOM], true)
    .setAllowInvalid(true)
    .build();
  for (var r = CE_FIRST_DATA_ROW; r <= lr; r++) {
    if (ceIsWeightsRow_(ceSheet, r)) {
      ceSheet.getRange(r, CE_ITEMNAME).setDataValidation(rule);
    }
  }
}

function handleCurrentEquipEdit(e, ceSheet) {
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row < CE_FIRST_DATA_ROW) return;

  var gearType = normalizeItemName(ceSheet.getRange(row, CE_GEARTYPE).getValue());
  if (!gearType || gearType.indexOf("---") === 0) return;

  if (gearType === CE_WEIGHT_ROW_LABEL) {
    ceHandleWeightsEdit_(e, ceSheet, row, col);
    return;
  }

  if (col >= CE_GEM_FIRST_COL && col <= CE_GEM_LAST_COL) {
    recalculateGearScoreForRow_(ceSheet, row);
    return;
  }

  if (col >= CE_STAT_FIRST_COL && col <= CE_STAT_LAST_COL) {
    recalculateGearScoreForRow_(ceSheet, row);
    return;
  }

  if (col !== CE_ITEMNAME) return;

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

  recalculateGearScoreForRow_(ceSheet, row);
  recalculateGearScoreForRow_(ceSheet, row);

  SpreadsheetApp.flush();
  bisTimingStep_(tCtx, "after flush");
  Utilities.sleep(EDIT_RECALC_WAIT_MS);
  bisTimingStep_(tCtx, "after sleep");
  refreshComparisonRichText(bpSheet, spec, null, tCtx, null);
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
 * Section matching must mirror ceFindWeightsRowForSpec_: substring checks (e.g. "Death Knight" inside
 * "--- DEATH KNIGHT TANK ---") pick the wrong spec and corrupt % Upgrade baselines.
 */
function findCERow(ceSheet, spec, gearType, gearColValuesOpt) {
  var data = gearColValuesOpt;
  if (data == null) {
    var lastRow = ceSheet.getLastRow();
    if (lastRow < 1) return null;
    data = ceSheet.getRange(1, CE_GEARTYPE, lastRow, 1).getValues();
  }
  var specWant = normalizeItemName(spec);
  var gearWant = normalizeItemName(gearType);
  if (!gearWant) return null;

  var curSpec = null;
  for (var i = 0; i < data.length; i++) {
    var val = String(data[i][0]);
    if (val.indexOf("---") === 0) {
      var parsedHdr = specDisplayNameFromSectionHeader(val);
      curSpec = parsedHdr ? normalizeItemName(parsedHdr) : null;
      continue;
    }
    if (curSpec && specWant && curSpec === specWant && normalizeItemName(val) === gearWant) {
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
    if (itemsNameMatch_(data[i][GG_NAME - 1], want)) {
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
    if (!itemsNameMatch_(data[i][GG_NAME - 1], want)) continue;
    if (normalizeItemName(data[i][GG_INTEREST - 1]) !== normalizeItemName(wantInterest)) continue;
    bpSheet.getRange(r, GG_INTEREST).setValue("");
    bpSheet.getRange(r, GG_INTEREST).setFontWeight("normal");
  }
  syncEquippedIndexForSpecGearFromPlannerData_(spec, pGear, data);
}

// ============================================================
// Stat Comparison column: Rich Text (Google Sheets only; CmpRaw = P, Stat Comparison = K)
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
 * Writes queued Stat Comparison (K) updates in batches (clear + setRichTextValues / setFormulas) to cut round-trips.
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
 * Full-sheet refresh: same logical layout as B:P indexed from Spec (B), without reading E,G,J–M.
 * partAD: A–D (Interest, Spec, Gear type, Name); partFI: F–I; partNOP: N–P (Equip, Equip2, CmpRaw).
 */
function plannerAssembleFullSheetBlockDisplayFromBands_(partAD, partFI, partNOP) {
  var n = partAD.length;
  var blockDisplay = new Array(n);
  for (var ri = 0; ri < n; ri++) {
    var ad = partAD[ri];
    var fi = partFI[ri];
    var nop = partNOP[ri];
    blockDisplay[ri] = [
      ad[1],
      ad[2],
      ad[3],
      null,
      fi[0],
      fi[1],
      fi[2],
      fi[3],
      null,
      null,
      null,
      null,
      nop[0],
      nop[1],
      nop[2],
    ];
  }
  return blockDisplay;
}

/** Stable key for force-refresh pass 2: skip CPU when CmpRaw-derived inputs match pass 1. */
function plannerCmpRawRefreshRowFingerprint_(dispStr, heroicDarkRow, intN, equippedHide) {
  var errCell = dispStr.length > 0 && dispStr.charAt(0) === "#" ? "1" : "0";
  return (
    normalizeComparisonDisplay(dispStr) +
    "\x1f" +
    (heroicDarkRow ? "1" : "0") +
    "\x1f" +
    String(intN || "") +
    "\x1f" +
    (equippedHide ? "1" : "0") +
    "\x1f" +
    errCell
  );
}

/**
 * Applies green/red Rich Text to Stat Comparison (K) from CmpRaw (P); fills % Upgrade (L).
 * Mirror fallback uses batched =N{row} (same link as R1C1 =RC[offset]).
 * @param {?string} specFilter - If set, only rows whose Spec (B) matches.
 * @param {?string} gearTypeFilter - With specFilter, only rows whose Gear type (C) matches (same slot
 *   as a CE or Interest edit). Omit or null to refresh every row in the spec (e.g. onOpen).
 * @param {?{start:number, last:number, rows:Array<string>}} timingCtx - optional; pass null from onOpen / menu.
 * @param {?Array<Array<*>>} plannerABCReuse - optional B:C (2 cols) or A:C (3 cols) getValues snapshot, rows aligned
 *   to BIS_FIRST_DATA_ROW; when spec+gear scoped, skips a second B+C read (equip path passes B:C only).
 * @param {?Object} optSharedPlannerCaches - from plannerRefreshCachesBuild_; when set (e.g. force refresh), skip rebuilding
 *   ItemDB/GemDB/CE caches so pass 2 reuses pass 1 memoization.
 * @param {?{captureFullSheetBands:?Object, reuseFullSheetBands:?Object}} optFullSheetIo_ - force-refresh pass 2 only:
 *   pass 1 uses captureFullSheetBands (plain object filled with partAD/partFI/partNOP); pass 2 uses reuseFullSheetBands
 *   to re-read CmpRaw (P) and K formulas only.
 */
function refreshComparisonRichText(
  bpSheet,
  specFilter,
  gearTypeFilter,
  timingCtx,
  plannerABCReuse,
  optSharedPlannerCaches,
  optFullSheetIo_
) {
  var lastRow = bpSheet.getLastRow();
  if (lastRow < BIS_FIRST_DATA_ROW) return;

  if (bisRefreshReentryDepth_ > 0) {
    bisTimingStep_(timingCtx, "refreshComparisonRichText: skipped (re-entry)");
    return;
  }
  bisRefreshReentryDepth_++;
  try {
    refreshComparisonRichTextInner_(
      bpSheet,
      specFilter,
      gearTypeFilter,
      timingCtx,
      plannerABCReuse,
      lastRow,
      optSharedPlannerCaches,
      optFullSheetIo_
    );
  } finally {
    bisRefreshReentryDepth_--;
  }
}

function refreshComparisonRichTextInner_(
  bpSheet,
  specFilter,
  gearTypeFilter,
  timingCtx,
  plannerABCReuse,
  lastRow,
  optSharedPlannerCaches,
  optFullSheetIo_
) {
  var specFilterNorm =
    specFilter != null && String(specFilter).replace(/^\s+|\s+$/g, "") !== ""
      ? normalizeItemName(specFilter)
      : null;
  var gearFilterNorm =
    gearTypeFilter != null && String(gearTypeFilter).replace(/^\s+|\s+$/g, "") !== ""
      ? normalizeItemName(gearTypeFilter)
      : null;

  bisTimingStep_(timingCtx, "refreshComparisonRichText: enter");

  // Sheet.getRange(row, col, numRows, numColumns) — 3rd/4th are counts, not end row/col.
  var cmpNumRows = Math.max(0, lastRow - BIS_FIRST_DATA_ROW + 1);
  var interestAll = [];
  var cmpRawLetter = colLetterFromNum_(CMP_RAW_COL);
  var offNInBlock = CMP_RAW_COL - GG_SPEC;
  var indices = [];
  var blockDisplay = null;
  var kFormulas = null;
  var rowNK = null;
  /** F:I per planner data row — Acquisition, Quest, Dungeon, Difficulty (only when block B:P not loaded). */
  var metaAcqDungeonDiff = null;

  if (!specFilterNorm && !gearFilterNorm) {
    var rData = BIS_FIRST_DATA_ROW;
    var wAD = GG_NAME - GG_INTEREST + 1;
    var wFI = GG_DIFFICULTY - GG_ACQ + 1;
    var wNOP = CMP_RAW_COL - GG_EQUIP + 1;
    var bandsReusedForP = false;
    if (cmpNumRows > 0) {
      var partAD;
      var partFI;
      var partNOP;
      var snap =
        optFullSheetIo_ && optFullSheetIo_.reuseFullSheetBands
          ? optFullSheetIo_.reuseFullSheetBands
          : null;
      if (
        snap &&
        snap.partAD &&
        snap.partAD.length === cmpNumRows &&
        snap.partFI &&
        snap.partFI.length === cmpNumRows &&
        snap.partNOP &&
        snap.partNOP.length === cmpNumRows
      ) {
        partAD = snap.partAD;
        partFI = snap.partFI;
        var prevNOP = snap.partNOP;
        var partPonly = bpSheet.getRange(rData, CMP_RAW_COL, cmpNumRows, 1).getValues();
        partNOP = new Array(cmpNumRows);
        for (var pi = 0; pi < cmpNumRows; pi++) {
          partNOP[pi] = [prevNOP[pi][0], prevNOP[pi][1], partPonly[pi][0]];
        }
        bandsReusedForP = true;
      } else {
        partAD = bpSheet.getRange(rData, GG_INTEREST, cmpNumRows, wAD).getValues();
        partFI = bpSheet.getRange(rData, GG_ACQ, cmpNumRows, wFI).getValues();
        partNOP = bpSheet.getRange(rData, GG_EQUIP, cmpNumRows, wNOP).getValues();
      }
      blockDisplay = plannerAssembleFullSheetBlockDisplayFromBands_(partAD, partFI, partNOP);
      interestAll = new Array(cmpNumRows);
      for (var ai = 0; ai < cmpNumRows; ai++) {
        interestAll[ai] = [partAD[ai][0]];
      }
      kFormulas = bpSheet.getRange(rData, CMP_DISP_COL, cmpNumRows, 1).getFormulas();
      if (
        optFullSheetIo_ &&
        optFullSheetIo_.captureFullSheetBands &&
        !bandsReusedForP
      ) {
        var cap = optFullSheetIo_.captureFullSheetBands;
        cap.partAD = partAD;
        cap.partFI = partFI;
        cap.partNOP = partNOP;
      }
    }
    bisTimingStep_(
      timingCtx,
      bandsReusedForP
        ? "refresh: after reads (CmpRaw P + K only; bands from pass 1)"
        : "refresh: after reads (A-D,F-I,N-P values + K formulas)"
    );
    for (var aj = 0; aj < cmpNumRows; aj++) indices.push(aj);
  } else if (specFilterNorm && gearFilterNorm) {
    if (cmpNumRows > 0) {
      interestAll = bpSheet
        .getRange(BIS_FIRST_DATA_ROW, GG_INTEREST, cmpNumRows, 1)
        .getDisplayValues();
    }
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
      var specG = bpSheet.getRange(r0g, GG_SPEC, hg, 1).getDisplayValues();
      for (var jg = 0; jg < hg; jg++) {
        rowNK[sg + jg] = {
          d: dG[jg][0],
          k: kG[jg][0],
          n: nameG[jg][0],
          e: equipG[jg][0],
          e2: equip2G[jg][0],
          g: gearG[jg][0],
          sp: specG[jg][0],
        };
      }
    }
    bisTimingStep_(timingCtx, "refresh: after scoped N+K reads (" + runsG.length + " run(s))");
  } else {
    if (cmpNumRows > 0) {
      interestAll = bpSheet
        .getRange(BIS_FIRST_DATA_ROW, GG_INTEREST, cmpNumRows, 1)
        .getDisplayValues();
    }
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
      var ix0 = runsS[si2][0];
      var ps = runsS[si2][1];
      var hs = ps - ix0 + 1;
      var r0s = BIS_FIRST_DATA_ROW + ix0;
      var dS = bpSheet.getRange(r0s, CMP_RAW_COL, hs, 1).getDisplayValues();
      var kS = bpSheet.getRange(r0s, CMP_DISP_COL, hs, 1).getFormulas();
      var nameS = bpSheet.getRange(r0s, GG_NAME, hs, 1).getDisplayValues();
      var equipS = bpSheet.getRange(r0s, GG_EQUIP, hs, 1).getDisplayValues();
      var equip2S = bpSheet.getRange(r0s, GG_EQUIP2, hs, 1).getDisplayValues();
      var gearS = bpSheet.getRange(r0s, GG_GEARTYPE, hs, 1).getDisplayValues();
      var specS = bpSheet.getRange(r0s, GG_SPEC, hs, 1).getDisplayValues();
      for (var js = 0; js < hs; js++) {
        rowNK[ix0 + js] = {
          d: dS[js][0],
          k: kS[js][0],
          n: nameS[js][0],
          e: equipS[js][0],
          e2: equip2S[js][0],
          g: gearS[js][0],
          sp: specS[js][0],
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

  var gearColsOk = plannerGearScoreColumnsReady_(bpSheet);
  if (!gearColsOk) {
    bisTimingStep_(
      timingCtx,
      "refresh: Q/R not Gear Score headers — skip writing Q/R (regenerate workbook from generate_bis_planner.py)"
    );
  }

  var writes = [];
  var upgradeWrites = [];
  var gearScoreWrites = [];
  var ss = bpSheet.getParent();
  var ceSheet = ss.getSheetByName(CURRENT_EQUIP_SHEET);
  var plannerCaches = null;
  if (ceSheet) {
    if (optSharedPlannerCaches) {
      plannerCaches = optSharedPlannerCaches;
      bisTimingStep_(timingCtx, "refresh: bulk cache (reused)");
    } else {
      plannerCaches = plannerRefreshCachesBuild_(ss, ceSheet);
      bisTimingStep_(timingCtx, "refresh: bulk cache (ItemDB + GemDB + CE column A)");
    }
  }

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
    var specRow;
    if (blockDisplay != null) {
      specRow = blockDisplay[i][0];
      nmN = blockDisplay[i][nameIx];
      gearRow = blockDisplay[i][gearIx];
      eqN = blockDisplay[i][equipIx];
      eq2N = blockDisplay[i][equip2Ix];
    } else {
      specRow = rowNK[i].sp;
      nmN = rowNK[i].n;
      gearRow = rowNK[i].g;
      eqN = rowNK[i].e;
      eq2N = rowNK[i].e2;
    }
    var intN =
      interestAll.length > i && interestAll[i]
        ? normalizeItemName(interestAll[i][0])
        : "";
    if (normalizeItemName(nmN) !== "") {
      if (gearIsRingOrTrinket_(gearRow)) {
        var e1 = normalizeItemName(eqN);
        var e2 = normalizeItemName(eq2N);
        if (intN === "Equipped 1" && e1 !== "" && itemsNameMatch_(nmN, eqN)) {
          dispStr = "";
        } else if (intN === "Equipped 2" && e2 !== "" && itemsNameMatch_(nmN, eq2N)) {
          dispStr = "";
        } else if (intN === "Equipped" && e1 !== "" && itemsNameMatch_(nmN, eqN)) {
          dispStr = "";
        }
      } else if (itemsNameMatch_(nmN, eqN)) {
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

    var equippedHide = plannerRowIsEquippedItemRow_(intN, gearRow, nmN, eqN, eq2N);
    var capBandSnap = optFullSheetIo_ && optFullSheetIo_.captureFullSheetBands;
    var reuseBandSnap = optFullSheetIo_ && optFullSheetIo_.reuseFullSheetBands;
    var fpKey = null;
    if (capBandSnap || (reuseBandSnap && reuseBandSnap.cmpSkipFp)) {
      fpKey = plannerCmpRawRefreshRowFingerprint_(dispStr, heroicDarkRow, intN, equippedHide);
      if (
        reuseBandSnap &&
        reuseBandSnap.cmpSkipFp &&
        Object.prototype.hasOwnProperty.call(reuseBandSnap.cmpSkipFp, r) &&
        reuseBandSnap.cmpSkipFp[r] === fpKey
      ) {
        continue;
      }
    }

    var itemBaseNum = "";
    var itemFullNum = "";
    if (ceSheet && normalizeItemName(specRow) && normalizeItemName(nmN)) {
      itemBaseNum = plannerItemGearScoreBaseOnly_(ss, ceSheet, specRow, nmN, plannerCaches);
      itemFullNum = plannerItemGearScore_(ss, ceSheet, specRow, nmN, plannerCaches);
    }
    gearScoreWrites.push({ r: r, base: itemBaseNum, full: itemFullNum });
    var baselineEq = ceSheet
      ? plannerBaselineEquippedScore_(ss, ceSheet, specRow, gearRow, eqN, eq2N, plannerCaches)
      : 0;

    // Do not skip when N display is empty: leaving K unchanged preserves stale rich text (e.g. after
    // equipping). Empty CmpRaw → build returns null → mirror =RC[] so K shows blank until N fills, then
    // the sheet updates the mirror live; onOpen / second pass still reapplies colored rich text.
    // CmpRaw display errors (#VALUE!, #N/A, etc.) — mirror CmpRaw, no rich text
    if (dispStr.length > 0 && dispStr.charAt(0) === "#") {
      if (capBandSnap && fpKey != null) {
        if (!capBandSnap.cmpSkipFp) capBandSnap.cmpSkipFp = {};
        capBandSnap.cmpSkipFp[r] = fpKey;
      }
      if (!cellIsCmpRawMirrorFormula_(kF, r)) {
        writes.push({ r: r, kind: "mirror", needClearBefore: true });
      }
      upgradeWrites.push({
        r: r,
        text: "",
        color: "#000000",
      });
      continue;
    }
    try {
      var rich = buildComparisonRichTextValue(dispStr, heroicDarkRow);
      if (rich) {
        writes.push({
          r: r,
          kind: "rich",
          rich: rich,
          needClearBefore: true,
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
    var hidePct =
      equippedHide ||
      normalizeItemName(nmN) === "" ||
      !normalizeItemName(specRow) ||
      !ceSheet;
    var uh = upgradePctFromScores_(
      Number(itemFullNum) || 0,
      baselineEq,
      heroicDarkRow,
      hidePct
    );
    upgradeWrites.push({ r: r, text: uh.text, color: uh.color });
    if (capBandSnap && fpKey != null) {
      if (!capBandSnap.cmpSkipFp) capBandSnap.cmpSkipFp = {};
      capBandSnap.cmpSkipFp[r] = fpKey;
    }
  }

  bisTimingStep_(timingCtx, "refresh: after CPU (filter + buildComparison queue)");
  flushComparisonWrites_(bpSheet, writes, cmpRawLetter, timingCtx);
  flushPlannerGearScoreCells_(bpSheet, gearColsOk ? gearScoreWrites : [], timingCtx);
  flushUpgradePctCells_(bpSheet, upgradeWrites, timingCtx);
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
  // Do not treat "Equipped slot 2" (no colon in sheet formula) as a header — it lives inside the
  // Slot 2 body; mis-detecting it split the block and styled slot 2 unlike slot 1.
  return (
    /^Slot 1 Comparison:$/.test(trimmed) || /^Slot 2 comparison$/.test(trimmed)
  );
}

/**
 * Ring/Trinket CmpRaw: blocks after "Slot 1 Comparison:" / "Slot 2 comparison".
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
  // No +/- stat runs — muted notice except "Current Equipment Not Specified" (always black).
  if (runs.length === 0 && line1End > 0) {
    var n1s = po;
    var n1e = Math.min(po + line1End, fullLen);
    if (n1s < n1e) {
      var plainL1 = String(line1 == null ? "" : line1).replace(/^\s+|\s+$/g, "");
      var noticeFill =
        plainL1 === CMP_RAW_EQUIP_NOT_SPECIFIED ? "#000000" : colNotice;
      builder.setTextStyle(
        n1s,
        n1e,
        SpreadsheetApp.newTextStyle().setForegroundColor(noticeFill).build()
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
    .addItem("Force refresh Stat Comparison & % Upgrade", "forceRefreshComparisonColors")
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
        "No timing saved yet. Set BIS_TIMING_DEBUG = true, then run Force refresh or edit Interest / Current Equipment."
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

/** Same idea as onOpen: sleeps + two full-sheet comparison passes (avoids per-spec timeout). */
function forceRefreshComparisonColors() {
  var bp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIS_PLANNER_SHEET);
  if (!bp) {
    try {
      SpreadsheetApp.getUi().alert('No sheet named "' + BIS_PLANNER_SHEET + '".');
    } catch (e0) {}
    return;
  }
  var tCtx = bisTimingCreate_();
  if (tCtx) bisTimingStep_(tCtx, "force: menu start");
  rebuildEquippedIndexFromPlanner_(bp);
  if (tCtx) bisTimingStep_(tCtx, "force: after rebuildEquippedIndexFromPlanner_");
  refreshComparisonAllSpecs_(bp, tCtx);
  if (tCtx) bisTimingFinish_(tCtx, "Force refresh Stat Comparison");
  try {
    SpreadsheetApp.getUi().alert("Stat Comparison and % Upgrade refresh finished.");
  } catch (e1) {}
}
