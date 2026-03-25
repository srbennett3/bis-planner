// ============================================================
// BIS Planner — Apps Script
//
// Copy this entire file to: Extensions > Apps Script
// Then save (Ctrl+S). The script runs automatically on cell edits.
//
// Created by Steven Bennett 2026
// ============================================================

var BIS_PLANNER_SHEET = "BIS Planner";
var CURRENT_EQUIP_SHEET = "Current Equipment";

// First data row on each sheet (must match generate_bis_planner.py: header row 1, data from row 2)
var BIS_FIRST_DATA_ROW = 2;
var CE_FIRST_DATA_ROW = 2;

/**
 * getLastRow() reflects any column on the sheet. A single stray value far below Current Equipment
 * makes bulk reads huge (timeouts / silent failure → gear scores never update). Cap how much of CE we scan.
 */
var CE_CONTENT_ROW_CAP = 3000;
/** Rows below a gear slot to search for the section "Weights" row (generator blocks are ~25 rows). */
var CE_WEIGHTS_SCAN_BELOW_SLOT = 150;

// BIS Planner columns (1-indexed)
var GG_INTEREST = 1;  // A
var GG_SPEC     = 2;  // B
var GG_GEARTYPE = 3;  // C
var GG_NAME     = 4;  // D
var GG_ACQ      = 6;  // F — Acquisition Type (heroic dungeon row detection; match _get_row_color)
var GG_DUNGEON  = 8;  // H
var GG_DIFFICULTY = 9; // I
/** Logical columns for in-memory block (old N/O/P); not exported on BIS Planner sheet after layout v2. */
var GG_EQUIP    = 14;
var GG_EQUIP2   = 15;
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
var CE_STAT_LAST_COL = 26;
var CE_GEAR_SCORE_COL = 27;
var CE_PAWN_SNAPSHOT_FIRST_COL = 28;
var CE_CUSTOM_STASH_FIRST_COL = 49;
var CE_META_WEIGHT_COL = 70;
var CE_WEIGHT_ROW_LABEL = "Weights";
var CE_WEIGHT_MODE_PAWN = "Pawn Default";
var CE_WEIGHT_MODE_CUSTOM = "Custom";

/** ItemDB VLOOKUP: first STAT_COLUMNS column index (Armor). */
var ITEMDB_FIRST_STAT_COL = 7;
/** 1-based column index of Attributes (after all stat columns); must match generate_bis_planner.py ItemDB layout. */
var ITEMDB_ATTR_COL_1BASED = ITEMDB_FIRST_STAT_COL + CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
var ITEMDB_SOCKR_COL = 3;
var ITEMDB_SOCKY_COL = 4;
var ITEMDB_SOCKB_COL = 5;
var ITEMDB_SOCKM_COL = 6;

var GEMDB_SHEET = "GemDB";
var ITEMDB_SHEET = "ItemDB";

/** Pawn stat keys aggregated into each planner stat column (same order as STAT_COLUMNS in Python). */
var CE_PAWN_KEYS_FOR_STAT = [
  ["Armor"],
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

/**
 * Stat labels in CmpRaw diff lines (+/- …); same order as generate_bis_planner.py STAT_COLUMNS /
 * CE_PAWN_KEYS_FOR_STAT (21 entries).
 */
var BIS_STAT_LABELS = [
  "Armor",
  "Str",
  "Agi",
  "Sta",
  "Int",
  "Spi",
  "Healing",
  "Spell Dmg",
  "AP",
  "MP5",
  "Defense",
  "Dodge",
  "Parry",
  "Block Rating",
  "Block Value",
  "Hit",
  "Crit",
  "Spell Hit",
  "Spell Crit",
  "Haste",
  "Resilience",
];

/** Non-meta GemDB categories: best weighted gem fills all colored sockets (Pawn-style). */
var CE_GEM_NON_META_CATS = ["red", "yellow", "blue", "orange", "purple", "green"];

var CE_GEM_LOCKED_BG = "#E0E0E0";

/** Manual (non–ItemDB) rows: colored socket count dropdown (column C). */
var CE_MANUAL_REG_SOCKET_OPTIONS = [
  "0 Sockets",
  "1 Socket",
  "2 Sockets",
  "3 Sockets",
  "4 Sockets",
];
/** Manual rows: meta socket count (column D). */
var CE_MANUAL_META_SOCKET_OPTIONS = ["0 Sockets", "1 Socket"];

var CE_STAT_PROTECTION_DESC = "BIS: ItemDB item stats (read-only)";

// BIS Planner sheet: A–M only. K = Stat Comparison (script); L = % Upgrade. CmpRaw / equip / scores are script-only.
var GG_UPGRADE_COL = 12;
var CMP_DISP_COL = 11;
/** Logical CmpRaw column index for block assembly; legacy sheets used column P (16) for mirror formulas. */
var CMP_RAW_COL = 16;
var CMP_RAW_LEGACY_SHEET_COL = 16;
/** Muted color when CmpRaw has no +/- stat segments (other plain notices). */
var CMP_NOTICE_COLOR = "#B06000";
/** Exact CmpRaw line for empty CE — shown in black (not CMP_NOTICE_COLOR). */
var CMP_RAW_EQUIP_NOT_SPECIFIED = "Current Equipment Not Specified";
/** Heroic dungeon rows (dark fill + white text in export): lighter Comparison colors for contrast */
var CMP_DELTA_POS_DARK_ROW = "#A5D6A7";
var CMP_DELTA_NEG_DARK_ROW = "#FFAB91";
var CMP_ATTR_LINE_DARK_ROW = "#C8E6C9";
var CMP_NOTICE_DARK_ROW = "#FFE082";
// Legacy K→P mirror: R1C1 offset from K (11) to P (16)
var CMP_RAW_R1C1_OFFSET = CMP_RAW_LEGACY_SHEET_COL - CMP_DISP_COL;

/** After CE / Interest edits: one short wait before reading CmpRaw (P). */
var EDIT_RECALC_WAIT_MS = 150;
/** Before full comparison refresh when a deferred trigger is not used (simple onOpen fallback). */
var OPEN_RECALC_WAIT_MS = 200;
/** Between two open passes so CmpRaw can finish recalculating before the second read. */
var OPEN_SECOND_PASS_SLEEP_MS = 220;
/**
 * When authorized, onOpen schedules comparison refresh this many ms later (separate execution, 6 min budget).
 * Simple onOpen often cannot create triggers; then we fall back to inline refresh after OPEN_RECALC_WAIT_MS.
 */
var OPEN_COMPARISON_DEFER_MS = 1200;

/**
 * Prevents nested refreshComparisonRichText (e.g. installable onEdit firing while script writes K/L/Q/R),
 * which can stack until timeout or appear to hang.
 */
var bisRefreshReentryDepth_ = 0;

/** Same-execution memo for ItemDB/GemDB bulk reads (force refresh pass 2, repeated lookups). */
var bisItemdbExecCache_ = { key: "", data: null };
var bisGemdbExecCache_ = { key: "", data: null };

/**
 * Same-execution memo: CE column A (rows 1..ceEffectiveLastRow) for plannerRefreshCachesBuild_.
 * Avoids a second full-column getValues when equip already read it or when nested refresh runs before outer.
 */
var bisCeGearColExecCache_ = { key: "", vals: null };

/** Document cache (bound spreadsheet): JSON must stay under ~100KB or we skip put. */
var BIS_ITEMDB_DOC_CACHE_PREFIX = "bis_idb_v1_";
var BIS_GEMDB_DOC_CACHE_PREFIX = "bis_gdb_v1_";
var BIS_SHEET_CACHE_TTL_SEC = 3600;
var BIS_SHEET_CACHE_MAX_JSON_CHARS = 95000;

/**
 * First dropdown entry: selecting it clears Interest (unequip). Must be a non-empty string for Sheets.
 */
var BIS_INTEREST_CLEAR = "----";

/** JSON map: equippedSlotKey_(spec, gear) → row number (1-based). Rebuilt on onOpen; reconciles clearOtherEquipped. */
var BIS_EQUIPPED_INDEX_PROP = "BIS_EQUIPPED_INDEX_JSON_V1";

/** Lazy CE map: specNorm → Weights row (1-based). Invalidated when weights row is edited. */
var bisCeWeightsRowMap_ = { key: "", specToWeights: null, colA: null };

/** Same-execution memo: ItemDB column A only (membership checks). */
var bisItemdbNamesExecCache_ = { key: "", names: null };

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
 * twice per spec and will hang or time out. Use one short wait + one full-sheet refresh.
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
  try {
    SpreadsheetApp.getUi()
      .createMenu("BIS Planner")
      .addItem("Recalculate stat comparisons (all rows)", "menuRecalculateStatComparisons_")
      .addToUi();
  } catch (eUi) {}
  SpreadsheetApp.flush();
  if (!bisScheduleDeferredOnOpenComparison_()) {
    Utilities.sleep(OPEN_RECALC_WAIT_MS);
    // Simple spreadsheet onOpen cannot create triggers → no defer. Plain K/L avoids per-row RichText (10×+ faster).
    refreshComparisonRichText(bp, null, null, null, null, null, null, null, true);
  }
}

/** Remove pending one-shot open comparison triggers (avoid stacking on rapid reopen). */
function bisClearDeferredOnOpenComparisonTriggers_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "bisDeferredOnOpenComparison_") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  } catch (e) {}
}

/**
 * @return {boolean} true if a deferred run was scheduled (onOpen returns quickly; K/L update ~1s later).
 * Fails for the normal spreadsheet simple onOpen (no ScriptApp.newTrigger); use menu for colored K/L after plain open.
 */
function bisScheduleDeferredOnOpenComparison_() {
  try {
    bisClearDeferredOnOpenComparisonTriggers_();
    ScriptApp.newTrigger("bisDeferredOnOpenComparison_")
      .timeBased()
      .after(OPEN_COMPARISON_DEFER_MS)
      .create();
    return true;
  } catch (e) {
    return false;
  }
}

/** Time-based trigger target: full comparison + % Upgrade (same as inline onOpen refresh). */
function bisDeferredOnOpenComparison_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var bp = ss.getSheetByName(BIS_PLANNER_SHEET);
    if (!bp) return;
    SpreadsheetApp.flush();
    refreshComparisonRichText(bp, null, null, null, null, null, null, null, false);
  } catch (e) {}
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

/**
 * Menu / repair: two full-sheet passes with a short sleep (CmpRaw may settle between reads).
 * Do not loop per-spec — that multiplies sleeps and can exceed the 6-minute limit.
 */
function refreshComparisonAllSpecs_(bp) {
  Utilities.sleep(OPEN_RECALC_WAIT_MS);
  var ss = bp.getParent();
  var ceSh = ss.getSheetByName(CURRENT_EQUIP_SHEET);
  var sharedCaches = ceSh ? plannerRefreshCachesBuild_(ss, ceSh) : null;
  var fullSheetBandSnap = {};
  refreshComparisonRichText(bp, null, null, null, sharedCaches, {
    captureFullSheetBands: fullSheetBandSnap,
  }, null, null);
  Utilities.sleep(OPEN_SECOND_PASS_SLEEP_MS);
  refreshComparisonRichText(bp, null, null, null, sharedCaches, {
    reuseFullSheetBands: fullSheetBandSnap,
  }, null, null);
}

/**
 * Spreadsheet menu (onOpen) and manual run: two-pass full-sheet comparison + % Upgrade refresh.
 * Use after bulk CE edits, script paste, or if column K looks stale without toggling Equipped.
 */
function menuRecalculateStatComparisons_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  var bp = ss.getSheetByName(BIS_PLANNER_SHEET);
  if (!bp) return;
  refreshComparisonAllSpecs_(bp);
}

// ============================================================
// BIS Planner edits (Interest column)
// ============================================================

function handleBISPlannerEdit(e, bpSheet) {
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row < BIS_FIRST_DATA_ROW) return;

  if (col === GG_INTEREST) {
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

    if (!normalizeItemName(spec) || !normalizeItemName(gearType)) {
      return;
    }

    var plannerBCReuse = null;
    var nv = newValue;
    if (!gearIsRingOrTrinket_(gearType) && (nv === "Equipped 1" || nv === "Equipped 2")) {
      e.range.setValue("Equipped");
      nv = "Equipped";
    }
    var ceGearColSync = null;
    var lastRowEq;
    if (interestIsEquippedState_(nv)) {
      e.range.setFontWeight("bold");
      var ceLabel = ceSlotLabelForPlannerInterest_(gearType, nv);
      lastRowEq = bpSheet.getLastRow();
      var nEq = lastRowEq - BIS_FIRST_DATA_ROW + 1;
      var eqMap = getEquippedIndexMap_();
      var eqKey = equippedSlotKey_(spec, gearType, nv);
      var prevEqRow = eqMap[eqKey];

      if (nEq <= 0) {
        plannerBCReuse = [];
        eqMap[eqKey] = row;
        saveEquippedIndexMap_(eqMap);
      } else if (prevEqRow == null || prevEqRow === "") {
        var abcCold = bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, nEq, GG_GEARTYPE).getValues();
        clearOtherEquipped(bpSheet, row, spec, gearType, nv, abcCold);
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
        var abcFull = bpSheet.getRange(BIS_FIRST_DATA_ROW, 1, nEq, GG_GEARTYPE).getValues();
        plannerBCReuse = [];
        for (var pbi = 0; pbi < abcFull.length; pbi++) {
          plannerBCReuse.push([
            abcFull[pbi][GG_SPEC - 1],
            abcFull[pbi][GG_GEARTYPE - 1],
          ]);
        }
        clearOtherEquippedUsingIndex_(bpSheet, row, spec, gearType, nv, abcFull);
      }
      var ssForCe = bpSheet.getParent();
      var ceSheetSync = ssForCe.getSheetByName(CURRENT_EQUIP_SHEET);
      ceGearColSync = null;
      if (ceSheetSync) {
        var ceLrSync = ceEffectiveLastRow_(ceSheetSync);
        if (ceLrSync >= 1) {
          ceGearColSync = ceSheetSync.getRange(1, CE_GEARTYPE, ceLrSync, 1).getValues();
        }
      }
      var ceRowEq = null;
      var oldCeB = "";
      if (ceSheetSync) {
        ceRowEq = findCERow(ceSheetSync, spec, ceLabel, ceGearColSync);
        if (ceRowEq) {
          oldCeB = normalizeItemName(ceSheetSync.getRange(ceRowEq, CE_ITEMNAME).getValue());
        }
      }
      setCEItem(spec, ceLabel, itemName, ceSheetSync, ceGearColSync);
      if (ceSheetSync && ceRowEq && normalizeItemName(itemName)) {
        ceRefreshStatsAndGearScoreForRow_(ceSheetSync, ceRowEq, oldCeB, itemName, ssForCe);
      }
    } else {
      e.range.setFontWeight("normal");
      if (interestIsEquippedState_(e.oldValue)) {
        var oldCe = ceSlotLabelForPlannerInterest_(gearType, e.oldValue);
        var ssU = bpSheet.getParent();
        var ceSheetU = ssU.getSheetByName(CURRENT_EQUIP_SHEET);
        var ceGearColU = null;
        if (ceSheetU) {
          var ceLrU = ceEffectiveLastRow_(ceSheetU);
          if (ceLrU >= 1) {
            ceGearColU = ceSheetU.getRange(1, CE_GEARTYPE, ceLrU, 1).getValues();
          }
        }
        clearCEItemIfMatch(spec, oldCe, itemName, ceSheetU, ceGearColU);
        clearEquippedIndexIfRow_(row, spec, gearType, e.oldValue);
      }
    }
    SpreadsheetApp.flush();
    Utilities.sleep(EDIT_RECALC_WAIT_MS);
    refreshComparisonRichText(
      bpSheet,
      spec,
      gearType,
      plannerBCReuse,
      null,
      null,
      ceGearColSync,
      lastRowEq
    );
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
    refreshComparisonRichText(bpSheet, specForRow, gearForRow, null, null, null, null, null);
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
/**
 * @param {?Array<Array<*>>} plannerABCForFallback - A:GG_GEARTYPE snapshot (same shape as clearOtherEquipped); avoids full read on index miss.
 */
function clearOtherEquippedUsingIndex_(
  bpSheet,
  currentRow,
  spec,
  gearType,
  interestValue,
  plannerABCForFallback
) {
  var key = equippedSlotKey_(spec, gearType, interestValue);
  var map = getEquippedIndexMap_();
  var prev = map[key];
  var specN = normalizeItemName(spec);
  var gearN = normalizeItemName(gearType);
  var fb = plannerABCForFallback != null && plannerABCForFallback.length > 0 ? plannerABCForFallback : null;

  if (prev == null || prev === "") {
    clearOtherEquipped(bpSheet, currentRow, spec, gearType, interestValue, fb);
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
      clearOtherEquipped(bpSheet, currentRow, spec, gearType, interestValue, fb);
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

/** Safe last row for CE bulk reads / column A scans (see CE_CONTENT_ROW_CAP). */
function ceEffectiveLastRow_(ceSheet) {
  var lr = ceSheet.getLastRow();
  if (lr < CE_FIRST_DATA_ROW) return lr;
  return Math.min(lr, CE_FIRST_DATA_ROW + CE_CONTENT_ROW_CAP - 1);
}

/**
 * Reuse column A matrix for rows 1..celr (1-based effective height).
 * Returns exactly celr rows (truncates if vals is longer).
 * Returns null if vals is null or shorter than celr (must read from sheet).
 */
function ceGearColSnapshotRowsAtCelr_(vals, celr) {
  if (vals == null) return null;
  if (celr < 1) {
    return vals.length === 0 ? [] : null;
  }
  if (vals.length < celr) return null;
  if (vals.length === celr) return vals;
  return vals.slice(0, celr);
}

/**
 * Same as ceGearColSnapshotRowsAtCelr_ but reads ceEffectiveLastRow_(ceSheet) once.
 * Do not call from plannerRefreshCachesBuild_: that function must use a single frozen celr for both
 * cacheKey and snapshot check — two consecutive getLastRow calls can disagree and force a duplicate read.
 */
function ceGearColSnapshotRows_(vals, ceSheet) {
  return ceGearColSnapshotRowsAtCelr_(vals, ceEffectiveLastRow_(ceSheet));
}

/** True when ceGearColSnapshotRows_ can supply column A without a sheet read. */
function ceGearColValsUsable_(vals, ceSheet) {
  return ceGearColSnapshotRows_(vals, ceSheet) != null;
}

function ceIsWeightsRow_(ceSheet, row) {
  var a = normalizeItemName(ceSheet.getRange(row, CE_GEARTYPE).getValue());
  return a === CE_WEIGHT_ROW_LABEL;
}

function ceWeightsMapKey_(ss, ceSheet) {
  return ss.getId() + ":" + ceSheet.getSheetId() + ":" + ceEffectiveLastRow_(ceSheet);
}

function ceInvalidateWeightsRowMap_() {
  bisCeWeightsRowMap_.key = "";
  bisCeWeightsRowMap_.specToWeights = null;
  bisCeWeightsRowMap_.colA = null;
}

/** Nearest section header above row1 (1-based): normalized spec name. */
function specForCERowFromColA_(colA, row1) {
  var i = row1 - 1;
  while (i >= CE_FIRST_DATA_ROW - 1 && i < colA.length) {
    var val = String(colA[i][0]).trim();
    if (val.indexOf("---") === 0) {
      var parsed = specDisplayNameFromSectionHeader(colA[i][0]);
      return parsed ? normalizeItemName(parsed) : null;
    }
    i--;
  }
  return null;
}

function ceEnsureWeightsSpecMap_(ss, ceSheet, colA) {
  var celr = ceEffectiveLastRow_(ceSheet);
  var key = ceWeightsMapKey_(ss, ceSheet);
  var m = {};
  var curSpec = null;
  for (var ri = CE_FIRST_DATA_ROW - 1; ri < colA.length && ri < celr; ri++) {
    var cell = String(colA[ri][0]).trim();
    if (cell.indexOf("---") === 0) {
      var ph = specDisplayNameFromSectionHeader(colA[ri][0]);
      curSpec = ph ? normalizeItemName(ph) : null;
      continue;
    }
    if (cell === CE_WEIGHT_ROW_LABEL && curSpec) {
      m[curSpec] = ri + 1;
    }
  }
  bisCeWeightsRowMap_.key = key;
  bisCeWeightsRowMap_.specToWeights = m;
  bisCeWeightsRowMap_.colA = colA;
}

/** Weights row for the CE section containing anyRow (one column A scan per ss+sheet+celr per execution). */
function ceFindWeightsRowForSection_(ceSheet, anyRow) {
  var ss = ceSheet.getParent();
  var celr = ceEffectiveLastRow_(ceSheet);
  if (anyRow > celr || anyRow < 1) return null;
  var colA;
  var key = ceWeightsMapKey_(ss, ceSheet);
  if (
    bisCeWeightsRowMap_.key === key &&
    bisCeWeightsRowMap_.specToWeights != null &&
    bisCeWeightsRowMap_.colA &&
    bisCeWeightsRowMap_.colA.length === celr
  ) {
    colA = bisCeWeightsRowMap_.colA;
  } else {
    colA = ceSheet.getRange(1, CE_GEARTYPE, celr, 1).getValues();
    ceEnsureWeightsSpecMap_(ss, ceSheet, colA);
  }
  var specHere = specForCERowFromColA_(colA, anyRow);
  if (!specHere) return null;
  var w = bisCeWeightsRowMap_.specToWeights[specHere];
  return w != null ? w : null;
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
  ceInvalidateWeightsRowMap_();
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

function bisGetDocumentCache_() {
  try {
    return CacheService.getDocumentCache();
  } catch (e0) {
    return null;
  }
}

function itemdbLookupSockets_(ss, itemName) {
  var z = { r: 0, y: 0, b: 0, m: 0 };
  if (!itemName) return z;
  var row = itemdbLookupStatsAndSockets_(ss, itemName);
  if (row) return row.socks;
  return z;
}

/**
 * ItemDB row for an item: socket counts + stat vector (same order as Current Equipment stat columns).
 * @return {?{socks:{r:number,y:number,b:number,m:number}, stats:Array<number>}}
 */
function itemdbLookupStatsAndSockets_(ss, itemName) {
  var want = normalizeItemName(itemName);
  if (!want) return null;
  var data = itemdbReadAllStatRows_(ss);
  return itemdbLookupStatsAndSocketsInData_(data, itemName);
}

/** One ItemDB read per script execution; Document Cache when JSON fits size cap. */
function itemdbReadAllStatRows_(ss) {
  var db = ss.getSheetByName(ITEMDB_SHEET);
  if (!db) return null;
  var lr = db.getLastRow();
  if (lr < 2) return null;
  var nStat = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var width = ITEMDB_ATTR_COL_1BASED;
  var execKey = ss.getId() + ":" + lr + ":" + width;
  if (bisItemdbExecCache_.key === execKey && bisItemdbExecCache_.data) {
    return bisItemdbExecCache_.data;
  }
  var docKey = BIS_ITEMDB_DOC_CACHE_PREFIX + ss.getId();
  var sc = bisGetDocumentCache_();
  if (sc) {
    try {
      var blob = sc.get(docKey);
      if (blob) {
        var parsed = JSON.parse(blob);
        if (
          parsed &&
          Number(parsed.lr) === lr &&
          Number(parsed.width) === width &&
          parsed.data
        ) {
          bisItemdbExecCache_.key = execKey;
          bisItemdbExecCache_.data = parsed.data;
          return parsed.data;
        }
      }
    } catch (eIdb) {}
  }
  var data = db.getRange(2, 1, lr, width).getValues();
  bisItemdbExecCache_.key = execKey;
  bisItemdbExecCache_.data = data;
  if (sc) {
    try {
      var payload = JSON.stringify({ lr: lr, width: width, data: data });
      if (payload.length <= BIS_SHEET_CACHE_MAX_JSON_CHARS) {
        sc.put(docKey, payload, BIS_SHEET_CACHE_TTL_SEC);
      }
    } catch (ePut) {}
  }
  return data;
}

/** ItemDB name column only (A2:A{lr}); same-execution memo for membership checks. */
function itemdbReadNameColumn_(ss) {
  var db = ss.getSheetByName(ITEMDB_SHEET);
  if (!db) return null;
  var lr = db.getLastRow();
  if (lr < 2) return null;
  var execKey = ss.getId() + ":idbnames:" + lr;
  if (bisItemdbNamesExecCache_.key === execKey && bisItemdbNamesExecCache_.names) {
    return bisItemdbNamesExecCache_.names;
  }
  var names = db.getRange(2, 1, lr, 1).getValues();
  bisItemdbNamesExecCache_.key = execKey;
  bisItemdbNamesExecCache_.names = names;
  return names;
}

/**
 * ItemDB stat cell → number or "" (blank). ItemDB often stores literal 0 for “no stat”; treat like empty for CE.
 * Scoring uses Number(v)||0 (blank and 0 both contribute 0).
 */
function itemdbStatSheetValue_(raw) {
  if (raw == null) return "";
  if (typeof raw === "string" && raw.replace(/^\s+|\s+$/g, "") === "") return "";
  var n = Number(raw);
  if (isNaN(n)) return "";
  if (n === 0) return "";
  return n;
}

function itemdbLookupStatsAndSocketsInData_(data, itemName) {
  if (!data || !itemName) return null;
  var nStat = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  for (var i = 0; i < data.length; i++) {
    if (!itemsNameMatch_(data[i][0], itemName)) continue;
    var stats = [];
    for (var s = 0; s < nStat; s++) {
      stats.push(itemdbStatSheetValue_(data[i][ITEMDB_FIRST_STAT_COL - 1 + s]));
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

/** Build GemDB category → rows (same row shape as sheet). Category cell is lowercased for keys. */
function gemdbBuildCategoryIndex_(data) {
  var by = {};
  for (var i = 0; i < data.length; i++) {
    var c = String(data[i][1] == null ? "" : data[i][1]).toLowerCase();
    if (!by[c]) by[c] = [];
    by[c].push(data[i]);
  }
  return by;
}

/**
 * GemDB load: rows + category index for bestGemInCategory_. Document Cache when JSON fits.
 * @return {?{rows:Array<Array<*>>, byCategory:Object, lr:number}}
 */
function gemdbReadAllRows_(ss) {
  var sh = ss.getSheetByName(GEMDB_SHEET);
  if (!sh) return null;
  var lr = sh.getLastRow();
  if (lr < 2) return null;
  var gemW = 4;
  var execKey = ss.getId() + ":gem:" + lr;
  if (bisGemdbExecCache_.key === execKey && bisGemdbExecCache_.data) {
    return bisGemdbExecCache_.data;
  }
  var docKey = BIS_GEMDB_DOC_CACHE_PREFIX + ss.getId();
  var sc = bisGetDocumentCache_();
  if (sc) {
    try {
      var blobG = sc.get(docKey);
      if (blobG) {
        var pg = JSON.parse(blobG);
        if (pg && Number(pg.lr) === lr && pg.rows) {
          var byG = pg.byCategory;
          if (!byG) byG = gemdbBuildCategoryIndex_(pg.rows);
          var wrapG = { rows: pg.rows, byCategory: byG, lr: lr };
          bisGemdbExecCache_.key = execKey;
          bisGemdbExecCache_.data = wrapG;
          return wrapG;
        }
      }
    } catch (eGdb) {}
  }
  var data = sh.getRange(2, 1, lr, gemW).getValues();
  var wrapped = { rows: data, byCategory: gemdbBuildCategoryIndex_(data), lr: lr };
  bisGemdbExecCache_.key = execKey;
  bisGemdbExecCache_.data = wrapped;
  if (sc) {
    try {
      var payloadG = JSON.stringify({
        lr: lr,
        rows: data,
        byCategory: wrapped.byCategory,
      });
      if (payloadG.length <= BIS_SHEET_CACHE_MAX_JSON_CHARS) {
        sc.put(docKey, payloadG, BIS_SHEET_CACHE_TTL_SEC);
      }
    } catch (ePutG) {}
  }
  return wrapped;
}

/** Weights row for a planner spec name (must match CE section title case). One bulk read of column A. */
function ceFindWeightsRowForSpec_(ceSheet, spec) {
  var want = normalizeItemName(spec);
  if (!want) return null;
  var lr = ceEffectiveLastRow_(ceSheet);
  if (lr < 1) return null;
  var colA = ceSheet.getRange(1, CE_GEARTYPE, lr, 1).getValues();
  var curSpec = null;
  for (var i = 0; i < colA.length; i++) {
    var a = String(colA[i][0]).trim();
    if (a.indexOf("---") === 0) {
      var parsed = specDisplayNameFromSectionHeader(a);
      curSpec = parsed ? normalizeItemName(parsed) : null;
      continue;
    }
    if (a === CE_WEIGHT_ROW_LABEL && curSpec && curSpec === want) {
      return i + 1;
    }
  }
  return null;
}

/**
 * Ring 1 / Ring 2 gear scores from Current Equipment (same source as % Upgrade baseline).
 * @return {{s1:number, s2:number}}
 */
function plannerRingEquippedSlotScores_(ceSheet, spec, optCaches) {
  var specN = normalizeItemName(spec);
  var sk = specN + "\0Ring\0slots";
  if (
    optCaches &&
    optCaches.ringSlotScores &&
    Object.prototype.hasOwnProperty.call(optCaches.ringSlotScores, sk)
  ) {
    return optCaches.ringSlotScores[sk];
  }
  var aCol = optCaches && optCaches.ceGearCol ? optCaches.ceGearCol : null;
  var r1 = findCERow(ceSheet, spec, "Ring 1", aCol);
  var r2 = findCERow(ceSheet, spec, "Ring 2", aCol);
  var s1;
  var s2;
  var nm1;
  var nm2;
  if (optCaches && optCaches.ceGearScoreCol) {
    s1 = r1 ? plannerCeGearScoreAtRowFromCaches_(optCaches, r1) : 0;
    s2 = r2 ? plannerCeGearScoreAtRowFromCaches_(optCaches, r2) : 0;
    nm1 = r1 ? plannerCeItemNameAtRowFromCaches_(optCaches, r1) : "";
    nm2 = r2 ? plannerCeItemNameAtRowFromCaches_(optCaches, r2) : "";
  } else {
    s1 = r1 ? Number(ceSheet.getRange(r1, CE_GEAR_SCORE_COL).getValue()) || 0 : 0;
    s2 = r2 ? Number(ceSheet.getRange(r2, CE_GEAR_SCORE_COL).getValue()) || 0 : 0;
    nm1 = r1 ? normalizeItemName(ceSheet.getRange(r1, CE_ITEMNAME).getValue()) : "";
    nm2 = r2 ? normalizeItemName(ceSheet.getRange(r2, CE_ITEMNAME).getValue()) : "";
  }
  var out = { s1: s1, s2: s2, hasItem1: nm1 !== "", hasItem2: nm2 !== "" };
  if (optCaches) {
    if (!optCaches.ringSlotScores) optCaches.ringSlotScores = {};
    optCaches.ringSlotScores[sk] = out;
  }
  return out;
}

/**
 * Trinket 1 / Trinket 2 gear scores from Current Equipment (same source as % Upgrade baseline).
 * @return {{s1:number, s2:number}}
 */
function plannerTrinketEquippedSlotScores_(ceSheet, spec, optCaches) {
  var specN = normalizeItemName(spec);
  var sk = specN + "\0Trinket\0slots";
  if (
    optCaches &&
    optCaches.trinketSlotScores &&
    Object.prototype.hasOwnProperty.call(optCaches.trinketSlotScores, sk)
  ) {
    return optCaches.trinketSlotScores[sk];
  }
  var aCol = optCaches && optCaches.ceGearCol ? optCaches.ceGearCol : null;
  var r1 = findCERow(ceSheet, spec, "Trinket 1", aCol);
  var r2 = findCERow(ceSheet, spec, "Trinket 2", aCol);
  var s1;
  var s2;
  var nm1;
  var nm2;
  if (optCaches && optCaches.ceGearScoreCol) {
    s1 = r1 ? plannerCeGearScoreAtRowFromCaches_(optCaches, r1) : 0;
    s2 = r2 ? plannerCeGearScoreAtRowFromCaches_(optCaches, r2) : 0;
    nm1 = r1 ? plannerCeItemNameAtRowFromCaches_(optCaches, r1) : "";
    nm2 = r2 ? plannerCeItemNameAtRowFromCaches_(optCaches, r2) : "";
  } else {
    s1 = r1 ? Number(ceSheet.getRange(r1, CE_GEAR_SCORE_COL).getValue()) || 0 : 0;
    s2 = r2 ? Number(ceSheet.getRange(r2, CE_GEAR_SCORE_COL).getValue()) || 0 : 0;
    nm1 = r1 ? normalizeItemName(ceSheet.getRange(r1, CE_ITEMNAME).getValue()) : "";
    nm2 = r2 ? normalizeItemName(ceSheet.getRange(r2, CE_ITEMNAME).getValue()) : "";
  }
  var out = { s1: s1, s2: s2, hasItem1: nm1 !== "", hasItem2: nm2 !== "" };
  if (optCaches) {
    if (!optCaches.trinketSlotScores) optCaches.trinketSlotScores = {};
    optCaches.trinketSlotScores[sk] = out;
  }
  return out;
}

/** True when Current Equipment row has an item name in column B (required before % Upgrade shows). */
function plannerCeRowHasEquippedItemName_(ceSheet, row, optCaches) {
  if (!row) return false;
  if (optCaches && optCaches.ceItemNameCol) {
    return plannerCeItemNameAtRowFromCaches_(optCaches, row) !== "";
  }
  return normalizeItemName(ceSheet.getRange(row, CE_ITEMNAME).getValue()) !== "";
}

/**
 * Non–ring/trinket gear: CE row for spec + gear has an item to compare against.
 */
function plannerCeHasEquippedItemForPlannerGear_(ceSheet, spec, plannerGearType, optCaches) {
  var g = normalizeItemName(plannerGearType);
  if (!g) return false;
  var aCol = optCaches && optCaches.ceGearCol ? optCaches.ceGearCol : null;
  var r = findCERow(ceSheet, spec, g, aCol);
  return plannerCeRowHasEquippedItemName_(ceSheet, r, optCaches);
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
 * Baseline for % Upgrade: read "Gear Score" from Current Equipment for the equipped slot(s).
 * Uses the same score shown in CE (manual custom items and ItemDB rows). Rings/trinkets: min of the
 * two slot scores when each is positive (weaker-slot baseline). Ring and trinket rows in L show both slots.
 */
function plannerBaselineEquippedScore_(ss, ceSheet, spec, plannerGearType, eqN, eq2N, optCaches) {
  var g = normalizeItemName(plannerGearType);
  var specN = normalizeItemName(spec);
  // One baseline per CE slot (spec + gear type); not keyed by planner Equip (N/O).
  var sk = specN + "\0" + g;
  if (optCaches && Object.prototype.hasOwnProperty.call(optCaches.baselineByKey, sk)) {
    return optCaches.baselineByKey[sk];
  }
  var aCol = optCaches && optCaches.ceGearCol ? optCaches.ceGearCol : null;
  var out = 0;
  var r0 = null;
  if (g === "Ring" || g === "Trinket") {
    if (g === "Ring") {
      var slotR = plannerRingEquippedSlotScores_(ceSheet, spec, optCaches);
      var dR = [];
      if (slotR.s1 > 0) dR.push(slotR.s1);
      if (slotR.s2 > 0) dR.push(slotR.s2);
      out = dR.length === 0 ? 0 : Math.min.apply(null, dR);
    } else {
      var slotT = plannerTrinketEquippedSlotScores_(ceSheet, spec, optCaches);
      var dT = [];
      if (slotT.s1 > 0) dT.push(slotT.s1);
      if (slotT.s2 > 0) dT.push(slotT.s2);
      out = dT.length === 0 ? 0 : Math.min.apply(null, dT);
    }
  } else {
    r0 = findCERow(ceSheet, spec, g, aCol);
    if (optCaches && optCaches.ceGearScoreCol) {
      out = r0 ? plannerCeGearScoreAtRowFromCaches_(optCaches, r0) : 0;
    } else {
      out = r0 ? Number(ceSheet.getRange(r0, CE_GEAR_SCORE_COL).getValue()) || 0 : 0;
    }
  }
  if (optCaches) optCaches.baselineByKey[sk] = out;
  return out;
}

function plannerCeGearScoreAtRowFromCaches_(caches, row1Based) {
  if (!caches || !caches.ceGearScoreCol || !row1Based) return 0;
  var i = row1Based - 1;
  if (i < 0 || i >= caches.ceGearScoreCol.length) return 0;
  return Number(caches.ceGearScoreCol[i][0]) || 0;
}

function plannerCeItemNameAtRowFromCaches_(caches, row1Based) {
  if (!caches || !caches.ceItemNameCol || !row1Based) return "";
  var i = row1Based - 1;
  if (i < 0 || i >= caches.ceItemNameCol.length) return "";
  return normalizeItemName(caches.ceItemNameCol[i][0]);
}

/** Precompute ideal reg/meta gem score per socket for each spec (weights fixed per spec; was repeated per ItemDB row). */
function plannerWarmWeightIdealGems_(caches) {
  var st = caches.weights.specToRow;
  for (var sk in st) {
    if (!Object.prototype.hasOwnProperty.call(st, sk)) continue;
    var w = caches.weights.get(sk);
    if (!w || w.idealRegScorePerSocket != null) continue;
    var bestRegScore = -1;
    for (var ci = 0; ci < CE_GEM_NON_META_CATS.length; ci++) {
      var best = bestGemInCategory_(caches.ss, CE_GEM_NON_META_CATS[ci], w.wVals, caches.gemData);
      if (best && best.score > bestRegScore) bestRegScore = best.score;
    }
    w.idealRegScorePerSocket = bestRegScore >= 0 ? bestRegScore : 0;
    var bm = bestGemInCategory_(caches.ss, "meta", w.wVals, caches.gemData);
    w.idealMetaGemPerSocket = bm ? bm.score : 0;
  }
}

/**
 * Bulk-read CE column A once + ItemDB + GemDB; lazy weight vectors per spec; memo baselines + item scores.
 * Used only inside refreshComparisonRichTextInner_ to avoid thousands of sheet reads.
 * @param {?Array<Array<*>>} optCeGearColVals - if length matches ceEffectiveLastRow_(ceSheet), reuse (no sheet read).
 * @param {?Array<Array<*>>} optItemdbData - when non-null, reuse instead of itemdbReadAllStatRows_.
 * @param {?{rows:Array<Array<*>>, byCategory:Object, lr:number}} optGemData - when non-null, reuse instead of gemdbReadAllRows_.
 */
function plannerRefreshCachesBuild_(ss, ceSheet, optCeGearColVals, optItemdbData, optGemData) {
  var itemdbData =
    optItemdbData != null ? optItemdbData : itemdbReadAllStatRows_(ss);
  var gemData = optGemData != null ? optGemData : gemdbReadAllRows_(ss);
  var celr = ceEffectiveLastRow_(ceSheet);
  var cacheKey = ss.getId() + ":" + ceSheet.getSheetId() + ":" + celr;
  var aCol;
  var colAFromSnap = ceGearColSnapshotRowsAtCelr_(optCeGearColVals, celr);
  var colAFromExec = null;
  if (
    colAFromSnap == null &&
    bisCeGearColExecCache_.key === cacheKey &&
    bisCeGearColExecCache_.vals != null
  ) {
    colAFromExec = ceGearColSnapshotRowsAtCelr_(bisCeGearColExecCache_.vals, celr);
  }
  if (colAFromSnap != null) {
    aCol = colAFromSnap;
    bisCeGearColExecCache_.key = cacheKey;
    bisCeGearColExecCache_.vals = aCol;
  } else if (colAFromExec != null) {
    aCol = colAFromExec;
  } else {
    aCol = celr >= 1 ? ceSheet.getRange(1, CE_GEARTYPE, celr, 1).getValues() : [];
    bisCeGearColExecCache_.key = cacheKey;
    bisCeGearColExecCache_.vals = aCol;
  }
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
      var wBandW = CE_META_WEIGHT_COL - CE_STAT_FIRST_COL + 1;
      var wBand = this.ceSheet.getRange(row, CE_STAT_FIRST_COL, 1, wBandW).getValues()[0];
      var wVals = wBand.slice(0, n);
      var metaSockW = Number(wBand[wBandW - 1]) || 0;
      var o = { wRow: row, wVals: wVals, metaSockW: metaSockW };
      this.wBySpec[k] = o;
      return o;
    },
  };
  var outCaches = {
    ss: ss,
    itemdbData: itemdbData,
    gemData: gemData,
    weights: weights,
    ceGearCol: aCol,
    ceSheet: ceSheet,
    baselineByKey: {},
    ringSlotScores: {},
    trinketSlotScores: {},
    itemScoreBaseByKey: {},
    itemScoreFullByKey: {},
    ceGearScoreCol:
      celr >= 1 ? ceSheet.getRange(1, CE_GEAR_SCORE_COL, celr, 1).getValues() : [],
    ceItemNameCol:
      celr >= 1 ? ceSheet.getRange(1, CE_ITEMNAME, celr, 1).getValues() : [],
  };
  plannerWarmWeightIdealGems_(outCaches);
  return outCaches;
}

/**
 * One ItemDB + gem-plan pass; fills both memo maps (avoids duplicate computeIdealGemPlan_ when callers
 * need base and full for the same row).
 * @return {{base:number, full:number}}
 */
function plannerItemGearScoresBothWithCaches_(caches, spec, itemName) {
  var nm = normalizeItemName(itemName);
  if (!nm) return { base: 0, full: 0 };
  var ck = normalizeItemName(spec) + "\0" + nm;
  var hasB = Object.prototype.hasOwnProperty.call(caches.itemScoreBaseByKey, ck);
  var hasF = Object.prototype.hasOwnProperty.call(caches.itemScoreFullByKey, ck);
  if (hasB && hasF) {
    return { base: caches.itemScoreBaseByKey[ck], full: caches.itemScoreFullByKey[ck] };
  }
  var w = caches.weights.get(spec);
  if (!w) {
    caches.itemScoreBaseByKey[ck] = 0;
    caches.itemScoreFullByKey[ck] = 0;
    return { base: 0, full: 0 };
  }
  var row = itemdbLookupStatsAndSocketsInData_(caches.itemdbData, itemName);
  if (!row) {
    caches.itemScoreBaseByKey[ck] = 0;
    caches.itemScoreFullByKey[ck] = 0;
    return { base: 0, full: 0 };
  }
  var base = 0;
  for (var i = 0; i < w.wVals.length; i++) {
    base += (Number(w.wVals[i]) || 0) * (Number(row.stats[i]) || 0);
  }
  var baseR = Math.round(base * 100) / 100;
  var gemScore;
  if (w.idealRegScorePerSocket != null) {
    var T = (row.socks.r || 0) + (row.socks.y || 0) + (row.socks.b || 0);
    var m = row.socks.m || 0;
    gemScore = 0;
    if (T > 0) gemScore += T * w.idealRegScorePerSocket;
    if (m > 0) {
      gemScore += m * (Number(w.metaSockW) || 0);
      gemScore += m * (w.idealMetaGemPerSocket || 0);
    }
  } else {
    gemScore = computeIdealGemPlan_(caches.ss, w.wVals, w.metaSockW, row.socks, caches.gemData).gemScore;
  }
  var fullR = Math.round((base + gemScore) * 100) / 100;
  caches.itemScoreBaseByKey[ck] = baseR;
  caches.itemScoreFullByKey[ck] = fullR;
  return { base: baseR, full: fullR };
}

function plannerItemGearScoreBaseOnlyWithCaches_(caches, spec, itemName) {
  return plannerItemGearScoresBothWithCaches_(caches, spec, itemName).base;
}

function plannerItemGearScoreWithCaches_(caches, spec, itemName) {
  return plannerItemGearScoresBothWithCaches_(caches, spec, itemName).full;
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
  var data;
  var filterByCat = false;
  if (optGemData && optGemData.byCategory) {
    data = optGemData.byCategory[cat] || [];
  } else if (Array.isArray(optGemData)) {
    data = optGemData;
    filterByCat = true;
  } else {
    data = null;
  }
  if (data == null) {
    var sh = ss.getSheetByName(GEMDB_SHEET);
    if (!sh) return null;
    var lr = sh.getLastRow();
    if (lr < 2) return null;
    data = sh.getRange(2, 1, lr, 4).getValues();
    filterByCat = true;
  }
  var bestScore = -1;
  var bestLabel = "";
  var bestStats = null;
  for (var i = 0; i < data.length; i++) {
    if (filterByCat && String(data[i][1]) !== cat) continue;
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
 * @param {boolean=} ceHasEquippedItem — when false, no item on CE for that slot → blank % (omit for ring/trinket dual path).
 */
function upgradePctFromScores_(itemScorePlusGems, equippedBaseline, heroicDark, hideRow, ceHasEquippedItem) {
  var out = { text: "", color: "#000000" };
  if (hideRow) return out;
  if (ceHasEquippedItem === false) return out;
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

/**
 * % Upgrade for one ring/trinket CE slot vs the planner row item gear score (stats × weights + ideal gems only).
 * Requires hasEquippedItem true (CE Item Name set for that ring/trinket slot). If that slot's gear score is ≤0
 * but an item is equipped, baseline is treated as 0: row score > 0 → +100%; row score ≤0 → 0%. If baseline > 0,
 * uses (item − baseline) / baseline × 100.
 */
function upgradePctForRingTrinketSlot_(itemScorePlusGems, slotBaseline, heroicDark, hideRow, hasEquippedItem) {
  var out = { text: "", color: "#000000" };
  if (hideRow) return out;
  if (hasEquippedItem === false) return out;
  var item = Number(itemScorePlusGems) || 0;
  var base = Number(slotBaseline) || 0;
  var posC = heroicDark ? CMP_DELTA_POS_DARK_ROW : "#0d652d";
  var negC = heroicDark ? CMP_DELTA_NEG_DARK_ROW : "#c5221f";
  var zeroC = heroicDark ? "#EEEEEE" : "#000000";
  if (base <= 0) {
    if (item <= 0) {
      out.text = "0%";
      out.color = zeroC;
      return out;
    }
    out.text = "+100%";
    out.color = posC;
    return out;
  }
  var pct = ((item - base) / base) * 100;
  var rounded = Math.round(pct * 10) / 10;
  var sign = rounded > 0 ? "+" : "";
  out.text = sign + rounded + "%";
  out.color = rounded > 0 ? posC : rounded < 0 ? negC : zeroC;
  return out;
}

/**
 * % Upgrade for ring/trinket: only slots with an equipped item (CE Item Name) get a line.
 * @return {?Object} RichTextValue from build(), or null if neither slot has an item to compare.
 */
function buildDualSlotUpgradeRichText_(itemScorePlusGems, s1, s2, hasItem1, hasItem2, heroicDark) {
  var itemNum = Number(itemScorePlusGems) || 0;
  var lines = [];
  if (hasItem1) {
    var p1 = upgradePctForRingTrinketSlot_(itemNum, s1, heroicDark, false, true);
    if (p1.text !== "") lines.push({ header: "Slot 1:\n", pct: p1.text, color: p1.color });
  }
  if (hasItem2) {
    var p2 = upgradePctForRingTrinketSlot_(itemNum, s2, heroicDark, false, true);
    if (p2.text !== "") lines.push({ header: "Slot 2:\n", pct: p2.text, color: p2.color });
  }
  if (lines.length === 0) return null;
  var full = "";
  var li;
  for (li = 0; li < lines.length; li++) {
    if (li > 0) full += "\n";
    full += lines[li].header + lines[li].pct;
  }
  var slotLabC = heroicDark ? CMP_NOTICE_DARK_ROW : CMP_NOTICE_COLOR;
  var slotLabStyle = SpreadsheetApp.newTextStyle().setForegroundColor(slotLabC).build();
  var b = SpreadsheetApp.newRichTextValue().setText(full);
  var pos = 0;
  for (li = 0; li < lines.length; li++) {
    if (li > 0) pos += 1;
    var ln = lines[li];
    var hLen = ln.header.length;
    var pctLen = ln.pct.length;
    b.setTextStyle(pos, pos + hLen, slotLabStyle);
    b.setTextStyle(
      pos + hLen,
      pos + hLen + pctLen,
      SpreadsheetApp.newTextStyle().setForegroundColor(ln.color).build()
    );
    pos += hLen + pctLen;
  }
  return b.build();
}

/** Plain % Upgrade text for ring/trinket rows (batched setValues; no per-cell RichText). */
function buildDualSlotUpgradePlainText_(itemScorePlusGems, s1, s2, hasItem1, hasItem2) {
  var itemNum = Number(itemScorePlusGems) || 0;
  var parts = [];
  if (hasItem1) {
    var p1 = upgradePctForRingTrinketSlot_(itemNum, s1, false, false, true);
    if (p1.text !== "") parts.push("Slot 1:\n" + p1.text);
  }
  if (hasItem2) {
    var p2 = upgradePctForRingTrinketSlot_(itemNum, s2, false, false, true);
    if (p2.text !== "") parts.push("Slot 2:\n" + p2.text);
  }
  if (parts.length === 0) return "";
  return parts.join("\n");
}

function flushUpgradePctCells_(bpSheet, upgradeWrites) {
  if (!upgradeWrites || upgradeWrites.length === 0) {
    return;
  }
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
    var anyRich = false;
    for (var jr = 0; jr < h; jr++) {
      if (seg[jr].rich != null) anyRich = true;
    }
    if (anyRich) {
      var jc = 0;
      while (jc < h) {
        if (seg[jc].rich != null) {
          var jEnd = jc;
          while (
            jEnd + 1 < h &&
            seg[jEnd + 1].rich != null &&
            seg[jEnd + 1].r === seg[jEnd].r + 1
          ) {
            jEnd++;
          }
          var r0u = seg[jc].r;
          var hu = seg[jEnd].r - r0u + 1;
          var rtm = [];
          for (var ju = jc; ju <= jEnd; ju++) {
            rtm.push([seg[ju].rich]);
          }
          bpSheet.getRange(r0u, GG_UPGRADE_COL, hu, 1).setRichTextValues(rtm);
          jc = jEnd + 1;
        } else {
          var w1 = seg[jc];
          var c1 = bpSheet.getRange(w1.r, GG_UPGRADE_COL);
          c1.setValue(w1.text == null ? "" : w1.text);
          try {
            c1.setFontColor(w1.color || "#000000");
          } catch (eOne) {}
          jc++;
        }
      }
      continue;
    }
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
}

function flushPlannerGearScoreCells_(bpSheet, gearWrites) {
  // Gear scores are computed in script only; planner layout v2 has no Q/R columns.
  void bpSheet;
  void gearWrites;
}

function ceItemNameInItemDb_(ss, itemName) {
  var want = normalizeItemName(itemName);
  if (!want) return false;
  var names = itemdbReadNameColumn_(ss);
  if (!names) return false;
  for (var ni = 0; ni < names.length; ni++) {
    if (!itemsNameMatch_(names[ni][0], itemName)) continue;
    var full = itemdbLookupStatsAndSockets_(ss, itemName);
    if (full == null) return false;
    return true;
  }
  return false;
}

function ceClearStatCellsToBlank_(ceSheet, row) {
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var blanks = [];
  for (var bi = 0; bi < n; bi++) blanks.push("");
  ceSheet.getRange(row, CE_STAT_FIRST_COL, 1, n).setValues([blanks]);
}

/** Matches generate_bis_planner.py ItemDB!A$2:…$lastRow for VLOOKUP width. */
function itemdbBoundedRangeA1_(ss) {
  var db = ss.getSheetByName(ITEMDB_SHEET);
  if (!db) return null;
  var lr = db.getLastRow();
  if (lr < 2) return null;
  var nStat = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var width = ITEMDB_FIRST_STAT_COL - 1 + nStat;
  return "ItemDB!A$2:" + colLetterFromNum_(width) + "$" + lr;
}

/** Empty item name: restore stat columns to ItemDB VLOOKUP (same as exported workbook). */
function ceRestoreStatLookupFormulasForRow_(ceSheet, row, ss, optItemdbRangeA1) {
  var rng =
    arguments.length >= 4 && optItemdbRangeA1 ? optItemdbRangeA1 : itemdbBoundedRangeA1_(ss);
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  if (!rng) {
    ceClearStatCellsToBlank_(ceSheet, row);
    return;
  }
  var formulas = [];
  for (var si = 0; si < n; si++) {
    var dbCol = ITEMDB_FIRST_STAT_COL + si;
    formulas.push("=IFERROR(VLOOKUP(B" + row + "," + rng + "," + dbCol + ',FALSE),"")');
  }
  ceSheet.getRange(row, CE_STAT_FIRST_COL, 1, n).setFormulas([formulas]);
}

/** Same as ceUpdate empty-item branch for gem/score/VLOOKUP; use after bulk-removing stat protections. */
function ceClearEmptyGearSlotRow_(ceSheet, row, ss, optItemdbRangeA1) {
  ceClearGemCellValidations_(ceSheet, row);
  ceRestoreStatLookupFormulasForRow_(ceSheet, row, ss, optItemdbRangeA1);
  ceSheet.getRange(row, CE_IDEAL_GEM_COL, 1, CE_TOTAL_GEM_STATS_COL - CE_IDEAL_GEM_COL + 1).setValues([["", "", ""]]);
  ceSheet.getRange(row, CE_IDEAL_GEM_COL, 1, CE_TOTAL_GEM_STATS_COL - CE_IDEAL_GEM_COL + 1).setBackground(CE_GEM_LOCKED_BG);
  ceSheet.getRange(row, CE_GEAR_SCORE_COL).setValue("");
}

/** One getProtections pass: remove our read-only stat protections for listed CE rows (rowSet keys = row number as string). */
function ceRemoveOurStatProtectionsForRows_(ceSheet, rowSet) {
  var protections = ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var toRemove = [];
  for (var pi = 0; pi < protections.length; pi++) {
    var pr = protections[pi];
    if (String(pr.getDescription()) !== CE_STAT_PROTECTION_DESC) continue;
    var rng = pr.getRange();
    if (rng.getSheet().getSheetId() !== ceSheet.getSheetId()) continue;
    if (rng.getColumn() !== CE_STAT_FIRST_COL || rng.getNumColumns() !== n) continue;
    var rr = rng.getRow();
    if (rowSet[String(rr)]) toRemove.push(pr);
  }
  for (var ti = 0; ti < toRemove.length; ti++) {
    try {
      toRemove[ti].remove();
    } catch (ePr) {}
  }
}

function ceClearGemCellValidations_(ceSheet, row) {
  ceSheet.getRange(row, CE_IDEAL_GEM_COL).setDataValidation(null);
  ceSheet.getRange(row, CE_IDEAL_META_COL).setDataValidation(null);
}

/** @param {?Array<GoogleAppsScript.Spreadsheet.Protection>} optRangeProtections - from one getProtections() pass (open refresh). */
function ceRemoveStatRowProtection_(ceSheet, row, optRangeProtections) {
  var protections = optRangeProtections || ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var pi = 0; pi < protections.length; pi++) {
    var pr = protections[pi];
    if (String(pr.getDescription()) !== CE_STAT_PROTECTION_DESC) continue;
    var rng = pr.getRange();
    if (rng.getSheet().getSheetId() !== ceSheet.getSheetId()) continue;
    if (rng.getRow() !== row) continue;
    if (rng.getColumn() !== CE_STAT_FIRST_COL) continue;
    if (rng.getNumColumns() !== CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1) continue;
    pr.remove();
  }
}

function ceProtectStatRow_(ceSheet, row, optRangeProtections) {
  ceRemoveStatRowProtection_(ceSheet, row, optRangeProtections);
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var rng = ceSheet.getRange(row, CE_STAT_FIRST_COL, 1, n);
  var prot = rng.protect().setDescription(CE_STAT_PROTECTION_DESC);
  try {
    var eds = prot.getEditors();
    for (var ei = 0; ei < eds.length; ei++) {
      prot.removeEditor(eds[ei]);
    }
  } catch (eRm) {}
}

function ceApplyItemDbStatsToRow_(ceSheet, row, ss, itemName) {
  var rowData = itemdbLookupStatsAndSockets_(ss, itemName);
  if (!rowData) return;
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var rowVals = [];
  for (var si = 0; si < n; si++) {
    var sv = rowData.stats[si];
    rowVals.push(sv === "" || sv == null || sv === 0 ? "" : sv);
  }
  ceSheet.getRange(row, CE_STAT_FIRST_COL, 1, n).setValues([rowVals]);
}

/**
 * Manual gear rows only (caller should ensure): turn 0 / "0" into blanks in CE stat columns.
 * Skips the Weights row. Leaves formula cells unchanged so VLOOKUP rows are not stripped.
 */
function ceScrubStatRowDisplayZerosToBlank_(ceSheet, row) {
  if (ceIsWeightsRow_(ceSheet, row)) return;
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var rng = ceSheet.getRange(row, CE_STAT_FIRST_COL, 1, n);
  var formulas = rng.getFormulas()[0];
  var vals = rng.getValues()[0];
  var out = [];
  var changed = false;
  for (var zi = 0; zi < n; zi++) {
    var f = formulas[zi];
    var fs = f ? String(f).replace(/^\s+|\s+$/g, "") : "";
    if (fs.charAt(0) === "=") {
      out.push(f);
      continue;
    }
    var v = vals[zi];
    if (v === 0 || v === "0") {
      out.push("");
      changed = true;
    } else {
      out.push(v);
    }
  }
  if (changed) rng.setValues([out]);
}

/** True if this CE row already has our read-only stat protection (avoids remove+protect on every onOpen). */
function ceStatRowHasOurProtection_(ceSheet, row, rangeProtections) {
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  for (var pi = 0; pi < rangeProtections.length; pi++) {
    var pr = rangeProtections[pi];
    if (String(pr.getDescription()) !== CE_STAT_PROTECTION_DESC) continue;
    var rng = pr.getRange();
    if (rng.getSheet().getSheetId() !== ceSheet.getSheetId()) continue;
    if (rng.getRow() === row && rng.getColumn() === CE_STAT_FIRST_COL && rng.getNumColumns() === n) return true;
  }
  return false;
}

function ceParseLeadingIntFromSocketLabel_(s) {
  var m = String(s == null ? "" : s).match(/^(\d+)/);
  if (!m) return 0;
  return Number(m[1]) || 0;
}

/** CE column A is Trinket 1 / Trinket 2 — no sockets in TBC item DB for trinkets. */
function ceManualRowIsTrinket_(ceSheet, row) {
  var gt = String(ceSheet.getRange(row, CE_GEARTYPE).getValue()).trim().toLowerCase();
  return gt === "trinket 1" || gt === "trinket 2";
}

function ceClearManualTrinketGemBlock_(ceSheet, row) {
  ceClearGemCellValidations_(ceSheet, row);
  ceSheet.getRange(row, CE_IDEAL_GEM_COL, 1, CE_TOTAL_GEM_STATS_COL - CE_IDEAL_GEM_COL + 1).setValues([["", "", ""]]);
  ceSheet.getRange(row, CE_IDEAL_GEM_COL, 1, CE_TOTAL_GEM_STATS_COL - CE_IDEAL_GEM_COL + 1).setBackground(CE_GEM_LOCKED_BG);
  ceSheet.getRange(row, CE_IDEAL_GEM_COL).setFontColor("#000000");
  ceSheet.getRange(row, CE_IDEAL_META_COL).setFontColor("#000000");
  ceSheet.getRange(row, CE_TOTAL_GEM_STATS_COL).setFontColor("#000000");
}

function ceSetupManualGemSocketDropdowns_(ceSheet, row) {
  if (ceManualRowIsTrinket_(ceSheet, row)) {
    ceClearGemCellValidations_(ceSheet, row);
    return;
  }
  var rIdeal = SpreadsheetApp.newDataValidation()
    .requireValueInList(CE_MANUAL_REG_SOCKET_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  var rMeta = SpreadsheetApp.newDataValidation()
    .requireValueInList(CE_MANUAL_META_SOCKET_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  var cellI = ceSheet.getRange(row, CE_IDEAL_GEM_COL);
  var cellM = ceSheet.getRange(row, CE_IDEAL_META_COL);
  var curI = cellI.getValue();
  var curM = cellM.getValue();
  cellI.setDataValidation(rIdeal);
  cellM.setDataValidation(rMeta);
  if (!curI || CE_MANUAL_REG_SOCKET_OPTIONS.indexOf(String(curI)) < 0) {
    cellI.setValue(CE_MANUAL_REG_SOCKET_OPTIONS[0]);
  }
  if (!curM || CE_MANUAL_META_SOCKET_OPTIONS.indexOf(String(curM)) < 0) {
    cellM.setValue(CE_MANUAL_META_SOCKET_OPTIONS[0]);
  }
}

/**
 * Manual CE row: C/D hold socket-count dropdowns; E is aggregated best-gem stats; score includes gems.
 * Scrubs display zeros in stat columns on each run (onOpen + any edit that hits this path).
 */
function ceRecalculateManualIdealGemsAndScore_(ceSheet, row, ss, wRow) {
  var itemChk = normalizeItemName(ceSheet.getRange(row, CE_ITEMNAME).getValue());
  if (itemChk && !ceItemNameInItemDb_(ss, itemChk)) {
    ceScrubStatRowDisplayZerosToBlank_(ceSheet, row);
  }
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var wVals = ceSheet.getRange(wRow, CE_STAT_FIRST_COL, 1, n).getValues()[0];
  if (ceManualRowIsTrinket_(ceSheet, row)) {
    ceClearManualTrinketGemBlock_(ceSheet, row);
    var scoreT = 0;
    var sValsT = ceSheet.getRange(row, CE_STAT_FIRST_COL, 1, n).getValues()[0];
    for (var ti = 0; ti < wVals.length; ti++) {
      scoreT += (Number(sValsT[ti]) || 0) * (Number(wVals[ti]) || 0);
    }
    ceSheet.getRange(row, CE_GEAR_SCORE_COL).setValue(Math.round(scoreT * 100) / 100);
    return;
  }
  var metaSockW = Number(ceSheet.getRange(wRow, CE_META_WEIGHT_COL).getValue()) || 0;
  var gemData = gemdbReadAllRows_(ss);
  var regLabel = String(ceSheet.getRange(row, CE_IDEAL_GEM_COL).getDisplayValue() || "");
  var metaLabel = String(ceSheet.getRange(row, CE_IDEAL_META_COL).getDisplayValue() || "");
  var T = ceParseLeadingIntFromSocketLabel_(regLabel);
  var m = ceParseLeadingIntFromSocketLabel_(metaLabel);
  if (m > 1) m = 1;
  var socks = { r: T, y: 0, b: 0, m: m };
  var plan = computeIdealGemPlan_(ss, wVals, metaSockW, socks, gemData);
  var acc = aggregateGemStats_(plan.regStats, plan.T, plan.metaStats, plan.m);
  ceSheet.getRange(row, CE_TOTAL_GEM_STATS_COL).setValue(formatAggregatedGemStats_(acc));
  ceSheet.getRange(row, CE_TOTAL_GEM_STATS_COL).setBackground(T > 0 || m > 0 ? null : CE_GEM_LOCKED_BG);
  ceSheet.getRange(row, CE_IDEAL_GEM_COL).setBackground(T > 0 ? null : CE_GEM_LOCKED_BG);
  ceSheet.getRange(row, CE_IDEAL_META_COL).setBackground(m > 0 ? null : CE_GEM_LOCKED_BG);
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

/** Updates ideal gem display, backgrounds, and Gear Score for one CE row. */
function ceUpdateIdealGemCellsAndScore_(ceSheet, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wRow = ceFindWeightsRowForSection_(ceSheet, row);
  if (!wRow) return;
  var item = normalizeItemName(ceSheet.getRange(row, CE_ITEMNAME).getValue());
  if (!item) {
    ceRemoveStatRowProtection_(ceSheet, row);
    ceClearGemCellValidations_(ceSheet, row);
    ceRestoreStatLookupFormulasForRow_(ceSheet, row, ss);
    ceSheet.getRange(row, CE_IDEAL_GEM_COL, 1, CE_TOTAL_GEM_STATS_COL - CE_IDEAL_GEM_COL + 1).setValues([["", "", ""]]);
    ceSheet.getRange(row, CE_IDEAL_GEM_COL, 1, CE_TOTAL_GEM_STATS_COL - CE_IDEAL_GEM_COL + 1).setBackground(CE_GEM_LOCKED_BG);
    ceSheet.getRange(row, CE_GEAR_SCORE_COL).setValue("");
    return;
  }
  var n = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  var wBandW = CE_META_WEIGHT_COL - CE_STAT_FIRST_COL + 1;
  var wBand = ceSheet.getRange(wRow, CE_STAT_FIRST_COL, 1, wBandW).getValues()[0];
  var wVals = wBand.slice(0, n);
  var metaSockW = Number(wBand[wBandW - 1]) || 0;
  var gemData = gemdbReadAllRows_(ss);
  var itemRowDb = itemdbLookupStatsAndSockets_(ss, item);
  if (itemRowDb) {
    ceClearGemCellValidations_(ceSheet, row);
    var socks = itemRowDb.socks;
    var Tdb = socks.r + socks.y + socks.b;
    var mdb = socks.m || 0;
    var planDb = computeIdealGemPlan_(ss, wVals, metaSockW, socks, gemData);
    var accDb = aggregateGemStats_(planDb.regStats, planDb.T, planDb.metaStats, planDb.m);
    var gemRow = [
      Tdb > 0 ? planDb.regLabel || "—" : "",
      mdb > 0 ? planDb.metaLabel || "—" : "",
      formatAggregatedGemStats_(accDb),
    ];
    var gemRng = ceSheet.getRange(row, CE_IDEAL_GEM_COL, 1, CE_TOTAL_GEM_STATS_COL - CE_IDEAL_GEM_COL + 1);
    gemRng.setValues([gemRow]);
    gemRng.setBackgrounds([
      [
        Tdb > 0 ? null : CE_GEM_LOCKED_BG,
        mdb > 0 ? null : CE_GEM_LOCKED_BG,
        Tdb > 0 || mdb > 0 ? null : CE_GEM_LOCKED_BG,
      ],
    ]);
    gemRng.setFontColor("#000000");
    var sValsDb = itemRowDb.stats;
    var scoreDb = 0;
    for (var idb = 0; idb < wVals.length; idb++) {
      scoreDb += (Number(sValsDb[idb]) || 0) * (Number(wVals[idb]) || 0);
    }
    scoreDb += planDb.gemScore;
    ceSheet.getRange(row, CE_GEAR_SCORE_COL).setValue(Math.round(scoreDb * 100) / 100);
    return;
  }
  ceSetupManualGemSocketDropdowns_(ceSheet, row);
  ceRecalculateManualIdealGemsAndScore_(ceSheet, row, ss, wRow);
}

function recalculateGearScoreForRow_(ceSheet, row) {
  ceUpdateIdealGemCellsAndScore_(ceSheet, row);
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

/** Full refresh (all gear rows). Prefer ceRefreshDerivedOnOpen_ from onOpen. */
function ceRefreshAllDerived_(ceSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lr = ceEffectiveLastRow_(ceSheet);
  if (lr < CE_FIRST_DATA_ROW) return;
  var numRows = lr - CE_FIRST_DATA_ROW + 1;
  var ab = ceSheet.getRange(CE_FIRST_DATA_ROW, 1, numRows, 2).getValues();
  var protCache = ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var itemdbRng = itemdbBoundedRangeA1_(ss);
  var emptyGearRows = [];
  for (var i = 0; i < numRows; i++) {
    var r = CE_FIRST_DATA_ROW + i;
    var gt = String(ab[i][0]).trim();
    if (normalizeItemName(gt) === CE_WEIGHT_ROW_LABEL) continue;
    if (!gt || gt.indexOf("---") === 0) continue;
    var item = normalizeItemName(ab[i][1]);
    if (item) {
      if (ceItemNameInItemDb_(ss, item)) {
        ceApplyItemDbStatsToRow_(ceSheet, r, ss, item);
        if (!ceStatRowHasOurProtection_(ceSheet, r, protCache)) {
          ceProtectStatRow_(ceSheet, r, protCache);
          protCache = ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        }
      } else {
        ceRemoveStatRowProtection_(ceSheet, r, protCache);
        protCache = ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      }
      recalculateGearScoreForRow_(ceSheet, r);
    } else {
      emptyGearRows.push(r);
    }
  }
  if (emptyGearRows.length > 0) {
    var es = {};
    for (var ei = 0; ei < emptyGearRows.length; ei++) es[String(emptyGearRows[ei])] = true;
    ceRemoveOurStatProtectionsForRows_(ceSheet, es);
    for (var ej = 0; ej < emptyGearRows.length; ej++) {
      ceClearEmptyGearSlotRow_(ceSheet, emptyGearRows[ej], ss, itemdbRng);
    }
  }
}

/** Gear rows with no item in B: clear gems / gear score / VLOOKUP (no full ceUpdate per row). */
function ceLockGemRowsWithoutItem_(ceSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lr = ceEffectiveLastRow_(ceSheet);
  if (lr < CE_FIRST_DATA_ROW) return;
  var numRows = lr - CE_FIRST_DATA_ROW + 1;
  var ab = ceSheet.getRange(CE_FIRST_DATA_ROW, 1, numRows, 2).getValues();
  var itemdbRng = itemdbBoundedRangeA1_(ss);
  var emptyGearRows = [];
  for (var i = 0; i < numRows; i++) {
    var r = CE_FIRST_DATA_ROW + i;
    var gt = String(ab[i][0]).trim();
    if (normalizeItemName(gt) === CE_WEIGHT_ROW_LABEL) continue;
    if (!gt || gt.indexOf("---") === 0) continue;
    if (normalizeItemName(ab[i][1])) continue;
    emptyGearRows.push(r);
  }
  if (emptyGearRows.length === 0) return;
  var es = {};
  for (var ei = 0; ei < emptyGearRows.length; ei++) es[String(emptyGearRows[ei])] = true;
  ceRemoveOurStatProtectionsForRows_(ceSheet, es);
  for (var ej = 0; ej < emptyGearRows.length; ej++) {
    ceClearEmptyGearSlotRow_(ceSheet, emptyGearRows[ej], ss, itemdbRng);
  }
}

/**
 * onOpen: rows with an item — re-sync ItemDB stats/protection, manual socket dropdowns, then scores.
 * Uses one getProtections snapshot + skip re-protect when already correct. Manual rows: stat zeros
 * scrubbed inside ceRecalculateManualIdealGemsAndScore_.
 */
function ceRefreshDerivedOnOpen_(ceSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lr = ceEffectiveLastRow_(ceSheet);
  if (lr < CE_FIRST_DATA_ROW) return;
  var numRows = lr - CE_FIRST_DATA_ROW + 1;
  var ab = ceSheet.getRange(CE_FIRST_DATA_ROW, 1, numRows, 2).getValues();
  var protCache = ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < numRows; i++) {
    var r = CE_FIRST_DATA_ROW + i;
    var gt = String(ab[i][0]).trim();
    if (normalizeItemName(gt) === CE_WEIGHT_ROW_LABEL) continue;
    if (!gt || gt.indexOf("---") === 0) continue;
    var item = normalizeItemName(ab[i][1]);
    if (!item) continue;
    var inDb = ceItemNameInItemDb_(ss, item);
    if (inDb) {
      ceApplyItemDbStatsToRow_(ceSheet, r, ss, item);
      if (!ceStatRowHasOurProtection_(ceSheet, r, protCache)) {
        ceProtectStatRow_(ceSheet, r, protCache);
        protCache = ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      }
    } else {
      ceRemoveStatRowProtection_(ceSheet, r, protCache);
      protCache = ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    }
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
    var gemColCount = CE_GEM_LAST_COL - CE_GEM_FIRST_COL + 1;
    ceSheet.getRange(1, CE_GEM_FIRST_COL, ceSheet.getMaxRows(), gemColCount).shiftColumnGroupDepth(1);
    PropertiesService.getDocumentProperties().setProperty("BIS_CE_GEM_GROUPED_V1", "1");
  } catch (eGrp) {}
}

function ceEnsureWeightsModeValidation_(ceSheet) {
  var lr = ceEffectiveLastRow_(ceSheet);
  if (lr < CE_FIRST_DATA_ROW) return;
  var numRows = lr - CE_FIRST_DATA_ROW + 1;
  var aVals = ceSheet.getRange(CE_FIRST_DATA_ROW, 1, numRows, 1).getValues();
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CE_WEIGHT_MODE_PAWN, CE_WEIGHT_MODE_CUSTOM], true)
    .setAllowInvalid(true)
    .build();
  var a1List = [];
  for (var wi = 0; wi < numRows; wi++) {
    if (normalizeItemName(aVals[wi][0]) === CE_WEIGHT_ROW_LABEL) {
      a1List.push(colLetterFromNum_(CE_ITEMNAME) + (CE_FIRST_DATA_ROW + wi));
    }
  }
  if (a1List.length === 0) return;
  try {
    ceSheet.getRangeList(a1List).setDataValidation(rule);
  } catch (eWl) {
    for (var wj = 0; wj < a1List.length; wj++) {
      ceSheet.getRange(a1List[wj]).setDataValidation(rule);
    }
  }
}

function refreshCEComparisonForRow_(ceSheet, bpSheet, row) {
  if (!bpSheet) return;
  var spec = findSpecForCERow(ceSheet, row);
  if (!spec) return;
  var ceSlotLabel = normalizeItemName(ceSheet.getRange(row, CE_GEARTYPE).getValue());
  if (!ceSlotLabel || ceSlotLabel.indexOf("---") === 0 || ceSlotLabel === CE_WEIGHT_ROW_LABEL) return;
  var plannerGear = plannerGearAndInterestFromCE_(ceSlotLabel).plannerGear;
  SpreadsheetApp.flush();
  Utilities.sleep(EDIT_RECALC_WAIT_MS);
  // Must pass gear type: spec-only refresh walks every row for that spec (can be 10k+ → minutes on onEdit).
  refreshComparisonRichText(bpSheet, spec, plannerGear, null, null, null, null, null);
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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bpSheet = ss.getSheetByName(BIS_PLANNER_SHEET);

  if (col >= CE_GEM_FIRST_COL && col <= CE_GEM_LAST_COL) {
    var itemGem = normalizeItemName(ceSheet.getRange(row, CE_ITEMNAME).getValue());
    if (!itemGem) {
      recalculateGearScoreForRow_(ceSheet, row);
    } else if (ceItemNameInItemDb_(ss, itemGem)) {
      recalculateGearScoreForRow_(ceSheet, row);
    } else {
      var wRowG = ceFindWeightsRowForSection_(ceSheet, row);
      if (wRowG) ceRecalculateManualIdealGemsAndScore_(ceSheet, row, ss, wRowG);
    }
    refreshCEComparisonForRow_(ceSheet, bpSheet, row);
    return;
  }

  if (col >= CE_STAT_FIRST_COL && col <= CE_STAT_LAST_COL) {
    var itemSt = normalizeItemName(ceSheet.getRange(row, CE_ITEMNAME).getValue());
    if (itemSt && ceItemNameInItemDb_(ss, itemSt)) {
      ceApplyItemDbStatsToRow_(ceSheet, row, ss, itemSt);
      refreshCEComparisonForRow_(ceSheet, bpSheet, row);
      return;
    }
    recalculateGearScoreForRow_(ceSheet, row);
    refreshCEComparisonForRow_(ceSheet, bpSheet, row);
    return;
  }

  if (col !== CE_ITEMNAME) return;

  var spec = findSpecForCERow(ceSheet, row);
  if (!spec) return;

  var newItemName = normalizeItemName(e.range.getValue());
  var oldItemName = normalizeItemName(e.oldValue);

  if (!newItemName) {
    ceRemoveStatRowProtection_(ceSheet, row);
    ceClearGemCellValidations_(ceSheet, row);
    ceRestoreStatLookupFormulasForRow_(ceSheet, row, ss);
    recalculateGearScoreForRow_(ceSheet, row);
    SpreadsheetApp.flush();
  } else {
    ceRefreshStatsAndGearScoreForRow_(ceSheet, row, oldItemName, newItemName, ss);
  }

  if (bpSheet) {
    if (oldItemName) {
      clearInterestForItem(bpSheet, spec, gearType, oldItemName);
    }
    if (newItemName) {
      setInterestForItem(bpSheet, spec, gearType, newItemName);
    }
  }

  refreshCEComparisonForRow_(ceSheet, bpSheet, row);
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

/**
 * After CE column B is set to the item (planner Equipped sync or user): apply ItemDB stats / manual gems,
 * recalculate gear score. Callers flush before the next sheet read (e.g. handleBISPlannerEdit, refreshCEComparisonForRow_).
 */
function ceRefreshStatsAndGearScoreForRow_(ceSheet, row, oldItemName, newItemName, ss) {
  var newN = normalizeItemName(newItemName);
  if (!newN) return;
  var oldN = normalizeItemName(oldItemName);
  var protCache = ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var newInDb = ceItemNameInItemDb_(ss, newN);
  if (newInDb) {
    ceApplyItemDbStatsToRow_(ceSheet, row, ss, newN);
    if (!ceStatRowHasOurProtection_(ceSheet, row, protCache)) {
      ceProtectStatRow_(ceSheet, row, protCache);
      protCache = ceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    }
    ceClearGemCellValidations_(ceSheet, row);
  } else {
    ceRemoveStatRowProtection_(ceSheet, row, protCache);
    if (oldN && ceItemNameInItemDb_(ss, oldN)) {
      ceClearStatCellsToBlank_(ceSheet, row);
    }
    ceSetupManualGemSocketDropdowns_(ceSheet, row);
  }
  recalculateGearScoreForRow_(ceSheet, row);
}

/**
 * @param {?GoogleAppsScript.Spreadsheet.Sheet} optCeSheet - pass from caller to avoid duplicate getSheetByName
 * @param {?Array<Array<*>>} optCeGearColVals - column A values rows 1..lastRow (from one getValues); skips full-column scan in findCERow
 */
function setCEItem(spec, gearType, itemName, optCeSheet, optCeGearColVals) {
  var ceSheet =
    optCeSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CURRENT_EQUIP_SHEET);
  if (!ceSheet) return;

  var ceRow = findCERow(ceSheet, spec, gearType, optCeGearColVals);
  if (ceRow) {
    ceSheet.getRange(ceRow, CE_ITEMNAME).setValue(normalizeItemName(itemName));
  }
}

function clearCEItemIfMatch(spec, gearType, itemName, optCeSheet, optCeGearColVals) {
  var ceSheet =
    optCeSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CURRENT_EQUIP_SHEET);
  if (!ceSheet) return;

  var ceRow = findCERow(ceSheet, spec, gearType, optCeGearColVals);
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
    var lastRow = ceEffectiveLastRow_(ceSheet);
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

/**
 * First occurrence of each (normalized spec, CE column A slot label) → 1-based row (same rules as findCERow).
 * Used to read stat columns from CE for Stat Comparison diffs (manual + ItemDB equipped items).
 */
function plannerCeSlotRowMapFromColA_(colA) {
  var map = {};
  if (!colA || colA.length === 0) return map;
  var curSpec = null;
  for (var i = 0; i < colA.length; i++) {
    var val = String(colA[i][0]);
    if (val.indexOf("---") === 0) {
      var parsedHdr = specDisplayNameFromSectionHeader(val);
      curSpec = parsedHdr ? normalizeItemName(parsedHdr) : null;
      continue;
    }
    if (!curSpec) continue;
    var gt = normalizeItemName(val);
    if (!gt || val.indexOf("---") === 0) continue;
    if (gt === normalizeItemName(CE_WEIGHT_ROW_LABEL)) continue;
    var key = curSpec + "\x1e" + gt;
    if (!(key in map)) {
      map[key] = i + 1;
    }
  }
  return map;
}

/** CE stat block row (1-based) → numeric array length nStat; null if out of range. */
function ceStatNumericArrayFromRowMatrix_(statMatrix, row1Based, nStat) {
  if (!statMatrix || row1Based < 1 || row1Based > statMatrix.length) return null;
  var row = statMatrix[row1Based - 1];
  if (!row || row.length < nStat) return null;
  var out = [];
  for (var si = 0; si < nStat; si++) {
    var v = row[si];
    if (v == null || v === "") {
      out.push(0);
    } else {
      var n = Number(v);
      out.push(isNaN(n) ? 0 : n);
    }
  }
  return out;
}

function plannerCeStatsForSpecSlot_(slotMap, statMatrix, spec, ceSlotLabel, nStat) {
  if (!slotMap || !statMatrix || !ceSlotLabel) return null;
  var key = normalizeItemName(spec) + "\x1e" + normalizeItemName(ceSlotLabel);
  var row = slotMap[key];
  if (!row) return null;
  return ceStatNumericArrayFromRowMatrix_(statMatrix, row, nStat);
}

/** In-memory CE column A snapshot; same matching rules as findCERow. */
function findCERowInSnapshot_(colA, spec, gearType) {
  if (!colA || colA.length === 0) return null;
  var specWant = normalizeItemName(spec);
  var gearWant = normalizeItemName(gearType);
  if (!gearWant) return null;
  var curSpec = null;
  for (var i = 0; i < colA.length; i++) {
    var val = String(colA[i][0]);
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

/** One A+B bulk read for equip resolution (full-sheet comparison). */
function plannerBuildCeEquipSnapshot_(ceSheet, ceGearColVals) {
  if (!ceSheet) {
    return { colA: [], colB: [], celr: 0 };
  }
  var celr = ceEffectiveLastRow_(ceSheet);
  if (celr < 1) {
    return { colA: [], colB: [], celr: 0 };
  }
  var colAFromSnap = ceGearColSnapshotRowsAtCelr_(ceGearColVals, celr);
  var colA =
    colAFromSnap != null
      ? colAFromSnap
      : ceSheet.getRange(1, CE_GEARTYPE, celr, 1).getValues();
  var colB = ceSheet.getRange(1, CE_ITEMNAME, celr, 1).getValues();
  return { colA: colA, colB: colB, celr: celr };
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

  var rowsSetEquipped = [];
  var rowsClearInterest = [];
  for (var i = 0; i < data.length; i++) {
    var r = i + BIS_FIRST_DATA_ROW;
    if (normalizeItemName(data[i][GG_SPEC - 1]) !== normalizeItemName(spec)) continue;
    if (normalizeItemName(data[i][GG_GEARTYPE - 1]) !== normalizeItemName(pGear)) continue;
    var curI = data[i][GG_INTEREST - 1];
    if (itemsNameMatch_(data[i][GG_NAME - 1], want)) {
      if (normalizeItemName(curI) !== normalizeItemName(wantInterest)) {
        rowsSetEquipped.push(r);
      }
    } else if (
      interestIsEquippedState_(curI) &&
      (!gearIsRingOrTrinket_(pGear) ||
        ringTrinketSlotBucket_(curI) === ringTrinketSlotBucket_(wantInterest))
    ) {
      rowsClearInterest.push(r);
    }
  }
  if (rowsSetEquipped.length > 0) {
    var rsEq = rowsSetEquipped.map(function (x) {
      return "A" + x;
    });
    var rlEq = bpSheet.getRangeList(rsEq);
    rlEq.setValue(wantInterest);
    rlEq.setFontWeight("bold");
  }
  if (rowsClearInterest.length > 0) {
    var rsCl = rowsClearInterest.map(function (x) {
      return "A" + x;
    });
    var rlCl = bpSheet.getRangeList(rsCl);
    rlCl.setValue("");
    rlCl.setFontWeight("normal");
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

  var rowsClear = [];
  for (var i = 0; i < data.length; i++) {
    var r = i + BIS_FIRST_DATA_ROW;
    if (normalizeItemName(data[i][GG_SPEC - 1]) !== normalizeItemName(spec)) continue;
    if (normalizeItemName(data[i][GG_GEARTYPE - 1]) !== normalizeItemName(pGear)) continue;
    if (!itemsNameMatch_(data[i][GG_NAME - 1], want)) continue;
    if (normalizeItemName(data[i][GG_INTEREST - 1]) !== normalizeItemName(wantInterest)) continue;
    rowsClear.push(r);
  }
  if (rowsClear.length > 0) {
    var rs = rowsClear.map(function (x) {
      return "A" + x;
    });
    var rl = bpSheet.getRangeList(rs);
    rl.setValue("");
    rl.setFontWeight("normal");
  }
  syncEquippedIndexForSpecGearFromPlannerData_(spec, pGear, data);
}

// ============================================================
// CmpRaw (column P): script-built plain text (replaces sheet Δ helpers + CmpRaw formulas)
// ============================================================

/**
 * @param {Array<Array<*>>} data - ItemDB grid including Attributes column (see itemdbReadAllStatRows_).
 * @return {Object<string, Array<*>>} lowercased trimmed name → row array
 */
function plannerItemdbNameToRow_(data) {
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var k = normalizeItemName(data[i][0]);
    if (!k) continue;
    map[k.toLowerCase()] = data[i];
  }
  return map;
}

/** @param {?Array<*>} rowArr @param {number} statIndex 0..nStat-1 */
function plannerItemdbStatNumericForDiff_(rowArr, statIndex) {
  if (!rowArr) return 0;
  var v = rowArr[ITEMDB_FIRST_STAT_COL - 1 + statIndex];
  if (v == null || v === "") return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

/** Attributes column (same row); nStat = number of stat columns in ItemDB. */
function plannerItemdbAttrString_(rowArr, nStat) {
  if (!rowArr) return "";
  var a = rowArr[ITEMDB_FIRST_STAT_COL - 1 + nStat];
  if (a == null) return "";
  return String(a);
}

/** Mirrors generate_bis_planner.py _cmp_attr_display_expr (SUBSTITUTE Use:/Equip:). */
function plannerCmpAttrDisplayFromRaw_(raw) {
  var t = raw == null ? "" : String(raw).replace(/^\s+|\s+$/g, "");
  if (!t) return "";
  t = t.replace(/ Use:/g, "\nUse:");
  t = t.replace(/ Equip:/g, "\nEquip:");
  return t;
}

/**
 * TEXTJOIN(", ",TRUE,…) over per-stat diff fragments for one equip slot.
 * Mirrors _build_stat_diff_helper_formula × nStat.
 * @param {?Array<number>} optEquipCeStats — CE stat columns for that equipped slot (manual or DB); when null/short, equipped side uses ItemDB only.
 */
function plannerStatDiffCommaJoin_(rowMap, itemName, equipName, nStat, optEquipCeStats) {
  if (normalizeItemName(equipName) === "") return "";
  var wantItem = normalizeItemName(itemName);
  var wantEq = normalizeItemName(equipName);
  var rowI = wantItem ? rowMap[wantItem.toLowerCase()] : null;
  var rowE = wantEq ? rowMap[wantEq.toLowerCase()] : null;
  var useCeEquip =
    optEquipCeStats != null && optEquipCeStats.length >= nStat;
  var parts = [];
  for (var si = 0; si < nStat; si++) {
    var itemV = plannerItemdbStatNumericForDiff_(rowI, si);
    var equipV = useCeEquip
      ? Number(optEquipCeStats[si]) || 0
      : plannerItemdbStatNumericForDiff_(rowE, si);
    var diff = itemV - equipV;
    if (diff === 0) continue;
    var rounded = Math.round(Math.abs(diff) * 10000) / 10000;
    var sign = diff > 0 ? "+" : "-";
    var lab = BIS_STAT_LABELS[si];
    if (!lab) continue;
    parts.push(sign + rounded + " " + lab);
  }
  return parts.length === 0 ? "" : parts.join(", ");
}

/**
 * Inner block for one slot: stats line + optional equipped attributes.
 * @param {string} equippedHeading - e.g. "Equipped slot 1:", "Equipped slot 2", "Currently Equipped:"
 */
function plannerCmpSlotInner_(statJoin, attrRaw, equippedHeading) {
  var tj = statJoin == null ? "" : String(statJoin).replace(/^\s+|\s+$/g, "");
  var va = attrRaw == null ? "" : String(attrRaw).replace(/^\s+|\s+$/g, "");
  if (tj === "" && va === "") return "";
  if (va === "") return statJoin == null ? "" : String(statJoin);
  var ad = plannerCmpAttrDisplayFromRaw_(attrRaw);
  var la = equippedHeading + "\n" + ad;
  if (tj === "") return la;
  return String(statJoin) + "\n" + la;
}

/**
 * Full CmpRaw string for one planner row (matches generate_bis_planner.py _build_cmp_raw_formula).
 * @param {?Array<number>} optCeStats1 — CE stat row for slot 1 / single-slot gear (when set, stat diff uses CE for equipped side).
 * @param {?Array<number>} optCeStats2 — CE stat row for ring/trinket slot 2.
 */
function plannerBuildCmpRawStringFromRowMap_(
  gearType,
  itemName,
  equipName,
  equip2Name,
  rowMap,
  nStat,
  optCeStats1,
  optCeStats2
) {
  var g = normalizeItemName(gearType);
  var nm = itemName;
  var eq1 = equipName;
  var eq2 = equip2Name;
  var rowItem = normalizeItemName(nm) ? rowMap[normalizeItemName(nm).toLowerCase()] : null;
  var wantE1 = normalizeItemName(eq1);
  var wantE2 = normalizeItemName(eq2);
  var rowE1 = wantE1 ? rowMap[wantE1.toLowerCase()] : null;
  var rowE2 = wantE2 ? rowMap[wantE2.toLowerCase()] : null;
  var va1 = plannerItemdbAttrString_(rowE1, nStat);
  var va2 = plannerItemdbAttrString_(rowE2, nStat);
  var tj1 = plannerStatDiffCommaJoin_(rowMap, nm, eq1, nStat, optCeStats1);
  var tj2 = plannerStatDiffCommaJoin_(rowMap, nm, eq2, nStat, optCeStats2);

  if (g === "Ring" || g === "Trinket") {
    var inner1 = plannerCmpSlotInner_(tj1, va1, "Equipped slot 1:");
    var inner2 = plannerCmpSlotInner_(tj2, va2, "Equipped slot 2");
    var same1 = itemsNameMatch_(nm, eq1);
    var same2 = itemsNameMatch_(nm, eq2);
    var b1 = wantE1 === "" || same1 ? "" : "Slot 1 Comparison:\n" + inner1;
    var b2 = wantE2 === "" || same2 ? "" : "Slot 2 comparison\n" + inner2;
    if (wantE1 === "" && wantE2 === "") {
      return CMP_RAW_EQUIP_NOT_SPECIFIED;
    }
    if (b1 === "" && b2 === "") return "";
    if (b1 === "") return b2;
    if (b2 === "") return b1;
    return b1 + "\n" + b2;
  }

  var tjS = tj1;
  var vaS = va1;
  var innerS = plannerCmpSlotInner_(tjS, vaS, "Currently Equipped:");
  var sameS = itemsNameMatch_(nm, eq1);
  if (wantE1 === "") return CMP_RAW_EQUIP_NOT_SPECIFIED;
  if (sameS) return "";
  if (innerS === "") return "";
  return innerS;
}

/**
 * Builds CmpRaw text into blockDisplay / rowNK only (layout v2: no CmpRaw sheet column).
 * @param {?GoogleAppsScript.Spreadsheet.Sheet} optCeSheet — when set with optCeGearColVals, equipped-side stat diffs use CE stat columns (manual items).
 * @param {?Array<Array<*>>} optCeGearColVals — CE column A rows 1..last (may be re-read if shorter than effective last row).
 */
function plannerWriteCmpRawForComparisonRefresh_(
  bpSheet,
  indices,
  blockDisplay,
  rowNK,
  offNInBlock,
  optItemdbData,
  optCeSheet,
  optCeGearColVals
) {
  var ss = bpSheet.getParent();
  var data;
  if (optItemdbData != null) {
    data = optItemdbData;
  } else {
    data = itemdbReadAllStatRows_(ss);
  }
  if (!data || data.length === 0) return;
  var nStat = CE_STAT_LAST_COL - CE_STAT_FIRST_COL + 1;
  if (BIS_STAT_LABELS.length !== nStat) return;
  var rowMap = plannerItemdbNameToRow_(data);
  var ceSlotMap = null;
  var ceStatMatrix = null;
  if (optCeSheet) {
    var celrW = ceEffectiveLastRow_(optCeSheet);
    if (celrW >= 1) {
      var colAForMap = optCeGearColVals;
      if (colAForMap == null || colAForMap.length < celrW) {
        colAForMap = optCeSheet.getRange(1, CE_GEARTYPE, celrW, 1).getValues();
      }
      ceSlotMap = plannerCeSlotRowMapFromColA_(colAForMap);
      ceStatMatrix = optCeSheet
        .getRange(1, CE_STAT_FIRST_COL, celrW, CE_STAT_LAST_COL)
        .getValues();
    }
  }
  for (var ii = 0; ii < indices.length; ii++) {
    var i = indices[ii];
    var sp;
    var g;
    var n;
    var e1;
    var e2;
    if (blockDisplay != null) {
      sp = blockDisplay[i][GG_SPEC - GG_SPEC];
      g = blockDisplay[i][GG_GEARTYPE - GG_SPEC];
      n = blockDisplay[i][GG_NAME - GG_SPEC];
      e1 = blockDisplay[i][GG_EQUIP - GG_SPEC];
      e2 = blockDisplay[i][GG_EQUIP2 - GG_SPEC];
    } else {
      var nk = rowNK[i];
      if (!nk) continue;
      sp = nk.sp;
      g = nk.g;
      n = nk.n;
      e1 = nk.e;
      e2 = nk.e2;
    }
    var gNorm = normalizeItemName(g);
    var stats1 = null;
    var stats2 = null;
    if (ceSlotMap && ceStatMatrix) {
      if (gNorm === "Ring") {
        stats1 = plannerCeStatsForSpecSlot_(ceSlotMap, ceStatMatrix, sp, "Ring 1", nStat);
        stats2 = plannerCeStatsForSpecSlot_(ceSlotMap, ceStatMatrix, sp, "Ring 2", nStat);
      } else if (gNorm === "Trinket") {
        stats1 = plannerCeStatsForSpecSlot_(ceSlotMap, ceStatMatrix, sp, "Trinket 1", nStat);
        stats2 = plannerCeStatsForSpecSlot_(ceSlotMap, ceStatMatrix, sp, "Trinket 2", nStat);
      } else if (gNorm) {
        stats1 = plannerCeStatsForSpecSlot_(ceSlotMap, ceStatMatrix, sp, gNorm, nStat);
      }
    }
    var s = plannerBuildCmpRawStringFromRowMap_(g, n, e1, e2, rowMap, nStat, stats1, stats2);
    if (blockDisplay != null) {
      blockDisplay[i][offNInBlock] = s;
    } else {
      rowNK[i].d = s;
    }
  }
}

// ============================================================
// Stat Comparison column K: Rich Text from script-built CmpRaw text (no CmpRaw sheet column in layout v2).
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

/** True if K mirrors legacy column P CmpRaw (=RC[offset] or =Prow). */
function cellIsCmpRawMirrorFormula_(formula, row) {
  if (formula == null || formula === "") return false;
  var f = String(formula).replace(/^\s+|\s+$/g, "");
  if (new RegExp("RC\\s*\\[\\s*" + CMP_RAW_R1C1_OFFSET + "\\s*\\]").test(f)) return true;
  var L = colLetterFromNum_(CMP_RAW_LEGACY_SHEET_COL);
  var re = new RegExp("^=\\s*\\$?" + L + "\\s*\\$?" + row + "\\s*$", "i");
  return re.test(f);
}

/**
 * Writes queued Stat Comparison (K) updates in batches (clear + setRichTextValues / setFormulas) to cut round-trips.
 * @param {Array<{r:number, kind:string, rich?:GoogleAppsScript.Spreadsheet.RichTextValue, needClearBefore:boolean}>} writes
 */
function flushComparisonWrites_(bpSheet, writes, cmpRawColLetter) {
  if (!writes || writes.length === 0) {
    return;
  }

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
        if (cmpRawColLetter) {
          var fmatrix = [];
          for (var tm = j; tm <= mEnd; tm++) {
            fmatrix.push(["=" + cmpRawColLetter + seg[tm].r]);
          }
          bpSheet.getRange(r0m, CMP_DISP_COL, hm, 1).setFormulas(fmatrix);
        }
        j = mEnd + 1;
      }
    }
  }
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
 * Equipped item name(s) from Current Equipment for this planner row (matches former N/O formulas).
 * @param {?{colA:Array<Array<*>>, colB:Array<Array<*>>, celr:number}} optEquipSnap - from plannerBuildCeEquipSnapshot_; avoids per-row CE reads.
 */
function plannerResolveEquippedNamesForPlannerRow_(ceSheet, specRaw, gearTypeRaw, ceGearColVals, optEquipSnap) {
  if (!ceSheet) return { e1: "", e2: "" };
  var sp = normalizeItemName(specRaw);
  var g = normalizeItemName(gearTypeRaw);
  if (!sp || !g) return { e1: "", e2: "" };
  var snap = optEquipSnap;
  var colA = snap ? snap.colA : ceGearColVals;
  var colB = snap ? snap.colB : null;
  function nameAtRow(r) {
    if (colB && r >= 1 && r <= colB.length) {
      return normalizeItemName(colB[r - 1][0]);
    }
    return normalizeItemName(ceSheet.getRange(r, CE_ITEMNAME).getValue());
  }
  if (g === "Ring" || g === "Trinket") {
    var s1 = g === "Ring" ? "Ring 1" : "Trinket 1";
    var s2 = g === "Ring" ? "Ring 2" : "Trinket 2";
    var r1 = snap ? findCERowInSnapshot_(colA, sp, s1) : findCERow(ceSheet, sp, s1, ceGearColVals);
    var r2 = snap ? findCERowInSnapshot_(colA, sp, s2) : findCERow(ceSheet, sp, s2, ceGearColVals);
    var e1 = r1 ? nameAtRow(r1) : "";
    var e2 = r2 ? nameAtRow(r2) : "";
    return { e1: e1, e2: e2 };
  }
  var r0 = snap ? findCERowInSnapshot_(colA, sp, g) : findCERow(ceSheet, sp, g, ceGearColVals);
  var e0 = r0 ? nameAtRow(r0) : "";
  return { e1: e0, e2: "" };
}

/**
 * Full-sheet refresh: logical B–P block from A–D + F–I bands + CE-resolved equip (no N–O on sheet).
 * Equip path uses spec+gear scoped refreshComparisonRichText (not full-sheet); onOpen/menu use this for all rows.
 */
function plannerAssembleFullSheetBlockFromBandsAndCE_(partAD, partFI, ceSheet, ceGearColVals) {
  var n = partAD.length;
  var snap = plannerBuildCeEquipSnapshot_(ceSheet, ceGearColVals);
  var blockDisplay = new Array(n);
  for (var ri = 0; ri < n; ri++) {
    var ad = partAD[ri];
    var fi = partFI[ri];
    var eqPair = plannerResolveEquippedNamesForPlannerRow_(ceSheet, ad[1], ad[2], ceGearColVals, snap);
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
      eqPair.e1,
      eqPair.e2,
      "",
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
 * @param {?Array<Array<*>>} plannerABCReuse - optional B:C (2 cols) or A:C (3 cols) getValues snapshot, rows aligned
 *   to BIS_FIRST_DATA_ROW; when spec+gear scoped, skips a second B+C read (equip path passes B:C only).
 * @param {?Object} optSharedPlannerCaches - from plannerRefreshCachesBuild_; when set (e.g. force refresh), skip rebuilding
 *   ItemDB/GemDB/CE caches so pass 2 reuses pass 1 memoization.
 * @param {?{captureFullSheetBands:?Object, reuseFullSheetBands:?Object}} optFullSheetIo_ - force-refresh pass 2 only:
 *   pass 1 fills captureFullSheetBands with partAD/partFI; pass 2 reuses them and refreshes CE-derived equip + K formulas.
 * @param {?Array<Array<*>>} optCeGearColVals - CE column A getValues (rows 1..effective last row); skips duplicate read in cache build.
 * @param {?number} optPlannerLastRow - when >= BIS_FIRST_DATA_ROW, skip getLastRow (equip path passes fresh lastRow).
 * @param {boolean=} optPlainComparisonDisp - full-sheet only: write plain text to K/L in bulk (fast onOpen fallback); omit for colored Rich Text.
 */
function refreshComparisonRichText(
  bpSheet,
  specFilter,
  gearTypeFilter,
  plannerABCReuse,
  optSharedPlannerCaches,
  optFullSheetIo_,
  optCeGearColVals,
  optPlannerLastRow,
  optPlainComparisonDisp
) {
  var lastRow;
  if (optPlannerLastRow != null && Number(optPlannerLastRow) >= BIS_FIRST_DATA_ROW) {
    lastRow = Number(optPlannerLastRow);
  } else {
    lastRow = bpSheet.getLastRow();
  }
  if (lastRow < BIS_FIRST_DATA_ROW) return;

  if (bisRefreshReentryDepth_ > 0) {
    return;
  }
  bisRefreshReentryDepth_++;
  try {
    refreshComparisonRichTextInner_(
      bpSheet,
      specFilter,
      gearTypeFilter,
      plannerABCReuse,
      lastRow,
      optSharedPlannerCaches,
      optFullSheetIo_,
      optCeGearColVals,
      optPlainComparisonDisp
    );
  } finally {
    bisRefreshReentryDepth_--;
  }
}

function refreshComparisonRichTextInner_(
  bpSheet,
  specFilter,
  gearTypeFilter,
  plannerABCReuse,
  lastRow,
  optSharedPlannerCaches,
  optFullSheetIo_,
  optCeGearColVals,
  optPlainComparisonDisp
) {
  var specFilterNorm =
    specFilter != null && String(specFilter).replace(/^\s+|\s+$/g, "") !== ""
      ? normalizeItemName(specFilter)
      : null;
  var gearFilterNorm =
    gearTypeFilter != null && String(gearTypeFilter).replace(/^\s+|\s+$/g, "") !== ""
      ? normalizeItemName(gearTypeFilter)
      : null;

  var ssInner = bpSheet.getParent();
  var ceSheetInner = ssInner.getSheetByName(CURRENT_EQUIP_SHEET);
  var ceGearColInner = null;
  if (ceSheetInner) {
    var celrInner = ceEffectiveLastRow_(ceSheetInner);
    var innerSnap = ceGearColSnapshotRowsAtCelr_(optCeGearColVals, celrInner);
    if (innerSnap != null) {
      ceGearColInner = innerSnap;
    } else if (celrInner >= 1) {
      ceGearColInner = ceSheetInner.getRange(1, CE_GEARTYPE, celrInner, 1).getValues();
    }
    // After a cold flush, getLastRow() can tick up on the next read; avoid a too-short matrix for rowNK/snap.
    var celrPost = ceEffectiveLastRow_(ceSheetInner);
    if (
      celrPost >= 1 &&
      (ceGearColInner == null || ceGearColInner.length < celrPost)
    ) {
      ceGearColInner = ceSheetInner.getRange(1, CE_GEARTYPE, celrPost, 1).getValues();
    }
  }
  var ceEquipSnap = ceSheetInner
    ? plannerBuildCeEquipSnapshot_(ceSheetInner, ceGearColInner)
    : null;

  // Sheet.getRange(row, col, numRows, numColumns) — 3rd/4th are counts, not end row/col.
  var cmpNumRows = Math.max(0, lastRow - BIS_FIRST_DATA_ROW + 1);
  var interestAll = [];
  var cmpRawLetter = null;
  var offNInBlock = CMP_RAW_COL - GG_SPEC;
  var indices = [];
  var blockDisplay = null;
  var kFormulas = null;
  var rowNK = null;
  /**
   * F:I per row for heroic/dungeon styling when blockDisplay is null.
   * Full-sheet path uses a 2D array (all rows); scoped path uses sparse object keyed by 0-based data index.
   */
  var metaAcqDungeonDiff = null;

  if (!specFilterNorm && !gearFilterNorm) {
    var rData = BIS_FIRST_DATA_ROW;
    var wAD = GG_NAME - GG_INTEREST + 1;
    var wFI = GG_DIFFICULTY - GG_ACQ + 1;
    if (cmpNumRows > 0) {
      var partAD;
      var partFI;
      var snap =
        optFullSheetIo_ && optFullSheetIo_.reuseFullSheetBands
          ? optFullSheetIo_.reuseFullSheetBands
          : null;
      if (
        snap &&
        snap.partAD &&
        snap.partAD.length === cmpNumRows &&
        snap.partFI &&
        snap.partFI.length === cmpNumRows
      ) {
        partAD = snap.partAD;
        partFI = snap.partFI;
      } else {
        partAD = bpSheet.getRange(rData, GG_INTEREST, cmpNumRows, wAD).getValues();
        partFI = bpSheet.getRange(rData, GG_ACQ, cmpNumRows, wFI).getValues();
      }
      blockDisplay = plannerAssembleFullSheetBlockFromBandsAndCE_(
        partAD,
        partFI,
        ceSheetInner,
        ceGearColInner
      );
      interestAll = new Array(cmpNumRows);
      for (var ai = 0; ai < cmpNumRows; ai++) {
        interestAll[ai] = [partAD[ai][0]];
      }
      kFormulas = bpSheet.getRange(rData, CMP_DISP_COL, cmpNumRows, 1).getFormulas();
      if (optFullSheetIo_ && optFullSheetIo_.captureFullSheetBands) {
        var cap = optFullSheetIo_.captureFullSheetBands;
        cap.partAD = partAD;
        cap.partFI = partFI;
      }
    }
    for (var aj = 0; aj < cmpNumRows; aj++) indices.push(aj);
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
      return;
    }
    interestAll = new Array(cmpNumRows);
    var scopedRunsSG = bisIndicesToRuns_(indices);
    for (var ir = 0; ir < scopedRunsSG.length; ir++) {
      var sInt = scopedRunsSG[ir][0];
      var pInt = scopedRunsSG[ir][1];
      var hInt = pInt - sInt + 1;
      var r0Int = BIS_FIRST_DATA_ROW + sInt;
      var intBlock = bpSheet.getRange(r0Int, GG_INTEREST, hInt, 1).getDisplayValues();
      for (var ji = 0; ji < hInt; ji++) {
        interestAll[sInt + ji] = intBlock[ji];
      }
    }
    rowNK = {};
    for (var gi = 0; gi < scopedRunsSG.length; gi++) {
      var sg = scopedRunsSG[gi][0];
      var pg = scopedRunsSG[gi][1];
      var hg = pg - sg + 1;
      var r0g = BIS_FIRST_DATA_ROW + sg;
      var kG = bpSheet.getRange(r0g, CMP_DISP_COL, hg, 1).getFormulas();
      var nameG = bpSheet.getRange(r0g, GG_NAME, hg, 1).getDisplayValues();
      var gearG = bpSheet.getRange(r0g, GG_GEARTYPE, hg, 1).getDisplayValues();
      var specG = bpSheet.getRange(r0g, GG_SPEC, hg, 1).getDisplayValues();
      for (var jg = 0; jg < hg; jg++) {
        var eqRg = plannerResolveEquippedNamesForPlannerRow_(
          ceSheetInner,
          specG[jg][0],
          gearG[jg][0],
          ceGearColInner,
          ceEquipSnap
        );
        rowNK[sg + jg] = {
          d: "",
          k: kG[jg][0],
          n: nameG[jg][0],
          e: eqRg.e1,
          e2: eqRg.e2,
          g: gearG[jg][0],
          sp: specG[jg][0],
        };
      }
    }
  } else {
    var bOnly = bpSheet.getRange(BIS_FIRST_DATA_ROW, GG_SPEC, cmpNumRows, 1).getValues();
    for (var ci = 0; ci < bOnly.length; ci++) {
      if (normalizeItemName(bOnly[ci][0]) !== specFilterNorm) continue;
      indices.push(ci);
    }
    if (indices.length === 0) {
      return;
    }
    interestAll = new Array(cmpNumRows);
    var scopedRunsSO = bisIndicesToRuns_(indices);
    for (var ir2 = 0; ir2 < scopedRunsSO.length; ir2++) {
      var sInt2 = scopedRunsSO[ir2][0];
      var pInt2 = scopedRunsSO[ir2][1];
      var hInt2 = pInt2 - sInt2 + 1;
      var r0Int2 = BIS_FIRST_DATA_ROW + sInt2;
      var intBlock2 = bpSheet.getRange(r0Int2, GG_INTEREST, hInt2, 1).getDisplayValues();
      for (var ji2 = 0; ji2 < hInt2; ji2++) {
        interestAll[sInt2 + ji2] = intBlock2[ji2];
      }
    }
    rowNK = {};
    for (var si2 = 0; si2 < scopedRunsSO.length; si2++) {
      var ix0 = scopedRunsSO[si2][0];
      var ps = scopedRunsSO[si2][1];
      var hs = ps - ix0 + 1;
      var r0s = BIS_FIRST_DATA_ROW + ix0;
      var kS = bpSheet.getRange(r0s, CMP_DISP_COL, hs, 1).getFormulas();
      var nameS = bpSheet.getRange(r0s, GG_NAME, hs, 1).getDisplayValues();
      var gearS = bpSheet.getRange(r0s, GG_GEARTYPE, hs, 1).getDisplayValues();
      var specS = bpSheet.getRange(r0s, GG_SPEC, hs, 1).getDisplayValues();
      for (var js = 0; js < hs; js++) {
        var eqRs = plannerResolveEquippedNamesForPlannerRow_(
          ceSheetInner,
          specS[js][0],
          gearS[js][0],
          ceGearColInner,
          ceEquipSnap
        );
        rowNK[ix0 + js] = {
          d: "",
          k: kS[js][0],
          n: nameS[js][0],
          e: eqRs.e1,
          e2: eqRs.e2,
          g: gearS[js][0],
          sp: specS[js][0],
        };
      }
    }
  }

  var threadItemdb = null;
  var threadGem = null;
  if (indices.length > 0) {
    threadItemdb = itemdbReadAllStatRows_(ssInner);
    plannerWriteCmpRawForComparisonRefresh_(
      bpSheet,
      indices,
      blockDisplay,
      rowNK,
      offNInBlock,
      threadItemdb,
      ceSheetInner,
      ceGearColInner
    );
    threadGem = gemdbReadAllRows_(ssInner);
  }


  if (blockDisplay == null && indices.length > 0) {
    var metaW = GG_DIFFICULTY - GG_ACQ + 1;
    var runsFi = bisIndicesToRuns_(indices);
    metaAcqDungeonDiff = {};
    for (var fi = 0; fi < runsFi.length; fi++) {
      var sf = runsFi[fi][0];
      var pf = runsFi[fi][1];
      var hf = pf - sf + 1;
      var r0f = BIS_FIRST_DATA_ROW + sf;
      var fiRows = bpSheet.getRange(r0f, GG_ACQ, hf, metaW).getDisplayValues();
      for (var jf = 0; jf < hf; jf++) {
        metaAcqDungeonDiff[sf + jf] = fiRows[jf];
      }
    }
  }

  var writes = [];
  var upgradeWrites = [];
  var ss = bpSheet.getParent();
  var ceSheet = ss.getSheetByName(CURRENT_EQUIP_SHEET);
  var plannerCaches = null;
  if (ceSheet) {
    if (optSharedPlannerCaches) {
      plannerCaches = optSharedPlannerCaches;
    } else {
      // Use ceGearColInner (resolved above), not optCeGearColVals: after CE writes in the same run,
      // effective last row can grow so the caller snapshot is too short — inner re-reads col A, but
      // passing the stale opt here forced a second full read in plannerRefreshCachesBuild_.
      plannerCaches = plannerRefreshCachesBuild_(
        ss,
        ceSheet,
        ceGearColInner,
        threadItemdb,
        threadGem
      );
    }
  }

  var usePlainFullSheet =
    optPlainComparisonDisp === true &&
    !specFilterNorm &&
    !gearFilterNorm &&
    cmpNumRows > 0;
  var kPlainMat = null;
  var lPlainTexts = null;
  var lPlainColors = null;
  if (usePlainFullSheet) {
    kPlainMat = [];
    lPlainTexts = [];
    lPlainColors = [];
    for (var pi = 0; pi < cmpNumRows; pi++) {
      kPlainMat.push([""]);
      lPlainTexts.push([""]);
      lPlainColors.push(["#000000"]);
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
        !usePlainFullSheet &&
        reuseBandSnap &&
        reuseBandSnap.cmpSkipFp &&
        Object.prototype.hasOwnProperty.call(reuseBandSnap.cmpSkipFp, r) &&
        reuseBandSnap.cmpSkipFp[r] === fpKey
      ) {
        continue;
      }
    }

    var itemFullNum = "";
    if (ceSheet && normalizeItemName(specRow) && normalizeItemName(nmN)) {
      itemFullNum = plannerItemGearScoresBothWithCaches_(plannerCaches, specRow, nmN).full;
    }
    var gearNorm = normalizeItemName(gearRow);
    var baselineEq = 0;
    if (ceSheet) {
      baselineEq = plannerBaselineEquippedScore_(ss, ceSheet, specRow, gearRow, eqN, eq2N, plannerCaches);
    }

    // Do not skip when N display is empty: leaving K unchanged preserves stale rich text (e.g. after
    // equipping). Empty CmpRaw → build returns null → mirror =RC[] so K shows blank until N fills, then
    // the sheet updates the mirror live; onOpen / second pass still reapplies colored rich text.
    // CmpRaw display errors (#VALUE!, #N/A, etc.) — mirror CmpRaw, no rich text
    if (dispStr.length > 0 && dispStr.charAt(0) === "#") {
      if (capBandSnap && fpKey != null) {
        if (!capBandSnap.cmpSkipFp) capBandSnap.cmpSkipFp = {};
        capBandSnap.cmpSkipFp[r] = fpKey;
      }
      if (usePlainFullSheet) {
        kPlainMat[i] = [dispStr];
        lPlainTexts[i] = [""];
        lPlainColors[i] = ["#000000"];
      } else {
        if (!cellIsCmpRawMirrorFormula_(kF, r)) {
          writes.push({ r: r, kind: "mirror", needClearBefore: true });
        }
        upgradeWrites.push({
          r: r,
          text: "",
          color: "#000000",
        });
      }
      continue;
    }
    if (usePlainFullSheet) {
      kPlainMat[i] = [normalizeComparisonDisplay(dispStr)];
    } else {
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
    }
    var hidePct =
      equippedHide ||
      normalizeItemName(nmN) === "" ||
      !normalizeItemName(specRow) ||
      !ceSheet;
    if (!hidePct && (gearNorm === "Ring" || gearNorm === "Trinket") && ceSheet && plannerCaches) {
      var s1s2 =
        gearNorm === "Ring"
          ? plannerRingEquippedSlotScores_(ceSheet, specRow, plannerCaches)
          : plannerTrinketEquippedSlotScores_(ceSheet, specRow, plannerCaches);
      if (usePlainFullSheet) {
        lPlainTexts[i] = [
          buildDualSlotUpgradePlainText_(
            Number(itemFullNum) || 0,
            s1s2.s1,
            s1s2.s2,
            s1s2.hasItem1,
            s1s2.hasItem2
          ),
        ];
        lPlainColors[i] = ["#000000"];
      } else {
        var dualRich = buildDualSlotUpgradeRichText_(
          Number(itemFullNum) || 0,
          s1s2.s1,
          s1s2.s2,
          s1s2.hasItem1,
          s1s2.hasItem2,
          heroicDarkRow
        );
        if (dualRich) {
          upgradeWrites.push({ r: r, rich: dualRich });
        } else {
          upgradeWrites.push({ r: r, text: "", color: "#000000" });
        }
      }
    } else {
      var ceHasEq =
        ceSheet && plannerCaches
          ? plannerCeHasEquippedItemForPlannerGear_(ceSheet, specRow, gearNorm, plannerCaches)
          : false;
      var uh = upgradePctFromScores_(
        Number(itemFullNum) || 0,
        baselineEq,
        heroicDarkRow,
        hidePct,
        ceHasEq
      );
      if (usePlainFullSheet) {
        lPlainTexts[i] = [uh.text == null ? "" : uh.text];
        lPlainColors[i] = [uh.color || "#000000"];
      } else {
        upgradeWrites.push({ r: r, text: uh.text, color: uh.color });
      }
    }
    if (capBandSnap && fpKey != null) {
      if (!capBandSnap.cmpSkipFp) capBandSnap.cmpSkipFp = {};
      capBandSnap.cmpSkipFp[r] = fpKey;
    }
  }

  if (usePlainFullSheet) {
    bpSheet.getRange(BIS_FIRST_DATA_ROW, CMP_DISP_COL, cmpNumRows, 1).setValues(kPlainMat);
    bpSheet.getRange(BIS_FIRST_DATA_ROW, GG_UPGRADE_COL, cmpNumRows, 1).setValues(lPlainTexts);
    try {
      bpSheet.getRange(BIS_FIRST_DATA_ROW, GG_UPGRADE_COL, cmpNumRows, 1).setFontColors(lPlainColors);
    } catch (ePl) {}
  } else {
    flushComparisonWrites_(bpSheet, writes, cmpRawLetter);
    flushUpgradePctCells_(bpSheet, upgradeWrites);
  }
  flushPlannerGearScoreCells_(bpSheet, []);
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
  var prevC = null;
  for (var si = 0; si < segs.length; si++) {
    if (si > 0) {
      var commaStart = out.s.length;
      out.s += ", ";
      if (prevC) {
        runs.push({ start: commaStart, end: out.s.length, color: prevC });
      }
    }
    var segStart = out.s.length;
    out.s += segs[si];
    var s = segs[si];
    var c = null;
    if (s.charAt(0) === "+") c = colPos;
    else if (s.charAt(0) === "-") c = colNeg;
    prevC = c;
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
  var prevC = null;
  for (si = 0; si < segs.length; si++) {
    if (si > 0) {
      var commaStart = out.length;
      out += ", ";
      if (prevC) {
        runs.push({ start: commaStart, end: out.length, color: prevC });
      }
    }
    var segStart = out.length;
    out += segs[si];
    var s = segs[si];
    var c = null;
    if (s.charAt(0) === "+") c = colPos;
    else if (s.charAt(0) === "-") c = colNeg;
    prevC = c;
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

