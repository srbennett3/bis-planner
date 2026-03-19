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

// BIS Planner columns (1-indexed)
var GG_INTEREST = 1;  // A
var GG_SPEC     = 2;  // B
var GG_GEARTYPE = 3;  // C
var GG_NAME     = 4;  // D

// Current Equipment columns (1-indexed)
var CE_GEARTYPE = 1;  // A
var CE_ITEMNAME = 2;  // B

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

// ============================================================
// BIS Planner edits (Interest column)
// ============================================================

function handleBISPlannerEdit(e, bpSheet) {
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (col !== GG_INTEREST || row <= 1) return;

  var newValue = e.range.getValue();
  var spec = bpSheet.getRange(row, GG_SPEC).getValue();
  var gearType = bpSheet.getRange(row, GG_GEARTYPE).getValue();
  var itemName = normalizeItemName(bpSheet.getRange(row, GG_NAME).getValue());

  if (!normalizeItemName(spec) || !normalizeItemName(gearType)) return;

  if (newValue === "Equipped") {
    e.range.setFontWeight("bold");
    clearOtherEquipped(bpSheet, row, spec, gearType);
    setCEItem(spec, gearType, itemName);
  } else {
    e.range.setFontWeight("normal");
    if (e.oldValue === "Equipped") {
      clearCEItemIfMatch(spec, gearType, itemName);
    }
  }
}

function clearOtherEquipped(bpSheet, currentRow, spec, gearType) {
  var lastRow = bpSheet.getLastRow();
  var data = bpSheet.getRange(2, 1, lastRow - 1, GG_NAME).getValues();

  for (var i = 0; i < data.length; i++) {
    var r = i + 2;
    if (r === currentRow) continue;
    if (data[i][GG_INTEREST - 1] === "Equipped" &&
        normalizeItemName(data[i][GG_SPEC - 1]) === normalizeItemName(spec) &&
        normalizeItemName(data[i][GG_GEARTYPE - 1]) === normalizeItemName(gearType)) {
      bpSheet.getRange(r, GG_INTEREST).setValue("");
      bpSheet.getRange(r, GG_INTEREST).setFontWeight("normal");
    }
  }
}

// ============================================================
// Current Equipment edits (Item Name column)
// ============================================================

function handleCurrentEquipEdit(e, ceSheet) {
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (col !== CE_ITEMNAME || row <= 1) return;

  var gearType = normalizeItemName(ceSheet.getRange(row, CE_GEARTYPE).getValue());
  if (!gearType || gearType.indexOf("---") === 0) return;

  var spec = findSpecForCERow(ceSheet, row);
  if (!spec) return;

  var newItemName = normalizeItemName(e.range.getValue());
  var oldItemName = normalizeItemName(e.oldValue);

  var bpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIS_PLANNER_SHEET);
  if (!bpSheet) return;

  if (oldItemName) {
    clearInterestForItem(bpSheet, spec, gearType, oldItemName);
  }

  if (newItemName) {
    setInterestForItem(bpSheet, spec, gearType, newItemName);
  }
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
  var data = bpSheet.getRange(2, 1, lastRow - 1, GG_NAME).getValues();

  for (var i = 0; i < data.length; i++) {
    var r = i + 2;
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
  var data = bpSheet.getRange(2, 1, lastRow - 1, GG_NAME).getValues();

  for (var i = 0; i < data.length; i++) {
    var r = i + 2;
    if (normalizeItemName(data[i][GG_SPEC - 1]) === normalizeItemName(spec) &&
        normalizeItemName(data[i][GG_GEARTYPE - 1]) === normalizeItemName(gearType) &&
        normalizeItemName(data[i][GG_NAME - 1]) === want &&
        data[i][GG_INTEREST - 1] === "Equipped") {
      bpSheet.getRange(r, GG_INTEREST).setValue("");
      bpSheet.getRange(r, GG_INTEREST).setFontWeight("normal");
    }
  }
}
