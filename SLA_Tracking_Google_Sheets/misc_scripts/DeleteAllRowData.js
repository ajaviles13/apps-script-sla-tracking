/**
 * ============================================================
 * SLA Tracker - Bulk Tab Data Clear Utility
 * ============================================================
 * For every tab in the spreadsheet EXCEPT the excluded sheets:
 *   1. Unhides any manually hidden rows
 *   2. Clears row data (A2:Q or A2:R, depending on TEMPLATE_TYPE)
 *      using open-ended row notation (e.g. "A2:Q") so it only
 *      clears down to each sheet's existing row count rather than
 *      forcing it to expand
 *   3. Resets the cell fill color to white on columns A:H (rows 2
 *      down, same rows as the content clear — header row is left
 *      alone), for BOTH template types
 *
 * Uses clearContent() (not clear()) so formatting and data
 * validation rules (e.g., dropdowns) are preserved — only cell
 * values/formulas are removed.
 *
 * NOTE: showRows() only un-hides rows that were manually hidden.
 * It does not affect rows hidden by an active Filter (as opposed
 * to a Filter View) — those are controlled separately.
 *
 * USAGE:
 * 1. Set TEMPLATE_TYPE below to "PDF" or "Word"
 * 2. Run clearAllTabData() from the Apps Script editor
 * 3. If your spreadsheet has a mix of PDF and Word tabs, flip
 *    TEMPLATE_TYPE and run again for the other set. This script
 *    applies the SAME range to every non-excluded tab on each run.
 * ============================================================
 */

// ---- CONFIG ----
var TEMPLATE_TYPE = "PDF"; // "PDF" or "Word"

var EXCLUDED_SHEET_NAMES = ["ℹ️ HOW TO USE", "SLA Control"];

var START_ROW = 2;

// Fill-color reset applies to these columns for BOTH template types
var COLOR_RESET_START_COL = "A";
var COLOR_RESET_END_COL = "H";
var WHITE = "#ffffff";

/**
 * Main entry point. Run this function from the Apps Script editor.
 */
function clearAllTabData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  var lastColumn = getLastColumnForTemplateType(TEMPLATE_TYPE);
  // Open-ended row notation (e.g. "A2:Q") targets the sheet's existing
  // row count instead of forcing it to expand — avoids blowing past the
  // workbook's 10,000,000 cell limit when many tabs are involved.
  var rangeA1 = "A" + START_ROW + ":" + lastColumn;

  var clearedSheets = [];
  var skippedSheets = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();

    if (EXCLUDED_SHEET_NAMES.indexOf(sheetName) !== -1) {
      skippedSheets.push(sheetName);
      continue;
    }

    // 1. Unhide any manually hidden rows
    var maxRows = sheet.getMaxRows();
    sheet.showRows(1, maxRows);

    // 2. Clear row data, preserving formatting + data validation rules
    var range = sheet.getRange(rangeA1);
    range.clearContent();

    // 3. Reset fill color to white on columns A:H (rows 2 down),
    //    same for both PDF and Word template types
    var colorRangeA1 = COLOR_RESET_START_COL + START_ROW + ":" + COLOR_RESET_END_COL;
    sheet.getRange(colorRangeA1).setBackground(WHITE);

    clearedSheets.push(sheetName);
  }

  Logger.log("Template type: " + TEMPLATE_TYPE);
  Logger.log("Range cleared: " + rangeA1);
  Logger.log("Cleared " + clearedSheets.length + " sheet(s): " + clearedSheets.join(", "));
  Logger.log("Skipped " + skippedSheets.length + " sheet(s): " + skippedSheets.join(", "));
}

/**
 * Maps TEMPLATE_TYPE to its last column letter.
 * PDF  -> Q
 * Word -> R
 */
function getLastColumnForTemplateType(templateType) {
  if (templateType === "PDF") {
    return "Q";
  } else if (templateType === "Word") {
    return "R";
  } else {
    throw new Error("Invalid TEMPLATE_TYPE: '" + templateType + "'. Must be 'PDF' or 'Word'.");
  }
}