/* =========================================================================
   Captana Ground Truth — Apps Script backend
   IMPORTANT: This script must be BOUND to the Captana Ground Truth spreadsheet.
   Open that spreadsheet → Extensions → Apps Script, and paste this there.
   ========================================================================= */

var SHEET_NAME = 'GroundTruth';
var HEADERS = [
  'Timestamp', 'Email', 'Store Type', 'Store Address', 'Category', 'Deployment Interval',
  'Camera Barcode', 'Out-of-Shelf Barcode', 'Classification',
  'Root Cause / FP Note', 'Reason Type', 'Record ID'
];

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Writes headers ONLY if the sheet is empty.
function setupHeaders() {
  getSheet();
}

// Force-rewrites row 1 with the current HEADERS, even if the sheet has data.
// Run this once after changing columns. Safe to run anytime.
function fixHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// Diagnostic: returns which spreadsheet this script is bound to + row count.
// Open the /exec URL with ?diag=1 to see it.
function diagInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  return {
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    hasGroundTruthTab: !!sheet,
    dataRows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0
  };
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'busy' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet();

    var id = data.id || '';
    if (id) {
      var idCol = HEADERS.length; // last column = Record ID
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var existing = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
        for (var i = 0; i < existing.length; i++) {
          if (existing[i][0] === id) {
            return ContentService
              .createTextOutput(JSON.stringify({ ok: true, duplicate: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
    }

    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.email || '',
      data.storeType || '',
      data.storeAddress || '',
      data.category || '',
      data.deploymentInterval || '',
      data.cameraBarcode || '',
      data.outOfShelfBarcode || '',
      data.classification || '',
      data.rootCause || '',
      data.reasonType || '',
      id
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  // ?diag=1 -> report which spreadsheet this is bound to (helps debugging).
  try {
    if (e && e.parameter && e.parameter.diag) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, diag: diagInfo() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, rows: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var values = sheet.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      var empty = true;
      for (var c = 0; c < r.length; c++) { if (r[c] !== '' && r[c] !== null) { empty = false; break; } }
      if (empty) continue;
      rows.push({
        timestamp: r[0], email: r[1], storeType: r[2], storeAddress: r[3],
        category: r[4], deploymentInterval: r[5], cameraBarcode: r[6],
        outOfShelfBarcode: r[7], classification: r[8], rootCause: r[9],
        reasonType: r[10], id: r[11]
      });
    }
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, rows: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
