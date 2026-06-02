var SHEET_NAME = 'GroundTruth';
var HEADERS = [
  'Timestamp', 'Email', 'Store Type', 'Store Address', 'Category',
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

function setupHeaders() {
  getSheet();
}

// Force-rewrite the header row even if the sheet already has data. Run once
// after updating columns. Safe to run anytime; only touches row 1.
function fixHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function doPost(e) {
  // Lock so two near-simultaneous requests can't both append the same id.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (lockErr) {
    // Couldn't get the lock in time; bail rather than risk a dup.
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'busy' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet();

    // De-dupe: if this record id already exists in the ID column, skip it.
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
  // Return all rows as JSON so the app's View Data screen can read them.
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, rows: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var values = sheet.getDataRange().getValues();
    var head = values[0];
    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      // skip fully-empty rows
      var empty = true;
      for (var c = 0; c < r.length; c++) { if (r[c] !== '' && r[c] !== null) { empty = false; break; } }
      if (empty) continue;
      rows.push({
        timestamp: r[0], email: r[1], storeType: r[2], storeAddress: r[3],
        category: r[4], cameraBarcode: r[5], outOfShelfBarcode: r[6],
        classification: r[7], rootCause: r[8], reasonType: r[9], id: r[10]
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
