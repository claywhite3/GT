var SHEET_NAME = 'GroundTruth';
var HEADERS = [
  'Timestamp', 'Email', 'Store Type', 'Store Address',
  'Camera Barcode', 'Out-of-Shelf Barcode', 'Classification', 'Root Cause / FP Note'
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

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet();
    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.email || '',
      data.storeType || '',
      data.storeAddress || '',
      data.cameraBarcode || '',
      data.outOfShelfBarcode || '',
      data.classification || '',
      data.rootCause || ''
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Ground Truth endpoint is live.');
}
