/**
 * Google Apps Script backend for Vusion TopStock Ground Truth.
 *
 * SETUP
 * -----
 * 1. Create a new Google Sheet.
 * 2. Extensions > Apps Script. Delete the default code, paste THIS file.
 * 3. Run `setupHeaders` once (from the editor) to create the header row.
 * 4. Deploy > New deployment > type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Copy the Web app URL.
 * 5. Paste that URL into GOOGLE_SHEET_ENDPOINT in app.js.
 *
 * Each POST body is JSON with: timestamp, email, storeType, storeAddress,
 * cameraBarcode, outOfShelfBarcode, rootCause. One row is appended per call.
 */

var SHEET_NAME = 'GroundTruth';
var HEADERS = [
  'Timestamp', 'Email', 'Store Type', 'Store Address',
  'Camera Barcode', 'Out-of-Shelf Barcode', 'Root Cause'
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
