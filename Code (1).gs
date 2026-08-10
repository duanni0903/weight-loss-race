/**
 * Weight-loss race — Google Sheets backend
 * -----------------------------------------
 * Paste this into the Apps Script editor of a Google Sheet
 * (Extensions > Apps Script), then deploy it as a Web App:
 *   Deploy > New deployment > type: Web app
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the resulting /exec URL into the site's .env file as
 * VITE_APPS_SCRIPT_URL.
 *
 * Storage model: everything (all competitions, participants,
 * check-ins, comments) is kept as ONE JSON blob in cell A1 of a
 * sheet named "Data". This keeps the API tiny — one GET to read
 * the whole app state, one POST to overwrite it — since the
 * frontend already keeps the full state in memory. Good enough
 * for a small friend-group competition; not meant for heavy
 * concurrent traffic.
 */

const SHEET_NAME = "Data";
const DEFAULT_VALUE = '{"competitions":{}}';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange("A1").setValue(DEFAULT_VALUE);
  }
  return sheet;
}

function doGet(e) {
  var sheet = getSheet_();
  var value = sheet.getRange("A1").getValue();
  var text = value && String(value).trim() ? String(value) : DEFAULT_VALUE;
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = getSheet_();
  try {
    var body = JSON.parse(e.postData.contents);
    var json = JSON.stringify(body.data);
    // Simple lock to avoid two near-simultaneous writes corrupting the cell.
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      sheet.getRange("A1").setValue(json);
    } finally {
      lock.releaseLock();
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
