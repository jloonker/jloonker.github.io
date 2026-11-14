/**
 * jai.dance mailing list → Google Sheet
 *
 * Setup:
 * 1. Create a Google Sheet and put its ID in SPREADSHEET_ID below.
 * 2. Create an Apps Script project (script.google.com), paste this whole file, and save.
 * 3. Deploy → New deployment → type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy, approve the permissions, and copy the Web app URL (ends in /exec).
 * 4. Paste that URL into MAILING_LIST_URL in index.html.
 *
 * Each sign-up becomes one row in the "Signups" tab.
 */

const SPREADSHEET_ID = "1aJMnd8_Yg443xAMR1l9b6sxUNLVpxEIW9EzvFrzWjMY"; // "jai.dance mailing list"
const SHEET_NAME = "Signups";
const HEADERS = ["Signed up", "Name", "Email", "Based in", "Dance level"];
const LEVELS = ["New to dance", "Beginner", "Intermediate", "Advanced"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || "{}");
    const name = clean(d.name, 100);
    const email = clean(d.email, 200);
    const city = clean(d.city, 100);
    const level = LEVELS.indexOf(d.level) >= 0 ? d.level : "";

    if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return ContentService.createTextOutput("invalid");
    }

    const sheet = getSheet();
    // Skip duplicates: update the existing row instead of adding a new one.
    const emails = sheet.getRange(2, 3, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat()
      .map(function (v) { return String(v).toLowerCase(); });
    const i = emails.indexOf(email.toLowerCase());
    const row = [new Date(), name, email, city, level];
    if (i >= 0 && sheet.getLastRow() > 1) {
      sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

// Trim, cap length, and stop spreadsheet formula injection (=, +, -, @).
function clean(v, max) {
  let s = String(v == null ? "" : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// Run once from the editor to create the "Signups" tab and approve permissions.
function setup() {
  getSheet();
}
