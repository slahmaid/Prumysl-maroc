/**
 * Saqr landing page → Google Sheet
 *
 * SETUP:
 * 1. Create a new Google Sheet.
 * 2. Extensions → Apps Script → paste this file → Save.
 * 3. Run setupSheet once (authorize when prompted).
 * 4. Deploy → New deployment → Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the /exec URL into saqr.html → SCRIPT_URL
 */

const SHEET_NAME = 'Saqr Orders';

function setupSheet() {
  const sheet = getOrCreateSheet_();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Name', 'City', 'Phone number', 'Quantity', 'Price']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function doGet() {
  return ContentService.createTextOutput('Saqr orders endpoint is running.');
}

function doPost(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const sheet = getOrCreateSheet_();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Date', 'Name', 'City', 'Phone number', 'Quantity', 'Price']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const name = String(p.name || '').trim();
    const city = String(p.city || '').trim();
    const phone = String(p.phone || '').replace(/\D/g, '');
    const quantity = String(p.quantity || '1').trim();
    const price = String(p.price || '1999').trim();

    sheet.appendRow([
      new Date(),
      name,
      city,
      phone,
      quantity,
      price
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
