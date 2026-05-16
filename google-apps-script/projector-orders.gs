/**
 * Projectors landing page → Google Sheet
 *
 * SETUP:
 * 1. Create a new Google Sheet.
 * 2. Extensions → Apps Script → paste this file → Save.
 * 3. Run setupSheet once (authorize when prompted).
 * 4. Deploy → New deployment → Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the /exec URL into projectors/index.html → SCRIPT_URL
 */

const SHEET_NAME = 'Projector Orders';

function parsePostParams_(e) {
  const p = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach((key) => {
      p[key] = e.parameter[key];
    });
  }
  if (Object.keys(p).length) return p;

  if (e && e.postData && e.postData.contents) {
    const type = String(e.postData.type || '').toLowerCase();
    if (type.indexOf('application/x-www-form-urlencoded') !== -1) {
      String(e.postData.contents).split('&').forEach((pair) => {
        const eq = pair.indexOf('=');
        const key = decodeURIComponent((eq >= 0 ? pair.substring(0, eq) : pair).replace(/\+/g, ' '));
        const val = decodeURIComponent((eq >= 0 ? pair.substring(eq + 1) : '').replace(/\+/g, ' '));
        p[key] = val;
      });
    }
  }
  return p;
}

function setupSheet() {
  const sheet = getOrCreateSheet_();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Name', 'City', 'Phone number', 'Model', 'Quantity', 'Price']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
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
  return ContentService.createTextOutput('Projector orders endpoint is running.');
}

function doPost(e) {
  try {
    const p = parsePostParams_(e);
    const sheet = getOrCreateSheet_();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Date', 'Name', 'City', 'Phone number', 'Model', 'Quantity', 'Price']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const name = String(p.name || '').trim();
    const city = String(p.city || '').trim();
    const phone = String(p.phone || '').replace(/\D/g, '');
    const model = String(p.model || '300w').trim();
    const quantity = String(p.quantity || '1').trim();
    const price = String(p.price || '699').trim();

    sheet.appendRow([
      new Date(),
      name,
      city,
      phone,
      model,
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
