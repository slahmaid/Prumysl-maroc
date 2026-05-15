/**
 * Prumysl — Moka orders (6 columns only)
 *
 * Setup:
 * 1. New Google Sheet → Extensions → Apps Script → paste this file.
 * 2. Run setupSheet once.
 * 3. Deploy as Web app (Execute as: Me, Anyone can access).
 * 4. Paste the /exec URL into moka/index.html → SCRIPT_URL
 */

var SHEET_NAME = 'Orders';
var UNIT_PRICE = 599;

var HEADERS = ['Date', 'Name', 'City', 'Phone', 'Price', 'Quantity'];

function getOrdersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function setupSheet() {
  var sheet = getOrdersSheet_();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function parseQuantity_(p) {
  if (p.quantity) {
    var q = parseInt(p.quantity, 10);
    if (!isNaN(q) && q > 0) return q;
  }
  return p.product_offer === '2_camera' ? 2 : 1;
}

function parsePrice_(p, quantity) {
  if (p.price) {
    var price = parseFloat(String(p.price).replace(',', '.'));
    if (!isNaN(price) && price > 0) return price;
  }
  return UNIT_PRICE * quantity;
}

function doGet() {
  return ContentService.createTextOutput('Moka orders endpoint OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    setupSheet();
    var p = (e && e.parameter) ? e.parameter : {};
    var quantity = parseQuantity_(p);
    var price = parsePrice_(p, quantity);

    getOrdersSheet_().appendRow([
      new Date(),
      p.name || '',
      p.city || '',
      p.phone || '',
      price,
      quantity
    ]);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: String(err) })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
