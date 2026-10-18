// Headroom — Google Apps Script backend
// Deploy: Extensions → Apps Script → Deploy → New deployment → Web app
// Execute as: Me | Access: Anyone with the link

const SHEET_NAME = 'Expenses';
const HEADERS = ['id', 'date', 'category', 'amount', 'note', 'freq', 'createdAt'];

function doGet(e) {
  return ok(sheetToRows(getSheet()));
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet();
  if (body.op === 'upsert') {
    upsertRow(sheet, body.row);
  } else if (body.op === 'delete') {
    deleteRow(sheet, body.id);
  } else if (body.op === 'bulk_upsert') {
    (body.rows || []).forEach(row => upsertRow(sheet, row));
  }
  return ok({ saved: true });
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
}

function sheetToRows(sheet) {
  ensureHeaders(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const hdrs = data[0].map(h => String(h).trim());
  return data.slice(1)
    .map(row => {
      const obj = {};
      hdrs.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    })
    .filter(r => r.id); // skip blank rows
}

function findRowIndex(sheet, id) {
  const idCol = HEADERS.indexOf('id');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

function upsertRow(sheet, row) {
  ensureHeaders(sheet);
  const hdrs = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length)).getValues()[0];
  const newRow = hdrs.map(h => (row[h] !== undefined && row[h] !== null) ? row[h] : '');
  const rowIdx = findRowIndex(sheet, row.id);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }
}

function deleteRow(sheet, id) {
  const rowIdx = findRowIndex(sheet, id);
  if (rowIdx > 0) sheet.deleteRow(rowIdx);
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
