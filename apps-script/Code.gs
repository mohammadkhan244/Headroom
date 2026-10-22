// Headroom — Apps Script backend (Expenses only)
// Setup: open this via Extensions → Apps Script inside your headroom Sheet
// Script Properties → add SHEET_SECRET (same value as in Vercel)
// Deploy → New deployment → Web app → Execute as: Me → Access: Anyone

const HEADERS = ['id','date','category','categories','amount','note','freq','type',
                 'reimbursable','reimb_status','recurring_id','account','createdAt'];

const SHEET_SECRET = PropertiesService.getScriptProperties().getProperty('SHEET_SECRET');

function doGet(e) {
  if ((e.parameter.secret || '') !== SHEET_SECRET) return unauthorized();
  return ok({ ok: true, rows: sheetToRows(expenses()) });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if ((body.secret || '') !== SHEET_SECRET) return unauthorized();
  const sh = expenses();
  if      (body.op === 'upsert') upsertRow(sh, body.row);
  else if (body.op === 'delete') deleteRow(sh, body.id);
  return ok({ ok: true });
}

// ── helpers ─────────────────────────────────────────────────────────

function expenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Expenses') || ss.insertSheet('Expenses');
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function sheetToRows(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(1, 1, last, HEADERS.length).getValues();
  const hdrs = data[0].map(h => String(h).trim());
  return data.slice(1)
    .map(row => { const o = {}; hdrs.forEach((h,i) => { o[h] = row[i] !== undefined ? row[i] : ''; }); return o; })
    .filter(r => r.id && r.date);
}

function upsertRow(sh, row) {
  const hdrs = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0].map(h => String(h).trim());
  if (!row.categories && row.category) row.categories = row.category;
  const newRow = hdrs.map(h => (row[h] !== undefined && row[h] !== null) ? row[h] : '');
  const idx = findRow(sh, row.id);
  if (idx > 0) sh.getRange(idx, 1, 1, newRow.length).setValues([newRow]);
  else sh.appendRow(newRow);
}

function deleteRow(sh, id) {
  const idx = findRow(sh, id);
  if (idx > 0) sh.deleteRow(idx);
}

function findRow(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) { if (String(col[i][0]) === String(id)) return i + 2; }
  return -1;
}

function unauthorized() {
  return ContentService.createTextOutput(JSON.stringify({ error: 'unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
