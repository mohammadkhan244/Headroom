// Headroom — Apps Script backend (Expenses + Balances)
// Setup: open via Extensions → Apps Script inside your headroom Sheet
// Script Properties → add SHEET_SECRET (same value as in Vercel)
// Deploy → New deployment → Web app → Execute as: Me → Access: Anyone

const EXP_HEADERS = ['id','date','category','categories','amount','note','freq','type',
                     'reimbursable','reimb_status','recurring_id','account','createdAt'];
const BAL_HEADERS = ['id','name','kind','balance','apy','shares','strike','cliff','note','asOf','updatedAt'];

const SHEET_SECRET = PropertiesService.getScriptProperties().getProperty('SHEET_SECRET');

function doGet(e) {
  if (!e || !e.parameter) return unauthorized();
  if ((e.parameter.secret || '') !== SHEET_SECRET) return unauthorized();
  // ?sheet=all → return both expenses and balances in one call
  if (e.parameter.sheet === 'all') {
    return ok({ ok: true, rows: sheetToRows(getSheet('Expenses', EXP_HEADERS)), balances: sheetToRows(getSheet('Balances', BAL_HEADERS)) });
  }
  return ok({ ok: true, rows: sheetToRows(getSheet('Expenses', EXP_HEADERS)) });
}

function doPost(e) {
  if (!e || !e.postData) return unauthorized();
  const body = JSON.parse(e.postData.contents);
  if ((body.secret || '') !== SHEET_SECRET) return unauthorized();
  const isBalances = body.sheet === 'Balances';
  const sh = isBalances ? getSheet('Balances', BAL_HEADERS) : getSheet('Expenses', EXP_HEADERS);
  const hdrs = isBalances ? BAL_HEADERS : EXP_HEADERS;
  if      (body.op === 'upsert') upsertRow(sh, hdrs, body.row);
  else if (body.op === 'delete') deleteRow(sh, body.id);
  return ok({ ok: true });
}

// ── helpers ─────────────────────────────────────────────────────────

function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function sheetToRows(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const lastCol = sh.getLastColumn();
  const data = sh.getRange(1, 1, last, lastCol).getValues();
  const hdrs = data[0].map(h => String(h).trim());
  return data.slice(1)
    .map(row => { const o = {}; hdrs.forEach((h,i) => { o[h] = row[i] !== undefined ? row[i] : ''; }); return o; })
    .filter(r => r.id);
}

function upsertRow(sh, hdrs, row) {
  const shHdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  if (!row.categories && row.category) row.categories = row.category;
  const newRow = shHdrs.map(h => (row[h] !== undefined && row[h] !== null) ? row[h] : '');
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
