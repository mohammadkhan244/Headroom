// Headroom — Google Apps Script backend v3
// Deploy: Extensions → Apps Script → Deploy → New deployment → Web app
// Execute as: Me | Access: Anyone with the link

const SHEET_NAMES = ['Expenses', 'Balances', 'Config'];

const HEADERS = {
  Expenses: ['id','date','category','categories','amount','note','freq','type',
             'reimbursable','reimb_status','recurring_id','account','createdAt'],
  Balances: ['id','name','kind','balance','apy','shares','strike','cliff',
             'note','asOf','updatedAt'],
  Config:   ['key','value']
};

const ALLOWED_SHEETS = new Set(SHEET_NAMES);

function getSheetName(e) {
  const name = (e && e.parameter && e.parameter.sheet) || 'Expenses';
  return ALLOWED_SHEETS.has(name) ? name : 'Expenses';
}

function doGet(e) {
  const secret = e && e.parameter && e.parameter.secret;
  if (!checkSecret(secret)) return unauthorized();
  const sheetParam = getSheetName(e);
  if (sheetParam === 'all') {
    // Not reachable via getSheetName — handled below
  }
  // sheet=all: return all three
  if (e && e.parameter && e.parameter.sheet === 'all') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ok({
      ok: true,
      expenses: sheetToRows(getOrCreate(ss, 'Expenses')),
      balances: sheetToRows(getOrCreate(ss, 'Balances')),
      config:   configToMap(getOrCreate(ss, 'Config'))
    });
  }
  return ok({ ok: true, rows: sheetToRows(getOrCreate(SpreadsheetApp.getActiveSpreadsheet(), sheetParam)) });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (!checkSecret(body.secret)) return unauthorized();

  const sheetName = ALLOWED_SHEETS.has(body.sheet) ? body.sheet : 'Expenses';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, sheetName);

  if (body.op === 'upsert') {
    upsertRow(sheet, sheetName, body.row);
  } else if (body.op === 'delete') {
    deleteRow(sheet, body.id);
  } else if (body.op === 'bulk_upsert') {
    (body.rows || []).forEach(row => upsertRow(sheet, sheetName, row));
  } else if (body.op === 'config_set') {
    configSet(sheet, body.key, body.value);
  }
  return ok({ ok: true, saved: true });
}

// ── Helpers ──────────────────────────────────────────────────────────────────

const SHEET_SECRET = PropertiesService.getScriptProperties().getProperty('SHEET_SECRET');
function checkSecret(s) { return s && s === SHEET_SECRET; }
function unauthorized() {
  return ContentService.createTextOutput(JSON.stringify({error:'unauthorized'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders(sheet, name) {
  const hdrs = HEADERS[name] || HEADERS['Expenses'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(hdrs);
    sheet.getRange(1, 1, 1, hdrs.length).setFontWeight('bold');
    return;
  }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  hdrs.forEach(h => {
    if (!existing.includes(h)) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(h).setFontWeight('bold');
    }
  });
}

function sheetToRows(sheet) {
  const name = sheet.getName();
  ensureHeaders(sheet, name);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 2) return [];
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const hdrs = data[0].map(h => String(h).trim());
  return data.slice(1)
    .map(row => {
      const obj = {};
      hdrs.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    })
    .filter(r => r.id);
}

function configToMap(sheet) {
  ensureHeaders(sheet, 'Config');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  data.forEach(([k, v]) => {
    if (!k) return;
    try { map[String(k)] = JSON.parse(String(v)); } catch { map[String(k)] = v; }
  });
  return map;
}

function configSet(sheet, key, value) {
  ensureHeaders(sheet, 'Config');
  const lastRow = sheet.getLastRow();
  const serialized = JSON.stringify(value);
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(key)) {
        sheet.getRange(i + 2, 2).setValue(serialized);
        return;
      }
    }
  }
  sheet.appendRow([key, serialized]);
}

function findRowIndex(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idCol = hdrs.map(h => String(h).trim()).indexOf('id');
  if (idCol < 0) return -1;
  const col = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) {
    if (String(col[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function upsertRow(sheet, sheetName, row) {
  ensureHeaders(sheet, sheetName);
  const lastCol = Math.max(sheet.getLastColumn(), (HEADERS[sheetName]||HEADERS['Expenses']).length);
  const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const rowToWrite = Object.assign({}, row);
  if (sheetName === 'Expenses' && !rowToWrite.categories && rowToWrite.category) {
    rowToWrite.categories = rowToWrite.category;
  }
  const newRow = hdrs.map(h => (rowToWrite[h] !== undefined && rowToWrite[h] !== null) ? rowToWrite[h] : '');
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
