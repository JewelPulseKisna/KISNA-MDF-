// ═══════════════════════════════════════════════════════════════
//  KISNA MDF — Google Apps Script Backend
//  ─────────────────────────────────────────────────────────────
//  HOW TO DEPLOY:
//  1. Open a new Google Sheet
//  2. Go to Extensions → Apps Script
//  3. Delete any existing code and paste THIS entire file
//  4. Save (Ctrl + S)
//  5. Click Deploy → New deployment
//  6. Type: Web App
//  7. Execute as: Me
//  8. Who has access: Anyone
//  9. Click Deploy → Copy the Web App URL
//  10. Paste that URL into the Kisna MDF app → Sheets → Connect
// ═══════════════════════════════════════════════════════════════

const IMPORT_SHEET = 'Import';
const STORES_SHEET   = 'Stores';
const INV_SHEET      = 'Inventory';
const EXP_SHEET      = 'Expenses';

// ── GET (health check) ──────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'Kisna MDF API is running ✓' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST (main handler) ─────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    switch (action) {
      case 'getAll':         result = getAll();            break;
      case 'addStore':       result = addStore(body);      break;
      case 'updateStore':    result = updateStore(body);   break;
      case 'deleteStore':    result = deleteStore(body);   break;
      case 'addInventory':   result = addInventory(body);  break;
      case 'deleteInventory':result = deleteInventory(body);break;
      case 'addExpense':     result = addExpense(body);    break;
      case 'deleteExpense':  result = deleteExpense(body); break;
      case 'importFromRaw': result = importFromRaw(); break;
      default:               result = { error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── INIT SHEETS ─────────────────────────────────────────────────
// Creates the three sheets with headers if they don't exist yet
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(STORES_SHEET)) {
    const s = ss.insertSheet(STORES_SHEET);
    s.appendRow(['id','name','storeCode','zone','state','city','softLaunchDate','inaugurationDate','openingMonth','phone','email','locationType','location','address','pinCode','storeSize','floor','gst']);
    s.setFrozenRows(1);
    formatHeader(s);
  }

  if (!ss.getSheetByName(INV_SHEET)) {
    const s = ss.insertSheet(INV_SHEET);
    s.appendRow(['id', 'storeId', 'month', 'type', 'value', 'budget', 'notes']);
    s.setFrozenRows(1);
    formatHeader(s);
  }

  if (!ss.getSheetByName(EXP_SHEET)) {
    const s = ss.insertSheet(EXP_SHEET);
    s.appendRow(['id', 'storeId', 'month', 'actual', 'cap', 'used', 'desc']);
    s.setFrozenRows(1);
    formatHeader(s);
  }
}

function formatHeader(sheet) {
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  hdr.setBackground('#1a1a2e').setFontColor('#C9A84C').setFontWeight('bold');
}

// ── GET ALL ─────────────────────────────────────────────────────
function getAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initSheets();

  // Stores → object keyed by id
  const storesSheet = ss.getSheetByName(STORES_SHEET);
  const storesData  = storesSheet.getDataRange().getValues();
  const stHdrs      = storesData[0];
  const stores      = {};
  for (let i = 1; i < storesData.length; i++) {
    const row = storesData[i];
    if (!row[0]) continue;
    const rec = {};
    stHdrs.forEach((h, j) => { rec[h] = String(row[j] != null ? row[j] : ''); });
    stores[rec.id] = rec;
  }

  // Inventory → array
  const invSheet  = ss.getSheetByName(INV_SHEET);
  const invData   = invSheet.getDataRange().getValues();
  const invHdrs   = invData[0];
  const numCols_i = ['value', 'budget'];
  const inventory = [];
  for (let i = 1; i < invData.length; i++) {
    const row = invData[i];
    if (!row[0]) continue;
    const rec = {};
    invHdrs.forEach((h, j) => {
      rec[h] = numCols_i.includes(h) ? Number(row[j]) : String(row[j] != null ? row[j] : '');
    });
    inventory.push(rec);
  }

  // Expenses → array
  const expSheet  = ss.getSheetByName(EXP_SHEET);
  const expData   = expSheet.getDataRange().getValues();
  const expHdrs   = expData[0];
  const numCols_e = ['actual', 'cap', 'used'];
  const expenses  = [];
  for (let i = 1; i < expData.length; i++) {
    const row = expData[i];
    if (!row[0]) continue;
    const rec = {};
    expHdrs.forEach((h, j) => {
      rec[h] = numCols_e.includes(h) ? Number(row[j]) : String(row[j] != null ? row[j] : '');
    });
    expenses.push(rec);
  }

  return { stores, inventory, expenses };
}

// ── STORES ──────────────────────────────────────────────────────
function addStore(data) {
  initSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STORES_SHEET);
  sheet.appendRow([
    data.id, data.name, data.storeCode||'', data.zone||'', data.state||'',
    data.city||'', data.softLaunchDate||'', data.inaugurationDate||'',
    data.openingMonth||'', data.phone||'', data.email||'',
    data.locationType||'', data.location||'', data.address||'',
    data.pinCode||'', data.storeSize||'', data.floor||'', data.gst||''
  ]);
  return { ok: true };
}

function updateStore(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STORES_SHEET);
  const rows  = sheet.getDataRange().getValues();
  const hdrs  = rows[0];
  const idCol = hdrs.indexOf('id');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(data.id)) {
      hdrs.forEach((h, j) => {
        if (data[h] !== undefined) sheet.getRange(i + 1, j + 1).setValue(data[h]);
      });
      return { ok: true };
    }
  }
  return { error: 'Store not found' };
}

function deleteStore(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  deleteRowById(ss.getSheetByName(STORES_SHEET), data.id);
  deleteRowsWhere(ss.getSheetByName(INV_SHEET),   'storeId', data.id);
  deleteRowsWhere(ss.getSheetByName(EXP_SHEET),   'storeId', data.id);
  return { ok: true };
}

// ── INVENTORY ────────────────────────────────────────────────────
function addInventory(data) {
  initSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INV_SHEET);
  sheet.appendRow([
    data.id, data.storeId, data.month, data.type,
    Number(data.value), Number(data.budget), data.notes || ''
  ]);
  return { ok: true };
}

function deleteInventory(data) {
  deleteRowById(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INV_SHEET), data.id);
  return { ok: true };
}

// ── EXPENSES ─────────────────────────────────────────────────────
function addExpense(data) {
  initSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXP_SHEET);
  sheet.appendRow([
    data.id, data.storeId, data.month,
    Number(data.actual), Number(data.cap), Number(data.used),
    data.desc || ''
  ]);
  return { ok: true };
}

function deleteExpense(data) {
  deleteRowById(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXP_SHEET), data.id);
  return { ok: true };
}

// ── HELPERS ──────────────────────────────────────────────────────
function deleteRowById(sheet, id) {
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}
function importFromRaw() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initSheets();
  const importSheet = ss.getSheetByName(IMPORT_SHEET);
  if (!importSheet) return { error: 'Import sheet not found' };
  const data = importSheet.getDataRange().getValues();
  if (data.length < 2) return { error: 'No data in Import sheet' };
  const storesSheet = ss.getSheetByName(STORES_SHEET);
  const lastRow = storesSheet.getLastRow();
  if (lastRow > 1) storesSheet.getRange(2, 1, lastRow - 1, storesSheet.getLastColumn()).clearContent();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    const storeCode    = String(r[0] || '');
    const zone         = String(r[1] || '');
    const state        = String(r[2] || '');
    const city         = String(r[3] || '');
    const softLaunch   = String(r[4] || '');
    const inauguration = String(r[5] || '');
    const location     = String(r[6] || '');
    const name         = String(r[7] || '');
    const phone        = String(r[8] || '');
    const email        = String(r[9] || '');
    const locationType = String(r[10] || '');
    const address      = String(r[11] || '');
    const pinCode      = String(r[12] || '');
    const storeSize    = String(r[13] || '');
    const floor        = String(r[14] || '');
    const gst          = String(r[15] || '');
    const openingMonth = inauguration || softLaunch || '';
    rows.push([storeCode,name,storeCode,zone,state,city,softLaunch,inauguration,openingMonth,phone,email,locationType,location,address,pinCode,storeSize,floor,gst]);
  }
  if (rows.length > 0) storesSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  return { ok: true, imported: rows.length };
}
function deleteRowsWhere(sheet, colName, value) {
  if (!sheet) return;
  const rows   = sheet.getDataRange().getValues();
  const hdrs   = rows[0];
  const colIdx = hdrs.indexOf(colName);
  if (colIdx === -1) return;
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][colIdx]) === String(value)) {
      sheet.deleteRow(i + 1);
    }
  }
}
