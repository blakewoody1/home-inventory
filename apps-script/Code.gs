/**
 * Home Inventory backend (Google Apps Script, bound to the "Home Inventory" sheet).
 *
 * The app sends POST requests with a JSON body: { key, action, ... }
 * Actions: ping, list, add, update, remove, upload, getPhoto
 * Every response is JSON: { ok: true, ... } or { ok: false, error: "..." }
 */

const KEY = 'CHANGE-ME-to-your-key';   // any word you like; the same one goes in the app's Settings
const TAB = 'Items';
const FOLDER_NAME = 'Home Inventory';

// Sheet header -> field name the app uses
const FIELDS = {
  'ID': 'id',
  'Item': 'item',
  'Category': 'category',
  'Room': 'room',
  'Brand': 'brand',
  'Model': 'model',
  'Serial': 'serial',
  'Purchase Date': 'purchaseDate',
  'Price': 'price',
  'Retailer': 'retailer',
  'Warranty Expires': 'warrantyExpires',
  'Photo': 'photo',
  'Manual Link': 'manualLink',
  'Receipt Link': 'receiptLink',
  'Notes': 'notes'
};
const DATE_FIELDS = ['purchaseDate', 'warrantyExpires'];
const UPLOAD_FIELDS = { photo: 'photo', manual: 'manualLink', receipt: 'receiptLink' };

function doPost(e) {
  let d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return out({ ok: false, error: 'Bad request' });
  }
  if (d.key !== KEY) return out({ ok: false, error: 'Wrong key' });

  try {
    switch (d.action) {
      case 'ping':     return out({ ok: true, sheet: SpreadsheetApp.getActive().getName() });
      case 'list':     return out({ ok: true, items: listItems() });
      case 'add':      return out(withLock(function () { return addItem(d.item || {}); }));
      case 'update':   return out(withLock(function () { return updateItem(d.id, d.item || {}); }));
      case 'remove':   return out(withLock(function () { return removeItem(d.id); }));
      case 'upload':   return out(withLock(function () { return uploadFile(d); }));
      case 'getPhoto': return out(getPhoto(d.fileId));
      default:         return out({ ok: false, error: 'Unknown action: ' + d.action });
    }
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}

// Opening the web app URL in a browser shows this, so you can tell it's deployed.
function doGet() {
  return ContentService.createTextOutput('Home Inventory backend is running.');
}

/* ---------- helpers ---------- */

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TAB);
  if (!sh) throw new Error('Tab "' + TAB + '" not found');
  return sh;
}

// Column number (1-based) for each field, read from the header row,
// so columns can be reordered in the sheet without breaking the app.
function columns(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cols = {};
  headers.forEach(function (h, i) {
    const f = FIELDS[String(h).trim()];
    if (f) cols[f] = i + 1;
  });
  if (!cols.id) throw new Error('No "ID" column in row 1');
  return cols;
}

function fmtDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v === '' || v == null ? '' : String(v);
}

function parseDate(s) {
  const p = String(s || '').split('-');
  if (p.length !== 3) return '';
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function toCell(field, value) {
  if (value == null) return '';
  if (DATE_FIELDS.indexOf(field) !== -1) return parseDate(value);
  if (field === 'price') {
    const n = Number(String(value).replace(/[$,\s]/g, ''));
    return value === '' || isNaN(n) ? '' : n;
  }
  return String(value).trim();
}

function findRow(sh, cols, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return -1;
  const ids = sh.getRange(2, cols.id, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function nextId(sh, cols) {
  const last = sh.getLastRow();
  let max = 0;
  if (last >= 2) {
    sh.getRange(2, cols.id, last - 1, 1).getValues().forEach(function (r) {
      const n = parseInt(String(r[0]).replace(/\D/g, ''), 10);
      if (!isNaN(n) && n > max) max = n;
    });
  }
  return ('0000' + (max + 1)).slice(-4);
}

function rowToItem(row, cols) {
  const item = {};
  Object.keys(cols).forEach(function (f) {
    const v = row[cols[f] - 1];
    item[f] = DATE_FIELDS.indexOf(f) !== -1 ? fmtDate(v) : (v === '' || v == null ? '' : v);
  });
  item.id = String(item.id);
  return item;
}

/* ---------- actions ---------- */

function listItems() {
  const sh = sheet();
  const cols = columns(sh);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues()
    .filter(function (r) { return String(r[cols.id - 1]) !== '' || String(r[(cols.item || cols.id) - 1]) !== ''; })
    .map(function (r) { return rowToItem(r, cols); });
}

function addItem(item) {
  if (!item.item || !String(item.item).trim()) throw new Error('Item name is required');
  const sh = sheet();
  const cols = columns(sh);
  const id = nextId(sh, cols);
  const row = [];
  for (let c = 0; c < sh.getLastColumn(); c++) row.push('');
  Object.keys(cols).forEach(function (f) {
    if (f === 'id') row[cols.id - 1] = id;
    else if (item[f] !== undefined) row[cols[f] - 1] = toCell(f, item[f]);
  });
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setValues([row]);
  return { ok: true, item: rowToItem(sh.getRange(r, 1, 1, row.length).getValues()[0], cols) };
}

// Only the fields sent are changed; everything else in the row is left alone.
function updateItem(id, item) {
  const sh = sheet();
  const cols = columns(sh);
  const r = findRow(sh, cols, id);
  if (r === -1) throw new Error('Item ' + id + ' not found');
  Object.keys(item).forEach(function (f) {
    if (f === 'id' || !cols[f]) return;
    sh.getRange(r, cols[f]).setValue(toCell(f, item[f]));
  });
  return { ok: true, item: rowToItem(sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0], cols) };
}

// Deletes the row. Photos and files stay in the Drive folder.
function removeItem(id) {
  const sh = sheet();
  const cols = columns(sh);
  const r = findRow(sh, cols, id);
  if (r === -1) throw new Error('Item ' + id + ' not found');
  sh.deleteRow(r);
  return { ok: true, removed: String(id) };
}

function folder() {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('FOLDER_ID');
  if (saved) {
    try {
      const f = DriveApp.getFolderById(saved);
      if (!f.isTrashed()) return f;
    } catch (err) { /* fall through and find or create */ }
  }
  const found = DriveApp.getFoldersByName(FOLDER_NAME);
  const f = found.hasNext() ? found.next() : DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

// d: { id, kind: 'photo'|'manual'|'receipt', data: base64, mimeType, name }
// Saves the file to the Drive folder and writes its link into the item's row.
function uploadFile(d) {
  const field = UPLOAD_FIELDS[d.kind || 'photo'];
  if (!field) throw new Error('Unknown upload kind: ' + d.kind);
  if (!d.data) throw new Error('No file data');
  const sh = sheet();
  const cols = columns(sh);
  const r = findRow(sh, cols, d.id);
  if (r === -1) throw new Error('Item ' + d.id + ' not found');
  if (!cols[field]) throw new Error('Column for ' + field + ' not found');

  const itemName = cols.item ? String(sh.getRange(r, cols.item).getValue()) : '';
  const ext = (String(d.name || '').match(/\.[A-Za-z0-9]+$/) || [d.mimeType === 'application/pdf' ? '.pdf' : '.jpg'])[0];
  const name = [d.id, itemName, d.kind || 'photo'].filter(String).join(' - ') + ext;
  const blob = Utilities.newBlob(Utilities.base64Decode(d.data), d.mimeType || 'image/jpeg', name);
  const file = folder().createFile(blob);

  sh.getRange(r, cols[field]).setValue(file.getUrl());
  return { ok: true, field: field, url: file.getUrl(), fileId: file.getId() };
}

// Returns a file from the inventory folder as base64, so the app can show
// photos without making anything in Drive public.
function getPhoto(fileId) {
  if (!fileId) throw new Error('No fileId');
  const file = DriveApp.getFileById(fileId);
  const parents = file.getParents();
  const folderId = folder().getId();
  let inFolder = false;
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId) { inFolder = true; break; }
  }
  if (!inFolder) throw new Error('File is not in the ' + FOLDER_NAME + ' folder');
  const blob = file.getBlob();
  return { ok: true, mimeType: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}
