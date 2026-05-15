/*
  Google Apps Script v2 — Outflow & Inflow Tracker
  ═══════════════════════════════════════════════════════════════
  DEPLOY STEPS:
    1. Open https://script.google.com/home → New project
    2. Paste this file's contents
    3. Deploy → New deployment → Web App
       • Execute as: Me (owner)
       • Who has access: Anyone
    4. Copy the /exec URL → paste it into index.html as GAS_URL
  ═══════════════════════════════════════════════════════════════

  Sheets created automatically in the user's spreadsheet:
    "Records"        — outflow/inflow records
    "Categories"     — category list (outflow + inflow, never deleted)
    "Accounts"       — account list (auto-populated from records)
    "Subcategories"  — subcategory list per category (Name, Category, Order)
    "Users"          — Telegram user registry

  Type values: "outflow" | "inflow"
  All API calls use doGet (query params) to avoid CORS preflight.
*/

// ══════════════════════════════════════════════════════════════
// 1. COLUMN CONFIGURATION (RECORD ORDER)
// ══════════════════════════════════════════════════════════════
// You can reorder these fields to change the column order in the "Records" sheet.
// Ensure you have at least 'date', 'amount', 'category', and 'type' for the app to function correctly.
//
// OPTIONAL FIELDS — comment/uncomment to enable or disable per instance:
//   'month'  — month number extracted from date (1–12)
//   'day'    — day number extracted from date (1–31)
//   'txnType'  — transaction type (Revenue / Expense / Deposit / Transfer / Refund etc.)
//              Maps to "Catecory" column in VS Tracker template.
//              Enable in the app by passing &txnType=Revenue in addRecord calls.
//
const RECORD_COLUMNS = [
  'timestamp',   // Row creation time
  // 'month',    // Month number extracted from date (1–12)  ← optional
  'date',        // Purchase date → "Day" column in sheet (YYYY-MM-DD)
  'amount',      // Transaction amount
  // 'txnType',    // Top-level group label (Revenue / Expense etc.) ← optional
  'category',    // Category 2 in VS Tracker
  'subcategory', // Category 3 in VS Tracker
  'description', // Note / Description
  'account',     // Account / Wallet
  'name',        // User name (from settings or TG)
  'type'         // "outflow" or "inflow"
];

// Display names for the headers (you can rename these if you like)
const COLUMN_DISPLAY_NAMES = {
  'timestamp':   'Timestamp',
  'month':       'Month',
  'date':        'Day',
  'amount':      'Amount',
  'txnType':     'Transaction Type',
  'category':    'Category 2',
  'subcategory': 'Category 3',
  'description': 'Description',
  'account':     'Account',
  'name':        'Name',
  'type':        'Type'
};

// ── Sheet names ────────────────────────────────────────────────
const SH_OUTFLOWS   = 'Records';
const SH_CATEGORIES = 'Categories';
const SH_ACCOUNTS   = 'Accounts';
const SH_TXNTYPES   = 'Transaction Types';
const SH_SUBCATS    = 'Subcategories';
const SH_USERS      = 'Users';

// ── Derived headers ─────────────────────────────────────────────
const HDR_OUTFLOWS   = RECORD_COLUMNS.map(key => COLUMN_DISPLAY_NAMES[key] || key);
const HDR_CATEGORIES = ['Name','Type','Order'];
const HDR_ACCOUNTS   = ['Name','Order'];
const HDR_TXNTYPES   = ['Name','Order'];
const HDR_SUBCATS    = ['Name','Category','Order'];
const HDR_USERS      = ['Name','First name','Last name','Username','Date added'];

// Helper to get column index by field key (0-based)
function getColIdx(key) {
  return RECORD_COLUMNS.indexOf(key);
}

// ── Default categories ─────────────────────────────────────────
const DEF_OUTFLOW = [
  'Office','Other'
];
const DEF_INFLOW = [
  'Refund','Other'
];

// ══════════════════════════════════════════════════════════════
// Entry points
// ══════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const p = e.parameter;
    if (!p.action) return err('action required');

    const ss = p.spreadsheetId ? SpreadsheetApp.openById(p.spreadsheetId) : null;

    switch (p.action) {
      case 'getCategories':  return getCategories(ss, p.type || 'outflow');
      case 'addCategory':    return addCategory(ss, p);
      case 'renameCategory': return renameCategory(ss, p);
      case 'getAccounts':    return getListSheet(ss, SH_ACCOUNTS, HDR_ACCOUNTS);
      case 'getTxnTypes':     return getListSheet(ss, SH_TXNTYPES,   HDR_TXNTYPES);
      case 'getSubcats':     return getSubcats(ss);
      case 'addListItem':    return addListItemAction(ss, p);
      case 'renameListItem': return renameListItemAction(ss, p);
      case 'addRecord':      return addRecord(ss, p);
      case 'checkDuplicate': return checkDuplicate(ss, p);
      case 'getSummary':     return getSummary(ss, p.period, p.groupBy || 'month', p.viewBy || 'category');
      case 'getTrend':       return getTrend(ss, p.endPeriod, Number(p.count) || 6, p.groupBy || 'month');
      default:               return err('unknown action: ' + p.action);
    }
  } catch (e) {
    return err(e.message);
  }
}

// Backward-compat: v1 posted directly without action
function doPost(e) {
  try {
    const p = e.parameter;
    if (p.action) return doGet({ parameter: p });
    const ss = SpreadsheetApp.openById(p.spreadsheetId);
    return addRecord(ss, { date: p.date, amount: p.amount, category: p.category, note: p.note, type: 'outflow' });
  } catch (e) {
    return err(e.message);
  }
}

// ══════════════════════════════════════════════════════════════
// Helpers
// ══════════════════════════════════════════════════════════════

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function err(msg) {
  return json({ status: 'error', message: msg });
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    
    // Auto-format Date column in Records sheet
    if (name === SH_OUTFLOWS) {
      const dateIdx = getColIdx('date');
      if (dateIdx !== -1) {
        const colLetter = String.fromCharCode(65 + dateIdx); // A, B, C...
        sheet.getRange(`${colLetter}2:${colLetter}`).setNumberFormat('yyyy-mm-dd');
      }
    }
  }
  return sheet;
}

function sheetData(sheet, startCol, numCols) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, startCol, last - 1, numCols).getValues();
}

// ══════════════════════════════════════════════════════════════
// Categories
// ══════════════════════════════════════════════════════════════

function getCategories(ss, type) {
  const sheet = getOrCreateSheet(ss, SH_CATEGORIES, HDR_CATEGORIES);

  if (sheet.getLastRow() < 2) {
    const rows = [];
    DEF_OUTFLOW.forEach((n, i) => rows.push([n, 'outflow', i + 1]));
    DEF_INFLOW.forEach((n, i)  => rows.push([n, 'inflow',  i + 1]));
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  const rows = sheetData(sheet, 1, 3);
  const cats = rows
    .filter(r => r[0] && String(r[1]) === type)
    .sort((a, b) => (Number(a[2]) || 0) - (Number(b[2]) || 0))
    .map(r => String(r[0]).trim());

  return json({ status: 'ok', categories: cats, type });
}

function addCategory(ss, p) {
  const name = String(p.name || '').trim();
  const type = String(p.type || 'outflow');
  if (!name) return err('name required');

  const sheet = getOrCreateSheet(ss, SH_CATEGORIES, HDR_CATEGORIES);
  const rows  = sheetData(sheet, 1, 2);

  const dup = rows.find(r => String(r[1]) === type && String(r[0]).trim().toLowerCase() === name.toLowerCase());
  if (dup) return json({ status: 'duplicate', existing: String(dup[0]).trim() });

  const orders = sheetData(sheet, 1, 3).filter(r => String(r[1]) === type).map(r => Number(r[2]) || 0);
  const nextOrder = orders.length ? Math.max(...orders) + 1 : 1;

  sheet.appendRow([name, type, nextOrder]);
  return json({ status: 'ok', name, type });
}

function renameCategory(ss, p) {
  const oldName = String(p.oldName || '').trim();
  const newName = String(p.newName || '').trim();
  const type    = String(p.type || 'outflow');
  if (!oldName || !newName) return err('oldName and newName required');

  const sheet = ss.getSheetByName(SH_CATEGORIES);
  if (!sheet) return err('Categories sheet not found');

  const rows = sheetData(sheet, 1, 2);
  const dup = rows.find(r =>
    String(r[1]) === type &&
    String(r[0]).trim().toLowerCase() === newName.toLowerCase() &&
    String(r[0]).trim() !== oldName
  );
  if (dup) return json({ status: 'duplicate', existing: String(dup[0]).trim() });

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]) === type && String(rows[i][0]).trim() === oldName) {
      sheet.getRange(i + 2, 1).setValue(newName);
      return json({ status: 'ok', newName });
    }
  }
  return err('Category not found');
}

// ══════════════════════════════════════════════════════════════
// Accounts & Subcategories (ordered lists, auto-populated)
// ══════════════════════════════════════════════════════════════

function getListSheet(ss, sheetName, headers) {
  const sheet = getOrCreateSheet(ss, sheetName, headers);
  const rows  = sheetData(sheet, 1, 2);
  const values = rows
    .filter(r => r[0])
    .sort((a, b) => (Number(a[1]) || 0) - (Number(b[1]) || 0))
    .map(r => String(r[0]).trim());
  return json({ status: 'ok', values });
}

// Migrate old Subcategories sheet (Name, Order) → new (Name, Category, Order)
function migrateSubcatsSheet(ss) {
  const sheet = ss.getSheetByName(SH_SUBCATS);
  if (!sheet || sheet.getLastRow() === 0) return;
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // Already migrated if 3+ columns
  if (hdr.length >= 3) return;
  // Insert Category column at position 2
  sheet.insertColumnAfter(1);
  sheet.getRange(1, 2).setValue('Category');
}

function getSubcats(ss) {
  migrateSubcatsSheet(ss);
  const sheet = getOrCreateSheet(ss, SH_SUBCATS, HDR_SUBCATS);
  const rows  = sheetData(sheet, 1, 3);
  const result = {};
  rows
    .filter(r => r[0])
    .sort((a, b) => (Number(a[2]) || 0) - (Number(b[2]) || 0))
    .forEach(r => {
      const name = String(r[0]).trim();
      const cat  = String(r[1]).trim();
      if (!cat) return;
      if (!result[cat]) result[cat] = [];
      result[cat].push(name);
    });
  return json({ status: 'ok', subcats: result });
}

function addListItemAction(ss, p) {
  const name = String(p.name || '').trim();
  const list = String(p.list || '');
  if (!name) return err('name required');

  if (list === 'accounts' || list === 'txnTypes') {
    const shName = list === 'accounts' ? SH_ACCOUNTS : SH_TXNTYPES;
    const hdr    = list === 'accounts' ? HDR_ACCOUNTS : HDR_TXNTYPES;
    const sheet  = getOrCreateSheet(ss, shName, hdr);
    const rows   = sheetData(sheet, 1, 1);
    const dup    = rows.find(r => String(r[0]).trim().toLowerCase() === name.toLowerCase());
    if (dup) return json({ status: 'duplicate', existing: String(dup[0]).trim() });
    const orders  = sheetData(sheet, 1, 2).map(r => Number(r[1]) || 0);
    const nextOrder = orders.length ? Math.max(...orders) + 1 : 1;
    sheet.appendRow([name, nextOrder]);
  } else {
    migrateSubcatsSheet(ss);
    const category = String(p.category || '').trim();
    if (!category) return err('category required for subcats');
    const sheet = getOrCreateSheet(ss, SH_SUBCATS, HDR_SUBCATS);
    const rows  = sheetData(sheet, 1, 2);
    const dup   = rows.find(r =>
      String(r[0]).trim().toLowerCase() === name.toLowerCase() &&
      String(r[1]).trim() === category
    );
    if (dup) return json({ status: 'duplicate', existing: String(dup[0]).trim() });
    const orders  = sheetData(sheet, 1, 3).map(r => Number(r[2]) || 0);
    const nextOrder = orders.length ? Math.max(...orders) + 1 : 1;
    sheet.appendRow([name, category, nextOrder]);
  }
  return json({ status: 'ok', name });
}

function renameListItemAction(ss, p) {
  const oldName = String(p.oldName || '').trim();
  const newName = String(p.newName || '').trim();
  const list    = String(p.list || '');
  if (!oldName || !newName) return err('oldName and newName required');

  if (list === 'accounts' || list === 'txnTypes') {
    const shName = list === 'accounts' ? SH_ACCOUNTS : SH_TXNTYPES;
    const sheet  = ss.getSheetByName(shName);
    if (!sheet) return err('Sheet not found');
    const rows = sheetData(sheet, 1, 1);
    const dup  = rows.find(r => String(r[0]).trim().toLowerCase() === newName.toLowerCase() && String(r[0]).trim() !== oldName);
    if (dup) return json({ status: 'duplicate', existing: String(dup[0]).trim() });
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === oldName) {
        sheet.getRange(i + 2, 1).setValue(newName);
        return json({ status: 'ok', newName });
      }
    }
    return err('Item not found');
  } else {
    migrateSubcatsSheet(ss);
    const category = String(p.category || '').trim();
    const sheet = ss.getSheetByName(SH_SUBCATS);
    if (!sheet) return err('Sheet not found');
    const rows = sheetData(sheet, 1, 2);
    const dup  = rows.find(r =>
      String(r[0]).trim().toLowerCase() === newName.toLowerCase() &&
      String(r[1]).trim() === category &&
      String(r[0]).trim() !== oldName
    );
    if (dup) return json({ status: 'duplicate', existing: String(dup[0]).trim() });
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === oldName && String(rows[i][1]).trim() === category) {
        sheet.getRange(i + 2, 1).setValue(newName);
        return json({ status: 'ok', newName });
      }
    }
    return err('Item not found');
  }
}

function upsertListItem(ss, sheetName, headers, name, category) {
  if (!name) return;
  const sheet = getOrCreateSheet(ss, sheetName, headers);
  if (sheetName === SH_SUBCATS) {
    const cat  = String(category || '').trim();
    const rows = sheetData(sheet, 1, 2);
    const exists = rows.some(r =>
      String(r[0]).trim().toLowerCase() === name.toLowerCase() &&
      String(r[1]).trim() === cat
    );
    if (exists) return;
    const orders = sheetData(sheet, 1, 3).map(r => Number(r[2]) || 0);
    const nextOrder = orders.length ? Math.max(...orders) + 1 : 1;
    sheet.appendRow([name, cat, nextOrder]);
  } else {
    const rows  = sheetData(sheet, 1, 1);
    const exists = rows.some(r => String(r[0]).trim().toLowerCase() === name.toLowerCase());
    if (exists) return;
    const orders = sheetData(sheet, 1, 2).map(r => Number(r[1]) || 0);
    const nextOrder = orders.length ? Math.max(...orders) + 1 : 1;
    sheet.appendRow([name, nextOrder]);
  }
}

// ══════════════════════════════════════════════════════════════
// Users
// ══════════════════════════════════════════════════════════════

function upsertUser(ss, firstName, lastName, username, appUserName) {
  const sheet = getOrCreateSheet(ss, SH_USERS, HDR_USERS);
  
  // Use appUserName if provided, otherwise fallback to TG name
  let nameKey = appUserName;
  if (!nameKey) {
    nameKey = firstName
      ? [firstName, lastName].filter(Boolean).join(' ')
      : (username ? '@' + username : 'Unknown');
  }
  
  const idStr = String(nameKey).trim();
  if (!idStr) return 'Unknown';

  const rows = sheetData(sheet, 1, HDR_USERS.length);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === idStr) {
      // Update metadata (First name, Last name, Username)
      sheet.getRange(i + 2, 2, 1, 3).setValues([[firstName || '', lastName || '', username || '']]);
      return idStr;
    }
  }

  // Name · First name · Last name · Username · Date added
  sheet.appendRow([idStr, firstName || '', lastName || '', username || '', new Date()]);
  return idStr;
}

// ══════════════════════════════════════════════════════════════
// Records
// ══════════════════════════════════════════════════════════════

function addRecord(ss, p) {
  const sheet       = getOrCreateSheet(ss, SH_OUTFLOWS, HDR_OUTFLOWS);
  const firstName   = String(p.firstName    || '');
  const lastName    = String(p.lastName     || '');
  const username    = String(p.username     || '');
  const type        = String(p.type         || 'outflow');
  const subcategory = String(p.subcategory  || '').trim();
  const account     = String(p.account      || '').trim();
  const txnType  = String(p.txnType        || '').trim();

  const userName = upsertUser(ss, firstName, lastName, username, p.userName);

  const category = String(p.category || '').trim();
  if (account)     upsertListItem(ss, SH_ACCOUNTS, HDR_ACCOUNTS, account);
  if (txnType)       upsertListItem(ss, SH_TXNTYPES,   HDR_TXNTYPES,   txnType);
  if (subcategory) { migrateSubcatsSheet(ss); upsertListItem(ss, SH_SUBCATS, HDR_SUBCATS, subcategory, category); }

  let purchaseDate = '';
  let purchaseMonth = '';
  if (p.date) {
    const [y, m, d] = String(p.date).split('-').map(Number);
    if (y && m && d) {
      purchaseDate  = new Date(y, m - 1, d);
      purchaseMonth = m;
    }
  }

  const amount = p.amount !== undefined ? Number(p.amount) : '';

  // Prepare row data based on RECORD_COLUMNS configuration
  const rowData = RECORD_COLUMNS.map(key => {
    switch(key) {
      case 'timestamp':   return new Date();
      case 'month':       return purchaseMonth;
      case 'date':        return purchaseDate;
      case 'amount':      return amount;
      case 'txnType':       return txnType;
      case 'category':    return p.category || '';
      case 'subcategory': return subcategory;
      case 'description': return p.note || '';
      case 'account':     return account;
      case 'name':        return userName;
      case 'type':        return type;
      default:            return '';
    }
  });

  sheet.appendRow(rowData);
  return json({ status: 'ok' });
}

function checkDuplicate(ss, p) {
  const sheet = ss.getSheetByName(SH_OUTFLOWS);
  if (!sheet || sheet.getLastRow() < 2) return json({ status: 'ok', duplicate: false });

  const today       = String(p.date        || '');
  const amount      = Number(p.amount      || 0);
  const category    = String(p.category    || '').toLowerCase();
  const subcategory = String(p.subcategory || '').trim().toLowerCase();
  const note        = String(p.note        || '').trim().toLowerCase();
  const account     = String(p.account     || '').trim().toLowerCase();
  const txnType  = String(p.txnType       || '').trim().toLowerCase();
  const userName    = String(p.userName     || '');
  const type        = String(p.type        || 'outflow');
  const tz          = Session.getScriptTimeZone();

  const rows = sheetData(sheet, 1, HDR_OUTFLOWS.length);
  let count = 0;

  const colIdxMap = {};
  RECORD_COLUMNS.forEach((key, i) => colIdxMap[key] = i);

  for (const row of rows) {
    const d = row[colIdxMap['date']];
    let dateStr = '';
    if (d instanceof Date && !isNaN(d)) dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    else dateStr = String(d).substring(0, 10);

    if (
      dateStr === today &&
      Math.abs(Number(row[colIdxMap['amount']])) === Math.abs(amount) &&
      String(row[colIdxMap['category']] || '').toLowerCase() === category &&
      String(row[colIdxMap['subcategory']] || '').trim().toLowerCase() === subcategory &&
      String(row[colIdxMap['description']] || '').trim().toLowerCase() === note &&
      String(row[colIdxMap['account']] || '').trim().toLowerCase() === account &&
      (colIdxMap['txnType'] === undefined || String(row[colIdxMap['txnType']] || '').trim().toLowerCase() === txnType) &&
      String(row[colIdxMap['type']] || 'outflow') === type &&
      (!userName || String(row[colIdxMap['name']]) === userName)
    ) count++;
  }

  return json({ status: 'ok', duplicate: count > 0, count });
}

// ══════════════════════════════════════════════════════════════
// Summary
// ══════════════════════════════════════════════════════════════

function prevPeriod(period, groupBy) {
  if (groupBy === 'year') {
    return String(Number(period) - 1);
  }
  const [y, m] = period.split('-').map(Number);
  if (m === 1) return `${y - 1}-12`;
  return `${y}-${String(m - 1).padStart(2, '0')}`;
}

function daysInPeriod(period, groupBy) {
  if (groupBy === 'year') {
    const y = Number(period);
    return ((y % 4 === 0 && y % 100 !== 0) || y % 400 === 0) ? 366 : 365;
  }
  const [y, m] = period.split('-').map(Number);
  return new Date(y, m, 0).getDate();
}

function getSummary(ss, period, groupBy, viewBy) {
  const sheet = ss.getSheetByName(SH_OUTFLOWS);
  if (!sheet || sheet.getLastRow() < 2)
    return json({ status: 'ok', period, groupBy, viewBy, outflows: [], income: [],
                  prevOutflowTotal: 0, prevIncomeTotal: 0, daysInPeriod: daysInPeriod(period, groupBy) });

  const tz   = Session.getScriptTimeZone();
  const rows = sheetData(sheet, 1, HDR_OUTFLOWS.length);

  const colIdxMap = {};
  RECORD_COLUMNS.forEach((key, i) => colIdxMap[key] = i);

  let colIdx = colIdxMap['category'];
  if (viewBy === 'subcategory') colIdx = colIdxMap['subcategory'];
  if (viewBy === 'account')     colIdx = colIdxMap['account'];

  const prev = prevPeriod(period, groupBy);

  // {label: {total, count}}
  const outMap = {}, inMap = {};
  let prevOutTotal = 0, prevInTotal = 0;

  for (const row of rows) {
    const d = row[colIdxMap['date']];
    if (!d) continue;
    let key = '';
    if (d instanceof Date && !isNaN(d))
      key = Utilities.formatDate(d, tz, groupBy === 'year' ? 'yyyy' : 'yyyy-MM');
    else
      key = String(d).substring(0, groupBy === 'year' ? 4 : 7);

    const amount = Number(row[colIdxMap['amount']]) || 0;
    const type   = String(row[colIdxMap['type']] || 'outflow');

    if (key === prev) {
      if (type === 'inflow') prevInTotal  += amount;
      else                   prevOutTotal += amount;
      continue;
    }

    if (key !== period) continue;

    const val = String(row[colIdx] || 'Other');
    if (type === 'inflow') {
      if (!inMap[val])  inMap[val]  = { total: 0, count: 0 };
      inMap[val].total  += amount; inMap[val].count++;
    } else {
      if (!outMap[val]) outMap[val] = { total: 0, count: 0 };
      outMap[val].total += amount; outMap[val].count++;
    }
  }

  const sorted = obj =>
    Object.entries(obj)
      .map(([label, v]) => ({ label, total: v.total, count: v.count }))
      .sort((a, b) => b.total - a.total);

  return json({
    status: 'ok', period, groupBy, viewBy,
    outflows: sorted(outMap),
    income:   sorted(inMap),
    prevOutflowTotal: prevOutTotal,
    prevIncomeTotal:  prevInTotal,
    daysInPeriod:     daysInPeriod(period, groupBy)
  });
}

// ══════════════════════════════════════════════════════════════
// Trend (last N periods)
// ══════════════════════════════════════════════════════════════

function getTrend(ss, endPeriod, count, groupBy) {
  const sheet = ss.getSheetByName(SH_OUTFLOWS);
  if (!sheet || sheet.getLastRow() < 2)
    return json({ status: 'ok', trend: [] });

  // Build list of periods from oldest to newest
  const periods = [];
  let p = endPeriod;
  for (let i = 0; i < count; i++) {
    periods.unshift(p);
    p = prevPeriod(p, groupBy);
  }

  const periodSet = new Set(periods);
  const totals = {};
  periods.forEach(pd => { totals[pd] = { outflow: 0, income: 0 }; });

  const tz   = Session.getScriptTimeZone();
  const rows = sheetData(sheet, 1, HDR_OUTFLOWS.length);
  const colIdxMap = {};
  RECORD_COLUMNS.forEach((key, i) => colIdxMap[key] = i);

  for (const row of rows) {
    const d = row[colIdxMap['date']];
    if (!d) continue;
    let key = '';
    if (d instanceof Date && !isNaN(d))
      key = Utilities.formatDate(d, tz, groupBy === 'year' ? 'yyyy' : 'yyyy-MM');
    else
      key = String(d).substring(0, groupBy === 'year' ? 4 : 7);

    if (!periodSet.has(key)) continue;

    const amount = Number(row[colIdxMap['amount']]) || 0;
    const type   = String(row[colIdxMap['type']] || 'outflow');
    if (type === 'inflow') totals[key].income  += amount;
    else                   totals[key].outflow += amount;
  }

  const trend = periods.map(pd => ({ period: pd, outflow: totals[pd].outflow, income: totals[pd].income }));
  return json({ status: 'ok', trend });
}

