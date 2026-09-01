// Self-check for the History tab's backend: a record is addressed by its row
// number, and a row number is not an identity — rows shift whenever the
// spreadsheet is edited by hand or by a second client. Every edit therefore
// carries the signature the row had when it was read, and the server must
// refuse the write when the row no longer matches it.
// Run: node test/record-editing.test.js
const fs = require('fs');
const vm = require('vm');

// ── Apps Script services, reduced to what these functions call ──────────
const sandbox = {
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  ContentService: {
    MimeType: { JSON: 'json' },
    createTextOutput: s => ({ setMimeType: () => JSON.parse(s) }),
  },
  Session: { getScriptTimeZone: () => 'UTC' },
  Utilities: {
    formatDate: (d, _tz, fmt) => {
      const p = n => String(n).padStart(2, '0');
      return fmt === 'HH:mm'
        ? `${p(d.getHours())}:${p(d.getMinutes())}`
        : `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
    },
  },
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(__dirname + '/../apps-script/Code.gs', 'utf8'), sandbox);

const { rowSignature, getRecords, updateRecord, deleteRecord, RECORD_COLUMNS, COLUMN_DISPLAY_NAMES } =
  vm.runInContext('({ rowSignature, getRecords, updateRecord, deleteRecord, RECORD_COLUMNS, COLUMN_DISPLAY_NAMES })', sandbox);

const idx = {};
RECORD_COLUMNS.forEach((key, i) => (idx[key] = i));

/** Minimal stand-in for a Sheet, backed by a 2D array whose row 0 is the header. */
const makeSheet = rows => ({
  rows,
  getLastRow: () => rows.length,
  getLastColumn: () => rows[0].length,
  getMaxColumns: () => rows[0].length,
  getRange: (r, c, nr, nc) => ({
    getValues: () => rows.slice(r - 1, r - 1 + nr).map(row => row.slice(c - 1, c - 1 + nc)),
    setValues: v => v.forEach((row, i) => (rows[r - 1 + i] = row.slice())),
  }),
  deleteRow: r => rows.splice(r - 1, 1),
});

const record = ({ ts, day, amount, category, type = 'Outflow', name = 'Kim' }) => {
  const row = new Array(RECORD_COLUMNS.length).fill('');
  row[idx.timestamp]   = ts;
  row[idx.date]        = day;
  row[idx.amount]      = amount;
  row[idx.category]    = category;
  row[idx.description] = '';
  row[idx.name]        = name;
  row[idx.type]        = type;
  return row;
};

const freshSheet = () => makeSheet([
  RECORD_COLUMNS.map(k => COLUMN_DISPLAY_NAMES[k] || k),
  record({ ts: new Date(2026, 7, 24, 9, 5),  day: new Date(2026, 7, 24), amount: 100, category: 'Rent' }),
  record({ ts: new Date(2026, 7, 25, 12, 0), day: new Date(2026, 7, 25), amount: 200, category: 'Food' }),
  record({ ts: new Date(2026, 7, 25, 18, 30), day: new Date(2026, 7, 25), amount: 300, category: 'Taxi' }),
  record({ ts: new Date(2026, 7, 26, 8, 0),  day: new Date(2026, 7, 26), amount: 900, category: 'Salary', type: 'Inflow' }),
]);
const asSpreadsheet = sheet => ({ getSheetByName: name => (name === 'Records' ? sheet : null) });

let failed = 0;
const check = (name, actual, expected) => {
  const [a, e] = [JSON.stringify(actual), JSON.stringify(expected)];
  if (a !== e) {
    console.error(`FAIL: ${name}\n  expected ${e}\n  got      ${a}`);
    failed++;
  }
};

// ── Listing ─────────────────────────────────────────────────────────────
let sheet = freshSheet();
const page = getRecords(asSpreadsheet(sheet), { type: 'outflow', limit: 2, offset: 0 });
check('outflows come back newest entry first', page.records.map(r => r.category), ['Taxi', 'Food']);
check('a full page reports there is more', page.hasMore, true);
check('rows are addressed by their sheet row number', page.records.map(r => r.row), [4, 3]);
check('the inflow of the same day is not listed among the outflows',
  page.records.every(r => r.type === 'outflow'), true);

const rest = getRecords(asSpreadsheet(sheet), { type: 'outflow', limit: 2, offset: 2 });
check('the next page continues where the first stopped, without repeating it',
  rest.records.map(r => r.category), ['Rent']);
check('and reports the end of the list', rest.hasMore, false);
check('the type filter selects inflows too',
  getRecords(asSpreadsheet(sheet), { type: 'inflow' }).records.map(r => r.category), ['Salary']);
check('dates and entry times are formatted for the app',
  [page.records[0].date, page.records[0].time], ['2026-08-25', '18:30']);

// A purchase entered days after it happened files under the day it happened,
// while the stamp the History tab prints under the amount belongs to the row.
const lateSheet = makeSheet([
  RECORD_COLUMNS.map(k => COLUMN_DISPLAY_NAMES[k] || k),
  record({ ts: new Date(2026, 8, 1, 14, 32), day: new Date(2026, 7, 20), amount: 50, category: 'Rent' }),
]);
const late = getRecords(asSpreadsheet(lateSheet), { type: 'outflow' }).records[0];
check('the entry stamp carries its own date, not the purchase day',
  [late.date, late.entryDate, late.time], ['2026-08-20', '2026-09-01', '14:32']);

// ── The guard: a row number alone must not be trusted ───────────────────
const taxi = page.records[0]; // row 4
check('the signature identifies the row it was read from',
  rowSignature(sheet.rows[taxi.row - 1], idx), taxi.sig);
check('and does not match its neighbour',
  rowSignature(sheet.rows[taxi.row - 2], idx) === taxi.sig, false);

sheet.deleteRow(3); // someone removes "Food" in the spreadsheet: Taxi is now row 3
check('deleting by a row number that has shifted is refused',
  deleteRecord(asSpreadsheet(sheet), { row: taxi.row, sig: taxi.sig }).status, 'stale');
check('and the record that moved into that row survives',
  sheet.rows.map(r => r[idx.category]), ['Category 2', 'Rent', 'Taxi', 'Salary']);

check('a row past the end of the sheet is refused',
  deleteRecord(asSpreadsheet(sheet), { row: 99, sig: taxi.sig }).status, 'stale');

// ── The happy paths ─────────────────────────────────────────────────────
sheet = freshSheet();
const food = getRecords(asSpreadsheet(sheet), { type: 'outflow' }).records[1];
check('an unchanged row deletes', deleteRecord(asSpreadsheet(sheet), food).status, 'ok');
check('and it is the picked record that is gone',
  sheet.rows.map(r => r[idx.category]), ['Category 2', 'Rent', 'Taxi', 'Salary']);

sheet = freshSheet();
const taxi2 = getRecords(asSpreadsheet(sheet), { type: 'outflow' }).records[0];
const stamp = sheet.rows[taxi2.row - 1][idx.timestamp];
check('an unchanged row updates',
  updateRecord(asSpreadsheet(sheet), {
    row: taxi2.row, sig: taxi2.sig,
    date: '2026-08-25', amount: 350, category: 'Transport', note: 'airport', type: 'outflow',
  }).status, 'ok');
check('the edited fields are written',
  [sheet.rows[taxi2.row - 1][idx.amount], sheet.rows[taxi2.row - 1][idx.category],
   sheet.rows[taxi2.row - 1][idx.description]], [350, 'Transport', 'airport']);
check('while who entered the row, and when, is left alone',
  [sheet.rows[taxi2.row - 1][idx.timestamp] === stamp, sheet.rows[taxi2.row - 1][idx.name]], [true, 'Kim']);
check('an edit sent twice is refused the second time, the row no longer matching',
  updateRecord(asSpreadsheet(sheet), {
    row: taxi2.row, sig: taxi2.sig, date: '2026-08-25', amount: 1, category: 'Transport', type: 'outflow',
  }).status, 'stale');

if (failed) process.exit(1);
console.log('record editing: 20 checks OK');
