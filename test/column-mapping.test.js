// Self-check for column resolution: a "Records" sheet keeps the column order it
// was created with, which differs between deployments sharing this script.
// Reading it by RECORD_COLUMNS positions would take amounts and dates from the
// wrong columns.
// Run: node test/column-mapping.test.js
const fs = require('fs');
const vm = require('vm');

const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(__dirname + '/../apps-script/Code.gs', 'utf8'), sandbox);

// `const` declarations stay in the context's lexical scope, not on the sandbox
// object, so they have to be pulled out by evaluating them there.
const { recordColIdx, recordWidth, RECORD_COLUMNS, COLUMN_DISPLAY_NAMES } =
  vm.runInContext('({ recordColIdx, recordWidth, RECORD_COLUMNS, COLUMN_DISPLAY_NAMES })', sandbox);

/** Minimal stand-in for a Sheet: only the calls recordColIdx/recordWidth make. */
const fakeSheet = headers => ({
  getLastRow: () => headers.length ? 5 : 0,
  getLastColumn: () => headers.length,
  getMaxColumns: () => Math.max(headers.length, 26),
  getRange: () => ({ getValues: () => [headers] }),
});

const positional = {};
RECORD_COLUMNS.forEach((key, i) => (positional[key] = i));
const defaultHeaders = RECORD_COLUMNS.map(k => COLUMN_DISPLAY_NAMES[k] || k);

let failed = 0;
const check = (name, actual, expected) => {
  const [a, e] = [JSON.stringify(actual), JSON.stringify(expected)];
  if (a !== e) {
    console.error(`FAIL: ${name}\n  expected ${e}\n  got      ${a}`);
    failed++;
  }
};

check('a sheet this script created maps to its own layout',
  recordColIdx(fakeSheet(defaultHeaders)), positional);

// The layout of the second bot: Description and Category 3 the other way round.
const swapped = ['Timestamp', 'Day', 'Amount', 'Category 2', 'Description', 'Category 3', 'Account', 'Name', 'Type'];
const bySwapped = recordColIdx(fakeSheet(swapped));
check('a differently ordered sheet is read by its headers',
  [bySwapped.description, bySwapped.subcategory], [4, 5]);
check('the fields the summary needs still resolve',
  [bySwapped.date, bySwapped.amount, bySwapped.category, bySwapped.type], [1, 2, 3, 8]);

check('headers are matched regardless of case and padding',
  recordColIdx(fakeSheet(defaultHeaders.map(h => ' ' + h.toUpperCase() + ' '))), positional);

check('unmatched headers fall back to the configured order rather than guessing',
  recordColIdx(fakeSheet(['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i'])), positional);

check('an empty sheet falls back to the configured order',
  recordColIdx(fakeSheet([])), positional);

// A sheet carrying extra columns of the user's own must not be truncated.
check('rows span the whole sheet, including columns this script does not know',
  recordWidth(fakeSheet(defaultHeaders.concat(['Budget', 'Notes']))), RECORD_COLUMNS.length + 2);

if (failed) process.exit(1);
console.log('column mapping: 7 checks OK');
