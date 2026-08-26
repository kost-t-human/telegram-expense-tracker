// Self-check for the date resolution the summary groups by.
// Run: node test/date-resolution.test.js
const fs = require('fs');
const vm = require('vm');

const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(__dirname + '/../apps-script/Code.gs', 'utf8'), sandbox);

const { getRowDate } = sandbox;
const colIdxMap = { timestamp: 0, date: 1 };
const ymd = d => d && `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

const cases = [
  // [name, timestamp cell, date cell, expected]
  ['purchase date wins over entry timestamp',
    new Date(2026, 7, 26), new Date(2026, 6, 15), '2026-07-15'],
  ['falls back to the timestamp when there is no purchase date',
    new Date(2026, 7, 26), '', '2026-08-26'],
  ['day-first text dates are read day-first, not month-first',
    new Date(2026, 7, 26), '07.03.2026', '2026-03-07'],
  ['slashed day-first text dates too',
    new Date(2026, 7, 26), '07/03/2026', '2026-03-07'],
  ['ISO text dates still parse',
    new Date(2026, 7, 26), '2026-03-07', '2026-03-07'],
];

let failed = 0;
for (const [name, timestamp, date, expected] of cases) {
  const actual = ymd(getRowDate([timestamp, date], colIdxMap));
  if (actual !== expected) {
    console.error(`FAIL: ${name}\n  expected ${expected}, got ${actual}`);
    failed++;
  }
}

if (failed) process.exit(1);
console.log(`getRowDate: ${cases.length} checks OK`);
