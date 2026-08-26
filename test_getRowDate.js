// Self-check for getRowDate: purchase date ("date" column) must win over
// the row-creation "timestamp" column when both are present.
// Run: node test_getRowDate.js
const fs = require('fs');
const vm = require('vm');

const src = fs.readFileSync(__dirname + '/gas_script_v2.js', 'utf8');
const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(src, sandbox);

const colIdxMap = { timestamp: 0, date: 1 };

// timestamp = entry logged today, date = purchase made last month
const row = [new Date(2026, 7, 26), new Date(2026, 6, 15)];
const result = sandbox.getRowDate(row, colIdxMap);
console.assert(result.getMonth() === 6 && result.getFullYear() === 2026,
  `expected purchase date (Jul 2026) to win, got ${result}`);

// falls back to timestamp when date column is missing/invalid
const rowNoDate = [new Date(2026, 7, 26), ''];
const fallback = sandbox.getRowDate(rowNoDate, colIdxMap);
console.assert(fallback.getMonth() === 7 && fallback.getFullYear() === 2026,
  `expected fallback to timestamp (Aug 2026), got ${fallback}`);

console.log('getRowDate: OK');
