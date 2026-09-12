const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const html = fs.readFileSync(new URL('../index.html', `file://${__filename}`), 'utf8');
const source = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)]
  .map(match => match[1]).filter(Boolean).at(-1);
const element = () => ({ classList: { add() {}, remove() {} }, style: {}, value: '', textContent: '', innerHTML: '', appendChild() {} });
const store = new Map();
const context = vm.createContext({
  console,
  alert() {},
  setTimeout(fn) { fn(); },
  clearTimeout() {},
  localStorage: {
    getItem(key) { return store.get(key) ?? null; },
    setItem(key, value) { store.set(key, String(value)); },
    removeItem(key) { store.delete(key); }
  },
  document: {
    addEventListener() {},
    getElementById() { return element(); },
    querySelectorAll() { return []; },
    createElement() { return element(); },
    body: { appendChild() {}, removeChild() {} }
  },
  URL: { createObjectURL() { return ''; }, revokeObjectURL() {} },
  Blob: class Blob {}
});
vm.runInContext(source, context);

assert.equal(vm.runInContext("/^6T19Q\\s*X$/.test('6T19Q X')", context), true);
assert.equal(vm.runInContext("/^6T19Q\\s*X$/.test('6T19QX')", context), true);
assert.equal(vm.runInContext("/^6T19Q\\s*X$/.test('6T19Q   X')", context), true);
assert.equal(vm.runInContext("/^6T19Q\\s*X$/.test('6T19Q *X')", context), false);

const rows = Array.from({ length: 10 }, () => Array(90).fill(''));
const start = new Date(2026, 0, 4);
for (let day = 0; day < 42; day++) {
  const date = new Date(start);
  date.setDate(start.getDate() + day);
  rows[7][6 + day * 2] = `${date.getMonth() + 1}/${date.getDate()}\n/${date.getFullYear()}`;
}
rows[8][0] = 'Cruz, Rowena';
rows[8][5] = 'AMGR-FT (AM), EB, ES, CN';
rows[8][7] = '6:45 AM x\nC 12:00';
rows[8][9] = 'PTO';
rows[8][11] = '6T19Q';
rows[8][13] = 'SICK';
context.XLSX = { utils: { sheet_to_json() { return rows; } } };
vm.runInContext("parseMyTimeExport({Sheets:{Sheet1:{}},SheetNames:['Sheet1']})", context);
const shifts = vm.runInContext("staffRoster.find(s => s.name === 'Cruz, Rowena').shifts", context);
const rawShifts = vm.runInContext("staffRoster.find(s => s.name === 'Cruz, Rowena').rawShifts", context);
const dates = vm.runInContext('rawData.dates.map(d => d.date)', context);
assert.equal(shifts[dates[0]], 'ANM_CN');
assert.equal(rawShifts[dates[0]], '6:45 AM x\nC 12:00');
assert.equal(shifts[dates[1]], 'PTO');
assert.equal(shifts[dates[2]], 'ANM_ADMIN');
assert.equal(shifts[dates[3]], 'SICK');
assert.equal(vm.runInContext("extractShiftFromBlock(['R 12:00\\n18t7q x'], false)", context), '7P');
assert.equal(vm.runInContext("extractShiftFromBlock(['6t19q\\nR 12:00'], false)", context), '7A');
assert.equal(vm.runInContext("extractShiftFromBlock(['Teach x\\nE 12:00'], false)", context), 'EDU');
assert.equal(vm.runInContext("parseShiftCode('7:30 AM', true)", context), 'ANM_ADMIN');
assert.equal(vm.runInContext("parseShiftCode('6:45 AM x', true)", context), 'ANM_CN');
assert.equal(vm.runInContext("countsTowardCommitment('7A')", context), true);
assert.equal(vm.runInContext("countsTowardCommitment('PTO')", context), true);
assert.equal(vm.runInContext("countsTowardCommitment('SICK')", context), false);
assert.equal(vm.runInContext("countsTowardCommitment('R/O')", context), false);
assert.equal(vm.runInContext("countsTowardCommitment('ADMIN')", context), false);
assert.equal(vm.runInContext("countsTowardCommitment('PAYRO')", context), false);

const fairnessOrder = vm.runInContext(`getFairnessRanking(
  ['Nurse, Never','Nurse, Recent','Nurse, Oldest'], 'PTO', [
    {employee:'Nurse, Recent',requestType:'PTO',decision:'Granted',grantedDate:'2026-08-01'},
    {employee:'Nurse, Oldest',requestType:'PTO',decision:'Granted',grantedDate:'2025-01-01'}
  ]
).map(item => item.name)`, context);
assert.deepEqual([...fairnessOrder], ['Nurse, Never','Nurse, Oldest','Nurse, Recent']);

assert.match(html, /Show only employees scheduled in this uploaded period/);
assert.match(html, /function recordRequestDecision\(/);
assert.match(html, /'Open Needs'/);
assert.match(html, /'Weekly Shift Review'/);
assert.match(html, /'Weekend Equity'/);
assert.match(html, /'Request History'/);
assert.match(html, /state:'frozen',xSplit:4,ySplit:2/);

console.log('audit-fixes tests passed');
