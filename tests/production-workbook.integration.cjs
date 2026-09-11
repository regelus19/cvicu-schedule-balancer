const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const rows = JSON.parse(fs.readFileSync(0, 'utf8'));
const html = fs.readFileSync(new URL('../index.html', `file://${__filename}`), 'utf8');
const source = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)].map(match => match[1]).filter(Boolean).at(-1);
const element = () => ({ classList: { add() {}, remove() {} }, style: {}, value: '', textContent: '', innerHTML: '', appendChild() {} });
const store = new Map();
const warnings = [];
const context = vm.createContext({
  console,
  alert(message) { warnings.push(String(message)); },
  setTimeout(fn) { fn(); }, clearTimeout() {},
  localStorage: { getItem(key) { return store.get(key) ?? null; }, setItem(key, value) { store.set(key, String(value)); } },
  document: { addEventListener() {}, getElementById() { return element(); }, querySelectorAll() { return []; }, createElement() { return element(); }, body: { appendChild() {}, removeChild() {} } },
  URL: { createObjectURL() { return ''; }, revokeObjectURL() {} }, Blob: class Blob {},
  XLSX: { utils: { sheet_to_json() { return rows; } } }
});
vm.runInContext(source, context);
vm.runInContext("parseMyTimeExport({Sheets:{Report:{}},SheetNames:['Report']})", context);

const result = vm.runInContext(`({
  dates: rawData.dates,
  staff: staffRoster.map(s => ({name:s.name, job:s.job, type:s.type, skill:s.skill, shift:s.shift, shifts:s.shifts})),
  classifications: staffRoster.flatMap(s => Object.values(s.shifts)),
  late: staffRoster.filter(s => detectLateSubmitter(s)).map(s => s.name)
})`, context);
assert.equal(result.dates.length, 42);
assert.equal(result.staff.length, 118);
assert.equal(result.dates[0].date, '10/4/ 2026');
assert.equal(result.dates.at(-1).date, '11/14/ 2026');
assert.ok(result.classifications.includes('ANM_CN'));
assert.ok(result.classifications.includes('ANM_ADMIN'));
assert.ok(result.classifications.includes('EDU'));
assert.ok(!warnings.some(message => message.startsWith('MyTime import warning:')));

const rowena = result.staff.find(staff => staff.name === 'Cruz, Rowena');
const bruce = result.staff.find(staff => staff.name === 'Bruce, Jessica');
assert.ok(rowena && Object.values(rowena.shifts).includes('ANM_CN'));
assert.ok(bruce && Object.values(bruce.shifts).includes('ANM_ADMIN'));

const balanced = vm.runInContext(`(() => {
  moveTracker = {}; moveCategories = {}; addedShiftsTracker = {};
  const day = balanceDayShift();
  const night = balanceNightShift();
  balancedResults = {day, night};
  const unresolved = getUnresolvedDeficits();
  return {
    dayMoves: day.moveList,
    nightMoves: night.moveList,
    dayDates: day.dates.length,
    nightDates: night.dates.length,
    daySimulation: day.simulation,
    nightSimulation: night.simulation,
    roster: staffRoster.map(s => ({name:s.name, job:s.job, type:s.type, skill:s.skill, shift:s.shift})),
    dates: day.dates,
    requirements: {...requirements},
    unresolved
  };
})()`, context);
assert.equal(balanced.dayDates, 42);
assert.equal(balanced.nightDates, 42);
assert.ok(!balanced.dayMoves.concat(balanced.nightMoves).some(move => ['Cruz, Rowena', 'Rodriguez, Althea Nadine', 'Bruce, Jessica', 'Sala, Jose F', 'Wells, Sydney N'].includes(move.name)));
assert.ok(balanced.unresolved.totalPositions > 0);
assert.equal(balanced.requirements.CN, 2);
assert.ok(vm.runInContext('generateUnresolvedDeficitsHTML(getUnresolvedDeficits())', context).includes('Draft Projection'));
assert.ok(vm.runInContext("scheduleStatus='closed'; generateUnresolvedDeficitsHTML(getUnresolvedDeficits())", context).includes('Partially Balanced'));

const coverage = (simulation, mode) => balanced.dates.map(date => {
  let rn = 0, pct = 0, cn = 0, eb = 0, es = 0, mid = 0;
  for (const staff of balanced.roster) {
    const token = simulation[staff.name]?.[date] || '';
    const counts = mode === 'day' ? ['7A', 'MID', 'ANM_CN'].includes(token) : token === '7P';
    if (!counts) continue;
    if (token === 'ANM_CN') { rn++; cn++; continue; }
    if (staff.job === 'RN') {
      rn++;
      if (staff.skill?.includes('CN')) cn++;
      if (staff.skill?.includes('EB')) eb++;
      if (staff.skill?.includes('ES')) es++;
    } else if (staff.job === 'PCT' || staff.job === 'HUC') pct++;
    if (token === 'MID') mid++;
  }
  const parsed = new Date(date.replace('/ ', '/'));
  const day = parsed.getDay();
  const weekend = mode === 'day' ? day === 0 || day === 6 : [0, 5, 6].includes(day);
  const rnMin = weekend ? balanced.requirements.RN_WE : balanced.requirements.RN_WD;
  const pctMin = weekend ? balanced.requirements.PCT_WE : balanced.requirements.PCT_WD;
  const rnMax = rnMin + (day === 6 ? balanced.requirements.CAP_SAT : day === 0 ? balanced.requirements.CAP_SUN : day === 4 ? balanced.requirements.CAP_THU : balanced.requirements.CAP_WD);
  const midMin = weekend ? balanced.requirements.MID_WE : balanced.requirements.MID_WD;
  cn = Math.min(cn, 4);
  return {date, rn, pct, cn, eb, es, mid, rnMin, rnMax, pctMin, midMin};
});
const dayCoverage = coverage(balanced.daySimulation, 'day');
const nightCoverage = coverage(balanced.nightSimulation, 'night');
const originalSimulation = Object.fromEntries(result.staff.map(staff => [staff.name, {...staff, ...staff.shifts}]));
const originalDayCoverage = coverage(originalSimulation, 'day');
const originalNightCoverage = coverage(originalSimulation, 'night');
const summarize = (rows, mode) => ({
  rnGapDays: rows.filter(r => r.rn < r.rnMin).length,
  rnOverCapDays: rows.filter(r => r.rn > r.rnMax).length,
  pctGapDays: rows.filter(r => r.pct < r.pctMin).length,
  cnGapDays: rows.filter(r => r.cn < balanced.requirements.CN).length,
  ebGapDays: rows.filter(r => r.eb < balanced.requirements.EB).length,
  esGapDays: rows.filter(r => r.es < balanced.requirements.ES).length,
  midGapDays: mode === 'day' ? rows.filter(r => r.mid < r.midMin).length : 0
});
const daySummary = summarize(dayCoverage, 'day');
const nightSummary = summarize(nightCoverage, 'night');

console.log(JSON.stringify({
  dates: result.dates.length,
  staff: result.staff.length,
  tokenCounts: result.classifications.reduce((counts, token) => ((counts[token] = (counts[token] || 0) + 1), counts), {}),
  lateSubmitters: result.late.length,
  dayMoves: balanced.dayMoves.length,
  nightMoves: balanced.nightMoves.length,
  unresolvedPositions: balanced.unresolved.totalPositions,
  unresolvedRoleDays: balanced.unresolved.totalRoleDays,
  dayCoverage: daySummary,
  nightCoverage: nightSummary,
  originalDayCoverage: summarize(originalDayCoverage, 'day'),
  originalNightCoverage: summarize(originalNightCoverage, 'night'),
  newlyOverCapDayDates: dayCoverage.filter(row => row.rn > row.rnMax && !(originalDayCoverage.find(before => before.date === row.date).rn > row.rnMax)).map(row => row.date),
  newlyOverCapNightDates: nightCoverage.filter(row => row.rn > row.rnMax && !(originalNightCoverage.find(before => before.date === row.date).rn > row.rnMax)).map(row => row.date)
}, null, 2));
