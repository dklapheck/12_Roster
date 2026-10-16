const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const load = (ctx, file) => vm.runInContext(fs.readFileSync(path.join(__dirname, '..', file), 'utf8'), ctx);
const plain = x => JSON.parse(JSON.stringify(x));
const headers = ['Student Number', 'FIRST NAME', 'LAST NAME', 'Welcome Email', 'SCC Completion', 'ISMA'];

function fixture() {
  const cells = [[1, 'Test', 'One', 'Sent', 'Attempt 1', 'Sent'], [2, 'Test', 'Two', '', 'Completed', 'Lead Approved']];
  const formulas = cells.map(r => r.map(() => ''));
  let sourceRows = [[2, 'Test', 'Two', '', 'Completed', 'Lead Approved'], [1, 'Test', 'One', 'Follow-up', 'Completed', '']];
  let sourceHeaders = headers.slice(), failWrite = false;
  const writes = [], cache = new Map(), logs = [], dialogs = [], extraSheets = {}, sourceSheets = {};
  const range = (r, c, height = 1, width = 1) => ({
    getValues: () => r === 1 ? [headers.slice(c - 1, c - 1 + width)] :
      cells.slice(r - 2, r - 2 + height).map(row => row.slice(c - 1, c - 1 + width)),
    getValue: () => cells[r - 2][c - 1],
    getFormula: () => formulas[r - 2][c - 1],
    getFormulas: () => formulas.slice(r - 2, r - 2 + height).map(row => row.slice(c - 1, c - 1 + width)),
    getNumberFormats: () => Array.from({length: height}, () => Array(width).fill('@')),
    setValues: values => {
      if (failWrite) { failWrite = false; throw Error('simulated write failure'); }
      writes.push({r, c, values: plain(values)});
      values.forEach((row, ri) => row.forEach((v, ci) => { cells[r - 2 + ri][c - 1 + ci] = v; }));
    }
  });
  const sheet = {getName: () => 'Master', getLastColumn: () => headers.length, getLastRow: () => cells.length + 1, getRange: range};
  const sourceSheet = {getName: () => 'Class Roster', getLastColumn: () => sourceHeaders.length,
    getLastRow: () => sourceRows.length + 1, getRange: r => ({getValues: () => r === 1 ? [sourceHeaders] : sourceRows})};
  const sourceBook = {getSheetByName: name => name === 'Class Roster' ? sourceSheet : sourceSheets[name] || null};
  const ss = {getId: () => '1q0y5VI2e1RYhiEriAu2msTGuOsyCwmun_qHmrjDUBJA', getSpreadsheetTimeZone: () => 'America/Los_Angeles',
    getSheetByName: name => name === 'Master' ? sheet : name === 'Info' ?
      {getRange: () => ({getValue: () => 'https://docs.google.com/spreadsheets/d/imports/edit'})} : extraSheets[name] || null};
  const ctx = vm.createContext({console: {log: x => logs.push(x), error: x => logs.push(x)}, Utilities: {getUuid: () => 'review-123', formatDate: v => String(v)},
    CacheService: {getUserCache: () => ({put: (k, v) => cache.set(k, v), get: k => cache.get(k), remove: k => cache.delete(k)})},
    LockService: {getDocumentLock: () => ({tryLock: () => true, releaseLock: () => {}})},
    SpreadsheetApp: {getActiveSpreadsheet: () => ss, openById: () => sourceBook, flush: () => {},
      getUi: () => ({showModalDialog: () => {}})},
    HtmlService: {createHtmlOutput: html => { dialogs.push(html); return {setWidth(){return this;},setHeight(){return this;}}; }}
  });
  load(ctx, 'Config.gs'); load(ctx, 'MasterSave.gs'); load(ctx, 'OnboardingReview.gs');
  const review = vm.runInContext('ONBOARDING_REVIEW', ctx);
  const master = vm.runInContext('MASTER_SAVE', ctx).readMaster(ss);
  const plan = () => review.plan(master, cells, formulas, review.sourceFrom(sourceBook));
  const open = () => review.section(plan(), false);
  return {ctx, review, master, cells, formulas, writes, cache, logs, dialogs, sheet, ss, sourceBook, extraSheets, sourceSheets,
    plan, open, source: sourceRows, setSource: rows => {sourceRows = rows;}, setHeaders: h => {sourceHeaders = h;},
    failNextWrite: () => {failWrite = true;}};
}

test('matches reordered exports by student ID and proposes all three fields, including explicit clears', () => {
  const f = fixture(), p = f.plan();
  assert.equal(p.changes.length, 3);
  assert.deepEqual(plain(p.changes.map(c => [c.studentNumber, c.header, c.after])), [
    ['1', 'Welcome Email', {value: 'Follow-up'}], ['1', 'SCC Completion', {value: 'Completed'}], ['1', 'ISMA', {value: ''}]
  ]);
  assert.match(p.changes[0].label, /Test One/);
  assert.equal(f.writes.length, 0);
});

test('rejecting keeps every value; accepting selected changes writes only those cells and consumes the selection', () => {
  const f = fixture(); f.open();
  f.review.decide('review-123', ['1|Welcome Email'], 'reject');
  assert.equal(f.writes.length, 0);
  assert.throws(() => f.review.decide('review-123', ['1|Welcome Email'], 'accept'), /no longer pending/);
  const r = f.review.decide('review-123', ['1|SCC Completion', '1|ISMA'], 'accept');
  assert.equal(r.count, 2);
  assert.deepEqual(f.cells[0].slice(3), ['Sent', 'Completed', '']);
  assert.equal(f.cells[1][4], 'Completed');
  assert.equal(f.cache.size, 0);
});

test('accept all uses the shared Master save path and rolls back a failed write', () => {
  const f = fixture(); f.open(); f.failNextWrite();
  const ids = plain(f.plan().changes.map(c => c.id));
  assert.throws(() => f.review.decide('review-123', ids, 'accept'), /Save failed/);
  assert.deepEqual(f.cells[0].slice(3), ['Sent', 'Attempt 1', 'Sent']);
  assert.equal(f.cache.size, 1);
  f.review.decide('review-123', ids, 'accept');
  assert.deepEqual(f.cells[0].slice(3), ['Follow-up', 'Completed', '']);
});

for (const mutation of ['local', 'source', 'formula', 'removed', 'duplicate', 'header']) {
  test('a stale ' + mutation + ' cancels the entire selected batch before any write', () => {
    const f = fixture(); f.open();
    const ids = plain(f.plan().changes.map(c => c.id));
    if (mutation === 'local') f.cells[0][5] = 'Locally updated';
    if (mutation === 'source') f.source[1][5] = 'Source updated';
    if (mutation === 'formula') f.formulas[0][5] = '=A1';
    if (mutation === 'removed') f.setSource([f.source[0]]);
    if (mutation === 'duplicate') f.setSource([...f.source, f.source[1]]);
    if (mutation === 'header') f.setHeaders(headers.map(h => h === 'ISMA' ? 'Other' : h));
    assert.throws(() => f.review.decide('review-123', ids, 'accept'), /Nothing accepted/);
    assert.equal(f.writes.length, 0);
    assert.equal(f.cells[0][3], 'Sent');
  });
}

test('acceptance resolves current Master row locations after rows move', () => {
  const f = fixture(); f.open(); f.cells.reverse();
  f.review.decide('review-123', ['1|ISMA'], 'accept');
  assert.equal(f.cells[0][5], 'Lead Approved');
  assert.equal(f.cells[1][5], '');
});

test('expired tokens, forged IDs, unrelated fields, and foreign workbook tokens are rejected', () => {
  const f = fixture();
  assert.throws(() => f.review.decide('gone', ['1|ISMA'], 'accept'), /expired/);
  f.open();
  assert.throws(() => f.review.decide('review-123', ['2|ISMA'], 'accept'), /no longer pending/);
  assert.throws(() => f.review.decide('review-123', ['1|SCC Notes'], 'accept'), /no longer pending/);
  const key = 'onboarding:review-123', data = JSON.parse(f.cache.get(key));
  data.workbook = 'other'; f.cache.set(key, JSON.stringify(data));
  assert.throws(() => f.review.decide('review-123', ['1|ISMA'], 'accept'), /another workbook/);
  assert.equal(f.writes.length, 0);
});

test('duplicates and formula destinations are skipped, absent students are named, missing headers are reported', () => {
  const f = fixture(); f.setSource([f.source[1], f.source[1]]);
  let p = f.plan();
  assert.equal(p.changes.length, 0);
  assert.match(p.issues.join(' '), /duplicate/);
  assert.match(p.missing[0], /2 — Test Two/);
  f.setSource([f.source[1]]); f.formulas[0][3] = '=A1';
  p = f.plan(); assert.equal(p.changes.length, 2); assert.match(p.issues.join(' '), /formula/);
  f.setHeaders(headers.map(h => h === 'ISMA' ? 'Other' : h));
  p = f.plan(); assert.match(p.issues.join(' '), /missing ISMA/);
  assert.equal(p.changes.length, 1);
});

test('preview has no review token or action buttons; dialog exposes bulk controls and valid client JavaScript', () => {
  const f = fixture(); f.cells[0][3] = '<script>alert(1)</script>';
  const preview = f.review.section(f.plan(), true);
  assert.equal(f.cache.size, 0);
  assert.doesNotMatch(preview, /Accept all|Reject all|reviewOnboardingChanges/);
  const html = f.open();
  assert.match(html, /Accept selected/); assert.match(html, /Reject all/); assert.match(html, /\(blank\)/);
  assert.match(html, /&lt;script&gt;/);
  const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
  assert.equal(scripts.length, 1);
  assert.doesNotThrow(() => new vm.Script(scripts[0][1]));
});

test('missing-student details survive unreadable AddDrop history', () => {
  const f = fixture(); load(f.ctx, 'RemoveStudents.gs');
  const remover = vm.runInContext('REMOVE_STUDENTS', f.ctx);
  const badSource = {getSheetByName: () => ({getLastColumn: () => 1, getRange: () => ({getValues: () => [['Wrong header']]})})};
  const details = remover.details(f.ss, ['1'], badSource);
  assert.equal(details.length, 1); assert.equal(details[0].name, 'Test One');
  assert.match(details[0].note, /AddDrop details unavailable/);
});

for (const preview of [true, false]) {
  test('full refresh ' + (preview ? 'preview' : 'save') + ' keeps onboarding proposals separate and names missing students', () => {
    const f = fixture(), packets = [], toasts = [];
    const tableSheet = (name, heads, rows) => ({getName: () => name, getLastColumn: () => heads.length,
      getLastRow: () => rows.length + 1, getRange: r => ({getValues: () => r === 1 ? [heads] : rows})});
    f.sourceSheets.Upstream = tableSheet('Upstream', ['STUDENT_NUMBER', 'CURRENT_ACTIVE_STUDENT'], [[1, 'Y']]);
    f.sourceSheets.ProRoster = tableSheet('ProRoster', ['Student Number', 'Name', 'Gender', 'Phone', 'Contact Name', 'Contact Phone', 'Contact Email'], []);
    f.extraSheets.Contacts = tableSheet('Contacts', ['Student Number', 'Student Source Name', 'Gender', 'Student Phone',
      'Contact Name', 'Relationship', 'Source Phone', 'Source Email', 'Effective Phone', 'Effective Email', 'Learning Coach', 'Contact Key'], []);
    f.ss.toast = (...args) => toasts.push(args);
    f.ctx.ADD_STUDENTS = {addAll: () => ({added: [], wouldAdd: [], skipped: [], notes: [], classRosterIds: new Set()})};
    f.ctx.GRADES_IMPORT = {sourceFrom: () => ({byId: new Map([['1', []]])}), findSlots: () => null,
      masterFields: () => null, run: () => ({issues: [], changes: [], added: [], studentsUpdated: 0})};
    f.ctx.STAR_IMPORT = {sourceFrom: () => ({issues: []}), boy: () => null, latest: () => null};
    f.ctx.REMOVE_STUDENTS = {details: () => {throw Error('history unavailable');}, section: () => ''};
    f.ctx.masterForTest = f.master;
    f.ctx.capturePacket = p => packets.push(p);
    vm.runInContext('saveToMaster = planner => {const p = planner(masterForTest); capturePacket(p); return {};}', f.ctx);
    load(f.ctx, 'RefreshStudentData.gs');
    const result = vm.runInContext('STUDENT_REFRESH', f.ctx).run(preview);
    assert.equal(result.onboarding.changes.length, 3);
    assert.equal(packets.length, 1);
    assert.equal(packets[0].fields.length, 0);
    assert.equal(packets[0].dryRun, preview);
    assert.equal(f.writes.length, 0);
    assert.equal(f.cache.size, preview ? 0 : 1);
    assert.equal(toasts.length, 1);
    assert.match(result.missingSources.find(s => s.source === 'Upstream').students[0], /2 — Test Two/);
    assert.equal(result.missingSources.find(s => s.source === 'ProRoster').students.length, 2);
    assert.match(result.missingSources.find(s => s.source === 'CourseGrades').students[0], /2 — Test Two/);
    assert.match(f.dialogs[0], /Missing students by source/);
    assert.match(f.dialogs[0], /2 — Test Two/);
    assert.equal(f.dialogs[0].includes('Accept all'), !preview);
  });
}
