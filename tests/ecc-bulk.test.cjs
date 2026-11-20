const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const crypto = require('node:crypto');
const load = (ctx, file) => vm.runInContext(fs.readFileSync(path.join(__dirname, '..', file), 'utf8'), ctx);
const plain = value => JSON.parse(JSON.stringify(value));
const HEADERS = ['Student Number', 'Preferred Name', 'FIRST NAME', 'LAST NAME', 'ECC Date',
  'ECC Overall', 'ECC Classes', 'ECC Grades', 'ECC Social', 'ECC Follow-up',
  'ECC Contact History', 'Contact Date', 'New Attempt Student Contact', 'Parent Communication'];
class Clock extends Date {
  constructor(...args) { super(...(args.length ? args : ['2026-10-24T00:30:00Z'])); }
}
function formatDate(date, timezone, format) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-US', {
    timeZone: timezone, year: 'numeric', month: 'numeric', day: 'numeric'
  }).formatToParts(date).map(part => [part.type, part.value]));
  if (format === 'yyyyMMdd') return parts.year + parts.month.padStart(2, '0') + parts.day.padStart(2, '0');
  return parts.month + '/' + parts.day + '/' + parts.year;
}
function fixture(headers = [...HEADERS]) {
  const old = new Clock('2026-10-21T12:00:00Z'), today = new Clock('2026-10-23T12:00:00Z');
  const cells = [
    ['11', 'Student One', '', '', old, 'Discussed progress', 'English work', '', '', 'Follow up', '', today, 'Called; no answer', ''],
    ['22', 'Followup Only', '', '', old, '', '', '', '', 'Follow-up by itself', '', '', '', ''],
    ['33', 'Student Three', '', '', '', '', '', '', '', '', '', old, 'Email reminder', '']
  ];
  const formulas = cells.map(row => row.map(() => ''));
  const writes = [], cache = new Map(), logs = [], dialogs = [];
  let uuid = 0, writeCalls = 0, failAt = 0;
  const range = (r, c, height = 1, width = 1) => ({
    getValues: () => r === 1 ? [headers.slice(c - 1, c - 1 + width)] :
      cells.slice(r - 2, r - 2 + height).map(row => row.slice(c - 1, c - 1 + width)),
    getFormulas: () => formulas.slice(r - 2, r - 2 + height).map(row => row.slice(c - 1, c - 1 + width)),
    getNumberFormats: () => Array.from({length: height}, () => Array(width).fill('@')),
    setValue: value => { if (r === 1) headers[c - 1] = value; else cells[r - 2][c - 1] = value; },
    setValues: values => {
      if (++writeCalls === failAt) throw Error('Simulated write failure');
      writes.push({r, c, values: plain(values)});
      values.forEach((row, ri) => row.forEach((value, ci) => { cells[r - 2 + ri][c - 1 + ci] = value; }));
    }
  });
  const sheet = {getName: () => 'Master', getLastColumn: () => headers.length,
    getLastRow: () => cells.length + 1, getRange: range};
  const settingsGrid = [
    ['Workflow', 'Opening Log Text', 'Log Type value', 'Subtype value', 'Log date field', 'Tag label'],
    ['ECC Conversation', 'Conversation opening', 'contact', 'success', 'log_date', 'ECC'],
    ['ECC Attempt', 'Attempt opening', 'contact', 'attempt', 'log_date', 'Attempt'],
    ['Parent Contact', 'Parent opening', 'contact', 'parent', 'log_date', '']
  ];
  const ss = {getId: () => '1q0y5VI2e1RYhiEriAu2msTGuOsyCwmun_qHmrjDUBJA',
    getSpreadsheetTimeZone: () => 'America/Los_Angeles', toast: () => {},
    getSheetByName: name => name === 'Master' ? sheet : name === 'Info' ?
      {getDataRange: () => ({getDisplayValues: () => settingsGrid})} : null};
  sheet.getParent = () => ss;
  const ctx = vm.createContext({Date: Clock,
    console: {log: value => logs.push(value), error: value => logs.push(value)},
    Utilities: {getUuid: () => 'uuid-' + (++uuid), formatDate,
      DigestAlgorithm: {SHA_256: 'sha256'}, Charset: {UTF_8: 'utf8'},
      computeDigest: (_, json) => [...crypto.createHash('sha256').update(json).digest()],
      base64EncodeWebSafe: json => Buffer.from(json).toString('base64url')},
    CacheService: {getUserCache: () => ({get: key => cache.get(key), put: (key, value) => cache.set(key, value)})},
    LockService: {getDocumentLock: () => ({tryLock: () => true, releaseLock: () => {}})},
    SpreadsheetApp: {getActiveSpreadsheet: () => ss, flush: () => {},
      getActiveRange: () => ({getSheet: () => sheet, getRow: () => 2, getNumRows: () => 1}),
      getUi: () => ({showModalDialog: (output, title) => dialogs.push({output, title})})},
    HtmlService: {createHtmlOutputFromFile: file => ({file, setWidth(){return this;}, setHeight(){return this;}}),
      createHtmlOutput: html => ({html, setWidth(){return this;}, setHeight(){return this;}})}
  });
  // Bulk may be loaded before LogECC by Apps Script. Dependencies must be lazy.
  load(ctx, 'EccBulk.gs'); load(ctx, 'Config.gs'); load(ctx, 'MasterSave.gs');
  load(ctx, 'EccPowerSchool.gs'); load(ctx, 'LogECC.gs');
  const bulk = vm.runInContext('ECC_BULK', ctx), ecc = vm.runInContext('ECC_LOG', ctx);
  const set = (row, header, value) => { cells[row][HEADERS.indexOf(header)] = value; };
  const get = (row, header) => cells[row][HEADERS.indexOf(header)];
  const decode = link => JSON.parse(Buffer.from(link.url.split(/#(?:ecc|parent)=/)[1], 'base64url'));
  const all = review => plain(review.entries.map(entry => entry.id));
  return {ctx, headers, cells, formulas, writes, cache, logs, dialogs, settingsGrid, ss, sheet, bulk, ecc, set, get, decode, all,
    failWrite: n => {failAt = writeCalls + n;}};
}

test('review includes both entry types, ignores followup-only rows, flags old dates in the sheet timezone, and writes nothing', () => {
  const f = fixture(), review = f.bulk.review();
  assert.deepEqual(plain(review.entries.map(e => [e.id, e.date, e.oldDate])), [
    ['11|Conversation', '10/21/2026', true], ['11|Attempt', '10/23/2026', false], ['33|Attempt', '10/21/2026', true]
  ]);
  assert.match(review.entries[0].note, /Classes: English work/);
  assert.match(review.entries[0].note, /Actions \/ Follow up: Follow up/);
  assert.equal(f.writes.length, 0);
  assert.equal(review.entries.some(e => e.studentNumber === '22'), false);
  assert.doesNotMatch([...f.cache.values()][0], /Discussed progress|Called; no answer/);
});

test('bulk saves both entries in date order, builds distinct PS handoffs with workflow settings, and preserves all source cells', () => {
  const f = fixture(), before = plain(f.cells), review = f.bulk.review();
  const result = f.bulk.archive(review.token, f.all(review));
  assert.equal(result.students, 2); assert.equal(result.links.length, 3);
  const history = f.get(0, 'ECC Contact History');
  assert.match(history, /^-- 10\/23\/2026 \[Attempt\]/);
  assert.match(history, /-- 10\/21\/2026 \[Conversation\]/);
  const payloads = result.links.map(f.decode);
  assert.deepEqual(plain(payloads.map(p => [p.studentNumber, p.outcome, p.settings.subtypeValue])), [
    ['11', 'ECC Conversation', 'success'], ['11', 'ECC Attempt', 'attempt'], ['33', 'ECC Attempt', 'attempt']
  ]);
  assert.equal(new Set(payloads.map(p => p.requestId)).size, 3);
  assert.equal(payloads[0].v, 1);
  assert.match(payloads[0].note, /^Conversation opening \| Discussed progress \| Classes: English work/);
  assert.doesNotMatch(payloads[0].note, /[\r\n]/);
  assert.match(payloads[1].note, /^Attempt opening \| Called; no answer$/);
  assert.equal(f.writes.every(w => w.c === HEADERS.indexOf('ECC Contact History') + 1), true);
  f.cells.forEach((row, index) => assert.deepEqual(plain(row.filter((_, c) => HEADERS[c] !== 'ECC Contact History')),
    before[index].filter((_, c) => HEADERS[c] !== 'ECC Contact History')));
});

test('unchecking entries saves only the selected type/student, even after Master rows are sorted', () => {
  const f = fixture(), review = f.bulk.review(); f.cells.reverse(); f.formulas.reverse();
  const result = f.bulk.archive(review.token, ['11|Attempt']);
  assert.equal(result.links.length, 1);
  assert.equal(f.get(0, 'ECC Contact History'), '');
  assert.match(f.get(2, 'ECC Contact History'), /\[Attempt\]/);
  assert.doesNotMatch(f.get(2, 'ECC Contact History'), /\[Conversation\]/);
});

test('new review and individual actions archive unchanged notes again; same-dialog response retry only recovers links', () => {
  const f = fixture(), review = f.bulk.review();
  const result = f.bulk.archive(review.token, ['11|Conversation']);
  const firstHistory = f.get(0, 'ECC Contact History'), writes = f.writes.length;
  assert.deepEqual(plain(f.bulk.archive(review.token, ['11|Conversation'])), plain(result));
  assert.equal(f.writes.length, writes);
  const next = f.bulk.review(); f.bulk.archive(next.token, ['11|Conversation']);
  assert.equal(f.get(0, 'ECC Contact History'), firstHistory + '\n\n' + firstHistory);
  // Select the real Master row with a real ID cell for the individual wrapper.
  const originalRange = f.sheet.getRange;
  f.sheet.getRange = (...args) => ({...originalRange(...args), getValue: () => '11'});
  f.ctx.archiveEccConversation();
  assert.equal(f.get(0, 'ECC Contact History').split('[Conversation]').length - 1, 3);
});

for (const mutation of ['notes', 'date', 'formula', 'removed', 'duplicate']) {
  test('stale ' + mutation + ' blocks the whole selected batch without modifying history', () => {
    const f = fixture(), review = f.bulk.review();
    if (mutation === 'notes') f.set(0, 'ECC Overall', 'Changed after review');
    if (mutation === 'date') f.set(0, 'ECC Date', new Clock('2026-10-20T12:00:00Z'));
    if (mutation === 'formula') f.formulas[0][HEADERS.indexOf('ECC Overall')] = '=A1';
    if (mutation === 'removed') f.cells[0][0] = '';
    if (mutation === 'duplicate') f.cells[1][0] = '11';
    assert.throws(() => f.bulk.archive(review.token, f.all(review)), /changed since|duplicate Student Number/);
    assert.equal(f.writes.length, 0);
  });
}

test('invalid dates are visible and require correction or deselection; older dates remain valid', () => {
  const f = fixture(); f.set(0, 'ECC Date', '');
  const review = f.bulk.review();
  assert.match(review.entries[0].error, /valid date in ECC Date/);
  assert.throws(() => f.bulk.archive(review.token, ['11|Conversation']), /valid date in ECC Date/);
  assert.equal(f.writes.length, 0);
  assert.equal(f.bulk.archive(review.token, ['33|Attempt']).links.length, 1);
});

test('combined conversation and attempt size is validated before any write; refreshed history is preserved', () => {
  const f = fixture(); f.set(0, 'ECC Overall', 'x'.repeat(26000)); f.set(0, 'New Attempt Student Contact', 'y'.repeat(26000));
  const review = f.bulk.review();
  assert.throws(() => f.bulk.archive(review.token, ['11|Conversation', '11|Attempt']), /50,000/);
  assert.equal(f.writes.length, 0);
  f.set(0, 'ECC Overall', 'Discussed progress'); f.set(0, 'New Attempt Student Contact', 'Called; no answer');
  const next = f.bulk.review();
  f.set(0, 'ECC Contact History', '-- 10/20/2026 [Attempt]\nOther editor saved this');
  f.bulk.archive(next.token, ['11|Conversation']);
  assert.match(f.get(0, 'ECC Contact History'), /Other editor saved this/);
});

test('failed batch writes roll back earlier rows and leave the review available for retry', () => {
  const f = fixture(), review = f.bulk.review(); f.failWrite(2);
  assert.throws(() => f.bulk.archive(review.token, f.all(review)), /Save failed/);
  assert.equal(f.get(0, 'ECC Contact History'), ''); assert.equal(f.get(2, 'ECC Contact History'), '');
  assert.equal(f.bulk.archive(review.token, f.all(review)).links.length, 3);
});

test('missing workflow settings cancel all archives before saving', () => {
  const f = fixture(), review = f.bulk.review(); f.settingsGrid.splice(2, 1);
  assert.throws(() => f.bulk.archive(review.token, f.all(review)), /Expected one PowerSchool settings row/);
  assert.equal(f.writes.length, 0);
});

test('near-limit archives support recovering the same links without reappending the note', () => {
  const f = fixture(); f.set(0, 'ECC Overall', 'x'.repeat(26000));
  const review = f.bulk.review(), result = f.bulk.archive(review.token, ['11|Conversation']);
  const size = f.get(0, 'ECC Contact History').length;
  assert.deepEqual(plain(f.bulk.archive(review.token, ['11|Conversation'])), plain(result));
  assert.equal(f.get(0, 'ECC Contact History').length, size);
});

test('rejects expired/foreign tokens, unreviewed IDs, empty and duplicate selections, and changed replay selections', () => {
  const f = fixture(), review = f.bulk.review();
  assert.throws(() => f.bulk.archive('expired', ['11|Conversation']), /expired/);
  assert.throws(() => f.bulk.archive(review.token, ['22|Conversation']), /Invalid ECC selection/);
  assert.throws(() => f.bulk.archive(review.token, []), /Select at least one/);
  assert.throws(() => f.bulk.archive(review.token, ['11|Attempt', '11|Attempt']), /Select at least one/);
  const key = 'ecc-bulk:' + review.token, data = JSON.parse(f.cache.get(key));
  data.workbook = 'other'; f.cache.set(key, JSON.stringify(data));
  assert.throws(() => f.bulk.archive(review.token, ['11|Conversation']), /another workbook/);
  data.workbook = f.ss.getId(); f.cache.set(key, JSON.stringify(data));
  f.bulk.archive(review.token, ['11|Conversation']);
  assert.throws(() => f.bulk.archive(review.token, ['11|Attempt']), /already been saved/);
});

test('Update exposes logging and demographics first and repeats both in the ECC submenu', () => {
  const f = fixture(), menus = [];
  f.ctx.SpreadsheetApp.getUi = () => ({showModalDialog: (output, title) => f.dialogs.push({output, title}), createMenu: title => {
    const menu = {title, items: [], addItem(label, handler){this.items.push({label, handler});return this;},
      addSeparator(){this.items.push({separator: true});return this;},
      addSubMenu(submenu){this.items.push(submenu);return this;}, addToUi(){menus.push(this);}};
    return menu;
  }});
  load(f.ctx, 'Menus.gs'); f.ctx.onOpen();
  const update = menus[0];
  assert.deepEqual(plain(update.items.slice(0, 2)), [
    {label: 'Archive and Log All ECC', handler: 'archiveAndLogAllEcc'},
    {label: 'Open Student Demographics', handler: 'openStudentDemographics'}
  ]);
  const submenu = update.items.find(item => item.title === 'ECC & Demographics');
  assert.deepEqual(submenu.items.filter(item => ['archiveAndLogAllEcc', 'openStudentDemographics'].includes(item.handler)).map(item => item.handler),
    ['archiveAndLogAllEcc', 'openStudentDemographics']);
  f.ctx.archiveAndLogAllEcc();
  assert.equal(f.dialogs[0].output.file, 'EccBulkDialog');
  assert.equal(f.writes.length, 0);
});

test('Parent Communication is offered as its own entry, dated by Contact Date, and logged as a Parent Contact handoff', () => {
  const f = fixture();
  f.set(2, 'Parent Communication', 'Spoke with mom about missing work');
  const review = f.bulk.review();
  const parent = review.entries.find(e => e.id === '33|Parent Contact');
  assert.equal(parent.type, 'Parent Contact');
  assert.equal(parent.date, '10/21/2026');
  assert.match(parent.note, /^-- 10\/21\/2026 \[Parent Contact\]\nSpoke with mom/);
  const result = f.bulk.archive(review.token, ['33|Attempt', '33|Parent Contact']);
  const link = result.links.find(l => l.type === 'Parent Contact');
  assert.match(link.url, /#parent=/);
  const payload = f.decode(link);
  assert.equal(payload.outcome, 'Parent Contact');
  assert.equal(payload.settings.subtypeValue, 'parent');
  assert.equal(payload.date, '10/21/2026');
  assert.equal(payload.note, 'Parent opening | Spoke with mom about missing work');
  assert.match(f.decode(result.links.find(l => l.type === 'Attempt')).note, /^Attempt opening/);
  const history = f.get(2, 'ECC Contact History');
  assert.match(history, /\[Parent Contact\]\nSpoke with mom/);
  assert.match(history, /\[Attempt\]\nEmail reminder/);
  assert.equal(f.get(2, 'Parent Communication'), 'Spoke with mom about missing work');
});

test('a Parent Contact needs a valid Contact Date and its own settings row', () => {
  const f = fixture();
  f.set(0, 'Parent Communication', 'Emailed dad');
  f.set(0, 'Contact Date', '');
  let review = f.bulk.review();
  assert.match(review.entries.find(e => e.id === '11|Parent Contact').error, /valid date in Contact Date/);
  f.set(0, 'Contact Date', new Clock('2026-10-23T12:00:00Z'));
  f.settingsGrid.pop();
  review = f.bulk.review();
  assert.throws(() => f.bulk.archive(review.token, ['11|Parent Contact']), /Expected one PowerSchool settings row for Parent Contact/);
  assert.equal(f.writes.length, 0);
});

test('individual Parent Contact action archives to history and opens a parent PowerSchool dialog', () => {
  const f = fixture();
  f.set(0, 'Parent Communication', 'Text with grandma');
  const originalRange = f.sheet.getRange;
  f.sheet.getRange = (...args) => ({...originalRange(...args), getValue: () => '11'});
  const result = f.ctx.archiveParentContactAndOpenPowerSchool();
  assert.equal(result.outcome, 'Parent Contact');
  assert.match(f.get(0, 'ECC Contact History'), /^-- 10\/23\/2026 \[Parent Contact\]\nText with grandma$/);
  assert.equal(f.dialogs.length, 1);
  assert.equal(f.dialogs[0].title, 'Open PowerSchool Parent Contact Log');
  const href = f.dialogs[0].output.html.match(/href="([^"]+)"/)[1];
  const payload = JSON.parse(Buffer.from(href.split('#parent=')[1], 'base64url'));
  assert.equal(payload.outcome, 'Parent Contact');
  assert.equal(payload.note, 'Parent opening | Text with grandma');
  assert.equal(payload.nearLimit, undefined);
});

test('setup renames ECC Attempt Date to Contact Date once, and validation explains the old header', () => {
  const old = HEADERS.map(h => h === 'Contact Date' ? 'ECC Attempt Date' : h);
  const f = fixture(old);
  const before = plain(f.cells);
  assert.throws(() => f.ecc.validate(), /Rename the Master header ECC Attempt Date to Contact Date/);
  assert.equal(f.ecc.renameContactDateHeader(), true);
  assert.deepEqual(f.headers, HEADERS);
  assert.deepEqual(plain(f.cells), before);
  assert.equal(f.ecc.renameContactDateHeader(), false);
  f.ecc.validate();
});

for (const field of ['Parent Communication', 'Contact Date']) {
  test('changed Parent Contact ' + field + ' cancels a mixed batch before any write', () => {
    const f = fixture(); f.set(0, 'Parent Communication', 'Spoke with mom');
    const review = f.bulk.review();
    f.set(0, field, field === 'Contact Date' ? new Clock('2026-10-22T12:00:00Z') : 'New message');
    assert.throws(() => f.bulk.archive(review.token, f.all(review)), /changed since/);
    assert.equal(f.writes.length, 0);
  });
}

test('Parent Contact formula notes cannot be archived', () => {
  const f = fixture(); f.set(0, 'Parent Communication', 'Generated note');
  f.formulas[0][HEADERS.indexOf('Parent Communication')] = '=A1';
  const review = f.bulk.review();
  assert.match(review.entries.find(e => e.id === '11|Parent Contact').error, /Use a value/);
  assert.throws(() => f.bulk.archive(review.token, f.all(review)), /Use a value/);
  assert.equal(f.writes.length, 0);
});

test('Parent Contact uses its own date in existing history and allows blank Type/Subtype settings', () => {
  const f = fixture();
  f.set(0, 'Parent Communication', 'Spoke with dad');
  f.set(0, 'ECC Contact History', '-- 10/24/2026 [Conversation]\nNewer\n\n-- 10/20/2026 [Attempt]\nOlder');
  f.settingsGrid[3][2] = ''; f.settingsGrid[3][3] = '';
  const review = f.bulk.review(), result = f.bulk.archive(review.token, ['11|Parent Contact']);
  assert.equal(f.get(0, 'ECC Contact History'), '-- 10/24/2026 [Conversation]\nNewer\n\n' +
    '-- 10/23/2026 [Parent Contact]\nSpoke with dad\n\n-- 10/20/2026 [Attempt]\nOlder');
  assert.equal(f.decode(result.links[0]).settings.typeValue, '');
  assert.equal(f.decode(result.links[0]).settings.subtypeValue, '');
});

test('Parent Contact and ECC writes roll back together if a later student fails', () => {
  const f = fixture(); f.set(0, 'Parent Communication', 'Called mom');
  const before = plain(f.cells), review = f.bulk.review(); f.failWrite(2);
  assert.throws(() => f.bulk.archive(review.token, f.all(review)), /Save failed/);
  assert.deepEqual(plain(f.cells), before);
  assert.equal(f.bulk.archive(review.token, f.all(review)).links.length, 4);
});

test('Parent Contact appears in the ECC submenu', () => {
  const f = fixture(), menus = [];
  f.ctx.SpreadsheetApp.getUi = () => ({createMenu: title => {
    const menu = {title, items: [], addItem(label, handler){this.items.push({label, handler});return this;},
      addSeparator(){return this;}, addSubMenu(submenu){this.items.push(submenu);return this;}, addToUi(){menus.push(this);}};
    return menu;
  }});
  load(f.ctx, 'Menus.gs'); f.ctx.onOpen();
  const submenu = menus[0].items.find(item => item.title === 'ECC & Demographics');
  assert.ok(submenu.items.some(item => item.handler === 'archiveParentContactAndOpenPowerSchool'));
});
