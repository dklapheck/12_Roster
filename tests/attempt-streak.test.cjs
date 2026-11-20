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
const STREAK = ['-- 10/22/2026 [Attempt]', 'Voicemail left for mom', '',
  '-- 10/20/2026 [Attempt]', 'ParentSquare:', 'Missed homeroom check-in, please reschedule', '',
  '-- 10/16/2026 [Attempt]', 'Hi Sam, you missed check-in today', '',
  '-- 10/15/2026 [Attempt]', 'Did not join homeroom', '',
  '-- 10/14/2026 [Attempt]', 'Texted student', '',
  '-- 10/1/2026 [Conversation]', 'Doing fine'].join('\n');
function fixture(headers = [...HEADERS]) {
  const old = new Clock('2026-10-21T12:00:00Z'), today = new Clock('2026-10-23T12:00:00Z');
  const cells = [
    ['11', 'Student One', '', '', '', '', '', '', '', '', STREAK, '', '', ''],
    ['22', 'Four Only', '', '', '', '', '', '', '', '', STREAK.split('\n').slice(3).join('\n'), '', '', ''],
    ['33', 'Student Three', '', '', '', '', '', '', '', '', '', today, 'Email reminder', '']
  ];
  const fills = [];
  const formulas = cells.map(row => row.map(() => ''));
  const writes = [], cache = new Map(), logs = [], dialogs = [];
  let uuid = 0, writeCalls = 0, failAt = 0;
  const range = (r, c, height = 1, width = 1) => ({
    getValues: () => r === 1 ? [headers.slice(c - 1, c - 1 + width)] :
      cells.slice(r - 2, r - 2 + height).map(row => row.slice(c - 1, c - 1 + width)),
    getFormulas: () => formulas.slice(r - 2, r - 2 + height).map(row => row.slice(c - 1, c - 1 + width)),
    setBackgrounds: values => { fills.push({r, c, values: plain(values)}); },
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
    ['Parent Contact', 'Parent opening', 'contact', 'parent', 'log_date', ''],
    ['5+ Attempts', '', 'contact', 'five', 'log_date', '']
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
  load(ctx, 'EccPowerSchool.gs'); load(ctx, 'LogECC.gs'); load(ctx, 'AttemptStreak.gs');
  const bulk = vm.runInContext('ECC_BULK', ctx), ecc = vm.runInContext('ECC_LOG', ctx);
  const set = (row, header, value) => { cells[row][HEADERS.indexOf(header)] = value; };
  const get = (row, header) => cells[row][HEADERS.indexOf(header)];
  const decode = link => JSON.parse(Buffer.from(link.url.split(/#(?:ecc|parent)=/)[1], 'base64url'));
  const all = review => plain(review.entries.map(entry => entry.id));
  const streak = vm.runInContext('ATTEMPT_STREAK', ctx);
  return {fills, streak, ctx, headers, cells, formulas, writes, cache, logs, dialogs, settingsGrid, ss, sheet, bulk, ecc, set, get, decode, all,
    failWrite: n => {failAt = writeCalls + n;}};
}

const FIVE = '5+ Attempt Log';
const entry = (date, type, note) => '-- ' + date + ' [' + type + ']\n' + note;
const join = (...parts) => parts.join('\n\n');

test('colors: none under 5, red at 5+, yellow once logged (even with more attempts), reset by a successful contact', () => {
  const {streak} = fixture();
  const attempts = n => Array.from({length: n}, (_, i) => entry('10/' + (10 + n - i) + '/2026', 'Attempt', 'Email ' + i));
  const base = entry('10/1/2026', 'Conversation', 'Fine');
  assert.equal(streak.status(join(...attempts(4), base)).color, '');
  const red = streak.status(join(...attempts(6), base));
  assert.equal(red.color, 'red'); assert.equal(red.due, true); assert.equal(red.count, 6);
  const logged = join(entry('10/20/2026', FIVE, 'summary'), ...attempts(5), base);
  assert.equal(streak.status(logged).color, 'yellow');
  const more = streak.status(join(...attempts(7).slice(0, 7), logged));
  assert.equal(more.color, 'yellow'); assert.equal(more.due, false);
  assert.equal(more.count, 12, 'the log itself is never counted as an attempt');
  for (const success of ['Conversation', 'Parent Contact']) {
    assert.equal(streak.status(join(entry('10/25/2026', success, 'Reached'), logged)).color, '');
    const again = streak.status(join(...attempts(5).map(e => e.replace('10/', '11/')), entry('10/25/2026', success, 'Reached'), logged));
    assert.equal(again.color, 'red', 'a new streak after ' + success + ' can be logged again');
  }
  assert.equal(streak.status(join(...attempts(5), '-- 9/14/2026: Good call')).lastSuccess, '9/14/2026');
});

test('summary lists every attempt once, oldest first, with one label each', () => {
  const {streak} = fixture();
  assert.equal(streak.summary(streak.status(STREAK)), [
    '5 unsuccessful student/parent contact attempts in a row since the last successful contact on 10/1/2026:',
    '10/14 text message; 10/15 absent from live class; 10/16 email; 10/20 ParentSquare message; 10/22 phone call/voicemail.',
    'No response received.'
  ].join('\n'));
  assert.equal(streak.methodOf('Dropped off the packet at the front office for the family to pick up'),
    'Dropped off the packet at the front office for…');
  assert.equal(streak.entries('-- 10/12/2026 [Attempt] Called mom')[0].note, 'Called mom');
});

test('bulk review offers one checked 5+ Attempt Log only to due students; archive logs it once and colors it yellow', () => {
  const f = fixture(), review = f.bulk.review();
  const five = review.entries.filter(e => e.type === FIVE);
  assert.deepEqual(plain(five.map(e => [e.id, e.date, e.oldDate])), [['11|' + FIVE, '10/23/2026', false]]);
  assert.match(five[0].note, /^-- 10\/23\/2026 \[5\+ Attempt Log\]\n5 unsuccessful/);
  const before = plain(f.cells);
  const result = f.bulk.archive(review.token, ['11|' + FIVE]);
  const history = f.get(0, 'ECC Contact History');
  assert.equal(history, five[0].note + '\n\n' + STREAK);
  const payload = f.decode(result.links[0]);
  assert.equal(result.links[0].url.includes('#parent='), true);
  assert.equal(payload.outcome, '5+ Attempts'); assert.equal(payload.settings.subtypeValue, 'five');
  assert.equal(payload.date, '10/23/2026');
  assert.match(payload.note, /^5 unsuccessful .* 10\/1\/2026: \| 10\/14 text message; .* \| No response received\.$/);
  f.cells.forEach((row, i) => assert.deepEqual(plain(row.filter((_, c) => HEADERS[c] !== 'ECC Contact History')),
    before[i].filter((_, c) => HEADERS[c] !== 'ECC Contact History')));
  assert.deepEqual(f.fills.at(-1).values, [['#FFE599'], [null], [null]]);
  assert.equal(f.bulk.review().entries.some(e => e.type === FIVE), false, 'never offered twice in a streak');
  // Lost response: the same dialog gets the same links back without saving again.
  const writes = f.writes.length;
  assert.deepEqual(plain(f.bulk.archive(review.token, ['11|' + FIVE])), plain(result));
  assert.equal(f.writes.length, writes);
});

test('an attempt archived in the same batch is included; a conversation in the same batch blocks the log', () => {
  const f = fixture();
  f.set(0, 'Contact Date', new Clock('2026-10-23T12:00:00Z')); f.set(0, 'New Attempt Student Contact', 'Emailed mom again');
  let review = f.bulk.review();
  f.bulk.archive(review.token, f.all(review));
  const history = f.get(0, 'ECC Contact History');
  assert.match(history, /^-- 10\/23\/2026 \[5\+ Attempt Log\]\n6 unsuccessful[^\n]*\n[^\n]*10\/23 email\.\n/);
  assert.match(history, /\n\n-- 10\/23\/2026 \[Attempt\]\nEmailed mom again/);
  assert.equal(f.streak.status(history).color, 'yellow');

  const g = fixture();
  g.set(0, 'ECC Date', new Clock('2026-10-23T12:00:00Z')); g.set(0, 'ECC Overall', 'Reached the student');
  review = g.bulk.review();
  assert.throws(() => g.bulk.archive(review.token, f.all(review)), /Student One: uncheck the 5\+ Attempt Log\. No 5\+ Attempt Log is due/);
  assert.equal(g.get(0, 'ECC Contact History'), STREAK, 'nothing is saved');
});

test('individual log: refuses students who are not due, logs a due student, and recolors', () => {
  const f = fixture();
  let row = 2;
  f.ctx.SpreadsheetApp.getActiveRange = () => ({getSheet: () => f.sheet, getRow: () => row, getNumRows: () => 1});
  const originalRange = f.sheet.getRange;
  f.sheet.getRange = (...args) => ({...originalRange(...args), getValue: () => f.cells[row - 2][0]});
  row = 3;
  assert.throws(() => f.ctx.archiveFiveAttemptLogAndOpenPowerSchool(), /only 4 attempts since the last successful contact \(needs 5\)/);
  row = 2;
  f.ctx.archiveFiveAttemptLogAndOpenPowerSchool();
  assert.match(f.get(0, 'ECC Contact History'), /^-- 10\/23\/2026 \[5\+ Attempt Log\]\n5 unsuccessful/);
  assert.equal(f.dialogs.length, 1);
  assert.match(f.dialogs[0].output.html, /#parent=/);
  assert.deepEqual(f.fills.at(-1).values, [['#FFE599'], [null], [null]]);
  assert.throws(() => f.ctx.archiveFiveAttemptLogAndOpenPowerSchool(), /already has a 5\+ Attempt Log \(10\/23\/2026\)/);
});

test('Refresh 5+ attempt colors marks due students red and leaves other rows blank', () => {
  const f = fixture();
  assert.deepEqual(plain(f.streak.recolor(true)), {red: 1, yellow: 0});
  assert.deepEqual(f.fills.at(-1).values, [['#EA9999'], [null], [null]]);
});
