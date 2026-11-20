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
  // v1.0 left a red fill on ECC Contact History; v1.1 clears only that fill.
  const backgrounds = cells.map(() => ['#ffffff']);
  backgrounds[1][0] = '#ea9999';
  const condition = formula => ({getCriteriaType: () => 'CUSTOM_FORMULA', getCriteriaValues: () => [formula]});
  // The teacher's every-other-row grey rule over the whole table.
  let rules = [{getBooleanCondition: () => condition('=ISEVEN(ROW())'), formula: '=ISEVEN(ROW())', background: '#EEEEEE'}];
  const writes = [], cache = new Map(), logs = [], dialogs = [];
  let uuid = 0, writeCalls = 0, failAt = 0;
  const range = (r, c, height = 1, width = 1) => ({
    getValues: () => r === 1 ? [headers.slice(c - 1, c - 1 + width)] :
      cells.slice(r - 2, r - 2 + height).map(row => row.slice(c - 1, c - 1 + width)),
    getNumRows: () => height,
    getFormulas: () => formulas.slice(r - 2, r - 2 + height).map(row => row.slice(c - 1, c - 1 + width)),
    getBackgrounds: () => backgrounds.slice(r - 2, r - 2 + height).map(row => [...row]),
    setBackgrounds: values => {
      fills.push({r, c, values: plain(values)});
      values.forEach((row, i) => { backgrounds[r - 2 + i][0] = row[0] || '#ffffff'; });
    },
    getNumberFormats: () => Array.from({length: height}, () => Array(width).fill('@')),
    setValue: value => { if (r === 1) headers[c - 1] = value; else cells[r - 2][c - 1] = value; },
    setValues: values => {
      if (++writeCalls === failAt) throw Error('Simulated write failure');
      writes.push({r, c, values: plain(values)});
      values.forEach((row, ri) => row.forEach((value, ci) => { cells[r - 2 + ri][c - 1 + ci] = value; }));
    }
  });
  const sheet = {getName: () => 'Master', getLastColumn: () => headers.length,
    getLastRow: () => cells.length + 1, getRange: range,
    getMaxColumns: () => headers.length, getMaxRows: () => 100,
    insertColumnAfter: column => {
      headers.splice(column, 0, '');
      cells.forEach(row => row.splice(column, 0, ''));
      formulas.forEach(row => row.splice(column, 0, ''));
    },
    getConditionalFormatRules: () => [...rules], setConditionalFormatRules: list => { rules = [...list]; }};
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
      BooleanCriteria: {CUSTOM_FORMULA: 'CUSTOM_FORMULA'},
      newConditionalFormatRule: () => {
        const rule = {};
        const builder = {
          whenFormulaSatisfied(formula) { rule.formula = formula; return builder; },
          setBackground(color) { rule.background = color; return builder; },
          setRanges(ranges) { rule.ranges = ranges; return builder; },
          build: () => Object.assign(rule, {getBooleanCondition: () => condition(rule.formula)})
        };
        return builder;
      },
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
  const status = row => cells[row][headers.indexOf('Attempt Streak')];
  // The fill a teacher sees on a Master row: the first matching rule wins.
  const shown = row => {
    const values = Object.fromEntries(headers.map((h, i) => [h, cells[row][i]]));
    for (const rule of rules) {
      if (rule.formula === '=ISEVEN(ROW())') { if ((row + 2) % 2 === 0) return rule.background; continue; }
      const m = rule.formula.match(/LEFT\(\$([A-Z]+)2,(\d+)\)="([^"]*)"/);
      const column = headers[m[1].charCodeAt(0) - 65];
      if (String(values[column]).slice(0, Number(m[2])) === m[3]) return rule.background;
    }
    return '#ffffff';
  };
  return {fills, backgrounds, status, shown, get rules() { return rules; }, streak, ctx, headers, cells, formulas, writes, cache, logs, dialogs, settingsGrid, ss, sheet, bulk, ecc, set, get, decode, all,
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
  // Only ECC Contact History and the script's Attempt Streak column change.
  f.cells.forEach((row, i) => assert.deepEqual(plain(row.slice(0, HEADERS.length).filter((_, c) => HEADERS[c] !== 'ECC Contact History')),
    before[i].filter((_, c) => HEADERS[c] !== 'ECC Contact History')));
  assert.equal(f.status(0), 'Logged 10/23/2026');
  assert.equal(f.shown(0), '#FFE599', 'yellow shows over the grey row rule');
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
  assert.equal(f.status(0), 'Logged 10/23/2026');
  assert.throws(() => f.ctx.archiveFiveAttemptLogAndOpenPowerSchool(), /already has a 5\+ Attempt Log \(10\/23\/2026\)/);
});

test('Refresh adds Attempt Streak once, fills it, and its rules show over an every-other-row grey rule', () => {
  const f = fixture();
  assert.deepEqual(plain(f.streak.recolor(true)), {red: 1, yellow: 0, added: true});
  assert.equal(f.headers.at(-1), 'Attempt Streak');
  assert.deepEqual([0, 1, 2].map(f.status), ['Due: 5 attempts', '4 attempts', '']);
  // Row 2 (student 11) is an even row, so the grey rule also matches it.
  assert.deepEqual([0, 1, 2].map(f.shown), ['#EA9999', '#ffffff', '#EEEEEE']);
  assert.equal(f.rules.length, 3);
  assert.equal(f.rules.at(-1).formula, '=ISEVEN(ROW())', 'the grey rule is kept, after ours');
  assert.deepEqual(plain(f.rules[0].ranges.map(r => r.getNumRows())), [99, 99], 'whole columns, so new students are covered');
  assert.equal(f.backgrounds[1][0], '#ffffff', 'the old v1.0 red fill is cleared');
  // Running again replaces our rules instead of adding more, and adds no column.
  const width = f.headers.length;
  assert.equal(f.streak.recolor(true).added, false);
  assert.equal(f.headers.length, width); assert.equal(f.rules.length, 3);
  // Values move with the row when Master is sorted; the rules follow them.
  f.cells.reverse(); f.formulas.reverse();
  assert.deepEqual([0, 1, 2].map(f.shown), ['#EEEEEE', '#ffffff', '#EA9999']);
});

test('Attempt Streak with a formula stops the refresh with an explanation', () => {
  const f = fixture();
  f.streak.recolor(true);
  f.formulas[0][f.headers.indexOf('Attempt Streak')] = '=1';
  assert.throws(() => f.streak.recolor(true), /must not contain formulas/);
});

test('recipients: codes anywhere, words only on label and greeting lines', () => {
  const {streak} = fixture();
  const names = {lc: 'Tameika Moore', student: 'RJ Hager'};
  const who = note => plain(streak.recipients(note, names));
  assert.deepEqual(who('LC Voicemail: call back to plan.\nParentSquare: Hello Learning Coach,'), ['LC']);
  assert.deepEqual(who('Email to student RJ Hager:\nHi RJ,\nbody\nEmail to LC Tameika Moore (Mother):\nHi Tameika,'), ['student', 'LC']);
  assert.deepEqual(who('ParentSquare:\nHi Tameika, quick grade update.'), ['LC']);
  assert.deepEqual(who('EC voicemail for Danielle Hayes: call back'), ['EC']);
  assert.deepEqual(who('Called LG, no answer'), ['LG']);
  assert.deepEqual(who('Email to legal guardian:\nHello,'), ['LG']);
  assert.deepEqual(who('Voicemail for Danielle Hayes [Relative]: call back'), [], 'a relationship alone is not EC or LG');
  assert.deepEqual(who('Hello Learning Coach,\nPlease check with your student and ask your learning coach.'), ['LC'],
    'body text does not add the student');
  assert.deepEqual(who('Did not join homeroom'), []);
  assert.deepEqual(who('ECC check missed'), [], 'ECC is not EC');
});

test('5+ log names who each attempt went to; each attempt PowerSchool note starts with To:', () => {
  const f = fixture();
  f.set(0, 'ECC Contact History', [
    '-- 10/22/2026 [Attempt]', 'EC Voicemail for Danielle Hayes [Relative]: call back', '',
    '-- 10/21/2026 [Attempt]', 'ParentSquare:', 'Hello Learning Coach,', '',
    '-- 10/20/2026 [Attempt]', 'Email to student:', 'Hi Student,', '',
    '-- 10/19/2026 [Attempt]', 'Called LG, no answer', '',
    '-- 10/16/2026 [Attempt]', 'Did not join homeroom', '',
    '-- 10/1/2026 [Conversation]', 'Fine'].join('\n'));
  f.set(2, 'New Attempt Student Contact', 'LC Voicemail: please call back');
  const review = f.bulk.review();
  const five = review.entries.find(e => e.type === FIVE);
  assert.match(five.note, /10\/16 absent from live class; 10\/19 phone call\/voicemail to LG; 10\/20 email to student; 10\/21 ParentSquare message to LC; 10\/22 phone call\/voicemail to EC\./);
  assert.match(five.note, /\nContacted: student, LC, LG and EC\.\nNo response received\.$/);
  const result = f.bulk.archive(review.token, f.all(review));
  const attempt = result.links.map(f.decode).find(p => p.studentNumber === '33');
  assert.equal(attempt.note, 'Attempt opening | To: LC | LC Voicemail: please call back');
  assert.doesNotMatch(f.get(2, 'ECC Contact History'), /To: LC/, 'history keeps the note as written');
});
