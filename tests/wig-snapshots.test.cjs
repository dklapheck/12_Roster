const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const context = vm.createContext({
  console,
  Utilities: {
    formatDate(date, tz, pattern) {
      assert.equal(pattern, 'yyyy-MM-dd');
      const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
      }).formatToParts(date);
      const get = type => parts.find(part => part.type === type).value;
      return get('year') + '-' + get('month') + '-' + get('day');
    },
  },
  CONFIG: {
    normalize: value => value == null ? '' : String(value).trim().replace(/\s+/g, ' '),
    blank: value => value == null || String(value).trim() === '',
    studentId: value => value == null ? '' : String(value).trim(),
  },
});

let source = fs.readFileSync('WigSnapshots.gs', 'utf8');
source = source.replace(
  'return Object.freeze({run, recolor});',
  'return Object.freeze({run, recolor, calculate, isTenthGrade, weekKey, destination});'
);
vm.runInContext(source, context);
const WIG = vm.runInContext('WIG_SNAPSHOTS', context);

test('Monday opens the existing Tuesday week and keeps it through Sunday', () => {
  const cases = [
    ['2026-11-01T20:00:00Z', '2026-10-27'], // Sunday
    ['2026-11-02T20:00:00Z', '2026-11-03'], // Monday
    ['2026-11-03T20:00:00Z', '2026-11-03'], // Tuesday
    ['2026-11-04T20:00:00Z', '2026-11-03'],
    ['2026-11-05T20:00:00Z', '2026-11-03'],
    ['2026-11-06T20:00:00Z', '2026-11-03'],
    ['2026-11-07T20:00:00Z', '2026-11-03'],
    ['2026-11-08T20:00:00Z', '2026-11-03'],
    ['2026-11-09T20:00:00Z', '2026-11-10'], // Next Monday
    ['2024-02-26T20:00:00Z', '2024-02-27'], // Leap week
    ['2024-03-03T20:00:00Z', '2024-02-27'],
    ['2024-03-04T20:00:00Z', '2024-03-05'],
    ['2024-12-30T20:00:00Z', '2024-12-31'], // Year boundary
    ['2025-01-05T20:00:00Z', '2024-12-31'],
    ['2025-01-06T20:00:00Z', '2025-01-07'],
  ];
  for (const [now, expected] of cases) {
    assert.equal(WIG.weekKey(new Date(now), 'America/Los_Angeles'), expected, now);
  }
});

test('week switches at Monday midnight in the spreadsheet timezone, including DST', () => {
  const cases = [
    ['2026-11-02T07:59:59Z', '2026-10-27'],
    ['2026-11-02T08:00:00Z', '2026-11-03'],
    ['2026-03-09T06:59:59Z', '2026-03-03'],
    ['2026-03-09T07:00:00Z', '2026-03-10'],
  ];
  for (const [now, expected] of cases) {
    assert.equal(WIG.weekKey(new Date(now), 'America/Los_Angeles'), expected, now);
  }
  const sameInstant = new Date('2026-11-02T07:30:00Z');
  assert.equal(WIG.weekKey(sameInstant, 'Etc/UTC'), '2026-11-03');
  assert.equal(WIG.weekKey(sameInstant, 'America/Los_Angeles'), '2026-10-27');
});

function destinationFixture(rows) {
  const headers = ['', 'HR Students', 'Chronically Absent', 'List of Students Chronically Absent',
    'Failing Math', 'Failing English', 'Failing CTE Courses', '', '', '', '', 'Missing Students / Data'];
  const sheet = {
    getRange(row, column) {
      const cell = rows[row - 2][column - 1];
      return {
        getValue: () => cell && typeof cell === 'object' ? '' : cell ?? '',
        getFormula: () => cell && typeof cell === 'object' ? cell.formula : '',
      };
    },
  };
  return {
    ss: {
      getSheetByName(name) { assert.equal(name, 'WIG'); return sheet; },
      getSpreadsheetTimeZone: () => 'America/Los_Angeles',
    },
    table: () => ({sheet, headers, columns: new Map(headers.filter(Boolean).map(h => [h, headers.indexOf(h)])), rows}),
  };
}

test('Monday selects the Tuesday row, reports missing/duplicate dates, and preserves occupied snapshots', () => {
  const key = WIG.weekKey(new Date('2026-11-02T20:00:00Z'), 'America/Los_Angeles');
  const fixture = destinationFixture([['10/27/2026', 20], ['11/3/2026'], ['11/10/2026']]);
  context.CONFIG.table = fixture.table;
  assert.equal(WIG.destination(fixture.ss, key).row, 3);
  assert.equal(WIG.destination(fixture.ss, key).occupied, false);
  for (const value of [0, 20, {formula: '=COUNTA(Master!A2:A)'}]) {
    const occupied = destinationFixture([['11/3/2026', value]]);
    context.CONFIG.table = occupied.table;
    assert.equal(WIG.destination(occupied.ss, key).occupied, true);
  }
  const missing = destinationFixture([['10/27/2026']]);
  context.CONFIG.table = missing.table;
  assert.throws(() => WIG.destination(missing.ss, key), /Add the date 2026-11-03/);
  const duplicate = destinationFixture([['11/3/2026'], ['2026-11-03']]);
  context.CONFIG.table = duplicate.table;
  assert.throws(() => WIG.destination(duplicate.ss, key), /duplicate week dates/);
});

test('Monday and Tuesday preview/save skip the same occupied week before opening the imports workbook', () => {
  for (const now of ['2026-11-02T20:00:00Z', '2026-11-03T20:00:00Z']) {
    for (const preview of [true, false]) {
      const fixture = destinationFixture([['10/27/2026', 20], ['11/3/2026', 0]]);
      const logs = [];
      const runContext = vm.createContext({
        console: {log() {}},
        Date: class extends Date { constructor(...args) { super(...(args.length ? args : [now])); } },
        Utilities: context.Utilities,
        SpreadsheetApp: {getActiveSpreadsheet: () => fixture.ss},
        CONFIG: {
          ...context.CONFIG,
          workbook: ss => ss,
          table: fixture.table,
          withLock: action => action(),
          sourceWorkbook() { throw new Error('Occupied week must not read source data.'); },
          log: (...args) => logs.push(args),
        },
      });
      vm.runInContext(source, runContext);
      const report = vm.runInContext('WIG_SNAPSHOTS', runContext).run(preview, false);
      assert.equal(report.week, '2026-11-03');
      assert.equal(report.row, 3);
      assert.equal(report.skipped, true);
      assert.equal(report.saved, false);
      assert.equal(logs.length, preview ? 0 : 1);
    }
  }
});

test('recognizes the grade formats used by Master', () => {
  for (const value of ['10', '10th', '10th Grade', 'Grade 10']) {
    assert.equal(WIG.isTenthGrade(value), true, value);
  }
  for (const value of ['', '9th Grade', '11th Grade', '12']) {
    assert.equal(WIG.isTenthGrade(value), false, value);
  }
});

test('Failing CTE lists only 10th-grade students', () => {
  const students = new Map([
    ['10', {name: 'Ten Student', first: 'Ten', last: 'Student', grade: '10th Grade', label: 'T Student, 10'}],
    ['11', {name: 'Eleven Student', first: 'Eleven', last: 'Student', grade: '11th Grade', label: 'E Student, 11'}],
  ]);
  const upstreamHeaders = ['CURRENT_ACTIVE_STUDENT', 'PAR_CHRONICALLY_ABSENT',
    'MATH_GRADE', 'MATH_PASSING_FLAG', 'ELA_GRADE', 'ELA_PASSING_FLAG'];
  const courseHeaders = ['CANVAS_PASSING_FLAG', 'COURSE_NAME',
    'HS_COURSE_OFFERINGS_CONTENT_CATEGORY', 'HS_COURSE_OFFERINGS_IS_CR_COURSE',
    'CANVAS_POSTED_COURSE_GRADE', 'CANVAS_UNPOSTED_COURSE_GRADE'];
  const columns = headers => new Map(headers.map((header, index) => [header, index]));
  const upstreamRow = ['Yes', 'No', 90, 'Yes', 90, 'Yes'];
  const failingCte = ['No', 'CAR015A Arts, AV, Comm, Principles', 'CTE Courses', 'No', 45, ''];
  const data = {
    u: {columns: columns(upstreamHeaders)},
    c: {columns: columns(courseHeaders)},
    upstream: new Map([['10', upstreamRow], ['11', upstreamRow]]),
    courses: new Map([['10', [failingCte]], ['11', [failingCte]]]),
    tz: 'America/Los_Angeles',
  };

  const report = WIG.calculate(students, data, '2026-10-13', 2);

  assert.equal(report.failingCounts[2], 1);
  assert.match(report.values[5], /T Student, 10/);
  assert.doesNotMatch(report.values[5], /E Student, 11/);
});
