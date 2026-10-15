const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const context = vm.createContext({
  console,
  CONFIG: {
    normalize: value => value == null ? '' : String(value).trim().replace(/\s+/g, ' '),
    blank: value => value == null || String(value).trim() === '',
    studentId: value => value == null ? '' : String(value).trim(),
  },
});

let source = fs.readFileSync('WigSnapshots.gs', 'utf8');
source = source.replace(
  'return Object.freeze({run, recolor});',
  'return Object.freeze({run, recolor, calculate, isTenthGrade});'
);
vm.runInContext(source, context);
const WIG = vm.runInContext('WIG_SNAPSHOTS', context);

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
