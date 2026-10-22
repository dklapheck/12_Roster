const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '..', 'CallDueColors.gs'), 'utf8');
const serial = date => Date.parse(date + 'T00:00:00Z') / 86400000 + 25569;
const today = serial('2026-10-22');

function workday(start, count) {
  assert.equal(typeof start, 'number', 'WORKDAY needs a numeric date');
  let date = start;
  for (let n = 0; n < count;) {
    date++;
    const day = new Date((date - 25569) * 86400000).getUTCDay();
    if (day !== 0 && day !== 6) n++;
  }
  return date;
}

// Evaluate the formula subset used by these rules, including lazy IF branches.
function evaluate(formula, cells) {
  const tokens = formula.slice(1).match(/"[^"]*"|\$?[A-Z]+\$?\d+|[A-Z]+|\d+|<>|<=|>=|[(),=<>]/g);
  let pos = 0;
  function atom() {
    const token = tokens[pos++];
    if (token[0] === '"') return ['literal', token.slice(1, -1)];
    if (/^\d+$/.test(token)) return ['literal', Number(token)];
    if (/^\$?[A-Z]+\$?\d+$/.test(token)) return ['cell', token.replace(/\$/g, '')];
    if (token === 'TRUE') return ['literal', true];
    assert.equal(tokens[pos++], '(');
    const args = [];
    if (tokens[pos] !== ')') {
      do { args.push(expression()); } while (tokens[pos] === ',' && ++pos);
    }
    assert.equal(tokens[pos++], ')');
    return ['function', token, args];
  }
  function expression() {
    const left = atom();
    if (['=', '<>', '<', '<=', '>', '>='].includes(tokens[pos])) {
      const operator = tokens[pos++];
      return ['comparison', operator, left, atom()];
    }
    return left;
  }
  function run(node) {
    if (node[0] === 'literal') return node[1];
    if (node[0] === 'cell') return cells[node[1]] ?? '';
    if (node[0] === 'comparison') {
      const a = run(node[2]), b = run(node[3]);
      return {'=': () => a === b, '<>': () => a !== b, '<': () => a < b,
        '<=': () => a <= b, '>': () => a > b, '>=': () => a >= b}[node[1]]();
    }
    const name = node[1], args = node[2];
    if (name === 'IF') return run(args[run(args[0]) ? 1 : 2]);
    const values = args.map(run);
    return {AND: () => values.every(Boolean), OR: () => values.some(Boolean),
      NOT: () => !values[0], N: () => typeof values[0] === 'number' ? values[0] : 0,
      ISNUMBER: () => typeof values[0] === 'number' && Number.isFinite(values[0]),
      TODAY: () => today, WORKDAY: () => workday(...values)}[name]();
  }
  const tree = expression();
  assert.equal(pos, tokens.length, 'All formula tokens parsed');
  return run(tree);
}

function load(globals = {}) {
  const context = vm.createContext(globals);
  vm.runInContext(source, context);
  return vm.runInContext('CALL_DUE_COLORS', context);
}
const api = load();
const rules = api.formulas('AF', 'AG', 'B');
function color(column, ecc, due, id = 'student') {
  return rules.find(r => r.column === column && evaluate(r.formula, {AF2: ecc, AG2: due, B2: id}))?.name;
}

test('each date can be overdue while the other column remains green', () => {
  assert.equal(color('AF', serial('2026-09-23'), serial('2026-11-06')), 'overdue');
  assert.equal(color('AG', serial('2026-09-23'), serial('2026-11-06')), 'green');
  assert.equal(color('AF', serial('2026-10-09'), serial('2026-10-21')), 'green');
  assert.equal(color('AG', serial('2026-10-09'), serial('2026-10-21')), 'overdue');
});

for (const [ecc, due, expected] of [
  ['2026-09-23', '2026-10-21', 'overdue'], ['2026-09-24', '2026-10-22', 'pink'],
  ['2026-09-28', '2026-10-26', 'pink'], ['2026-09-29', '2026-10-27', 'orange'],
  ['2026-10-01', '2026-10-29', 'orange'], ['2026-10-02', '2026-10-30', 'yellow'],
  ['2026-10-08', '2026-11-05', 'yellow'], ['2026-10-09', '2026-11-06', 'green']
]) {
  test('ECC Date boundary ' + ecc + ' uses 20 weekdays and becomes ' + expected, () => {
    assert.equal(color('AF', serial(ecc), serial('2027-01-01')), expected);
  });
  test('Call Due boundary ' + due + ' uses its own date and becomes ' + expected, () => {
    assert.equal(color('AG', serial('2027-01-01'), serial(due)), expected);
  });
}

for (const value of ['', 'not a date']) {
  test('missing or invalid ECC Date does not borrow Call Due: ' + JSON.stringify(value), () => {
    assert.equal(color('AF', value, serial('2026-11-06')), 'missing');
    assert.equal(color('AG', value, serial('2026-11-06')), 'green');
  });
  test('missing or invalid Call Due does not borrow ECC Date: ' + JSON.stringify(value), () => {
    assert.equal(color('AF', serial('2026-10-09'), value), 'green');
    assert.equal(color('AG', serial('2026-10-09'), value), 'missing');
  });
}

test('empty student rows receive no missing-date fill', () => {
  assert.equal(color('AF', '', '', ''), undefined);
  assert.equal(color('AG', '', '', ''), undefined);
});

function sheetFixture(ecc = 32, due = 33, student = 2) {
  const condition = formula => ({getBooleanCondition: () => ({
    getCriteriaType: () => 'CUSTOM_FORMULA', getCriteriaValues: () => [formula]})});
  const old = condition('=AND(N("HR10DUE")=0,MAX($AF2,$AG2)<TODAY())');
  const legacy = condition('=AND(ISNUMBER($AD2),$AD2<TODAY())');
  const unrelated = condition('=AND(ISNUMBER(AW2),AW2<60)');
  let installed = [old, legacy, unrelated];
  const sheet = {getName: () => 'Master', getMaxRows: () => 100,
    getRange: (row, column, height, width) => ({row, column, height, width}),
    getConditionalFormatRules: () => installed, setConditionalFormatRules: value => {installed = value;}};
  const globals = {
    CONFIG: {workbook: () => ({getSheetByName: () => sheet}),
      headerColumns: () => ({columns: new Map([['ECC Date', ecc], ['Call Due', due], ['Student Number', student]])}),
      columnLetter: number => {let text = ''; for (; number; number = Math.floor((number - 1) / 26)) text = String.fromCharCode(65 + (number - 1) % 26) + text; return text;}},
    SpreadsheetApp: {BooleanCriteria: {CUSTOM_FORMULA: 'CUSTOM_FORMULA'},
      newConditionalFormatRule: () => {
        const rule = {};
        const builder = {whenFormulaSatisfied: value => {rule.formula = value; return builder;},
          setBackground: value => {rule.background = value; return builder;},
          setRanges: value => {rule.ranges = value; return builder;},
          setFontColor: value => {rule.font = value; return builder;},
          build: () => Object.assign(rule, condition(rule.formula))};
        return builder;
      }}
  };
  return {api: load(globals), unrelated, installed: () => installed};
}

for (const columns of [[32, 33, 2], [10, 17, 4]]) {
  test('refresh replaces shared rules with single-column rules and stays repeatable: ' + columns.join(','), () => {
    const f = sheetFixture(...columns);
    assert.equal(f.api.apply(true)[0].replaced, 2);
    const installed = f.installed();
    assert.equal(installed.length, 13);
    assert.equal(installed[12], f.unrelated);
    for (let i = 0; i < 12; i++) {
      assert.equal(installed[i].ranges.length, 1);
      assert.equal(installed[i].ranges[0].column, columns[i < 6 ? 0 : 1]);
      assert.equal(installed[i].ranges[0].row, 2);
      assert.equal(installed[i].ranges[0].height, 99);
    }
    const expectedColors = ['#7B1113', '#CC0000', '#FF66C4', '#F6B26B', '#FFE599', '#93C47D'];
    assert.deepEqual(Array.from(installed.slice(0, 6), r => r.background), expectedColors);
    assert.deepEqual(Array.from(installed.slice(6, 12), r => r.background), expectedColors);
    assert.equal(f.api.apply(true)[0].replaced, 12);
    assert.equal(f.installed().length, 13);
    assert.equal(f.installed()[12], f.unrelated);
  });
}
