const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const ctx = vm.createContext({console});
for (const file of ['Config.gs', 'GradesImport.gs', 'GradeTracking.gs']) vm.runInContext(fs.readFileSync(__dirname + '/../' + file, 'utf8'), ctx);
const T = vm.runInContext('GRADE_TRACKING', ctx);
const c = (id, course, grade, period = 'Fall') => ({id, name: 'Student ' + id, course, grade, period});
const meta = {date: '2026-10-15', imported: '2026-10-15 04:15:00'};
const snapshot = courses => T.plan([], courses, {date: '2026-10-09', imported: 'old'}, 'now', 'baseline').rows;
test('flags only drops greater than ten points; supports zero and extra credit', () => {
  const old = snapshot([c('1','A',85), c('2','A',85), c('3','A',10.01), c('4','A',110)]);
  const r = T.plan(old,[c('1','A',75),c('2','A',74.99),c('3','A',0),c('4','A',99)],meta,'now','new');
  assert.deepEqual(Array.from(r.rows, r => r[11]), ['Decreased','RED FLAG','RED FLAG','RED FLAG']);
  assert.equal(r.rows[0][10], -10);
});
test('matches identity and course regardless of row order; flags every course row for affected student', () => {
  const old = snapshot([c('1','A',90),c('1','B',80),c('2','A',90)]);
  const r = T.plan(old,[c('2','A',90),c('1','B',82),c('1','A',70)],meta,'now','new');
  const report = T.reportRows(r.rows, old);
  assert.equal(report.filter(r => r[0] === 'RED FLAG').length, 2);
  assert.equal(report.filter(r => r[7] === 'RED FLAG').length, 1);
  assert.equal(report[0][2], '1');
});
test('same-source retries preserve flags without duplicate history, including a report-render retry', () => {
  const old = snapshot([c('1','A',90)]);
  const courses = [c('1','A',70)];
  const first = T.plan(old,courses,meta,'now','new');
  const history = old.concat(Array.from(first.rows));
  const retry = T.plan(history,courses,meta,'later','retry');
  assert.equal(retry.append.length,0);
  assert.equal(retry.rows[0][11],'RED FLAG');
  assert.equal(T.reportRows(retry.rows,history)[0][9],'2026-10-09');
  const next = T.plan(history,courses,{...meta,date:'2026-10-16'},'later','next');
  assert.equal(next.rows[0][11],'Unchanged');
});
test('missing, blank, new and new-term grades never create artificial drops', () => {
  const old = snapshot([c('1','A',90),c('2','A',''),c('3','A',90),c('4','A',90)]);
  const r = T.plan(old,[c('1','A',''),c('2','A',70),c('4','A',50,'Spring'),c('5','A',20)],meta,'now','new');
  assert.equal(r.rows.filter(r => r[11] === 'RED FLAG').length,0);
  assert.equal(r.rows.filter(r => r[11] === 'No longer in source').length,2);
  assert.equal(r.rows.find(r => r[4] === '2')[11],'No previous numeric grade');
});
test('bootstrap matches courses when slots move; numeric percent strings are parsed on 0..100 scale', () => {
  const headers = ['Preferred Name','Student Number','LAST NAME','FIRST NAME','GRADE','START DATE','MTH','MTH%','Course 1','Course 1%'];
  const master = [{id:'1',values:['Student',1,'Example','Test','','']}];
  const source = {byId:new Map([['1',[{name:'MTH208A Geometry',grade:70,period:'Fall'}]]])};
  const base = T.baseline(headers,[['Student',1,'Example','Test','','','','','MTH208A Geometry','85.5%']],master,source);
  assert.equal(base[0].grade,85.5);
  assert.equal(base[0].period,'Fall');
  assert.equal(T.plan(snapshot(base),T.sourceCourses(master,source),meta,'now','new').rows[0][11],'RED FLAG');
});
test('source snapshot includes courses beyond visible slots and excludes students outside Master', () => {
  const source = {byId:new Map([['1',Array.from({length:14},(_,i)=>({name:'Course '+i,grade:80,period:'Fall'}))],['9',[{name:'Other',grade:50}]]])};
  assert.equal(T.sourceCourses([{id:'1',values:['Student']}],source).length,14);
  assert.throws(()=>T.sourceCourses([],source), /No roster/);
});
test('Master underline targets follow student and course identity, not slot position', () => {
  const headers=['Preferred Name','Student Number','LAST NAME','FIRST NAME','GRADE','START DATE','MTH','MTH%','Course 1','Course 1%'];
  const rows=[['Student',1,'','','','','Other',90,'Moved course',70]];
  const history=[['','','','',1,'Student','Moved course','Fall',70,80,-10,'Decreased'],
    ['','','','',1,'Student','Other','Fall',90,90,0,'Unchanged']];
  const targets=T.underlineTargets(headers,rows,history);
  assert.equal(targets.all.length,2);
  assert.deepEqual(JSON.parse(JSON.stringify(targets.decreased)),[{id:'1',row:2,course:9,pct:10}]);
});
test('duplicate records are rejected before writes',()=>{
  assert.throws(()=>T.plan([],[c('1','A',80),c('1','A',70)],meta,'now','new'),/Duplicate/);
});
test('Grades preview is read-only; saved no-op imports capture history; failed writes do not capture new grades', () => {
  const headers=['Preferred Name','Student Number','LAST NAME','FIRST NAME','GRADE','START DATE','MTH','MTH%'];
  let saved=[['Student',1,'Example','Test','','','MTH208A Geometry',80]], fail=false;
  const calls=[];
  const sheet={getLastColumn:()=>8,getLastRow:()=>2,getMaxRows:()=>1000,
    getRange:(r,c,n,w)=>({getValues:()=>r===1?[headers]:saved,
      getFormulas:()=>[Array(8).fill('')],setValues:v=>{if(fail)throw Error('write failed');saved=v;}})};
  const master={columns:new Map(headers.slice(0,6).map((h,i)=>[h,i])),rows:[saved[0].slice(0,6)]};
  const context=vm.createContext({console,CONFIG:{text:v=>String(v??'').trim(),blank:v=>v===''||v==null,
    same:(a,b)=>a===b,normalize:v=>String(v??'').trim(),studentId:v=>String(v??''),
    withLock:fn=>fn(),table:()=>master,prepareValues:(range,v)=>v,columnLetter:()=>''},
    GRADE_TRACKING:{seed:()=>calls.push('seed'),capture:()=>{calls.push('capture');return {studentsFlagged:0};}},
    SpreadsheetApp:{flush:()=>calls.push('flush')}});
  vm.runInContext(fs.readFileSync(__dirname+'/../GradesImport.gs','utf8'),context);
  const importer=vm.runInContext('GRADES_IMPORT',context);
  const ss={getSheetByName:()=>sheet};
  const source={byId:new Map([['1',[{name:'MTH208A Geometry',grade:80}]]]),issues:[],importDate:'today'};
  importer.run(ss,source,true);
  assert.equal(calls.length,0);
  const result=importer.run(ss,source,false);
  assert.equal(result.writes.length,0);
  assert.deepEqual(calls,['seed','flush','capture']);
  calls.length=0; fail=true; source.byId.get('1')[0].grade=60;
  assert.throws(()=>importer.run(ss,source,false),/write failed/);
  assert.deepEqual(calls,['seed']);
});
