const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const html = fs.readFileSync(path.join(__dirname, '..', 'EccBulkDialog.html'), 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];

class Element {
  constructor(tag = 'div') { this.tag = tag; this.children = []; this.listeners = {}; this.value = ''; }
  append(...children) { this.children.push(...children); }
  insertCell() { const cell = new Element('td'); this.append(cell); return cell; }
  setAttribute(key, value) { this[key] = value; }
  addEventListener(event, callback) { this.listeners[event] = callback; }
  emit(event) { this.listeners[event]?.(); }
}
function fixture(entries) {
  // Derive IDs from the current HTML so removed controls cannot linger in tests.
  const elements = Object.fromEntries([...html.matchAll(/\bid="([^"]+)"/g)]
    .map(([, id]) => [id, new Element()]));
  ['search', 'selectAll', 'clearAll'].forEach(id => {elements[id].className = 'reviewControl';});
  const descendants = root => [root, ...root.children.filter(c => c instanceof Element).flatMap(descendants)];
  const document = {getElementById: id => elements[id], createElement: tag => new Element(tag), createTextNode: text => text,
    querySelectorAll: () => Object.values(elements).flatMap(descendants).filter(e => e.className === 'reviewControl')};
  const requests = [];
  let onSuccess, onFailure, closedDialog = false;
  const window = {open() { assert.fail('PowerSchool tabs must be opened manually.'); }};
  const api = (success, failure) => ({
    withSuccessHandler: callback => api(callback, failure), withFailureHandler: callback => api(success, callback),
    getEccBulkReview: () => success({token: 'review-1', entries}),
    archiveSelectedEcc: (token, ids) => {requests.push({token, ids: [...ids]});onSuccess = success; onFailure = failure;}
  });
  const ctx = vm.createContext({document, window,
    google: {script: {run: api(), host: {close: () => {closedDialog = true;}}}}});
  vm.runInContext(script, ctx);
  return {elements, requests,
    boxes: () => descendants(elements.entries).filter(e => e.type === 'checkbox'),
    succeed: result => onSuccess(JSON.parse(JSON.stringify(result))), fail: message => onFailure({message}),
    closedDialog: () => closedDialog, start: () => elements.save.emit('click')};
}
const entries = [
  {id: '11|Conversation', label: 'Student One', studentNumber: '11', type: 'Conversation', date: '10/21/2026', oldDate: true, note: '<script>unsafe</script>', error: ''},
  {id: '11|Attempt', label: 'Student One', studentNumber: '11', type: 'Attempt', date: '10/23/2026', oldDate: false, note: 'Called', error: ''},
  {id: '33|Attempt', label: 'Student Three', studentNumber: '33', type: 'Attempt', date: '10/23/2026', oldDate: false, note: 'Email reminder', error: ''}
];
const result = {students: 2, links: entries.map((e, i) => ({...e, url: 'https://californiak12.powerschool.com/teachers/home.html#ecc=payload' + i}))};

test('dialog checks every entry, search does not alter selection, and older dates and note text render safely', () => {
  const f = fixture(entries);
  assert.equal(f.boxes().every(box => box.checked), true);
  assert.match(f.elements.summary.textContent, /2 students selected · 1 conversations · 2 attempts/);
  const first = f.elements.entries.children[0];
  assert.equal(first.children[3].children[0].textContent, 'Older than today');
  assert.equal(first.children[4].children[0].children[1].textContent, '<script>unsafe</script>');
  f.elements.search.value = 'Three'; f.elements.search.emit('input');
  assert.equal(first.hidden, true);
  assert.match(f.elements.summary.textContent, /2 students/);
  f.elements.clearAll.emit('click'); assert.equal(f.elements.save.disabled, true);
  f.elements.selectAll.emit('click'); assert.equal(f.boxes().every(box => box.checked), true);
});

test('archive displays manual links after saving without opening tabs or archiving twice', () => {
  const f = fixture(entries); f.start();
  assert.deepEqual(f.requests, [{token: 'review-1', ids: entries.map(e => e.id)}]);
  assert.equal(f.elements.results.children.length, 0);
  assert.equal(f.elements.close.disabled, true);
  assert.equal(f.elements.save.disabled, true);
  f.start(); assert.equal(f.requests.length, 1);
  f.succeed(result);
  assert.equal(f.elements.review.hidden, true);
  assert.equal(f.elements.links.hidden, false);
  assert.equal(f.elements.save.hidden, true);
  assert.equal(f.elements.results.children.length, 3);
  f.elements.results.children.forEach((li, i) => {
    const a = li.children[0];
    assert.equal(a.tag, 'a');
    assert.equal(a.href, result.links[i].url);
    assert.equal(a.target, '_blank');
    assert.equal(a.rel, 'noopener noreferrer');
    assert.equal(a.textContent, entries[i].label + ' — ' + entries[i].studentNumber + ' — ' + entries[i].type + ' — ' + entries[i].date);
  });
  f.start(); assert.equal(f.requests.length, 1);
});

test('archive failure permits an explicit retry and shows no links or success claim', () => {
  const f = fixture(entries); f.start(); f.fail('Notes changed; reopen review.');
  assert.equal(f.elements.results.children.length, 0);
  assert.equal(f.elements.save.disabled, false);
  assert.equal(f.elements.status.textContent, 'Notes changed; reopen review.');
  assert.equal(f.elements.close.disabled, false);
  assert.equal(f.boxes().every(box => !box.disabled), true);
  f.start(); assert.equal(f.requests.length, 2);
});

test('invalid dates can be unchecked and only the remaining checked entries are archived', () => {
  const f = fixture([{...entries[0], error: 'Enter a valid date in ECC Date.'}, entries[1]]);
  f.start();
  assert.equal(f.requests.length, 0);
  f.boxes()[0].checked = false; f.boxes()[0].emit('change'); f.elements.save.emit('click');
  assert.deepEqual(f.requests[0].ids, ['11|Attempt']);
});

test('dialog instructs users to submit and close each student tab before the next link', () => {
  assert.match(html, /Open one student at a time\. Submit the log, close that PowerSchool tab, then click the next link\./);
});

test('link clicks count openings once without claiming submission or verification', () => {
  const f = fixture(entries); f.start(); f.succeed(result);
  assert.match(f.elements.status.textContent, /3 entries archived\. 0 of 3 links opened\./);
  const li = f.elements.results.children[0], a = li.children[0];
  a.emit('click'); a.emit('click');
  assert.equal(li.className, 'opened');
  assert.equal(li.children.length, 2);
  assert.equal(li.children[1].textContent, '✓ opened');
  assert.match(f.elements.status.textContent, /1 of 3 links opened/);
  assert.doesNotMatch(f.elements.status.textContent, /submitted|confirmed|verified/);
  for (const row of f.elements.results.children.slice(1)) row.children[0].emit('click');
  assert.match(f.elements.status.textContent, /3 of 3 links opened/);
  assert.equal(f.requests.length, 1);
});

test('student labels and near-limit warnings render as text in the link list', () => {
  const f = fixture(entries); f.start();
  f.succeed({links: [{...result.links[0], label: '<script>unsafe</script>', nearLimit: true}]});
  const li = f.elements.results.children[0];
  assert.match(li.children[0].textContent, /^<script>unsafe<\/script>/);
  assert.match(li.children[1], /History is close to the cell size limit/);
});

test('dialog cannot close while archiving but can close after success or failure', () => {
  for (const success of [true, false]) {
    const f = fixture(entries); f.start(); f.elements.close.emit('click');
    assert.equal(f.closedDialog(), false);
    if (success) f.succeed(result); else f.fail('Try again');
    f.elements.close.emit('click'); assert.equal(f.closedDialog(), true);
  }
});

test('empty review cannot archive and displays a useful result', () => {
  const f = fixture([]);
  assert.equal(f.elements.save.disabled, true);
  assert.equal(f.elements.status.textContent, 'No ECC conversations or attempts to archive.');
});
