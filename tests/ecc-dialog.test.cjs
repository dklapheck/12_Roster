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
  const elements = Object.fromEntries(['status', 'summary', 'save', 'review', 'remaining', 'close', 'results', 'entries', 'search', 'selectAll', 'clearAll']
    .map(id => [id, new Element()]));
  ['search', 'selectAll', 'clearAll'].forEach(id => {elements[id].className = 'reviewControl';});
  const descendants = root => [root, ...root.children.filter(c => c instanceof Element).flatMap(descendants)];
  const document = {getElementById: id => elements[id], createElement: tag => new Element(tag), createTextNode: text => text,
    querySelectorAll: () => Object.values(elements).flatMap(descendants).filter(e => e.className === 'reviewControl')};
  const windows = [], navigations = [], requests = [];
  let blocked = 0, onSuccess, onFailure;
  const window = {open: () => {
    if (blocked > 0) {blocked--;return null;}
    const tab = {closed: false, document: {body: {}}, location: {replace: url => navigations.push(url)},
      close(){this.closed = true;}};
    windows.push(tab); return tab;
  }};
  const api = (success, failure) => ({
    withSuccessHandler: callback => api(callback, failure), withFailureHandler: callback => api(success, callback),
    getEccBulkReview: () => success({token: 'review-1', entries}),
    archiveSelectedEcc: (token, ids) => {requests.push({token, ids: [...ids]});onSuccess = success; onFailure = failure;}
  });
  const ctx = vm.createContext({document, window, google: {script: {run: api(), host: {close: () => {}}}}});
  vm.runInContext(script, ctx);
  return {ctx, elements, windows, requests, navigations,
    boxes: () => descendants(elements.entries).filter(e => e.type === 'checkbox'),
    block: count => {blocked = count;}, succeed: result => onSuccess(result), fail: message => onFailure({message})};
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

test('tabs are reserved during the click, never navigate before save, and blocked tabs can reopen without another archive', () => {
  const f = fixture(entries); f.block(1); f.elements.save.emit('click');
  assert.deepEqual(f.requests, [{token: 'review-1', ids: entries.map(e => e.id)}]);
  assert.equal(f.windows.length, 2); assert.equal(f.navigations.length, 0);
  assert.equal(f.windows.every(tab => tab.opener === null), true);
  assert.equal(f.elements.close.disabled, true);
  f.succeed(result);
  assert.equal(f.navigations.length, 2);
  assert.equal(f.elements.remaining.hidden, false);
  assert.match(f.elements.status.textContent, /3 entries archived. 2 PowerSchool tabs opened/);
  f.elements.remaining.emit('click');
  assert.equal(f.navigations.length, 3); assert.equal(f.elements.remaining.hidden, true);
  assert.equal(f.requests.length, 1);
  assert.equal(f.elements.results.children.length, 3);
  assert.equal(f.elements.results.children[0].children[0].rel, 'noopener noreferrer');
  f.elements.save.emit('click'); assert.equal(f.requests.length, 1);
});

test('archive failure closes placeholders and allows an explicit retry without claiming entries were saved', () => {
  const f = fixture(entries); f.elements.save.emit('click'); f.fail('Notes changed; reopen review.');
  assert.equal(f.windows.every(tab => tab.closed), true);
  assert.equal(f.navigations.length, 0);
  assert.equal(f.elements.save.disabled, false);
  assert.equal(f.elements.status.textContent, 'Notes changed; reopen review.');
  assert.equal(f.elements.close.disabled, false);
});

test('invalid dates can be unchecked and only the remaining checked entries are archived', () => {
  const f = fixture([{...entries[0], error: 'Enter a valid date in ECC Date.'}, entries[1]]);
  f.elements.save.emit('click');
  assert.equal(f.requests.length, 0); assert.equal(f.windows.length, 0);
  f.boxes()[0].checked = false; f.boxes()[0].emit('change'); f.elements.save.emit('click');
  assert.deepEqual(f.requests[0].ids, ['11|Attempt']);
});

test('empty review cannot archive and displays a useful result', () => {
  const f = fixture([]);
  assert.equal(f.elements.save.disabled, true);
  assert.equal(f.elements.status.textContent, 'No ECC conversations or attempts to archive.');
});
