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
  // Use only IDs that exist in the shipped dialog, so removed controls cannot
  // silently survive in the test fixture.
  const elements = Object.fromEntries([...html.matchAll(/<[^>]+\bid="([^"]+)"[^>]*>/g)].map(([tag, id]) => {
    const el = new Element();
    el.hidden = /\shidden\b/.test(tag);
    el.disabled = /\sdisabled\b/.test(tag);
    if (/class="reviewControl"/.test(tag)) el.className = 'reviewControl';
    return [id, el];
  }));
  const descendants = root => [root, ...root.children.filter(c => c instanceof Element).flatMap(descendants)];
  const document = {getElementById: id => elements[id], createElement: tag => new Element(tag), createTextNode: text => text,
    querySelectorAll: () => Object.values(elements).flatMap(descendants).filter(e => e.className === 'reviewControl')};
  const windows = [], requests = [];
  let onSuccess, onFailure, closedDialog = false;
  const window = {open: (...args) => {windows.push(args);}};
  const api = (success, failure) => ({
    withSuccessHandler: callback => api(callback, failure), withFailureHandler: callback => api(success, callback),
    getEccBulkReview: () => success({token: 'review-1', entries}),
    archiveSelectedEcc: (token, ids) => {requests.push({token, ids: [...ids]});onSuccess = success; onFailure = failure;}
  });
  const ctx = vm.createContext({document, window,
    google: {script: {run: api(), host: {close: () => {closedDialog = true;}}}}});
  vm.runInContext(script, ctx);
  return {elements, windows, requests,
    boxes: () => descendants(elements.entries).filter(e => e.type === 'checkbox'),
    links: () => elements.results.children.map(li => li.children[0]),
    succeed: result => onSuccess(structuredClone(result)), fail: message => onFailure({message}),
    closedDialog: () => closedDialog, start: () => elements.save.emit('click')};
}
const entries = [
  {id: '11|Conversation', label: 'Student One', studentNumber: '11', type: 'Conversation', date: '10/21/2026', oldDate: true, note: '<script>unsafe</script>', error: ''},
  {id: '11|Attempt', label: 'Student One', studentNumber: '11', type: 'Attempt', date: '10/23/2026', oldDate: false, note: 'Called', error: ''},
  {id: '33|Attempt', label: 'Student Three', studentNumber: '33', type: 'Attempt', date: '10/23/2026', oldDate: false, note: 'Email reminder', error: ''},
  {id: '33|Parent Contact', label: 'Student Three', studentNumber: '33', type: 'Parent Contact', date: '10/23/2026', oldDate: false, note: 'Spoke with parent', error: ''}
];
const result = {students: 2, links: entries.map((e, i) => ({...e,
  url: 'https://californiak12.powerschool.com/teachers/home.html#' + (e.type === 'Parent Contact' ? 'parent=' : 'ecc=') + 'payload' + i}))};

test('dialog checks all three entry types, search preserves selection, and note text renders safely', () => {
  const f = fixture(entries);
  assert.equal(f.boxes().every(box => box.checked), true);
  assert.equal(f.elements.summary.textContent, '2 students selected · 1 conversations · 2 attempts · 1 parent contacts');
  const first = f.elements.entries.children[0];
  assert.equal(first.children[3].children[0].textContent, 'Older than today');
  assert.equal(first.children[4].children[0].children[1].textContent, '<script>unsafe</script>');
  assert.equal(f.elements.entries.children[3].children[2].textContent, 'Parent contact (successful)');
  f.elements.search.value = 'Three'; f.elements.search.emit('input');
  assert.equal(first.hidden, true);
  assert.match(f.elements.summary.textContent, /2 students/);
  f.elements.clearAll.emit('click'); assert.equal(f.elements.save.disabled, true);
  f.elements.selectAll.emit('click'); assert.equal(f.boxes().every(box => box.checked), true);
});

test('archive shows one link per selected entry without opening PowerSchool automatically', () => {
  const f = fixture(entries); f.start();
  assert.deepEqual(f.requests, [{token: 'review-1', ids: entries.map(e => e.id)}]);
  assert.equal(f.elements.close.disabled, true);
  f.succeed(result);
  assert.equal(f.elements.links.hidden, false);
  assert.equal(f.elements.review.hidden, true);
  assert.equal(f.elements.save.hidden, true);
  assert.deepEqual(f.links().map(a => a.href), result.links.map(link => link.url));
  assert.equal(f.links().every(a => a.target === '_blank' && a.rel === 'noopener noreferrer'), true);
  assert.equal(f.links()[3].textContent, 'Student Three — 33 — Parent Contact — 10/23/2026');
  assert.equal(f.windows.length, 0);
  assert.equal(f.elements.status.textContent, '4 entries archived. 0 of 4 links opened.');
});

test('archive failure restores controls without displaying links or claiming entries were saved', () => {
  const f = fixture(entries); f.start(); f.fail('Notes changed; reopen review.');
  assert.equal(f.elements.results.children.length, 0);
  assert.equal(f.windows.length, 0);
  assert.equal(f.elements.save.disabled, false);
  assert.equal(f.elements.status.textContent, 'Notes changed; reopen review.');
  assert.equal(f.elements.close.disabled, false);
  f.start(); assert.equal(f.requests.length, 2);
});

test('invalid dates can be unchecked and only the remaining checked entries are archived', () => {
  const f = fixture([{...entries[0], error: 'Enter a valid date in ECC Date.'}, entries[3]]);
  f.start(); assert.equal(f.requests.length, 0);
  f.boxes()[0].checked = false; f.boxes()[0].emit('change'); f.start();
  assert.deepEqual(f.requests[0].ids, ['33|Parent Contact']);
});

test('repeated archive clicks while saving or after success never archive twice', () => {
  const f = fixture(entries); f.start(); f.start();
  assert.equal(f.requests.length, 1);
  f.succeed(result); f.start();
  assert.equal(f.requests.length, 1);
});

test('opening a link tracks that entry without claiming that PowerSchool saved it', () => {
  const f = fixture(entries); f.start(); f.succeed(result);
  f.links()[3].emit('click');
  assert.equal(f.elements.results.children[3].className, 'opened');
  assert.equal(f.elements.results.children[3].children[1].textContent, '✓ opened');
  assert.equal(f.elements.status.textContent, '4 entries archived. 1 of 4 links opened.');
  assert.equal(f.requests.length, 1);
});

test('retrying a link keeps its count and marker stable without rearchiving', () => {
  const f = fixture(entries); f.start(); f.succeed(result);
  const link = f.links()[0]; link.emit('click'); link.emit('click');
  assert.equal(link.href, result.links[0].url);
  assert.equal(f.elements.results.children[0].children.length, 2);
  assert.equal(f.elements.status.textContent, '4 entries archived. 1 of 4 links opened.');
  assert.equal(f.requests.length, 1);
});

test('opening every link counts only openings, with no saved-log confirmation', () => {
  const f = fixture(entries); f.start(); f.succeed(result);
  f.links().forEach(a => a.emit('click'));
  assert.equal(f.elements.status.textContent, '4 entries archived. 4 of 4 links opened.');
  assert.doesNotMatch(f.elements.status.textContent, /confirmed|submitted/);
  assert.equal(f.windows.length, 0);
});

test('the dialog can close before or after archive, but cannot close during a save', () => {
  const cancelled = fixture(entries); cancelled.elements.close.emit('click');
  assert.equal(cancelled.closedDialog(), true);
  const f = fixture(entries); f.start(); f.elements.close.emit('click');
  assert.equal(f.closedDialog(), false);
  f.succeed(result); f.elements.close.emit('click');
  assert.equal(f.closedDialog(), true);
});

test('history-size warnings remain attached to the correct result link', () => {
  const f = fixture(entries); f.start();
  f.succeed({...result, links: result.links.map((link, i) => ({...link, nearLimit: i === 3}))});
  assert.equal(f.elements.results.children[0].children.length, 1);
  assert.match(f.elements.results.children[3].children[1], /History is close to the cell size limit/);
});

test('empty review cannot archive and displays a useful result', () => {
  const f = fixture([]); f.start();
  assert.equal(f.elements.save.disabled, true);
  assert.equal(f.requests.length, 0);
  assert.equal(f.elements.status.textContent, 'No ECC conversations, attempts or parent contacts to archive.');
});
