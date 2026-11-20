const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

test('Update exposes individual ECC logging at the top level and in its existing submenu', () => {
  const menus = [];
  const ui = {
    createMenu(name) {
      const menu = {
        name, entries: [],
        addItem(label, handler) { this.entries.push({label, handler}); return this; },
        addSeparator() { this.entries.push({separator: true}); return this; },
        addSubMenu(submenu) { this.entries.push({submenu}); return this; },
        addToUi() { menus.push(this); return this; },
      };
      return menu;
    },
  };
  const context = vm.createContext({SpreadsheetApp: {getUi: () => ui}});
  vm.runInContext(fs.readFileSync('Menus.gs', 'utf8'), context);
  vm.runInContext(fs.readFileSync('LogECC.gs', 'utf8'), context);
  vm.runInContext('onOpen()', context);

  assert.equal(menus.length, 1);
  const update = menus[0];
  assert.equal(update.name, 'Update');
  const firstSeparator = update.entries.findIndex(entry => entry.separator);
  const quick = update.entries.slice(0, firstSeparator);
  const ecc = update.entries.find(entry => entry.submenu?.name === 'ECC & Demographics').submenu;
  for (const [word, handler] of [
    ['conversation', 'archiveEccConversationAndOpenPowerSchool'],
    ['attempt', 'archiveEccAttemptAndOpenPowerSchool'],
  ]) {
    const item = quick.find(entry => entry.handler === handler);
    assert.ok(item, 'Quick access for ' + word);
    assert.match(item.label, new RegExp('selected student.*' + word));
    assert.ok(ecc.entries.some(entry => entry.handler === handler && entry.label === item.label));
    assert.equal(vm.runInContext('typeof ' + handler, context), 'function');
  }
  for (const handler of ['archiveAndLogAllEcc', 'openStudentDemographics']) {
    assert.ok(quick.some(entry => entry.handler === handler), handler);
  }
  const wig = update.entries.find(entry => entry.submenu?.name === 'WIG').submenu;
  assert.ok(wig.entries.some(entry => entry.handler === 'previewWigSnapshot'));
  assert.ok(wig.entries.some(entry => entry.handler === 'saveWigSnapshot'));
});
