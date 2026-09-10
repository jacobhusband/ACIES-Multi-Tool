const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '..', 'script.js'), 'utf8');
function scope() {
  const c = vm.createContext({});
  vm.runInContext(source.slice(source.indexOf('const STATUS_CANON'), source.indexOf('const HELP_TOPICS')), c);
  vm.runInContext(source.slice(source.indexOf('function canonStatus('), source.indexOf('function normalizeRef(')), c);
  return c;
}
test('blank and legacy statuses normalize consistently without changing finished work', () => {
  const c = scope();
  for (const item of [{}, { statuses: [] }, { status: 'none' }, { status: 'Working' }, { statusTags: ['working'] }]) {
    c.migrateStatusFields(item); c.syncStatusArrays(item);
    assert.equal(item.status, 'In progress');
    assert.equal(item.statuses.join(), 'In progress');
    assert.equal(item.statusTags.join(), 'inProgress');
    c.migrateStatusFields(item); c.syncStatusArrays(item);
    assert.equal(item.statuses.length, 1);
  }
  for (const status of ['Waiting', 'On hold', 'Pending Review', 'Complete', 'Completed (by others)', 'Delivered']) {
    const item = { status };
    c.migrateStatusFields(item); c.syncStatusArrays(item);
    assert.equal(item.status, status);
  }
});
test('selection cannot toggle off or clear, and completion still checks tasks', () => {
  const c = scope();
  const item = { tasks: [{ done: false }] };
  c.setSingleStatus(item, 'In progress');
  c.toggleStatus(item, 'In progress');
  assert.equal(item.status, 'In progress');
  c.setSingleStatus(item, 'Complete');
  assert.equal(item.tasks[0].done, true);
  c.toggleStatus(item, 'Complete');
  assert.equal(item.status, 'Complete');
  c.setSingleStatus(item, '');
  assert.equal(item.status, 'In progress');
});
test('saved no-status board column migrates in place without losing visibility settings', () => {
  const c = scope();
  vm.runInContext(source.slice(source.indexOf('const DEFAULT_PROJECT_CARD_COLUMNS'), source.indexOf('let userSettings')), c);
  vm.runInContext(source.slice(source.indexOf('function getDefaultProjectCardColumn('), source.indexOf('function syncProjectViewPreferencesFromSettings(')), c);
  const cols = c.normalizeProjectCardColumns([{key:'pinned'}, {key:'none', label:'No status', hidden:true}, {key:'Waiting'}]);
  assert.equal(cols[1].key, 'In progress');
  assert.equal(cols[1].label, 'In progress');
  assert.equal(cols[1].hidden, true);
  assert.equal(cols.filter(x => x.key === 'In progress').length, 1);
  assert.equal(cols.some(x => x.key === 'none'), false);
});
