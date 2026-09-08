const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const script = fs.readFileSync(path.join(__dirname, '../script.js'), 'utf8');

function harness(captured, options = {}) {
  let selections = 0;
  let confirmations = 0;
  const context = {
    console,
    window: { pywebview: { api: { save_active_outlook_selection() {} } } },
    captureEmailDropSources: () => captured,
    resolveEmailRefFromCapturedDrop: async () => {
      if (options.error) throw new Error('Virtual file unavailable');
      return options.result || {};
    },
    saveActiveOutlookSelectionRef: async () => { selections++; return { raw: 'saved.msg' }; },
    confirm: () => { confirmations++; return options.confirm === true; },
  };
  vm.createContext(context);
  vm.runInContext(script.slice(script.indexOf('const OUTLOOK_DROP_TYPE_PREFIXES ='),
    script.indexOf('function captureEmailDropSources(dt) {')), context);
  vm.runInContext(script.slice(script.indexOf('async function resolvePageEmailDropRef('),
    script.indexOf('async function requestPageEmailRef(')), context);
  return { run: () => context.resolvePageEmailDropRef({}), counts: () => ({ selections, confirmations }) };
}
const empty = { files: [], entries: [], urlCandidate: '', types: [] };

test('Outlook virtual-file read failure still saves the selected email', async () => {
  const h = harness({ ...empty, types: ['FileContents'] }, { error: true });
  assert.equal((await h.run()).status, 'success');
  assert.deepEqual(h.counts(), { selections: 1, confirmations: 0 });
});
test('readable email uses the actual dropped file without consulting Outlook', async () => {
  const h = harness({ ...empty, types: ['Files'] }, { result: { emailRef: { raw: 'actual.eml' } } });
  assert.equal((await h.run()).emailRef.raw, 'actual.eml');
  assert.equal(h.counts().selections, 0);
});
test('empty generic file drop only uses Outlook after source confirmation', async () => {
  for (const accepted of [false, true]) {
    const h = harness({ ...empty, types: ['Files'] }, { confirm: accepted });
    assert.equal((await h.run()).status, accepted ? 'success' : 'error');
    assert.deepEqual(h.counts(), { selections: Number(accepted), confirmations: 1 });
  }
});
test('unsupported actual file never attaches an unrelated selected email', async () => {
  const h = harness({ ...empty, files: [{ name: 'report.pdf' }], types: ['Files'] });
  assert.equal((await h.run()).status, 'error');
  assert.deepEqual(h.counts(), { selections: 0, confirmations: 0 });
});
test('deliverable menu has no Attachments action', () => {
  const menu = script.slice(script.indexOf('function createDeliverableActionsDropdown('),
    script.indexOf('function createCardHeader('));
  assert.ok(!menu.includes('label: "Attachments"'));
});
