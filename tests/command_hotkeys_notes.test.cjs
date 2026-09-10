const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../script.js'), 'utf8');
function extract(name) {
  return source.match(new RegExp(`^(?:async )?function ${name}\\([\\s\\S]*?^}\\r?$`, 'm'))[0];
}
test('hotkeys retain explicit empty bindings, ignore missing workflows, and pass target context', async () => {
  const calls = [];
  const context = vm.createContext({
    userSettings: { commandHotkeys: { 1: 'workflow:w', 2: '', 3: 'workflow:deleted' } },
    getUserWorkflows: () => [{ id: 'w', name: 'Prepare and publish' }],
    getDeliverableToolMenuEntries: () => [], DELIVERABLE_QUICK_ACCESS_ACTIONS: [],
    buildProjectsTabToolLaunchContext: (project, deliverable) => ({ project, deliverable }),
    runWorkflow: async (...args) => calls.push(args),
  });
  vm.runInContext(['getWorkflowDisplayName', 'getCommandHotkeyBindings', 'getCommandHotkeyOptions', 'getCommandHotkeyEntries'].map(extract).join('\n'), context);
  const entries = context.getCommandHotkeyEntries();
  assert.equal(entries.length, 1);
  await entries[0].run({ project: 'p', deliverable: 'd' });
  assert.equal(calls[0][0], 'w');
  assert.equal(calls[0][1].project, 'p');
  assert.equal(calls[0][1].deliverable, 'd');
  context.userSettings.commandHotkeys = {};
  assert.equal(context.getCommandHotkeyEntries().length, 0);
});
test('attachment migration preserves files, URLs and email metadata, escapes HTML, and runs once', () => {
  const context = vm.createContext({ normalizeAttachments: a => a, Date });
  vm.runInContext(extract('escapeHtml') + '\n' + extract('migrateDeliverableAttachmentsToProjectNotes'), context);
  const project = { page: { html: '<p>Existing notes</p>' }, deliverables: [{ name: '<Permit>', attachments: [
    { type: 'path', target: 'C:\\Plans\\a.pdf', description: 'Plan "A"' },
    { type: 'url', target: 'https://example.com/?a=1&b=2', description: 'Reference' },
    { type: 'email', description: 'Approval', emailRef: { raw: 'mail.msg', messageId: '123' } },
  ] }] };
  context.migrateDeliverableAttachmentsToProjectNotes(project);
  const html = project.page.html;
  assert.ok(html.startsWith('<p>Existing notes</p>'));
  assert.ok(html.includes('&lt;Permit&gt;'));
  assert.ok(html.includes('data-attachment-target="C:\\Plans\\a.pdf"'));
  assert.ok(html.includes('a=1&amp;b=2'));
  assert.ok(html.includes('data-email-message-id="123"'));
  context.migrateDeliverableAttachmentsToProjectNotes(project);
  assert.equal(project.page.html, html);
  assert.equal(project.deliverables[0].attachments.length, 3);
});
