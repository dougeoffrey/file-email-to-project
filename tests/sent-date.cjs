const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const ts = require('typescript');
const source = fs.readFileSync('src/taskpane/outlook.ts', 'utf8');
const code = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText;
async function check(sentDateTime, zone, expected, failure = false) {
  const elements = new Map();
  const element = id => {
    if (!elements.has(id)) elements.set(id, { value: id === 'folder-description' ? 'Markups' : 'folder', textContent: '', disabled: false });
    return elements.get(id);
  };
  const context = vm.createContext({
    exports: {}, require: () => ({}),
    fetch: async (url, options) => {
      assert.equal(url, 'https://graph.microsoft.com/v1.0/me/messages/rest%2Femail%2B1%3D?$select=sentDateTime');
      assert.equal(options.headers.Authorization, 'Bearer test-token');
      return { ok: !failure, status: 403, text: async () => 'Access denied', json: async () => ({ sentDateTime }) };
    },
    document: { getElementById: element },
    Office: { onReady() {}, MailboxEnums: { RestVersion: { v2_0: 'v2.0' } }, context: { mailbox: {
      // No internet headers API: internal/sent messages must still work.
      item: { itemId: 'email-1' },
      convertToRestId(id, version) { assert.equal(id, 'email-1'); assert.equal(version, 'v2.0'); return 'rest/email+1='; },
      convertToLocalClientTime(date) {
        const parts = Object.fromEntries(new Intl.DateTimeFormat('en-US', { timeZone: zone, year: 'numeric', month: 'numeric', day: 'numeric' }).formatToParts(date).map(p => [p.type, p.value]));
        return { year: +parts.year, month: +parts.month - 1, date: +parts.day };
      }
    } } }
  });
  vm.runInContext(code, context);
  vm.runInContext('selectedEmailAttachmentCount = 1; currentAccessToken = "test-token"', context);
  await vm.runInContext('loadSelectedEmailSentDate()', context);
  assert.equal(vm.runInContext('buildFolderName()', context), expected ? `${expected}_Markups` : '');
  assert.equal(element('save-button').disabled, !expected);
  if (expected) {
    assert.equal(element('folder-name-preview').textContent, `${expected}_Markups`);
    element('folder-description').value = '';
    vm.runInContext('updateSaveButton()', context);
    assert.equal(element('folder-name-preview').textContent, `${expected}_…`);
    element('folder-description').value = 'Markups';
    vm.runInContext('Office.context.mailbox.item.itemId = "email-2"', context);
    assert.equal(vm.runInContext('buildFolderName()', context), '');
  }
}
(async () => {
  await check('2026-01-01T01:30:00Z', 'America/Los_Angeles', '251231');
  await check('2026-09-23T06:30:00Z', 'Asia/Tokyo', '260923');
  await check('2026-03-08T09:30:00Z', 'America/Los_Angeles', '260308');
  await check('2026-03-08T10:30:00Z', 'America/Los_Angeles', '260308');
  await check(undefined, 'UTC', '');
  await check('invalid', 'UTC', '');
  await check('', 'UTC', '', true);
  console.log('7 sent-date cases passed: headerless messages, rollover, DST, missing/invalid date, API failure; preview and stale-item guards checked.');
})().catch(error => { console.error(error); process.exitCode = 1; });
