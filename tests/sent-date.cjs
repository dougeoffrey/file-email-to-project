const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const ts = require('typescript');
const source = fs.readFileSync('src/taskpane/outlook.ts', 'utf8');
const code = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText;
async function check(header, zone, expected, failure = false) {
  const elements = new Map();
  const element = id => {
    if (!elements.has(id)) elements.set(id, { value: id === 'folder-description' ? 'Markups' : 'folder', textContent: '', disabled: false });
    return elements.get(id);
  };
  const context = vm.createContext({
    exports: {}, require: () => ({}),
    document: { getElementById: element },
    Office: { onReady() {}, AsyncResultStatus: { Succeeded: 'ok' }, context: { mailbox: {
      item: { itemId: 'email-1', getAllInternetHeadersAsync(callback) { callback(failure ? { status: 'error', error: { message: 'Unavailable' } } : { status: 'ok', value: header }); } },
      convertToLocalClientTime(date) {
        const parts = Object.fromEntries(new Intl.DateTimeFormat('en-US', { timeZone: zone, year: 'numeric', month: 'numeric', day: 'numeric' }).formatToParts(date).map(p => [p.type, p.value]));
        return { year: +parts.year, month: +parts.month - 1, date: +parts.day };
      }
    } } }
  });
  vm.runInContext(code, context);
  vm.runInContext('selectedEmailAttachmentCount = 1', context);
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
  await check('Date: Thu, 1 Jan 2026 01:30:00 +0000\r\n', 'America/Los_Angeles', '251231');
  await check('dAtE: Tue, 22 Sep 2026\r\n\t23:30:00 -0700\r\n', 'Asia/Tokyo', '260923');
  await check('Date: Sun, 8 Mar 2026 09:30:00 +0000\r\n', 'America/Los_Angeles', '260308');
  await check('Date: Sun, 8 Mar 2026 10:30:00 +0000\r\n', 'America/Los_Angeles', '260308');
  await check('Received: today\r\n', 'UTC', '');
  await check('Date: invalid\r\n', 'UTC', '');
  await check('', 'UTC', '', true);
  console.log('7 sent-date cases passed: rollover, folded headers, DST, missing/invalid date, API failure; preview and stale-item guards checked.');
})().catch(error => { console.error(error); process.exitCode = 1; });
