// ci/guard.js: заливка лише в порожній скрипт або в скрипт із шаблоном; чужий код (наприклад, особистої таблиці) → стоп.
const assert = require('assert'), fs = require('fs'), os = require('os'), path = require('path');
const { check } = require('../ci/guard');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const dir = files => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'g-')); Object.keys(files).forEach(f => { fs.mkdirSync(path.dirname(path.join(d, f)), { recursive: true }); fs.writeFileSync(path.join(d, f), files[f]); }); return d; };

t('fresh script: only manifest, or default Code with empty myFunction → ok', () => {
  assert.ok(check(dir({ 'appsscript.json': '{}', '.clasp.json': '{}' })).ok);
  assert.ok(check(dir({ 'appsscript.json': '{}', 'Code.js': 'function myFunction() {\n  \n}\n' })).ok);
});

t('script that already has the template mark → ok', () => {
  assert.ok(check(dir({ '00_Config.js': "const TEMPLATE_MARK = 'monitoring-budgets-template';", '10_Journal.js': 'x' })).ok);
});

t('any other code without the mark → refused (wrong SCRIPT_ID)', () => {
  const r = check(dir({ '01_Journal.js': 'function onOpen() {}', 'migrations/m01.js': 'x', 'appsscript.json': '{}' }));
  assert.ok(!r.ok && /SCRIPT_ID/.test(r.why), r.why);
  assert.ok(!check(dir({ 'Code.js': 'function myFunction() { doSomething(); }' })).ok);
});

console.log(n + ' tests passed');
