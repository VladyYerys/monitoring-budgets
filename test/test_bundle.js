// «Весь код одним файлом» для Apps Script: той самий результат, що й файли окремо (можна вставити в Code.gs без GitHub Actions).
const assert = require('assert');
const { makeEnv } = require('./gasmock');
const { codeFile } = require('../tools/bundle');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('one file builds the same template as separate files', () => {
  const code = codeFile();
  assert.ok(code.indexOf('monitoring-budgets-template') > 0, 'мітка шаблону в коді (її шукає ci/guard.js)');
  const a = makeEnv({ code }), b = makeEnv();
  [a, b].forEach(e => { e.ss.insertSheet('Аркуш1'); e.ctx.setupTemplate(); });
  const shape = e => JSON.stringify(e.ss.getSheets().map(s => [s.getName(), Array.from(s.cells.values()).filter(x => x.f).length]));
  assert.strictEqual(shape(a), shape(b));
  assert.strictEqual(a.triggers.length, b.triggers.length);
});

console.log(n + ' tests passed');
