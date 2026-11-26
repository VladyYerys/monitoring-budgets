// Перевірка приватності: сам репозиторій чистий; правила ловлять токени, картки, пошти, Script ID; секрет не друкується.
const assert = require('assert'), fs = require('fs'), path = require('path');
const P = require('../tools/privacy_check');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const hit = s => P.RULES.filter(r => (s.match(r.re) || []).some(m => !r.ok || !r.ok(m))).map(r => r.name);

t('the repository itself: nothing found', () => {
  const r = P.run();
  assert.deepStrictEqual(r.found, []);
});

t('rules catch secrets and personal data (samples built at run time, not stored in the repo)', () => {
  const card = ['4111', '1111', '1111', '1111'].join(' ');                  // тестова картка Visa (проходить Луна)
  assert.ok(hit('картка ' + card).indexOf('номер картки') >= 0);
  assert.deepStrictEqual(hit('сума 4111 1111 1111 1112'), []);              // не проходить Луна
  assert.ok(hit(['12345678', ':', 'A'.repeat(35)].join('')).indexOf('токен Telegram-бота') >= 0);
  assert.ok(hit('UA' + '1'.repeat(27)).indexOf('IBAN') >= 0);
  assert.ok(hit('пишіть на name' + '@' + 'mail.com').indexOf('пошта') >= 0);
  assert.deepStrictEqual(hit('1234+me' + '@' + 'users.noreply.github.com'), []);
  assert.ok(hit('id ' + '1' + 'a'.repeat(56)).length > 0);
  assert.ok(hit('"refresh' + '_token": "x"').indexOf('OAuth / вхід clasp') >= 0);
});

t('every external host in the code is described in PRIVACY.md', () => {
  const priv = fs.readFileSync(path.join(__dirname, '..', 'PRIVACY.md'), 'utf8');
  P.hostsInCode().forEach((f, h) => assert.ok(priv.indexOf(h) >= 0, h));
});

console.log(n + ' tests passed');
