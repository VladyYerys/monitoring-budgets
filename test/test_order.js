// Порядок журналу (майбутнє зверху, лінія, рядки для вводу, минуле від нового) і onEdit (автодата, автокатегорія, «навчання» правил).
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);

function built() {
  const env = makeEnv();
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  return env;
}
const row = (key, b, c, s, bank) => ({ v: [key, b, c, '', '', '', '', ''], b: bank || null, s: s === undefined ? '' : s, key: key });

t('planOrder_: future (nearest first) → 4 empty rows → past (newest first); within a day by write order', () => {
  const c = makeEnv().ctx, today = '2026-11-25';
  const rows = [row('2026-11-20', 10, 'A', 1), row('2026-12-05', 20, 'B', 2), row('2026-11-25', 30, 'C', 3), row('2026-11-25', 40, 'D', 4),
    row('2026-11-28', 50, 'E', 5)];
  for (let i = 0; i < 10; i++) rows.push(row('', '', '', ''));
  const o = c.planOrder_(rows, today, { now: 1e12, waitIncomplete: false });
  const keys = o.map(x => x.v[1]);
  assert.strictEqual(JSON.stringify(keys), JSON.stringify([50, 20, '', '', '', '', 40, 30, 10]));
  assert.strictEqual(c.planOrder_(o, today, { now: 1e12, waitIncomplete: false }), null, 'вдруге: нічого міняти');
});

t('planOrder_: half-written row waits (unless 10 min passed); bank rows without category do not block; new rows get a write stamp', () => {
  const c = makeEnv().ctx, today = '2026-11-25';
  const rows = [row('2026-11-25', 100, '', 1), row('2026-11-24', 5, 'A', 2)];
  for (let i = 0; i < 6; i++) rows.push(row('', '', '', ''));
  assert.strictEqual(c.planOrder_(rows, today, { now: 1, waitIncomplete: true }), null);
  const bank = [row('2026-11-24', 5, 'A', 2), row('2026-11-25', 100, '', 1, ['id1', 'opis', -100, 5411, '25.11.2026 10:00', -50])];
  for (let i = 0; i < 6; i++) bank.push(row('', '', '', ''));
  const o = c.planOrder_(bank, today, { now: 1, waitIncomplete: true });
  assert.ok(o && o[4].b[0] === 'id1', 'рядок з банку без категорії не блокує і їде разом зі своїми колонками банку');
  const fresh = [row('2026-11-24', 5, 'A', '')];
  for (let i = 0; i < 6; i++) fresh.push(row('', '', '', ''));
  assert.strictEqual(c.planOrder_(fresh, today, { now: 777, waitIncomplete: false })[4].s, 777);
});

t('sortCore_ on the sheet: A…H and Y…AE move together, formulas stay in place', () => {
  const env = built(), j = env.ss.getSheetByName('Журнал'), d = k => env.Utilities.parseDate(k, 'Europe/Kiev', 'yyyy-MM-dd');
  j.getRange('A5:H5').setValues([[d('2026-01-10'), 10, 'Продукти', 'старе', '✅ факт', '', '', '']]);
  j.getRange('Y5:AE5').setValues([['b1', 'АТБ', -10, 5411, '10.01.2026 10:00', 100, 5]]);
  j.getRange('A6:H6').setValues([[d('2026-01-12'), 20, 'Таксі', 'нове', '✅ факт', '', '', '']]);
  const f = j.getRange('I5').getFormula();
  assert.strictEqual(env.ctx.sortCore_(env.ss), true);
  const top = j.getRange('A9:H9').getValues()[0], bank = j.getRange('Y10:AE10').getValues()[0];
  assert.strictEqual(top[3], 'нове');
  assert.strictEqual(j.getRange('D10').getValue(), 'старе');
  assert.strictEqual(bank[0], 'b1');
  assert.strictEqual(j.getRange('I5').getFormula(), f);
});

t('onEdit: date filled automatically; category from a comment rule; changing category of a bank row suggests a rule', () => {
  const env = built(), j = env.ss.getSheetByName('Журнал'), c = env.ctx;
  j.getRange('B9:D9').setValues([[55, '', 'вранці кава біля роботи']]);
  c.onEdit({ range: j.getRange('B9:D9') });
  assert.ok(j.getRange('A9').getValue() instanceof Date);
  assert.strictEqual(j.getRange('C9').getValue(), 'Кава та перекуси');
  assert.ok(Number(env.docProps.getProperty('lastEdit')) > 0);
  j.getRange('A10:E10').setValues([[new Date(), 300, 'Інші витрати', '', '🏦 чернетка']]);
  j.getRange('Y10:Z10').setValues([['b9', 'Квіткова лавка 24']]);
  c.onEdit({ range: j.getRange('C10') });
  const rules = env.ss.getSheetByName('Правила').getRange('A5:G204').getValues().filter(r => r[0] === '✨ нове, підтвердіть');
  assert.strictEqual(rules.length, 1);
  assert.strictEqual(rules[0][1], 'Квіткова лавка');
  assert.strictEqual(rules[0][4], 'Інші витрати');
});

t('matchRule_: MCC ranges, direction, comment-only, first match wins', () => {
  const c = makeEnv().ctx;
  const rules = c.parseRules_([['так', 'bolt food', 'опис банку', 'списання', 'Доставка їжі'], ['так', 'bolt', 'опис банку', 'списання', 'Таксі'],
    ['так', '5812-5814', 'MCC', 'усі', 'Кафе й ресторани'], ['так', 'зарплата', 'будь-де', 'надходження', 'Зарплата'], ['ні', 'atb', 'опис банку', 'усі', 'Продукти']]);
  assert.strictEqual(c.matchRule_(rules, { desc: 'BOLT FOOD Kyiv' }, -100).category, 'Доставка їжі');
  assert.strictEqual(c.matchRule_(rules, { desc: 'Bolt.eu' }, -100).category, 'Таксі');
  assert.strictEqual(c.matchRule_(rules, { desc: 'Кав\'ярня', mcc: 5813 }, -50).category, 'Кафе й ресторани');
  assert.strictEqual(c.matchRule_(rules, { desc: 'Зарплата за жовтень' }, -100), null, 'напрям «надходження»');
  assert.strictEqual(c.matchRule_(rules, { desc: 'ATB 123' }, -10), null, 'вимкнене правило');
  assert.strictEqual(c.matchRule_(rules, { comment: 'зарплата' }, 0, true).category, 'Зарплата');
});

t('mbsTick: sorts after quiet time and closes the previous month in «Історія місяців»', () => {
  const env = built(), c = env.ctx;
  env.docProps.setProperty('lastEdit', String(Date.now() - 60000));
  const set = env.ss.getSheetByName('Налаштування');
  c.mbsTick();
  assert.strictEqual(env.ss.getSheetByName('Історія місяців').getRange('A4').getValue(), '', 'місяць до старту обліку не пишеться');
  env.docProps.deleteProperty('mh_last');
  set.getRange('B4').setValue(new Date(Date.UTC(2025, 0, 1, 10)));
  c.mbsTick();
  assert.ok(env.docProps.getProperty('sort_day'));
  assert.ok(env.docProps.getProperty('mh_last'), 'минулий місяць записано в «Історію місяців»');
  const h = env.ss.getSheetByName('Історія місяців').getRange('A4').getValue();
  assert.strictEqual(String(h), env.docProps.getProperty('mh_last'));
});

console.log(n + ' tests passed');
