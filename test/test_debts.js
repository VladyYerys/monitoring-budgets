// «Борги й розстрочки»: аркуш, плани платежів у журналі («[Борг] …»), сплачене з фактів, вимкнення модуля, ранкові нагадування.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

function setup() {
  const env = makeEnv();
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  env.ctx.modSet_(env.ss, 'debts', true);
  return env;
}
const debt = (o) => Object.assign({ name: 'Банк', dir: 'я винен', total: 10000, monthly: 3000, day: 10, from: '', who: '', cat: 'Повернення боргу', paid: 0, left: null }, o);
const fact = (key, d, b) => ({ key: key, d: d, e: '✅ факт', m: 0, o: b, b: b, c: 'Повернення боргу', g: '' });

t('wanted plans: one per month from today to the horizon, the last one is the remainder, paid month skipped, future start respected', () => {
  const c = setup().ctx;
  let w = c.debtWanted_([debt()], [], '2026-11-25', '2026-12-31');
  assert.strictEqual(w.length, 1, '10.11 уже минуло');
  assert.strictEqual(w[0].key, '[Борг] Банк: 12.2026'); assert.strictEqual(w[0].date, '2026-12-10'); assert.strictEqual(w[0].amount, 3000);
  w = c.debtWanted_([debt({ total: 4000 })], [fact('2026-11-10', '[Борг] Банк: 11.2026', 3000)], '2026-11-05', '2026-12-31');
  assert.strictEqual(w.length, 1, 'листопад сплачено');
  assert.strictEqual(w[0].amount, 1000, 'останній платіж: залишок');
  assert.strictEqual(c.debtWanted_([debt({ total: 3000 })], [fact('2026-10-10', '[Борг] Банк', 3000)], '2026-11-05', '2026-12-31').length, 0, 'усе сплачено (тег без місяця теж рахується)');
  assert.strictEqual(c.debtWanted_([debt({ from: '2027-01-05' })], [], '2026-11-05', '2026-12-31').length, 0, 'старт пізніше');
  assert.strictEqual(c.debtWanted_([debt({ monthly: 0 })], [], '2026-11-05', '2026-12-31').length, 0, 'без щомісячного платежу планів немає');
  assert.strictEqual(c.debtPaid_([fact('2026-11-10', '[Борг] Банкір: 11.2026', 500)], 'Банк'), 0, 'інший борг з подібною назвою не рахується');
});

t('sheet: formulas «Сплачено» (tag with and without month), «Залишок», «Лишилось платежів»; «Головне» sums by direction', () => {
  const env = setup(), d = env.ss.getSheetByName('Борги');
  assert.ok(/SUMIFS\('Журнал'!\$O\$5:\$O\$3004,'Журнал'!\$D\$5:\$D\$3004,"\[Борг\] "&\$A5\)\+SUMIFS\(.*"\[Борг\] "&\$A5&":\*"\)/.test(d.getRange('I5').getFormula()), d.getRange('I5').getFormula());
  assert.strictEqual(d.getRange('J5').getFormula(), '=IF($A5="","",MAX(0,N($C5)-N($I5)))');
  assert.ok(/ROUNDUP\(\$J5\/\$D5,0\)/.test(d.getRange('K5').getFormula()));
  assert.ok(d.getRange('B5').getDataValidation().reject);
  const m = env.ss.getSheetByName('Головне'), v = m.getRange('A1:B120').getValues(), f = m.getRange('A1:B120').getFormulas();
  const i = v.findIndex(r => r[0] === 'Борги: я винен (залишок)');
  assert.strictEqual(f[i][1], "=SUMIFS('Борги'!$J$5:$J$24,'Борги'!$B$5:$B$24,\"я винен\")");
});

t('sync: plans added as «🔁 план регулярний» with the tag and «Хто», updated in place, removed when the module is off; your rows untouched', () => {
  const env = setup(), c = env.ctx, d = env.ss.getSheetByName('Борги'), j = env.ss.getSheetByName('Журнал');
  const tz = 'Europe/Kiev', today = env.Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  d.getRange('A5:H6').setValues([['Розстрочка', 'я винен', 20000, 2500, 28, '', '', ''], ['Іван', 'мені винні', 3000, 1000, 15, '', '', '']]);
  j.getRange('A5:E5').setValues([[new Date(), 99, 'Продукти', 'мій рядок', '✅ факт']]);
  const s1 = c.debtsPlansSync_(env.ss);
  assert.ok(s1.added >= 2, JSON.stringify(s1));
  const plans = () => c.readJournal_(env.ss).filter(o => o.d.indexOf('[Борг] ') === 0);
  assert.ok(plans().every(o => o.e === '🔁 план регулярний'));
  assert.ok(plans().some(o => /^\[Борг\] Розстрочка: \d\d\.\d{4}$/.test(o.d) && o.b === 2500 && o.c === 'Повернення боргу'));
  assert.ok(plans().some(o => o.d.indexOf('[Борг] Іван') === 0 && o.c === 'Мені повернули борг' && o.b === 1000));
  assert.ok(plans().every(o => o.key >= today));
  d.getRange('D5').setValue(3000);
  const s2 = c.debtsPlansSync_(env.ss);
  assert.ok(s2.updated >= 1 && s2.added === 0, JSON.stringify(s2));
  assert.ok(plans().filter(o => o.d.indexOf('[Борг] Розстрочка') === 0).every(o => o.b === 3000));
  const st = env.ss.getSheetByName('Налаштування'), r = 44 + vm.runInContext('MODS', c).findIndex(m => m.id === 'debts');
  st.getRange(r, 2).setValue('ні');
  c.mbsEdit({ range: st.getRange(r, 2) });
  assert.strictEqual(plans().length, 0, 'модуль вимкнено: майбутні плани прибрано');
  assert.strictEqual(env.ss.getSheetByName('Борги').isSheetHidden(), true);
  assert.strictEqual(d.getRange('A5').getValue(), 'Розстрочка', 'дані аркуша лишились');
  assert.ok(c.readJournal_(env.ss).some(o => o.d === 'мій рядок'));
});

t('morning: remaining debts line and a reminder the day before and on the day', () => {
  const env = setup(), c = env.ctx, d = env.ss.getSheetByName('Борги'), j = env.ss.getSheetByName('Журнал');
  const tz = 'Europe/Kiev', today = env.Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  d.getRange('A5:H5').setValues([['Розстрочка', 'я винен', 20000, 2500, 28, '', '', '']]);
  d.put(5, 10, 17500, d.cell(5, 10).f);
  j.getRange('A5:E5').setValues([[env.Utilities.parseDate(c.addDays_(today, 1), tz, 'yyyy-MM-dd'), 2500, 'Повернення боргу', '[Борг] Розстрочка: 12.2026', '🔁 план регулярний']]);
  const r = c.buildMorning_(env.ss, {});
  assert.ok(/🤝 Борги: я винен 17 500 ₴/.test(r.text), r.text);
  assert.ok(/🤝 <b>Розстрочка: 12\.2026<\/b>: 2 500 ₴ \(Повернення боргу\), завтра/.test(r.text), r.text);
  c.modSet_(env.ss, 'debts', false);
  assert.ok(!/🤝/.test(c.buildMorning_(env.ss, {}).text));
});

console.log(n + ' tests passed');
