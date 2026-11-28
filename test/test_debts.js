// «Борги й розстрочки»: аркуш, плани платежів у журналі («[Борг] …»), сплачене з фактів, вимкнення модуля, ранкові нагадування.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
const K = (c, s) => vm.runInContext(s, c);
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
  assert.strictEqual(d.getRange('J5').getFormula(), '=IF($A5="","",IF($N5<>"",0,MAX(0,N($C5)-N($I5))))', 'закрито (N): залишок 0');
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

/* ---------- 💳 покупки частинами й розстрочки (0.5.4) ---------- */

const TOKEN = 'uTestToken' + 'q'.repeat(30);
function bankSetup(items) {
  const state = { items: items || [], accounts: [{ id: 'acc1', currencyCode: 980, type: 'black', creditLimit: 0, balance: 4000000, iban: 'UA00TEST', maskedPan: ['000000******4321'] }] };
  const env = makeEnv({ mono: url => {
    if (/client-info$/.test(url)) return { code: 200, body: { name: 'Тест', accounts: state.accounts } };
    if (/\/personal\/statement\//.test(url)) return { code: 200, body: state.items };
    return { code: 404, body: 'no' };
  }, promptText: TOKEN });
  env.ss.insertSheet('Аркуш1'); env.ctx.setupTemplate(); env.ctx.modSet_(env.ss, 'debts', true);
  env.state = state; env.ctx.connectMono();
  return env;
}
const dkey = (env, d) => env.Utilities.formatDate(new Date(Date.now() + d * 86400000), 'Europe/Kiev', 'yyyy-MM-dd');
const dsec = (env, key, hhmm) => Math.floor(env.Utilities.parseDate(key + ' ' + (hhmm || '12:00') + ':00', 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss').getTime() / 1000);

t('💳 pure partsOf_: «Платіж МАГАЗИН - monomarket» first payment, «Погашення …» payment, «Повне погашення …» final (also without a shop: instalment to the card); a plain purchase or «Платіж» without monomarket: no', () => {
  const c = setup().ctx, P = x => JSON.stringify(c.partsOf_(x));
  assert.strictEqual(P('Платіж MEBLIMAX - monomarket'), JSON.stringify({ kind: 'first', shop: 'MEBLIMAX', mono: true }));
  assert.strictEqual(P('Погашення MEBLIMAX - monomarket'), JSON.stringify({ kind: 'pay', shop: 'MEBLIMAX', mono: true }));
  assert.strictEqual(P('Повне погашення Petshop Test - monomarket'), JSON.stringify({ kind: 'final', shop: 'Petshop Test', mono: true }));
  assert.strictEqual(P('Повне погашення'), JSON.stringify({ kind: 'final', shop: '', mono: false }));
  assert.strictEqual(P('Погашення'), JSON.stringify({ kind: 'pay', shop: '', mono: false }));
  ['monomarket', 'Платіж за послуги', 'АТБ', 'Переказ на картку', ''].forEach(x => assert.strictEqual(c.partsOf_(x), null, x));
});

t('💳 bank: the first payment adds a row on «Борги» (name and bank name = shop, payment, day, first date, kind «покупка частинами», sum empty → «⚠️ впишіть суму»); the payment is a fact «Кредит / розстрочка» tagged «[Борг] …: мм.рррр» (counts as paid)', () => {
  const env = bankSetup(), c = env.ctx, k = dkey(env, -2);
  env.state.items = [{ id: 'f1', time: dsec(env, k, '06:25'), description: 'Платіж MEBLIMAX - monomarket', mcc: 5732, amount: -45030, balance: 3954970 }];
  const r = c.importBank_(env.ss), sh = env.ss.getSheetByName('Борги');
  const row = sh.getRange('A5:N5').getValues()[0];
  assert.strictEqual(row[0], 'MEBLIMAX'); assert.strictEqual(row[1], 'я винен'); assert.strictEqual(row[2], ''); assert.strictEqual(row[3], 450.3);
  assert.strictEqual(row[4], Math.min(28, Number(k.slice(8)))); assert.strictEqual(c.dateKey_(row[5], 'Europe/Kiev'), k);
  assert.strictEqual(JSON.stringify([row[11], row[12], row[13]]), JSON.stringify(['покупка частинами', 'MEBLIMAX', '']));
  const j = c.readJournal_(env.ss).find(o => o.bank[0] === 'f1');
  assert.strictEqual(JSON.stringify([j.c, j.e, j.b]), JSON.stringify(['Кредит / розстрочка', '✅ факт', 450.3]));
  assert.ok(j.d.indexOf('[Борг] MEBLIMAX: ' + k.slice(5, 7) + '.' + k.slice(0, 4) + ' · Платіж MEBLIMAX - monomarket') === 0, j.d);
  assert.strictEqual(r.parts, 1); assert.strictEqual(JSON.stringify(r.partsNew), JSON.stringify(['MEBLIMAX']));
  assert.ok(c.importReport_(r).some(l => /Нова покупка частинами на «Борги»: <b>MEBLIMAX<\/b>/.test(l)));
  assert.ok(/O5/.test('O5') && /впишіть суму \(C\)/.test(sh.getRange('O5').getFormula()), 'прогрес: без суми підказує');
  // удруге той самий імпорт: рядок не задвоюється
  c.importBank_(env.ss);
  assert.strictEqual(sh.getRange('A5:A8').getValues().filter(x => x[0]).length, 1);
});

t('💳 bank: with the purchase sum the payments become plans; «Погашення МАГАЗИН» closes its plan whatever the category (even yours: «Дім і побут»); «Повне погашення» writes «Закрито» and later plans disappear', () => {
  const env = bankSetup(), c = env.ctx, sh = env.ss.getSheetByName('Борги'), today = dkey(env, 0);
  const D = x => env.Utilities.parseDate(x, 'Europe/Kiev', 'yyyy-MM-dd');
  const day = Math.min(28, Number(dkey(env, 1).slice(8)));
  sh.getRange('A5:H5').setValues([['Навушники', 'я винен', 5400, 450.3, day, '', '', 'Дім і побут']]);
  sh.getRange('L5:M5').setValues([['покупка частинами', 'MEBLIMAX']]);
  c.debtsPlansSync_(env.ss);
  const plans = () => c.readJournal_(env.ss).filter(o => o.d.indexOf('[Борг] Навушники:') === 0 && o.e === '🔁 план регулярний');
  const p0 = plans();
  assert.ok(p0.length >= 1 && p0[0].c === 'Дім і побут', JSON.stringify(p0.map(o => [o.key, o.c])));
  const pk = p0.map(o => o.key).sort()[0];                              // найближчий план
  env.state.items = [{ id: 'p1', time: dsec(env, today, '06:25'), description: 'Погашення MEBLIMAX - monomarket', mcc: 5732, amount: -45030, balance: 3954970 }];
  let r = c.importBank_(env.ss);
  const row = c.readJournal_(env.ss).find(o => o.bank[0] === 'p1');
  assert.strictEqual(row.c, 'Дім і побут', 'категорія плану'); assert.strictEqual(row.e, '✅ факт'); assert.ok(row.d.indexOf('[Борг] Навушники: ' + pk.slice(5, 7) + '.') === 0, row.d);
  assert.strictEqual(r.plans, 1); assert.strictEqual(r.parts, 1);
  assert.ok(!plans().some(o => o.key === pk), 'план місяця закрито, не лишився поруч з фактом');
  // повне погашення: «Закрито», плани прибрано
  env.state.items.push({ id: 'p2', time: dsec(env, today, '14:51'), description: 'Повне погашення MEBLIMAX - monomarket', mcc: 5732, amount: -405000, balance: 3549970 });
  r = c.importBank_(env.ss);
  assert.strictEqual(c.dateKey_(sh.getRange('N5').getValue(), 'Europe/Kiev'), today);
  assert.strictEqual(JSON.stringify(r.partsClosed), JSON.stringify(['Навушники']));
  assert.strictEqual(plans().length, 0, 'закрито: майбутніх планів немає');
  assert.strictEqual(c.debtsList_(env.ss).length, 0); assert.strictEqual(c.debtsAll_(env.ss)[0].closed, today);
  assert.ok(/✓ закрито/.test(sh.getRange('O5').getFormula()));
});

t('💳 pure planImport_: instalment to the card «Повне погашення» without a shop closes the only open «розстрочка» with that payment; an unknown «[Борг]» plan is taken only by the same sum ±1 % and ±5 days; module off: just «Кредит / розстрочка», no new rows', () => {
  const env = setup(), c = env.ctx, k = '2026-11-26';
  const jr = o => Object.assign({ row: 100, key: k, b: 0, c: '', d: '', e: '', bank: ['', '', '', '', '', ''], acct: '', blank: false }, o);
  const ctx = (items, journal, debts) => ({ items: items, journal: journal, logged: new Set(), rules: c.readRules_(env.ss), dateOf: () => k, timeOf: () => '',
    acc: { name: 'Mono', kind: 'основна картка', who: '' }, types: c.catTypes_(env.ss), autoFact: true, debts: debts });
  const inst = { row: 7, name: 'Розстрочка на картку', bank: '', dir: 'я винен', kind: 'розстрочка', monthly: 550, cat: 'Кредит / розстрочка', closed: '' };
  let p = c.planImport_(ctx([{ id: 'x1', time: 1, amount: -55000, description: 'Повне погашення', mcc: 0 }], [], [inst]));
  assert.strictEqual(JSON.stringify(p.closeDebts), JSON.stringify([{ row: 7, name: 'Розстрочка на картку', key: k }]));
  assert.ok(p.drafts[0].d.indexOf('[Борг] Розстрочка на картку: 11.2026 · ') === 0, p.drafts[0].d);
  const plan = jr({ row: 50, key: '2026-11-24', b: 900, c: 'Повернення боргу', d: '[Борг] Крісло: 11.2026', e: '🔁 план регулярний' });
  p = c.planImport_(ctx([{ id: 'x2', time: 2, amount: -90000, description: 'Погашення Невідомий - monomarket', mcc: 0 }], [plan], []));
  assert.strictEqual(p.takes.length, 1, 'та сама сума, 2 дні'); assert.strictEqual(p.takes[0].c, 'Повернення боргу');
  p = c.planImport_(ctx([{ id: 'x3', time: 3, amount: -140000, description: 'Погашення Невідомий - monomarket', mcc: 0 }], [plan], []));
  assert.strictEqual(p.takes.length, 0, 'інша сума: не той план');
  p = c.planImport_(ctx([{ id: 'x4', time: 4, amount: -45030, description: 'Платіж MEBLIMAX - monomarket', mcc: 0 }], [], null));
  assert.strictEqual(JSON.stringify([p.drafts[0].c, p.drafts[0].d, p.newDebts.length]), JSON.stringify(['Кредит / розстрочка', 'Платіж MEBLIMAX - monomarket', 0]));
});

t('💳 limits: used = remaining of «покупка частинами» and «розстрочка» rows; «так» in G writes total = available in Mono + used; formulas check against Mono; header K shows debts of the month and the limit; step 17 upgrades an old «Борги»; the sheet is on view with the module', () => {
  const env = setup(), c = env.ctx, sh = env.ss.getSheetByName('Борги');
  c.modulesApply_(env.ss, true);
  assert.strictEqual(sh.isSheetHidden(), false, 'модуль увімкнено: на виду');
  sh.getRange('A5:F7').setValues([['Навушники', 'я винен', 24000, 1000, 26, ''], ['Крісло', 'я винен', 18000, 900, 26, ''], ['Позика', 'я винен', 5000, 1000, 5, '']]);
  sh.getRange('L5:L7').setValues([['покупка частинами'], ['розстрочка'], ['позика']]);
  const j = env.ss.getSheetByName('Журнал'), D = x => env.Utilities.parseDate(x, 'Europe/Kiev', 'yyyy-MM-dd');
  j.getRange('A20:E20').setValues([[D('2026-10-26'), 1000, 'Кредит / розстрочка', '[Борг] Навушники: 10.2026', '✅ факт']]);
  const rows = c.readJournal_(env.ss).map(o => Object.assign(o, { m: 0, o: o.b }));
  assert.strictEqual(c.debtsUsed_(c.debtsAll_(env.ss), rows, ['покупка частинами', 'розстрочка']), 24000 - 1000 + 18000);
  sh.getRange('C28').setValue(80000); sh.getRange('G28').setValue('так');
  c.mbsEdit({ range: sh.getRange('G28') });
  assert.strictEqual(sh.getRange('B28').getValue(), Math.round((80000 + 24000 - 1000 + 18000) * 100) / 100);
  assert.strictEqual(sh.getRange('G28').getValue(), '', '«так» прибрано');
  assert.ok(/SUMIFS\(\$J\$5:\$J\$24,\$L\$5:\$L\$24,"покупка частинами"\)\+SUMIFS\(\$J\$5:\$J\$24,\$L\$5:\$L\$24,"розстрочка"\)/.test(sh.getRange('D28').getFormula()), sh.getRange('D28').getFormula());
  assert.ok(/✓ сходиться з Mono/.test(sh.getRange('F28').getFormula()) && /не сходиться/.test(sh.getRange('F28').getFormula()));
  assert.ok(/SUMIFS\(\$J\$5:\$J\$24,\$L\$5:\$L\$24,"кредит готівкою"\)/.test(sh.getRange('D29').getFormula()));
  // шапка журналу K1:K3
  const k = j.getRange('K1:K3').getFormulas().map(x => x[0]);
  assert.ok(/Борги цього місяця/.test(k[0]) && /Найближчий/.test(k[1]) && /Ліміт частинами/.test(k[2]) && /'Борги'!\$B\$28/.test(k[2]), k.join('\n'));
  c.modSet_(env.ss, 'debts', false); c.modulesApply_(env.ss, true);
  assert.strictEqual(j.getRange('K1').getFormula(), '', 'модуль вимкнено: K порожня'); assert.strictEqual(sh.isSheetHidden(), true);
  c.modSet_(env.ss, 'debts', true); c.modulesApply_(env.ss, true);
  // крок 17 для старого аркуша: лише A…K, без лімітів
  sh.getRange('L4:O4').setValues([['', '', '', '']]); sh.getRange('A27:G29').setValues([['', '', '', '', '', '', ''], ['', '', '', '', '', '', ''], ['', '', '', '', '', '', '']]);
  sh.getRange('J5').setFormula('=IF($A5="","",MAX(0,N($C5)-N($I5)))');
  env.docProps.setProperty('schema_v', '16');
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 17).done(env.ss), false);
  c.schemaTick_(env.ss);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.strictEqual(sh.getRange('L4').getValue(), 'Вид'); assert.strictEqual(sh.getRange('A27').getValue(), '💳 Ліміти Mono (застосунок → «Кредити»)');
  assert.ok(/\$N5<>""/.test(sh.getRange('J5').getFormula()));
  assert.strictEqual(sh.getRange('A5').getValue(), 'Навушники', 'ваші рядки на місці'); assert.strictEqual(sh.getRange('L6').getValue(), 'розстрочка');
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 17).done(env.ss), true);
  assert.ok(/частинами й розстрочок: 2/.test(c.debtsState_(env.ss)), c.debtsState_(env.ss));
});

console.log(n + ' tests passed');
