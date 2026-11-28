// «Регулярні»: правила → плани «🔁 план регулярний» у журналі; факт у межах ±5 днів і ±25 % закриває місяць; прострочені; вимкнення.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
const K = (c, x) => require('vm').runInContext(x, c);
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const tz = 'Europe/Kiev';

function setup() {
  const env = makeEnv();
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  return env;
}
const rule = (o) => Object.assign({ i: 0, name: 'Оренда', cat: 'Оренда житла', amt: 10000, day: 5, per: 'щомісяця', from: '', to: '', who: '', on: true }, o);
const row = (o) => Object.assign({ row: 100, key: '', d: '', e: '✅ факт', b: 0, c: 'Оренда житла', g: '', blank: false }, o);

t('payment date by period: monthly, quarterly and yearly from the month of «З», day 31 → last day, «з» / «до» respected', () => {
  const c = setup().ctx;
  assert.strictEqual(c.regDate_(rule(), '2026-12', '2026-11-26'), '2026-12-05');
  assert.strictEqual(c.regDate_(rule({ day: 31 }), '2027-02', '2026-11-26'), '2027-02-28');
  const q = rule({ per: 'щокварталу', from: '2026-03-10' });
  assert.strictEqual(c.regDate_(q, '2026-12', '2026-11-26'), '2026-12-05');
  assert.strictEqual(c.regDate_(q, '2026-11', '2026-11-26'), '', 'не квартальний місяць');
  assert.strictEqual(c.regDate_(rule({ per: 'щороку', from: '2026-01-20' }), '2027-01', '2026-11-26'), '2027-01-05');
  assert.strictEqual(c.regDate_(rule({ per: 'щороку', from: '2026-01-20' }), '2026-12', '2026-11-26'), '');
  assert.strictEqual(c.regDate_(rule({ per: 'щокварталу' }), '2027-02', '2026-11-26'), '2027-02-05', 'без «З»: від поточного місяця');
  assert.strictEqual(c.regDate_(rule({ from: '2027-01-01' }), '2026-12', '2026-11-26'), '', 'ще не почалось');
  assert.strictEqual(c.regDate_(rule({ to: '2026-12-01' }), '2026-12', '2026-11-26'), '', 'уже скінчилось');
});

t('wanted plans: from today to the horizon; a fact of the same category within ±5 days and ±25 % (or with the name) closes the month; one fact closes one rule', () => {
  const c = setup().ctx;
  let w = c.regWanted_([rule()], [], '2026-11-26', '2026-12-31');
  assert.strictEqual(w.want.length, 1, 'листопадовий день минув: заднім числом план не ставимо');
  assert.strictEqual(w.want[0].key, '[Регулярний] Оренда: 12.2026'); assert.strictEqual(w.want[0].date, '2026-12-05'); assert.strictEqual(w.want[0].amount, 10000);
  assert.strictEqual(w.state[0].next, '2026-12-05');
  // заплатили раніше: 03.12 при плані 05.12
  w = c.regWanted_([rule()], [row({ key: '2026-12-03', b: 9500 })], '2026-12-04', '2027-01-31');
  assert.strictEqual(JSON.stringify(w.want.map(x => x.date)), JSON.stringify(['2027-01-05']));
  assert.ok(w.paid.indexOf('[Регулярний] Оренда: 12.2026') >= 0);
  // сума інша на 40 % і без назви: не те
  w = c.regWanted_([rule()], [row({ key: '2026-12-03', b: 6000 })], '2026-12-04', '2026-12-31');
  assert.strictEqual(w.want.length, 1);
  // сума інша, але в коментарі назва правила: те
  w = c.regWanted_([rule()], [row({ key: '2026-12-03', b: 6000, d: 'Оренда за грудень' })], '2026-12-04', '2026-12-31');
  assert.strictEqual(w.want.length, 0);
  // факт з коментарем плану (план відмітили «✅ факт») закриває місяць незалежно від дати
  w = c.regWanted_([rule()], [row({ key: '2026-12-20', b: 10000, d: '[Регулярний] Оренда: 12.2026' })], '2026-12-04', '2026-12-31');
  assert.strictEqual(w.want.length, 0);
  // два правила тієї ж категорії: один факт закриває лише одне
  w = c.regWanted_([rule(), rule({ i: 1, name: 'Гараж', amt: 9000 })], [row({ key: '2026-12-05', b: 9800 })], '2026-12-04', '2026-12-31');
  assert.strictEqual(w.want.length, 1);
  // план чи чернетка з банку: чернетка теж закриває, план ні
  assert.strictEqual(c.regWanted_([rule()], [row({ key: '2026-12-05', b: 10000, e: '🏦 чернетка' })], '2026-12-04', '2026-12-31').want.length, 0);
  assert.strictEqual(c.regWanted_([rule()], [row({ key: '2026-12-05', b: 10000, e: 'план' })], '2026-12-04', '2026-12-31').want.length, 1);
  // вимкнено: планів немає
  assert.strictEqual(c.regWanted_([rule({ on: false })], [], '2026-11-26', '2026-12-31').want.length, 0);
});

t('overdue: a past plan without a fact stays and is reported, a rescheduled one keeps your date; plans of a removed or disabled rule are dropped even in the past', () => {
  const c = setup().ctx;
  const old = row({ key: '2026-10-05', b: 10000, d: '[Регулярний] Оренда: 10.2026', e: '🔁 план регулярний' });
  let w = c.regWanted_([rule()], [old], '2026-11-26', '2026-12-31');
  assert.strictEqual(w.state[0].over, '2026-10-05');
  assert.ok(w.paid.indexOf(old.d) < 0, 'прострочений план лишається');
  w = c.regWanted_([rule()], [old, row({ key: '2026-10-08', b: 10000 })], '2026-11-26', '2026-12-31');
  assert.ok(w.paid.indexOf(old.d) >= 0, 'факт знайшовся: план прибрати');
  assert.ok(c.regWanted_([rule({ on: false })], [old], '2026-11-26', '2026-12-31').paid.indexOf(old.d) >= 0, 'правило вимкнено');
  assert.ok(c.regWanted_([], [old], '2026-11-26', '2026-12-31').paid.indexOf(old.d) >= 0, 'правила немає');
  // прострочений жовтневий платіж ви перенесли на 30.11: план лишається з вашою датою й сумою
  w = c.regWanted_([rule()], [row({ key: '2026-11-30', b: 9000, d: '[Регулярний] Оренда: 10.2026', e: '🔁 план регулярний' })], '2026-11-26', '2026-12-31');
  const moved = w.want.find(x => x.key === '[Регулярний] Оренда: 10.2026');
  assert.ok(moved && moved.date === '2026-11-30' && moved.amount === 9000, JSON.stringify(w.want));
});

t('sheet and sync: lists, plans in the journal with the tag, updated in place, past plan closed by a fact, state column; your rows untouched', () => {
  const env = setup(), c = env.ctx, r = env.ss.getSheetByName('Регулярні'), j = env.ss.getSheetByName('Журнал');
  assert.strictEqual(r.getRange('A4').getValue(), 'Назва');
  assert.ok(r.getRange('E5').getDataValidation().reject && r.getRange('I5').getDataValidation().reject, 'списки «як часто» і «так / ні»');
  const today = env.Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), prev = c.prevYm_(today.slice(0, 7));
  const P = k => env.Utilities.parseDate(k, tz, 'yyyy-MM-dd');
  r.getRange('A5:I5').setValues([['Оренда', 'Оренда житла', 10000, 28, 'щомісяця', '', '', '', 'так']]);
  j.getRange('A5:E6').setValues([[P(today), 99, 'Продукти', 'мій рядок', '✅ факт'],
    [P(prev + '-10'), 10000, 'Оренда житла', '[Регулярний] Оренда: ' + prev.slice(5) + '.' + prev.slice(0, 4), '🔁 план регулярний']]);
  const s1 = c.regPlansSync_(env.ss);
  const plans = () => c.readJournal_(env.ss).filter(o => o.d.indexOf('[Регулярний] ') === 0);
  assert.ok(s1.added >= 1, JSON.stringify(s1));
  assert.ok(plans().every(o => o.e === '🔁 план регулярний' && o.c === 'Оренда житла'));
  assert.ok(plans().filter(o => o.key >= today).every(o => o.b === 10000 && /^\[Регулярний\] Оренда: \d\d\.\d{4}$/.test(o.d)));
  assert.ok(plans().some(o => o.key < today), 'минулий план без факту лишився');
  assert.ok(/⚠️ прострочено 10\./.test(r.getRange('J5').getValue()) && /наступний/.test(r.getRange('J5').getValue()), r.getRange('J5').getValue());
  // нова сума: оновлено на місці
  r.getRange('C5').setValue(11000);
  const s2 = c.regPlansSync_(env.ss);
  assert.ok(s2.updated >= 1 && s2.added === 0, JSON.stringify(s2));
  assert.ok(plans().filter(o => o.key >= today).every(o => o.b === 11000));
  // факт минулого місяця: прострочений план прибрано, факт лишився
  const free = c.readJournal_(env.ss).find(o => o.blank).row;
  j.getRange(free, 1, 1, 5).setValues([[P(prev + '-12'), 10000, 'Оренда житла', 'переказ орендодавцю', '✅ факт']]);
  c.regPlansSync_(env.ss);
  assert.ok(!plans().some(o => o.key < today), 'план сплаченого місяця прибрано');
  assert.ok(c.readJournal_(env.ss).some(o => o.d === 'переказ орендодавцю') && c.readJournal_(env.ss).some(o => o.d === 'мій рядок'));
  assert.ok(!/прострочено/.test(r.getRange('J5').getValue()));
  // вимкнули: майбутні плани геть
  r.getRange('I5').setValue('ні');
  c.regPlansSync_(env.ss);
  assert.strictEqual(plans().length, 0);
  assert.strictEqual(r.getRange('J5').getValue(), 'вимкнено');
});

t('schedule: once a day or after an edit on «Регулярні» (onEdit marks it, mbsEdit asks for a soon run); menu item exists', () => {
  const env = setup(), c = env.ctx, r = env.ss.getSheetByName('Регулярні');
  r.getRange('A5:I5').setValues([['Підписка', 'Підписки', 299, 28, 'щомісяця', '', '', '', 'так']]);
  env.docProps.deleteProperty('reg_day');                                   // наступний день (збирання вже раз оновило «Стан» прикладів)
  assert.ok(c.regIfDue_(env.ss), 'перший раз за день');
  assert.strictEqual(c.regIfDue_(env.ss), null, 'вдруге того ж дня: ні');
  env.docProps.setProperty('reg_t', String(Date.now() - 10000));
  c.onEdit({ range: r.getRange('C5') });
  assert.ok(Number(env.docProps.getProperty('reg_edit')) > 0);
  assert.ok(c.regIfDue_(env.ss), 'після правки: так');
  const before = env.triggers.length;
  env.docProps.deleteProperty('sort_soon');
  c.mbsEdit({ range: r.getRange('C5') });
  assert.strictEqual(env.triggers.length, before + 1, 'скоро mbsTick');
  env.calls.menu = null; c.onOpen();
  assert.ok(env.calls.menu.indexOf('regularNow') >= 0);
});

t('examples: 21 kinds (subscriptions, rent, utilities, fuel, sport, insurance, salary, loan, installment, debts, savings) with «ні», no plans, a hint in «Стан»; yours count, examples do not; added once; a step adds them to an older table after your rows', () => {
  const env = setup(), c = env.ctx, r = env.ss.getSheetByName('Регулярні'), n = K(c, 'REG_EXAMPLES').length;
  const rows = r.getRange('A5:I54').getValues().filter(x => x[0]);
  assert.strictEqual(rows.length, n);
  assert.ok(rows.every(x => x[8] === 'ні'), 'усі вимкнені');
  ['Netflix', 'Бензин', 'Оренда квартири', 'Страховка авто', 'Зарплата', 'Кредит у банку', 'Розстрочка за телефон', 'Повертаю борг другу', 'Мені повертають борг']
    .forEach(name => assert.ok(rows.some(x => x[0] === name), name));
  assert.ok(rows.some(x => x[4] === 'щокварталу') && rows.some(x => x[4] === 'щороку') && rows.some(x => x[6] !== ''), 'щокварталу, щороку, з датою «До»');
  const noteA4 = r.getRange('A4').getNotes()[0][0];
  assert.ok(/приклади/.test(noteA4) && /«так»/.test(noteA4), 'підказка про приклади: у примітці на «Назва» (рядки 1–3: шапка)');
  assert.ok(r.getRange('J5:J' + (4 + n)).getValues().every(x => /^💡 приклад, «ні»: /.test(x[0])), 'підказка в «Стані»');
  assert.ok(/Борги/.test(r.getRange('J18').getValue()), 'кредит: на «Борги»');
  assert.strictEqual(c.readJournal_(env.ss).filter(o => o.d.indexOf('[Регулярний] ') === 0).length, 0, 'прикладів у журналі немає');
  assert.strictEqual(c.regOwn_(env.ss).length, 0, 'приклади не рахуються як ваші');
  assert.ok(/Як додати/.test(env.ss.getSheetByName('Борги').getRange('A3').getValue()));
  // ваш: приклад з «так» стає планом
  r.getRange('I5').setValue('так');
  c.regPlansSync_(env.ss);
  assert.strictEqual(c.regOwn_(env.ss).length, 1);
  assert.ok(c.readJournal_(env.ss).some(o => o.d.indexOf('[Регулярний] Netflix: ') === 0));
  assert.ok(!/приклад/.test(r.getRange('J5').getValue()));
  // ще раз не додаються
  assert.strictEqual(c.regExamplesAdd_(env.ss), 0);
  // стара таблиця: ваші рядки зверху, приклади після них
  const env2 = setup(), r2 = env2.ss.getSheetByName('Регулярні');
  r2.getRange('A5:J54').clearContent(); r2.getRange('A5:A54').setNotes(Array.from({ length: 50 }, () => ['']));
  r2.getRange('A5:I6').setValues([['Моя оренда', 'Оренда житла', 9000, 1, 'щомісяця', '', '', '', 'так'], ['Моя підписка', 'Підписки', 199, 9, 'щомісяця', '', '', '', 'так']]);
  env2.docProps.deleteProperty('reg_ex');
  env2.docProps.setProperty('schema_v', '9');
  env2.ctx.schemaTick_(env2.ss);
  assert.strictEqual(r2.getRange('A5').getValue(), 'Моя оренда');
  assert.strictEqual(r2.getRange('A7').getValue(), 'Netflix');
  assert.strictEqual(env2.ctx.regOwn_(env2.ss).length, 2);
});

t('📊 header (rows 1–3): dark tiles like «Charts»; first tile «💰 Собі й інвестиції»; then connected, left over, this month, 30 days with a column chart; «Борги» tile (K:L) only with the module; helper columns N…U hidden; old title and hints gone', () => {
  const env = setup(), c = env.ctx, r = env.ss.getSheetByName('Регулярні');
  const f = a => r.getRange(a).getFormula(), v = a => String(r.getRange(a).getValue());
  ['A1', 'A2', 'A3'].forEach(a => assert.ok(!/Регулярні платежі й надходження|Кожен рядок стає планом|Нижче приклади/.test(v(a)), a + ': ' + v(a)));
  assert.ok(/СОБІ Й ІНВЕСТИЦІЇ/.test(f('A1')) && /SPARKLINE\(\$S\$2:\$S\$3,\$T\$1:\$U\$4\)/.test(f('A3')), f('A1') + '\n' + f('A3'));
  assert.strictEqual(f('A2').indexOf('=TEXT($S$4'), 0);
  assert.ok(/ПІДКЛЮЧЕНО/.test(v('C1')) && f('C2') === '=$S$10');
  assert.ok(/ЛИШАЄТЬСЯ/.test(f('E1')) && f('E2') === '=$S$8' && /SPARKLINE\(\$S\$26:\$S\$27/.test(f('E3')));
  assert.ok(/ЦЕЙ МІСЯЦЬ/.test(f('G1')) && /ПРОСТРОЧЕНО/.test(f('G1')) && /SPARKLINE\(\$S\$14:\$S\$15/.test(f('G3')));
  assert.ok(/30 ДНІВ/.test(f('I1')) && /SPARKLINE\(\$Q\$5:\$Q\$34,\$T\$16:\$U\$18\)/.test(f('I3')));
  assert.strictEqual(v('K1'), '', 'модуль боргів вимкнено: плитки «Борги» немає');
  assert.ok(['A1:B1', 'A2:B2', 'A3:B3', 'I1:J1'].every(m => (r.merges || []).indexOf(m) >= 0), JSON.stringify(r.merges));
  // службові: на місяць (щокварталу /3, щороку /12, лише «так», не після «До»), тип з «Довідника», дні, підсумки
  assert.ok(/IF\(\$E5="щокварталу",1\/3,IF\(\$E5="щороку",1\/12,1\)\)/.test(f('N5')) && /\$I5<>"так"/.test(f('N5')) && /N\(\$G5\)<TODAY\(\)/.test(f('N5')), f('N5'));
  assert.ok(/VLOOKUP\(\$B5,'Довідник'!\$A\$5:\$B\$/.test(f('O5')));
  assert.ok(/"\[Регулярний\] \*"/.test(f('Q5')) && /"\[Борг\] \*"/.test(f('Q5')) && /TODAY\(\)\+29/.test(f('Q34')));
  assert.ok(/"Заплати собі"/.test(f('S2')) && /"Інвестиції"/.test(f('S3')) && /FILTER\(/.test(f('S19')));
  assert.ok(r.isColumnHiddenByUser(14) && r.isColumnHiddenByUser(21) && !r.isColumnHiddenByUser(10));
  assert.ok(/Кожен рядок стає планом/.test(r.getRange('A4').getNotes()[0][0]));
  // модуль боргів: плитка K:L і підсумки з «Боргів»; вимкнули: знову порожньо
  c.modSet_(env.ss, 'debts', true); c.modulesApply_(env.ss, true);
  assert.ok(/БОРГИ/.test(v('K1')) && f('K2') === '=$S$6' && /'Борги'!\$D\$5:\$D\$24/.test(f('S6')), f('S6'));
  c.modSet_(env.ss, 'debts', false); c.modulesApply_(env.ss, true);
  assert.strictEqual(v('K1'), ''); assert.strictEqual(f('S6'), '');
  // повторно: та сама шапка, без подвоєння
  c.regHeader_(env.ss);
  assert.strictEqual((r.merges || []).filter(m => m === 'A1:B1').length, 1);
  const all = r.getRange('A1:U4').getValues().map(x => x.join(' ')).join(' ') + r.getRange('A1:U4').getFormulas().map(x => x.join(' ')).join(' ');
  assert.ok(all.indexOf('—') < 0, 'без довгих тире');
});

t('📊 categories «Заплати собі» and «Інвестиції» (type «Заощадження») and two examples; upgrade of a version-17 table: old rows 1–3 replaced by the header, categories and examples added once after your rows', () => {
  const env = setup(), c = env.ctx, d = env.ss.getSheetByName('Довідник');
  const cats = () => d.getRange('A5:C84').getValues().filter(x => x[0]);
  ['Заплати собі', 'Інвестиції'].forEach(n => assert.strictEqual(JSON.stringify(cats().find(x => x[0] === n).slice(1)), JSON.stringify(['Заощадження', '💰 Заощадження']), n));
  const r = env.ss.getSheetByName('Регулярні');
  assert.ok(r.getRange('A5:B54').getValues().some(x => x[0] === 'Заплати собі' && x[1] === 'Заплати собі'));
  // стара таблиця: без категорій, без нових прикладів, старі рядки 1–3
  const env2 = setup(), c2 = env2.ctx, d2 = env2.ss.getSheetByName('Довідник'), r2 = env2.ss.getSheetByName('Регулярні');
  d2.getRange('A5:A84').getValues().forEach((x, i) => { if (x[0] === 'Заплати собі' || x[0] === 'Інвестиції') d2.getRange(5 + i, 1, 1, 3).setValues([['', '', '']]); });
  r2.getRange('A5:J54').getValues().forEach((x, i) => { if (x[0] === 'Заплати собі' || x[0] === 'Інвестиції') r2.getRange(5 + i, 1, 1, 10).setValues([Array(10).fill('')]); });
  r2.getRange('A1:A3').setValues([['🔁 Регулярні платежі й надходження'], ['Кожен рядок стає планом…'], ['Нижче приклади з «Увімкнено: ні»…']]);
  r2.getRange('S10').setValue(''); r2.getRange('N5').setFormula('');
  env2.docProps.deleteProperty('reg_ex2');
  env2.docProps.setProperty('schema_v', '17');
  assert.strictEqual(K(c2, 'STEPS').find(x => x.v === 18).done(env2.ss), false);
  c2.schemaTick_(env2.ss);
  assert.strictEqual(env2.docProps.getProperty('schema_v'), String(c2.schemaLast_()));
  assert.ok(d2.getRange('A5:A84').getValues().some(x => x[0] === 'Заплати собі') && d2.getRange('A5:A84').getValues().some(x => x[0] === 'Інвестиції'));
  const names = r2.getRange('A5:A54').getValues().map(x => x[0]);
  assert.ok(names.indexOf('Заплати собі') >= 0 && names.indexOf('Інвестиції') >= 0 && names.indexOf('Netflix') === 0, names.join(','));
  assert.ok(/СОБІ Й ІНВЕСТИЦІЇ/.test(r2.getRange('A1').getFormula()) && String(r2.getRange('A2').getValue()).indexOf('Кожен рядок') < 0);
  assert.strictEqual(K(c2, 'STEPS').find(x => x.v === 18).done(env2.ss), true);
  assert.strictEqual(c2.regExamplesMore_(env2.ss), 0, 'вдруге не додаються');
});

console.log(n + ' tests passed');
