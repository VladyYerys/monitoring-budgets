// «Регулярні»: правила → плани «🔁 план регулярний» у журналі; факт у межах ±5 днів і ±25 % закриває місяць; прострочені; вимкнення.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
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

console.log(n + ' tests passed');
