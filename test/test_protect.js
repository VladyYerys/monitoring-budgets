// Журнал без стирання (0.5.9, версія таблиці 22): стерте записане повертається з копії, видалені рядки вставляються назад;
// «🧹 Зачистити дані…»: усе, дата, період чи місяць, з копією таблиці; після неї імпорт з банку завантажує ці дні знову.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);
const TOKEN = 'uTestToken' + 'z'.repeat(30);

function setup(opts) {
  const state = { items: [], accounts: [{ id: 'acc1', currencyCode: 980, type: 'black', creditLimit: 0, balance: 4000000, iban: 'UA00TEST', maskedPan: ['000000******4321'] }] };
  const env = makeEnv(Object.assign({ mono: url => {
    if (/client-info$/.test(url)) return { code: 200, body: { name: 'Тест', accounts: state.accounts } };
    if (/\/personal\/statement\//.test(url)) return { code: 200, body: state.items };
    return { code: 404, body: 'no' };
  }, promptText: TOKEN }, opts || {}));
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  env.state = state;
  return env;
}
const day = (env, d) => env.Utilities.formatDate(new Date(Date.now() + (d || 0) * 86400000), 'Europe/Kiev', 'yyyy-MM-dd');
const sec = (env, key, hhmm) => Math.floor(env.Utilities.parseDate(key + ' ' + (hhmm || '12:00') + ':00', 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss').getTime() / 1000);
const dmy = k => k.slice(8, 10) + '.' + k.slice(5, 7) + '.' + k.slice(0, 4);
const J = env => env.ss.getSheetByName('Журнал');

/** Mono: 3 операції за 2 дні тому й 2 за вчора, плюс ваш рядок готівкою вчора. */
function filled(opts) {
  const env = setup(opts), c = env.ctx, k2 = day(env, -2), k1 = day(env, -1);
  const it = (id, k, desc, mcc, amount, hhmm) => ({ id: id, time: sec(env, k, hhmm), description: desc, mcc: mcc, amount: amount, balance: 4000000 });
  env.state.items = [it('a1', k2, 'АТБ', 5411, -15000, '09:00'), it('a2', k2, 'Glovo', 5812, -32000, '13:00'), it('a3', k2, 'Банкомат DN00TEST', 6011, -50000, '18:00'),
    it('b1', k1, 'Uklon', 4121, -12000, '08:00'), it('b2', k1, 'MEBLIMAX', 5712, -84000, '19:00')];
  c.connectMono();
  c.importBank_(env.ss);
  const j = J(env), free = c.readJournal_(env.ss).find(o => o.blank).row;
  j.getRange(free, 1, 1, 5).setValues([[env.Utilities.parseDate(k1, 'Europe/Kiev', 'yyyy-MM-dd'), 85, 'Кава та перекуси', 'кава готівкою', '✅ факт']]);
  c.sortCore_(env.ss);
  return { env: env, c: c, k1: k1, k2: k2 };
}
const rowOf = (env, id) => env.ctx.readJournal_(env.ss).find(o => o.bank && o.bank[0] === id);
const cash = env => env.ctx.readJournal_(env.ss).find(o => o.d === 'кава готівкою');
const logIds = env => { const b = env.ss.getSheetByName('Банк'), last = b.getLastRow(); return last < 5 ? [] : b.getRange(5, 2, last - 4, 1).getValues().map(r => String(r[0])).filter(Boolean); };
const edit = (env, a1, extra) => env.ctx.onEdit(Object.assign({ range: J(env).getRange(a1) }, extra || {}));

t('pure: «що зачистити»: усе, дата, період (обидва напрями, тире, «з … по»), місяць, рррр-мм-дд; нісенітниця й неіснуючі дати: null', () => {
  const c = setup().ctx, w = s => JSON.stringify(c.wipeRange_(s));
  assert.strictEqual(w('усе'), JSON.stringify({ all: true, from: '', to: '' }));
  assert.strictEqual(w(' Все '), JSON.stringify({ all: true, from: '', to: '' }));
  assert.strictEqual(w('14.09.2026'), JSON.stringify({ all: false, from: '2026-09-14', to: '2026-09-14' }));
  assert.strictEqual(w('01.09.2026-30.09.2026'), JSON.stringify({ all: false, from: '2026-09-01', to: '2026-09-30' }));
  assert.strictEqual(w('30.09.2026 – 01.09.2026'), JSON.stringify({ all: false, from: '2026-09-01', to: '2026-09-30' }));
  assert.strictEqual(w('з 1.9.2026 по 5.9.2026'), JSON.stringify({ all: false, from: '2026-09-01', to: '2026-09-05' }));
  assert.strictEqual(w('09.2026'), JSON.stringify({ all: false, from: '2026-09-01', to: '2026-09-30' }));
  assert.strictEqual(w('02.2028'), JSON.stringify({ all: false, from: '2028-02-01', to: '2028-02-29' }));
  assert.strictEqual(w('2026-09-14'), JSON.stringify({ all: false, from: '2026-09-14', to: '2026-09-14' }));
  ['', 'вчора', '31.02.2026', '13.2026', '14.09.2026 і ще щось', '1.1.2026-2.1.2026-3.1.2026'].forEach(s => assert.strictEqual(c.wipeRange_(s), null, s));
  assert.strictEqual(c.wipeMonths_('2026-09-14', '2026-11-28'), 3);
  assert.strictEqual(c.wipeMonths_('2026-11-27', '2026-11-28'), 1);
  assert.strictEqual(JSON.stringify(c.rowRuns_([9, 5, 6, 7, 12])), JSON.stringify([[5, 3], [9, 1], [12, 1]]));
});

t('pure: копія рядка з датою туди й назад; де видалено рядки (і коли копія не збігається)', () => {
  const c = setup().ctx, d = new Date(Date.UTC(2026, 8, 14));
  const s = c.shadowEnc_([d, 100, 'Продукти', '', '✅ факт', '', '', '', 'id1', 'АТБ', -100, 5411, '', '', '', '']);
  const back = c.shadowDec_(s);
  assert.strictEqual(back[0].getTime(), d.getTime()); assert.strictEqual(back[1], 100); assert.strictEqual(back[8], 'id1');
  assert.strictEqual(c.shadowEnc_(['', '', null, undefined]), '');
  assert.strictEqual(c.rowsGone_(['a', 'b', 'c', 'd', ''], ['a', 'c', 'd', ''], 1), 1);
  assert.strictEqual(c.rowsGone_(['a', 'b', 'c', 'd', ''], ['a', 'b', 'c', 'd'], 1), 4, 'в кінці');
  assert.strictEqual(c.rowsGone_(['a', 'b', 'c', 'd', 'e'], ['a', 'd', 'e'], 2), 1);
  assert.strictEqual(c.rowsGone_(['a', 'b', 'c', 'd'], ['a', 'x', 'd'], 1), -1, 'копія не та');
});

t('стерли дату, суму, категорію чи статус або весь рядок: повертається одразу, з підказкою; коментар і щойно вписане стирати можна', () => {
  const { env, c } = filled(), j = J(env);
  const r = rowOf(env, 'a2').row, before = j.getRange(r, 1, 1, 8).getValues()[0];
  j.getRange(r, 3).setValue(''); edit(env, 'C' + r);
  assert.strictEqual(j.getRange(r, 3).getValue(), before[2], 'категорія повернулась');
  assert.ok(env.calls.toasts.some(x => /стирати не можна/.test(x[1]) && /Зачистити дані/.test(x[1])), JSON.stringify(env.calls.toasts));
  j.getRange(r, 1, 1, 8).setValues([['', '', '', '', '', '', '', '']]); edit(env, 'A' + r + ':H' + r);
  assert.strictEqual(JSON.stringify(j.getRange(r, 1, 1, 8).getValues()[0]), JSON.stringify(before), 'весь рядок повернувся');
  // ваш рядок готівкою: так само
  const m = cash(env).row, mb = j.getRange(m, 1, 1, 8).getValues()[0];
  j.getRange(m, 1, 1, 5).setValues([['', '', '', '', '']]); edit(env, 'A' + m + ':E' + m);
  assert.strictEqual(JSON.stringify(j.getRange(m, 1, 1, 8).getValues()[0]), JSON.stringify(mb));
  // коментар стерти можна
  j.getRange(m, 4).setValue(''); edit(env, 'D' + m);
  assert.strictEqual(j.getRange(m, 4).getValue(), '');
  // щойно вписане (ще не в копії): стерти можна
  const free = c.readJournal_(env.ss).find(o => o.blank).row;
  j.getRange(free, 2, 1, 2).setValues([[40, 'Таксі']]); edit(env, 'B' + free + ':C' + free);
  j.getRange(free, 1, 1, 3).setValues([['', '', '']]); edit(env, 'A' + free + ':C' + free);
  assert.strictEqual(j.getRange(free, 2).getValue(), '');
  // вписали значення в одну клітинку: нічого не читаємо й не повертаємо
  j.getRange(r, 2).setValue(321); edit(env, 'B' + r, { value: '321' });
  assert.strictEqual(j.getRange(r, 2).getValue(), 321);
});

t('захист не заважає: план можна стерти, недописаний рядок теж; рядок, що переїхав, не дублюється; стерті формули праворуч пишуться знову', () => {
  const { env, c } = filled(), j = J(env);
  // недописаний ваш рядок (без категорії), уже в копії
  const free = c.readJournal_(env.ss).find(o => o.blank).row;
  j.getRange(free, 1, 1, 2).setValues([[new Date(), 55]]); c.sortCore_(env.ss);
  const inc = c.readJournal_(env.ss).find(o => Number(o.b) === 55).row;
  j.getRange(inc, 1, 1, 2).setValues([['', '']]); edit(env, 'A' + inc + ':B' + inc);
  assert.strictEqual(j.getRange(inc, 2).getValue(), '', 'недописаний: стерти можна');
  // план
  const f2 = c.readJournal_(env.ss).find(o => o.blank).row;
  j.getRange(f2, 1, 1, 5).setValues([[new Date(Date.now() + 5 * 86400000), 700, 'Комуналка', 'на газ', 'план']]); c.sortCore_(env.ss);
  const p = c.readJournal_(env.ss).find(o => Number(o.b) === 700).row;
  j.getRange(p, 1, 1, 5).setValues([['', '', '', '', '']]); edit(env, 'A' + p + ':E' + p);
  assert.strictEqual(j.getRange(p, 2).getValue(), '', 'план: стерти можна');
  // переїхав: вирізали рядок готівки й вставили в порожній
  const m = cash(env).row, mv = j.getRange(m, 1, 1, 8).getValues()[0], to = c.readJournal_(env.ss).find(o => o.blank).row;
  j.getRange(to, 1, 1, 8).setValues([mv]); j.getRange(m, 1, 1, 8).setValues([['', '', '', '', '', '', '', '']]);
  edit(env, 'A' + m + ':H' + m);
  assert.strictEqual(j.getRange(m, 2).getValue(), '', 'не дублюємо');
  assert.strictEqual(c.readJournal_(env.ss).filter(o => o.d === 'кава готівкою').length, 1);
  // формула в I стерта (після попередження Google)
  const r = rowOf(env, 'b1').row, fI = j.getRange(r, 9).getFormula();
  assert.ok(fI);
  j.getRange(r, 9).setValue(''); edit(env, 'I' + r);
  assert.strictEqual(j.getRange(r, 9).getFormula(), fI);
});

t('«виділити все» й Delete: дані повертаються з копії, шапка й заголовки колонок пишуться знову', () => {
  const { env, c } = filled(), j = J(env), r = rowOf(env, 'a1').row, v = j.getRange(r, 1, 1, 8).getValues()[0];
  const a1 = j.getRange('A1').getValue(), b3 = j.getRange('B3').getFormula();
  assert.ok(a1 && b3);
  j.getRange(1, 1, 60, 8).setValues(new Array(60).fill(0).map(() => new Array(8).fill('')));
  edit(env, 'A1:H60');
  assert.strictEqual(j.getRange('A1').getValue(), a1); assert.strictEqual(j.getRange('B3').getFormula(), b3);
  assert.strictEqual(j.getRange('A4').getValue(), 'Дата'); assert.strictEqual(j.getRange('C4').getValue(), 'Категорія');
  assert.strictEqual(JSON.stringify(j.getRange(r, 1, 1, 8).getValues()[0]), JSON.stringify(v));
  assert.strictEqual(c.readJournal_(env.ss).filter(o => o.bank && o.bank[0]).length, 5);
  assert.ok(cash(env), 'і ваш рядок');
});

t('видалили рядки журналу: вставляються назад на те саме місце, зі значеннями, банком і формулами; видалення на іншому аркуші й шапки: нічого не вгадуємо', () => {
  const { env, c } = filled(), j = J(env), y = 25;
  const r = rowOf(env, 'b2').row, vals = j.getRange(r, 1, 1, 8).getValues()[0], max = j.getMaxRows(), nextId = j.getRange(r + 1, y).getValue();
  j.deleteRows(r, 2);
  c.mbsChange({ changeType: 'REMOVE_ROW' });
  assert.strictEqual(j.getMaxRows(), max);
  assert.strictEqual(JSON.stringify(j.getRange(r, 1, 1, 8).getValues()[0]), JSON.stringify(vals));
  assert.strictEqual(j.getRange(r, y).getValue(), 'b2'); assert.strictEqual(j.getRange(r + 1, y).getValue(), nextId);
  assert.ok(j.getRange(r, 9).getFormula() && j.getRange(r + 1, 13).getFormula(), 'формули на місці');
  assert.ok(env.calls.toasts.some(x => /видаляти не можна: повернув 2 рядки/.test(x[1])), JSON.stringify(env.calls.toasts));
  assert.strictEqual(c.readJournal_(env.ss).filter(o => o.bank && o.bank[0]).length, 5);
  // інший аркуш: журнал не змінився
  const d = env.ss.getSheetByName('Довідник'), dm = d.getMaxRows();
  d.deleteRows(90, 1); c.mbsChange({ changeType: 'REMOVE_ROW' });
  assert.strictEqual(j.getMaxRows(), max); assert.strictEqual(d.getMaxRows(), dm - 1);
  // шапку: не вгадуємо, підказка Ctrl+Z
  env.calls.toasts.length = 0;
  j.deleteRows(4, 1); c.mbsChange({ changeType: 'REMOVE_ROW' });
  assert.strictEqual(j.getMaxRows(), max - 1);
  assert.ok(env.calls.toasts.some(x => /шапки/.test(x[1]) && /Ctrl\+Z/.test(x[1])), JSON.stringify(env.calls.toasts));
});

t('🧹 зачистити дату: лише ці рядки (з банку й ваші) і записи «Банк» за цей день, з копією таблиці; без імпорту одразу підказка; імпорт повертає їх', () => {
  const { env, c, k1, k2 } = filled({ alertAnswer: (t, m) => (/Завантажити з Mono/.test(m) ? 'NO' : 'YES') });
  env.opts.promptText = [dmy(k1)];
  const files0 = env.files.length;
  c.wipeNow();
  const conf = env.calls.alerts.find(a => /Зачистити «/.test(a[1]));
  assert.ok(/рядків журналу: 3 \(з банку 2, ваших 1, планів 0\)/.test(conf[1]) && /записів на аркуші «Банк»: 2/.test(conf[1]), conf[1]);
  assert.strictEqual(env.files.length, files0 + 1, 'копія перед змінами');
  assert.strictEqual(rowOf(env, 'b1'), undefined); assert.strictEqual(rowOf(env, 'b2'), undefined); assert.strictEqual(cash(env), undefined);
  assert.ok(rowOf(env, 'a1') && rowOf(env, 'a2') && rowOf(env, 'a3'), 'інший день не чіпали');
  assert.deepStrictEqual(logIds(env).sort(), ['a1', 'a2', 'a3']);
  const dlg = env.calls.dialogs.pop();
  assert.ok(/Зачищено/.test(dlg.html) && /Копія перед змінами/.test(dlg.html) && /Імпорт з банку за період/.test(dlg.html), dlg.html);
  c.importBank_(env.ss);
  assert.ok(rowOf(env, 'b1') && rowOf(env, 'b2'), 'імпорт завантажив знову');
  assert.strictEqual(c.readJournal_(env.ss).filter(o => o.bank && o.bank[0]).length, 5, 'без дублів');
  // зачищене не повертається захистом: копія оновилась
  assert.ok(c.readJournal_(env.ss).every(o => o.d !== 'кава готівкою'));
});

t('🧹 зачистити «усе» й одразу завантажити з Mono знову; «ні» в підтвердженні: нічого не змінено; під час імпорту за період: спершу зупинити', () => {
  const { env, c, k2 } = filled();
  env.opts.alertAnswer = 'NO'; env.opts.promptText = ['усе'];
  c.wipeNow();
  assert.strictEqual(c.readJournal_(env.ss).filter(o => o.bank && o.bank[0]).length, 5, '«ні»: нічого');
  env.opts.alertAnswer = 'YES'; env.opts.promptText = ['усе'];
  c.wipeNow();
  const dlg = env.calls.dialogs.pop();
  assert.ok(/Зачищено «усе»/.test(dlg.html) && /Рядків журналу: <b>6<\/b>/.test(dlg.html), dlg.html);
  assert.ok(/Картка: <b>/.test(dlg.html) && new RegExp('з <b>' + dmy(k2) + '</b>').test(dlg.html), 'імпорт за період з найранішої зачищеної дати: ' + dlg.html);
  assert.strictEqual(c.readJournal_(env.ss).filter(o => o.bank && o.bank[0]).length, 5, 'з банку знову всі 5');
  assert.strictEqual(cash(env), undefined, 'ваш рядок готівкою зачищено (у копії таблиці він є)');
  env.userProps.setProperty('period_job', JSON.stringify({ name: 'x', fromKey: k2, done: 0, total: 2, rep: {} }));
  env.calls.prompts.length = 0;
  c.wipeNow();
  assert.strictEqual(env.calls.prompts.length, 0); assert.ok(/спершу зупиніть/i.test(env.calls.dialogs.pop().html));
});

t('🧹 не зрозумів чи порожньо: пояснення; меню має пункт; аркуш копії схований, а копія оновлюється при сортуванні', () => {
  const { env, c } = filled();
  env.opts.promptText = ['позавчора'];
  c.wipeNow();
  assert.ok(/Не зрозумів/.test(env.calls.dialogs.pop().html));
  env.opts.promptText = ['01.01.2020'];
  c.wipeNow();
  assert.ok(/нічого немає/.test(env.calls.dialogs.pop().html));
  c.onOpen();
  assert.ok(env.calls.menu.indexOf('wipeNow') >= 0);
  const sw = env.ss.getSheetByName('🛡 Захист журналу');
  assert.ok(sw && sw.isSheetHidden());
  const r = rowOf(env, 'a1').row;
  assert.ok(/"a1"/.test(sw.getRange(r, 1).getValue()), 'рядок у рядок з журналом');
});

t('оновлення таблиці версії 21 до 22: копія журналу й розклад «видалені рядки назад»; дані журналу не змінюються', () => {
  const { env, c } = filled(), before = JSON.stringify(c.readJournal_(env.ss).map(o => [o.b, o.c, o.e]));
  env.ss.deleteSheet(env.ss.getSheetByName('🛡 Захист журналу'));
  env.triggers.filter(x => x.getHandlerFunction() === 'mbsChange').forEach(x => env.triggers.splice(env.triggers.indexOf(x), 1));
  env.docProps.setProperty('schema_v', '21');
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 22).done(env.ss), false);
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), '22');
  assert.ok(env.ss.getSheetByName('🛡 Захист журналу').isSheetHidden());
  assert.ok(env.triggers.some(x => x.getHandlerFunction() === 'mbsChange'));
  assert.strictEqual(JSON.stringify(c.readJournal_(env.ss).map(o => [o.b, o.c, o.e])), before);
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 22).done(env.ss), true);
});

console.log(n + ' tests passed');
