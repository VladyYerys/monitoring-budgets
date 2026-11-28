// Категорії з банку (0.4.9): нові стартові правила, правила за отримувачем, між своїми, план закривається переказом,
// категорія за історією, одразу «✅ факт», «🔎 Сирі дані Mono» (вікно без запису в таблицю), оновлення таблиці версії 12 до 13 і 14.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);
const TOKEN = 'uTestToken' + 'z'.repeat(30), IBAN = 'UA' + '4'.repeat(27);

function setup(items) {
  const state = { items: items || [], accounts: [{ id: 'acc1', currencyCode: 980, type: 'black', creditLimit: 0, balance: 4000000, iban: 'UA00TEST', maskedPan: ['000000******4321'] }] };
  const env = makeEnv({ mono: url => {
    if (/client-info$/.test(url)) return { code: 200, body: { name: 'Тест', accounts: state.accounts } };
    if (/\/personal\/statement\//.test(url)) return { code: 200, body: state.items };
    return { code: 404, body: 'no' };
  }, promptText: TOKEN });
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  env.state = state;
  return env;
}
const day = (env, d) => env.Utilities.formatDate(new Date(Date.now() + (d || 0) * 86400000), 'Europe/Kiev', 'yyyy-MM-dd');
const sec = (env, key, hhmm) => Math.floor(env.Utilities.parseDate(key + ' ' + (hhmm || '12:00') + ':00', 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss').getTime() / 1000);
const cat = (c, rules, it) => c.categorize_(Object.assign({ id: 'x', mcc: 0, comment: '' }, it), it.amount / 100, rules, { name: '', kind: 'основна картка' }).cat;

t('starter rules close 14 of the 15 real examples: shops by name, «погашення» of parts, salary by the payment comment; your own rules still win', () => {
  const env = setup(), c = env.ctx, rules = c.readRules_(env.ss);
  const want = [
    ['ЗаТак', -14340, 'Продукти'], ['VEGETABLES AND FRUITS', -12364, 'Продукти'], ['Експрес Маркет Тест', -21290, 'Продукти'], ['Testvitamin', -8000, 'Продукти'],
    ['Barkinoteatru Test', -6500, 'Розваги'], ['Громадський транспорт', -2500, 'Транспорт'], ['Аврора', -6200, 'Дім і побут'], ['Zoomagazin Test', -50000, 'Тварини'],
    ['Повне погашення Магазин - monomarket', -74583, 'Кредит / розстрочка'], ['Повне погашення ', -66802, 'Кредит / розстрочка'],
    ['Погашення наступного платежу Магазин', -129495, 'Кредит / розстрочка'],
  ];
  want.forEach(w => assert.strictEqual(cat(c, rules, { description: w[0], amount: w[1], mcc: 4829 }), w[2], w[0]));
  // зарплата: «заробітна плата» в коментарі (так Mono показує призначення платежу), сума щомісяця різна
  const salary = { description: 'ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ "ТЕСТ"', comment: 'ІВАНЕНКО ІВАН заробітна плата за 11.2026 року, без ПДВ', mcc: 4829 };
  assert.strictEqual(cat(c, rules, Object.assign({ amount: 2142856 }, salary)), 'Зарплата');
  assert.strictEqual(cat(c, rules, Object.assign({ amount: -2142856 }, salary)), '↔ Переказ (уточніть)', 'лише надходження');
  // MCC: нові групи
  assert.strictEqual(cat(c, rules, { description: 'Shop 1', amount: -100, mcc: 5995 }), 'Тварини');
  assert.strictEqual(cat(c, rules, { description: 'X', amount: -100, mcc: 7832 }), 'Розваги');
  assert.strictEqual(cat(c, rules, { description: 'X', amount: -100, mcc: 3600 }), 'Подорожі', 'діапазон 3501-3999');
  // 15-й: «Переказ на картку» без імені й коментаря: лишається червоним
  assert.strictEqual(cat(c, rules, { description: 'Переказ на картку', amount: -100000, mcc: 4829 }), '↔ Переказ (уточніть)');
  // monomarket без «погашення» (звичайна покупка): не розстрочка
  assert.notStrictEqual(cat(c, rules, { description: 'monomarket', amount: -50000, mcc: 5311 }), 'Кредит / розстрочка');
  // ваше правило вище за стартові
  env.ss.getSheetByName('Правила').getRange('A5:E5').setValues([['так', 'затак', 'опис банку', 'списання', 'Кава та перекуси']]);
  assert.strictEqual(cat(c, c.readRules_(env.ss), { description: 'ЗаТак', amount: -100 }), 'Кава та перекуси');
});

t('rules by recipient: ЄДРПОУ, part of IBAN, name; direction respected; «будь-де» also looks at the recipient name; the draft note shows the recipient', () => {
  const env = setup(), c = env.ctx, sh = env.ss.getSheetByName('Правила');
  const free = sh.getRange('A5:B204').getValues().findIndex(r => !r[0] && !r[1]);
  sh.getRange(5 + free, 1, 3, 5).setValues([
    ['так', '12345678', 'отримувач', 'надходження', 'Зарплата'],
    ['так', IBAN.slice(-12), 'отримувач', 'списання', 'Комуналка'],
    ['так', 'осбб світанок', 'отримувач', 'усі', 'Комуналка']]);
  const rules = c.readRules_(env.ss);
  assert.strictEqual(cat(c, rules, { description: 'ТОВ «Тест»', amount: 2000000, counterEdrpou: '12345678' }), 'Зарплата');
  assert.notStrictEqual(cat(c, rules, { description: 'ТОВ «Тест»', amount: -2000000, counterEdrpou: '12345678' }), 'Зарплата', 'напрям');
  assert.strictEqual(cat(c, rules, { description: 'Оплата', amount: -150000, counterIban: IBAN }), 'Комуналка');
  assert.strictEqual(cat(c, rules, { description: 'Платіж', amount: -90000, counterName: 'ОСББ «Світанок»' }), 'Комуналка');
  assert.ok(!c.whoMatch_(['12345'], { cedrpou: '123456789' }), 'код: лише повний збіг');
  const note = c.noteFor_({ description: 'Платіж', counterName: 'ОСББ Світанок', comment: 'внесок' }, -900);
  assert.ok(/Платіж · ОСББ Світанок · «внесок»/.test(note), note);
  assert.ok(!/Світанок · Світанок/.test(c.noteFor_({ description: 'ОСББ Світанок', counterName: 'ОСББ Світанок' }, -1)), 'назва вже в описі: не дублюємо');
});

t('pure planImport_: between own accounts both rows become «між своїми»; a plan within ±5 days and ±10 % is taken (category, comment); history decides at ≥ 3 and ≥ 80 %, otherwise a hint', () => {
  const env = setup(), c = env.ctx, k = '2026-11-10';
  const types = c.catTypes_(env.ss);
  const row = (o) => Object.assign({ row: 100, key: k, b: 0, c: '', d: '', e: '', bank: ['', '', '', '', '', ''], acct: '', blank: false }, o);
  const ctx = (items, journal, acc) => ({ items: items, journal: journal, logged: new Set(), rules: c.readRules_(env.ss), dateOf: () => k, timeOf: () => '10.11.2026 10:00',
    acc: acc || { name: 'Mono Чорна *4321', kind: 'основна картка', who: '' }, types: types });
  // між своїми: ПриватБанк уже приніс +1000, Mono: −1000 того ж дня
  let p = c.planImport_(ctx([{ id: 'm1', time: 1, amount: -100000, description: 'Переказ на картку', mcc: 4829 }],
    [row({ row: 50, b: 1000, c: '↔ Переказ (уточніть)', e: '🏦 чернетка', bank: ['pb:1', 'З картки Mono', 1000, 4829, '', ''], acct: 'Приват' })]));
  assert.strictEqual(p.drafts[0].c, 'Переказ між своїми рахунками');
  assert.strictEqual(JSON.stringify(p.recats), JSON.stringify([{ row: 50, c: 'Переказ між своїми рахунками', e: '🏦 чернетка' }]), 'без autoFact: чернетка');
  assert.strictEqual(p.stats.own, 1); assert.strictEqual(p.stats.transfers, 0);
  // той самий рахунок чи інша сума: ні
  p = c.planImport_(ctx([{ id: 'm1', time: 1, amount: -100000, description: 'Переказ на картку', mcc: 4829 }],
    [row({ row: 50, b: 999, c: '', e: '🏦 чернетка', bank: ['pb:1', 'x', 999, 0, '', ''], acct: 'Приват' })]));
  assert.strictEqual(p.drafts[0].c, '↔ Переказ (уточніть)');
  // план оренди 15 000 на 08.11: переказ 15 200 10.11 → рядок плану стає чернеткою
  const plan = row({ row: 60, key: '2026-11-08', b: 15000, c: 'Оренда житла', d: '[Регулярний] Оренда: 11.2026', e: '🔁 план регулярний' });
  p = c.planImport_(ctx([{ id: 'm2', time: 2, amount: -1520000, description: 'Переказ на картку', mcc: 4829 }], [plan]));
  assert.strictEqual(p.drafts.length, 0);
  assert.strictEqual(p.takes.length, 1); assert.strictEqual(p.takes[0].row, 60); assert.strictEqual(p.takes[0].c, 'Оренда житла');
  assert.ok(p.takes[0].d.indexOf('[Регулярний] Оренда: 11.2026 · Переказ на картку') === 0, p.takes[0].d);
  assert.strictEqual(p.stats.plans, 1);
  // за межами: 12 % суми чи 6 днів, чи інший напрям: ні
  [[-1680000, '2026-11-08'], [-1500000, '2026-11-02'], [1500000, '2026-11-08']].forEach(x => {
    const q = c.planImport_(ctx([{ id: 'm3', time: 3, amount: x[0], description: 'Переказ', mcc: 4829 }], [Object.assign({}, plan, { key: x[1] })]));
    assert.strictEqual(q.takes.length, 0, JSON.stringify(x));
  });
  // зарплата: план-надходження; сума інша щомісяця (21 428,56 проти 22 000)
  const sal = row({ row: 61, key: '2026-11-12', b: 22000, c: 'Зарплата', d: '[Регулярний] Зарплата: 11.2026', e: '🔁 план регулярний' });
  p = c.planImport_(ctx([{ id: 'm4', time: 4, amount: 2142856, description: 'ТОВ «Тест»', comment: 'заробітна плата за 11.2026', mcc: 4829 }], [sal]));
  assert.strictEqual(p.takes.length, 1, 'правило «Зарплата» + план тієї ж категорії');
  // історія: 3 з 3 → категорія; 2 → лише підказка
  const hist = i => row({ row: 70 + i, key: '2026-10-0' + (i + 1), b: 5000, c: 'Оренда житла', e: '✅ факт', bank: ['h' + i, 'Олена К.', -5000, 4829, '', ''] });
  p = c.planImport_(ctx([{ id: 'm5', time: 5, amount: -500000, description: 'Олена К.', mcc: 4829 }], [hist(0), hist(1), hist(2)]));
  assert.strictEqual(p.drafts[0].c, 'Оренда житла'); assert.strictEqual(p.stats.learned, 1);
  assert.ok(/за історією: Оренда житла \(3 з 3\)/.test(p.drafts[0].d), p.drafts[0].d);
  p = c.planImport_(ctx([{ id: 'm5', time: 5, amount: -500000, description: 'Олена К.', mcc: 4829 }],
    [hist(0), hist(1), row({ row: 80, key: '2026-10-05', b: 700, c: 'Подарунки', e: '✅ факт', bank: ['h9', 'Олена К.', -700, 4829, '', ''] })]));
  assert.strictEqual(p.drafts[0].c, '↔ Переказ (уточніть)');
  assert.ok(/зазвичай: Оренда житла \(2 з 3\), Подарунки \(1\)/.test(p.drafts[0].d), p.drafts[0].d);
  // надходження з тим самим описом: окрема історія
  p = c.planImport_(ctx([{ id: 'm6', time: 6, amount: 500000, description: 'Олена К.', mcc: 4829 }], [hist(0), hist(1), hist(2)]));
  assert.notStrictEqual(p.drafts[0].c, 'Оренда житла');
});

t('import from Mono: the rent plan in the journal becomes the bank operation («✅ факт» at once) and is not added back; the report counts it', () => {
  const env = setup(), c = env.ctx, k = day(env, 1), j = env.ss.getSheetByName('Журнал');
  const D = x => env.Utilities.parseDate(x, 'Europe/Kiev', 'yyyy-MM-dd'), ym = k.slice(5, 7) + '.' + k.slice(0, 4);
  const reg = env.ss.getSheetByName('Регулярні');
  const free = reg.getRange('A6:A60').getValues().findIndex(r => !r[0]);
  const r0 = 6 + free;
  reg.getRange(r0, 1, 1, 5).setValues([['Оренда', 'Оренда житла', 15000, Number(k.slice(8)), 'щомісяця']]);   // план на завтра; банк: сьогодні
  c.regPlansSync_(env.ss);
  const plans = () => c.readJournal_(env.ss).filter(o => o.d.indexOf('[Регулярний] Оренда:') === 0);
  const before = plans();
  assert.ok(before.some(o => o.e === '🔁 план регулярний' && o.key === k), JSON.stringify(before.map(o => [o.key, o.e, o.d])));
  env.state.items = [{ id: 'rent1', time: sec(env, day(env, 0), '09:00'), description: 'Переказ на картку', mcc: 4829, amount: -1490000, balance: 2510000 }];
  c.connectMono();
  const r = c.importBank_(env.ss);
  assert.strictEqual(r.plans, 1); assert.strictEqual(r.transfers, 0);
  assert.ok(/закрито планів 1/.test(c.importSummary_(r)));
  const got = plans().find(o => o.bank[0] === 'rent1');
  assert.ok(got && got.e === '✅ факт' && got.c === 'Оренда житла' && Number(got.b) === 14900, JSON.stringify(got && [got.e, got.c, got.b, got.d]));
  assert.ok(got.d.indexOf('[Регулярний] Оренда: ' + ym + ' · ') === 0);
  c.regPlansSync_(env.ss);
  assert.ok(!plans().some(o => o.e === '🔁 план регулярний' && o.d === '[Регулярний] Оренда: ' + ym), 'план цього місяця не повернувся');
});

t('«🔎 Сирі дані Mono»: all fields, checks of both guesses (salary in the comment, a separate purchase in parts), IBAN and personal codes masked, nothing written', () => {
  const env = setup(), c = env.ctx, k = day(env, -2);
  env.state.items = [
    { id: 'r1', time: sec(env, k, '14:19'), description: 'ТОВАРИСТВО «ТЕСТ»', comment: 'ІВАНЕНКО ІВАН заробітна плата за 11.2026 року, без ПДВ', mcc: 4829, amount: 2142856,
      counterName: 'ТОВАРИСТВО «ТЕСТ»', counterEdrpou: '12345678', counterIban: IBAN },
    { id: 'r2', time: sec(env, k, '15:00'), description: 'Повне погашення Магазин - monomarket', mcc: 4829, amount: -74583 },
    { id: 'r3', time: sec(env, k, '15:01'), description: 'Покупка частинами Магазин', mcc: 5995, amount: -300000 },
    { id: 'r4', time: sec(env, k, '16:00'), description: 'Переказ на картку', mcc: 4829, amount: -100000, counterEdrpou: '1234567890', originalMcc: 6012 },
  ];
  c.connectMono();
  const snap = () => env.ss.getSheets().map(s => s.getName() + ':' + Array.from(s.cells.values()).map(x => String(x.v) + String(x.f)).join('|')).join('#');
  const before = snap();
  const ui = c.SpreadsheetApp.getUi(), p0 = ui.prompt;
  ui.prompt = () => ({ getSelectedButton: () => 'OK', getResponseText: () => '7' });
  c.monoRawNow();
  ui.prompt = p0;
  const d = env.calls.dialogs.pop();
  assert.ok(/Сирі дані Mono/.test(d.t), d.t);
  const h = d.html;
  assert.ok(/за 7 дн\.: операцій <b>4<\/b>/.test(h), h.slice(0, 800));
  assert.ok(/«Заробітна плата»: <b>1<\/b> опер\.; знайдено в: коментар \(comment\) 1/.test(h), 'перше припущення');
  assert.ok(/погашень <b>1<\/b>, інших рядків «частинами» \(можлива сама покупка\) <b>1<\/b>/.test(h), 'друге припущення');
  assert.ok(h.indexOf(IBAN) < 0 && h.indexOf('UA44…4444') >= 0, 'IBAN з прихованою серединою');
  assert.ok(h.indexOf('1234567890') < 0 && h.indexOf('12…90') >= 0 && h.indexOf('12345678') >= 0, 'ІПН людини прихований, ЄДРПОУ ні');
  assert.ok(/4829 \(було 6012\)/.test(h));
  assert.ok(h.indexOf(TOKEN) < 0);
  assert.ok(/Зарплата/.test(h) && /Кредит \/ розстрочка/.test(h), 'що поставить шаблон');
  assert.doesNotThrow(() => new vm.Script(h.split('<script>')[1].split('</script>')[0]));
  assert.strictEqual(snap(), before, 'таблицю не змінено');
});

t('upgrade of a version-12 table to 13: new starter rules go to free rows after yours, «отримувач» in the list; a rule you delete later does not come back', () => {
  const env = setup(), c = env.ctx, sh = env.ss.getSheetByName('Правила');
  const v13 = K(c, 'RULES_V13');
  // таблиця версії 12: нових правил ще немає
  const vals = sh.getRange('A5:G204').getValues().map(r => v13.some(x => String(r[1]) === x[1] && String(r[2]) === x[2]) ? ['', '', '', '', '', '', ''] : r);
  const own = vals.filter(r => r.some(x => x !== '' && x !== null));
  sh.getRange('A5:G204').setValues(own.concat(Array.from({ length: 200 - own.length }, () => ['', '', '', '', '', '', ''])));
  sh.getRange('A' + (5 + own.length) + ':E' + (5 + own.length)).setValues([['так', 'моє правило', 'опис банку', 'усі', 'Подарунки']]);
  env.docProps.deleteProperty('rules_v13');
  env.docProps.setProperty('schema_v', '12');
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  const after = sh.getRange('A5:G204').getValues();
  const at = x => after.findIndex(r => String(r[1]) === x[1] && String(r[2]) === x[2]);
  assert.ok(v13.every(x => at(x) > own.length), 'нові правила після ваших');
  assert.ok(/стартове, \d+\.\d+\.\d+/.test(after[at(v13[0])][6]));
  assert.ok(sh.getRange('C5').getDataValidation().spec.list.indexOf('отримувач') >= 0, '«отримувач» у списку');
  // видалили одне нове правило: наступна перевірка його не повертає
  sh.getRange(5 + at(v13[3]), 1, 1, 7).setValues([['', '', '', '', '', '', '']]);
  env.docProps.setProperty('schema_v', '12');
  c.schemaTick_(env.ss);
  assert.strictEqual(sh.getRange('A5:G204').getValues().filter(r => String(r[1]) === v13[3][1]).length, 0, 'видалене не повернулось');
  // ще раз з нуля: без дублів
  env.docProps.deleteProperty('rules_v13');
  c.rulesUpgrade13_(env.ss);
  v13.filter((x, i) => i !== 3).forEach(x => assert.strictEqual(sh.getRange('A5:G204').getValues().filter(r => String(r[1]) === x[1] && String(r[2]) === x[2]).length, 1, x[1]));
});

t('auto «✅ факт»: a known category closes at once; red transfer, empty category and «⏳ в обробці» stay drafts; when the hold ends, the next import closes it; «ні» in B35 keeps everything as drafts', () => {
  const env = setup(), c = env.ctx, k = day(env, 0), j = () => c.readJournal_(env.ss).filter(o => !o.blank), by = id => j().find(o => o.bank[0] === id);
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('A35').getValue(), K(c, 'AUTO_FACT_LABEL'));
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('B35').getValue(), 'так');
  env.state.items = [
    { id: 'a1', time: sec(env, k, '09:00'), description: 'АТБ', mcc: 5411, amount: -12000, balance: 3988000 },
    { id: 'a2', time: sec(env, k, '09:10'), description: 'Сільпо', mcc: 5411, amount: -30000, balance: 3958000, hold: true },
    { id: 'a3', time: sec(env, k, '09:20'), description: 'Переказ на картку', mcc: 4829, amount: -50000, balance: 3908000 },
    { id: 'a4', time: sec(env, k, '09:30'), description: 'Щось нове', mcc: 7399, amount: -1000, balance: 3907000 },
  ];
  env.state.accounts[0].balance = 3907000;
  c.connectMono();
  const r = c.importBank_(env.ss);
  assert.strictEqual(r.facts, 1);
  assert.strictEqual(by('a1').e, '✅ факт'); assert.strictEqual(by('a2').e, '🏦 чернетка'); assert.ok(/⏳ в обробці/.test(by('a2').d));
  assert.strictEqual(by('a3').e, '🏦 чернетка'); assert.strictEqual(by('a4').e, '🏦 чернетка');
  // «в обробці» завершилось: наступний імпорт ставить факт і прибирає позначку
  env.state.items[1] = Object.assign({}, env.state.items[1], { hold: false });
  const r2 = c.importBank_(env.ss);
  assert.strictEqual(r2.facts, 1);
  assert.strictEqual(by('a2').e, '✅ факт'); assert.ok(!/⏳/.test(by('a2').d), by('a2').d);
  assert.strictEqual(by('a3').e, '🏦 чернетка', 'червоний не чіпається');
  // «ні»: усе як раніше, чернетками
  const env2 = setup(), c2 = env2.ctx;
  env2.ss.getSheetByName('Налаштування').getRange('B35').setValue('ні');
  env2.state.items = [{ id: 'b1', time: sec(env2, day(env2, 0), '10:00'), description: 'АТБ', mcc: 5411, amount: -12000, balance: 3988000 }];
  env2.state.accounts[0].balance = 3988000;
  c2.connectMono();
  assert.strictEqual(c2.importBank_(env2.ss).facts, 0);
  assert.strictEqual(c2.readJournal_(env2.ss).find(o => o.bank[0] === 'b1').e, '🏦 чернетка');
});

t('auto «✅ факт» in pure planImport_: between own accounts the other red draft also becomes a fact; a plan taken while «в обробці» stays a draft', () => {
  const env = setup(), c = env.ctx, k = '2026-11-10', types = c.catTypes_(env.ss);
  const row = o => Object.assign({ row: 100, key: k, b: 0, c: '', d: '', e: '', bank: ['', '', '', '', '', ''], acct: '', blank: false }, o);
  const ctx = (items, journal) => ({ items: items, journal: journal, logged: new Set(), rules: c.readRules_(env.ss), dateOf: () => k, timeOf: () => '10.11.2026 10:00',
    acc: { name: 'Mono Чорна *4321', kind: 'основна картка', who: '' }, types: types, autoFact: true });
  let p = c.planImport_(ctx([{ id: 'm1', time: 1, amount: -100000, description: 'Переказ на картку', mcc: 4829 }],
    [row({ row: 50, b: 1000, c: '↔ Переказ (уточніть)', d: 'З картки Mono · надходження', e: '🏦 чернетка', bank: ['pb:1', 'З картки Mono', 1000, 4829, '', ''], acct: 'Приват' })]));
  assert.strictEqual(p.drafts[0].e, '✅ факт');
  assert.strictEqual(JSON.stringify(p.recats), JSON.stringify([{ row: 50, c: 'Переказ між своїми рахунками', e: '✅ факт' }]));
  assert.strictEqual(p.stats.facts, 2);
  const plan = row({ row: 60, key: '2026-11-08', b: 15000, c: 'Оренда житла', d: '[Регулярний] Оренда: 11.2026', e: '🔁 план регулярний' });
  p = c.planImport_(ctx([{ id: 'm2', time: 2, amount: -1500000, description: 'Переказ на картку', mcc: 4829, hold: true }], [plan]));
  assert.strictEqual(p.takes[0].e, '🏦 чернетка', '«в обробці»: чернетка');
  p = c.planImport_(Object.assign(ctx([{ id: 'm2', time: 2, amount: -1500000, description: 'Переказ на картку', mcc: 4829 }], [plan]), { autoFact: false }));
  assert.strictEqual(p.takes[0].e, '🏦 чернетка', 'B35 «ні»');
});

t('upgrade of a version-13 table to 14: «одразу «✅ факт»» row added on «Налаштування» (B35 «так»); your «ні» stays', () => {
  const env = setup(), c = env.ctx, set = env.ss.getSheetByName('Налаштування');
  set.getRange('A35:C35').setValues([['', '', '']]);
  assert.strictEqual(c.autoFact_(env.ss), true, 'рядка ще немає: типово так');
  env.docProps.setProperty('schema_v', '13');
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(set.getRange('A35').getValue(), K(c, 'AUTO_FACT_LABEL'));
  assert.strictEqual(set.getRange('B35').getValue(), 'так');
  set.getRange('B35').setValue('ні');
  env.docProps.setProperty('schema_v', '13');
  c.schemaTick_(env.ss);
  assert.strictEqual(set.getRange('B35').getValue(), 'ні');
  assert.strictEqual(c.autoFact_(env.ss), false);
});

/* ---------- ↩ повернення за покупку (0.5.3) ---------- */

function refundCtx(env, items, journal, extra) {
  const c = env.ctx;
  return Object.assign({ items: items, journal: journal, logged: new Set(), rules: c.readRules_(env.ss), dateOf: it => it.k, timeOf: () => '27.11.2026 01:24',
    acc: { name: 'Mono Чорна *4321', kind: 'основна картка', who: '' }, types: c.catTypes_(env.ss), autoFact: true }, extra || {});
}
const jr = o => Object.assign({ row: 100, key: '2026-11-21', b: 0, c: '', d: '', e: '✅ факт', bank: ['', '', '', '', '', ''], acct: '', blank: false, l: '' }, o);
const buy = (row, key, amt, cat, extra) => jr(Object.assign({ row: row, key: key, b: amt, c: cat, bank: ['b' + row, 'MEBLIMAX', -amt, 5712, '', ''] }, extra || {}));
const refund = (id, k, amt, o) => Object.assign({ id: id, k: k, time: 1000 + Number(id.replace(/\D/g, '') || 0), amount: Math.round(amt * 100), description: 'Скасування. MEBLIMAX', mcc: 5712 }, o || {});

t('↩ refund «Скасування. МАГАЗИН»: the purchase category with a minus, «✅ факт», a note with the purchase date; partial refunds too; bigger than the purchase, older than 60 days or another account: no purchase, so a rule or a draft without category', () => {
  const env = setup(), c = env.ctx;
  let p = c.planImport_(refundCtx(env, [refund('r1', '2026-11-27', 525)], [buy(50, '2026-11-21', 525, 'Дім і побут')]));
  const d = p.drafts[0];
  assert.strictEqual(JSON.stringify([d.b, d.c, d.e]), JSON.stringify([-525, 'Дім і побут', '✅ факт']));
  assert.ok(/^↩ повернення за покупку 21\.11 · Скасування\. MEBLIMAX$/.test(d.d), d.d);
  assert.strictEqual(p.stats.refunds, 1); assert.strictEqual(p.stats.facts, 1); assert.strictEqual(p.stats.uncategorized, 0); assert.strictEqual(p.stats.transfers, 0);
  assert.ok(/↩ повернення → Дім і побут, мінусом \(за покупку 21\.11\)/.test(p.log[0][8]), p.log[0][8]);
  // частково (200 з 525): та сама покупка
  p = c.planImport_(refundCtx(env, [refund('r2', '2026-11-27', 200)], [buy(50, '2026-11-21', 525, 'Дім і побут')]));
  assert.strictEqual(p.drafts[0].b, -200); assert.strictEqual(p.drafts[0].c, 'Дім і побут');
  // більше за покупку, старше 60 днів, інший рахунок: покупки немає → без правила: чернетка без категорії, мінусом
  const odd = { mcc: 7999 };                                            // MCC без стартового правила
  [[refund('r3', '2026-11-27', 600, odd), buy(50, '2026-11-21', 525, 'Дім і побут')],
   [refund('r4', '2026-11-27', 525, odd), buy(50, '2026-09-27', 525, 'Дім і побут')],
   [refund('r5', '2026-11-27', 525, odd), buy(50, '2026-11-21', 525, 'Дім і побут', { acct: 'Приват' })]].forEach(x => {
    const q = c.planImport_(refundCtx(env, [x[0]], [x[1]]));
    assert.strictEqual(JSON.stringify([q.drafts[0].b, q.drafts[0].c, q.drafts[0].e]), JSON.stringify([-x[0].amount / 100, '', '🏦 чернетка']), x[0].id);
    assert.ok(/^↩ повернення: оберіть категорію покупки/.test(q.drafts[0].d), q.drafts[0].d);
    assert.strictEqual(q.stats.uncategorized, 1); assert.strictEqual(q.stats.refunds, 1);
  });
  // покупки немає, але MCC магазину має стартове правило (меблі → «Дім і побут»): воно
  p = c.planImport_(refundCtx(env, [refund('r9', '2026-11-27', 600)], [buy(50, '2026-11-21', 525, 'Подарунки')]));
  assert.strictEqual(JSON.stringify([p.drafts[0].c, p.drafts[0].e]), JSON.stringify(['Дім і побут', '✅ факт']));
  // покупки немає, але є правило магазину (стартове «АТБ» → Продукти, для списань): воно
  p = c.planImport_(refundCtx(env, [refund('r6', '2026-11-27', 99.9, { description: 'Скасування. АТБ', mcc: 5411 })], []));
  assert.strictEqual(JSON.stringify([p.drafts[0].b, p.drafts[0].c, p.drafts[0].e]), JSON.stringify([-99.9, 'Продукти', '✅ факт']));
  // два повернення по 525 за дві покупки: кожне своїй (новіша спершу)
  p = c.planImport_(refundCtx(env, [refund('r7', '2026-11-27', 525), refund('r8', '2026-11-27', 525)],
    [buy(50, '2026-11-21', 525, 'Дім і побут'), buy(51, '2026-11-10', 525, 'Подарунки')]));
  assert.strictEqual(JSON.stringify(p.drafts.map(x => [x.c, x.d.slice(0, 29)])),
    JSON.stringify([['Дім і побут', '↩ повернення за покупку 21.11'], ['Подарунки', '↩ повернення за покупку 10.11']]));
});

t('↩ refund without the word: same shop and MCC as a purchase → a draft with the purchase category (check it); other MCC or a person-to-person transfer: not a refund; your own −525 row is not duplicated', () => {
  const env = setup(), c = env.ctx;
  let p = c.planImport_(refundCtx(env, [refund('w1', '2026-11-27', 525, { description: 'MEBLIMAX' })], [buy(50, '2026-11-21', 525, 'Дім і побут')]));
  assert.strictEqual(JSON.stringify([p.drafts[0].b, p.drafts[0].c, p.drafts[0].e]), JSON.stringify([-525, 'Дім і побут', '🏦 чернетка']));
  assert.ok(/перевірте: схоже на повернення/.test(p.drafts[0].d), p.drafts[0].d);
  p = c.planImport_(refundCtx(env, [refund('w2', '2026-11-27', 525, { description: 'MEBLIMAX', mcc: 5999 })], [buy(50, '2026-11-21', 525, 'Дім і побут')]));
  assert.strictEqual(p.stats.refunds, 0); assert.ok(p.drafts[0].b > 0, 'інший MCC: звичайне надходження');
  // ваш рядок «−525 Дім і побут» без номера банку: операція банку пристає до нього
  p = c.planImport_(refundCtx(env, [refund('w3', '2026-11-27', 525)], [jr({ row: 90, key: '2026-11-27', b: -525, c: 'Дім і побут', l: 525 })]));
  assert.strictEqual(p.drafts.length, 0); assert.strictEqual(p.marks[0].row, 90);
});

t('↩ «Скасування … в обробці»: for a purchase already settled it is a draft «⏳ в обробці» (not removed as «cancelled by bank»); for a pending purchase or no purchase it is skipped as before; when settled it becomes a fact and keeps the minus', () => {
  const env = setup(), c = env.ctx, held = refund('h1', '2026-11-27', 525, { hold: true });
  let p = c.planImport_(refundCtx(env, [], [buy(50, '2026-11-21', 525, 'Дім і побут')], { pending: [held] }));
  assert.strictEqual(JSON.stringify([p.drafts[0].b, p.drafts[0].c, p.drafts[0].e]), JSON.stringify([-525, 'Дім і побут', '🏦 чернетка']));
  assert.ok(/⏳ в обробці/.test(p.drafts[0].d)); assert.strictEqual(p.stats.skipped, 0);
  // покупка ще «в обробці» (банк скасовує неcписану оплату) чи покупки немає: пропуск, рядок у журналі «Банк» один раз
  [[buy(50, '2026-11-25', 525, 'Дім і побут', { e: '🏦 чернетка', d: 'MEBLIMAX · ⏳ в обробці' })], []].forEach(j => {
    const q = c.planImport_(refundCtx(env, [], j, { pending: [held] }));
    assert.strictEqual(q.drafts.length, 0); assert.strictEqual(q.stats.skipped, 1); assert.strictEqual(q.log[0][0], 'skip:h1');
    assert.strictEqual(c.planImport_(refundCtx(env, [], j, { pending: [held], logged: new Set(['skip:h1']) })).log.length, 0);
  });
  // уже в журналі чернеткою «в обробці»: не прибирається, поки Mono тримає її «в обробці»
  const draft = jr({ row: 60, key: '2026-11-27', b: -525, c: 'Дім і побут', e: '🏦 чернетка', d: '↩ повернення за покупку 21.11 · Скасування. MEBLIMAX · ⏳ в обробці',
    bank: ['h1', 'Скасування. MEBLIMAX', 525, 5712, '', ''] });
  p = c.planImport_(refundCtx(env, [], [buy(50, '2026-11-21', 525, 'Дім і побут'), draft], { pending: [held], coverFrom: '2026-11-25', coverTo: '2026-11-28' }));
  assert.strictEqual(p.vanish.length, 0); assert.strictEqual(p.updates.length, 0);
  // проведено: «✅ факт», сума лишається −525
  p = c.planImport_(refundCtx(env, [refund('h1', '2026-11-27', 525)], [buy(50, '2026-11-21', 525, 'Дім і побут'), draft]));
  assert.strictEqual(p.updates.length, 1); assert.strictEqual(p.updates[0].b, null, 'мінус не перетворюється на плюс'); assert.strictEqual(p.updates[0].e, '✅ факт');
  assert.ok(!/⏳/.test(p.updates[0].d)); assert.strictEqual(p.stats.changed, 0);
});

t('↩ refund through the Mono import: purchase, then «Скасування. …» in processing, then settled; the journal has −525 in the purchase category; report and Telegram text count refunds; the raw window marks it', () => {
  const env = setup(), c = env.ctx, k0 = day(env, -3), k1 = day(env, 0);
  env.state.items = [{ id: 'p1', time: sec(env, k0, '15:26'), description: 'MEBLIMAX', mcc: 5712, amount: -52500, balance: 3947500 }];
  c.connectMono();
  c.importBank_(env.ss);
  const row = id => c.readJournal_(env.ss).find(o => o.bank[0] === id);
  env.ss.getSheetByName('Журнал').getRange(row('p1').row, 3).setValue('Дім і побут');        // ви поставили категорію покупці
  env.state.items.push({ id: 'r1', time: sec(env, k1, '01:24'), description: 'Скасування. MEBLIMAX', mcc: 5712, amount: 52500, hold: true, balance: 3947500 });
  let r = c.importBank_(env.ss);
  assert.strictEqual(JSON.stringify([row('r1').b, row('r1').c, row('r1').e]), JSON.stringify([-525, 'Дім і побут', '🏦 чернетка']));
  assert.strictEqual(r.refunds, 1); assert.strictEqual(r.skipped, 0);
  assert.ok(c.importReport_(r).some(l => /Повернень за покупки: <b>1<\/b>/.test(l)));
  env.state.items[1].hold = false;
  r = c.importBank_(env.ss);
  assert.strictEqual(JSON.stringify([row('r1').b, row('r1').c, row('r1').e]), JSON.stringify([-525, 'Дім і побут', '✅ факт']));
  assert.strictEqual(c.readJournal_(env.ss).filter(o => o.bank[0] === 'r1').length, 1);
  assert.ok(/↩ Повернень за покупки: 1/.test(c.importText_(env.ss, Object.assign(c.emptyReport_(), { refunds: 1, accounts: [] }), { at: '12:45', attempt: 0 })));
  const raw = c.rawRows_([env.state.items[1]], c.readRules_(env.ss), { name: 'x', kind: 'основна картка' }, () => '', {});
  assert.ok(/↩ повернення/.test(raw.rows[0].cat), raw.rows[0].cat);
});

t('↩ formulas: «Списання» (T) counts a minus in expenses as money back to «мої»; upgrade of a version-15 table to 16 rewrites only T and the hint', () => {
  const env = setup(), c = env.ctx, T = K(c, 'journalFormulas_()').T;
  assert.ok(T.indexOf('IF(OR(N(AH{r})<0,AND(N(O{r})<0,N(AH{r})>0)),O{r},0)') >= 0, T);
  const j = env.ss.getSheetByName('Журнал'), old = '=IF(OR(Q5="",N(Q5)<1,AG5="рахунок ФОП"),0,IF(N(AH5)<0,O5,0))';
  j.getRange('T5').setFormula(old); j.getRange('T3004').setFormula(old.replace(/5/g, '3004'));
  const s0 = j.getRange('S5').getFormula();
  j.getRange('A3').setNote('Пишіть у рядки нижче жовтої лінії: дата, сума (завжди додатна), категорія зі списку, статус.');
  env.docProps.setProperty('schema_v', '15');
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 16).done(env.ss), false);
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), '16');
  assert.strictEqual(j.getRange('T5').getFormula(), T.split('{r}').join('5'));
  assert.strictEqual(j.getRange('S5').getFormula(), s0, 'інші колонки не чіпали');
  assert.ok(/з мінусом, у категорії покупки/.test(j.getRange('A3').getNotes()[0][0]), j.getRange('A3').getNotes()[0][0]);
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 16).done(env.ss), true);
});

console.log(n + ' tests passed');
