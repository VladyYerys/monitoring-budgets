// «🚀 Підключення»: панель, чекліст «Старту», демо-місяць і його очищення; сім'я (учасники, «Хто», бюджети, група в Telegram); оновлення таблиці версії 4.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
const { panelSetup } = require('./panel');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const BOT = '123456789:' + 'B'.repeat(35);

t('«🚀 Підключення» in an empty table: builds, opens the side panel; family of two, debit card, income day, FOP group, demo month; checklist updated', () => {
  const env = makeEnv(), c = env.ctx;
  env.ss.insertSheet('Аркуш1');
  c.startWizard();
  const sb = env.calls.sidebars.pop();
  assert.strictEqual(sb.t, '🚀 Підключення');
  ['Додайте картку', 'Сповіщення в Telegram', 'Регулярні платежі', 'Передайте Gemini', 'google.script.run'].forEach(x => assert.ok(sb.html.indexOf(x) >= 0, x));
  assert.ok(!/\$\{/.test(sb.html), 'без незаміщених ${…}');
  let s = c.cnState();
  assert.strictEqual(s.family, false); assert.strictEqual(s.bank.accounts.length, 0); assert.ok(/<svg/.test(s.qrBotFather));
  assert.ok(s.cats.some(x => x.name === 'Продукти'));
  s = c.cnMode(true, 'Олена, Іван');
  assert.strictEqual(JSON.stringify(s.who), JSON.stringify(['Олена', 'Іван', 'спільне']));
  c.cnCard(false); c.cnBday(1); c.cnFop('3', 'Олена');
  assert.throws(() => c.cnBday(40), /від 1 до 31/);
  assert.throws(() => c.cnFop('4', 'Олена'), /1, 2 або 3/);
  c.cnDemo(true);
  const fin = c.cnFinish();
  assert.ok(fin.need > 0);
  const set = env.ss.getSheetByName('Налаштування');
  assert.strictEqual(set.getRange('B13').getValue(), 'сім\'я');
  assert.strictEqual(JSON.stringify(set.getRange('B14:B16').getValues().map(r => r[0])), JSON.stringify(['Олена', 'Іван', '']));
  assert.strictEqual(set.getRange('B7').getValue(), 'дебетова');
  assert.strictEqual(set.getRange('B5').getValue(), 1);
  assert.strictEqual(JSON.stringify(env.ss.getSheetByName('ФОП').getRange('A14:E14').getValues()[0]), JSON.stringify(['Олена', '3', '', 'так', 'так']));
  const rows = c.readJournal_(env.ss).filter(o => !o.blank);
  assert.ok(rows.length > 30 && rows.every(o => o.d.indexOf('[демо] ') === 0 || o.d.indexOf('[Регулярний] [демо] ') === 0));
  assert.ok(rows.some(o => o.d.indexOf('[Регулярний] [демо] Комуналка: ') === 0 && o.e === '🔁 план регулярний'), 'демо-правила «Регулярних» стали планами');
  assert.ok(rows.some(o => o.g === 'Олена') && rows.some(o => o.g === 'спільне'), 'у сім\'ї демо розкладене за «Хто»');
  const st = env.ss.getSheetByName('Старт');
  assert.strictEqual(st.getRange('A6').getValue(), '✓', '«Готово» натиснуто');
  const mod = name => st.getRange('A26:B34').getValues().find(r => r[0] === name)[1];
  assert.strictEqual(mod('💳 Кредитний ліміт і грейс'), 'ні', 'дебетова картка: без ліміту');
  assert.strictEqual(mod('🧾 ФОП'), 'так'); assert.strictEqual(mod('🤝 Борги й розстрочки'), 'ні'); assert.strictEqual(mod('💾 Копії й захист'), 'так');
  assert.strictEqual(env.ss.getSheetByName('ФОП').isSheetHidden(), false, 'ФОП є: аркуш видно');
  assert.strictEqual(env.ss.getSheetByName('Борги').isSheetHidden(), true, 'борги вимкнено: аркуш сховано');
  assert.strictEqual(JSON.parse(env.docProps.getProperty('mods')).credit, false);
  assert.ok(env.docProps.getProperty('fop_day'), 'модуль ФОП застосовано: плани податків синхронізовано');
  assert.ok(/^Готово: \d+ з \d+/.test(st.getRange('A3').getValue()));
  assert.ok(['B19', 'B20', 'B21', 'B22', 'B23'].some(a => /HYPERLINK\("https:\/\/send\.monobank\.ua\/jar\//.test(st.getRange(a).getFormula())), 'посилання «Підтримати» під чеклістом');
});

t('demo clear: only «[демо]» rows, its checkpoint, category plans and the start date go back; your row stays', () => {
  const env = panelSetup(makeEnv(), { credit: true, bday: 25, demo: true }), c = env.ctx;
  const set = () => env.ss.getSheetByName('Налаштування');
  const j = env.ss.getSheetByName('Журнал'), key = v => env.Utilities.formatDate(v, 'Europe/Kiev', 'yyyy-MM-dd');
  const startBefore = env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM') + '-01';
  assert.ok(key(set().getRange('B4').getValue()) < startBefore, 'демо зсунуло старт обліку на минулий місяць');
  const demoRows = c.readJournal_(env.ss).filter(o => !o.blank).length;
  assert.ok(demoRows > 20);
  const free = c.readJournal_(env.ss).find(o => o.blank && o.row > 20).row;
  j.getRange(free, 1, 1, 5).setValues([[new Date(), 99, 'Продукти', 'мій рядок', '✅ факт']]);
  assert.ok(env.ss.getSheetByName('Довідник').getRange('F5').getValue() === 9000);
  c.demoClear();
  const left = c.readJournal_(env.ss).filter(o => !o.blank);
  assert.strictEqual(left.length, 1); assert.strictEqual(left[0].d, 'мій рядок');
  assert.strictEqual(env.ss.getSheetByName('Довідник').getRange('F5').getValue(), '');
  assert.ok(!env.ss.getSheetByName('Залишки банку').getRange('B5:B10').getValues().some(r => r[0] === -15000));
  assert.strictEqual(key(set().getRange('B4').getValue()), startBefore, 'старт обліку як був');
});

t('family: «Хто» list from members, Місяці filter by member keeps plans of own budgets, personal budget counts only its owner', () => {
  const env = makeEnv(), c = env.ctx;
  env.ss.insertSheet('Аркуш1'); c.setupTemplate();
  const j = env.ss.getSheetByName('Журнал'), m = env.ss.getSheetByName('Місяці');
  const dv = j.getRange('G5').getDataValidation();
  assert.strictEqual(dv.spec.kind, 'range'); assert.strictEqual(dv.spec.range, 'B14:B19'); assert.strictEqual(dv.spec.sheet, 'Налаштування');
  assert.strictEqual(m.getRange('D2').getValue(), 'уся сім\'я');
  assert.ok(/IF\(OR\(\$D\$2="",\$D\$2="уся сім'я"\),"<>✱",\$D\$2\)/.test(m.getRange('A3').getFormula()));
  assert.ok(/'Довідник'!\$G\$5=\$D\$2/.test(m.getRange('D7').getFormula()), m.getRange('D7').getFormula());
  env.ss.getSheetByName('Довідник').getRange('F5:G5').setValues([[1000, 'Олена']]);
  const today = env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  const rows = [{ key: today, c: 'Продукти', e: '✅ факт', m: 0, o: 700, b: 700, g: 'Олена', bank: ['', '', '', '', '', ''], kind: 'основна картка' },
    { key: today, c: 'Продукти', e: '✅ факт', m: 0, o: 900, b: 900, g: 'Іван', bank: ['', '', '', '', '', ''], kind: 'основна картка' }];
  const b = c.budgetLines_(env.ss, rows, null)[0];
  assert.ok(/Продукти \(Олена\)/.test(b.text) && /лишилось <b>300 ₴/.test(b.text) && !b.over, b.text);
});

t('Telegram family group: /start in a group is accepted; morning shows yesterday by member', () => {
  const sent = [];
  const env = makeEnv({ promptText: BOT, mono: (url, o) => {
    if (/getMe$/.test(url)) return { code: 200, body: { ok: true, result: { username: 'fam_bot' } } };
    if (/getUpdates$/.test(url)) return { code: 200, body: { ok: true, result: [{ message: { text: '/start@fam_bot', chat: { id: -100500, type: 'supergroup', title: 'Сім\'я' } } }] } };
    if (/sendMessage$/.test(url)) { sent.push(JSON.parse(o.payload)); return { code: 200, body: { ok: true, result: {} } }; }
    return { code: 404, body: {} };
  } }), c = env.ctx;
  env.ss.insertSheet('Аркуш1'); c.setupTemplate();
  c.connectTelegram();
  assert.strictEqual(env.userProps.getProperty('tg_chat'), '-100500');
  assert.ok(/група «Сім'я»/.test(env.calls.dialogs.pop().html));
  const set = env.ss.getSheetByName('Налаштування');
  set.getRange('B13').setValue('сім\'я'); set.getRange('B14:B15').setValues([['Олена'], ['Іван']]);
  const j = env.ss.getSheetByName('Журнал'), today = env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd'), yday = c.addDays_(today, -1);
  const D = env.Utilities.parseDate(yday, 'Europe/Kiev', 'yyyy-MM-dd');
  [[5, 300, 'Олена'], [6, 120, 'Іван'], [7, 50, 'спільне']].forEach(x => {
    j.getRange(x[0], 1, 1, 8).setValues([[D, x[1], 'Продукти', '', '✅ факт', '', x[2], '']]);
    j.put(x[0], 9, 'Витрати', j.cell(x[0], 9).f); j.put(x[0], 13, 0, j.cell(x[0], 13).f); j.put(x[0], 15, x[1], j.cell(x[0], 15).f);
  });
  const r = c.buildMorning_(env.ss, {});
  assert.ok(/👥 Вчора: Олена 300 ₴ · Іван 120 ₴ · спільне 50 ₴/.test(r.text), r.text);
});

t('upgrade of a version-4 table: backup first, new sheets and columns, your journal and categories untouched, «Місяці» keep the year', () => {
  const env = makeEnv(), c = env.ctx;
  env.ss.insertSheet('Аркуш1'); c.setupTemplate();
  // імітуємо таблицю версії 4: без нових аркушів, журнал без AF…AH, налаштування без «Сім'ї», довідник без «Чий бюджет»
  ['Старт', 'Рахунки', 'ФОП'].forEach(n => env.ss.deleteSheet(env.ss.getSheetByName(n)));
  const j = env.ss.getSheetByName('Журнал');
  j.maxC = 31; for (const k of Array.from(j.cells.keys())) if (+k.split(',')[1] > 31) j.cells.delete(k);
  const set = env.ss.getSheetByName('Налаштування'); set.getRange('A12:B20').clearContent();
  env.ss.getSheetByName('Довідник').getRange('G4').clearContent();
  env.docProps.setProperty('schema_v', '4');
  j.getRange('A9:E9').setValues([[new Date(), 250, 'Продукти', 'мій рядок', '✅ факт']]);
  env.ss.getSheetByName('Довідник').getRange('A5').setValue('Моя їжа');
  env.ss.getSheetByName('Місяці').getRange('B2').setValue(2025);
  const files0 = env.files.length;
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.strictEqual(env.files.length, files0 + 1, 'копія перед змінами');
  ['Старт', 'Рахунки', 'ФОП'].forEach(n => assert.ok(env.ss.getSheetByName(n), n));
  assert.ok(/^Рахунок/.test(j.getRange('AF4').getValue()) && j.getRange('AH5').getFormula());
  assert.strictEqual(j.getRange('D9').getValue(), 'мій рядок');
  assert.strictEqual(env.ss.getSheetByName('Довідник').getRange('A5').getValue(), 'Моя їжа');
  assert.strictEqual(env.ss.getSheetByName('Довідник').getRange('G4').getValue(), 'Чий бюджет');
  assert.strictEqual(set.getRange('A12').getValue(), 'Сім\'я');
  assert.strictEqual(Number(env.ss.getSheetByName('Місяці').getRange('B2').getValue()), 2025);
  assert.ok(/AG\$5:\$AG\$3004/.test(env.ss.getSheetByName('Місяці').getRange('E7').getFormula()));
  assert.ok(/ФОП: дохід понад ліміт/.test(env.ss.getSheetByName('Перевірка').getRange('A5:A30').getValues().map(r => r[0]).join('|')));
});

console.log(n + ' tests passed');
