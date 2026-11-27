// «🚀 Підключення»: панель, «⚡ Швидке підключення» (стан і інструкція для ШІ), демо-місяць і його очищення; сім'я (учасники, «Хто», бюджети, група в Telegram); оновлення таблиці версії 4.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
const K = (c, x) => require('vm').runInContext(x, c);
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
  const st = env.ss.getSheetByName('⚡ Швидке підключення');
  const state = st.getRange('A8:C22').getValues().filter(r => r[0]), val = k => (state.find(r => r[0] === k) || [])[1];
  assert.ok(/^сім'я: Олена, Іван/.test(val('Режим')), JSON.stringify(state));
  assert.strictEqual(val('Банк'), 'не підключено');
  assert.ok(/ФОП/.test(val('ФОП')) || /1 ФОП/.test(val('ФОП')), val('ФОП'));
  assert.ok(/модуль вимкнено/.test(val('Борги й розстрочки')), 'вимкнений модуль: так і пише');
  assert.ok(/фактів/.test(val('Операції в журналі')));
  const mod = name => env.ss.getSheetByName('Налаштування').getRange('A44:B52').getValues().find(r => r[0] === name)[1];
  assert.strictEqual(mod('💳 Кредитний ліміт і грейс'), 'ні', 'дебетова картка: без ліміту');
  assert.strictEqual(mod('🧾 ФОП'), 'так'); assert.strictEqual(mod('🤝 Борги й розстрочки'), 'ні'); assert.strictEqual(mod('💾 Копії й захист'), 'так');
  assert.ok(env.ss.getSheetByName('ФОП') && env.ss.getSheetByName('ФОП').isSheetHidden(), 'ФОП є: аркуш є, схований разом з рештою');
  assert.strictEqual(env.ss.getSheetByName('Борги').isSheetHidden(), true, 'борги вимкнено: аркуш сховано');
  assert.strictEqual(JSON.parse(env.docProps.getProperty('mods')).credit, false);
  assert.ok(env.docProps.getProperty('fop_day'), 'модуль ФОП застосовано: плани податків синхронізовано');
  assert.ok(/основного: \d+ з 6/.test(st.getRange('A5').getValue()), st.getRange('A5').getValue());
  const all = st.getRange('A1:C110').getValues();
  assert.ok(/HYPERLINK\("https:\/\/send\.monobank\.ua\/jar\//.test(st.getRange('C1').getFormula()), 'посилання «Підтримати» біля назви');
  assert.ok(all.every(r => String(r[1]) !== 'Увімкнено'), 'блоку модулів немає');
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
  ['⚡ Швидке підключення', 'Рахунки', 'ФОП'].forEach(n => env.ss.deleteSheet(env.ss.getSheetByName(n)));
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
  ['⚡ Швидке підключення', 'Рахунки', 'ФОП'].forEach(n => assert.ok(env.ss.getSheetByName(n), n));
  assert.ok(/^Рахунок/.test(j.getRange('AF4').getValue()) && j.getRange('AH5').getFormula());
  assert.strictEqual(j.getRange('D9').getValue(), 'мій рядок');
  assert.strictEqual(env.ss.getSheetByName('Довідник').getRange('A5').getValue(), 'Моя їжа');
  assert.strictEqual(env.ss.getSheetByName('Довідник').getRange('G4').getValue(), 'Чий бюджет');
  assert.strictEqual(set.getRange('A12').getValue(), 'Сім\'я');
  assert.strictEqual(Number(env.ss.getSheetByName('Місяці').getRange('B2').getValue()), 2025);
  assert.ok(/AG\$5:\$AG\$3004/.test(env.ss.getSheetByName('Місяці').getRange('E7').getFormula()));
  assert.ok(/ФОП: дохід понад ліміт/.test(env.ss.getSheetByName('Перевірка').getRange('A5:A30').getValues().map(r => r[0]).join('|')));
});

t('«⚡ Швидке підключення»: first and visible; state for the AI, the step-by-step guide, map with columns, how to answer, what not to do, ready prompts; no secrets', () => {
  const env = makeEnv(), c = env.ctx;
  env.ss.insertSheet('Аркуш1'); c.setupTemplate();
  const Q = '⚡ Швидке підключення', st = env.ss.getSheetByName(Q);
  assert.strictEqual(env.ss.getSheets()[0].getName(), Q); assert.strictEqual(st.isSheetHidden(), false);
  assert.ok(!env.ss.getSheetByName('Старт'));
  const txt = st.getRange('A1:C110').getValues().map(r => r.join(' | ')).join('\n');
  ['# 1. СТАН', '# 2. ДЛЯ ШІ: ЯК ВЕСТИ ЛЮДИНУ', '# 3. КАРТА', '# 4. ЯК ВІДПОВІДАТИ', '# 5. НЕ МОЖНА', '# 6. ГОТОВІ ЗАПИТИ',
    'Прочитай аркуш «⚡ Швидке підключення» і налаштуй зі мною таблицю', 'ЛИШЕ на пряме прохання людини', 'Токени, номери карток, IBAN не проси й не приймай',
    '«Регулярні», перший порожній рядок з 5-го: A назва, B категорія', '«Довідник»: рядок категорії (колонка A), колонка F', 'B27 пріоритет',
    'Колонки з I і далі не чіпати', 'Не видаляй і не сортуй рядки', '«Доступно моїх сьогодні»', 'щомісяця / щокварталу / щороку']
    .forEach(x => assert.ok(txt.indexOf(x) >= 0, x + '\n' + txt.slice(0, 1500)));
  const state = st.getRange('A8:A22').getValues().map(r => r[0]);
  ['Режим', 'Банк', 'Імпорт з банку щодня', 'Telegram-бот', 'Регулярні платежі й доходи', 'Бюджети категорій на місяць', 'Мої правила для радника', 'Модулі увімкнено', 'Операції в журналі']
    .forEach(k => assert.ok(state.indexOf(k) >= 0, k));
  assert.ok(txt.indexOf('—') < 0, 'без довгих тире');
  // з банком і токеном: токена на аркуші немає
  env.userProps.setProperty('mono_token', 'uSecret' + 'x'.repeat(30));
  c.startRefresh_(env.ss);
  const t2 = st.getRange('A1:C110').getValues().map(r => r.join(' ')).join(' ');
  assert.ok(t2.indexOf('uSecret') < 0);
  c.startNow();
  assert.ok(/Готово основного: \d+ з 6/.test(env.calls.dialogs.pop().html.replace(/<[^>]+>/g, '')));
});

t('«⚡ Швидке підключення»: five short lines on view (what to open, what to write, the panel, «Готово x з 6»); everything for the AI below is one folded row group; a refresh keeps one group; an old-layout sheet is redrawn the same day', () => {
  const env = makeEnv(), c = env.ctx, Q = '⚡ Швидке підключення';
  env.ss.insertSheet('Аркуш1'); c.setupTemplate();
  const st = env.ss.getSheetByName(Q), top = st.getRange('A1:A5').getValues().map(r => String(r[0]));
  assert.strictEqual(top[0], Q);
  assert.ok(/Gemini/.test(top[1]) && /зірочка справа вгорі/.test(top[1]), top[1]);
  assert.ok(top[2].indexOf('«Прочитай аркуш «' + Q + '» і налаштуй зі мною таблицю»') >= 0, top[2]);
  assert.ok(/🚀 Підключення/.test(top[3]) && /ніколи в чат/.test(top[3]), top[3]);
  assert.ok(/^Готово основного: \d+ з 6\. Нижче згорнуто інструкцію для ШІ/.test(top[4]), top[4]);
  assert.ok(top.every(x => x.length <= 120), 'рядки короткі');
  assert.ok(/HYPERLINK\(.*Підтримати автора/.test(st.getRange('C1').getFormula()), 'підтримка біля назви');
  // з рядка 6: одна згорнута група до кінця; «+» на рядку 5
  assert.strictEqual(st.getRowGroupDepth(5), 0);
  assert.strictEqual(st.getRowGroupDepth(6), 1); assert.strictEqual(st.getRowGroupDepth(110), 1);
  assert.strictEqual(st.getRowGroup(6, 1).isCollapsed(), true);
  assert.strictEqual(st.groupCtl, 'BEFORE');
  assert.ok(/^# 1\. СТАН/.test(st.getRange('A6').getValue()), 'стан у згорнутому');
  // людина розгорнула; оновлення: група та сама (не друга), знову згорнута
  st.getRowGroup(6, 1).expand();
  c.startRefresh_(env.ss);
  assert.strictEqual(st.groups.length, 1); assert.strictEqual(st.getRowGroupDepth(6), 1);
  assert.strictEqual(st.getRowGroup(6, 1).isCollapsed(), true);
  // аркуш старого вигляду (версія 0.5.0, уже оновлений сьогодні): перемальовується одразу, не завтра
  st.groups = []; st.getRange('A2').setValue('З ШІ: відкрийте Gemini…');
  env.docProps.deleteProperty('start_layout');
  assert.ok(c.startIfDue_(env.ss), 'старий вигляд: оновити зараз');
  assert.strictEqual(st.getRowGroupDepth(6), 1);
  assert.ok(/^1\. Відкрийте Gemini/.test(st.getRange('A2').getValue()));
  assert.strictEqual(c.startIfDue_(env.ss), null, 'новий вигляд і вже сьогодні: нічого');
  const txt = st.getRange('A1:C110').getValues().map(r => r.join(' ')).join(' ');
  assert.ok(txt.indexOf('—') < 0, 'без довгих тире');
});

t('upgrade of a version-14 table to 15: «Старт» becomes «⚡ Швидке підключення» (first, visible), old content replaced; a second «Старт» is kept hidden as «Старт (старий)»', () => {
  const env = makeEnv(), c = env.ctx, Q = '⚡ Швидке підключення';
  env.ss.insertSheet('Аркуш1'); c.setupTemplate();
  const st = env.ss.getSheetByName(Q).setName('Старт');
  st.getRange('A1:C3').setValues([['🚀 Старт', '', ''], ['Меню…', '', ''], ['▰▰▱  2 з 3', '', '']]);
  st.getRange('A6:C6').setValues([['✅ 1', 'Картка й банк', '«🚀 Підключення», крок 1']]);
  env.docProps.setProperty('schema_v', '14');
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.ok(!env.ss.getSheetByName('Старт'));
  const q = env.ss.getSheetByName(Q);
  assert.strictEqual(q, st, 'той самий аркуш, перейменований');
  assert.strictEqual(q.getRange('A1').getValue(), Q);
  assert.ok(!/✅ 1/.test(q.getRange('A6:C6').getValues()[0].join('')), 'старі кроки прибрано');
  assert.strictEqual(env.ss.getSheets()[0].getName(), Q); assert.strictEqual(q.isSheetHidden(), false);
  // обидва аркуші (рідкісне): старий ховається, нічого не видаляється
  const env2 = makeEnv(), c2 = env2.ctx;
  env2.ss.insertSheet('Аркуш1'); c2.setupTemplate();
  env2.ss.insertSheet('Старт').getRange('A1').setValue('🚀 Старт');
  env2.docProps.setProperty('schema_v', '14');
  c2.schemaTick_(env2.ss);
  const old = env2.ss.getSheetByName('Старт (старий)');
  assert.ok(old && old.isSheetHidden() && !env2.ss.getSheetByName('Старт'));
  assert.strictEqual(K(c2, 'STEPS').find(x => x.v === 15).done(env2.ss), true);
});

console.log(n + ' tests passed');
