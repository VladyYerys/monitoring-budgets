// Модулі-перемикачі (панель, крок 5; блок на «Налаштуваннях»): кожен вмикається й вимикається окремо; грейс Mono на «Головному»; оновлення 5→6 і 11→12.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);
const BOT = '123456789:' + 'B'.repeat(35);

function setup(opts) {
  const env = makeEnv(opts || {});
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  return env;
}
const st = env => env.ss.getSheetByName('Налаштування');
const row = (env, name) => 44 + K(env.ctx, 'MODS').findIndex(m => m.name === name);
/** Як людина: змінила клітинку модуля (блок «🧩 Модулі» на «Налаштуваннях»), спрацював розклад mbsEdit. */
function toggle(env, name, v) {
  const r = st(env).getRange(row(env, name), 2);
  r.setValue(v);
  env.ctx.mbsEdit({ range: r });
}
const mainRow = (env, label) => 1 + env.ss.getSheetByName('Головне').getRange('A1:A120').getValues().findIndex(r => r[0] === label);
const trig = (env, fn) => env.triggers.some(x => x.getHandlerFunction() === fn);

t('new table: block «🧩 Модулі» on «Налаштування» (rows 42…) with 9 modules, typical values, mods saved, FOP and debts hidden; «Старт» without modules', () => {
  const env = setup(), c = env.ctx;
  assert.strictEqual(st(env).getRange('A42').getValue(), '🧩 Модулі');
  const v = st(env).getRange('A44:B52').getValues();
  assert.strictEqual(JSON.stringify(v.map(r => r[1])), JSON.stringify(['так', 'так', 'ні', 'так', 'ні', 'так', 'ні', 'ні', 'ні']));
  assert.strictEqual(v[8][0], '🗂 Усі аркуші');
  assert.ok(st(env).getRange('B44').getDataValidation().reject, 'лише так / ні');
  const mods = JSON.parse(env.docProps.getProperty('mods'));
  assert.strictEqual(mods.tg, true); assert.strictEqual(mods.fop, false); assert.strictEqual(mods.debts, false);
  assert.strictEqual(env.ss.getSheetByName('ФОП').isSheetHidden(), true);
  assert.strictEqual(env.ss.getSheetByName('Борги').isSheetHidden(), true);
  const start = env.ss.getSheetByName('Старт');
  assert.ok(start.getRange('A1:C40').getValues().every(r => r.join('').indexOf('🧩 Модулі') < 0 || /крок 5/.test(r.join(''))), 'на «Старті» блоку модулів немає');
  const p = start.prots.find(x => /Monitoring Budgets/.test(x.getDescription()));
  assert.ok(p && p.unprot.length === 0, '«Старт» пише лише скрипт: вільних клітинок немає');
  assert.ok(mainRow(env, 'Грейс: внести до строку') > 0 && mainRow(env, 'Борги: я винен (залишок)') > 0);
  assert.ok(env.ss.getSheetByName('Головне').isRowHiddenByUser(mainRow(env, 'Борги: я винен (залишок)')), 'борги вимкнено: рядки сховано');
});

t('credit off: limit and grace rows hidden on «Головне», «Дні» F:H hidden, morning without limit and grace; on again: all back', () => {
  const env = setup(), c = env.ctx, m = env.ss.getSheetByName('Головне'), d = env.ss.getSheetByName('Дні');
  toggle(env, '💳 Кредитний ліміт і грейс', 'ні');
  assert.strictEqual(c.modOn_('credit'), false);
  ['Кредитний ліміт', 'Борг за лімітом зараз', 'Грейс: внести до строку', 'Грейс: строк'].forEach(l => assert.ok(m.isRowHiddenByUser(mainRow(env, l)), l));
  assert.ok(!m.isRowHiddenByUser(mainRow(env, 'На картці зараз')), 'спільні рядки лишаються');
  assert.ok(d.isColumnHiddenByUser(6) && d.isColumnHiddenByUser(8));
  const mo = env.ss.getSheetByName('Місяці'), mrow = l => c.monthsSummaryRow_(l);
  assert.ok(mo.isRowHiddenByUser(mrow('Витрачено з ліміту')) && mo.isRowHiddenByUser(mrow('ЧОМУ З ЛІМІТУ')) && mo.isRowHiddenByUser(mrow('Разом з ліміту')));
  assert.ok(!mo.isRowHiddenByUser(mrow('Мої надходження')) && !mo.isRowHiddenByUser(mrow('Непомітні разом')), 'решта «Місяців» видно');
  m.put(mainRow(env, 'Кредитний ліміт'), 2, 50000, m.cell(mainRow(env, 'Кредитний ліміт'), 2).f);
  m.put(mainRow(env, 'Грейс: внести до строку'), 2, 9000, m.cell(mainRow(env, 'Грейс: внести до строку'), 2).f);
  let r = c.buildMorning_(env.ss, {});
  assert.ok(!/Борг за лімітом|Грейс/.test(r.text) && /💳 На картці/.test(r.text), r.text);
  toggle(env, '💳 Кредитний ліміт і грейс', 'так');
  assert.ok(!m.isRowHiddenByUser(mainRow(env, 'Кредитний ліміт')) && !d.isColumnHiddenByUser(6) && !mo.isRowHiddenByUser(mrow('ЧОМУ З ЛІМІТУ')));
  const tz = 'Europe/Kiev', today = env.Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  m.put(mainRow(env, 'Грейс: строк'), 2, env.Utilities.parseDate(c.lastDayOf_(c.addDays_(c.lastDayOf_(today.slice(0, 7)), 1).slice(0, 7)), tz, 'yyyy-MM-dd'), m.cell(mainRow(env, 'Грейс: строк'), 2).f);
  r = c.buildMorning_(env.ss, {});
  assert.ok(/💳 Грейс до \d\d\.\d\d: внести <b>9 000 ₴<\/b>/.test(r.text), r.text);
});

t('grace formulas on «Головне»: debt at the end of last month, what already came, deadline, transfers and fee from the rate on «Налаштування»', () => {
  const env = setup(), c = env.ctx, m = env.ss.getSheetByName('Головне'), F = l => m.getRange(mainRow(env, l), 2).getFormula();
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('B10').getValue(), 0.035);
  assert.ok(/'Налаштування'!\$B\$10/.test(F('Грейс: ставка комісії за переказ з ліміту')));
  assert.ok(/MAX\(0,N\(IFERROR\(INDEX\('Дні'!\$F\$6:\$F\$98,MATCH\(EOMONTH\(TODAY\(\),-1\)/.test(F('Грейс: борг на кінець минулого місяця')), F('Грейс: борг на кінець минулого місяця'));
  const left = 'B' + mainRow(env, 'Грейс: лишилось внести до кінця місяця');
  assert.strictEqual(F('Грейс: строк'), '=IF(' + left + '>0,EOMONTH(TODAY(),0),EOMONTH(TODAY(),1))');
  assert.ok(F('Грейс: внести до строку').indexOf(left) > 0);
  assert.ok(/"Надходження".*"основна картка"/.test(F('Грейс: надійде за планом до строку')));
  assert.strictEqual(F('Грейс: комісія, якщо все переказами'), '=ROUND(B' + mainRow(env, 'Грейс: внести до строку') + '*B' + mainRow(env, 'Грейс: ставка комісії за переказ з ліміту') + ',0)');
});

t('Telegram off: no sends, morning trigger removed, token kept; on again: trigger back; connecting the bot turns the module on', () => {
  const env = setup({ mono: (url, o) => (/getMe/.test(url) ? { code: 200, body: { ok: true, result: { username: 'mb_bot' } } } :
    /getUpdates/.test(url) ? { code: 200, body: { ok: true, result: [{ message: { chat: { id: 77, type: 'private' }, text: '/start' } }] } } :
    { code: 200, body: { ok: true, result: {} } }), promptText: BOT }), c = env.ctx;
  toggle(env, '📨 Telegram', 'ні');
  c.connectTelegram();
  assert.strictEqual(c.modOn_('tg'), true, 'підключення вмикає модуль');
  assert.strictEqual(st(env).getRange(row(env, '📨 Telegram'), 2).getValue(), 'так');
  assert.ok(trig(env, 'mbsMorning'));
  toggle(env, '📨 Telegram', 'ні');
  assert.ok(!trig(env, 'mbsMorning'));
  assert.ok(env.userProps.getProperty('tg_token'), 'токен лишився');
  const sent0 = env.calls.fetch.filter(u => /sendMessage/.test(u)).length;
  c.mbsMorning();
  assert.throws(() => c.tgSend_('x'), /вимкнено/);
  assert.strictEqual(env.calls.fetch.filter(u => /sendMessage/.test(u)).length, sent0, 'нічого не надіслано');
  assert.strictEqual(c.tgReady_(), false);
  toggle(env, '📨 Telegram', 'так');
  assert.ok(trig(env, 'mbsMorning'));
  assert.strictEqual(c.tgReady_(), true);
});

t('copies and protection off: nightly trigger and our protection gone, mbsBackup idle, step 4 still «done»; on: both back', () => {
  const env = setup(), c = env.ctx;
  assert.ok(trig(env, 'mbsBackup'));
  toggle(env, '💾 Копії й захист', 'ні');
  assert.ok(!trig(env, 'mbsBackup'));
  assert.strictEqual(env.ss.getSheetByName('Журнал').prots.length, 0);
  const f0 = env.files.length;
  c.mbsBackup();
  assert.strictEqual(env.files.length, f0, 'копії немає');
  assert.strictEqual(K(c, 'STEPS')[3].done(env.ss), true);
  c.tableStatus();
  assert.ok(/модуль вимкнено/.test(env.calls.dialogs.pop().html));
  toggle(env, '💾 Копії й захист', 'так');
  assert.ok(trig(env, 'mbsBackup'));
  assert.ok(env.ss.getSheetByName('Журнал').prots.length === 1 && c.protectionOn_(env.ss));
});

t('calendar: «так» on «Старт» creates and fills the calendar; menu «вимкнути» writes «ні» back; budgets off: no budget alerts', () => {
  const env = setup({ alertAnswer: 'NO' }), c = env.ctx;
  toggle(env, '📆 Календар платежів', 'так');
  assert.strictEqual(env.cals.length, 1);
  assert.strictEqual(c.calOn_(), true);
  c.calendarOff();
  assert.strictEqual(st(env).getRange(row(env, '📆 Календар платежів'), 2).getValue(), 'ні');
  assert.strictEqual(c.calOn_(), false);
  const d = env.ss.getSheetByName('Довідник');
  d.getRange('F5').setValue(100);
  const rows = [{ key: env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd'), c: d.getRange('A5').getValue(), e: '✅ факт', m: 0, o: 500, b: 500, g: '', kind: 'основна картка', bank: ['', '', '', '', '', ''] }];
  assert.ok(c.budgetLines_(env.ss, rows, null).some(x => x.over));
  toggle(env, '💰 Бюджети категорій', 'ні');
  assert.strictEqual(c.budgetLines_(env.ss, rows, null).length, 0);
  assert.ok(c.budgetLines_(env.ss, rows, null, true).length === 1, 'знімок для радника рахує бюджети все одно');
});

t('everything off: a plain journal with «Місяці», «Дні» and charts; only the order schedule stays; tick applies a missed edit', () => {
  const env = setup(), c = env.ctx;
  K(c, 'MODS').forEach((m, i) => st(env).getRange(44 + i, 2).setValue('ні'));
  c.mbsTick();                                                            // правку пропущено: застосує розклад
  assert.ok(K(c, 'MODS').every(m => c.modOn_(m.id) === false));
  assert.strictEqual(JSON.stringify(env.triggers.map(x => x.getHandlerFunction()).sort()), JSON.stringify(['mbsEdit', 'mbsTick']));
  ['Старт', 'Журнал', 'Регулярні', 'Charts', 'charts_mob'].forEach(n => assert.strictEqual(env.ss.getSheetByName(n).isSheetHidden(), false, n));
  ['Місяці', 'Дні', 'Довідник', 'Налаштування', 'Головне'].forEach(n => assert.strictEqual(env.ss.getSheetByName(n).isSheetHidden(), true, n + ': працює схованим'));
  assert.ok(env.ss.getSheetByName('Charts').charts.length >= 6 && env.ss.getSheetByName('charts_mob').charts.length >= 4);
  const r = c.buildMorning_(env.ss, {});
  assert.ok(!/Нічна копія/.test(r.text), 'самоперевірка не скаржиться на вимкнені копії');
  c.modulesNow();
  const sb = env.calls.sidebars.pop();
  assert.ok(sb.t === '🚀 Підключення' && /var GO = 5,/.test(sb.html), 'меню «🧩 Модулі»: панель одразу на кроці 5');
});

t('upgrade of a version-5 table to 6: backup, modules from what already works (calendar on, FOP filled), «Борги», grace rows, settings rows; your data untouched', () => {
  const env = setup(), c = env.ctx;
  // таблиця версії 5: без «Боргів», без модулів, без грейсу й нових рядків «Налаштувань»
  env.ss.deleteSheet(env.ss.getSheetByName('Борги'));
  st(env).getRange('A42:C52').clearContent();
  const m = env.ss.getSheetByName('Головне'), g = mainRow(env, 'Грейс: ставка комісії за переказ з ліміту');
  m.getRange('A' + g + ':C' + (g + 12)).clearContent();
  const set = env.ss.getSheetByName('Налаштування');
  set.getRange('A10:B10').clearContent(); set.getRange('A22:B24').clearContent();
  env.docProps.deleteProperty('mods'); K(c, 'MODS_CACHE_ = null');
  env.docProps.setProperty('cal', 'on');
  env.ss.getSheetByName('ФОП').showSheet().getRange('A14:E14').setValues([['Олена', '2', '', 'так', 'так']]);
  env.docProps.setProperty('schema_v', '5');
  env.ss.getSheetByName('Журнал').getRange('A9:E9').setValues([[new Date(), 250, 'Продукти', 'мій рядок', '✅ факт']]);
  const files0 = env.files.length;
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.strictEqual(env.files.length, files0 + 1, 'копія перед змінами');
  assert.ok(env.ss.getSheetByName('Борги'));
  const v = st(env).getRange('A44:B52').getValues(), mod = name => v.find(r => r[0] === name)[1];
  assert.strictEqual(mod('📆 Календар платежів'), 'так', 'календар уже працював');
  assert.strictEqual(mod('🧾 ФОП'), 'так', 'ФОП заповнено');
  assert.strictEqual(mod('🤝 Борги й розстрочки'), 'ні');
  assert.strictEqual(env.ss.getSheetByName('ФОП').isSheetHidden(), true, 'модуль увімкнено, але аркуш схований, як і решта («🗂 Усі аркуші»: ні)');
  assert.ok(mainRow(env, 'Грейс: строк') > 0);
  assert.strictEqual(set.getRange('B10').getValue(), 0.035);
  assert.strictEqual(set.getRange('B23').getValue(), 'категорії');
  const mine = c.readJournal_(env.ss).filter(o => o.d === 'мій рядок');
  assert.ok(mine.length === 1 && mine[0].b === 250 && mine[0].c === 'Продукти', 'ваш рядок на місці (порядок міг змінитись: додались плани ФОП)');
  assert.strictEqual(K(c, 'STEPS')[5].done(env.ss), true);
});

t('upgrade of a version-11 table to 12: the module block moves from «Старт» to «Налаштування» with your choices; «Старт» keeps only the steps; nothing switches', () => {
  const env = setup(), c = env.ctx, start = env.ss.getSheetByName('Старт'), set = st(env), MODS = K(c, 'MODS');
  toggle(env, '📆 Календар платежів', 'так'); toggle(env, '💳 Кредитний ліміт і грейс', 'ні');
  const before = env.docProps.getProperty('mods');
  // таблиця версії 11: блок на «Старті» (рядки 24…), на «Налаштуваннях» його немає
  set.getRange('A42:C52').clearContent();
  start.getRange('A24:C25').setValues([['🧩 Модулі', '', 'так / ні'], ['Модуль', 'Увімкнено', 'Що робить']]);
  start.getRange(26, 1, MODS.length, 3).setValues(MODS.map(m => [m.name, c.modOn_(m.id) ? 'так' : 'ні', m.what]));
  start.getRange('B26:B34').setDataValidation(K(c, 'SpreadsheetApp').newDataValidation().requireValueInList(['так', 'ні']).setAllowInvalid(false).build());
  assert.strictEqual(c.startHasModules_(env.ss), true); assert.strictEqual(c.modOldOnStart_(env.ss), true);
  // поки блок на «Старті»: оновлення кроків його не чіпає, правка там працює
  c.startRefresh_(env.ss);
  assert.strictEqual(start.getRange('A24').getValue(), '🧩 Модулі');
  start.getRange(26 + MODS.findIndex(m => m.id === 'budgets'), 2).setValue('ні');
  c.mbsEdit({ range: start.getRange(26 + MODS.findIndex(m => m.id === 'budgets'), 2) });
  assert.strictEqual(c.modOn_('budgets'), false, 'стара клітинка на «Старті» ще працює');
  env.docProps.setProperty('schema_v', '11');
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.strictEqual(set.getRange('A42').getValue(), '🧩 Модулі');
  const v = set.getRange('A44:B52').getValues(), mod = id => v[MODS.findIndex(m => m.id === id)][1];
  assert.ok(mod('cal') === 'так' && mod('credit') === 'ні' && mod('budgets') === 'ні' && mod('tg') === 'так', JSON.stringify(v));
  assert.ok(start.getRange('A24:C36').getValues().every(r => r.join('') === '' || /Підтримати|безкоштовний/.test(r.join(''))), 'блок зі «Старту» прибрано');
  assert.strictEqual(start.getRange('B26').getDataValidation(), null, 'без списку «так / ні»');
  assert.ok(c.modOn_('cal') && !c.modOn_('credit') && !c.modOn_('budgets'), 'нічого не перемкнулось');
  assert.notStrictEqual(before, null);
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 12).done(env.ss), true);
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 6).done(env.ss), true, 'старі кроки не повторюються');
  // ще раз: нічого не робить
  env.docProps.setProperty('schema_v', '11');
  const files0 = env.files.length;
  c.schemaTick_(env.ss);
  assert.strictEqual(env.files.length, files0, 'крок уже зроблено: без копії');
});

t('panel step 5: a checkbox switches a module at once (sheets, rows, schedules), the block on «Налаштування» follows; «apply all» repairs; unknown id refused', () => {
  const env = setup(), c = env.ctx;
  let s = c.cnState();
  assert.strictEqual(s.mods.length, 9); assert.strictEqual(s.modsSeen, false);
  const r = c.cnModule('debts', true);
  assert.ok(r.lines.some(l => /Борги й розстрочки: увімкнено/.test(l)), r.lines.join(' | '));
  assert.strictEqual(c.modOn_('debts'), true);
  assert.strictEqual(st(env).getRange(row(env, '🤝 Борги й розстрочки'), 2).getValue(), 'так');
  assert.ok(r.s.mods.find(m => m.id === 'debts').on && r.s.modsSeen);
  c.cnModule('service', true);
  assert.strictEqual(env.ss.getSheetByName('Налаштування').isSheetHidden(), false, 'усі аркуші видно');
  c.cnModule('service', false);
  assert.strictEqual(env.ss.getSheetByName('Налаштування').isSheetHidden(), true);
  c.cnModule('tg', false);
  assert.ok(!env.triggers.some(x => x.getHandlerFunction() === 'mbsMorning'));
  const a = c.cnModsApply();
  assert.ok(a.lines.length >= 9, 'усі модулі заново');
  assert.throws(() => c.cnModule('nope', true), /Невідомий модуль/);
  const start = env.ss.getSheetByName('Старт').getRange('A6:C25').getValues();
  assert.ok(start.some(x => /^✅ 5$/.test(x[0]) && /Модулі/.test(x[1])), 'на «Старті»: ✅ 5 Модулі');
});

console.log(n + ' tests passed');
