// «🚀 Підключення»: панель з чотирма кроками (картка, Telegram, регулярні, Gemini), QR-коди, пошук регулярних, меню без «⚙️ Налаштування».
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);
const TOKEN = 'uTestToken' + 'x'.repeat(30), BOT = '123456789:' + 'B'.repeat(35);
const tz = 'Europe/Kiev';

function setup(state) {
  state = Object.assign({ accounts: [
    { id: 'acc1', currencyCode: 980, type: 'black', creditLimit: 5000000, balance: 4000000, iban: 'UA00TEST', maskedPan: ['000000******4321'] },
    { id: 'fop1', currencyCode: 980, type: 'fop', creditLimit: 0, balance: 100000, iban: 'UA00FOP9876', maskedPan: [] },
    { id: 'usd1', currencyCode: 840, type: 'black', creditLimit: 0, balance: 0, iban: 'UA00USD', maskedPan: [] }], items: [], updates: [], sent: [], statements: 0 }, state || {});
  const env = makeEnv({ mono: (url, o) => {
    if (/client-info$/.test(url)) return { code: 200, body: { name: 'Тест', accounts: state.accounts } };
    if (/\/personal\/statement\//.test(url)) { state.statements++; return { code: 200, body: state.items }; }
    if (/getMe$/.test(url)) return { code: 200, body: { ok: true, result: { username: 'family_budget_bot' } } };
    if (/getUpdates$/.test(url)) return { code: 200, body: { ok: true, result: state.updates } };
    if (/sendMessage$/.test(url)) { state.sent.push(JSON.parse(o.payload)); return { code: 200, body: { ok: true, result: {} } }; }
    return { code: 404, body: 'no' };
  } });
  env.ss.insertSheet('Аркуш1');
  env.ctx.startWizard();
  env.state = state;
  return env;
}
const cells = env => env.ss.getSheets().flatMap(s => Array.from(s.cells.values()).map(x => String(x.v) + String(x.f))).join('\n');

t('menu: «🚀 Підключення» opens the panel; no «⚙️ Налаштування» items (connect / disconnect live in the panel)', () => {
  const env = setup(), c = env.ctx;
  env.calls.menu = null; c.onOpen();
  assert.strictEqual(env.calls.menu[0], 'startWizard');
  ['connectMono', 'disconnectMono', 'connectTelegram', 'enableDailyImport', 'connectPrivat', 'calendarOn', 'removeProtection']
    .forEach(fn => assert.ok(env.calls.menu.indexOf(fn) < 0, fn));
  const html = env.calls.sidebars.pop().html;
  ['disconnectMono', 'disconnectPrivat', 'disconnectTelegram', 'calendarOff', 'removeProtection', 'connectPrivat', 'importStatements']
    .forEach(fn => { assert.ok(html.indexOf("'" + fn + "'") >= 0, fn); assert.strictEqual(typeof c[fn], 'function', fn); });
  (html.match(/run\('(cn\w+)'/g) || []).map(x => x.slice(5, -1)).forEach(fn => assert.strictEqual(typeof c[fn], 'function', fn));
  assert.ok(!/^function /m.test(html), 'функції панелі з відступом: не плутаються з функціями скрипта');
  assert.doesNotThrow(() => new vm.Script(html.split('<script>')[1].split('</script>')[0]), 'скрипт панелі без синтаксичних помилок');
  (html.match(/google\.script\.run(?:\.with\w+\([^)]*\))*\.(\w+)\(/g) || []).map(x => x.match(/\.(\w+)\($/)[1]).filter(fn => !/^with/.test(fn))
    .forEach(fn => assert.strictEqual(typeof c[fn], 'function', 'з панелі: ' + fn));
  ['id="ap"', 'id="ac"', 'id="af"', 'id="ad"', 'id="ao"', 'id="ar"', 'Скопіювати для Gemini', 'advSkillCopied'].forEach(x => assert.ok(html.indexOf(x) >= 0, x));
});

t('step 1, Mono: token checked without saving; picked accounts saved (main card + FOP), card type and modules from the bank, schedule on; token never in cells or panel state', () => {
  const env = setup(), c = env.ctx;
  assert.throws(() => c.cnMonoCheck('коротко'), /не схоже на токен/);
  const list = c.cnMonoCheck(TOKEN);
  assert.strictEqual(list.length, 2, 'лише гривневі');
  assert.strictEqual(list[0].label, 'Mono Чорна *4321'); assert.strictEqual(list[0].limit, 50000); assert.strictEqual(list[0].pick, true);
  assert.ok(list[1].fop && list[1].pick && list[1].label === 'Mono ФОП *9876');
  assert.strictEqual(env.userProps.getProperty('mono_token'), null, 'перевірка нічого не зберігає');
  assert.throws(() => c.cnMonoSave(TOKEN, [], ''), /хоча б один/);
  const r = c.cnMonoSave(TOKEN, ['acc1', 'fop1'], '');
  assert.strictEqual(r.credit, true); assert.strictEqual(r.fop, true);
  assert.strictEqual(env.userProps.getProperty('mono_token'), TOKEN);
  const acc = env.ss.getSheetByName('Рахунки').getRange('A5:D6').getValues().map(x => x[0] + '|' + x[3]);
  assert.strictEqual(JSON.stringify(acc), JSON.stringify(['Mono Чорна *4321|основна картка', 'Mono ФОП *9876|рахунок ФОП']));
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('B7').getValue(), 'кредитна з лімітом');
  const mods = JSON.parse(env.docProps.getProperty('mods'));
  assert.strictEqual(mods.credit, true); assert.strictEqual(mods.fop, true);
  assert.ok(env.ss.getSheetByName('ФОП') && env.ss.getSheetByName('ФОП').isSheetHidden(), 'модуль ФОП застосовано; аркуш схований разом з рештою');
  assert.ok(env.triggers.filter(x => x.getHandlerFunction() === 'mbsImport').length === 3, 'імпорт за розкладом');
  const s = c.cnState();
  assert.ok(s.bank.fopAcc && s.bank.schedule && s.bank.mono.length === 2);
  assert.ok(JSON.stringify(s).indexOf(TOKEN) < 0 && cells(env).indexOf(TOKEN) < 0, 'токен ні в клітинках, ні в панелі');
  c.cnFop('2', '');
  assert.strictEqual(JSON.stringify(env.ss.getSheetByName('ФОП').getRange('A14:B14').getValues()[0]), JSON.stringify(['Я', '2']));
  c.cnSchedule(false);
  assert.strictEqual(env.triggers.filter(x => x.getHandlerFunction() === 'mbsImport').length, 0);
  const imp = c.cnImport();
  assert.ok(typeof imp.text === 'string' && imp.own === -10000 && imp.drafts === 0, JSON.stringify(imp));
});

t('step 1 without a bank: card type switches the credit module; income day 1…31', () => {
  const env = setup(), c = env.ctx;
  c.cnCard(false);
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('B7').getValue(), 'дебетова');
  assert.strictEqual(JSON.parse(env.docProps.getProperty('mods')).credit, false);
  assert.strictEqual(c.cnCard(true).bank.credit, true);
  assert.strictEqual(c.cnBday(20).bank.bday, 20);
  assert.throws(() => c.cnSchedule(true), /Спершу підключіть/);
});

t('step 2, Telegram: token → bot name, deep links and QR codes; panel polls: /start with our code (also in a group) connects, sends the first message, module on', () => {
  const env = setup(), c = env.ctx, st = env.state;
  assert.throws(() => c.cnTgToken('abc'), /не схоже на токен бота/);
  const r = c.cnTgToken(BOT);
  const code = env.userProps.getProperty('tg_link');
  assert.ok(/^[0-9a-f]{12}$/.test(code));
  assert.strictEqual(r.private, 'https://t.me/family_budget_bot?start=' + code);
  assert.strictEqual(r.group, 'https://t.me/family_budget_bot?startgroup=' + code);
  assert.ok(/^<svg viewBox="0 0 \d+ \d+"/.test(r.qr) && r.qr.indexOf(BOT) < 0, 'QR без токена');
  assert.ok(c.cnState().tg.link.qr, 'панель після перезапуску бачить QR');
  assert.strictEqual(c.cnTgPoll().ok, false);
  st.updates = [{ update_id: 7, message: { date: 1, text: '/start@family_budget_bot ' + code, chat: { id: -100500, type: 'supergroup', title: 'Сім\'я' } } }];
  const p = c.cnTgPoll();
  assert.ok(p.ok && /група «Сім'я»/.test(p.chat), JSON.stringify(p));
  assert.strictEqual(env.userProps.getProperty('tg_chat'), '-100500');
  assert.ok(st.sent.some(m => /підключено/.test(m.text)));
  assert.strictEqual(JSON.parse(env.docProps.getProperty('mods')).tg, true);
  assert.ok(env.triggers.some(x => x.getHandlerFunction() === 'mbsMorning'));
  assert.ok(cells(env).indexOf(BOT) < 0 && JSON.stringify(c.cnState()).indexOf(BOT) < 0);
});

t('Telegram /start finder: our code first; otherwise a fresh /start in a private chat; old /start and other chats ignored', () => {
  const c = setup().ctx, msg = (text, type, date, id) => ({ message: { text: text, date: date, chat: { id: id, type: type } } });
  assert.strictEqual(c.cnTgFind_([msg('/start', 'private', 500, 1), msg('/start abc', 'private', 10, 2)], 'abc', 100).chat.id, 2);
  assert.strictEqual(c.cnTgFind_([msg('/start', 'private', 500, 1)], 'abc', 100).chat.id, 1);
  assert.strictEqual(c.cnTgFind_([msg('/start', 'private', 50, 1)], 'abc', 100), null, 'старий /start');
  assert.strictEqual(c.cnTgFind_([msg('/start', 'group', 500, 1), msg('привіт', 'private', 500, 3)], 'abc', 100), null);
  assert.strictEqual(c.cnTgFind_([msg('/start xyz', 'channel', 500, 1)], 'xyz', 100), null);
});

t('regular finder: monthly payments and income found (amount ±25 %, day ±5); daily coffee, varying sums and existing rules skipped', () => {
  const c = setup().ctx;
  const ops = [];
  const add = (key, amt, desc, cat) => ops.push({ key: key, amt: amt, desc: desc, cat: cat });
  ['2026-08', '2026-09', '2026-10'].forEach((m, i) => {
    add(m + '-12', -299, 'Netflix.com 12' + i, 'Підписки');
    add(m + '-05', -12000 - i * 300, 'Переказ Олені К.', '↔ Переказ (уточніть)');
    add(m + '-25', 42000, 'Зарплата ТОВ Ромашка', 'Зарплата');
    add(m + '-10', -250, 'Київстар', 'Інтернет і зв\'язок');
    for (let d = 1; d <= 20; d++) add(m + '-' + String(d).padStart(2, '0'), -75, 'Кавʼярня', 'Кава та перекуси');
    add(m + '-15', i === 1 ? -5000 : -800, 'Розетка', 'Дрібні покупки');
  });
  add('2026-10-13', -299, 'Netflix.com 99', 'Підписки');
  const r = c.regSuggest_(ops, [{ name: 'Мобільний', cat: 'Інтернет і зв\'язок', amt: 260 }]);
  const names = r.map(x => x.name);
  assert.ok(names.indexOf('Netflix.com 99') >= 0 || names.some(x => /^Netflix/.test(x)), names.join());
  const nf = r.find(x => /^Netflix/.test(x.name));
  assert.strictEqual(nf.amount, 299); assert.strictEqual(nf.day, 12); assert.strictEqual(nf.cat, 'Підписки'); assert.strictEqual(nf.income, false);
  const rent = r.find(x => /Олені/.test(x.name));
  assert.ok(rent && rent.cat === '' && rent.day === 5 && rent.amount === 12300, JSON.stringify(rent));
  const pay = r.find(x => x.income);
  assert.ok(pay && pay.cat === 'Зарплата' && pay.day === 25 && pay.amount === 42000);
  assert.ok(!names.some(x => /Кав|Розетка|Київстар/.test(x)), 'кава щодня, різні суми, уже є правило: ' + names.join());
  assert.strictEqual(r[0].income, true, 'від більших сум');
});

t('step 3: search in the journal, add picked ones to «Регулярні» (plans appear), wrong category skipped; deep search reads Mono for 2 months without writing the journal', () => {
  const env = setup(), c = env.ctx, j = env.ss.getSheetByName('Журнал');
  const today = env.Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), P = k => env.Utilities.parseDate(k, tz, 'yyyy-MM-dd');
  const m1 = c.prevYm_(today.slice(0, 7)), m2 = c.prevYm_(m1);
  j.getRange('A5:E6').setValues([[P(m2 + '-12'), 299, 'Підписки', 'Netflix', '✅ факт'], [P(m1 + '-12'), 299, 'Підписки', 'Netflix', '✅ факт']]);
  [5, 6].forEach(r => j.put(r, 9, 'Витрати', j.cell(r, 9).f));
  const r = c.cnRegSuggest(false);
  assert.strictEqual(r.items.length, 1); assert.strictEqual(r.items[0].name, 'Netflix'); assert.strictEqual(r.canDeep, false, 'Mono не підключено');
  const add = c.cnRegAdd([{ name: 'Netflix', cat: 'Підписки', amount: 299, day: 12, per: 'щомісяця' }, { name: 'Щось', cat: 'Немає такої', amount: 10, day: 1 }]);
  assert.strictEqual(add.added, 1); assert.strictEqual(JSON.stringify(add.skipped), JSON.stringify(['Щось']));
  assert.ok(add.plans >= 1);
  assert.strictEqual(env.ss.getSheetByName('Регулярні').getRange('A7').getValue(), 'Netflix', 'під «Заплати собі» й «Інвестиції»');
  assert.ok(c.readJournal_(env.ss).some(o => /^\[Регулярний\] Netflix: /.test(o.d)));
  assert.strictEqual(c.cnRegSuggest(false).items.length, 0, 'уже на аркуші');
  // глибше: Mono, 2 запити по 31 день, нічого не пишеться в журнал
  c.cnMonoSave(TOKEN, ['acc1'], '');
  const at = k => Math.floor(env.Utilities.parseDate(k + ' 10:00:00', tz, 'yyyy-MM-dd HH:mm:ss').getTime() / 1000);
  env.state.items = [{ id: 'k1', time: at(m2 + '-07'), amount: -25000, description: 'Київстар', mcc: 4814 }, { id: 'k2', time: at(m1 + '-07'), amount: -25000, description: 'Київстар', mcc: 4814 }];
  const before = c.readJournal_(env.ss).filter(o => !o.blank).length, s0 = env.state.statements;
  const deep = c.cnRegSuggest(true);
  assert.strictEqual(env.state.statements - s0, 2, 'два вікна по 31 дню');
  assert.ok(deep.items.some(x => x.name === 'Київстар' && x.amount === 250 && x.day === 7), JSON.stringify(deep.items));
  assert.strictEqual(c.readJournal_(env.ss).filter(o => !o.blank).length, before, 'журнал не змінився');
});

t('step 4 and finish: snapshot on Drive with the skill texts, module on; «Готово» marks the checklist; demo on and off from the panel', () => {
  const env = setup(), c = env.ctx;
  const a = c.cnAdvisor();
  assert.ok(/^https:\/\/docs\.google\.com\//.test(a.doc) || a.doc, JSON.stringify(a));
  assert.strictEqual(a.gemini, 'https://gemini.google.com/');
  assert.ok(/Monitoring Budgets/.test(a.instructions) && a.starters.length >= 3);
  assert.strictEqual(JSON.parse(env.docProps.getProperty('mods')).advisor, true);
  assert.ok(c.cnState().adv.doc);
  assert.strictEqual(c.cnDemo(true).demo, true);
  assert.ok(c.readJournal_(env.ss).some(o => /^\[демо\]/.test(o.d)));
  assert.strictEqual(c.cnDemo(false).demo, false);
  const f = c.cnFinish();
  assert.ok(env.docProps.getProperty('wiz_done') && f.need > 0);
  assert.ok(/основного: \d+ з 6/.test(env.ss.getSheetByName('⚡ Швидке підключення').getRange('A5').getValue()), '«Готово»: стан оновлено');
});

t('QR: version grows with the text, finder patterns in three corners, quiet zone in the SVG', () => {
  const c = setup().ctx;
  const m = c.qrMatrix_('https://t.me/BotFather'), big = c.qrMatrix_('https://t.me/some_very_long_bot_name_budget_bot?startgroup=0123456789abcdef');
  assert.strictEqual(m.length, 25); assert.ok(big.length > m.length);
  const n = m.length;
  [[0, 0], [0, n - 7], [n - 7, 0]].forEach(p => {
    assert.strictEqual(m[p[0]].slice(p[1], p[1] + 7), '1111111');
    assert.strictEqual(m[p[0] + 1].slice(p[1], p[1] + 7), '1000001');
    assert.strictEqual(m[p[0] + 3].slice(p[1], p[1] + 7), '1011101');
  });
  assert.ok(/viewBox="0 0 33 33"/.test(c.qrSvg_('https://t.me/BotFather')), '25 + 2 × 4');
});

console.log(n + ' tests passed');
