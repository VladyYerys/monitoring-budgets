// Імпорт з Mono: чернетки, без дублів з вашими рядками, комісії, перекази, контрольні точки, журнал «Банк»; токен лише в UserProperties.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);
const TOKEN = 'uTestToken' + 'x'.repeat(30);

function nowKey(env) { return env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd'); }
function sec(env, key, hhmm) { return Math.floor(env.Utilities.parseDate(key + ' ' + hhmm + ':00', 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss').getTime() / 1000); }

function setup(items, accounts) {
  const state = { items: items || [], accounts: accounts || [{ id: 'acc1', currencyCode: 980, type: 'black', creditLimit: 5000000, balance: 4000000,
    iban: 'UA00TEST', maskedPan: ['000000******4321'] }] };
  const env = makeEnv({ mono: (url) => {
    if (/client-info$/.test(url)) return { code: 200, body: { name: 'Тест', accounts: state.accounts } };
    if (/\/personal\/statement\//.test(url)) return { code: 200, body: state.items };
    return { code: 404, body: 'no' };
  }, promptText: TOKEN });
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  env.state = state;
  return env;
}

t('connect: token only in user properties, card saved as «Mono Чорна *4321» (last 4 digits), nothing in cells', () => {
  const env = setup();
  env.ctx.connectMono();
  assert.strictEqual(env.userProps.getProperty('mono_token'), TOKEN);
  assert.strictEqual(JSON.parse(env.userProps.getProperty('mono_accs'))[0].last4, '4321');
  const acc = env.ss.getSheetByName('Рахунки').getRange('A5:D5').getValues()[0];
  assert.strictEqual(JSON.stringify([acc[0], acc[1], acc[3]]), JSON.stringify(['Mono Чорна *4321', 'Mono', 'основна картка']));
  assert.strictEqual(env.ctx.savedCardLabel_(), 'Mono Чорна *4321');
  const all = env.ss.getSheets().flatMap(s => Array.from(s.cells.values()).map(x => String(x.v) + String(x.f)));
  assert.ok(all.every(v => v.indexOf(TOKEN) < 0 && v.indexOf('000000') < 0));
  assert.ok(/Mono підключено/.test(env.calls.dialogs.pop().t));
});

t('import: drafts with rule categories, fee as its own row, transfer red, your own row not duplicated, checkpoints, bank log', () => {
  const env = setup(), today = nowKey(env), c = env.ctx;
  const j = env.ss.getSheetByName('Журнал');
  j.getRange('A12:E12').setValues([[env.Utilities.parseDate(today, 'Europe/Kiev', 'yyyy-MM-dd'), 250, 'Кафе й ресторани', 'обід', '✅ факт']]);
  env.state.items = [
    { id: 'm1', time: sec(env, today, '09:10'), description: 'АТБ', mcc: 5411, amount: -12000, balance: 3990000, commissionRate: 0 },
    { id: 'm2', time: sec(env, today, '10:00'), description: 'Обід', mcc: 5812, amount: -25000, balance: 3965000, commissionRate: 0 },
    { id: 'm3', time: sec(env, today, '11:00'), description: 'Переказ на картку', mcc: 4829, amount: -100000, balance: 3860000, commissionRate: 500 },
    { id: 'm4', time: sec(env, today, '12:00'), description: 'Невідомий магазин', mcc: 7999, amount: -3000, balance: 3857000, commissionRate: 0 },
  ];
  env.state.accounts[0].balance = 3857000;
  c.connectMono();
  const r = c.importBank_(env.ss);
  assert.strictEqual(r.matched, 1, 'обід уже був у журналі');
  assert.strictEqual(r.drafts, 4, 'АТБ, переказ, його комісія, невідомий магазин');
  assert.strictEqual(r.transfers, 1); assert.strictEqual(r.uncategorized, 1); assert.strictEqual(r.fees, 1);
  const rows = c.readJournal_(env.ss).filter(o => !o.blank);
  const by = id => rows.find(o => o.bank[0] === id);
  assert.strictEqual(by('m1').c, 'Продукти'); assert.strictEqual(by('m1').e, '🏦 чернетка'); assert.strictEqual(by('m1').b, 120);
  assert.strictEqual(by('m2').d, 'обід', 'ваш рядок лишився, отримав номер операції');
  assert.strictEqual(by('m3').c, '↔ Переказ (уточніть)');
  assert.strictEqual(by('m3#fee').c, 'Комісії й відсотки банку'); assert.strictEqual(by('m3#fee').f, '🏦 Банківські комісії');
  assert.strictEqual(by('m4').c, '');
  assert.ok(rows.every(o => o.ord !== '' && o.ord !== null), 'черговість запису є в усіх рядків');
  const cp = env.ss.getSheetByName('Залишки банку').getRange('A5:C7').getValues();
  assert.strictEqual(cp.filter(x => x[0]).length, 3);
  assert.strictEqual(cp[2][2], 50000, 'ліміт');
  const log = env.ss.getSheetByName('Банк');
  assert.ok(log.getLastRow() >= 9 && /нових чернеток 4/.test(log.getRange('D2').getValue()));
  assert.strictEqual(log.getRange('F2').getValue(), 'Mono Чорна *4321');
  // другий імпорт: нічого нового, без дублів
  const r2 = c.importBank_(env.ss);
  assert.strictEqual(r2.drafts, 0);
  assert.strictEqual(c.readJournal_(env.ss).filter(o => !o.blank).length, rows.length);
});

t('bank cancelled a pending payment: the draft disappears; a pending «скасування» is not imported', () => {
  const env = setup(), today = nowKey(env), c = env.ctx;
  env.state.items = [{ id: 'p1', time: sec(env, today, '08:00'), description: 'Магазин', mcc: 5311, amount: -5000, balance: 3995000, hold: true }];
  env.state.accounts[0].balance = 3995000;
  c.connectMono(); c.importBank_(env.ss);
  assert.ok(c.readJournal_(env.ss).some(o => o.bank[0] === 'p1'));
  env.state.items = [{ id: 'p2', time: sec(env, today, '08:05'), description: 'Скасування. Магазин', mcc: 5311, amount: 5000, hold: true }];
  env.state.accounts[0].balance = 4000000;
  const r = c.importBank_(env.ss);
  assert.strictEqual(r.vanished, 1); assert.strictEqual(r.skipped, 1);
  assert.ok(!c.readJournal_(env.ss).some(o => o.bank[0] === 'p1' || o.bank[0] === 'p2'));
});

t('debit card: the account without a credit limit is chosen; scheduled import retries and reports to Telegram without leaking the bot token', () => {
  const accounts = [{ id: 'd1', currencyCode: 980, type: 'white', creditLimit: 0, balance: 150000, maskedPan: ['1111'] }];
  const env = setup([], accounts), c = env.ctx;
  env.ss.getSheetByName('Налаштування').getRange('B7').setValue('дебетова');
  c.connectMono();
  assert.strictEqual(JSON.parse(env.userProps.getProperty('mono_accs'))[0].id, 'd1');
  c.enableDailyImport();
  const imp = env.triggers.filter(t => t.getHandlerFunction() === 'mbsImport');
  assert.strictEqual(imp.length, 3); assert.ok(imp.every(t => t.spec.minute === 45));
  const bot = '123456:' + 'Z'.repeat(35);
  env.userProps.setProperties({ tg_token: bot, tg_chat: '42' });
  env.state.accounts = 'broken';
  const sent = [];
  env.ctx.UrlFetchApp.fetch = (url, o) => {
    env.calls.fetch.push(url);
    if (/api\.telegram\.org/.test(url)) { sent.push(JSON.parse(o.payload).text); return { getResponseCode: () => 200, getContentText: () => '{"ok":true,"result":{}}' }; }
    return { getResponseCode: () => 500, getContentText: () => 'oops' };
  };
  c.mbsImport();
  assert.ok(env.triggers.some(t => t.getHandlerFunction() === 'mbsImportRetry'), 'повтор заплановано');
  assert.ok(/Імпорт з банку не вдався/.test(sent[0]) && /Mono/.test(sent[0]) && sent[0].indexOf(bot) < 0, sent[0]);
  const hist = JSON.parse(env.docProps.getProperty('imp_hist'));
  assert.strictEqual(hist[0].ok, false);
});

t('pure: expandOps_ learns whether the fee is inside the amount; importWindowFrom_ overlaps one day and stays within 31 days', () => {
  const c = makeEnv().ctx;
  const a = c.expandOps_([{ id: 'a', time: 1, amount: -1000, balance: 10000 }, { id: 'b', time: 2, amount: -10100, commissionRate: 100, balance: -100 }], '');
  assert.strictEqual(a.learned, 'так');
  assert.strictEqual(JSON.stringify(a.ops.map(o => o.amount)), JSON.stringify([-1000, -10000, -100]));
  const now = 1e9;
  assert.strictEqual(c.importWindowFrom_(now, now - 3 * 86400, now - 3600), now - 3 * 86400);
  assert.strictEqual(c.importWindowFrom_(now, now - 90 * 86400, 0), now - 2682000 + 60);
});

t('family: two people, each with own Mono in own account; the second card is «інша картка» (no checkpoints), drafts get «Хто»; FOP account of Mono goes to «рахунок ФОП»', () => {
  const k = env => env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  const accounts = [{ id: 'olena', currencyCode: 980, type: 'black', creditLimit: 3000000, balance: 2500000, maskedPan: ['1111'] },
    { id: 'olena-fop', currencyCode: 980, type: 'fop', creditLimit: 0, balance: 900000, iban: 'UA' + '7'.repeat(27) }];
  const env = setup([], accounts), c = env.ctx, today = k(env);
  const set = env.ss.getSheetByName('Налаштування');
  set.getRange('B13').setValue('сім\'я'); set.getRange('B14:B15').setValues([['Олена'], ['Іван']]);
  env.ctx.SpreadsheetApp.getUi();
  const answers = [TOKEN, '1, 2', '1'];
  env.calls.prompts.length = 0;
  const ui = env.ctx.SpreadsheetApp.getUi(), prompt0 = ui.prompt;
  ui.prompt = (t, m) => { const a = answers.shift(); env.calls.prompts.push([t, m]); return { getSelectedButton: () => 'OK', getResponseText: () => a }; };
  c.connectMono();
  const list = c.accList_(env.ss);
  assert.strictEqual(JSON.stringify(list.map(a => [a.name, a.kind, a.who])), JSON.stringify([['Mono Чорна *1111', 'основна картка', 'Олена'], ['Mono ФОП *7777', 'рахунок ФОП', 'Олена']]));
  env.state.items = [{ id: 'o1', time: sec(env, today, '09:00'), description: 'Сільпо', mcc: 5411, amount: -30000, balance: 2470000 }];
  c.importBank_(env.ss);
  assert.ok(env.calls.sleep >= 61000, 'між виписками двох рахунків: пауза 61 с');
  // Іван: свій акаунт Google, свій токен, своя картка
  Object.keys(env.userProps.getProperties()).forEach(x => env.userProps.deleteProperty(x));
  env.state.accounts = [{ id: 'ivan', currencyCode: 980, type: 'white', creditLimit: 0, balance: 400000, maskedPan: ['2222'] }];
  answers.push(TOKEN.replace('x', 'y'), '2');
  c.connectMono();
  ui.prompt = prompt0;
  assert.strictEqual(c.accList_(env.ss)[2].kind, 'інша картка');
  env.state.items = [{ id: 'i1', time: sec(env, today, '10:00'), description: 'АЗС WOG', mcc: 5541, amount: -120000, balance: 280000 }];
  const cps0 = env.ss.getSheetByName('Залишки банку').getRange('A5:C10').getValues().filter(r => r[0]).length;
  const r = c.importBank_(env.ss);
  assert.strictEqual(r.checkpoints.length, 0, 'картка Івана: без контрольних точок');
  assert.strictEqual(env.ss.getSheetByName('Залишки банку').getRange('A5:C10').getValues().filter(r => r[0]).length, cps0);
  const rows = c.readJournal_(env.ss).filter(o => !o.blank);
  const ivan = rows.find(o => o.bank[0] === 'i1'), olena = rows.find(o => o.bank[0] === 'o1');
  assert.strictEqual(JSON.stringify([ivan.g, ivan.acct, ivan.kind, ivan.c]), JSON.stringify(['Іван', 'Mono Біла *2222', 'інша картка', 'Пальне']));
  assert.strictEqual(JSON.stringify([olena.g, olena.kind]), JSON.stringify(['Олена', 'основна картка']));
});

console.log(n + ' tests passed');
