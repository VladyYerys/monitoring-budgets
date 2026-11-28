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
  assert.strictEqual(by('m1').c, 'Продукти'); assert.strictEqual(by('m1').e, '✅ факт', 'категорію визначено: одразу факт');
  assert.strictEqual(by('m3').e, '🏦 чернетка'); assert.strictEqual(by('m4').e, '🏦 чернетка'); assert.strictEqual(by('m3#fee').e, '✅ факт');
  assert.strictEqual(r.facts, 2, 'АТБ і комісія'); assert.strictEqual(by('m1').b, 120);
  assert.strictEqual(by('m2').d, 'обід', 'ваш рядок лишився, отримав номер операції');
  assert.strictEqual(by('m3').c, '↔ Переказ (уточніть)');
  assert.strictEqual(by('m3#fee').c, 'Комісії й відсотки банку'); assert.strictEqual(by('m3#fee').f, '🏦 Банківські комісії');
  assert.strictEqual(by('m4').c, '');
  assert.ok(rows.every(o => o.ord !== '' && o.ord !== null), 'черговість запису є в усіх рядків');
  const cp = env.ss.getSheetByName('Залишки банку').getRange('A5:C7').getValues();
  assert.strictEqual(cp.filter(x => x[0]).length, 3);
  assert.strictEqual(cp[2][2], 50000, 'ліміт');
  const log = env.ss.getSheetByName('Банк');
  assert.ok(log.getLastRow() >= 9 && /нових операцій 4/.test(log.getRange('D2').getValue()) && /одразу «✅ факт» 2/.test(log.getRange('D2').getValue()));
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

t('Mono «VPN_BLOCKED» (Google servers look like a VPN to Mono): two quiet retries; still blocked → a plain explanation, not raw JSON; other errors show Mono\'s own text', () => {
  let h = null;
  const env = makeEnv({ mono: url => h(url) }), c = env.ctx, blocked = { code: 400, body: { errCode: 'VPN_BLOCKED', errTitle: 'З\'єднання заблоковано', errText: 'Змініть IP-адресу або вимкніть VPN', errButton: 'Зрозуміло' } };
  let left = 1;
  h = () => (left-- > 0 ? blocked : { code: 200, body: { name: 'Тест', accounts: [] } });
  const s0 = env.calls.sleep;
  assert.strictEqual(c.monoFetch_(TOKEN, '/personal/client-info').name, 'Тест', 'вдруге пройшло');
  assert.strictEqual(env.calls.sleep - s0, 5000);
  h = () => blocked;
  const f0 = env.calls.fetch.length;
  assert.throws(() => c.monoFetch_(TOKEN, '/personal/client-info'), e => /сервери Google/.test(e.message) && /Зачекайте 5 хвилин, оновіть сторінку/.test(e.message) && /10 і 30 хв/.test(e.message) && !/errCode|\{|400/.test(e.message));
  assert.strictEqual(env.calls.fetch.length - f0, 3, 'запит і два повтори');
  h = () => ({ code: 400, body: { errorDescription: 'Period must be no more than 31 days' } });
  assert.throws(() => c.monoFetch_(TOKEN, '/personal/statement/x/1/2'), e => e.message === 'Mono відповів кодом 400: Period must be no more than 31 days');
});

/* ---------- 🗓 імпорт за період ---------- */

// Mono з вікнами: віддає лише операції в [from, to] запиту; запити й Telegram записуються
function periodSetup(items, extra) {
  const env = setup(items), c = env.ctx, reqs = [], sent = [];
  env.reqs = reqs; env.sent = sent; env.fail = null;
  env.ctx.UrlFetchApp.fetch = (url, o) => {
    env.calls.fetch.push(url);
    const R = (code, body) => ({ getResponseCode: () => code, getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body)) });
    if (/api\.telegram\.org/.test(url)) { sent.push(JSON.parse(o.payload).text); return R(200, { ok: true, result: {} }); }
    if (/client-info$/.test(url)) return R(200, { name: 'Тест', accounts: env.state.accounts });
    const m = url.match(/\/personal\/statement\/([^/]+)\/(\d+)\/(\d+)$/);
    if (m) {
      if (env.fail) { const f = env.fail; if (f.times !== undefined && --f.times < 0) env.fail = null; else return R(f.code, f.body); }
      const from = +m[2], to = +m[3]; reqs.push([m[1], from, to]);
      return R(200, env.state.items.filter(it => it.time >= from && it.time <= to).sort((a, b) => b.time - a.time));
    }
    return R(404, 'no');
  };
  Object.assign(env.opts, extra || {});
  return env;
}
const daysAgo = (env, n, hhmm) => sec(env, env.ctx.addDays_(nowKey(env), -n), hhmm || '12:00');
const run = (env, name) => env.triggers.filter(t => t.getHandlerFunction() === name).length;

t('🗓 period: 3 months → first 31 days at once, the rest by a one-minute schedule, window by window (≤ 31 days, contiguous, newest first); no duplicates; checkpoints, start date and «last import» untouched; summary in status and Telegram', () => {
  const env = periodSetup([]), c = env.ctx, today = nowKey(env);
  const D = n => env.Utilities.parseDate(c.addDays_(today, -n), 'Europe/Kiev', 'yyyy-MM-dd');
  env.state.items = [
    { id: 'p1', time: daysAgo(env, 2, '09:00'), description: 'АТБ', mcc: 5411, amount: -15000, balance: 3985000 },
    { id: 'p2', time: daysAgo(env, 20), description: 'Сільпо', mcc: 5411, amount: -40000, balance: 4000000 },
    { id: 'p3', time: daysAgo(env, 45), description: 'Обід', mcc: 5812, amount: -32000, balance: 4040000 },
    { id: 'p4', time: daysAgo(env, 70), description: 'Переказ на картку', mcc: 4829, amount: -500000, balance: 4072000 },
    { id: 'p5', time: daysAgo(env, 80), description: 'Невідомий магазин', mcc: 7999, amount: -9900, balance: 4572000 },
    { id: 'old', time: daysAgo(env, 120), description: 'АТБ', mcc: 5411, amount: -1000, balance: 4581900 },
  ];
  c.connectMono();
  c.importBank_(env.ss);                                                   // щоденний імпорт: лише останні дні (p1)
  const lastTo = env.docProps.getProperty('bank_last_to:' + c.accList_(env.ss)[0].key);
  const j = env.ss.getSheetByName('Журнал');
  j.getRange('A20:E20').setValues([[D(44), 320, 'Кафе й ресторани', 'обід з колегами', '✅ факт']]);   // ваш рядок: та сама сума, ±1 день
  const start0 = env.ss.getSheetByName('Налаштування').getRange('B4').getValue();
  const cps0 = JSON.stringify(env.ss.getSheetByName('Залишки банку').getRange('A5:C12').getValues());
  env.userProps.setProperties({ tg_token: '123456:' + 'Z'.repeat(35), tg_chat: '42' });
  env.reqs.length = 0;
  env.opts.promptText = ['3'];
  c.monoPeriodNow();
  const job = JSON.parse(env.userProps.getProperty('period_job'));
  assert.strictEqual(job.total, 3); assert.strictEqual(job.done, 1); assert.strictEqual(job.fromKey, c.addDays_(c.monthsBack_(today, 3), 1));
  assert.ok(!/\*{6}|UA00TEST/.test(env.userProps.getProperty('period_job')), 'без номера картки й IBAN (лише назва з 4 цифрами, як на «Рахунках»)');
  assert.strictEqual(run(env, 'mbsPeriod'), 1); assert.strictEqual(env.triggers.find(t => t.getHandlerFunction() === 'mbsPeriod').spec.after, 65000);
  const dlg = env.calls.dialogs.pop().html.replace(/<[^>]+>/g, '');
  assert.ok(/вікон по 31 дню: 3/.test(dlg) && /нових 1/.test(dlg) && /Решта \(2 вікна\) сама у фоні/.test(dlg) && /Telegram/.test(dlg) && /Зупинити: цей самий пункт меню/.test(dlg), dlg);
  let rows = () => c.readJournal_(env.ss).filter(o => !o.blank), by = id => rows().filter(o => o.bank[0] === id);
  assert.strictEqual(by('p1').length, 1, 'p1 уже був: без дубля'); assert.strictEqual(by('p2').length, 1); assert.strictEqual(by('p3').length, 0, 'старше: ще ні');
  c.mbsPeriod();
  assert.strictEqual(by('p3').length, 1); assert.strictEqual(by('p3')[0].d, 'обід з колегами', 'ваш рядок отримав номер операції, без дубля');
  assert.strictEqual(rows().filter(o => Number(o.b) === 320).length, 1);
  assert.strictEqual(run(env, 'mbsPeriod'), 1, 'наступне вікно заплановано');
  c.mbsPeriod();
  assert.strictEqual(by('p4')[0].c, '↔ Переказ (уточніть)'); assert.strictEqual(by('p4')[0].e, '🏦 чернетка');
  assert.strictEqual(by('p5')[0].e, '🏦 чернетка'); assert.strictEqual(by('old').length, 0, 'поза періодом');
  assert.strictEqual(env.userProps.getProperty('period_job'), null); assert.strictEqual(run(env, 'mbsPeriod'), 0, 'усе: розкладу немає');
  // вікна: 3 запити, кожен ≤ 31 доби, суцільно від початку періоду до «зараз», новіші спершу
  const st = env.reqs.filter(r => r[0] === 'acc1');
  assert.strictEqual(st.length, 3);
  st.forEach((r, i) => { assert.ok(r[2] - r[1] <= 31 * 86400, 'вікно ≤ 31 доби'); if (i) assert.strictEqual(r[2], st[i - 1][1], 'вікна стикуються'); });
  assert.strictEqual(st[2][1], sec(env, job.fromKey, '00:00'), 'з початку дня');
  const last = JSON.parse(env.userProps.getProperty('period_last'));
  assert.strictEqual(last.ok, true); assert.strictEqual(last.done, 3);
  assert.strictEqual(last.rep.drafts, 3, 'p2, p4, p5 (p1 уже був, p3: ваш рядок)'); assert.strictEqual(last.rep.matched, 1);
  assert.ok(env.sent.length === 1 && /Імпорт за період/.test(env.sent[0]) && /вікон 3 з 3/.test(env.sent[0]) && env.sent[0].indexOf('Z'.repeat(35)) < 0, env.sent.join('\n'));
  assert.ok(c.bankStatusLines_(env.ss).some(l => /Імпорт за період .*✓ нових операцій/.test(l)));
  // не змінено: старт обліку, контрольні точки, вікно щоденного імпорту
  assert.strictEqual(String(env.ss.getSheetByName('Налаштування').getRange('B4').getValue()), String(start0));
  assert.strictEqual(JSON.stringify(env.ss.getSheetByName('Залишки банку').getRange('A5:C12').getValues()), cps0);
  assert.strictEqual(env.docProps.getProperty('bank_last_to:' + c.accList_(env.ss)[0].key), lastTo);
  // удруге той самий період: нічого нового
  env.opts.promptText = ['3']; c.monoPeriodNow(); c.mbsPeriod(); c.mbsPeriod();
  assert.strictEqual(JSON.parse(env.userProps.getProperty('period_last')).rep.drafts, 0, 'повтор: без дублів');
});

t('🗓 period: menu while running shows progress and stops on «Так»; disconnect stops too; 1 month is one window; Mono 4xx stops at once, 5xx is retried then stops after 3 fails; busy script: tries again in a minute', () => {
  const env = periodSetup([]), c = env.ctx;
  env.state.items = [{ id: 'q1', time: daysAgo(env, 3), description: 'АТБ', mcc: 5411, amount: -5000, balance: 3995000 }];
  c.connectMono();
  env.userProps.setProperties({ tg_token: '123456:' + 'Z'.repeat(35), tg_chat: '42' });
  env.opts.promptText = ['1']; c.monoPeriodNow();
  assert.strictEqual(env.userProps.getProperty('period_job'), null, '1 місяць: одне вікно, одразу готово');
  assert.strictEqual(JSON.parse(env.userProps.getProperty('period_last')).total, 1);
  assert.ok(/Готово/.test(env.calls.dialogs.pop().html));
  assert.strictEqual(env.sent.length, 0, 'одне вікно: без Telegram (видно у вікні)');
  // триває → меню показує стан; «Так»: зупинено
  env.opts.promptText = ['6']; c.monoPeriodNow();
  assert.strictEqual(JSON.parse(env.userProps.getProperty('period_job')).total, 6);
  env.opts.alertAnswer = 'NO'; c.monoPeriodNow();
  assert.ok(/Пройдено вікон 1 з 6/.test(env.calls.alerts.pop()[1])); assert.ok(env.userProps.getProperty('period_job'), '«Ні»: триває');
  env.opts.alertAnswer = 'YES'; c.monoPeriodNow();
  assert.strictEqual(env.userProps.getProperty('period_job'), null); assert.strictEqual(run(env, 'mbsPeriod'), 0);
  assert.strictEqual(JSON.parse(env.userProps.getProperty('period_last')).err, 'зупинено вручну');
  // Mono 400 (наприклад, так далеко не віддає): зупинка одразу, з текстом Mono
  env.opts.promptText = ['4']; c.monoPeriodNow();
  env.fail = { code: 400, body: { errorDescription: 'Statement period is too old' } };
  c.mbsPeriod();
  let last = JSON.parse(env.userProps.getProperty('period_last'));
  assert.ok(!last.ok && /кодом 400: Statement period is too old/.test(last.err), last.err); assert.strictEqual(run(env, 'mbsPeriod'), 0);
  assert.ok(c.bankStatusLines_(env.ss).some(l => /Імпорт за період .*⚠️/.test(l)));
  // 500: повтор через хвилину; третя невдача поспіль: зупинка
  env.fail = null; env.opts.promptText = ['4']; c.monoPeriodNow();
  env.fail = { code: 500, body: 'oops' };
  c.mbsPeriod();
  assert.ok(env.userProps.getProperty('period_job') && run(env, 'mbsPeriod') === 1, 'збій: ще раз за хвилину');
  assert.ok(c.bankStatusLines_(env.ss).some(l => /триває.*остання спроба: Mono відповів кодом 500/.test(l)));
  c.mbsPeriod(); c.mbsPeriod();
  last = JSON.parse(env.userProps.getProperty('period_last'));
  assert.ok(!last.ok && /кодом 500/.test(last.err)); assert.strictEqual(run(env, 'mbsPeriod'), 0);
  // інша дія скрипта триває: нічого не робить, пробує за хвилину
  env.fail = null; env.opts.promptText = ['4']; c.monoPeriodNow();
  const done0 = JSON.parse(env.userProps.getProperty('period_job')).done;
  env.opts.lockBusy = true; c.mbsPeriod(); env.opts.lockBusy = false;
  assert.strictEqual(JSON.parse(env.userProps.getProperty('period_job')).done, done0); assert.strictEqual(run(env, 'mbsPeriod'), 1);
  // Google обірвав виконання тричі поспіль (спроба рахується ще до запиту): зупинка з поясненням, без зависання
  const jb = JSON.parse(env.userProps.getProperty('period_job')); jb.fails = 3; env.userProps.setProperty('period_job', JSON.stringify(jb));
  c.mbsPeriod();
  assert.ok(/три спроби поспіль обірвались/.test(JSON.parse(env.userProps.getProperty('period_last')).err)); assert.strictEqual(run(env, 'mbsPeriod'), 0);
  // зупинили з меню, поки йшло вікно: робота не воскресає
  env.opts.promptText = ['4']; c.monoPeriodNow();
  const fetch0 = env.ctx.UrlFetchApp.fetch;
  env.ctx.UrlFetchApp.fetch = (url, o) => { if (/statement/.test(url)) env.userProps.deleteProperty('period_job'); return fetch0(url, o); };
  c.mbsPeriod();
  env.ctx.UrlFetchApp.fetch = fetch0;
  assert.strictEqual(env.userProps.getProperty('period_job'), null); assert.strictEqual(run(env, 'mbsPeriod'), 0);
  // відключили Mono: роботу й розклад прибрано
  env.opts.promptText = ['4']; c.monoPeriodNow();
  assert.strictEqual(run(env, 'mbsPeriod'), 1);
  c.disconnectMono();
  assert.strictEqual(env.userProps.getProperty('period_job'), null); assert.strictEqual(run(env, 'mbsPeriod'), 0);
  c.monoPeriodNow();
  assert.ok(/Mono ще не підключено/.test(env.calls.dialogs.pop().html));
});

t('🗓 period: the journal runs out of rows → stops and says so; what fitted stays; pure monthsBack_', () => {
  const env = periodSetup([]), c = env.ctx;
  env.state.items = [0, 1, 2, 3].map(i => ({ id: 'f' + i, time: daysAgo(env, 3 + i), description: 'АТБ', mcc: 5411, amount: -1000 * (i + 1), balance: 3990000 }));
  c.connectMono();
  const j = env.ss.getSheetByName('Журнал'), J_LAST = K(c, 'J_LAST'), d = env.Utilities.parseDate('2026-01-05', 'Europe/Kiev', 'yyyy-MM-dd');
  j.getRange(5, 1, J_LAST - 5 - 1, 5).setValues(Array.from({ length: J_LAST - 6 }, () => [d, 1, 'Продукти', 'заповнювач', '✅ факт']));   // вільно 2 рядки
  env.opts.promptText = ['2']; c.monoPeriodNow();
  const last = JSON.parse(env.userProps.getProperty('period_last'));
  assert.ok(!last.ok && /закінчились вільні рядки/.test(last.err), JSON.stringify(last)); assert.strictEqual(run(env, 'mbsPeriod'), 0);
  assert.ok(/не вмістилось 2/.test(env.calls.dialogs.pop().html));
  assert.strictEqual(c.readJournal_(env.ss).filter(o => /^f\d$/.test(o.bank[0])).length, 2);
  assert.strictEqual(c.monthsBack_('2026-11-28', 3), '2026-08-28'); assert.strictEqual(c.monthsBack_('2026-03-31', 1), '2026-02-28');
  assert.strictEqual(c.monthsBack_('2026-01-15', 12), '2025-01-15'); assert.strictEqual(c.monthsBack_('2024-03-30', 1), '2024-02-29');
});

console.log(n + ' tests passed');
