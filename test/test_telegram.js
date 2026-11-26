// Telegram: підключення (токен лише в UserProperties), ранковий підсумок з «Головного», тривоги, самоперевірка; тиждень і місяць.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const BOT = '123456789:' + 'A'.repeat(35);

function setup(extra) {
  const sent = [];
  const env = makeEnv(Object.assign({ promptText: BOT, mono: (url, o) => {
    if (/getMe$/.test(url)) return { code: 200, body: { ok: true, result: { username: 'test_mb_bot' } } };
    if (/getUpdates$/.test(url)) return { code: 200, body: { ok: true, result: [{ message: { text: '/start', chat: { id: 77, type: 'private' } } }] } };
    if (/sendMessage$/.test(url)) { sent.push(JSON.parse(o.payload).text); return { code: 200, body: { ok: true, result: {} } }; }
    return { code: 404, body: { ok: false, description: 'nope for ' + url } };
  } }, extra || {}));
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  env.sent = sent;
  return env;
}

/** «Пораховані» значення «Головного» (мок формул не рахує): підпис → значення. */
function mainValues(env, vals) {
  const sh = env.ss.getSheetByName('Головне');
  for (let r = 4; r < 40; r++) {
    const l = sh.getRange(r, 1).getValue();
    if (l in vals) sh.put(r, 2, vals[l], sh.cell(r, 2).f);
  }
}
const D = k => new Date(Date.UTC(+k.slice(0, 4), +k.slice(5, 7) - 1, +k.slice(8, 10), 10));

t('connect: token and chat only in user properties; «підключено» sent; morning trigger at 05:00', () => {
  const env = setup();
  env.ctx.connectTelegram();
  assert.strictEqual(env.userProps.getProperty('tg_token'), BOT);
  assert.strictEqual(env.userProps.getProperty('tg_chat'), '77');
  assert.ok(/підключено/.test(env.sent[0]));
  assert.ok(env.triggers.some(t => t.getHandlerFunction() === 'mbsMorning' && t.spec.hour === 5));
  const cells = env.ss.getSheets().flatMap(s => Array.from(s.cells.values()).map(x => String(x.v)));
  assert.ok(cells.every(v => v.indexOf(BOT) < 0));
});

t('a Telegram error never shows the bot token', () => {
  const env = setup({ mono: () => ({ code: 401, body: { ok: false, description: 'Unauthorized ' + BOT } }) });
  assert.throws(() => env.ctx.tgCall_(BOT, 'getMe'), e => /Unauthorized/.test(e.message) && e.message.indexOf(BOT) < 0);
});

t('morning: numbers from «Головне», yesterday from the journal, alerts (mine below zero, red rows, plan due, limit without reason, forecast), self-check', () => {
  const env = setup(), c = env.ctx, today = env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd'), yday = c.addDays_(today, -1);
  env.userProps.setProperties({ tg_token: BOT, tg_chat: '77' });
  mainValues(env, { 'Сьогодні': D(today), 'Доступно моїх сьогодні': -1200, 'Можна на день': 0, 'Днів до кінця місяця': 6, 'Кредитний ліміт': 50000,
    'Борг за лімітом зараз': 20000, 'На картці зараз': -20000, 'Прогноз на кінець місяця (за планом)': -26000, 'Межа «близько до ліміту»': 5000,
    'Найнижча точка «як зараз»': -48000, 'Дата найнижчої «як зараз»': D(c.addDays_(today, 20)), 'Витрачаєте ≈ на день': 900, 'Як зараз: на картці на кінець місяця': -30000 });
  const j = env.ss.getSheetByName('Журнал');
  const put = (r, vals, svc) => { j.getRange(r, 1, 1, 8).setValues([vals]); Object.keys(svc || {}).forEach(col => j.put(r, env.colNum(col), svc[col], j.cell(r, env.colNum(col)).f)); };
  put(5, [D(yday), 300, 'Продукти', '', '✅ факт', '', '', ''], { I: 'Витрати', M: 0, O: 300, W: 300, X: 'без причини' });
  put(6, [D(yday), 50, '↔ Переказ (уточніть)', '', '✅ факт', '', '', ''], { M: 2 });
  put(7, [D(c.addDays_(today, -3)), 700, 'Комуналка', '', 'план', '', '', ''], { I: 'Витрати', M: 1 });
  const r = c.buildMorning_(env.ss, {});
  assert.ok(/Вчора витрачено: <b>300 ₴<\/b>, з ліміту 300 ₴/.test(r.text), r.text);
  assert.ok(/Доступно «моїх»: <b>−1 200 ₴/.test(r.text));
  assert.ok(r.alerts.some(a => /«Мої» закінчились/.test(a)));
  assert.ok(r.alerts.some(a => /З ліміту без причини: 1/.test(a)));
  assert.ok(r.alerts.some(a => /Червоні рядки/.test(a)));
  assert.ok(r.alerts.some(a => /План настав, підтвердіть: 1/.test(a)));
  assert.ok(r.alerts.some(a => /Прогноз/.test(a) && /2 000 ₴/.test(a)), r.alerts.join('\n'));
  assert.ok(/Самоперевірка/.test(r.text));
  env.ss.getSheetByName('Налаштування').getRange('B4').setValue(new Date(Date.UTC(2025, 0, 1, 10)));   // старт обліку раніше: є тиждень і місяць для звітів
  c.mbsMorning();
  assert.strictEqual(env.docProps.getProperty('tg_sent_day'), today);
  const kinds = env.sent.map(s => s.split('\n')[0]);
  assert.ok(/Підсумок/.test(kinds[0]) && kinds.some(k => /Тиждень/.test(k)) && kinds.some(k => /Підсумок: /.test(k)), kinds.join(' | '));
  const n0 = env.sent.length; c.mbsMorning();
  assert.strictEqual(env.sent.length, n0, 'вдруге того ж дня: нічого');
});

t('budgets: category plan from «Довідник», spent = facts + bank drafts of this month; over budget is an alert', () => {
  const env = setup(), c = env.ctx, today = env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  env.ss.getSheetByName('Довідник').getRange('F5').setValue(1000);                    // Продукти: 1 000 на місяць
  const rows = [{ key: today, c: 'Продукти', e: '✅ факт', m: 0, o: 800, b: 800, bank: ['', '', '', '', '', ''] },
    { key: today, c: 'Продукти', e: '🏦 чернетка', m: 2, o: 0, b: 300, bank: ['x', '', -300, '', '', ''] }];
  const b = c.budgetLines_(env.ss, rows, null);
  assert.strictEqual(b.length, 1);
  assert.ok(b[0].over && /перевищено на 100 ₴/.test(b[0].text) && /у чернетках 300 ₴/.test(b[0].text), b[0].text);
});

console.log(n + ' tests passed');
