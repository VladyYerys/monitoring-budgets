// Спільний вхід банків: «Рахунки», ПриватБанк «Автоклієнт», виписки CSV / Google Таблиці з папки, без дублів між джерелами.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const PB_TOKEN = 'pbTestToken' + 'q'.repeat(30), IBAN = 'UA' + '1'.repeat(27);

function setup(opts) {
  const env = makeEnv(opts || {});
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  return env;
}
const today = env => env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
const ddmm = k => k.slice(8, 10) + '.' + k.slice(5, 7) + '.' + k.slice(0, 4);

t('parseStatement_: header not on the first row, signed amounts with spaces and commas, debit/credit columns, Windows-1251 CSV, stable ids', () => {
  const env = makeEnv(), c = env.ctx, digest = c.md5Short_;
  const rows = [['Виписка по картці', ''], ['', ''], ['Дата', 'Опис операції', 'MCC', 'Сума в валюті картки'],
    ['01.11.2026 10:15', 'АТБ', '5411', '−1 234,50'], ['02.11.2026', 'Зарплата', '', '42 000,00'], ['', 'підсумок', '', '40765,5'], ['01.11.2026 10:15', 'АТБ', '5411', '−1 234,50']];
  const p = c.parseStatement_(rows, 'Europe/Kiev', 'Ощад', digest);
  assert.strictEqual(p.items.length, 3); assert.strictEqual(p.skipped, 1);
  assert.strictEqual(JSON.stringify(p.items.map(i => i.amount)), JSON.stringify([-123450, 4200000, -123450]));
  assert.strictEqual(p.items[0].mcc, 5411);
  assert.notStrictEqual(p.items[0].id, p.items[2].id, 'дві однакові операції за день: різні номери');
  assert.strictEqual(c.parseStatement_(rows, 'Europe/Kiev', 'Ощад', digest).items[2].id, p.items[2].id, 'вдруге ті самі номери');
  const dc = c.parseStatement_([['Дата операції', 'Призначення платежу', 'Дебет', 'Кредит'], ['2026-11-03', 'Оренда', '15000.00', ''], ['2026-11-04', 'Оплата', '', '3,000.50']],
    'Europe/Kiev', 'x', digest);
  assert.strictEqual(JSON.stringify(dc.items.map(i => i.amount)), JSON.stringify([-1500000, 300050]));
  const text = 'Дата;Деталі операції;Сума\n05.11.2026;Кав\'ярня;-85,00\n';
  assert.strictEqual(c.csvRows_(text)[1][2], '-85,00');
  const f = env.addFile(env.root, 'oschad.csv', 'text/csv', Buffer.from(new TextEncoder().encode('')));
  f.content = Buffer.from(require('child_process').execSync('printf "%s" "$T" | iconv -f utf-8 -t cp1251', { env: { T: text } }));
  const got = c.stmtFileRows_(f);
  assert.strictEqual(got.rows[1][1], 'Кав\'ярня', 'Windows-1251 розпізнано');
});

t('statements from the folder: Google Sheet read-only, personal-looking table is skipped, drafts on the chosen account, second run adds nothing', () => {
  const k = '2026-11-05';
  const env = setup({ promptText: ['Ощад картка'] }), c = env.ctx;
  const folder = c.stmtFolder_(env.ss);
  env.addSheetFile(folder, 'Виписка листопад', [['Дата', 'Опис', 'Сума'], [ddmm(k), 'Сільпо', '-512,40'], [ddmm(k), 'Переказ на картку', '-1000']]);
  env.addSheetFile(folder, 'Monitoring Budget (особиста)', [['Дата', 'Сума'], ['01.11.2026', '1']], ['m', 'y', 'Журнал']);
  c.importStatements();
  const html = env.calls.dialogs.pop().html;
  assert.ok(/нових чернеток 2/.test(html) && /особисту таблицю/.test(html), html);
  const rows = c.readJournal_(env.ss).filter(o => !o.blank);
  assert.strictEqual(rows.length, 2);
  assert.ok(rows.every(o => o.acct === 'Ощад картка' && o.e === '🏦 чернетка' && /^csv:/.test(o.bank[0])));
  assert.strictEqual(rows.find(o => o.b === 512.4).c, 'Продукти');
  assert.strictEqual(rows.find(o => o.b === 1000).c, '↔ Переказ (уточніть)');
  assert.strictEqual(c.accList_(env.ss)[0].kind, 'основна картка', 'перший рахунок: основна картка');
  c.importStatements();
  assert.ok(/Нових файлів немає/.test(env.calls.dialogs.pop().html));
});

t('ПриватБанк: token and IBAN only in user properties; only processed UAH operations; FOP account defaults; pages; no duplicate with a statement of the same account', () => {
  const k = '2026-11-20';
  const pages = { '': { status: 'SUCCESS', exist_next_page: true, next_page_id: 'p2', transactions: [
      { ID: '901', TRANTYPE: 'C', SUM: '15000.00', CCY: 'UAH', DAT_OD: ddmm(k), TIM_P: '10:00', OSND: 'Оплата за послуги згідно рахунку', PR_PR: 'r' },
      { ID: '902', TRANTYPE: 'D', SUM: '450.00', CCY: 'UAH', DAT_OD: ddmm(k), TIM_P: '11:00', OSND: 'Єдиний податок за листопад', PR_PR: 'r' },
      { ID: '903', TRANTYPE: 'D', SUM: '100.00', CCY: 'UAH', DAT_OD: ddmm(k), TIM_P: '12:00', OSND: 'в обробці', PR_PR: 'p' }] },
    'p2': { status: 'SUCCESS', exist_next_page: false, transactions: [
      { ID: '904', TRANTYPE: 'D', SUM: '5000.00', CCY: 'UAH', DAT_OD: ddmm(k), TIM_P: '15:00', OSND: 'Переказ власних коштів на картку', PR_PR: 'r' },
      { ID: '905', TRANTYPE: 'D', SUM: '10.00', CCY: 'USD', DAT_OD: ddmm(k), TIM_P: '15:00', OSND: 'usd', PR_PR: 'r' }] } };
  const heads = [];
  const env = setup({ promptText: [PB_TOKEN, IBAN.toLowerCase()], mono: (url, o) => {
    heads.push(o.headers);
    if (/statements\/transactions/.test(url)) { const f = (url.match(/followId=([^&]+)/) || [])[1] || ''; return { code: 200, body: /limit=20$/.test(url) ? pages['p2'] : pages[f] }; }
    return { code: 404, body: {} };
  } }), c = env.ctx;
  c.connectPrivat();
  assert.strictEqual(env.userProps.getProperty('pb_token'), PB_TOKEN);
  assert.ok(heads.every(h => h.token === PB_TOKEN && /charset=utf8/.test(h['Content-Type']) && h['User-Agent']));
  const acc = c.accList_(env.ss)[0];
  assert.strictEqual(JSON.stringify([acc.name, acc.kind]), JSON.stringify(['ПриватБанк ФОП *1111', 'рахунок ФОП']));
  const r = c.importBank_(env.ss);
  assert.strictEqual(r.drafts, 3, 'проведені гривневі: 901, 902, 904');
  const rows = c.readJournal_(env.ss).filter(o => !o.blank), by = id => rows.find(o => o.bank[0] === 'pb:' + id);
  assert.strictEqual(by('901').c, 'Дохід ФОП');
  assert.strictEqual(by('902').c, 'Єдиний податок', 'правило податків');
  assert.strictEqual(by('904').c, 'Переказ між своїми рахунками');
  assert.ok(rows.every(o => o.kind === 'рахунок ФОП'));
  assert.ok(env.ss.getSheets().every(s => Array.from(s.cells.values()).every(x => String(x.v).indexOf(PB_TOKEN) < 0 && String(x.v).indexOf(IBAN) < 0)));
  // виписка того ж рахунку з тією ж операцією: без дубля
  const res = c.bankIngest_(env.ss, { acc: acc, items: [{ id: 'csv:abc', time: rows[0].a.getTime() / 1000, amount: 1500000, description: 'дохід' }],
    dateOf: () => k, timeOf: () => ddmm(k), src: 'csv' });
  assert.strictEqual(res.written, 0); assert.strictEqual(res.plan.stats.other, 1);
});

t('privatItem_ sign and id; FOP account rows do not touch «мої» or the card balance (formulas), other cards count in «мої» only', () => {
  const env = setup(), c = env.ctx;
  const it = c.privatItem_({ ID: '', REF: 'R1', REFN: '2', TRANTYPE: 'D', SUM: '12,50', DAT_OD: '03.11.2026', TIM_P: '09:30', PR_PR: 'r' }, 'Europe/Kiev');
  assert.strictEqual(it.id, 'pb:R1.2'); assert.strictEqual(it.amount, -1250);
  const j = env.ss.getSheetByName('Журнал');
  const L = j.getRange('L5').getFormula(), T = j.getRange('T5').getFormula(), S = j.getRange('S5').getFormula(), AG = j.getRange('AG5').getFormula();
  assert.ok(/AG5="основна картка"/.test(L) && /AH5/.test(L), L);
  assert.ok(/AG5="рахунок ФОП"/.test(T) && /N\(AH5\)<0/.test(T), T);
  assert.ok(/AG5="рахунок ФОП"/.test(S), S);
  assert.ok(/VLOOKUP\(AF5,'Рахунки'!\$A\$5:\$D\$34,4,FALSE\)/.test(AG), AG);
  const m = env.ss.getSheetByName('Місяці').getRange('E7').getFormula();
  assert.ok(/'Журнал'!\$AG\$5:\$AG\$3004,"<>рахунок ФОП"/.test(m) && /'Журнал'!\$G\$5:\$G\$3004,\$A\$3/.test(m), m);
});

console.log(n + ' tests passed');
