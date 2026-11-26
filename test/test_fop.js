// «ФОП»: ставки в таблиці, податки кожного ФОП окремо (групи 1, 2, 3), плани податків у журналі, ліміт доходу в Telegram.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

function setup(opts) {
  const env = makeEnv(opts || {});
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  env.ctx.modSet_(env.ss, 'fop', true);                              // модуль «🧾 ФОП» (у новій таблиці типово вимкнено)
  return env;
}
const rates = { ep1: 332.8, ep2: 1729.4, ep3: 5, esv: 1902.34, vz12: 864.7, vz3: 1, limit: [0, 1444049, 7211598, 10091049] };
const inc = (key, b, who) => ({ key: key, c: 'Дохід ФОП', e: '✅ факт', m: 0, o: b, b: b, g: who || '' });

t('sheet: 2026 rates row, formulas for income / limit / quarters; one FOP needs no «Хто», several do', () => {
  const env = setup(), f = env.ss.getSheetByName('ФОП');
  assert.strictEqual(JSON.stringify(f.getRange('A6:L6').getValues()[0]), JSON.stringify([2026, 8647, 3328, 332.8, 1729.4, 5, 1902.34, 864.7, 1, 1444049, 7211598, 10091049]));
  assert.ok(/SUMIFS\('Журнал'!\$O\$5:\$O\$3004,'Журнал'!\$C\$5:\$C\$3004,"Дохід ФОП",'Журнал'!\$G\$5:\$G\$3004,\$K14/.test(f.getRange('F14').getFormula()));
  assert.ok(/COUNTA\(\$A\$14:\$A\$18\)<=1,"<>✱"/.test(f.getRange('K14').getFormula()));
  assert.ok(/IF\(IFERROR\(VALUE\(\$B\$14&""\),0\)=3,\$C23\*/.test(f.getRange('D23').getFormula()), f.getRange('D23').getFormula());
  assert.strictEqual(f.getRange('B26').getFormula(), '="IV кв."');
  assert.ok(/DATE\(YEAR\(\$N23\),MONTH\(\$N23\),20\)/.test(f.getRange('I23').getFormula()));
});

t('fopWanted_: group 2 monthly ЄП + ВЗ by the 20th, ЄСВ after the quarter, group 3 from the quarter income 50 days later; paid ones skipped', () => {
  const c = makeEnv().ctx, today = '2026-11-25', horizon = '2026-12-31';
  const fops = [{ name: 'Олена', group: 3, own: null, vz: true, esv: true }, { name: 'Іван', group: 2, own: 1200, vz: true, esv: false }];
  const rows = [inc('2026-08-03', 100000, 'Олена'), inc('2026-09-15', 50000, 'Олена'), inc('2026-10-10', 70000, 'Олена'), inc('2026-09-20', 999, 'Іван')];
  const w = c.fopWanted_(fops, rates, rows, today, horizon, {});
  const by = k => w.find(x => x.key === k);
  assert.strictEqual(by('[ФОП] Іван: ЄП за грудень 2026').amount, 1200);
  assert.strictEqual(by('[ФОП] Іван: ВЗ за грудень 2026').date, '2026-12-20');
  assert.ok(!w.some(x => /Іван: ЄСВ/.test(x.key)), 'ЄСВ: ні');
  const ep = by('[ФОП] Олена: ЄП за III кв. 2026');
  assert.ok(!ep, 'строк III кв. (19.11) уже минув');
  assert.strictEqual(by('[ФОП] Олена: ЄСВ за IV кв. 2026'), undefined, 'ЄСВ за IV кв.: 20.01, за горизонтом');
  const w2 = c.fopWanted_(fops, rates, rows, '2026-10-15', '2026-11-30', {});
  const q3 = w2.find(x => x.key === '[ФОП] Олена: ЄП за III кв. 2026');
  assert.strictEqual(q3.date, '2026-11-19'); assert.strictEqual(q3.amount, 7500); assert.strictEqual(q3.who, 'Олена');
  assert.strictEqual(w2.find(x => x.key === '[ФОП] Олена: ВЗ за III кв. 2026').amount, 1500);
  assert.strictEqual(w2.find(x => x.key === '[ФОП] Олена: ЄСВ за III кв. 2026').amount, 5707.02);
  const paid = rows.concat([{ key: '2026-11-10', c: 'Єдиний податок', e: '✅ факт', m: 0, o: 7500, b: 7500, g: 'Олена' }]);
  assert.ok(!c.fopWanted_(fops, rates, paid, '2026-10-15', '2026-11-30', {}).some(x => x.key === '[ФОП] Олена: ЄП за III кв. 2026'), 'сплачено → без плану');
  assert.ok(!c.fopWanted_(fops, rates, rows, '2026-10-15', '2026-11-30', { 'Олена|III кв. 2026': true }).some(x => /Олена: (ЄП|ВЗ|ЄСВ) за III/.test(x.key)), '✅ → без плану');
});

t('plans sync: adds rows «🔁 план регулярний» with «[ФОП]», updates amounts in place, removes a future plan that is no longer needed, never touches your rows', () => {
  const env = setup(), c = env.ctx, f = env.ss.getSheetByName('ФОП'), j = env.ss.getSheetByName('Журнал');
  f.getRange('A14:E14').setValues([['Іван', '2', '', 'так', 'так']]);
  j.getRange('A5:H5').setValues([[new Date(), 777, 'Єдиний податок', 'мій ручний план', 'план', '', '', '']]);
  const st = c.fopPlansSync_(env.ss);
  assert.ok(st.added >= 2, JSON.stringify(st));
  const plans = () => c.readJournal_(env.ss).filter(o => o.d.indexOf('[ФОП]') === 0);
  assert.ok(plans().every(o => o.e === '🔁 план регулярний' && o.g === ''), 'один ФОП: «Хто» порожнє');
  const today = env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd'), nm = c.addDays_(c.lastDayOf_(today.slice(0, 7)), 1).slice(0, 7);
  const label = vm.runInContext('MONTHS_UA', c)[Number(nm.slice(5)) - 1] + ' ' + nm.slice(0, 4);
  assert.ok(plans().some(o => o.d === '[ФОП] Іван: ЄП за ' + label && o.b === 1729.4 && o.key === nm + '-20'), plans().map(o => o.d + ' ' + o.b + ' ' + o.key).join('; '));
  assert.ok(plans().some(o => o.d === '[ФОП] Іван: ВЗ за ' + label && o.b === 864.7));
  f.getRange('C14').setValue(1000);
  const st2 = c.fopPlansSync_(env.ss);
  assert.ok(st2.updated >= 1 && st2.added === 0, JSON.stringify(st2));
  assert.ok(plans().filter(o => /: ЄП за/.test(o.d)).every(o => o.b === 1000));
  f.getRange('D14').setValue('ні');
  const st3 = c.fopPlansSync_(env.ss);
  assert.ok(st3.removed >= 1 && !plans().some(o => /: ВЗ за/.test(o.d)), 'ВЗ вимкнули → майбутні плани ВЗ прибрано');
  const mine = c.readJournal_(env.ss).find(o => o.d === 'мій ручний план');
  assert.ok(mine && mine.b === 777);
});

t('Telegram: limit alert once per threshold (80 → 90 → 100 %), tax reminder 5 days before', () => {
  const env = setup(), c = env.ctx, f = env.ss.getSheetByName('ФОП'), j = env.ss.getSheetByName('Журнал');
  const today = env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd'), y = today.slice(0, 4);
  f.getRange('A14:E14').setValues([['Олена', '1', '', 'так', 'так']]);
  const put = (r, key, b, cat, note, status, svc) => { j.getRange(r, 1, 1, 8).setValues([[env.Utilities.parseDate(key, 'Europe/Kiev', 'yyyy-MM-dd'), b, cat, note, status, '', '', '']]);
    Object.keys(svc).forEach(col => j.put(r, env.colNum(col), svc[col], j.cell(r, env.colNum(col)).f)); };
  put(5, y + '-01-15', 1200000, 'Дохід ФОП', '', '✅ факт', { M: 0, O: 1200000 });
  put(6, c.addDays_(today, 5), 3099.84, 'ЄСВ', '[ФОП] Олена: ЄСВ за IV кв. ' + y, '🔁 план регулярний', { M: 1 });
  const m1 = c.buildMorning_(env.ss, {});
  assert.ok(m1.alerts.some(a => /ФОП Олена/.test(a) && /83 %/.test(a)), m1.alerts.join('\n'));
  assert.ok(m1.alerts.some(a => /Олена: ЄСВ за IV кв/.test(a) && /через 5 днів/.test(a)));
  assert.ok(!c.buildMorning_(env.ss, {}).alerts.some(a => /ФОП Олена/.test(a)), 'вдруге той самий поріг: без тривоги');
  j.getRange('B5').setValue(1450000); j.put(5, 15, 1450000, j.cell(5, 15).f);
  assert.ok(c.buildMorning_(env.ss, {}).alerts.some(a => /⛔/.test(a) && /100 %/.test(a)));
});

console.log(n + ' tests passed');
