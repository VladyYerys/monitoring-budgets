// «Історія місяців»: значення з «Місяців» за потрібний місяць; рік на «Місяцях» повертається; перезапис без дублів; сортування.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);

function setup() {
  const env = makeEnv({ sheetsLike: true });
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  const c = env.ctx, m = env.ss.getSheetByName('Місяці');
  // «пораховані» значення (мок формул не рахує): жовтень 2026, факт
  const put = (label, v, k) => { const r = c.monthsSummaryRow_(label); m.put(r, c.mCols_(k || 10).fact, v, m.cell(r, c.mCols_(k || 10).fact).f); };
  const putCat = (i, name, type, plan, fact, k) => {
    const r = K(c, 'M_FIRST') + i, col = c.mCols_(k || 10);
    m.put(r, 1, name, m.cell(r, 1).f); m.put(r, 2, type, m.cell(r, 2).f); m.put(r, 3, 'Група', m.cell(r, 3).f);
    m.put(r, col.plan, plan, m.cell(r, col.plan).f); m.put(r, col.fact, fact, m.cell(r, col.fact).f);
  };
  return { env, c, m, put, putCat };
}

t('October: summary row + categories, values (not formulas), top-5, reasons and holes as text; month key stays text', () => {
  const x = setup();
  x.m.getRange('B2').setValue(2026);
  x.put('Витрати', 30000); x.put('Усі списання', 41000); x.put('Всього надходжень', 45000); x.put('Мої на кінець місяця', 1200);
  x.put('На картці на кінець місяця', -50000); x.put('Борг за лімітом на кінець місяця', 50000); x.put('Непомітні разом', 1500);
  x.putCat(0, 'Продукти', 'Витрати', 12000, 14000); x.putCat(1, 'Кафе й ресторани', 'Витрати', 3000, 7000); x.putCat(2, 'Зарплата', 'Надходження', 44000, 45000);
  const hr = x.c.monthsSummaryRow_('НЕПОМІТНІ ВИТРАТИ') + 1;
  x.m.put(hr, 1, '☕ Кава та перекуси', x.m.cell(hr, 1).f); x.m.put(hr, x.c.mCols_(10).fact, 900, x.m.cell(hr, x.c.mCols_(10).fact).f);
  const rr = x.c.monthsSummaryRow_('З ліміту: без причини');
  x.m.put(rr, x.c.mCols_(10).fact, 1200, x.m.cell(rr, x.c.mCols_(10).fact).f);
  const r = x.c.historyWrite_(x.env.ss, '2026-10');
  assert.strictEqual(r.cats, 3);
  const h = x.env.ss.getSheetByName('Історія місяців'), cols = K(x.c, 'H_COLS');
  const row = h.getRange(4, 1, 1, cols.length).getValues()[0], at = l => row[cols.indexOf(l)];
  assert.strictEqual(at('Місяць'), '2026-10');
  assert.strictEqual(JSON.stringify([at('Витрати'), at('Усі списання'), at('Всього надходжень'), at('Мої на кінець місяця'), at('На картці на кінець місяця')]),
    JSON.stringify([30000, 41000, 45000, 1200, -50000]));
  assert.strictEqual(at('Топ-5 категорій'), 'Продукти 14 000 ₴; Кафе й ресторани 7 000 ₴');
  assert.strictEqual(at('Непомітні: з чого'), '☕ Кава та перекуси 900 ₴');
  assert.strictEqual(at('З ліміту: причини'), 'без причини 1 200 ₴');
  assert.ok(h.getRange(4, 1, 1, cols.length).getFormulas()[0].every(f => !f));
  const cs = x.env.ss.getSheetByName('Історія категорій').getRange(4, 1, 3, 6).getValues();
  assert.strictEqual(JSON.stringify(cs[0]), JSON.stringify(['2026-10', 'Продукти', 'Витрати', 'Група', 12000, 14000]));
});

t('rewriting the same month: no duplicates; another year: «Місяці» switch for the read and switch back', () => {
  const x = setup();
  x.m.getRange('B2').setValue(2026);
  x.c.historyWrite_(x.env.ss, '2026-10'); x.c.historyWrite_(x.env.ss, '2026-10'); x.c.historyWrite_(x.env.ss, '2025-12');
  const h = x.env.ss.getSheetByName('Історія місяців');
  assert.strictEqual(JSON.stringify(h.getRange(4, 1, 3, 1).getValues().map(r => String(r[0]))), JSON.stringify(['2025-12', '2026-10', '']));
  assert.strictEqual(Number(x.m.getRange('B2').getValue()), 2026);
});

console.log(n + ' tests passed');
