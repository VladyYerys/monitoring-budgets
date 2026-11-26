// Шапка журналу (рядки 1…3, A…J): мої й картка, три найгірші бюджети зі смужкою, наступний місяць, грейс; модулі; крок версії 7.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);

function setup() {
  const env = makeEnv();
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  return env;
}
const J = env => env.ss.getSheetByName('Журнал');
const F = (env, a) => J(env).getRange(a).getFormula();
const V = (env, a) => J(env).getRange(a).getValue();

t('new table: three header rows A…J with labels and formulas from «Головне», link to the input row, merged bar cells, coloured header row', () => {
  const env = setup(), j = J(env);
  assert.strictEqual(V(env, 'A1'), '💚 Мої зараз'); assert.strictEqual(V(env, 'A2'), '💳 На картці'); assert.strictEqual(V(env, 'A3'), '✍️ Писати у рядок');
  assert.ok(/MATCH\("Доступно моїх сьогодні"/.test(F(env, 'B1')) && /MATCH\("На картці зараз"/.test(F(env, 'B2')));
  assert.ok(F(env, 'B3').indexOf('HYPERLINK("#gid=' + j.getSheetId() + '&range=A"&') > 0, F(env, 'B3'));
  assert.ok(/MATCH\(TRUE,INDEX\(\$A\$5:\$A\$3004="",0\),0\)\+4/.test(F(env, 'B3')) && /N\(INDEX\(\$A:\$A,/.test(F(env, 'B3')), 'перший порожній рядок, після жовтої лінії');
  [1, 2, 3].forEach(k => {
    assert.ok(F(env, 'C' + k).indexOf('LARGE(\'Довідник\'!$P$5:$P$84,' + k + ')') > 0, 'бюджет №' + k);
    assert.ok(/SPARKLINE\(INDEX\('Довідник'!\$Q\$5:\$R\$84,.*'Довідник'!\$T\$5:\$U\$8\)/.test(F(env, 'D' + k)), F(env, 'D' + k));
    assert.ok(j.merges.indexOf('D' + k + ':E' + k) >= 0);
  });
  assert.ok(/" · \+"&TEXT/.test(F(env, 'F1')) && /швидше норми/.test(F(env, 'F1')));
  assert.ok(/CHOOSE\(MONTH\(EOMONTH\(TODAY\(\),1\)\),"Січень",.*"Грудень"\)&": вільно ≈ "/.test(F(env, 'G1')));
  assert.ok(/Прогноз на кінець наступного місяця/.test(F(env, 'H1')) && /'Дні'!\$M\$6/.test(F(env, 'H2')) && /Найнижча точка за планом/.test(F(env, 'H3')));
  assert.ok(/Грейс: строк/.test(F(env, 'I1')) && /Грейс: внести до строку/.test(F(env, 'J1')) && /Грейс: комісія, якщо все переказами/.test(F(env, 'J2')));
  assert.ok(/Чернеток чекають підтвердження/.test(F(env, 'J3')));
  for (const a of ['A1', 'B1', 'C1', 'D1', 'F1', 'G1', 'H1', 'I1', 'J1', 'B3']) assert.ok((F(env, a) || '').indexOf('{') < 0, 'без масивів у формулах: ' + a);
  assert.ok(j.notes && /жовтої лінії/.test(j.notes.A3), 'підказка в примітці');
  assert.strictEqual(V(env, 'A4'), 'Дата', 'рядок заголовків на місці');
  const cf = j.cf.filter(r => r.spec.ranges.every(g => g.getLastRow() <= 3));
  assert.ok(cf.some(r => r.spec.startsWith === '⛔') && cf.some(r => r.spec.startsWith === '✓') && cf.some(r => r.spec.contains === ' · +'));
  assert.ok(j.cf.some(r => r.spec.formula && /\$A5=""/.test(r.spec.formula)), 'кольори журналу (жовта лінія) лишились');
});

t('«Довідник» helper columns M…R: spent this month (facts + drafts, without FOP account, by owner), norm, state, order, bar parts; bar settings T5:U8; hidden', () => {
  const env = setup(), d = env.ss.getSheetByName('Довідник'), f = a => d.getRange(a).getFormula();
  const m = f('M5');
  assert.ok(/^=IF\(AND\(\$A5<>"",\$B5="Витрати",N\(\$F5\)>0\),IF\(\$G5="",SUMIFS\('Журнал'!\$O\$5:\$O\$3004,/.test(m), m);
  assert.ok(/'Журнал'!\$AG\$5:\$AG\$3004,"<>рахунок ФОП"/.test(m) && /'Журнал'!\$E\$5:\$E\$3004,"🏦 чернетка"/.test(m) && /'Журнал'!\$G\$5:\$G\$3004,\$G5/.test(m));
  assert.ok(/DAY\(TODAY\(\)\)\/DAY\(EOMONTH\(TODAY\(\),0\)\)/.test(f('N5')));
  assert.strictEqual(f('O5'), '=IF($M5="","",IF($M5>N($F5)+0.005,"⛔",IF($M5>$N5*1.1,"▲","✓")))');
  assert.ok(/ROW\(\)\/1E6/.test(f('P5')) && /MIN\(\$M5,N\(\$F5\)\)\/MAX/.test(f('Q5')) && /MAX\(0,\$M5-N\(\$F5\)\)\/MAX/.test(f('R5')));
  assert.ok(f('M84') && !f('M85'), 'усі 80 категорій');
  assert.strictEqual(JSON.stringify(d.getRange('T5:U8').getValues()), JSON.stringify([['charttype', 'bar'], ['max', 1], ['color1', '#4a86e8'], ['color2', '#e06666']]));
  assert.ok(d.hidden.some(h => h[0] === 13 && h[1] === 9), 'M…U сховано');
});

t('modules: credit off → «Можна на день» instead of grace; budgets off → month spending instead of budgets; on again: back; journal rows untouched', () => {
  const env = setup(), c = env.ctx, j = J(env);
  j.getRange('A5:E6').setValues([[new Date(), 120, 'Продукти', 'мій рядок', '✅ факт'], [new Date(), 50, 'Кава та перекуси', 'ще', '✅ факт']]);
  const before = JSON.stringify(j.getRange('A5:H6').getValues());
  c.modSet_(env.ss, 'credit', false); c.journalHeader_(env.ss);
  assert.strictEqual(V(env, 'I1'), 'Можна на день'); assert.ok(/Можна на день/.test(F(env, 'J1')) && /Днів до кінця місяця/.test(F(env, 'J2')));
  c.modSet_(env.ss, 'budgets', false); c.journalHeader_(env.ss);
  assert.strictEqual(V(env, 'C1'), 'Витрати місяця'); assert.ok(/Витрати цього місяця: факт/.test(F(env, 'F1')) && !F(env, 'D1'));
  c.modSet_(env.ss, 'credit', true); c.modSet_(env.ss, 'budgets', true); c.journalHeader_(env.ss);
  assert.ok(/Грейс: строк/.test(F(env, 'I1')) && /LARGE/.test(F(env, 'C1')));
  assert.strictEqual(JSON.stringify(j.getRange('A5:H6').getValues()), before, 'дані журналу не змінились');
  assert.strictEqual(j.cf.filter(r => r.spec.ranges.every(g => g.getLastRow() <= 3)).length, 7, 'правила шапки не дублюються');
  // через перемикач на «Старті»
  const st = env.ss.getSheetByName('Старт'), r = 26 + K(c, 'MODS').findIndex(m => m.id === 'credit');
  st.getRange(r, 2).setValue('ні'); c.mbsEdit({ range: st.getRange(r, 2) });
  assert.strictEqual(V(env, 'I1'), 'Можна на день');
});

t('upgrade of a version-6 table to 7: old one-line header replaced, backup first; the link follows the sheet id after a copy (once a day)', () => {
  const env = setup(), c = env.ctx, j = J(env), d = env.ss.getSheetByName('Довідник');
  j.getRange('A1:J3').clearContent();
  j.getRange('A1').setValue('📒 Журнал'); j.getRange('A3').setFormula('="💚 Мої: "&0');
  d.getRange('T5:U8').clearContent();
  env.docProps.setProperty('schema_v', '6');
  const f0 = env.files.length;
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.strictEqual(env.files.length, f0 + 1, 'копія перед змінами');
  assert.strictEqual(c.headerIsNew_(env.ss), true);
  j.gid = 4242;                                                          // копія таблиці: інший id аркуша
  env.docProps.deleteProperty('hdr_day');
  assert.strictEqual(c.headerIfDue_(env.ss), true);
  assert.ok(F(env, 'B3').indexOf('#gid=4242&') > 0);
  assert.strictEqual(c.headerIfDue_(env.ss), null, 'раз на день');
});

console.log(n + ' tests passed');
