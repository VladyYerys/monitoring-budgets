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
  assert.ok(/MATCH\("Наступний місяць: вільно"/.test(F(env, 'H1')) && /Наступний місяць: вільно на день/.test(F(env, 'G1')), 'вільно: з «моїх» наступного місяця');
  assert.ok(/'Дні'!\$M\$6/.test(F(env, 'H2')) && /на картці/.test(F(env, 'G2')) && /Найнижча точка за планом/.test(F(env, 'H3')) && /Мін\. на картці/.test(F(env, 'G3')));
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
  // через перемикач (блок «🧩 Модулі» на «Налаштуваннях»)
  const st = env.ss.getSheetByName('Налаштування'), r = 44 + K(c, 'MODS').findIndex(m => m.id === 'credit');
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

t('next month «вільно» from «мої» (version 19): income into the budget month (facts after pay day and plans) minus set-aside, minus spending plans, minus budgets above what is planned; per day; «Мін.» only from today', () => {
  const env = setup(), c = env.ctx, m = env.ss.getSheetByName('Головне');
  const rows = m.getRange('A1:B120').getValues(), fx = m.getRange('B1:B120').getFormulas();
  const at = l => { const i = rows.findIndex(r => r[0] === l); assert.ok(i >= 0, 'немає «' + l + '»'); return fx[i][0]; };
  const nm = '(EOMONTH(TODAY(),0)+1)';
  assert.ok(at('Наступний місяць: «мої» вже є').indexOf("SUMIFS('Журнал'!$S$5:$S$3004,'Журнал'!$Q$5:$Q$3004," + nm + ')') > 0);
  assert.ok(at('Наступний місяць: дохід за планом').indexOf("'Журнал'!$R$5:$R$3004,\"надходження\"") > 0, at('Наступний місяць: дохід за планом'));
  assert.ok(at('Наступний місяць: відкласти на погашення').indexOf("'Журнал'!$R$5:$R$3004,\"відкладення\"") > 0);
  ['Витрати', 'Заощадження', 'Податки'].forEach(t => assert.ok(at('Наступний місяць: плани витрат').indexOf('"' + t + '"') > 0, t));
  assert.ok(!/"Надходження"|"Резерв"/.test(at('Наступний місяць: плани витрат')), 'дохід і резерв не віднімаються');
  assert.strictEqual(at('Наступний місяць: бюджети понад плани'), "=SUM('Довідник'!$V$5:$V$84)");
  assert.ok(/^=B\d+\+B\d+-B\d+-B\d+-B\d+$/.test(at('Наступний місяць: вільно')), at('Наступний місяць: вільно'));
  assert.ok(/\/DAY\(EOMONTH\(TODAY\(\),1\)\)$/.test(at('Наступний місяць: вільно на день')));
  const d = env.ss.getSheetByName('Довідник'), v5 = d.getRange('V5').getFormula();
  assert.ok(/\$B5="Витрати",N\(\$F5\)>0\),MAX\(0,N\(\$F5\)-SUMIFS\('Журнал'!\$N\$5:\$N\$3004,'Журнал'!\$C\$5:\$C\$3004,\$A5,'Журнал'!\$Q\$5/.test(v5), v5);
  assert.ok(d.isColumnHiddenByUser(22));
  const days = env.ss.getSheetByName('Дні');
  assert.ok(/^=MINIFS\(\$L\$6:\$L\$98,\$A\$6:\$A\$98,">="&TODAY\(\)\)$/.test(days.getRange('B2').getFormula()), days.getRange('B2').getFormula());
  assert.ok(/FILTER\(\$A\$6:\$A\$98,\$L\$6:\$L\$98=B2,\$A\$6:\$A\$98>=TODAY\(\)\)/.test(days.getRange('C2').getFormula()));
  assert.ok(/^=MINIFS\(\$M\$6/.test(days.getRange('B4').getFormula()));
  // крок 19: стара таблиця (без рядків, без V, мінімум з усього вікна) → нова
  m.getRange('A1:D120').getValues().forEach((r, i) => { if (/^Наступний місяць/.test(r[0])) m.getRange(i + 1, 1, 1, 4).setValues([['', '', '', '']]); });
  days.getRange('B2').setFormula('=MIN(L6:L98)'); d.getRange('V5').setFormula('');
  env.docProps.setProperty('schema_v', '18');
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 19).done(env.ss), false);
  c.schemaTick_(env.ss);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.ok(/MINIFS/.test(days.getRange('B2').getFormula()) && d.getRange('V5').getFormula() !== '' && c.mainHasLabel_(env.ss, 'Наступний місяць: вільно'));
  assert.ok(/Наступний місяць: вільно"/.test(env.ss.getSheetByName('Журнал').getRange('H1').getFormula()));
  assert.strictEqual(K(c, 'STEPS').find(x => x.v === 19).done(env.ss), true);
});

console.log(n + ' tests passed');
