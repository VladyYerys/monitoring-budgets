// Збирання шаблону в порожній таблиці: аркуші, заголовки, формули, списки; формули посилаються лише на наші аркуші.
const assert = require('assert'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);

function built() {
  const env = makeEnv();
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  return env;
}
const allFormulas = env => env.ss.getSheets().flatMap(sh => Array.from(sh.cells.entries()).filter(([, x]) => x.f).map(([k, x]) => [sh.getName(), k, x.f]));

t('all sheets, the empty starter sheet removed, marked as template, success dialog', () => {
  const env = built();
  assert.strictEqual(JSON.stringify(env.ss.getSheets().map(s => s.getName())), JSON.stringify(
    ['Старт', 'Журнал', 'Регулярні', 'Charts', 'charts_mob', 'Місяці', 'Дні', 'Довідник', 'Налаштування', 'ФОП', 'Борги', 'Головне', 'Правила', 'Рахунки', 'Залишки банку',
      'Банк', 'Історія місяців', 'Історія категорій', 'Перевірка', 'Дані графіків']));
  const hidden = env.ss.getSheets().filter(s => s.isSheetHidden()).map(s => s.getName());
  const shown = env.ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName());
  assert.strictEqual(JSON.stringify(shown), JSON.stringify(['Старт', 'Журнал', 'Регулярні', 'Charts', 'charts_mob']), 'на виду лише п\'ять вкладок: ' + shown);
  assert.strictEqual(hidden.length, 15);
  assert.strictEqual(env.docProps.getProperty('mbs'), 'monitoring-budgets-template');
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(env.ctx.schemaLast_()));
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('A40').getValue(), 'monitoring-budgets-template');
  const dlg = env.calls.dialogs.pop(); assert.ok(/Шаблон готовий/.test(dlg.t) && /версію <b>10<\/b>/.test(dlg.html), dlg.t + dlg.html);
  assert.strictEqual(env.files.length, 0, 'у порожній таблиці копія перед змінами не потрібна');
});

t('journal: headers, formulas in every row 5…3004, lists for status / category / hole / reason, service columns hidden', () => {
  const env = built(), j = env.ss.getSheetByName('Журнал');
  assert.strictEqual(JSON.stringify(j.getRange('A4:J4').getValues()[0].slice(0, 5)), JSON.stringify(['Дата', 'Сума, ₴', 'Категорія', 'Коментар', 'Статус']));
  ['I', 'J', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X'].forEach(c => {
    assert.ok(j.getRange(c + '5').getFormula().indexOf('5') > 0, c);
    assert.ok(j.getRange(c + '3004').getFormula().indexOf('3004') > 0, c + '3004');
    assert.strictEqual(j.cell(3005, env.colNum(c)).f, '', c + '3005');
  });
  assert.ok(/'Довідник'!\$A\$5:\$B\$84/.test(j.getRange('I5').getFormula()));
  assert.strictEqual(JSON.stringify(j.getRange('E5').getDataValidation().getCriteriaValues()[0]), JSON.stringify(['✅ факт', '🏦 чернетка', 'план', '🔁 план регулярний']));
  assert.ok(j.hidden.some(h => h[0] === 12 && h[1] === 23));                        // L…AH
  assert.strictEqual(j.getRange('Y4').getValue(), 'Банк: номер операції');
  assert.strictEqual(j.getRange('AE4').getValue(), 'Черговість запису (службовий)');
  assert.ok(j.cf.length >= 5 && j.cf.every(r => !r.spec.formula || !/,/.test(r.spec.formula.replace(/"[^"]*"/g, ''))), 'кольори без ком: працюють у будь-якому регіоні');
});

t('dictionary: neutral starter set, types and «мої гроші» flags, holes and reasons', () => {
  const env = built(), d = env.ss.getSheetByName('Довідник');
  const cats = d.getRange('A5:E84').getValues().filter(r => r[0]);
  assert.ok(cats.length >= 40 && cats.length <= 80);
  assert.ok(cats.every(r => r[1] === '' || K(env.ctx, 'TYPES').indexOf(r[1]) >= 0));
  assert.ok(cats.some(r => r[4] === 'надходження') && cats.some(r => r[4] === 'відкладення'));
  assert.ok(d.getRange('I5').getValue() && d.getRange('K5').getValue());
});

t('Місяці: year in B2, month starts in hidden row 5, category and summary formulas; Дні: 93 rows of formulas', () => {
  const env = built(), m = env.ss.getSheetByName('Місяці'), dd = env.ss.getSheetByName('Дні');
  assert.strictEqual(String(m.getRange('B2').getValue()), String(new Date().getFullYear()));
  assert.strictEqual(m.getRange('D5').getFormula(), '=DATE($B$2,1,1)');
  assert.strictEqual(m.getRange('AK5').getFormula(), '=DATE($B$2,12,1)');
  assert.ok(m.hiddenRows.some(h => h[0] === 5));
  assert.ok(/SUMIFS\('Журнал'!\$O\$5:\$O\$3004,'Журнал'!\$C\$5:\$C\$3004,\$A7/.test(m.getRange('E7').getFormula()));
  const row = l => env.ctx.monthsSummaryRow_(l);
  assert.ok(row('Усі списання') > 86 && /^=SUM\(E\d+:E\d+\)$/.test(m.getRange('E' + row('Усі списання')).getFormula()));
  assert.ok(/MAXIFS\('Залишки банку'!\$A\$5:\$A\$2004/.test(m.getRange('E' + row('На картці на кінець місяця')).getFormula()));
  assert.strictEqual(m.getRange('D' + row('На картці на кінець місяця')).getFormula(), '');      // «план» для залишків порожній
  assert.strictEqual(dd.getRange('A6').getFormula(), '=EOMONTH(TODAY(),-2)+1');
  assert.ok(dd.getRange('M98').getFormula() && !dd.getRange('A99').getFormula());
});

t('every formula: balanced brackets and quotes; refers only to sheets of the template', () => {
  const env = built(), names = new Set(env.ss.getSheets().map(s => s.getName()));
  const fs = allFormulas(env);
  assert.ok(fs.length > 40000, fs.length);
  assert.ok(!env.calls.valueFormulas, 'формули лише через setFormulas (з урахуванням регіону): ' + env.calls.valueFormulas);
  fs.forEach(([sh, k, f]) => {
    let depth = 0, q = false;
    for (const ch of f) { if (ch === '"') q = !q; else if (!q && ch === '(') depth++; else if (!q && ch === ')') depth--; assert.ok(depth >= 0, sh + k + f); }
    assert.ok(depth === 0 && !q, sh + ' ' + k + ' ' + f);
    (f.match(/'([^']+)'!/g) || []).forEach(x => assert.ok(names.has(x.slice(1, -2)), sh + ' ' + k + ' → ' + x));
  });
});

t('Місяці stays light: under 6 000 formulas per year (the personal model has tens of thousands per year)', () => {
  const env = built(), m = env.ss.getSheetByName('Місяці');
  const cnt = Array.from(m.cells.values()).filter(x => x.f).length;
  assert.ok(cnt < 6000, cnt);
});

t('Ukrainian-style table («;» between arguments): every formula is written with «;», decimals with «,»', () => {
  const env = makeEnv({ sep: ';' });
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  assert.strictEqual(env.docProps.getProperty('fsep'), ';');
  const fs = allFormulas(env);
  fs.forEach(([sh, k, f]) => assert.ok(!/,/.test(f.replace(/"[^"]*"/g, '').replace(/\d,\d/g, '')), sh + ' ' + k + ' ' + f));
  const chk = env.ss.getSheetByName('Перевірка').getRange('D5').getFormula();
  assert.ok(/<0,005/.test(chk), chk);
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('Z50').getFormula(), '', 'проба прибрана');
});

t('rules, bank log, main (labels the bot reads), checks, charts, schedules, protection', () => {
  const env = built(), c = env.ctx;
  const r = env.ss.getSheetByName('Правила').getRange('A5:E5').getValues()[0];
  assert.strictEqual(r[0], 'так'); assert.ok(r[4]);
  const cats = new Set(env.ss.getSheetByName('Довідник').getRange('A5:A84').getValues().map(x => x[0]));
  env.ss.getSheetByName('Правила').getRange('E5:E40').getValues().filter(x => x[0]).forEach(x => assert.ok(cats.has(x[0]), 'категорія правила є в довіднику: ' + x[0]));
  const main = env.ss.getSheetByName('Головне'), labels = main.getRange('A4:A40').getValues().map(x => x[0]);
  ['Доступно моїх сьогодні', 'Можна на день', 'На картці зараз', 'Прогноз на кінець місяця (за планом)', 'Чернеток чекають підтвердження', 'Межа «близько до ліміту»']
    .forEach(l => assert.ok(labels.indexOf(l) >= 0, l));
  assert.strictEqual((main.charts || []).length, 0, 'графіки переїхали на «Charts»');
  assert.strictEqual(env.ss.getSheetByName('Charts').charts.length, 11, '«Charts»: усі графіки (з модулем ліміту)');
  assert.strictEqual(env.ss.getSheetByName('charts_mob').charts.length, 8, '«charts_mob»: вісім графіків');
  assert.ok(!env.ss.getSheetByName('📱 Телефон'), '«📱 Телефон» більше не будується');
  assert.ok(/Доступно моїх сьогодні/.test(env.ss.getSheetByName('Журнал').getRange('B1').getFormula()));
  const chk = env.ss.getSheetByName('Перевірка');
  assert.ok(/ПЕРЕВІРИТИ/.test(chk.getRange('B2').getFormula()) && chk.getRange('B5').getFormula());
  const h = env.triggers.map(t => t.getHandlerFunction()).sort();
  assert.strictEqual(JSON.stringify(h), JSON.stringify(['mbsBackup', 'mbsEdit', 'mbsTick']));
  const j = env.ss.getSheetByName('Журнал');
  assert.ok(j.prots.length === 1 && j.prots[0].warn && j.prots[0].unprot[0] === 'A5:H3004');
  assert.ok(env.ss.getSheetByName('Довідник').prots.length === 0, 'довідник вільний для вводу');
  env.calls.menu = null; c.onOpen();
  assert.ok(env.calls.menu.length > 15 && env.calls.menu.every(fn => typeof c[fn] === 'function'), 'кожен пункт меню має функцію');
});

t('setup again (e.g. in a copy of the template): data sheets untouched, only missing pieces added', () => {
  const env = built();
  const d = env.ss.getSheetByName('Довідник'); d.getRange('A5').setValue('Моя категорія');
  env.ss.getSheetByName('Журнал').getRange('A5:C5').setValues([[new Date(), 100, 'Моя категорія']]);
  env.triggers.splice(0, env.triggers.length);                              // копія файлу: розклади не копіюються
  env.docProps.deleteProperty('schema_v'); env.docProps.deleteProperty('mbs');  // і властивості документа теж можуть не скопіюватись
  env.ctx.setupTemplate();
  assert.strictEqual(d.getRange('A5').getValue(), 'Моя категорія');
  assert.strictEqual(env.ss.getSheetByName('Журнал').getRange('B5').getValue(), 100);
  assert.strictEqual(env.triggers.length, 3);
  assert.strictEqual(env.docProps.getProperty('mbs'), 'monitoring-budgets-template', 'мітка відновлена з прихованої клітинки');
  assert.strictEqual(env.files.length, 1, 'перед змінами в таблиці з даними: копія');
});

console.log(n + ' tests passed');
