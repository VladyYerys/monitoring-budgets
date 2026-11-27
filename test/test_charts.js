// Дашборди «Charts» і «charts_mob» (темні, як в особистій таблиці), «Дані графіків» (схований); модуль «🗂 Службові аркуші»; оновлення таблиці версії 8 до 9.
const assert = require('assert'), vm = require('vm'), fs = require('fs'), path = require('path');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);
const same = (a, b, m) => assert.strictEqual(JSON.stringify(a), JSON.stringify(b), m);

function setup() {
  const env = makeEnv();
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  return env;
}
const st = env => env.ss.getSheetByName('⚡ Швидке підключення');
function toggle(env, id, v) {
  const r = env.ss.getSheetByName('Налаштування').getRange(44 + K(env.ctx, 'MODS').findIndex(m => m.id === id), 2);
  r.setValue(v);
  env.ctx.mbsEdit({ range: r });
}
const titles = sh => (sh.charts || []).map(ch => ch.opts.title);
const by = (sh, s) => sh.charts.find(ch => ch.opts.title.indexOf(s) >= 0);
const SERVICE = ['Місяці', 'Дні', 'Довідник', 'Налаштування', 'Головне', 'Правила', 'Рахунки', 'Залишки банку', 'Банк', 'Історія місяців', 'Історія категорій', 'Перевірка'];
const VISIBLE = ['⚡ Швидке підключення', 'Журнал', 'Регулярні', 'Charts', 'charts_mob'];

t('«Дані графіків»: hidden; key figures by name; categories, «що зросло», groups, types, 60 days, pace, forecast, holes with limits from «Charts», reasons', () => {
  const env = setup(), c = env.ctx, d = env.ss.getSheetByName('Дані графіків');
  assert.strictEqual(d.isSheetHidden(), true);
  assert.strictEqual(d.getRange('A1').getValue(), 'Показник');
  const keys = c.chartsKeys_(), names = K(c, 'CH_KEYS');
  assert.ok(keys.length === names.length && keys.every((k, i) => k[0] === names[i]), 'ключі й CH_KEYS в одному порядку');
  assert.strictEqual(d.getRange('A2').getValue(), 'Сьогодні'); assert.strictEqual(d.getRange('B2').getFormula(), '=TODAY()');
  assert.strictEqual(c.chK_('daysIn'), '$B$13');
  assert.ok(d.getRange('B2:B60').getFormulas().every(r => !/\{\{/.test(r[0])), 'усі {{ключі}} замінено');
  assert.ok(/MINIFS\('Дні'!\$L\$6:\$L\$98,'Дні'!\$A\$6:\$A\$98,">="&\$B\$2\)/.test(d.getRange(c.chK_('low').slice(1).replace(/\$/g, '')).getFormula()), 'найнижча точка: від сьогодні');
  // категорії
  assert.strictEqual(d.getRange('D2').getFormula(), "='Довідник'!$A$5&\"\"");
  const g2 = d.getRange('G2').getFormula();
  assert.ok(/SUMIFS\('Журнал'!\$O\$5:\$O\$3004,'Журнал'!\$C\$5:\$C\$3004,\$D2,'Журнал'!\$A\$5:\$A\$3004,">="&\$B\$3,'Журнал'!\$A\$5:\$A\$3004,"<="&\$B\$2,/.test(g2) && /"<>рахунок ФОП"/.test(g2), g2);
  assert.ok(/">="&\$B\$7,.*"<="&\$B\$8/.test(d.getRange('I2').getFormula()), 'квартал');
  assert.strictEqual(d.getRange('L2').getFormula(), '=IF(AND($E2="Витрати",N($K2)>0.005),$K2+ROW()/1E9,"")');
  assert.ok(/Поки що жодна категорія/.test(d.getRange('O2').getFormula()) && !/Поки/.test(d.getRange('O3').getFormula()));
  // групи: назви пише скрипт, за спаданням
  const groups = d.getRange('U2:U13').getValues().map(r => r[0]).filter(Boolean);
  const dict = env.ss.getSheetByName('Довідник').getRange('A5:C84').getValues().filter(r => r[1] === 'Витрати' && r[2]).map(r => r[2]);
  assert.ok(groups.length >= 3 && groups.every(g => dict.indexOf(g) >= 0) && new Set(groups).size === groups.length, groups.join());
  assert.strictEqual(d.getRange('U14').getValue(), 'Усі витрати');
  assert.ok(/LARGE\(\$Z\$2:\$Z\$13,1\)/.test(d.getRange('AC2').getFormula()));
  same(d.getRange('AJ2:AJ7').getValues().map(r => r[0]), ['Витрати', 'Борги', 'Заощадження', 'Податки', 'Разом списано (4 типи)', 'Надходження']);
  assert.strictEqual(d.getRange('AK6').getFormula(), '=SUM(AK2:AK5)');
  // 60 днів, темп, прогноз
  assert.strictEqual(d.getRange('AQ2').getFormula(), '=$B$2-59'); assert.strictEqual(d.getRange('AQ61').getFormula(), '=$B$2');
  assert.ok(/'Дні'!\$E\$6:\$E\$98/.test(d.getRange('AR2').getFormula()) && /MINIFS/.test(d.getRange('AR2').getFormula()), 'до вікна «Днів»: назад від контрольної точки');
  assert.ok(/NA\(\)/.test(d.getRange('AU2').getFormula()));
  assert.strictEqual(d.getRange('AX2').getFormula(), '=IF(1>$B$13,"",1)');
  assert.strictEqual(d.getRange('BC2').getFormula(), '=IF($B$2-30>EOMONTH($B$2,1),"",$B$2-30)');
  assert.strictEqual(d.getRange('BC32').getFormula(), '=IF($B$2>EOMONTH($B$2,1),"",$B$2)');
  assert.ok(/\$B\$2\+61/.test(d.getRange('BC93').getFormula()));
  // «чорні діри»: ліміт з «Charts», статус
  assert.strictEqual(d.getRange('BQ2').getFormula(), "='Довідник'!$I$5&\"\"");
  assert.strictEqual(d.getRange('BW2').getFormula(), "=IF(N('Charts'!$H$106)>0,'Charts'!$H$106,\"\")");
  assert.ok(/▲▲ понад/.test(d.getRange('BZ2').getFormula()) && /○ без бази/.test(d.getRange('BZ2').getFormula()));
  assert.strictEqual(d.getRange('BQ10').getValue(), 'Усі непомітні витрати');
  assert.strictEqual(d.getRange('CD1').getFormula(), '=$BQ$2&""');
  // причини
  assert.strictEqual(d.getRange('CM22').getFormula(), '="без причини"');
  assert.strictEqual(d.getRange('CM23').getValue(), 'Разом з ліміту');
  assert.ok(/LARGE\(\$CQ\$2:\$CQ\$22,21\)/.test(d.getRange('CS22').getFormula()));
});

t('«Charts»: Grafana-dark tiles, sections and eleven charts in the personal layout; no em dash anywhere', () => {
  const env = setup(), sh = env.ss.getSheetByName('Charts');
  assert.strictEqual(sh.getRange('B1').getValue(), 'Фінанси зараз');
  assert.ok(/\$B\$2/.test(sh.getRange('P1').getFormula()));
  same(['B4', 'F4', 'J4', 'N4'].map(a => sh.getRange(a).getValue()),
    ['ЗАЛИШОК НА КІНЕЦЬ ДНЯ', 'ВИТРАЧЕНО ЦЬОГО МІСЯЦЯ', 'ДОХІД ЦЬОГО МІСЯЦЯ', 'ДОСТУПНО З КРЕДИТНИМ ЛІМІТОМ (як у банку)']);
  assert.ok(/'Дані графіків'!\$B\$16/.test(sh.getRange('B5').getFormula()), 'залишок на кінець дня');
  assert.ok(sh.merges.indexOf('B5:E6') >= 0 && sh.merges.indexOf('B4:E4') >= 0, 'плитка: заголовок і велике число об\'єднано');
  assert.strictEqual(sh.noGrid, true);
  assert.strictEqual(sh.getRange('H105').getValue(), 'Ліміт на місяць ✎');
  assert.ok(!sh.getRange('H106').getFormula(), 'ліміти: ваші числа');
  assert.ok(/'Дані графіків'!\$BW\$10/.test(sh.getRange('H114').getFormula()));
  assert.strictEqual(sh.charts.length, 11);
  const pos = sh.charts.map(ch => [ch.row, ch.col, ch.opts.width, ch.opts.height]);
  same(pos.slice(0, 4), [[10, 2, 1408, 420], [31, 2, 690, 399], [31, 10, 690, 399], [53, 2, 690, 546]]);
  const b60 = by(sh, 'останні 60 днів');
  assert.strictEqual(b60.type, 'COMBO'); assert.strictEqual(b60.opts.seriesType, 'area'); assert.strictEqual(b60.opts.areaOpacity, 0.3);
  same(b60.opts.series[2].lineDashStyle, [6, 4], 'межа ліміту пунктиром');
  assert.strictEqual(b60.opts.backgroundColor, '#181B1F'); assert.strictEqual(b60.opts.titleTextStyle.color, '#D8D9DA');
  const pie = by(sh, 'Цей місяць за типами');
  assert.strictEqual(pie.type, 'PIE'); assert.strictEqual(pie.opts.pieHole, 0.55);
  same(pie.opts.colors, ['#3987E5', '#D95926', '#199E70', '#C98500']);
  assert.ok(by(sh, 'Витрати: з моїх чи з ліміту').opts.isStacked);
  assert.strictEqual(by(sh, 'Витрати за групами').type, 'BAR');
  const holes6 = by(sh, 'по місяцях');
  assert.strictEqual(holes6.ranges[0], 'Дані графіків!CC1:CI7', 'шість «чорних дір» з «Довідника»');
  assert.ok(by(sh, 'Залишок: факт').ranges.indexOf('Дані графіків!BF1:BF93') >= 0);
  assert.ok(sh.charts.every(ch => ch.opts.title.indexOf('—') < 0));
  ['27_Charts.gs', '28_Dash.gs'].forEach(f => assert.ok(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8').indexOf('—') < 0, f + ': без тире «—»'));
  assert.ok(sh.prots.some(p => p.unprot.indexOf('H106:H113') >= 0), 'захист: ліміти «чорних дір» вільні');
});

t('«charts_mob»: three narrow columns for the phone, short chart titles 336 px wide', () => {
  const env = setup(), m = env.ss.getSheetByName('charts_mob');
  same(['B1', 'G1', 'L1'].map(a => m.getRange(a).getValue()), ['Фінанси зараз', 'Непомітні витрати', 'Мої гроші й ліміт']);
  assert.strictEqual(m.getRange('B3').getValue(), 'ЗАЛИШОК НА КІНЕЦЬ ДНЯ');
  assert.ok(/UPPER/.test(m.getRange('B11').getFormula()));
  assert.strictEqual(m.charts.length, 8);
  assert.ok(m.charts.every(ch => ch.opts.width === 336 && [2, 7, 12].indexOf(ch.col) >= 0));
  same(m.charts.map(ch => ch.opts.title).slice(0, 2), ['Залишок на кінець дня: 60 днів', 'Цей місяць за типами']);
  assert.strictEqual(by(m, 'Цей місяць за типами').opts.legend.position, 'bottom');
  assert.strictEqual(m.getRange('L18').getValue(), 'Чому з ліміту');
});

t('credit module off → no limit tiles, charts or reasons (8 + 6 charts); on → back (11 + 8); your own chart and hole limits stay', () => {
  const env = setup(), sh = env.ss.getSheetByName('Charts'), m = env.ss.getSheetByName('charts_mob');
  sh.getRange('H106').setValue(1500);
  sh.insertChart(sh.newChart().setChartType('LINE').addRange(sh.getRange('A1:B5')).setPosition(1, 1, 0, 0).setOption('title', 'Мій графік').build());
  toggle(env, 'credit', 'ні');
  assert.strictEqual(sh.charts.length, 9, '8 наших + ваш');
  assert.strictEqual(m.charts.length, 6);
  assert.ok(titles(sh).indexOf('Мій графік') >= 0);
  assert.ok(!titles(sh).concat(titles(m)).some(x => /ліміт|Борг|Чому/.test(x)), titles(sh).join(' | '));
  assert.strictEqual(sh.getRange('N4').getValue(), '≈ МОЖНА НА ДЕНЬ ДО КІНЦЯ МІСЯЦЯ');
  assert.strictEqual(sh.getRange('B147').getValue(), 'Мої гроші й прогноз');
  assert.strictEqual(sh.getMaxRows(), 177);
  assert.strictEqual(m.getRange('L1').getValue(), 'Мої гроші й прогноз');
  assert.ok(!/Чому/.test(m.getRange('L18').getValue()));
  assert.strictEqual(sh.getRange('H106').getValue(), 1500, 'ліміт «чорної діри» зберігся');
  toggle(env, 'credit', 'так');
  assert.strictEqual(sh.charts.length, 12);
  assert.strictEqual(m.charts.length, 8);
  assert.strictEqual(sh.getRange('H106').getValue(), 1500);
  assert.strictEqual(sh.getMaxRows(), 221);
});

t('module «🗂 Усі аркуші»: off by default (only ⚡ Швидке підключення, Журнал, Регулярні, Charts, charts_mob on view; the rest works hidden), on → all shown (ФОП and Борги only with their modules), off → hidden', () => {
  const env = setup();
  const hidden = () => SERVICE.map(s => env.ss.getSheetByName(s).isSheetHidden());
  const shown = () => env.ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName());
  assert.ok(hidden().every(Boolean));
  same(shown(), VISIBLE);
  toggle(env, 'fop', 'так');
  same(shown(), VISIBLE, 'ФОП увімкнено, але аркуш схований разом з рештою');
  toggle(env, 'service', 'так');
  assert.ok(hidden().every(x => !x));
  assert.strictEqual(env.ss.getSheetByName('ФОП').isSheetHidden(), false, 'ФОП: модуль увімкнено');
  assert.strictEqual(env.ss.getSheetByName('Борги').isSheetHidden(), true, 'Борги: модуль вимкнено');
  assert.strictEqual(env.ss.getSheetByName('Дані графіків').isSheetHidden(), true);
  toggle(env, 'service', 'ні');
  assert.ok(hidden().every(Boolean));
  same(shown(), VISIBLE);
  env.ctx.startRefresh_(env.ss);
  assert.ok(env.ss.getSheetByName('⚡ Швидке підключення').getRange('A7:C21').getValues().some(r => r[0] === 'Модулі увімкнено' && !/Усі аркуші/.test(r[1])), 'стан: «Усі аркуші» вимкнено');
});

t('upgrade of a version-8 table to 9: backup first, «Charts» and «charts_mob» built, old chart data rebuilt, «📱 Телефон» removed, old charts off «Головне» (yours kept)', () => {
  const env = setup(), c = env.ctx, ss = env.ss;
  ['Charts', 'charts_mob'].forEach(x => ss.deleteSheet(ss.getSheetByName(x)));
  const d = ss.getSheetByName('Дані графіків'); d.clear(); d.getRange('A1:E1').setValues([['Дата', 'Факт', 'За планом', 'Як зараз', 'Межа ліміту']]);
  const ph = ss.insertSheet('📱 Телефон'), main = ss.getSheetByName('Головне');
  main.showSheet();
  ['🍩 Цей місяць за типами', '📊 Витрати по місяцях: план і факт', 'Мій графік'].forEach(x =>
    main.insertChart(main.newChart().setChartType('LINE').addRange(main.getRange('E3:G15')).setPosition(20, 5, 0, 0).setOption('title', x).build()));
  ph.insertChart(ph.newChart().setChartType('PIE').addRange(main.getRange('E3:G15')).setPosition(21, 1, 0, 0).setOption('title', '🍩 Цей місяць за типами').build());
  ss.insertSheet('Мій аркуш').getRange('A1').setValue('моє');
  env.docProps.setProperty('schema_v', '8');
  const f0 = env.files.length;
  assert.strictEqual(c.schemaTick_(ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.strictEqual(env.files.length, f0 + 1, 'копія перед змінами');
  assert.ok(!ss.getSheetByName('📱 Телефон'));
  assert.strictEqual(d.getRange('A1').getValue(), 'Показник', '«Дані графіків» нової будови');
  assert.strictEqual(ss.getSheetByName('Charts').charts.length, 11);
  assert.strictEqual(ss.getSheetByName('charts_mob').charts.length, 8);
  same(titles(main), ['Мій графік'], 'наші старі графіки прибрано, ваш лишився');
  assert.strictEqual(main.isSheetHidden(), true, '«Головне» тепер службовий');
  const names = ss.getSheets().map(x => x.getName());
  same(names.slice(0, 5), VISIBLE);
  assert.ok(names.indexOf('Мій аркуш') > names.indexOf('Дані графіків') && ss.getSheetByName('Мій аркуш').getRange('A1').getValue() === 'моє', 'ваш аркуш на місці');
  assert.strictEqual(c.schemaTick_(ss), false, 'вдруге нічого');
});

t('upgrade of a version-9 table to 10: module «🗂 Службові аркуші» becomes «🗂 Усі аркуші» with your choice kept; Місяці, Дні, Довідник, Налаштування hidden; order Старт, Журнал, Регулярні, Charts, charts_mob', () => {
  const env = setup(), c = env.ctx, ss = env.ss, i = K(c, 'MODS').findIndex(m => m.id === 'service'), row = 44 + i;
  env.ss.getSheetByName('Налаштування').getRange(row, 1).setValue('🗂 Службові аркуші');
  ['Місяці', 'Дні', 'Довідник', 'Налаштування'].forEach(n => ss.getSheetByName(n).showSheet());
  env.docProps.setProperty('schema_v', '9');
  assert.strictEqual(c.schemaTick_(ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange(row, 1).getValue(), '🗂 Усі аркуші');
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange(row, 2).getValue(), 'ні', 'ваш вибір лишився');
  assert.ok(/Довідник/.test(env.ss.getSheetByName('Налаштування').getRange(row, 3).getValue()));
  same(ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName()), VISIBLE);
  same(ss.getSheets().map(s => s.getName()).slice(0, 5), VISIBLE);
  // «так» на старій таблиці: усе видно й після кроку
  const env2 = setup(), r2 = 44 + K(env2.ctx, 'MODS').findIndex(m => m.id === 'service');
  toggle(env2, 'service', 'так');
  env2.ss.getSheetByName('Налаштування').getRange(r2, 1).setValue('🗂 Службові аркуші');
  env2.docProps.setProperty('schema_v', '9');
  env2.ctx.schemaTick_(env2.ss);
  assert.strictEqual(env2.ss.getSheetByName('Налаштування').getRange(r2, 2).getValue(), 'так');
  assert.ok(SERVICE.every(n => !env2.ss.getSheetByName(n).isSheetHidden()));
});

t('once a day: a renamed group reaches the data; a new «чорна діра» in «Довідник» rebuilds the charts with one more series', () => {
  const env = setup(), c = env.ctx, d = env.ss.getSheetByName('Довідник');
  const i = d.getRange('A5:B84').getValues().findIndex(r => r[1] === 'Витрати');
  d.getRange(5 + i, 3).setValue('Моя група');
  d.getRange('I11').setValue('🎮 Ігри');
  env.docProps.deleteProperty('ch_day');
  assert.strictEqual(c.chartsIfDue_(env.ss), true);
  assert.ok(env.ss.getSheetByName('Дані графіків').getRange('U2:U13').getValues().some(r => r[0] === 'Моя група'));
  assert.strictEqual(by(env.ss.getSheetByName('Charts'), 'по місяцях').ranges[0], 'Дані графіків!CC1:CJ7', 'сьома «чорна діра»');
  assert.strictEqual(c.chartsIfDue_(env.ss), null, 'того ж дня вдруге: ні');
});

console.log(n + ' tests passed');
