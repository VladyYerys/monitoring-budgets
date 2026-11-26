// Графіки: «Дані графіків» (схований), сітка на «Головному», «📱 Телефон»; модуль «🗂 Службові аркуші»; оновлення таблиці версії 7 до 8.
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
const st = env => env.ss.getSheetByName('Старт');
function toggle(env, id, v) {
  const r = st(env).getRange(26 + K(env.ctx, 'MODS').findIndex(m => m.id === id), 2);
  r.setValue(v);
  env.ctx.mbsEdit({ range: r });
}
const titles = sh => (sh.charts || []).map(ch => ch.opts.title);
const SERVICE = ['Правила', 'Рахунки', 'Залишки банку', 'Банк', 'Історія місяців', 'Історія категорій', 'Перевірка'];

t('«Дані графіків»: hidden, headers, formulas with gaps (NA), types, holes, groups from «Довідник», top-8 reasons, category helper', () => {
  const env = setup(), d = env.ss.getSheetByName('Дані графіків');
  assert.strictEqual(d.isSheetHidden(), true);
  assert.strictEqual(JSON.stringify(d.getRange('A1:E1').getValues()[0]), JSON.stringify(['Дата', 'Факт', 'За планом', 'Як зараз', 'Межа ліміту']));
  assert.strictEqual(d.getRange('A2').getFormula(), "=IF('Дні'!$A$6=\"\",\"\",'Дні'!$A$6)");
  assert.ok(/IF\(\$A2<=TODAY\(\),N\('Дні'!\$E\$6\),NA\(\)\)/.test(d.getRange('B2').getFormula()), d.getRange('B2').getFormula());
  assert.ok(/'Дні'!\$L\$98/.test(d.getRange('C94').getFormula()), 'останній рядок вікна «Днів»');
  assert.strictEqual(d.getRange('G2').getFormula(), '=IF(1>DAY(EOMONTH(TODAY(),0)),"","1")');
  assert.strictEqual(JSON.stringify(d.getRange('L2:L5').getValues().map(r => r[0])), JSON.stringify(['Витрати', 'Заощадження', 'Податки', 'Борги']));
  assert.ok(/SUMIFS\('Журнал'!\$O\$5:\$O\$3004,'Журнал'!\$I\$5:\$I\$3004,\$L2,/.test(d.getRange('M2').getFormula()) && /"<>рахунок ФОП"/.test(d.getRange('M2').getFormula()));
  assert.ok(/"\?\*"/.test(d.getRange('P2').getFormula()), '«чорні діри»: непорожня діра');
  assert.ok(/CHOOSE\(MONTH\(EDATE\(DATE\(YEAR\(TODAY\(\)\),MONTH\(TODAY\(\)\),1\),-5\)\),"січ"/.test(d.getRange('O2').getFormula()));
  assert.strictEqual(d.getRange('R2').getFormula(), "='Довідник'!$I$5&\"\"");
  assert.ok(/MAXIFS/.test(d.getRange('Z2').getFormula()) && d.getRange('Z8').getFormula() === '=NA()', 'борг: факт до цього місяця');
  assert.ok(d.getRange('AB2').getFormula() === '=NA()' && /\$Z\$10-N\('Налаштування'!\$B\$6\)\*2\)$/.test(d.getRange('AB8').getFormula()), d.getRange('AB8').getFormula());
  assert.ok(/EOMONTH\(TODAY\(\),-1\)/.test(d.getRange('Z10').getFormula()), 'база: борг на кінець минулого місяця');
  const groups = d.getRange('AD2:AD13').getValues().map(r => r[0]).filter(Boolean);
  const dict = env.ss.getSheetByName('Довідник').getRange('A5:C84').getValues().filter(r => r[1] === 'Витрати' && r[2]).map(r => r[2]);
  assert.ok(groups.length >= 3 && groups.every(g => dict.indexOf(g) >= 0) && new Set(groups).size === groups.length, groups.join());
  assert.ok(/SUMIFS\(\$BC\$2:\$BC\$81,\$BB\$2:\$BB\$81,\$AD2\)/.test(d.getRange('AE2').getFormula()));
  assert.strictEqual(d.getRange('AH22').getFormula(), '="без причини"');
  assert.ok(/LARGE\(\$AJ\$2:\$AJ\$22,8\)/.test(d.getRange('AL9').getFormula()));
  assert.ok(/'Довідник'!\$B\$5<>"Витрати"/.test(d.getRange('BC2').getFormula()));
});

t('«Головне»: ten charts in a two-column grid from row 18 (pie with a hole, stacked columns); credit module off → seven, back on → ten; your own chart stays', () => {
  const env = setup(), m = env.ss.getSheetByName('Головне');
  assert.strictEqual(m.charts.length, 10);
  assert.strictEqual(JSON.stringify(m.charts.slice(0, 4).map(ch => [ch.row, ch.col])), JSON.stringify([[18, 5], [18, 10], [34, 5], [34, 10]]));
  const by = s => m.charts.find(ch => ch.opts.title.indexOf(s) >= 0);
  assert.strictEqual(by('за типами').type, 'PIE'); assert.strictEqual(by('за типами').opts.pieHole, 0.5);
  assert.strictEqual(by('з моїх чи з ліміту').opts.isStacked, true);
  assert.strictEqual(by('Групи').type, 'BAR');
  assert.ok(by('межа ліміту').ranges.indexOf('Дані графіків!E1:E94') >= 0);
  assert.ok(by('план і факт').ranges[0] === 'Головне!E3:G15', 'місяці: таблиця на «Головному»');
  m.insertChart(m.newChart().setChartType('LINE').addRange(m.getRange('A1:B5')).setPosition(1, 1, 0, 0).setOption('title', 'Мій графік').build());
  toggle(env, 'credit', 'ні');
  assert.strictEqual(m.charts.length, 8, '7 наших + ваш');
  assert.ok(titles(m).indexOf('Мій графік') >= 0);
  assert.ok(!titles(m).some(x => /ліміт|Борг|Чому/.test(x)), titles(m).join(' | '));
  assert.ok(by('факт і прогноз').ranges.every(r => !/E1:E94/.test(r)), 'без межі ліміту');
  assert.strictEqual(env.ss.getSheetByName('Дані графіків').getRange('W1').getValue(), 'Понад мої');
  assert.strictEqual(env.ss.getSheetByName('📱 Телефон').charts.length, 6);
  toggle(env, 'credit', 'так');
  assert.strictEqual(m.charts.length, 11);
  assert.strictEqual(env.ss.getSheetByName('📱 Телефон').charts.length, 8);
});

t('«📱 Телефон»: key figures from «Головне», three worst budgets (only with the budgets module), charts one under another', () => {
  const env = setup(), p = env.ss.getSheetByName('📱 Телефон');
  const a = p.getRange('A1:A20').getValues().map(r => r[0]);
  ['💚 Мої сьогодні', '📅 Можна на день', '💳 На картці', '🔴 Борг за лімітом', '🛒 Витрати місяця', '💳 Грейс: внести до строку', '💰 Бюджети: найгірші три'].forEach(l => assert.ok(a.indexOf(l) >= 0, l));
  assert.ok(/MATCH\("Доступно моїх сьогодні",'Головне'!\$A\$1:\$A\$120,0\)/.test(p.getRange('B4').getFormula()));
  assert.ok(p.charts.every(ch => ch.col === 1 && ch.row >= 21) && p.charts.every(ch => ch.opts.width === 340));
  assert.ok(!titles(p).some(x => /Групи|Чому/.test(x)), 'лише-настільні графіки не на телефоні');
  toggle(env, 'budgets', 'ні');
  assert.ok(p.getRange('A1:A20').getValues().every(r => r[0] !== '💰 Бюджети: найгірші три'));
  toggle(env, 'credit', 'ні');
  assert.ok(p.getRange('A1:A20').getValues().every(r => r[0] !== '🔴 Борг за лімітом'));
});

t('module «🗂 Службові аркуші»: off by default (service sheets hidden, they still work), on → shown, off → hidden; «Дані графіків» always hidden', () => {
  const env = setup();
  const hidden = () => SERVICE.map(s => env.ss.getSheetByName(s).isSheetHidden());
  assert.ok(hidden().every(Boolean));
  toggle(env, 'service', 'так');
  assert.ok(hidden().every(x => !x));
  assert.strictEqual(env.ss.getSheetByName('Дані графіків').isSheetHidden(), true);
  assert.ok(env.ss.getSheetByName('Старт').getRange('A6:C25').getValues().every(r => !/сховано/.test(r[2])), 'підказка «сховано» зникла');
  toggle(env, 'service', 'ні');
  assert.ok(hidden().every(Boolean));
  env.ctx.startRefresh_(env.ss);
  assert.ok(env.ss.getSheetByName('Старт').getRange('A6:C25').getValues().some(r => /сховано/.test(r[2])), 'де шукати сховане');
});

t('upgrade of a version-7 table to 8: backup first, «Регулярні», charts and phone added, old charts replaced, module row added without touching chosen ones, tabs reordered, your sheet kept', () => {
  const env = setup(), c = env.ctx, ss = env.ss, s = st(env);
  ['Регулярні', 'Дані графіків', '📱 Телефон'].forEach(x => ss.deleteSheet(ss.getSheetByName(x)));
  s.getRange('A34:C34').clearContent();
  s.getRange('B26').setValue('ні');                                          // Telegram: вибір людини
  env.docProps.setProperty('mods', JSON.stringify(Object.assign(JSON.parse(env.docProps.getProperty('mods')), { tg: false })));
  const mods = JSON.parse(env.docProps.getProperty('mods')); delete mods.service; env.docProps.setProperty('mods', JSON.stringify(mods));
  const m = ss.getSheetByName('Головне'); m.charts = [];
  ['📈 На картці по днях: факт, за планом, як зараз', '📊 Витрати по місяцях: план і факт'].forEach(x =>
    m.insertChart(m.newChart().setChartType('LINE').addRange(m.getRange('E3:G15')).setPosition(20, 5, 0, 0).setOption('title', x).build()));
  SERVICE.forEach(x => ss.getSheetByName(x).showSheet());
  ss.insertSheet('Мій аркуш').getRange('A1').setValue('моє');
  env.docProps.setProperty('schema_v', '7');
  const f0 = env.files.length;
  assert.strictEqual(c.schemaTick_(ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), '8');
  assert.strictEqual(env.files.length, f0 + 1, 'копія перед змінами');
  assert.ok(['Регулярні', 'Дані графіків', '📱 Телефон'].every(x => !!ss.getSheetByName(x)));
  assert.strictEqual(m.charts.length, 10, 'старі два замінено сіткою');
  assert.strictEqual(s.getRange('A34').getValue(), '🗂 Службові аркуші'); assert.strictEqual(s.getRange('B34').getValue(), 'ні');
  assert.strictEqual(s.getRange('B26').getValue(), 'ні', 'вибір людини не змінився');
  assert.ok(SERVICE.every(x => ss.getSheetByName(x).isSheetHidden()), 'службові сховано');
  const names = ss.getSheets().map(x => x.getName());
  assert.deepStrictEqual(names.slice(0, 7).join(), ['Старт', 'Головне', 'Журнал', 'Регулярні', 'Місяці', 'Дні', '📱 Телефон'].join());
  assert.ok(names.indexOf('Мій аркуш') > names.indexOf('Дані графіків') && ss.getSheetByName('Мій аркуш').getRange('A1').getValue() === 'моє', 'ваш аркуш на місці');
  assert.ok(ss.getSheetByName('Регулярні').prots.some(p => p.unprot.indexOf('A5:I54') >= 0), 'захист: правила вільні');
  assert.strictEqual(c.schemaTick_(ss), false, 'вдруге нічого');
});

t('groups refresh once a day: a renamed group in «Довідник» reaches the chart data', () => {
  const env = setup(), c = env.ctx, d = env.ss.getSheetByName('Довідник');
  const i = d.getRange('A5:B84').getValues().findIndex(r => r[1] === 'Витрати');
  d.getRange(5 + i, 3).setValue('Моя група');
  env.docProps.deleteProperty('ch_day');
  assert.strictEqual(c.chartsIfDue_(env.ss), true);
  assert.ok(env.ss.getSheetByName('Дані графіків').getRange('AD2:AD13').getValues().some(r => r[0] === 'Моя група'));
  assert.strictEqual(c.chartsIfDue_(env.ss), null, 'того ж дня вдруге: ні');
});

console.log(n + ' tests passed');
