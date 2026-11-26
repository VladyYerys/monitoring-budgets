// Оновлення: чи вийшла нова версія (публічний version.json, раз на день, без даних людей) і «📦 Перенести дані зі старої копії».
const assert = require('assert'), fs = require('fs'), path = require('path'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);
const COPY = 'https://docs.google.com/spreadsheets/d/' + 'X'.repeat(30) + '/copy';
const ver = (v, extra) => (url, o) => (/version\.json$/.test(url) ? { code: 200, body: Object.assign({ version: v, copy: COPY, notes: ['модулі', 'радник'] }, extra || {}) } : { code: 404, body: '' });

function built(opts) {
  const env = makeEnv(opts || {});
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  return env;
}

t('version compare and version.json in the repository matches the code version', () => {
  const env = makeEnv(), c = env.ctx;
  assert.strictEqual(c.verCmp_('0.4.0', '0.4.0'), 0); assert.strictEqual(c.verCmp_('0.10.0', '0.9.9'), 1); assert.strictEqual(c.verCmp_('0.3.9', '0.4.0'), -1);
  const j = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'version.json'), 'utf8'));
  assert.strictEqual(j.version, K(c, 'TEMPLATE_VERSION'));
  assert.strictEqual(j.version, JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'package.json'), 'utf8')).version);
  assert.ok(Array.isArray(j.notes) && j.notes.length > 0);
  assert.ok(j.copy === '' || /^https:\/\/docs\.google\.com\/spreadsheets\/d\/[\w-]+\/copy/.test(j.copy), 'copy: порожньо або посилання «Зробити копію»');
  assert.strictEqual(K(c, 'VERSION_URL'), 'https://raw.githubusercontent.com/VladyYerys/monitoring-budgets/main/version.json');
});

t('newer version: «🩺 Стан таблиці» shows it with the copy link and how to move data; once a day only; the request carries nothing about the table', () => {
  const env = built({ mono: ver('9.0.0') }), c = env.ctx;
  c.tableStatus();
  const h = env.calls.dialogs.pop().html;
  assert.ok(/Вийшла версія 9\.0\.0/.test(h) && h.indexOf(COPY) >= 0 && /Перенести дані зі старої копії/.test(h) && /модулі; радник/.test(h), h);
  const calls = env.calls.fetch.filter(u => /version\.json/.test(u));
  assert.ok(calls.length === 1 && calls[0] === K(c, 'VERSION_URL'), 'рівно ця адреса, без параметрів');
  c.tableStatus();
  assert.strictEqual(env.calls.fetch.filter(u => /version\.json/.test(u)).length, 1, 'кеш: удруге не питає');
  env.docProps.deleteProperty('ver_day');
  c.verIfDue_(env.ss); c.verIfDue_(env.ss);
  assert.strictEqual(env.calls.fetch.filter(u => /version\.json/.test(u)).length, 2, 'розклад: раз на день');
  const r1 = c.buildMorning_(env.ss, {}), r2 = c.buildMorning_(env.ss, {});
  assert.ok(/🆕 Вийшла версія шаблону 9\.0\.0/.test(r1.text) && !/🆕/.test(r2.text), 'у Telegram один раз');
});

t('same version, broken file, no network, foreign copy link: calm text, no crash, copy link not shown', () => {
  let env = built({ mono: ver(vm.runInContext('TEMPLATE_VERSION', makeEnv().ctx)) });
  env.ctx.versionNow();
  assert.ok(/у вас найновіша/.test(env.calls.dialogs.pop().html));
  env = built({ mono: ver('9.0.0', { copy: 'https://evil.example/copy' }) });
  env.ctx.versionNow();
  let h = env.calls.dialogs.pop().html;
  assert.ok(/Вийшла версія 9\.0\.0/.test(h) && h.indexOf('evil') < 0, h);
  env = built({ mono: () => ({ code: 200, body: 'not json' }) });
  env.ctx.versionNow();
  h = env.calls.dialogs.pop().html;
  assert.ok(/не читається/.test(h), h);
  env = built({ mono: () => { throw new Error('offline'); } });
  env.ctx.tableStatus();
  assert.ok(/немає зв'язку з GitHub/.test(env.calls.dialogs.pop().html));
});

/** Стара копія: демо, сім'я, свої налаштування, правило, борг, модулі, рядок історії. */
function oldCopy() {
  const env = makeEnv({ promptText: ['2', 'Олена, Іван', '2', '20', '', '', '1, 4, 6', '', 'так'] }), c = env.ctx;
  env.ss.insertSheet('Аркуш1'); c.startWizard();
  env.ss.getSheetByName('Правила').getRange('A100:E100').setValues([['так', 'моя кавʼярня', 'опис банку', 'списання', 'Кава та перекуси']]);
  c.modSet_(env.ss, 'debts', true);
  env.ss.getSheetByName('Борги').getRange('A5:F5').setValues([['Розстрочка', 'я винен', 12000, 1000, 15, '']]);
  c.debtsPlansSync_(env.ss);                                            // плани боргу вже в журналі старої копії
  env.ss.getSheetByName('Налаштування').getRange('B10').setValue(0.02);
  env.ss.getSheetByName('Налаштування').getRange('B24').setValue('ні');
  const j = env.ss.getSheetByName('Журнал');
  j.getRange('Y5:AA5').setValues([['bank-op-1', 'Опис з банку', -250]]);
  env.ss.getSheetByName('Історія місяців').getRange('A4:B4').setValues([['2026-10', 13402]]);
  return env;
}
const OLD_ID = 'OLD_COPY_' + 'a'.repeat(20);
const snap = env => env.ss.getSheets().map(s => s.getName() + ':' + JSON.stringify(Array.from(s.cells.entries()))).join('|');

t('move data: journal with bank fields, dictionary, rules, settings by label, debts, modules by name, history; the old copy only read; check OK', () => {
  const old = oldCopy(), before = snap(old);
  const env = built({ promptText: 'https://docs.google.com/spreadsheets/d/' + OLD_ID + '/edit#gid=0' }), c = env.ctx;
  env.addSpreadsheet(OLD_ID, old.ss);
  c.migrateFromOld();
  const d = env.calls.dialogs.pop();
  assert.ok(/Дані перенесено/.test(d.t), d.t + d.html);
  const val = (e, sh, a) => JSON.stringify(e.ss.getSheetByName(sh).getRange(a).getValues());
  const same = (sh, a, msg) => assert.ok(val(env, sh, a) === val(old, sh, a), msg || sh + ' ' + a);
  same('Журнал', 'A5:H400', 'журнал A…H'); same('Журнал', 'Y5:AF400', 'банківські поля, черговість, рахунок');
  same('Довідник', 'A5:K84'); same('Правила', 'A5:G204'); same('Борги', 'A5:H24'); same('Залишки банку', 'A5:C40');
  assert.strictEqual(val(env, 'Історія місяців', 'A4:B4'), JSON.stringify([['2026-10', 13402]]));
  const set = env.ss.getSheetByName('Налаштування');
  assert.strictEqual(set.getRange('B13').getValue(), 'сім\'я'); assert.strictEqual(set.getRange('B14').getValue(), 'Олена');
  assert.strictEqual(set.getRange('B7').getValue(), 'дебетова'); assert.strictEqual(set.getRange('B5').getValue(), 20);
  assert.strictEqual(set.getRange('B10').getValue(), 0.02); assert.strictEqual(set.getRange('B24').getValue(), 'ні');
  assert.strictEqual(set.getRange('B9').getValue(), K(c, 'TEMPLATE_VERSION'), 'версію не переносимо');
  const mods = JSON.parse(env.docProps.getProperty('mods'));
  assert.strictEqual(mods.debts, true, 'модулі за назвою'); assert.strictEqual(mods.credit, false); assert.strictEqual(mods.tg, true);
  assert.strictEqual(env.ss.getSheetByName('Борги').isSheetHidden(), false);
  assert.ok(/Підключіть заново/.test(d.html) && /Перевірка/.test(d.html));
  assert.strictEqual(snap(old), before, 'стара копія не змінилась');
  assert.ok(!env.userProps.getProperty('mono_token') && !env.userProps.getProperty('tg_token'), 'токени не переносяться');
});

t('move data refuses: the personal table (sheets «m» and «y»), a non-template table, not a link; a used journal asks and backs up first', () => {
  const env = built({ promptText: 'https://docs.google.com/spreadsheets/d/' + 'P'.repeat(25) + '/edit' }), c = env.ctx;
  const personal = makeEnv(); personal.ss.insertSheet('m'); personal.ss.insertSheet('y');
  env.addSpreadsheet('P'.repeat(25), personal.ss);
  c.migrateFromOld();
  assert.ok(/особисту таблицю/.test(env.calls.dialogs.pop().html));
  const other = makeEnv(); other.ss.insertSheet('Аркуш1').getRange('A1').setValue('щось');
  const e2 = built({ promptText: 'Q'.repeat(25) }); e2.addSpreadsheet('Q'.repeat(25), other.ss);
  e2.ctx.migrateFromOld();
  assert.ok(/не копія шаблону/.test(e2.calls.dialogs.pop().html));
  const e3 = built({ promptText: 'привіт' }); e3.ctx.migrateFromOld();
  assert.ok(/не схоже на посилання/.test(e3.calls.dialogs.pop().html));
  const old = oldCopy();
  const e4 = built({ promptText: OLD_ID, alertAnswer: 'NO' });
  e4.addSpreadsheet(OLD_ID, old.ss);
  e4.ss.getSheetByName('Журнал').getRange('A5:E5').setValues([[new Date(), 77, 'Продукти', 'новий рядок', '✅ факт']]);
  e4.ctx.migrateFromOld();
  assert.strictEqual(e4.ctx.readJournal_(e4.ss).filter(o => !o.blank).length, 1, '«Ні»: нічого не змінено');
  const e5 = built({ promptText: OLD_ID, alertAnswer: 'YES' });
  e5.addSpreadsheet(OLD_ID, old.ss);
  e5.ss.getSheetByName('Журнал').getRange('A5:E5').setValues([[new Date(), 77, 'Продукти', 'новий рядок', '✅ факт']]);
  const f0 = e5.files.length;
  e5.ctx.migrateFromOld();
  assert.ok(e5.files.length > f0, 'копія перед перенесенням');
  assert.ok(/Копія цієї таблиці перед перенесенням/.test(e5.calls.dialogs.pop().html));
});

t('move data from an older copy (no «Борги», no modules, no new settings rows): skipped with a note, defaults of the new copy kept', () => {
  const old = oldCopy();
  old.ss.deleteSheet(old.ss.getSheetByName('Борги'));
  old.ss.getSheetByName('Старт').getRange('A24:C33').clearContent();
  old.ss.getSheetByName('Налаштування').getRange('A10:B10').clearContent();
  old.ss.getSheetByName('Налаштування').getRange('A22:B24').clearContent();
  const env = built({ promptText: OLD_ID }), c = env.ctx;
  env.addSpreadsheet(OLD_ID, old.ss);
  c.migrateFromOld();
  const h = env.calls.dialogs.pop().html;
  assert.ok(/«Борги»: у старій копії немає/.test(h) && /модулів у старій копії ще не було/.test(h), h);
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('B10').getValue(), 0.035);
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('B24').getValue(), 'так');
  assert.strictEqual(env.ss.getSheetByName('Налаштування').getRange('B13').getValue(), 'сім\'я');
});

console.log(n + ' tests passed');
