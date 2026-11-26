// Копії (лише свої, ротація 30 / 10), захист з попередженням (чужий не чіпаємо), кроки версій (копія перед змінами, зупинка на помилці).
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

t('backups: own folder (not found by name), 30 nightly + 10 manual kept, other files in the folder untouched', () => {
  const env = setup(), c = env.ctx;
  const folder = c.backupFolder_(env.ss);
  const other = env.ctx.DriveApp.getFileById('SS1').makeCopy('Monitoring Budget: копія 2026-01-01 00:00 (авто)', folder);   // особиста таблиця: інша назва
  const mine = env.files.length;
  for (let i = 0; i < 33; i++) c.mbsBackup();
  for (let i = 0; i < 12; i++) c.backup_(env.ss, 'ручна');
  const live = env.files.filter(f => !f.trashed && f.parent === folder);
  assert.strictEqual(live.filter(f => /\(авто\)$/.test(f.name) && f.name.indexOf('Test: копія') === 0).length, 30);
  assert.strictEqual(live.filter(f => /\(ручна\)$/.test(f.name)).length, 10);
  assert.ok(!other.trashed, 'чужа копія лишилась');
  assert.ok(JSON.parse(env.docProps.getProperty('hb_backup')).ok);
  assert.ok(mine >= 1);
});

t('protection: warning-only on formula sheets, input cells free; your own protection is not touched; remove only ours', () => {
  const env = setup(), c = env.ctx, m = env.ss.getSheetByName('Місяці');
  m.prots.forEach(p => p.remove());
  m.protect().setDescription('моя');
  const r = c.applyProtection_(env.ss);
  assert.ok(r.skipped.some(s => /Місяці/.test(s)));
  assert.strictEqual(m.prots.length, 1);
  assert.strictEqual(m.prots[0].getDescription(), 'моя');
  c.removeProtection();
  assert.strictEqual(env.ss.getSheetByName('Журнал').prots.length, 0);
  assert.strictEqual(m.prots.length, 1, 'власний захист лишився');
  assert.strictEqual(env.docProps.getProperty('prot'), 'off');
  assert.strictEqual(K(c, 'STEPS')[3].done(env.ss), true, 'вимкнений захист не вважається «невиконаним кроком»');
});

t('steps: new code with a new step → backup first, step applied, version up; a failing step stops and is reported', () => {
  const env = setup(), c = env.ctx;
  env.ss.getSheetByName('Журнал').getRange('A5:C5').setValues([[new Date(), 10, 'Продукти']]);
  const steps = K(c, 'STEPS'), v1 = c.schemaLast_() + 1, V1 = String(v1);
  let ran = 0;
  steps.push({ v: v1, title: 'тестовий крок', auto: () => { ran++; return ['готово']; }, done: () => ran > 0 });
  const files0 = env.files.length;
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(ran, 1);
  assert.strictEqual(env.docProps.getProperty('schema_v'), V1);
  assert.strictEqual(env.files.length, files0 + 1, 'копія перед змінами');
  assert.strictEqual(c.schemaTick_(env.ss), false, 'вдруге нічого');
  steps.push({ v: v1 + 1, title: 'зламаний крок', auto: () => { throw new Error('бум'); }, done: () => false });
  c.schemaTick_(env.ss);
  assert.strictEqual(env.docProps.getProperty('schema_v'), V1);
  assert.ok(/бум/.test(JSON.parse(env.docProps.getProperty('schema_fail')).msg));
  assert.strictEqual(c.schemaTick_(env.ss), false, 'після помилки: не частіше ніж раз на 6 год');
});

t('status window shows every section without throwing; support dialog has the jar link only', () => {
  const env = setup(), c = env.ctx;
  c.tableStatus();
  const d = env.calls.dialogs.pop();
  ['Версія', 'Перевірка', 'Копії', 'Імпорт з банку', 'Telegram', 'Календар', 'Захист', 'Модулі', 'Радник Gemini', 'Нова версія'].forEach(s => assert.ok(d.html.indexOf(s) >= 0, s));
  c.supportAuthor();
  assert.ok(/send\.monobank\.ua\/jar\//.test(env.calls.dialogs.pop().html));
});

console.log(n + ' tests passed');
