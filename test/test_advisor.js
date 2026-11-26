// Радник Gemini: знімок фінансів (зміст, рівні деталей, приватність), документ на Диску, щоденне оновлення, навичка (SKILL.md).
const assert = require('assert'), fs = require('fs'), path = require('path'), vm = require('vm');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);

/** Таблиця з майстром (сім'я, демо-місяць) і модулем радника. */
function setup(opts) {
  const env = makeEnv(Object.assign({ promptText: ['2', 'Олена, Іван', '1', '25', 'Олена:3', '', '1, 2, 3, 4, 6, 7, 8', '', 'так'] }, opts || {}));
  env.ss.insertSheet('Аркуш1');
  env.ctx.startWizard();
  return env;
}
const tz = 'Europe/Kiev';
/** Мок формул не рахує: ставимо результати службових колонок журналу (тип I, стан M, факт O, вид рахунку AG), як їх порахували б Таблиці. */
function calc(env) {
  const j = env.ss.getSheetByName('Журнал'), d = env.ss.getSheetByName('Довідник').getRange('A5:B84').getValues(), type = {};
  d.forEach(r => { if (r[0]) type[r[0]] = r[1]; });
  const today = env.Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  for (let r = 5; r <= 400; r++) {
    const v = j.getRange(r, 1, 1, 8).getValues()[0];
    if (v[1] === '' || v[1] === null) continue;
    const key = env.Utilities.formatDate(v[0], tz, 'yyyy-MM-dd'), plan = v[4] === 'план' || v[4] === '🔁 план регулярний' || (v[4] === '' && key > today);
    const m = v[4] === '🏦 чернетка' ? 2 : plan ? 1 : 0;
    [[9, type[v[2]] || ''], [13, m], [15, m >= 1 ? 0 : v[1]], [33, 'основна картка']].forEach(x => j.put(r, x[0], x[1], j.cell(r, x[0]).f));
  }
}

t('snapshot text: date, glossary, main figures, budgets, months by member, plans, FOP; no comments, no bank descriptions except subscription names', () => {
  const env = setup(), c = env.ctx, j = env.ss.getSheetByName('Журнал');
  const today = env.Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), y = c.addDays_(today, -1);
  const free = c.readJournal_(env.ss).find(o => o.blank).row;
  j.getRange(free, 1, 2, 5).setValues([[env.Utilities.parseDate(y, tz, 'yyyy-MM-dd'), 299, 'Підписки', 'мій секретний коментар', '✅ факт'],
    [env.Utilities.parseDate(y, tz, 'yyyy-MM-dd'), 450, 'Продукти', 'ще коментар', '✅ факт']]);
  j.getRange(free, 25, 2, 2).setValues([['op1', 'Netflix 4441114441114444'], ['op2', 'ATB Market Kyiv']]);
  const m = env.ss.getSheetByName('Головне'), mr = l => 1 + m.getRange('A1:A120').getValues().findIndex(r => r[0] === l);
  m.put(mr('Доступно моїх сьогодні'), 2, 12345, m.cell(mr('Доступно моїх сьогодні'), 2).f);
  calc(env);
  const s = c.snapshotText_(env.ss);
  ['# Monitoring Budgets: знімок для радника', 'Дата знімка: ', '## Як читати', 'Бюджетний місяць: дохід з 25-го', '## Головне на ', '- Доступно моїх сьогодні: 12 345 грн',
    '## Місяці (факти)', 'По учасниках (витрати): ', '## Плани до ', '## Підписки', 'Netflix …', '## ФОП', 'Олена: 3 група', '## Чого тут немає']
    .forEach(x => assert.ok(s.indexOf(x) >= 0, x + '\n' + s));
  ['мій секретний коментар', 'ще коментар', '[демо]', 'супермаркет', 'ATB', '4441114441114444', 'token'].forEach(x => assert.ok(s.indexOf(x) < 0, 'не має бути: ' + x));
  assert.ok(/Категорії: .*Продукти \d/.test(s), 'рівень «категорії»');
  assert.ok(s.indexOf('## Операції') < 0);
  assert.ok(!/Кредитний ліміт|Грейс/.test(s) || c.modOn_('credit'), 'ліміт лише з модулем');
});

t('detail levels: «підсумки» without categories and plan list, «операції» with operations (still without comments)', () => {
  const env = setup(), c = env.ctx;
  calc(env);
  const s0 = c.snapshotText_(env.ss, { level: 'підсумки' }), s2 = c.snapshotText_(env.ss, { level: 'операції' });
  assert.ok(s0.indexOf('- Категорії: ') < 0 && /надходжень за планом .* списань/.test(s0), s0);
  assert.ok(/## Операції за 60 днів \(без коментарів\)\n- \d\d\.\d\d\.\d{4} · /.test(s2), s2.slice(-2000));
  assert.ok(s2.indexOf('[демо]') < 0 && s2.indexOf('супермаркет') < 0);
  env.ss.getSheetByName('Налаштування').getRange('B23').setValue('підсумки');
  assert.strictEqual(c.advLevel_(env.ss), 'підсумки');
});

t('document on Drive: created once in «Monitoring Budgets: радник», rewritten in place with headings; daily refresh only when the module is on', () => {
  const env = setup(), c = env.ctx;
  assert.strictEqual(c.modOn_('advisor'), true);
  const id = env.docProps.getProperty('adv_doc_id');
  assert.ok(id, 'модуль увімкнено майстром: знімок одразу записано');
  const doc = env.docs[id];
  assert.strictEqual(doc.file.parent.name, 'Monitoring Budgets: радник');
  assert.strictEqual(doc.paras[0].h, 'HEADING1');
  assert.ok(doc.paras.some(p => p.t === 'Головне на ' + (() => { const k = env.Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'); return k.slice(8) + '.' + k.slice(5, 7) + '.' + k.slice(0, 4); })() && p.h === 'HEADING2'));
  c.snapshotWrite_(env.ss);
  assert.strictEqual(env.docProps.getProperty('adv_doc_id'), id, 'той самий документ');
  assert.strictEqual(env.calls.docs, 1);
  env.docProps.deleteProperty('adv_day');
  const hour = Number(env.Utilities.formatDate(new Date(), tz, 'HH'));
  const r = c.advisorIfDue_(env.ss);
  assert.ok(hour < 5 ? r === null : r && r.chars > 500);
  if (hour >= 5) assert.strictEqual(c.advisorIfDue_(env.ss), null, 'раз на день');
  c.modSet_(env.ss, 'advisor', false);
  env.docProps.deleteProperty('adv_day');
  assert.strictEqual(c.advisorIfDue_(env.ss), null, 'модуль вимкнено');
  const e2 = setup({ docFail: 'Drive недоступний' });
  e2.docProps.deleteProperty('adv_day');
  assert.strictEqual(e2.ctx.advisorIfDue_(e2.ss), null);
  if (hour >= 5) assert.ok(/Drive недоступний/.test(JSON.parse(e2.docProps.getProperty('adv_err')).msg));
});

t('Gemini skill: SKILL.md in the repository equals the text in the code; role, Ukrainian, snapshot date, no investment advice, starter questions', () => {
  const env = makeEnv(), c = env.ctx;
  const md = c.advSkillMd_(), file = path.join(__dirname, '..', 'gemini', 'monitoring-budgets-advisor', 'SKILL.md');
  assert.strictEqual(fs.readFileSync(file, 'utf8'), md, 'оновіть файл: node tools/skill.js');
  assert.ok(/^---\nname: monitoring-budgets-advisor\ndescription: .+\n---\n/.test(md));
  ['українською', 'дату знімка', 'Не давай інвестиційних порад', 'Ніколи не проси номери карток', 'Де ми переплачуємо за останні 3 місяці?', 'Які підписки варто скасувати?']
    .forEach(x => assert.ok(md.indexOf(x) >= 0, x));
  assert.ok(md.indexOf('—') < 0, 'без довгих тире');
});

t('menu: «Як підключити навичку» shows the snapshot link, steps, texts to copy and the SKILL.md link; «Оновити знімок» turns the module on', () => {
  const env = makeEnv(), c = env.ctx;
  env.ss.insertSheet('Аркуш1'); c.setupTemplate();
  c.advisorNow();
  let d = env.calls.dialogs.pop();
  assert.ok(/Знімка ще немає/.test(d.html) && /<textarea/.test(d.html) && /gemini\/monitoring-budgets-advisor\/SKILL\.md/.test(d.html) && /Create manually/.test(d.html), d.html);
  c.advisorSnapshotNow();
  assert.strictEqual(c.modOn_('advisor'), true);
  assert.ok(/Знімок оновлено/.test(env.calls.dialogs.pop().t));
  c.advisorNow();
  d = env.calls.dialogs.pop();
  assert.ok(/https:\/\/docs\.google\.com\/document\/d\//.test(d.html));
  c.tableStatus();
  assert.ok(/Знімок оновлено/.test(env.calls.dialogs.pop().html));
});

console.log(n + ' tests passed');
