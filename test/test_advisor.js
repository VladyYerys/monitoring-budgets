// Радник Gemini: знімок фінансів (зміст, рівні деталей, приватність), «Мої правила», документ на Диску, щоденне оновлення, навичка (SKILL.md, повідомлення для чату).
const assert = require('assert'), fs = require('fs'), path = require('path'), vm = require('vm');
const { makeEnv } = require('./gasmock');
const { panelSetup } = require('./panel');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const K = (c, s) => vm.runInContext(s, c);

/** Таблиця з панелі «🚀 Підключення» (сім'я, ФОП, демо-місяць) і модулем радника. */
function setup(opts) {
  const env = makeEnv(opts || {});
  return panelSetup(env, { family: 'Олена, Іван', credit: true, bday: 25, fop: [['Олена', '3']], mods: ['tg', 'credit', 'debts', 'budgets', 'backup', 'fop', 'advisor'], demo: true });
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
  ['українською', 'За знімком від', 'Не давай інвестиційних порад', 'Ніколи не проси номери карток', 'Де ми переплачуємо за останні 3 місяці?', 'Які підписки варто скасувати?']
    .forEach(x => assert.ok(md.indexOf(x) >= 0, x));
  assert.ok(md.indexOf('—') < 0, 'без довгих тире');
});

t('menu: «Як підключити навичку» shows the snapshot link, steps, texts to copy and the SKILL.md link; «Оновити знімок» turns the module on', () => {
  const env = makeEnv(), c = env.ctx;
  env.ss.insertSheet('Аркуш1'); c.setupTemplate();
  c.advisorNow();
  let d = env.calls.dialogs.pop();
  assert.ok(/Знімка ще немає/.test(d.html) && /<textarea/.test(d.html) && /gemini\/monitoring-budgets-advisor\/SKILL\.md/.test(d.html) && /Create manually/.test(d.html), d.html);
  assert.ok(/📋 Скопіювати для Gemini/.test(d.html) && /Створи нову навичку/.test(d.html) && /advSkillCopied/.test(d.html) && /Мої правила<\/b>: ще не заповнено/.test(d.html), d.html);
  assert.doesNotThrow(() => new vm.Script(d.html.split('<script>')[1].split('</script>')[0]), 'скрипт вікна без синтаксичних помилок');
  c.advisorSnapshotNow();
  assert.strictEqual(c.modOn_('advisor'), true);
  assert.ok(/Знімок оновлено/.test(env.calls.dialogs.pop().t));
  c.advisorNow();
  d = env.calls.dialogs.pop();
  assert.ok(/https:\/\/docs\.google\.com\/document\/d\//.test(d.html));
  c.tableStatus();
  assert.ok(/Знімок оновлено/.test(env.calls.dialogs.pop().html));
});

t('«Мої правила»: block on «Налаштування» (B27…B32) with lists; empty rules: hint in the snapshot; skill version and «Де що в таблиці» in the snapshot', () => {
  const env = setup(), c = env.ctx, set = env.ss.getSheetByName('Налаштування');
  assert.strictEqual(set.getRange('A26').getValue(), 'Мої правила для радника Gemini');
  assert.strictEqual(set.getRange('B27').getValue(), 'не вказано');
  assert.ok(set.getRange('B27').getDataValidation(), 'список пріоритетів');
  assert.strictEqual(set.getRange('B30').getValue(), 'так');
  const r = c.advRules_(env.ss);
  assert.strictEqual(r.filled, false); assert.strictEqual(r.debtsMust, true);
  calc(env);
  const s = c.snapshotText_(env.ss);
  ['Навичка: версія ' + K(c, 'ADV_SKILL.v'), '## Мої правила', 'Розстрочки й кредити: обов\'язкові платежі', 'Більше правил ще немає', '«🚀 Підключення», крок 4',
    '## Де що в таблиці', '«Журнал»: кожна операція рядком: A дата, B сума', 'F план на місяць', '«Регулярні»', 'рядки 27–32', '«Борги»', '«ФОП»']
    .forEach(x => assert.ok(s.indexOf(x) >= 0, x + '\n' + s.slice(0, 3000)));
  assert.ok(s.indexOf('## Мої правила') < s.indexOf('## Головне на '), 'правила одразу після «Як читати»');
  assert.ok(s.indexOf('## Де що в таблиці') < s.indexOf('## Чого тут немає'));
  assert.ok(s.indexOf('—') < 0, 'без довгих тире');
});

t('«Мої правила» saved from the panel: percents, floor, installments, FOP, free text in the snapshot; card numbers, IBAN and tokens refused or masked; 500 characters max', () => {
  const env = setup(), c = env.ctx, set = env.ss.getSheetByName('Налаштування');
  const r = c.advRulesSave_(env.ss, { prio: 'зібрати подушку безпеки', cushion: 10, floor: 3000, debtsMust: false, fopPct: 5, rules: 'На подарунки не більше 3 000 ₴.\nТаксі лише вночі.' });
  assert.strictEqual(set.getRange('B27').getValue(), 'зібрати подушку безпеки');
  assert.strictEqual(set.getRange('B28').getValue(), 0.1); assert.strictEqual(set.getRange('B29').getValue(), 3000);
  assert.strictEqual(set.getRange('B30').getValue(), 'ні'); assert.strictEqual(set.getRange('B31').getValue(), 0.05);
  assert.ok(r.filled && r.cushion === 0.1 && r.fopPct === 0.05 && r.debtsMust === false);
  calc(env);
  const s = c.snapshotText_(env.ss);
  ['Пріоритет №1: зібрати подушку безпеки.', 'Подушка: відкладати 10 % від кожного доходу.', 'Не опускатися нижче: 3 000 грн на картці', 'Розстрочки й кредити: можна переносити',
    'ФОП: відкладати 5 % з доходу ФОП', 'Ще: На подарунки не більше 3 000 ₴.; Таксі лише вночі.'].forEach(x => assert.ok(s.indexOf(x) >= 0, x + '\n' + s.slice(0, 2500)));
  assert.ok(s.indexOf('Більше правил ще немає') < 0);
  // «10 %», набране просто в клітинку як 10: теж 10 %
  set.getRange('B28').setValue(10);
  assert.strictEqual(c.advRules_(env.ss).cushion, 0.1);
  // номер картки, IBAN, токен: з панелі не записуються
  ['Картка 4441 1144 4111 4444', 'IBAN UA' + '21' + '0'.repeat(25), 'токен uXyZ1234567890abcdefGHIJKLmnop'].forEach(x =>
    assert.throws(() => c.advRulesSave_(env.ss, { rules: x }), /номер картки, IBAN чи токен/, x));
  assert.throws(() => c.advRulesSave_(env.ss, { rules: 'а'.repeat(501) }), /до 500 знаків/);
  assert.strictEqual(set.getRange('B32').getValue(), 'На подарунки не більше 3 000 ₴.\nТаксі лише вночі.', 'відмова нічого не змінила');
  // вписали в клітинку самі: у знімку довгі числа й ключі стають «…»
  set.getRange('B32').setValue('Моя картка 4441114441114444, ключ uXyZ1234567890abcdefGHIJKLmnop');
  const s2 = c.snapshotText_(env.ss);
  assert.ok(s2.indexOf('4441114441114444') < 0 && s2.indexOf('uXyZ1234567890') < 0 && s2.indexOf('Ще: Моя картка …, ключ …') >= 0, s2.slice(0, 2500));
});

t('upgrade of a version-10 table to 11: «Мої правила» added under the advisor block, your settings kept, snapshot rewritten; a table that already has the block: untouched', () => {
  const env = setup(), c = env.ctx, set = env.ss.getSheetByName('Налаштування');
  set.getRange('A26:B33').setValues(Array.from({ length: 8 }, () => ['', '']));
  set.getRange('B5').setValue(20); set.getRange('B23').setValue('підсумки');
  const docs0 = env.calls.docs;
  env.docProps.setProperty('schema_v', '10');
  assert.strictEqual(c.schemaTick_(env.ss), true);
  assert.strictEqual(env.docProps.getProperty('schema_v'), String(c.schemaLast_()));
  assert.strictEqual(set.getRange('A26').getValue(), 'Мої правила для радника Gemini');
  assert.strictEqual(set.getRange('B27').getValue(), 'не вказано');
  assert.strictEqual(set.getRange('B5').getValue(), 20); assert.strictEqual(set.getRange('B23').getValue(), 'підсумки');
  assert.ok(/Мої правила/.test(env.docs[env.docProps.getProperty('adv_doc_id')].paras.map(p => p.t).join('\n')), 'знімок переписано з новим розділом');
  assert.strictEqual(env.calls.docs, docs0, 'той самий документ');
  // блок уже є (і заповнений): крок нічого не чіпає
  set.getRange('B28').setValue(0.15);
  env.docProps.setProperty('schema_v', '10');
  c.schemaTick_(env.ss);
  assert.strictEqual(set.getRange('B28').getValue(), 0.15);
});

t('skill v2: one message for the Gemini chat creates the skill; description with trigger words and YAML-safe; every figure the skill names exists in the snapshot', () => {
  const env = setup(), c = env.ctx, d = K(c, 'ADV_SKILL.description'), msg = c.advSkillMsg_(), ins = c.advSkillInstructions_();
  assert.ok(/^Створи нову навичку \(skill\)/.test(msg) && msg.indexOf('Назва: Радник Monitoring Budgets') >= 0 && msg.indexOf('Опис: ' + d) >= 0 && msg.indexOf(ins) >= 0);
  assert.ok(d.length <= 1024 && d.indexOf(': ') < 0 && d.indexOf('#') < 0, 'опис: до 1024 знаків, без «: » і «#» (YAML)');
  ['бюджет', 'скільки можна витратити', 'зарплат', 'кредитний ліміт', 'підписки', 'борги', 'ФОП'].forEach(x => assert.ok(d.indexOf(x) >= 0, 'опис: ' + x));
  ['Версія навички: 2', 'Готові цифри не перераховуй', '«Мої правила»', 'Де що в таблиці', '## 4. Типові питання', 'Чи вистачить на X?', 'Прийшла зарплата', 'лише читає знімок']
    .forEach(x => assert.ok(ins.indexOf(x) >= 0, x));
  calc(env);
  const m = env.ss.getSheetByName('Головне'), labs = m.getRange('A1:A120').getValues().map(r => String(r[0]));
  const L = ['Можна на день', 'Доступно моїх сьогодні', 'Найнижча точка за планом', 'Чернеток чекають підтвердження', 'Борг за лімітом зараз', 'Орієнтир погашення на місяць',
    'Грейс: внести до строку', 'Грейс: строк'];
  env.ss.getSheetByName('Борги').getRange('A5:F5').setValues([['Розстрочка', 'я винен', 12000, 1000, 15, '']]);
  L.forEach(l => { const r = 1 + labs.indexOf(l); assert.ok(r > 0, '«Головне»: ' + l); m.put(r, 2, l === 'Грейс: строк' ? new Date() : 100, m.cell(r, 2).f); });
  const s = c.snapshotText_(env.ss);
  ['Можна на день', 'Доступно моїх сьогодні', 'Найнижча точка за планом', 'Чернеток чекають підтвердження', 'Борг за лімітом зараз', 'Орієнтир погашення на місяць',
    'Грейс: внести до строку', 'Грейс: строк'].forEach(l => { assert.ok(ins.indexOf('«' + l + '»') >= 0, 'навичка: ' + l); assert.ok(s.indexOf('- ' + l + ': ') >= 0, 'знімок: ' + l); });
  ['Головне', 'Бюджети', 'Місяці', 'Плани', 'Підписки', 'Борги', 'ФОП'].forEach(x => assert.ok(new RegExp('^## ' + x, 'm').test(s) || x === 'Підписки', 'розділ знімка: ' + x));
  assert.ok(ins.indexOf('—') < 0 && msg.indexOf('—') < 0);
});

t('panel step 4: rules saved with the snapshot, the message returned; «скопійовано» remembers the skill version; an older copied skill is flagged', () => {
  const env = setup(), c = env.ctx;
  const a = c.cnAdvisor({ prio: 'швидше закрити борги й розстрочки', cushion: 0, floor: 1500, debtsMust: true, fopPct: 0, rules: '' });
  assert.ok(/^Створи нову навичку/.test(a.msg) && a.v === 2 && a.rules.floor === 1500);
  let st = c.cnState().adv;
  assert.ok(st.rules.filled && st.prio.length === 5 && st.max === 500 && st.copied === 0 && st.old === false && st.fop === true, JSON.stringify(st));
  assert.strictEqual(c.advSkillCopied(), true);
  st = c.cnState().adv;
  assert.ok(st.copied === 2 && st.old === false);
  env.docProps.setProperty('adv_skill_v', '1');
  assert.strictEqual(c.cnState().adv.old, true, 'навичку v1 скопіювали раніше: панель підкаже скопіювати ще раз');
  assert.throws(() => c.cnAdvisor({ rules: '4441 1144 4111 4444' }), /номер картки/);
});

console.log(n + ' tests passed');
