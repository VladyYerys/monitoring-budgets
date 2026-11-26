/** 41_Steps.gs: Версія таблиці кроками: збирання з нуля, оновлення після нового коду, «🩺 Стан таблиці». */

/*
 * Будова таблиці: кроки з номерами (STEPS). Таблиця пам'ятає, до якого кроку дійшла (властивість документа schema_v).
 * «🚀 Почати» у порожній таблиці проходить усі кроки. Після нового коду mbsTick (кожні 15 хв) бачить нові кроки,
 * робить копію таблиці й застосовує їх по черзі. Кожен крок спершу перевіряє, чи вже зроблено (done), і будує лише те,
 * чого бракує: аркуші з вашими даними (журнал, довідник, налаштування, залишки, правила) ніколи не перебудовуються.
 * Новий крок: функція тут або в окремому файлі, рядок у STEPS.
 */

const STEPS = [
  { v: 1, title: '📒 Основа: журнал, довідник, налаштування, залишки банку, місяці, дні, історія',
    auto: ss => stepBase_(ss), done: ss => [SETTINGS_SHEET, DICT_SHEET, CP_SHEET, JOURNAL, MONTHS_SHEET, DAYS_SHEET, HISTORY_SHEET, H_CAT_SHEET].every(n => !!ss.getSheetByName(n)) },
  { v: 2, title: '🏦 Банк: «Правила» і «Банк»', auto: ss => stepBank_(ss), done: ss => !!ss.getSheetByName(RULES_SHEET) && !!ss.getSheetByName(BANK_SHEET) },
  { v: 3, title: '🧭 «Головне» з графіками, «Перевірка», шапка журналу', auto: ss => stepMain_(ss), done: ss => !!ss.getSheetByName(MAIN_SHEET) && !!ss.getSheetByName(CHECK_SHEET) },
  { v: 4, title: '⏱ Розклади: порядок журналу, нічні копії; захист з попередженням', auto: ss => stepSchedules_(ss),
    done: ss => hasTrigger_(TRG.tick) && hasTrigger_(TRG.edit) && (!modOn_('backup') || hasTrigger_(TRG.backup)) &&
      (!modOn_('backup') || PropertiesService.getDocumentProperties().getProperty('prot') === 'off' || protectionOn_(ss)) },
  { v: 5, title: '🚀 «Старт», «Рахунки», сім\'я, «ФОП», виписки', auto: ss => stepFamily_(ss),
    done: ss => [START_SHEET, ACC_SHEET, FOP_SHEET].every(n => !!ss.getSheetByName(n)) && !!ss.getSheetByName(JOURNAL) &&
      journalHasAccounts_(ss) && String(ss.getSheetByName(SETTINGS_SHEET).getRange('A12').getValue()) === 'Сім\'я' },
  { v: 6, title: '🧩 Модулі на «Старті», «Борги», грейс Mono, радник Gemini, оновлення', auto: ss => stepModules_(ss),
    done: ss => !!ss.getSheetByName(DEBTS_SHEET) && startHasModules_(ss) && !!PropertiesService.getDocumentProperties().getProperty('mods') &&
      mainHasLabel_(ss, 'Грейс: строк') &&
      String(ss.getSheetByName(SETTINGS_SHEET).getRange('A22').getValue()) === 'Радник і нагадування' },
];
/** Остання версія таблиці в цьому коді. */
function schemaLast_() { return STEPS[STEPS.length - 1].v; }
const SCHEMA_RETRY_MS = 6 * 3600 * 1000;
const SHEET_ORDER = [START_SHEET, MAIN_SHEET, JOURNAL, MONTHS_SHEET, DAYS_SHEET, FOP_SHEET, DEBTS_SHEET, DICT_SHEET, RULES_SHEET, ACC_SHEET, CP_SHEET, BANK_SHEET, HISTORY_SHEET,
  H_CAT_SHEET, CHECK_SHEET, SETTINGS_SHEET];

function buildIfMissing_(ss, name, fn, out) { if (!ss.getSheetByName(name)) { fn(ss); out.push('«' + name + '» ✓'); } }

function stepBase_(ss) {
  const out = [];
  buildIfMissing_(ss, SETTINGS_SHEET, buildSettings_, out);
  buildIfMissing_(ss, DICT_SHEET, buildDictionary_, out);
  buildIfMissing_(ss, ACC_SHEET, buildAccounts_, out);
  buildIfMissing_(ss, CP_SHEET, buildCheckpoints_, out);
  buildIfMissing_(ss, JOURNAL, buildJournal_, out);
  buildIfMissing_(ss, MONTHS_SHEET, buildMonths_, out);
  buildIfMissing_(ss, DAYS_SHEET, buildDays_, out);
  if (!ss.getSheetByName(HISTORY_SHEET) || !ss.getSheetByName(H_CAT_SHEET)) { buildHistory_(ss); out.push('«' + HISTORY_SHEET + '» ✓'); }
  return out.length ? ['Додано: ' + out.join(', ') + '.'] : ['Усе вже було.'];
}

function stepBank_(ss) {
  const out = [];
  buildIfMissing_(ss, RULES_SHEET, buildRules_, out);
  buildIfMissing_(ss, BANK_SHEET, buildBankLog_, out);
  return out.length ? ['Додано: ' + out.join(', ') + '. Підключити Mono: «⚙️ Налаштування → 🏦 Підключити Mono».'] : ['Усе вже було.'];
}

function stepMain_(ss) {
  const out = [];
  buildIfMissing_(ss, DEBTS_SHEET, buildDebts_, out);            // «Головне» рахує залишки боргів
  buildIfMissing_(ss, MAIN_SHEET, buildMain_, out);
  buildIfMissing_(ss, FOP_SHEET, buildFop_, out);
  buildIfMissing_(ss, CHECK_SHEET, buildCheck_, out);
  buildIfMissing_(ss, START_SHEET, buildStart_, out);
  return out.length ? ['Додано: ' + out.join(', ') + '.'] : ['Усе вже було.'];
}

function stepSchedules_(ss) {
  ensureTickTriggers_(ss);
  const bk = modOn_('backup');
  if (bk && !hasTrigger_(TRG.backup)) ensureBackupTrigger_(ss);
  const dp = PropertiesService.getDocumentProperties(), lines = ['Порядок журналу: після кожної правки і кожні ' + TICK_MIN + ' хв ✓',
    bk ? 'Нічні копії: щодня о ' + pad2_(BACKUP_HOUR) + ':00 ✓' : 'Нічні копії: модуль вимкнено'];
  if (bk && dp.getProperty('prot') !== 'off') { const r = applyProtection_(ss); lines.push('Захист з попередженням: ' + r.done.length + ' аркушів ✓' + (r.skipped.length ? ' (пропущено: ' + r.skipped.join('; ') + ')' : '')); }
  return lines;
}

/**
 * Версія 5 для вже зібраної таблиці: нові аркуші, колонки журналу «Рахунок» і службові формули, блок «Сім'я», «Чий бюджет»,
 * правила податків; «Місяці», «Головне» й «Перевірка» перебудовуються (там лише формули; рік і «Хто» на «Місяцях» лишаються).
 * Ваші дані (журнал A…H, довідник, правила, залишки, налаштування) не змінюються.
 */
function stepFamily_(ss) {
  const out = [].concat(settingsUpgrade_(ss), dictUpgrade_(ss), rulesUpgrade_(ss));
  buildIfMissing_(ss, ACC_SHEET, buildAccounts_, out);
  buildIfMissing_(ss, FOP_SHEET, buildFop_, out);
  const j = ss.getSheetByName(JOURNAL);
  if (j && !journalHasAccounts_(ss)) { buildJournal_(ss); out.push('журнал: колонка «Рахунок» і службові формули'); }
  const b = ss.getSheetByName(BANK_SHEET);
  if (b && String(b.getRange('L4').getValue()) !== 'Рахунок') { ensureSize_(b, 0, BANK_LOG_N).getRange('L4').setValue('Рахунок').setFontWeight('bold'); }
  if (ss.getSheetByName(MONTHS_SHEET)) { buildMonths_(ss); out.push('«' + MONTHS_SHEET + '»: фільтр «Хто», без рахунків ФОП'); }
  if (ss.getSheetByName(MAIN_SHEET)) buildMain_(ss);
  if (ss.getSheetByName(CHECK_SHEET)) { buildCheck_(ss); out.push('«' + CHECK_SHEET + '»: ФОП і рахунки'); }
  buildIfMissing_(ss, START_SHEET, buildStart_, out);
  if (PropertiesService.getDocumentProperties().getProperty('prot') !== 'off') applyProtection_(ss);
  return out.length ? ['Додано: ' + out.join(', ') + '.'] : ['Усе вже було.'];
}

/**
 * Версія 6: модулі на «Старті» (стан з того, що вже працює), аркуш «Борги», рядки грейсу й боргів на «Головному»,
 * ставка комісії, радник і нагадування на «Налаштуваннях». Ваші дані не змінюються. Наприкінці модулі застосовуються.
 */
function stepModules_(ss) {
  const out = [].concat(settingsUpgrade6_(ss));
  buildIfMissing_(ss, DEBTS_SHEET, buildDebts_, out);
  const st = ss.getSheetByName(START_SHEET);
  if (st && !startHasModules_(ss)) { ensureSize_(st, MOD_FIRST + MODS.length + 5, 4); startModulesBlock_(ss, st); out.push('блок «🧩 Модулі» на «' + START_SHEET + '»'); }
  if (ss.getSheetByName(MAIN_SHEET) && !mainHasLabel_(ss, 'Грейс: строк')) { buildMain_(ss); out.push('«' + MAIN_SHEET + '»: грейс Mono й борги'); }
  if (PropertiesService.getDocumentProperties().getProperty('prot') !== 'off' && modOn_('backup')) applyProtection_(ss);
  const lines = modulesApply_(ss, true);
  return (out.length ? ['Додано: ' + out.join(', ') + '.'] : ['Усе вже було.']).concat(lines);
}

function mainHasLabel_(ss, label) {
  const sh = ss.getSheetByName(MAIN_SHEET);
  return !!sh && sh.getRange(1, 1, Math.min(Math.max(sh.getLastRow(), 1), 120), 1).getValues().some(r => String(r[0]) === label);
}

/** У журналі вже є колонка «Рахунок» (AF)? Старий журнал коротший: читати за його межами не можна. */
function journalHasAccounts_(ss) {
  const j = ss.getSheetByName(JOURNAL);
  return !!j && j.getMaxColumns() >= colNum_(JC.acct) && String(j.getRange(JC.acct + J_HEAD).getValue()).indexOf('Рахунок') === 0;
}

/** Аркуші в зручному порядку; порожній стартовий аркуш прибрати. */
function tidySheets_(ss) {
  const ours = SHEET_ORDER.filter(n => !!ss.getSheetByName(n));
  ss.getSheets().filter(s => ours.indexOf(s.getName()) < 0 && s.getLastRow() === 0 && s.getLastColumn() === 0).forEach(s => { if (ss.getSheets().length > 1) ss.deleteSheet(s); });
  if (ss.setActiveSheet && ss.moveActiveSheet) ours.forEach((n, i) => { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
}

/* ---------- запуск кроків ---------- */

/** З mbsTick: є нові кроки → копія й кроки по черзі. true: щось застосовано. */
function schemaTick_(ss) {
  const dp = PropertiesService.getDocumentProperties();
  if (Number(dp.getProperty('schema_v') || 0) >= schemaLast_()) return false;
  const fail = jsonProp_(dp, 'schema_fail');
  if (fail && fail.to === schemaLast_() && Date.now() - Number(fail.t || 0) < SCHEMA_RETRY_MS) return false;
  return !!schemaRun_(ss).ran;
}

/**
 * Застосовує кроки після поточної версії. Зроблені лише відмічає. Перед першим справжнім кроком у таблиці з даними: копія.
 * Крок упав: версія лишається на попередньому, причина в «🩺 Стан таблиці» і в Telegram.
 */
function schemaRun_(ss, opt) {
  guard_(ss, !!(opt && opt.allowEmpty));
  const dp = PropertiesService.getDocumentProperties();
  let lock = null;
  try { lock = LockService.getDocumentLock(); } catch (e) { lock = null; }
  if (lock && !lock.tryLock(1000)) return { ran: false, lines: ['⚠️ Оновлення таблиці вже йде, спробуйте за хвилину.'] };
  try {
    dp.deleteProperty('fsep'); FSEP_ = null;                          // регіон таблиці могли змінити: роздільник формул визначимо заново
    const from = Number(dp.getProperty('schema_v') || 0), hadData = !!ss.getSheetByName(JOURNAL);
    const rep = { t: Date.now(), from: from, to: from, steps: [], backup: null, err: '' };
    let ran = false;
    for (const s of STEPS) {
      let done = false;
      try { done = !!s.done(ss); } catch (e) { done = false; }
      if (s.v <= from && done) continue;
      if (done) { rep.steps.push({ v: s.v, title: s.title, lines: ['уже було'] }); }
      else {
        if (hadData && !rep.backup) {
          try { const c = backup_(ss, 'ручна'); rep.backup = { name: c.name, url: c.url }; }
          catch (e) { rep.err = 'копію перед змінами зробити не вдалося (' + String(e && e.message || e) + '), нічого не змінено'; break; }
        }
        try { rep.steps.push({ v: s.v, title: s.title, lines: (s.auto(ss) || []).filter(Boolean) }); ran = true; }
        catch (e) { rep.err = 'крок ' + s.v + ' «' + s.title + '»: ' + String(e && e.message || e); break; }
      }
      if (s.v > rep.to) { rep.to = s.v; dp.setProperty('schema_v', String(s.v)); }
    }
    if (ran && !hadData) { try { tidySheets_(ss); } catch (e) { /* порядок аркушів не критичний */ } }     // лише перше збирання
    dp.setProperties({ schema_last: JSON.stringify(rep), mbs_version: TEMPLATE_VERSION });
    if (rep.err) dp.setProperty('schema_fail', JSON.stringify({ t: rep.t, to: schemaLast_(), msg: rep.err }));
    else dp.deleteProperty('schema_fail');
    if ((ran && hadData) || rep.err) { try { if (tgReady_()) tgSend_(schemaTg_(rep)); } catch (e) { /* звіт і так у «🩺 Стан таблиці» */ } }
    return { ran: ran, rep: rep, lines: schemaReport_(rep) };
  } finally { if (lock) lock.releaseLock(); }
}

function schemaReport_(rep) {
  const out = [(rep.err ? '⚠️ Оновлення зупинилось на версії <b>' : '✅ Таблиця має версію <b>') + rep.to + '</b>' + (rep.from !== rep.to ? ' (було ' + rep.from + ')' : '') + '.'];
  if (rep.backup) out.push('💾 Копія перед змінами: ' + link_(rep.backup.url, rep.backup.name));
  rep.steps.forEach(s => out.push('<b>' + s.v + ' · ' + esc_(s.title) + '</b><br>' + s.lines.map(esc_).join('<br>')));
  if (rep.err) out.push('⚠️ ' + esc_(rep.err) + '<br>Сам спробує ще раз за 6 год або одразу: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».');
  return out;
}

function schemaTg_(rep) {
  const out = [(rep.err ? '⚠️ <b>Оновлення таблиці зупинилось</b> на версії ' : '🛠 <b>Таблицю оновлено</b> до версії ') + rep.to + (rep.from !== rep.to ? ' (було ' + rep.from + ')' : '')];
  rep.steps.forEach(s => out.push('• ' + s.v + ' ' + esc_(s.title) + ': ' + esc_(s.lines[0] || 'готово')));
  if (rep.backup) out.push('💾 Копія перед змінами: ' + esc_(rep.backup.name));
  if (rep.err) out.push('⚠️ ' + esc_(rep.err));
  return out.join('\n');
}

/* ---------- меню ---------- */

function schemaUpdateNow() {
  if (!menuOk_('🛠 Оновлення таблиці')) return;
  show_('🛠 Оновлення таблиці', schemaRun_(SpreadsheetApp.getActive()).lines);
}

/** Усе про стан таблиці в одному вікні. */
function tableStatus() {
  if (!menuOk_('🩺 Стан таблиці')) return;
  const ss = SpreadsheetApp.getActive(), dp = PropertiesService.getDocumentProperties();
  const sec = (t, l) => ['<b>' + t + '</b>'].concat(l.filter(Boolean)).join('<br>');
  const cur = Number(dp.getProperty('schema_v') || 0), fail = jsonProp_(dp, 'schema_fail'), cs = checkState_(ss);
  const ver = ['Код шаблону: ' + TEMPLATE_VERSION + ' · версія таблиці: <b>' + cur + '</b> з ' + schemaLast_() + (cur >= schemaLast_() ? ' ✓' : ': нові кроки застосуються самі протягом 15 хв'),
    verLines_(ss, true).join('<br>'),
    fail && fail.to === schemaLast_() ? '⚠️ Оновлення зупинилось: ' + esc_(fail.msg) : '',
    hasTrigger_(TRG.tick) ? '' : '⚠️ Розкладу немає: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».',
    'Роздільник у формулах: «' + esc_(fsep_(ss)) + '» · часовий пояс: ' + esc_(ss.getSpreadsheetTimeZone()),
    'Код відкритий: ' + link_(REPO_URL, 'github.com/VladyYerys/monitoring-budgets') + ' (MIT).'];
  show_('🩺 Стан таблиці', [sec('🛠 Версія', ver),
    sec('✔️ Перевірка', cs.none ? ['аркуша немає'] : cs.bad.length ? ['⚠️ ' + cs.bad.length + ': ' + cs.bad.map(esc_).join('; ')] : ['усе OK (' + cs.ok + ')']),
    sec('💾 Копії', backupStatusLines_(ss)), sec('🏦 Імпорт з банку', bankStatusLines_(ss)), sec('📨 Telegram', tgStatusLines_(ss)),
    sec('📆 Календар платежів', calLines_(ss)), sec('🛡 Захист', [protectionOn_(ss) ? 'з попередженням ✓' : 'вимкнено']),
    sec('🧩 Модулі', modLines_().concat(['Змінити: «' + START_SHEET + '», блок «🧩 Модулі».'])), sec('🤖 Радник Gemini', advisorLines_(ss))]);
}
