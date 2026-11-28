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
    done: ss => (!!ss.getSheetByName(START_SHEET) || !!ss.getSheetByName(START_OLD)) && [ACC_SHEET, FOP_SHEET].every(n => !!ss.getSheetByName(n)) && !!ss.getSheetByName(JOURNAL) &&
      journalHasAccounts_(ss) && String(ss.getSheetByName(SETTINGS_SHEET).getRange('A12').getValue()) === 'Сім\'я' },
  { v: 6, title: '🧩 Модулі, «Борги», грейс Mono, радник Gemini, оновлення', auto: ss => stepModules_(ss),
    done: ss => !!ss.getSheetByName(DEBTS_SHEET) && startHasModules_(ss) && !!PropertiesService.getDocumentProperties().getProperty('mods') &&
      mainHasLabel_(ss, 'Грейс: строк') &&
      String(ss.getSheetByName(SETTINGS_SHEET).getRange('A22').getValue()) === 'Радник і нагадування' },
  { v: 7, title: '🧭 Шапка журналу: мої, бюджети зі смужками, прогноз, грейс', auto: ss => stepHeader_(ss), done: ss => headerIsNew_(ss) },
  { v: 8, title: '🗂 Менше вкладок: «Регулярні», «Дані графіків», службові аркуші сховано', auto: ss => stepTabs_(ss),
    done: ss => [REG_SHEET, CH_SHEET].every(n => !!ss.getSheetByName(n)) && modBlockFull_(ss) },
  { v: 9, title: '📊 «Charts» і «charts_mob» (темні дашборди) замість «Головного» й «📱 Телефону»', auto: ss => stepDash_(ss),
    done: ss => [DASH_SHEET, MOB_SHEET].every(n => !!ss.getSheetByName(n)) && !ss.getSheetByName(PHONE_SHEET) && !chartsDataOld_(ss) },
  { v: 10, title: '🗂 На виду лише «Старт», «Журнал», «Регулярні», «Charts» і «charts_mob»; приклади регулярних платежів і боргів', auto: ss => stepVisible_(ss),
    done: ss => (!startHasModules_(ss) || modBlockFull_(ss)) && (!ss.getSheetByName(REG_SHEET) || !!PropertiesService.getDocumentProperties().getProperty('reg_ex')) },
  { v: 11, title: '🤖 «Мої правила» для радника Gemini на «Налаштуваннях»; навичка Gemini версії ' + ADV_SKILL.v, auto: ss => stepRules_(ss),
    done: ss => !ss.getSheetByName(SETTINGS_SHEET) || String(ss.getSheetByName(SETTINGS_SHEET).getRange('A26').getValue()) === ADV_RULES_HEAD },
  { v: 12, title: '🚀 Простіший «Старт»: кроки з номерами панелі; перемикачі модулів у «🚀 Підключення», крок 5', auto: ss => stepStart12_(ss),
    done: ss => !ss.getSheetByName(SETTINGS_SHEET) || (!modOldOnStart_(ss) && !!modBlock_(ss)) },
  { v: 13, title: '🧭 Категорії з банку: зарплата за коментарем, погашення розстрочок, нові мережі й MCC, правила за отримувачем', auto: ss => stepRules13_(ss),
    done: ss => !ss.getSheetByName(RULES_SHEET) || !!PropertiesService.getDocumentProperties().getProperty('rules_v13') },
  { v: 14, title: '✅ Операції з банку з визначеною категорією одразу «✅ факт» (вимкнути: «Налаштування», B35)', auto: ss => settingsAutoFactUpgrade_(ss).map(x => 'Додано: ' + x + '.'),
    done: ss => !ss.getSheetByName(SETTINGS_SHEET) || String(ss.getSheetByName(SETTINGS_SHEET).getRange('A35').getValue()) === AUTO_FACT_LABEL },
  { v: 15, title: '⚡ «Старт» став «Швидким підключенням»: стан таблиці й покрокова інструкція для ШІ (Gemini чи іншого)', auto: ss => stepQuick15_(ss),
    done: ss => !ss.getSheetByName(START_OLD) && !!ss.getSheetByName(START_SHEET) && String(ss.getSheetByName(START_SHEET).getRange('A1').getValue()) === QUICK_TITLE },
  { v: 16, title: '↩ Повернення за покупку: мінус у категорії покупки зменшує її витрати й повертає гроші в «мої»; імпорт з банку розпізнає «Скасування. …»',
    auto: ss => stepRefund16_(ss), done: ss => !ss.getSheetByName(JOURNAL) || refundFormulaOk_(ss) },
  { v: 17, title: '💳 Покупки частинами й розстрочки на «Боргах»: вид, назва в банку, «Закрито», прогрес, ліміти Mono з перевіркою; борги місяця й ліміт у шапці журналу',
    auto: ss => stepParts17_(ss), done: ss => !ss.getSheetByName(DEBTS_SHEET) || debtsNew_(ss) },
  { v: 18, title: '📊 «Регулярні»: шапка-плитки як у «Charts» (собі й інвестиції, підключено, лишається, цей місяць, 30 днів, борги); категорії «Заплати собі» й «Інвестиції»',
    auto: ss => stepReg18_(ss), done: ss => !ss.getSheetByName(REG_SHEET) || regHeaderNew_(ss) },
  { v: 19, title: '🗓 «Вільно» наступного місяця з «моїх» (дохід мінус відкладене, плани й бюджети), «Мін.» лише від сьогодні',
    auto: ss => stepFree19_(ss), done: ss => freeNew_(ss) },
  { v: 20, title: '🎨 «Регулярні»: угорі «Заплати собі» й «Інвестиції», далі ваші рядки, приклади внизу; кольори груп, вимкнені сірі',
    auto: ss => stepRegOrder20_(ss), done: ss => !ss.getSheetByName(REG_SHEET) || String(ss.getSheetByName(REG_SHEET).getRange('P' + REG_FIRST).getFormula()) !== '' },
  { v: 21, title: '🛍 Категорії «' + CAT_PROM + '», «' + CAT_MONO + '», «' + CAT_NP + '», «' + CAT_CASH + '» з правилами; «З Білої картки»: між своїми; надходження від ТОВ: зарплата; правила «' + RULE_SPARE + '»; старі чернетки отримують категорію',
    auto: ss => stepShops21_(ss), done: ss => !!PropertiesService.getDocumentProperties().getProperty('shops_v21') || !ss.getSheetByName(RULES_SHEET) },
  { v: 22, title: '🛡 Журнал без стирання: стерте записане й видалені рядки повертаються; прибрати рядки: меню «' + WIPE_T + '…»',
    auto: ss => stepProtect22_(ss), done: ss => !ss.getSheetByName(JOURNAL) || (!!ss.getSheetByName(SHADOW_SHEET) && hasTrigger_(TRG.change)) },
];
/** Остання версія таблиці в цьому коді. */
function schemaLast_() { return STEPS[STEPS.length - 1].v; }
const SCHEMA_RETRY_MS = 6 * 3600 * 1000;
/** Порядок вкладок: робочі зліва, службові (зазвичай сховані) справа. */
const SHEET_ORDER = [START_SHEET, JOURNAL, REG_SHEET, DASH_SHEET, MOB_SHEET, MONTHS_SHEET, DAYS_SHEET, DICT_SHEET, SETTINGS_SHEET, FOP_SHEET, DEBTS_SHEET,
  MAIN_SHEET, RULES_SHEET, ACC_SHEET, CP_SHEET, BANK_SHEET, HISTORY_SHEET, H_CAT_SHEET, CHECK_SHEET, CH_SHEET, SHADOW_SHEET];

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
  return out.length ? ['Додано: ' + out.join(', ') + '. Підключити банк: «🚀 Підключення», крок 1.'] : ['Усе вже було.'];
}

function stepMain_(ss) {
  const out = [];
  buildIfMissing_(ss, DEBTS_SHEET, buildDebts_, out);            // «Головне» рахує залишки боргів
  buildIfMissing_(ss, MAIN_SHEET, buildMain_, out);
  buildIfMissing_(ss, FOP_SHEET, buildFop_, out);
  buildIfMissing_(ss, CHECK_SHEET, buildCheck_, out);
  if (!ss.getSheetByName(START_OLD)) buildIfMissing_(ss, START_SHEET, buildStart_, out);   // старий «Старт» перейменує крок 15
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
  if (!ss.getSheetByName(START_OLD)) buildIfMissing_(ss, START_SHEET, buildStart_, out);   // старий «Старт» перейменує крок 15
  if (PropertiesService.getDocumentProperties().getProperty('prot') !== 'off') applyProtection_(ss);
  return out.length ? ['Додано: ' + out.join(', ') + '.'] : ['Усе вже було.'];
}

/**
 * Версія 6: модулі (стан з того, що вже працює), аркуш «Борги», рядки грейсу й боргів на «Головному»,
 * ставка комісії, радник і нагадування на «Налаштуваннях». Ваші дані не змінюються. Наприкінці модулі застосовуються.
 */
function stepModules_(ss) {
  const out = [].concat(settingsUpgrade6_(ss));
  buildIfMissing_(ss, DEBTS_SHEET, buildDebts_, out);
  if (ss.getSheetByName(SETTINGS_SHEET) && !startHasModules_(ss)) { modBlockBuild_(ss); out.push('перемикачі модулів (' + MOD_WHERE + ')'); }
  if (ss.getSheetByName(MAIN_SHEET) && !mainHasLabel_(ss, 'Грейс: строк')) { buildMain_(ss); out.push('«' + MAIN_SHEET + '»: грейс Mono й борги'); }
  if (PropertiesService.getDocumentProperties().getProperty('prot') !== 'off' && modOn_('backup')) applyProtection_(ss);
  const lines = modulesApply_(ss, true);
  return (out.length ? ['Додано: ' + out.join(', ') + '.'] : ['Усе вже було.']).concat(lines);
}

/** Версія 7: нова шапка журналу (рядки 1…3) і службові колонки «Довідника» M…U. Дані журналу не змінюються. */
function stepHeader_(ss) {
  if (!ss.getSheetByName(JOURNAL) || !ss.getSheetByName(MAIN_SHEET)) return ['Журналу чи «Головного» ще немає: шапку додасть збирання таблиці.'];
  journalHeader_(ss);
  return ['Шапка журналу: мої й картка, три найгірші бюджети зі смужкою, наступний місяць, ' + (modOn_('credit') ? 'грейс' : 'на день') + '; «№ … ↓» веде на рядок для вводу.'];
}

/**
 * Версія 8: аркуш «Регулярні» (плани з правил), «Дані графіків» (схований), модуль «🗂 Службові аркуші»
 * (типово «ні»: службові аркуші сховано), вкладки в новому порядку. Ваші дані не змінюються.
 */
function stepTabs_(ss) {
  const out = [];
  buildIfMissing_(ss, REG_SHEET, buildRegular_, out);
  if (ss.getSheetByName(DAYS_SHEET) && ss.getSheetByName(DICT_SHEET) && !ss.getSheetByName(CH_SHEET)) { buildChartsData_(ss); out.push('«' + CH_SHEET + '» (схований)'); }
  if (ss.getSheetByName(SETTINGS_SHEET) && !modBlockFull_(ss)) { modBlockAdd_(ss); out.push('модуль «' + modName_('service') + '»'); }
  if (PropertiesService.getDocumentProperties().getProperty('prot') !== 'off' && modOn_('backup')) applyProtection_(ss);
  try { sheetsReorder_(ss); } catch (e) { /* порядок вкладок не критичний */ }
  const lines = modulesApply_(ss, false);
  sheetsVisibility_(ss);
  return (out.length ? ['Додано: ' + out.join(', ') + '.'] : ['Усе вже було.'])
    .concat(lines.filter(l => l.indexOf('Змін немає') < 0), [modOn_('service') ? 'Усі аркуші видно.' : 'Решту аркушів сховано: показати можна модулем «' + modName_('service') + '» (' + MOD_WHERE + ').']);
}

/**
 * Версія 9: «Charts» (комп'ютер) і «charts_mob» (телефон) у темному стилі, як в особистій таблиці: плитки, таблиці, графіки.
 * «Дані графіків» нової будови; «Головне» стає службовим (його цифри читають Telegram і шапка журналу), графіки з нього прибрано;
 * «📱 Телефон» видалено (там були лише формули й графіки). Ваші дані не змінюються.
 */
function stepDash_(ss) {
  if (!ss.getSheetByName(DAYS_SHEET) || !ss.getSheetByName(DICT_SHEET)) return ['«' + DAYS_SHEET + '» чи «' + DICT_SHEET + '» ще немає: дашборди додасть збирання таблиці.'];
  const out = [];
  buildChartsData_(ss);
  const n = dashRefresh_(ss);
  out.push('«' + DASH_SHEET + '» і «' + MOB_SHEET + '»: графіків ' + n);
  const main = ss.getSheetByName(MAIN_SHEET);
  if (main) { dashRemoveOurs_(main, CH_OLD_TITLES); out.push('«' + MAIN_SHEET + '» тепер службовий (графіки переїхали)'); }
  const ph = ss.getSheetByName(PHONE_SHEET);
  if (ph && ss.getSheets().length > 1) { ss.deleteSheet(ph); out.push('«' + PHONE_SHEET + '» видалено: замість нього «' + MOB_SHEET + '»'); }
  if (PropertiesService.getDocumentProperties().getProperty('prot') !== 'off' && modOn_('backup')) applyProtection_(ss);
  try { sheetsReorder_(ss); } catch (e) { /* порядок вкладок не критичний */ }
  sheetsVisibility_(ss);
  try { if (ss.setActiveSheet) ss.setActiveSheet(ss.getSheetByName(DASH_SHEET)); } catch (e) { /* ignore */ }
  return out;
}

/**
 * Версія 10: на виду лише «Старт», «Журнал», «Регулярні», «Charts», «charts_mob» (як в особистій таблиці: журнал, регулярні, графіки);
 * «Місяці», «Дні», «Довідник», «Налаштування», «ФОП», «Борги» сховано разом зі службовими. Модуль «🗂 Службові аркуші»
 * перейменовано на «🗂 Усі аркуші» (ваше «так» / «ні» лишається): «так» показує все.
 * На «Регулярних» приклади з «Увімкнено: ні» (підписки, оренда, бензин, страховка, зарплата, кредит, розстрочка, борги…) у вільних рядках
 * після ваших; на «Боргах» підказка, як додати борг. Ваші рядки й дані не змінюються.
 */
function stepVisible_(ss) {
  const st = ss.getSheetByName(START_SHEET), out = [];
  const nEx = regExamplesAdd_(ss);
  if (nEx) out.push('«' + REG_SHEET + '»: ' + nEx + ' прикладів з «ні» (планів не ставлять)');
  const debts = ss.getSheetByName(DEBTS_SHEET);
  if (debts && String(debts.getRange(3, 1).getValue()) === '') { debts.getRange(3, 1).setValue(DEBT_HINT).setFontColor('#1a73e8').setFontWeight('bold'); out.push('«' + DEBTS_SHEET + '»: як додати борг'); }
  const b = modBlock_(ss);
  if (b) {
    const i = MODS.findIndex(m => m.id === 'service'), m = MODS[i], cell = b.sh.getRange(b.first + i, 1);
    if (String(cell.getValue()) !== m.name) { cell.setValue(m.name); b.sh.getRange(b.first + i, 3).setValue(m.what); out.push('модуль «' + m.name + '» (був «🗂 Службові аркуші»)'); }
    if (!modBlockFull_(ss)) modBlockAdd_(ss);
  }
  try { sheetsReorder_(ss); } catch (e) { /* порядок вкладок не критичний */ }
  try { if (ss.setActiveSheet && st) ss.setActiveSheet(st); } catch (e) { /* активний аркуш не сховати */ }
  sheetsVisibility_(ss);
  try { startRefresh_(ss); } catch (e) { /* чекліст оновиться завтра */ }
  return (out.length ? ['Оновлено: ' + out.join(', ') + '.'] : []).concat([modOn_('service') ? 'Усі аркуші видно (модуль «' + modName_('service') + '» увімкнено).'
    : 'На виду: «' + VISIBLE_SHEETS.join('», «') + '». Решта працює схованою; показати: модуль «' + modName_('service') + '» (' + MOD_WHERE + ').']);
}

/** Версія 11: блок «Мої правила» і новий знімок (розділи «Мої правила», «Де що в таблиці»), якщо радник увімкнений. */
function stepRules_(ss) {
  const out = settingsRulesUpgrade_(ss).map(x => 'Додано: ' + x + '.');
  if (modOn_('advisor')) {
    try { snapshotWrite_(ss); out.push('Знімок для Gemini оновлено.'); } catch (e) { /* знімок оновиться сам уранці */ }
    out.push('Навичка Gemini тепер версії ' + ADV_SKILL.v + ': скопіюйте її ще раз («🚀 Підключення», крок 4) і заповніть «Мої правила».');
  }
  return out;
}

/**
 * Версія 12: блок «🧩 Модулі» переїжджає зі «Старту» на «Налаштування» (рядки 42…) з вашими «так» / «ні»; перемикати: галочки
 * в «🚀 Підключення», крок 5. «Старт»: лише кроки з номерами панелі. Модулі не змінюються.
 */
function stepStart12_(ss) {
  const out = [], was = modRead_(ss);
  if (!modBlock_(ss) || modBlock_(ss).old) { modBlockBuild_(ss, was); out.push('перемикачі модулів: ' + MOD_WHERE); }
  const st = ss.getSheetByName(START_OLD);
  if (st && modOldOnStart_(ss)) {
    const r = st.getRange(MOD_OLD_HEAD, 1, Math.min(MODS.length + 4, st.getMaxRows() - MOD_OLD_HEAD + 1), 4);
    r.clearContent(); r.clearDataValidations(); r.setFontWeight('normal').setFontSize(10).setFontColor('#000000').setHorizontalAlignment('left');
    out.push('«' + START_OLD + '»: без блоку модулів');
  }
  if (PropertiesService.getDocumentProperties().getProperty('prot') !== 'off' && modOn_('backup')) { try { applyProtection_(ss); } catch (e) { /* захист: наступного разу */ } }
  try { modulesIfChanged_(ss); } catch (e) { /* застосує mbsTick */ }
  try { startRefresh_(ss); } catch (e) { /* чекліст оновиться завтра */ }
  return out.length ? ['Оновлено: ' + out.join('; ') + '. Увімкнені модулі не змінились.'] : ['Усе вже було.'];
}

/** Версія 13: нові стартові правила у вільні рядки «Правил» (ваші йдуть раніше й виграють), «отримувач» у «Де шукати». */
function stepRules13_(ss) {
  const add = rulesUpgrade13_(ss);
  return [add.length ? 'Додано стартових правил: ' + add.length + ' (' + add.filter((x, i, a) => a.indexOf(x) === i).join(', ') + ').' : 'Нових стартових правил не треба.',
    'Без правила: переказ між своїми рахунками, закриття плану й категорія за історією ставляться самі. «Де шукати: отримувач»: назва, IBAN чи ЄДРПОУ (меню «🔎 Сирі дані Mono»).'];
}

/** Версія 15: «Старт» → «⚡ Швидке підключення» (стан і інструкція для ШІ); першою вкладкою, на виду. Модулі й дані не змінюються. */
function stepQuick15_(ss) {
  const out = [], old = ss.getSheetByName(START_OLD);
  if (old && !ss.getSheetByName(START_SHEET)) { old.setName(START_SHEET); out.push('«' + START_OLD + '» тепер «' + START_SHEET + '»'); }
  else if (old) { old.setName(START_OLD + ' (старий)'); old.hideSheet(); out.push('старий «' + START_OLD + '» сховано як «' + START_OLD + ' (старий)»'); }
  const sh = ss.getSheetByName(START_SHEET);
  if (!sh) { buildStart_(ss); out.push('аркуш «' + START_SHEET + '»'); }
  else {
    const r = sh.getRange(1, 1, Math.min(sh.getMaxRows(), START_ROWS + 5), Math.min(sh.getMaxColumns(), 4));
    r.clearContent(); r.clearDataValidations();
    startRefresh_(ss);
  }
  sheetsReorder_(ss);
  sheetsVisibility_(ss);
  if (PropertiesService.getDocumentProperties().getProperty('prot') !== 'off' && modOn_('backup')) { try { applyProtection_(ss); } catch (e) { /* захист: наступного разу */ } }
  return ['Оновлено: ' + out.join('; ') + '.', 'Відкрийте Gemini в таблиці (чи інший ШІ) і напишіть: «Прочитай аркуш «' + START_SHEET + '» і налаштуй зі мною таблицю».'];
}

/** Колонка T журналу вже з поверненнями (перший і останній рядок). */
function refundFormulaOk_(ss) {
  const sh = ss.getSheetByName(JOURNAL), f = journalFormulas_().T, col = colNum_(JC.out);
  return sh.getRange(J_FIRST, col).getFormula() === rowF_(f, J_FIRST) && sh.getRange(J_LAST, col).getFormula() === rowF_(f, J_LAST);
}

/** Версія 16: формула «Списання» (T) рахує мінус у витратах як повернення; підказка про суму. Ваші рядки й інші колонки не змінюються. */
function stepRefund16_(ss) {
  const sh = ss.getSheetByName(JOURNAL);
  if (!sh) return ['Журналу немає: нічого оновлювати.'];
  const f = journalFormulas_().T, rows = [];
  for (let r = J_FIRST; r <= J_LAST; r++) rows.push([rowF_(f, r)]);
  setF_(sh.getRange(J_FIRST, colNum_(JC.out), rows.length, 1), rows);
  if (headerIsNew_(ss)) sh.getRange('A3').setNote(headerNote_());
  else if (/завжди додатна/.test(String(sh.getRange(2, 1).getValue()))) sh.getRange(2, 1).setValue(String(sh.getRange(2, 1).getValue()).replace('сума (завжди додатна)', J_SUM_HINT));
  try { startRefresh_(ss); } catch (e) { /* «Швидке підключення» оновиться завтра */ }
  return ['Повернення за покупку: пишіть суму з мінусом у категорії покупки (наприклад, «Одяг −1 200»). Витрати категорії зменшаться, гроші повернуться в «мої».',
    'Імпорт з банку робить це сам для «Скасування. …» від Mono: бере категорію покупки з журналу.'];
}

/** Версія 21: категорії й правила для маркетплейсів, потім правила для старих чернеток без категорії. */
/** Версія 22: схована копія журналу і розклад onChange (видалені рядки назад). Дані журналу не змінюються. */
function stepProtect22_(ss) {
  const out = [];
  if (!ss.getSheetByName(SHADOW_SHEET)) { buildShadow_(ss); out.push('копія журналу «' + SHADOW_SHEET + '» (схована)'); } else shadowSave_(ss);
  if (!hasTrigger_(TRG.change)) { ensureChangeTrigger_(ss); out.push('розклад «видалені рядки назад»'); }
  if (PropertiesService.getDocumentProperties().getProperty('prot') === 'on') applyProtection_(ss);
  sheetShow_(ss, SHADOW_SHEET, false);
  return ['Оновлено: ' + (out.join('; ') || 'нічого') + '. Стерте записане в журналі повертається само; прибрати рядки: меню «' + WIPE_T + '…».'];
}

function stepShops21_(ss) {
  const out = dictAddCats_(ss, [[CAT_PROM, T_SPEND, '🛍 Покупки'], [CAT_MONO, T_SPEND, '🛍 Покупки'], [CAT_NP, T_SPEND, '🛍 Покупки'], [CAT_CASH, T_SPEND, '📦 Інше']]);
  rulesUpgrade21_(ss).forEach(c => out.push('правило «' + c + '»'));
  const n = ss.getSheetByName(JOURNAL) ? draftsRecat_(ss) : 0;
  if (n) out.push('чернеток отримали категорію: ' + n);
  PropertiesService.getDocumentProperties().setProperty('shops_v21', '1');
  return ['Оновлено: ' + (out.join('; ') || 'нічого') + '.'];
}

/** Версія 20: порядок рядків «Регулярних» і кольори. Значення рядків не змінюються, лише місце. */
function stepRegOrder20_(ss) {
  const moved = regOrder_(ss);
  regColors_(ss, true);
  if (ss.getSheetByName(JOURNAL)) regPlansSync_(ss);
  return ['«' + REG_SHEET + '»: ' + (moved ? 'рядки впорядковано (угорі «Заплати собі» й «Інвестиції», далі ваші, приклади внизу)' : 'порядок уже такий') + '; кольори груп, вимкнені сірі.'];
}

/** «Вільно» з «моїх» і «Мін.» від сьогодні вже є? */
function freeNew_(ss) {
  const days = ss.getSheetByName(DAYS_SHEET), d = ss.getSheetByName(DICT_SHEET);
  return (!ss.getSheetByName(MAIN_SHEET) || mainHasLabel_(ss, MAIN_NEXT_FREE)) && (!days || String(days.getRange('B2').getFormula()).indexOf('=MINIFS(') === 0) &&
    (!d || String(d.getRange('V' + D_FIRST).getFormula()) !== '');
}

/** Версія 19: «Головне» (рядки наступного місяця), «Довідник» V, «Дні» B2:C4, шапка журналу. Ваші дані не змінюються. */
function stepFree19_(ss) {
  dictNextHelper_(ss);
  const main = ss.getSheetByName(MAIN_SHEET);
  if (main) {                                                      // рядки зсунулись: сховані за модулями ховаємо знову за підписами
    buildMain_(ss);
    try { main.showRows(1, Math.min(main.getMaxRows(), 120)); } catch (e) { /* ignore */ }
    mainRowsShow_(ss, MAIN_CREDIT, modOn_('credit')); mainRowsShow_(ss, MAIN_DEBTS, modOn_('debts'));
  }
  const days = ss.getSheetByName(DAYS_SHEET);
  if (days) daysTop_(days);
  if (headerIsNew_(ss)) journalHeader_(ss);
  return ['«Вільно» в шапці журналу тепер з «моїх» наступного місяця: дохід, що йде в місяць, мінус «на погашення ліміту», плани й бюджети категорій.',
    '«Мін.» і «як зараз»: на картці, лише від сьогодні.'];
}

/** Версія 18: категорії «Заплати собі» й «Інвестиції», два нові приклади на «Регулярних», шапка-плитки. Ваші рядки не змінюються. */
function stepReg18_(ss) {
  const out = dictSelfInvest_(ss);
  const n = regExamplesMore_(ss);
  if (n) out.push('прикладів на «' + REG_SHEET + '»: ' + n + ' (з «ні»)');
  if (regHeader_(ss)) out.push('шапка «' + REG_SHEET + '»: плитки в стилі «Charts»');
  return ['Оновлено: ' + (out.join('; ') || 'нічого') + '.', 'Рядки 1–3 «' + REG_SHEET + '» тепер шапка; пояснення в примітці на заголовку «Назва».'];
}

/** «Борги» вже з колонками L…O і блоком лімітів? */
function debtsNew_(ss) {
  const sh = ss.getSheetByName(DEBTS_SHEET);
  return !!sh && sh.getMaxColumns() >= DEBT_COLS.length && String(sh.getRange(DEBT_HEAD, 12).getValue()) === DEBT_COLS[11] &&
    String(sh.getRange(DEBT_LIM_HEAD, 1).getValue()) === DEBT_LIM_COLS[0];
}

/** Версія 17: на «Боргах» колонки L…O (вид, назва в банку, закрито, прогрес) і ліміти (рядки 27–29); шапка журналу K; «Борги» на виду з модулем. Ваші рядки не змінюються. */
function stepParts17_(ss) {
  const sh = ss.getSheetByName(DEBTS_SHEET), out = [];
  if (!sh) return ['«' + DEBTS_SHEET + '» з\'являться з модулем «🤝 Борги й розстрочки».'];
  ensureSize_(sh, DEBT_LIMITS[DEBT_LIMITS.length - 1].row + 2, DEBT_COLS.length);
  const busy = sh.getRange(DEBT_LIM_HEAD, 1, DEBT_LIMITS[DEBT_LIMITS.length - 1].row - DEBT_LIM_HEAD + 1, DEBT_LIM_COLS.length).getValues()
    .some(r => r.some(v => v !== '' && v !== null)) && String(sh.getRange(DEBT_LIM_HEAD, 1).getValue()) !== DEBT_LIM_COLS[0];
  debtsColumns_(ss, sh);
  out.push('«' + DEBTS_SHEET + '»: колонки «Вид», «Назва в банку», «Закрито», «Прогрес»');
  if (busy) out.push('⚠️ рядки ' + DEBT_LIM_HEAD + '–' + DEBT_LIMITS[DEBT_LIMITS.length - 1].row + ' на «' + DEBTS_SHEET + '» зайняті вашими даними: блок лімітів не додано (звільніть і запустіть «⬆️ Оновити зараз»)');
  else { debtsLimits_(sh); out.push('ліміти Mono (рядки ' + DEBT_LIMITS[0].row + '–' + DEBT_LIMITS[DEBT_LIMITS.length - 1].row + ')'); }
  sheetsVisibility_(ss);
  if (headerIsNew_(ss)) headerParts_(ss);
  return ['Оновлено: ' + out.join('; ') + '.', 'Покупки частинами з Mono тепер додаються на «' + DEBTS_SHEET + '» самі з першого платежу: впишіть лише суму покупки (C).'];
}

/** Блок модулів має рядок для кожного модуля (нові модулі дописуються в кінець). */
function modBlockFull_(ss) {
  const b = modBlock_(ss);
  if (!b) return false;
  const v = b.sh.getRange(b.first, 1, MODS.length, 1).getValues();
  return MODS.every((m, i) => String(v[i][0]) === m.name);
}

/** Дописати порожні рядки нових модулів, не чіпаючи вже вибраних «так» / «ні». Інша назва в рядку: блок заново. */
function modBlockAdd_(ss) {
  const b = modBlock_(ss);
  if (!b) { modBlockBuild_(ss); return; }
  const st = ensureSize_(b.sh, b.first + MODS.length + 5, 4), F = b.first;
  const v = st.getRange(F, 1, MODS.length, 1).getValues(), init = modInitial_(ss);
  if (MODS.some((m, i) => String(v[i][0]) !== '' && String(v[i][0]) !== m.name)) { modBlockBuild_(ss, modRead_(ss)); return; }
  MODS.forEach((m, i) => {
    if (String(v[i][0]) !== '') return;
    st.getRange(F + i, 1, 1, 3).setValues([[m.name, init[m.id] ? MOD_YES : MOD_NO, m.what]]);
    st.getRange(F + i, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([MOD_YES, MOD_NO]).setAllowInvalid(false).build())
      .setHorizontalAlignment('center').setFontWeight('bold');
    st.getRange(F + i, 3).setFontColor('#5f6368');
  });
}

/** Наші вкладки в порядку SHEET_ORDER; чужі лишаються після них. Нічого не видаляє. */
function sheetsReorder_(ss) {
  if (!ss.setActiveSheet || !ss.moveActiveSheet) return;
  const was = ss.getActiveSheet ? ss.getActiveSheet() : null;
  SHEET_ORDER.filter(n => !!ss.getSheetByName(n)).forEach((n, i) => { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
  try { if (was) ss.setActiveSheet(was); } catch (e) { /* ignore */ }
}

/** Видимість вкладок (і після перестановки: активація могла показати схований аркуш): VISIBLE_SHEETS завжди; решта за модулем «🗂 Усі аркуші» (all: його нове значення); «ФОП» і «Борги» ще й за своїми модулями. */
function sheetsVisibility_(ss, all) {
  const on = all === undefined ? modOn_('service') : all;
  VISIBLE_SHEETS.forEach(n => sheetShow_(ss, n, true));
  SERVICE_SHEETS.forEach(n => sheetShow_(ss, n, on));
  sheetShow_(ss, FOP_SHEET, on && modOn_('fop'));
  sheetShow_(ss, DEBTS_SHEET, modOn_('debts'));                        // покупки частинами й ліміти: на виду, поки модуль увімкнено
  sheetShow_(ss, CH_SHEET, false);
  sheetShow_(ss, SHADOW_SHEET, false);                                  // копія журналу для захисту: завжди схована
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
  sheetsVisibility_(ss);
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
    sec('🧩 Модулі', modLines_().concat(['Змінити: ' + MOD_WHERE + '.'])), sec('🤖 Радник Gemini', advisorLines_(ss))]);
}
