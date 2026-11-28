/** 42_Modules.gs: Модулі-перемикачі: кожен вмикається й вимикається окремо. Перемикати: «🚀 Підключення», крок 5 (меню «🧩 Модулі»). */

/*
 * Вимкнений модуль ховає свої аркуші, рядки «Головного» й повідомлення та знімає свої розклади. Дані не видаляються:
 * увімкнули знову, і все повертається. Усе вимкнено: простий журнал з «Місяцями», «Днями» й графіками.
 * «🗂 Усі аркуші» (типово «ні»): на виду лише «⚡ Швидке підключення», «Журнал», «Регулярні», «Charts» і «charts_mob» (VISIBLE_SHEETS);
 * «Місяці», «Дні», «Довідник», «Налаштування», «ФОП», «Борги» і службові працюють схованими. «так» показує їх
 * («ФОП» і «Борги» лише з увімкненими модулями).
 * Перемикачі: галочки в панелі «🚀 Підключення», крок 5. Вони пишуть «так» / «ні» в блок «🧩 Модулі» на «Налаштуваннях»
 * (рядки 42…; до версії 12 блок був на «Старті») і одразу застосовують. Правка блоку вручну теж працює (розклад mbsEdit,
 * а ще кожні 15 хв, якщо правку пропущено). Клітинки блоку: джерело правди (переносяться «📦 Перенести дані»);
 * властивість документа mods: їх копія, щоб розклади читали швидко.
 */

const MOD_TITLE = '🧩 Модулі';
const MOD_HEAD = 42, MOD_FIRST = 44;                                     // «Налаштування» (з версії 12)
const MOD_OLD_HEAD = 24, MOD_OLD_FIRST = 26;                             // «Старт» (до версії 12)
const MOD_WHERE = '«🚀 Підключення», крок 5 (меню «🧩 Модулі»)';           // де перемикати: для підказок

const MOD_YES = 'так', MOD_NO = 'ні';
const MODS = [
  { id: 'tg', name: '📨 Telegram', what: 'ранковий підсумок, тижневий і місячний звіти, тривоги', def: true },
  { id: 'credit', name: '💳 Кредитний ліміт і грейс', what: 'борг за лімітом, «мої» проти ліміту, причини, грейс Mono і комісія за перекриття', def: true },
  { id: 'debts', name: '🤝 Борги й розстрочки', what: 'аркуш «' + DEBTS_SHEET + '»: щомісячні платежі стають планами в журналі', def: false },
  { id: 'budgets', name: '💰 Бюджети категорій', what: 'план на місяць («' + DICT_SHEET + '», F): «перевищено» і «швидше норми» в Telegram', def: true },
  { id: 'cal', name: '📆 Календар платежів', what: 'плани з журналу в окремому Google Календарі з нагадуваннями', def: false },
  { id: 'backup', name: '💾 Копії й захист', what: 'нічна копія на Диску; захист аркушів з формулами «з попередженням»', def: true },
  { id: 'fop', name: '🧾 ФОП', what: 'аркуш «' + FOP_SHEET + '»: податки, строки, ліміт доходу групи', def: false },
  { id: 'advisor', name: '🤖 Радник Gemini', what: 'щоденний знімок фінансів на вашому Диску для навички Gemini', def: false },
  { id: 'service', name: '🗂 Усі аркуші', what: 'показати «' + MONTHS_SHEET + '», «' + DAYS_SHEET + '», «' + DICT_SHEET + '», «' + SETTINGS_SHEET + '», «' + FOP_SHEET +
    '», «' + DEBTS_SHEET + '» і службові (схованими вони працюють так само)', def: false },
];
/** Завжди на виду. */
const VISIBLE_SHEETS = [START_SHEET, JOURNAL, REG_SHEET, DASH_SHEET, MOB_SHEET];
/** Аркуші, які ховає вимкнений модуль «🗂 Усі аркуші» («ФОП» і «Борги» ще й за своїми модулями). «Дані графіків» схований завжди. */
const SERVICE_SHEETS = [MONTHS_SHEET, DAYS_SHEET, DICT_SHEET, SETTINGS_SHEET, MAIN_SHEET, RULES_SHEET, ACC_SHEET, CP_SHEET, BANK_SHEET, HISTORY_SHEET, H_CAT_SHEET, CHECK_SHEET];

let MODS_CACHE_ = null;
function modState_() {
  if (!MODS_CACHE_) MODS_CACHE_ = jsonProp_(PropertiesService.getDocumentProperties(), 'mods') || {};
  return MODS_CACHE_;
}
function modDef_(id) { const m = MODS.find(x => x.id === id); return m ? m.def : true; }
/** Модуль увімкнено? Ще не налаштовано: типове значення. */
function modOn_(id) { const st = modState_(); return Object.prototype.hasOwnProperty.call(st, id) ? !!st[id] : modDef_(id); }
function modName_(id) { const m = MODS.find(x => x.id === id); return m ? m.name : id; }

/** Початковий стан для таблиці без блоку модулів: що вже працює, те й увімкнено. */
function modInitial_(ss) {
  const dp = PropertiesService.getDocumentProperties(), saved = jsonProp_(dp, 'mods');
  const o = {};
  MODS.forEach(m => { o[m.id] = saved && Object.prototype.hasOwnProperty.call(saved, m.id) ? !!saved[m.id] : m.def; });
  if (!saved) {
    o.credit = cardIsCredit_(ss);
    o.cal = dp.getProperty('cal') === 'on';
    o.fop = ss.getSheetByName(FOP_SHEET) ? fopList_(ss).length > 0 : false;
    o.backup = dp.getProperty('prot') !== 'off' || hasTrigger_(TRG.backup);
  }
  return o;
}

/** Де блок «🧩 Модулі»: { sh, first, old } («Налаштування»; old: ще на «Старті», до версії 12) або null. */
function modBlock_(ss) {
  const at = (name, head, first) => {
    const sh = ss.getSheetByName(name);
    return sh && sh.getMaxRows() >= first + MODS.length - 1 && String(sh.getRange(head, 1).getValue()) === MOD_TITLE ? { sh: sh, first: first, old: name === START_OLD } : null;
  };
  return at(SETTINGS_SHEET, MOD_HEAD, MOD_FIRST) || at(START_OLD, MOD_OLD_HEAD, MOD_OLD_FIRST);
}
function startHasModules_(ss) { return !!modBlock_(ss); }
/** Старий блок ще стоїть на «Старті». */
function modOldOnStart_(ss) {
  const sh = ss.getSheetByName(START_OLD);
  return !!sh && sh.getMaxRows() >= MOD_OLD_HEAD && String(sh.getRange(MOD_OLD_HEAD, 1).getValue()) === MOD_TITLE;
}

/** Блок «🧩 Модулі» на «Налаштуваннях». Значення: want ({ id: true/false }), інакше з mods або початкові. */
function modBlockBuild_(ss, want) {
  const sh = ensureSize_(ss.getSheetByName(SETTINGS_SHEET), MOD_FIRST + MODS.length + 5, 4), st = want || modInitial_(ss);
  sh.getRange(MOD_HEAD, 1, 1, 3).setValues([[MOD_TITLE, '', 'так / ні. Зручніше галочками: «🚀 Підключення», крок 5. Вимкнений модуль ховає свої аркуші, рядки й повідомлення; дані не видаляються.']]);
  sh.getRange(MOD_HEAD, 1).setFontWeight('bold').setFontSize(12);
  sh.getRange(MOD_HEAD, 3).setFontColor('#5f6368');
  sh.getRange(MOD_HEAD + 1, 1, 1, 3).setValues([['Модуль', 'Увімкнено', 'Що робить']]).setFontWeight('bold');
  sh.getRange(MOD_FIRST, 1, MODS.length, 3).setValues(MODS.map(m => [m.name, (Object.prototype.hasOwnProperty.call(st, m.id) ? st[m.id] : m.def) ? MOD_YES : MOD_NO, m.what]));
  sh.getRange(MOD_FIRST, 2, MODS.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([MOD_YES, MOD_NO]).setAllowInvalid(false).build())
    .setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange(MOD_FIRST, 3, MODS.length, 1).setFontColor('#5f6368');
}

/** Що стоїть у клітинках блоку: { id: true/false } або null (блоку немає). */
function modRead_(ss) {
  const b = modBlock_(ss);
  if (!b) return null;
  const v = b.sh.getRange(b.first, 1, MODS.length, 2).getValues(), o = {};
  MODS.forEach((m, i) => { const x = String(v[i][1]).trim().toLowerCase(); o[m.id] = x === '' ? m.def : x !== MOD_NO; });
  return o;
}

/** Змінити один модуль (клітинка + властивість) без застосування: для пунктів меню, що вже роблять своє (календар, Telegram). */
function modSet_(ss, id, on) {
  const i = MODS.findIndex(m => m.id === id);
  if (i < 0) return;
  const b = modBlock_(ss);
  if (b) b.sh.getRange(b.first + i, 2).setValue(on ? MOD_YES : MOD_NO);
  const dp = PropertiesService.getDocumentProperties(), st = jsonProp_(dp, 'mods') || modInitial_(ss);
  st[id] = !!on;
  dp.setProperty('mods', JSON.stringify(st));
  MODS_CACHE_ = st;
}

/* ---------- застосування ---------- */

function sheetShow_(ss, name, on) {
  const sh = ss.getSheetByName(name);
  if (!sh) return;
  try { if (on) sh.showSheet(); else if (ss.getSheets().length > 1) sh.hideSheet(); } catch (e) { /* аркуш активний або останній видимий */ }
}

/** Рядки «Головного» за підписами: показати або сховати. */
function mainRowsShow_(ss, labels, on) {
  const sh = ss.getSheetByName(MAIN_SHEET);
  if (!sh) return;
  const v = sh.getRange(1, 1, Math.min(Math.max(sh.getLastRow(), 1), 120), 1).getValues();
  v.forEach((r, i) => { if (labels.indexOf(String(r[0])) >= 0) { if (on) sh.showRows(i + 1, 1); else sh.hideRows(i + 1, 1); } });
}

/** Рядки «Місяців» про ліміт: [[перший рядок, скільки]]. Блок «ЧОМУ З ЛІМІТУ» цілком. */
function monthsCreditRows_() {
  const one = ['Витрачено з ліміту', 'Відкладено на погашення (бюджетний місяць)', 'Погашено за бюджетний місяць (відкладено + невитрачене)', 'Орієнтир погашення',
    'Борг за лімітом на кінець місяця'].map(l => [monthsSummaryRow_(l), 1]);
  const a = monthsSummaryRow_('ЧОМУ З ЛІМІТУ'), b = monthsSummaryRow_('Разом з ліміту');
  return one.concat(a && b ? [[a, b - a + 1]] : []).filter(g => g[0] > 0);
}

/** Що робить кожен модуль, коли його ввімкнули (on) чи вимкнули. was: попередній стан (undefined: ще не було). → рядок звіту. */
const MOD_APPLY = {
  tg: (ss, on) => {
    if (!on) { removeTriggers_(TRG.morning); removeTriggers_(TRG.morningRetry); return 'вимкнено: повідомлень не буде (бот і токен лишаються)'; }
    if (!tgConnected_()) return 'увімкнено; бот ще не підключено: «🚀 Підключення», крок 2';
    if (!hasTrigger_(TRG.morning)) ensureMorningTrigger_();
    return 'увімкнено ✓';
  },
  credit: (ss, on) => {
    mainRowsShow_(ss, MAIN_CREDIT, on);
    const d = ss.getSheetByName(DAYS_SHEET), m = ss.getSheetByName(MONTHS_SHEET);
    if (d) { if (on) d.showColumns(6, 3); else d.hideColumns(6, 3); }
    if (m) monthsCreditRows_().forEach(g => { if (on) m.showRows(g[0], g[1]); else m.hideRows(g[0], g[1]); });
    if (headerIsNew_(ss)) journalHeader_(ss);                       // шапка: грейс або «на день»
    dashRefresh_(ss);                                               // «Charts», «charts_mob»: з лімітом, боргом і причинами чи без
    return on ? 'увімкнено ✓' + (cardIsCredit_(ss) ? '' : ' (на «' + SETTINGS_SHEET + '» картка дебетова: ліміт буде 0)') : 'вимкнено: рядки ліміту й грейсу сховано';
  },
  debts: (ss, on) => {
    if (on && !ss.getSheetByName(DEBTS_SHEET)) buildDebts_(ss);
    sheetShow_(ss, DEBTS_SHEET, on);                                // на виду з модулем: покупки частинами, розстрочки, ліміти
    mainRowsShow_(ss, MAIN_DEBTS, on);
    if (headerIsNew_(ss)) headerParts_(ss);                         // шапка журналу: борги місяця й ліміт частинами
    const st = debtsPlansSync_(ss);
    return (on ? 'увімкнено ✓: аркуш «' + DEBTS_SHEET + '» на виду (покупки частинами, розстрочки, ліміти)' : 'вимкнено: аркуш сховано') + (st.added + st.removed ? ', планів у журналі +' + st.added + ' −' + st.removed : '');
  },
  budgets: (ss, on) => {
    if (headerIsNew_(ss)) journalHeader_(ss);                       // шапка: три бюджети або витрати місяця
    return on ? 'увімкнено ✓' : 'вимкнено: тривог про бюджети не буде (плани категорій у «' + DICT_SHEET + '» лишаються для «' + MONTHS_SHEET + '»)';
  },
  cal: (ss, on, was) => {
    if (!on) return 'вимкнено: календар більше не оновлюється (сам календар лишився; видалити: «🚀 Підключення → Інше → Календар: видалити»)';
    if (was === true) return 'увімкнено ✓';
    try { const st = calSync_(ss); return 'увімкнено ✓: ' + calStatText_(st); } catch (e) { return '⚠️ увімкнено, але оновити не вдалося: ' + esc_(e && e.message || e); }
  },
  backup: (ss, on, was) => {
    if (!on) { removeTriggers_(TRG.backup); const n = protectionRemove_(ss); return 'вимкнено: нічних копій не буде' + (n ? ', захист знято' : ''); }
    if (!hasTrigger_(TRG.backup)) ensureBackupTrigger_(ss);
    if (was === false) applyProtection_(ss);
    return 'увімкнено ✓: нічні копії о ' + pad2_(BACKUP_HOUR) + ':00' + (protectionOn_(ss) ? ', захист з попередженням' : '');
  },
  fop: (ss, on) => {
    if (on && !ss.getSheetByName(FOP_SHEET)) buildFop_(ss);
    sheetShow_(ss, FOP_SHEET, on && modOn_('service'));
    const st = fopPlansSync_(ss);
    return (on ? 'увімкнено ✓: аркуш «' + FOP_SHEET + '»' + (modOn_('service') ? '' : ' (схований, як і решта: «' + modName_('service') + '»: так)') : 'вимкнено: аркуш сховано') + (st.added + st.removed ? ', планів податків +' + st.added + ' −' + st.removed : '');
  },
  service: (ss, on) => {
    sheetsVisibility_(ss, on);
    try { startRefresh_(ss); } catch (e) { /* чекліст: підказка «сховано» оновиться завтра */ }
    return on ? 'увімкнено ✓: усі аркуші видно' : 'вимкнено: на виду лише «' + VISIBLE_SHEETS.join('», «') + '» (решта працює як і раніше; показати: «так»)';
  },
  advisor: (ss, on, was) => {
    if (!on) return 'вимкнено: знімок більше не оновлюється (документ на Диску лишився)';
    if (was === true) return 'увімкнено ✓';
    try { const r = snapshotWrite_(ss); return 'увімкнено ✓: знімок «' + esc_(ADV_DOC) + '» (' + r.chars + ' знаків)'; }
    catch (e) { return '⚠️ увімкнено, але знімок не записано: ' + esc_(e && e.message || e); }
  },
};

/**
 * Перечитати клітинки «Старту» й застосувати. force: усі модулі заново (не лише змінені). → рядки звіту.
 * Під час інших дій скрипта (замок зайнятий) нічого не робить: спробує наступний розклад.
 */
function modulesApply_(ss, force) {
  const want = modRead_(ss);
  if (!want) return ['Блоку «' + MOD_TITLE + '» ще немає: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».'];
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20 * 1000)) return ['⚠️ Скрипт зайнятий іншою дією; модулі застосуються самі за 15 хв.'];
  const dp = PropertiesService.getDocumentProperties(), was = jsonProp_(dp, 'mods') || {}, out = [];
  try {
    dp.setProperty('mods', JSON.stringify(want));
    MODS_CACHE_ = want;
    MODS.forEach(m => {
      const before = Object.prototype.hasOwnProperty.call(was, m.id) ? !!was[m.id] : undefined;
      if (!force && before === want[m.id]) return;
      let line;
      try { line = MOD_APPLY[m.id](ss, want[m.id], before); } catch (e) { line = '⚠️ ' + esc_(e && e.message || e); }
      out.push(m.name + ': ' + line);
    });
  } finally { lock.releaseLock(); }
  return out.length ? out : ['Змін немає: модулі вже застосовано.'];
}

/** З mbsTick і mbsEdit: клітинки змінились, а властивість ще ні → застосувати. */
function modulesIfChanged_(ss) {
  const want = modRead_(ss);
  if (!want) return null;
  const was = jsonProp_(PropertiesService.getDocumentProperties(), 'mods') || {};
  if (MODS.every(m => Object.prototype.hasOwnProperty.call(was, m.id) && !!was[m.id] === want[m.id])) return null;
  return modulesApply_(ss, false);
}

/** Чи правка зачепила клітинки модулів («Налаштування»; у старих таблицях ще «Старт»). */
function isModulesEdit_(range) {
  const name = range.getSheet().getName(), first = name === SETTINGS_SHEET ? MOD_FIRST : name === START_OLD ? MOD_OLD_FIRST : 0;
  return first > 0 && range.getColumn() <= 2 && range.getLastColumn() >= 2 && range.getLastRow() >= first && range.getRow() <= first + MODS.length - 1;
}

function modLines_() {
  return MODS.map(m => m.name + ': ' + (modOn_(m.id) ? '<b>так</b>' : 'ні'));
}

/* ---------- меню ---------- */

/** Меню «🧩 Модулі»: панель «🚀 Підключення» одразу на кроці 5 (галочки). */
function modulesNow() { cnOpen_(5); }
