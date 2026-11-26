/** 42_Modules.gs: Модулі-перемикачі на «Старті»: кожен вмикається й вимикається окремо («так» / «ні»). */

/*
 * Вимкнений модуль ховає свої аркуші, рядки «Головного» й повідомлення та знімає свої розклади. Дані не видаляються:
 * увімкнули знову, і все повертається. Усе вимкнено: простий журнал з «Місяцями», «Днями» й графіками.
 * «🗂 Службові аркуші» (типово «ні»): на виду лише робочі вкладки; правила банку, рахунки, залишки, історія й перевірка
 * працюють схованими.
 * Застосовується одразу після правки на «Старті» (розклад mbsEdit), а ще кожні 15 хв (якщо правку пропущено),
 * з меню «🧩 Модулі: застосувати» і з майстра «🚀 Почати».
 * Клітинки на «Старті»: джерело правди; властивість документа mods: їх копія, щоб розклади читали швидко.
 */

const MOD_HEAD = 24, MOD_FIRST = 26;
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
  { id: 'service', name: '🗂 Службові аркуші', what: 'показати «' + RULES_SHEET + '», «' + ACC_SHEET + '», «' + CP_SHEET + '», «' + BANK_SHEET +
    '», історію й «' + CHECK_SHEET + '» (схованими вони працюють так само)', def: false },
];
/** Аркуші, які ховає вимкнений модуль «🗂 Службові аркуші». «Дані графіків» схований завжди. */
const SERVICE_SHEETS = [RULES_SHEET, ACC_SHEET, CP_SHEET, BANK_SHEET, HISTORY_SHEET, H_CAT_SHEET, CHECK_SHEET];

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

/** Блок «🧩 Модулі» на «Старті». Значення: з mods або початкові. */
function startModulesBlock_(ss, sh) {
  const st = modInitial_(ss);
  sh.getRange(MOD_HEAD, 1, 1, 3).setValues([['🧩 Модулі', '', 'так / ні: вимкнений модуль ховає свої аркуші, рядки й повідомлення; дані не видаляються.']]);
  sh.getRange(MOD_HEAD, 1).setFontWeight('bold').setFontSize(12);
  sh.getRange(MOD_HEAD, 3).setFontColor('#5f6368');
  sh.getRange(MOD_HEAD + 1, 1, 1, 3).setValues([['Модуль', 'Увімкнено', 'Що робить']]).setFontWeight('bold');
  sh.getRange(MOD_FIRST, 1, MODS.length, 3).setValues(MODS.map(m => [m.name, st[m.id] ? MOD_YES : MOD_NO, m.what]));
  sh.getRange(MOD_FIRST, 2, MODS.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([MOD_YES, MOD_NO]).setAllowInvalid(false).build())
    .setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange(MOD_FIRST, 3, MODS.length, 1).setFontColor('#5f6368');
}

function startHasModules_(ss) {
  const sh = ss.getSheetByName(START_SHEET);
  return !!sh && sh.getMaxRows() >= MOD_FIRST + MODS.length - 1 && String(sh.getRange(MOD_HEAD, 1).getValue()) === '🧩 Модулі';
}

/** Що стоїть у клітинках «Старту»: { id: true/false } або null (блоку немає). */
function modRead_(ss) {
  if (!startHasModules_(ss)) return null;
  const v = ss.getSheetByName(START_SHEET).getRange(MOD_FIRST, 1, MODS.length, 2).getValues(), o = {};
  MODS.forEach((m, i) => { const x = String(v[i][1]).trim().toLowerCase(); o[m.id] = x === '' ? m.def : x !== MOD_NO; });
  return o;
}

/** Змінити один модуль (клітинка + властивість) без застосування: для пунктів меню, що вже роблять своє (календар, Telegram). */
function modSet_(ss, id, on) {
  const i = MODS.findIndex(m => m.id === id);
  if (i < 0) return;
  if (startHasModules_(ss)) ss.getSheetByName(START_SHEET).getRange(MOD_FIRST + i, 2).setValue(on ? MOD_YES : MOD_NO);
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
    if (!tgConnected_()) return 'увімкнено; бот ще не підключено: «⚙️ Налаштування → 📨 Підключити Telegram-бота»';
    if (!hasTrigger_(TRG.morning)) ensureMorningTrigger_();
    return 'увімкнено ✓';
  },
  credit: (ss, on) => {
    mainRowsShow_(ss, MAIN_CREDIT, on);
    const d = ss.getSheetByName(DAYS_SHEET), m = ss.getSheetByName(MONTHS_SHEET);
    if (d) { if (on) d.showColumns(6, 3); else d.hideColumns(6, 3); }
    if (m) monthsCreditRows_().forEach(g => { if (on) m.showRows(g[0], g[1]); else m.hideRows(g[0], g[1]); });
    if (headerIsNew_(ss)) journalHeader_(ss);                       // шапка: грейс або «на день»
    chartsRefresh_(ss);                                             // графіки: з лімітом, боргом і причинами чи без
    return on ? 'увімкнено ✓' + (cardIsCredit_(ss) ? '' : ' (на «' + SETTINGS_SHEET + '» картка дебетова: ліміт буде 0)') : 'вимкнено: рядки ліміту й грейсу сховано';
  },
  debts: (ss, on) => {
    if (on && !ss.getSheetByName(DEBTS_SHEET)) buildDebts_(ss);
    sheetShow_(ss, DEBTS_SHEET, on);
    mainRowsShow_(ss, MAIN_DEBTS, on);
    const st = debtsPlansSync_(ss);
    return (on ? 'увімкнено ✓: аркуш «' + DEBTS_SHEET + '»' : 'вимкнено: аркуш сховано') + (st.added + st.removed ? ', планів у журналі +' + st.added + ' −' + st.removed : '');
  },
  budgets: (ss, on) => {
    if (headerIsNew_(ss)) journalHeader_(ss);                       // шапка: три бюджети або витрати місяця
    if (ss.getSheetByName(PHONE_SHEET)) buildPhone_(ss);            // «📱 Телефон»: три бюджети
    return on ? 'увімкнено ✓' : 'вимкнено: тривог про бюджети не буде (плани категорій у «' + DICT_SHEET + '» лишаються для «' + MONTHS_SHEET + '»)';
  },
  cal: (ss, on, was) => {
    if (!on) return 'вимкнено: календар більше не оновлюється (сам календар лишився; видалити: «⚙️ Налаштування → 📆 Календар платежів: вимкнути»)';
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
    sheetShow_(ss, FOP_SHEET, on);
    const st = fopPlansSync_(ss);
    return (on ? 'увімкнено ✓: аркуш «' + FOP_SHEET + '»' : 'вимкнено: аркуш сховано') + (st.added + st.removed ? ', планів податків +' + st.added + ' −' + st.removed : '');
  },
  service: (ss, on) => {
    SERVICE_SHEETS.forEach(n => sheetShow_(ss, n, on));
    sheetShow_(ss, CH_SHEET, false);
    try { startRefresh_(ss); } catch (e) { /* чекліст: підказка «сховано» оновиться завтра */ }
    return on ? 'увімкнено ✓: службові аркуші видно' : 'вимкнено: службові аркуші сховано (працюють як і раніше; показати: «так»)';
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
  if (!want) return ['Блоку «🧩 Модулі» на «' + START_SHEET + '» ще немає: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».'];
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

/** Чи правка зачепила клітинки модулів на «Старті». */
function isModulesEdit_(range) {
  return range.getSheet().getName() === START_SHEET && range.getColumn() <= 2 && range.getLastColumn() >= 2 &&
    range.getLastRow() >= MOD_FIRST && range.getRow() <= MOD_FIRST + MODS.length - 1;
}

function modLines_() {
  return MODS.map(m => m.name + ': ' + (modOn_(m.id) ? '<b>так</b>' : 'ні'));
}

/* ---------- меню ---------- */

function modulesNow() {
  if (!menuOk_('🧩 Модулі')) return;
  const ss = SpreadsheetApp.getActive();
  if (!startHasModules_(ss)) { try { schemaRun_(ss); } catch (e) { show_('🧩 Модулі', ['⚠️ ' + esc_(e.message)]); return; } }
  const lines = modulesApply_(ss, true);
  show_('🧩 Модулі', lines.concat(['Змінити: аркуш «' + START_SHEET + '», блок «🧩 Модулі» (так / ні); застосовується саме.']));
}
