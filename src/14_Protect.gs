/** 14_Protect.gs: журнал без стирання (записане повертається) і «🧹 Зачистити дані»: єдиний спосіб прибрати рядки. */

/*
 * Додавати й змінювати в журналі можна. Стирати й видаляти записане: ні.
 *  - Копія: схований аркуш «🛡 Захист журналу», рядок у рядок з журналом (A…H і службові Y…AF одним текстом).
 *    Оновлюється, коли журнал упорядковується (імпорт, ≈ хвилина після правки, «🧹 Зачистити дані») і кожні 15 хв.
 *  - Стерли дату, суму, категорію чи статус записаного рядка або весь рядок: onEdit одразу повертає з копії.
 *    Коментар, «Непомітна», «Хто», «Причина» стирати можна. Щойно вписане (ще не в копії) теж. Плани не захищаються:
 *    їх ставлять «Регулярні», «Борги», ФОП. Стерли формули праворуч (після попередження Google): пишуться знову.
 *  - Видалили рядки журналу: розклад mbsChange (onChange) вставляє їх назад з копії на те саме місце, з формулами.
 *    Видалили рядок шапки чи копія не збігається: нічого не вгадуємо, підказка «Ctrl+Z».
 *  - Прибрати рядки: «🧹 Зачистити дані…» (усе, дата, період чи місяць), з копією таблиці перед змінами. Записи цих днів
 *    на «Банк» теж ідуть, тож імпорт з банку завантажить ці операції знову.
 */

const SHADOW_SHEET = '🛡 Захист журналу', KEEP_COLS = [0, 1, 2, 4];       // A дата, B сума, C категорія, E статус
const PROTECT_T = '🛡 Журнал', WIPE_T = '🧹 Зачистити дані';

const blank_ = x => x === '' || x === null || x === undefined;
const isDate_ = x => Object.prototype.toString.call(x) === '[object Date]';

/** Чиста функція: рядок (A…H, Y…AF) → текст для копії; порожній → ''. */
function shadowEnc_(v) {
  if (!v || v.every(blank_)) return '';
  return JSON.stringify(v.map(x => (isDate_(x) ? { d: x.getTime() } : blank_(x) ? '' : x)));
}

/** Чиста функція: текст копії → рядок або null. */
function shadowDec_(s) {
  if (!s) return null;
  try { return JSON.parse(String(s)).map(x => (x && typeof x === 'object' && typeof x.d === 'number' ? new Date(x.d) : x)); } catch (e) { return null; }
}

function buildShadow_(ss) {
  const sw = ensureSize_(sheetOf_(ss, SHADOW_SHEET), J_LAST, 2);
  sw.getRange('A1:A' + J_LAST).setNumberFormat('@');
  sw.getRange(1, 2).setValue('Пише скрипт: копія журналу, щоб записане не стиралось і видалені рядки повертались. Не змінюйте й не видаляйте цей аркуш.');
  shadowSave_(ss);
  try { sw.hideSheet(); } catch (e) { /* останній видимий аркуш не ховається */ }
  return sw;
}

/** Копія журналу. ah, bk: A…H і Y…AF усіх рядків, якщо вже прочитані (sortCore_). */
function shadowSave_(ss, ah, bk) {
  const sh = ss.getSheetByName(JOURNAL), sw = ss.getSheetByName(SHADOW_SHEET);
  if (!sh || !sw) return false;
  const n = J_LAST - J_FIRST + 1;
  if (!ah || !bk) { ah = sh.getRange(J_FIRST, 1, n, J_INPUT_N).getValues(); bk = sh.getRange(J_FIRST, colNum_(JC.bankId), n, J_MOVE_N).getValues(); }
  const out = [];
  for (let i = 0; i < n; i++) out.push([shadowEnc_((ah[i] || []).concat(bk[i] || []))]);
  sw.getRange(J_FIRST, 1, n, 1).setValues(out);
  sw.getRange(1, 1).setValue(JSON.stringify({ max: sh.getMaxRows(), head: sh.getRange(J_HEAD, 1, 1, J_INPUT_N).getValues()[0], at: Date.now() }));
  return true;
}

/** Не дасть збій сортуванню чи розкладу. Аркуш копії хтось видалив: створюється знову. */
function shadowKeep_(ss, ah, bk) {
  try {
    if (!ss.getSheetByName(SHADOW_SHEET) && ss.getSheetByName(JOURNAL)) { buildShadow_(ss); if (!ah) return true; }
    return shadowSave_(ss, ah, bk);
  } catch (e) { return false; }
}

/** Кожні 15 хв: копія, лише якщо журнал змінився (скрипт міг писати без сортування). */
function shadowIfStale_(ss) {
  const sh = ss.getSheetByName(JOURNAL), sw = ss.getSheetByName(SHADOW_SHEET);
  if (!sh || !sw) return false;
  const n = J_LAST - J_FIRST + 1, ah = sh.getRange(J_FIRST, 1, n, J_INPUT_N).getValues(), bk = sh.getRange(J_FIRST, colNum_(JC.bankId), n, J_MOVE_N).getValues();
  const old = sw.getRange(J_FIRST, 1, n, 1).getValues();
  if (ah.every((v, i) => shadowEnc_(v.concat(bk[i])) === String(old[i][0] || ''))) return false;
  return shadowSave_(ss, ah, bk);
}

/* ---------- стерли: повернути ---------- */

/**
 * Чиста функція: що повернути. saved: рядки з копії (або null), ah/bk: зараз (A…H, Y…AF), elsewhere(i, v, id): чи цей запис
 * є деінде в журналі (переїхав: сортування, вирізали-вставили). Захищені: з банку й ваші повні (дата, сума, категорія). → [{ i, v, b, n }]
 */
function journalKeep_(saved, ah, bk, elsewhere) {
  const out = [], same1 = (a, b) => (isDate_(a) || isDate_(b) ? isDate_(a) && isDate_(b) && a.getTime() === b.getTime() : String(a) === String(b));
  saved.forEach((S, i) => {
    if (!S) return;
    const sv = S.slice(0, J_INPUT_N), sb = S.slice(J_INPUT_N);
    if (sv.every(blank_) || isPlan_(sv[4])) return;                        // плани ставить скрипт: їх не тримаємо
    const cv = ah[i].slice(), cb = bk[i].slice(), sid = String(sb[0] || ''), cid = String(cb[0] || '');
    if (!sid && (blank_(sv[0]) || blank_(sv[1]) || blank_(sv[2]))) return;   // недописаний ваш рядок: стирайте, якщо треба
    if (sid && cid && sid !== cid) return;                                  // інший рядок: копія ще не оновилась
    const wiped = cv.every(blank_), byId = !!sid && cid === sid;
    const same = byId || (!wiped && cv.every((x, k) => blank_(x) || same1(x, sv[k])));   // усе, що лишилось, як у копії
    if (!same && (!wiped || cid)) return;                                   // тут уже інший запис
    if (wiped && !byId && elsewhere(i, sv, sid)) return;                    // рядок переїхав, а не стерся
    let n = 0;
    (wiped ? sv.map((x, k) => k) : KEEP_COLS).forEach(k => { if (blank_(cv[k]) && !blank_(sv[k])) { cv[k] = sv[k]; n++; } });
    sb.forEach((x, k) => { if (blank_(cb[k]) && !blank_(x)) { cb[k] = x; n++; } });
    if (n) out.push({ i: i, v: cv, b: cb, n: n });
  });
  return out;
}

/** Колонки з формулами журналу, що перетинаються з [c0, c1]. */
function formulaCols_(c0, c1) { return Object.keys(journalFormulas_()).filter(c => colNum_(c) >= c0 && colNum_(c) <= c1); }

/** onEdit у журналі: стерте записане повертається. → скільки клітинок повернуто */
function protectEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== JOURNAL) return 0;
  const c0 = e.range.getColumn(), c1 = e.range.getLastColumn(), one = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
  const cleared = blank_(e.value) || typeof e.value === 'object';           // стерта клітинка: value немає (буває { oldValue })
  if (one && c0 <= J_INPUT_N && (!cleared || KEEP_COLS.indexOf(c0 - 1) < 0)) return 0;   // вписали значення чи стерли необов'язкове
  const ss = sh.getParent();
  let fixed = 0;
  if (e.range.getRow() <= J_HEAD && cleared && (blank_(sh.getRange(1, 1).getValue()) || blank_(sh.getRange(J_HEAD, 1).getValue()))) {
    journalHeadBack_(ss, sh);                                               // стерли шапку (скажімо, «виділити все» й Delete)
    fixed++;
  }
  const r0 = Math.max(e.range.getRow(), J_FIRST), r1 = Math.min(e.range.getLastRow(), J_LAST);
  if (r1 < r0) return fixed;
  const sw = ss.getSheetByName(SHADOW_SHEET), y = colNum_(JC.bankId), n = r1 - r0 + 1;
  const fc = formulaCols_(c0, c1);
  if (fc.length) {                                                          // формули праворуч стерли: пишемо знову
    const fs = journalFormulas_();
    fc.forEach(c => { const rows = []; for (let r = r0; r <= r1; r++) rows.push([rowF_(fs[c], r)]); setF_(sh.getRange(r0, colNum_(c), n, 1), rows); });
    fixed += fc.length * n;
  }
  if (!sw || (c0 > J_INPUT_N && c1 < y) || c0 > y + J_MOVE_N - 1) return fixed;
  const saved = sw.getRange(r0, 1, n, 1).getValues().map(r => shadowDec_(r[0]));
  if (!saved.some(Boolean)) return fixed;
  const ah = sh.getRange(r0, 1, n, J_INPUT_N).getValues(), bk = sh.getRange(r0, y, n, J_MOVE_N).getValues();
  let all = null, ids = null;
  const elsewhere = (i, sv, id) => {
    const me = r0 + i - J_FIRST;
    if (id) {
      if (!ids) ids = sh.getRange(J_FIRST, y, J_LAST - J_FIRST + 1, 1).getValues().map(r => String(r[0] || ''));
      return ids.some((x, j) => j !== me && x === id);
    }
    if (!all) all = sh.getRange(J_FIRST, 1, J_LAST - J_FIRST + 1, J_INPUT_N).getValues().map(shadowEnc_);
    const k = shadowEnc_(sv);
    return all.some((x, j) => j !== me && x === k);
  };
  const back = journalKeep_(saved, ah, bk, elsewhere), by = {};
  back.forEach(x => { by[x.i] = x; fixed += x.n; });
  rowRuns_(back.map(x => x.i)).forEach(run => {                              // суцільні шматки одним записом: «виділити все» теж встигає
    const xs = [];
    for (let i = run[0]; i < run[0] + run[1]; i++) xs.push(by[i]);
    sh.getRange(r0 + run[0], 1, run[1], J_INPUT_N).setValues(xs.map(x => x.v));
    sh.getRange(r0 + run[0], y, run[1], J_MOVE_N).setValues(xs.map(x => x.b));
  });
  if (back.length) {
    try { ss.toast('Записане в журналі стирати не можна: повернув ' + back.length + ' ' + rowsWord_(back.length) + '. Прибрати рядки: меню «📊 ' + APP_NAME + ' → ' + WIPE_T + '…».', PROTECT_T, 10); }
    catch (x) { /* без підказки */ }
  }
  return fixed;
}

/** Шапка журналу знову: заголовки колонок і рядки 1…3 (нова шапка) або назва (стара). Дані не чіпає. */
function journalHeadBack_(ss, sh) {
  const last = Math.min(colNum_(J_SERVICE_LAST), sh.getMaxColumns()), heads = [];
  for (let c = 1; c <= last; c++) heads.push(J_HEADERS[colStr_(c)] || '');
  sh.getRange(J_HEAD, 1, 1, last).setValues([heads]);
  const d = ss.getSheetByName(DICT_SHEET);
  if (d && d.getMaxColumns() >= colNum_('U') && String(d.getRange('T5').getValue()) === 'charttype') journalHeader_(ss);
  else if (blank_(sh.getRange(1, 1).getValue())) sh.getRange(1, 1).setValue('📒 Журнал');
  try { ss.toast('Шапку журналу стирати не можна: написав знову.', PROTECT_T, 8); } catch (e) { /* без підказки */ }
}

function rowsWord_(n) { return n % 10 === 1 && n % 100 !== 11 ? 'рядок' : [2, 3, 4].indexOf(n % 10) >= 0 && [12, 13, 14].indexOf(n % 100) < 0 ? 'рядки' : 'рядків'; }

/* ---------- видалили рядки: вставити назад ---------- */

/** Чиста функція: де видалено k рядків (індекс у копії) або -1, якщо копія не збігається з тим, що лишилось. */
function rowsGone_(saved, cur, k) {
  let i = 0;
  while (i < cur.length && i < saved.length && cur[i] === saved[i]) i++;
  for (let j = i; j < cur.length && j + k < saved.length; j++) if (cur[j] !== saved[j + k]) return -1;
  return i;
}

/** Після видалення рядків: вставити з копії. → { n, err? } */
function journalRowsBack_(ss) {
  const sh = ss.getSheetByName(JOURNAL), sw = ss.getSheetByName(SHADOW_SHEET);
  if (!sh || !sw) return { n: 0 };
  let meta = {};
  try { meta = JSON.parse(String(sw.getRange(1, 1).getValue() || '{}')); } catch (e) { meta = {}; }
  const k = Number(meta.max || 0) - sh.getMaxRows();
  if (k <= 0) return { n: 0 };
  if (JSON.stringify(sh.getRange(J_HEAD, 1, 1, J_INPUT_N).getValues()[0]) !== JSON.stringify(meta.head || [])) return { n: 0, err: 'head' };
  const L = J_LAST - J_FIRST + 1, m = Math.min(L, sh.getMaxRows() - J_FIRST + 1), y = colNum_(JC.bankId);
  const saved = sw.getRange(J_FIRST, 1, L, 1).getValues().map(r => String(r[0] || ''));
  const ah = sh.getRange(J_FIRST, 1, m, J_INPUT_N).getValues(), bk = sh.getRange(J_FIRST, y, m, J_MOVE_N).getValues();
  const at = rowsGone_(saved, ah.map((v, i) => shadowEnc_(v.concat(bk[i]))), k);
  if (at < 0) return { n: 0, err: 'diff' };
  const R = J_FIRST + at, cols = sh.getMaxColumns();
  if (R <= sh.getMaxRows()) sh.insertRowsBefore(R, k); else sh.insertRowsAfter(sh.getMaxRows(), k);
  const src = R + k <= sh.getMaxRows() ? R + k : R - 1;
  if (src >= J_FIRST) {
    [SpreadsheetApp.CopyPasteType.PASTE_FORMAT, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION].forEach(t => {
      try { sh.getRange(src, 1, 1, cols).copyTo(sh.getRange(R, 1, k, cols), t, false); } catch (e) { /* формат не критичний */ }
    });
  }
  const rows = saved.slice(at, at + k).map(s => shadowDec_(s) || new Array(J_INPUT_N + J_MOVE_N).fill(''));
  sh.getRange(R, 1, k, J_INPUT_N).setValues(rows.map(v => v.slice(0, J_INPUT_N)));
  sh.getRange(R, y, k, J_MOVE_N).setValues(rows.map(v => v.slice(J_INPUT_N)));
  const fs = journalFormulas_();
  Object.keys(fs).forEach(c => { const f = []; for (let r = R; r < R + k; r++) f.push([rowF_(fs[c], r)]); setF_(sh.getRange(R, colNum_(c), k, 1), f); });
  shadowKeep_(ss);
  return { n: k, filled: rows.filter(v => !v.every(blank_)).length };
}

/** Розклад onChange: видалили рядки журналу → назад; вставили → оновити копію (рядки зсунулись). */
function mbsChange(e) {
  const type = String(e && e.changeType || '');
  if ((type !== 'REMOVE_ROW' && type !== 'INSERT_ROW') || !quietOk_()) return;
  const ss = SpreadsheetApp.getActive(), lock = LockService.getScriptLock();
  if (!lock.tryLock(120 * 1000)) { try { ss.toast('Скрипт зайнятий: видалені рядки журналу не повернув. Поверніть: Ctrl+Z (на Mac Cmd+Z).', PROTECT_T, 15); } catch (x) { /* ignore */ } return; }
  let r = { n: 0 };
  try { r = type === 'REMOVE_ROW' ? journalRowsBack_(ss) : (shadowKeep_(ss), { n: 0 }); }
  finally { lock.releaseLock(); }
  try {
    if (r.n) ss.toast('Рядки журналу видаляти не можна: повернув ' + r.n + ' ' + rowsWord_(r.n) + '. Прибрати рядки: меню «📊 ' + APP_NAME + ' → ' + WIPE_T + '…».', PROTECT_T, 10);
    else if (r.err) ss.toast((r.err === 'head' ? 'Видалено рядок шапки журналу.' : 'Видалено рядки журналу, а копія вже інша.') +
      ' Сам не повертаю, щоб не наплутати: Ctrl+Z (на Mac Cmd+Z) чи «Файл → Історія версій».', PROTECT_T, 20);
  } catch (x) { /* без підказки */ }
}

function ensureChangeTrigger_(ss) {
  if (!hasTrigger_(TRG.change)) ScriptApp.newTrigger(TRG.change).forSpreadsheet(ss).onChange().create();
}

/* ---------- 🧹 Зачистити дані ---------- */

/** Чиста функція: «усе» | дата | період | місяць → { all, from, to } або null. Дати: дд.мм.рррр, рррр-мм-дд; місяць: мм.рррр. */
function wipeRange_(text) {
  const s = String(text || '').trim().toLowerCase();
  if (/^(усе|все|all)$/.test(s)) return { all: true, from: '', to: '' };
  const re = /(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})[./](\d{1,2})[./](\d{4})|(\d{1,2})[./](\d{4})/g, got = [];
  let m;
  const day = (y, mo, d) => {
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return '';
    const k = y + '-' + pad2_(mo) + '-' + pad2_(d);
    return addDays_(k, 0) === k ? k : '';
  };
  while ((m = re.exec(s))) {
    if (m[1]) got.push({ a: day(+m[1], +m[2], +m[3]) });
    else if (m[4]) got.push({ a: day(+m[6], +m[5], +m[4]) });
    else { const mo = +m[7]; if (mo < 1 || mo > 12) return null; const ym = m[8] + '-' + pad2_(mo); got.push({ a: ym + '-01', b: lastDayOf_(ym) }); }
  }
  if (!got.length || got.length > 2 || got.some(x => !x.a)) return null;
  if (s.replace(re, ' ').replace(/із|від|по|до|з/g, '').replace(/[\s\-\u2013\u2014…,.]/g, '')) return null;   // зайві слова: краще перепитати
  let from = got[0].a, to = (got[got.length - 1].b || got[got.length - 1].a);
  if (got.length === 2 && got[1].a < from) { from = got[1].a; to = got[0].b || got[0].a; }
  return { all: false, from: from, to: to };
}

function wipeLabel_(g) { return g.all ? 'усе' : g.from === g.to ? dd_(g.from) + '.' + g.from.slice(0, 4) : dd_(g.from) + '.' + g.from.slice(0, 4) + '–' + dd_(g.to) + '.' + g.to.slice(0, 4); }

function wipeHit_(g, key) { return g.all || (!!key && key >= g.from && key <= g.to); }

/** Чиста функція: які рядки журналу прибрати. → { rows, bank, plans, mine, first } */
function wipeJournal_(journal, g) {
  const out = { rows: [], bank: 0, plans: 0, mine: 0, first: '' };
  journal.forEach(o => {
    if (o.blank || !wipeHit_(g, o.key)) return;
    out.rows.push(o.row);
    if (o.bank && o.bank[0]) out.bank++; else if (isPlan_(o.e) || Number(o.m) === 1) out.plans++; else out.mine++;
    if (o.key && (!out.first || o.key < out.first)) out.first = o.key;
  });
  return out;
}

/** Чиста функція: записи «Банк» (з BANK_LOG_FIRST), що лишаються. Дата: колонка C. → { keep, n, first } */
function wipeLog_(values, g, tz) {
  const keep = [];
  let n = 0, first = '';
  values.forEach(r => {
    if (r.every(blank_)) return;
    const key = dateKey_(r[2], tz);
    if (!wipeHit_(g, key)) { keep.push(r); return; }
    n++;
    if (key && (!first || key < first)) first = key;
  });
  return { keep: keep, n: n, first: first };
}

/** Чиста функція: суцільні шматки рядків → [[перший, скільки]]. */
function rowRuns_(rows) {
  const out = [];
  rows.slice().sort((a, b) => a - b).forEach(r => { const l = out[out.length - 1]; if (l && l[0] + l[1] === r) l[1]++; else out.push([r, 1]); });
  return out;
}

function wipePreview_(ss, g) {
  const b = ss.getSheetByName(BANK_SHEET), last = b.getLastRow(), tz = ss.getSpreadsheetTimeZone();
  const log = last >= BANK_LOG_FIRST ? wipeLog_(b.getRange(BANK_LOG_FIRST, 1, last - BANK_LOG_FIRST + 1, BANK_LOG_N).getValues(), g, tz) : { keep: [], n: 0, first: '' };
  return { j: wipeJournal_(readJournal_(ss), g), log: log };
}

/** Зачистка (копію таблиці вже зроблено). → { rows, bank, plans, mine, logs, first } */
function wipeRun_(ss, g) {
  guard_(ss);
  const sh = ss.getSheetByName(JOURNAL), b = ss.getSheetByName(BANK_SHEET), tz = ss.getSpreadsheetTimeZone(), y = colNum_(JC.bankId);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) throw new Error('Інша дія скрипта ще триває (імпорт?), спробуйте за хвилину. Нічого не змінено.');
  try {
    const p = wipePreview_(ss, g);
    rowRuns_(p.j.rows).forEach(x => {
      sh.getRange(x[0], 1, x[1], J_INPUT_N).setValues(new Array(x[1]).fill(0).map(() => new Array(J_INPUT_N).fill('')));
      sh.getRange(x[0], y, x[1], J_MOVE_N).setValues(new Array(x[1]).fill(0).map(() => emptyBank_().concat(['', ''])));
    });
    if (p.log.n) {
      const last = b.getLastRow();
      b.getRange(BANK_LOG_FIRST, 1, last - BANK_LOG_FIRST + 1, BANK_LOG_N).clearContent();
      if (p.log.keep.length) b.getRange(BANK_LOG_FIRST, 1, p.log.keep.length, BANK_LOG_N).setValues(p.log.keep);
    }
    PropertiesService.getDocumentProperties().deleteProperty('stmt_done');     // виписки з папки теж можна завантажити знову (дублів не буде)
    sortCore_(ss);
    const first = [p.j.first, p.log.first].filter(Boolean).sort()[0] || '';
    return { rows: p.j.rows.length, bank: p.j.bank, plans: p.j.plans, mine: p.j.mine, logs: p.log.n, first: first };
  } finally {
    lock.releaseLock();
    try { if (ss.getSheetByName(REG_SHEET)) regPlansSync_(ss); } catch (e) { /* плани поставить розклад */ }
    try { if (modOn_('debts') && ss.getSheetByName(DEBTS_SHEET)) debtsPlansSync_(ss); } catch (e) { /* плани поставить розклад */ }
    shadowKeep_(ss);
  }
}

/** Меню «🧹 Зачистити дані…»: що прибрати, підтвердження, копія, зачистка, за бажанням одразу імпорт з Mono знову. */
function wipeNow() {
  if (!menuOk_(WIPE_T)) return;
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi(), up = PropertiesService.getUserProperties();
  if (!ss.getSheetByName(JOURNAL) || !ss.getSheetByName(BANK_SHEET)) { show_(WIPE_T, ['Немає аркушів журналу чи банку: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».']); return; }
  if (up.getProperty('period_job')) { show_(WIPE_T, ['Зараз іде «' + PERIOD_T + '». Спершу зупиніть його (той самий пункт меню → «Так»), потім зачищайте.']); return; }
  const today = todayKey_(ss), ex = dd_(today) + '.' + today.slice(0, 4);
  const a = ui.prompt(WIPE_T, 'Що прибрати з журналу? Імпорт з банку потім завантажить ці дні знову.\n\n' +
    '• усе\n• дата: ' + ex + '\n• період: 01.' + today.slice(5, 7) + '.' + today.slice(0, 4) + '-' + ex + '\n• місяць: ' + today.slice(5, 7) + '.' + today.slice(0, 4) +
    '\n\nДалі буде підтвердження з кількістю рядків.', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const g = wipeRange_(a.getResponseText());
  if (!g) { show_(WIPE_T, ['⚠️ Не зрозумів «' + esc_(String(a.getResponseText() || '').slice(0, 60)) + '». Напишіть «усе», дату ' + ex + ', період 01.09.' + today.slice(0, 4) + '-30.09.' + today.slice(0, 4) + ' чи місяць 09.' + today.slice(0, 4) + '.']); return; }
  const p = wipePreview_(ss, g), label = wipeLabel_(g);
  if (!p.j.rows.length && !p.log.n) { show_(WIPE_T, ['За «' + label + '» у журналі й на аркуші «' + BANK_SHEET + '» нічого немає.']); return; }
  const ok = ui.alert(WIPE_T, 'Зачистити «' + label + '»?\n\n' +
    '• рядків журналу: ' + p.j.rows.length + ' (з банку ' + p.j.bank + ', ваших ' + p.j.mine + ', планів ' + p.j.plans + ')\n' +
    '• записів на аркуші «' + BANK_SHEET + '»: ' + p.log.n + ' (без них імпорт завантажить ці операції знову)\n\n' +
    'Перед змінами: копія таблиці на Диску. Плани з «Регулярних» і «Боргів» з\'являться знову самі. «Довідник», «Правила», «Регулярні», «Борги», ' +
    '«Налаштування», залишки банку й «Історія місяців» не змінюються.\n\nТочно зачистити?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  let copy;
  try { copy = backup_(ss, 'ручна'); } catch (e) { show_(WIPE_T, ['⚠️ Копію зробити не вдалося, нічого не змінено: ' + esc_(e.message)]); return; }
  let r;
  try { r = wipeRun_(ss, g); } catch (e) { show_(WIPE_T, ['⚠️ ' + esc_(e.message)]); return; }
  const lines = ['Зачищено «' + esc_(label) + '» ✓ Рядків журналу: <b>' + r.rows + '</b> (з банку ' + r.bank + ', ваших ' + r.mine + ', планів ' + r.plans + '), записів «' + BANK_SHEET + '»: ' + r.logs + '.',
    '💾 Копія перед змінами: ' + link_(copy.url, copy.name)];
  const accs = up.getProperty('mono_token') ? monoAccs_() : [], min = addDays_(monthsBack_(today, PERIOD_MAX_MONTHS), 1);
  let from = g.all ? r.first : g.from;
  if (from && from < min) from = min;
  if (accs.length && from && from <= today) {
    const again = ui.alert(WIPE_T, 'Зачищено ✓ Завантажити з Mono знову з ' + dd_(from) + '.' + from.slice(0, 4) + ' по сьогодні' +
      (accs.length > 1 ? ' (картка «' + accs[0].name + '»; інші: «' + PERIOD_T + '…»)' : '') + '?', ui.ButtonSet.YES_NO);
    if (again === ui.Button.YES) { show_(WIPE_T, lines.concat(periodBegin_(ss, accs[0], wipeMonths_(from, today), from))); return; }
  }
  lines.push('Завантажити знову: «' + PERIOD_T + '…» (Mono), «🏦 Імпорт з банку зараз» (останні дні) чи «📥 Імпорт виписок з папки».');
  show_(WIPE_T, lines);
}

/** Чиста функція: скільки місяців назад (1…12) покриває дату from. */
function wipeMonths_(from, today) {
  for (let m = 1; m < PERIOD_MAX_MONTHS; m++) if (addDays_(monthsBack_(today, m), 1) <= from) return m;
  return PERIOD_MAX_MONTHS;
}
