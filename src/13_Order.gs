/** 13_Order.gs: Журнал: автодата, автокатегорія, «навчання» правил, порядок рядків і розклад (mbsTick). */

/*
 * Автодата: внесли суму чи категорію, а дату ні: стане сьогоднішня.
 * Автокатегорія: категорія порожня, а в коментарі є слово з «Правил» (де шукати: коментар або будь-де).
 * «Навчання»: змінили категорію рядка з банку → на «Правилах» з'являється пропозиція «✨ нове, підтвердіть».
 * Порядок: майбутнє (найближче зверху) → жовтий рядок-лінія → 3 порожні рядки для вводу → минуле від нового до старого.
 * Сортує сам: за хвилину після правки (коли 45 с тиші) і кожні 15 хв (страховка: новий день, зміни від скрипта).
 * Недописаний рядок (без суми чи категорії) чекає до 10 хв, потім більше не блокує порядок.
 */

const IDLE_MS = 45 * 1000, INCOMPLETE_WAIT_MS = 10 * 60 * 1000, SORT_SOON_MS = 60 * 1000, TICK_MIN = 15;
const INPUT_ROWS = 3;

function onEdit(e) {
  if (!e || !e.range) return;
  if (!quietOk_()) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== JOURNAL) return;
  const r0 = e.range.getRow(), r1 = e.range.getLastRow(), c0 = e.range.getColumn(), c1 = e.range.getLastColumn();
  if (r1 < J_FIRST || c0 > J_INPUT_N) return;
  PropertiesService.getDocumentProperties().setProperty('lastEdit', String(Date.now()));
  const top = Math.max(r0, J_FIRST), n = Math.min(r1, J_LAST) - top + 1;
  if (n < 1 || n > 50 || c1 < 2) return;
  const ss = sh.getParent(), vals = sh.getRange(top, 1, n, 6).getValues(), today = todayString_();
  let rules = null;
  vals.forEach((v, i) => {
    const row = top + i;
    if (v[0] === '' && (v[1] !== '' || v[2] !== '')) sh.getRange(row, 1).setValue(Utilities.parseDate(today, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'));
    if (v[2] === '' && v[3] !== '' && c0 <= 4) {
      if (!rules) rules = readRules_(ss);
      const hit = matchRule_(rules, { comment: String(v[3]) }, 0, true);
      if (hit) {
        sh.getRange(row, 3).setValue(hit.category);
        if (hit.hole && v[5] === '') sh.getRange(row, 6).setValue(hit.hole);
      }
    }
  });
  if (n === 1 && c0 <= 3 && c1 >= 3) {
    const cat = String(vals[0][2] || '');
    if (cat && cat !== CAT_TRANSFER) {
      const bank = sh.getRange(top, colNum_(JC.bankId), 1, 2).getValues()[0];
      if (bank[0] && bank[1]) learnRule_(ss, String(bank[1]), cat, today);
    }
  }
}

/* ---------- розклад ---------- */

/** Страховка кожні 15 хв + запуск після кожної правки (mbsEdit → mbsSoon за хвилину). Повторний виклик дублів не створює. */
function ensureTickTriggers_(ss) {
  if (!hasTrigger_(TRG.tick)) ScriptApp.newTrigger(TRG.tick).timeBased().everyMinutes(TICK_MIN).create();
  if (!hasTrigger_(TRG.edit)) ScriptApp.newTrigger(TRG.edit).forSpreadsheet(ss).onEdit().create();
}

/** Правка в журналі (A…H): упорядкувати за хвилину. Правка модулів на «Старті»: застосувати одразу. */
function mbsEdit(e) {
  if (!e || !e.range || !quietOk_()) return;
  if (isModulesEdit_(e.range)) { try { modulesIfChanged_(e.range.getSheet().getParent()); } catch (x) { /* застосує mbsTick */ } return; }
  const name = e.range.getSheet().getName();
  if (name !== JOURNAL || e.range.getLastRow() < J_FIRST || e.range.getColumn() > J_INPUT_N) return;
  soonPlan_(SORT_SOON_MS);
}

function soonPlan_(ms) {
  const dp = PropertiesService.getDocumentProperties();
  if (Number(dp.getProperty('sort_soon') || 0) > Date.now() - 2 * 60 * 1000) return;
  ScriptApp.newTrigger(TRG.soon).timeBased().after(ms).create();
  dp.setProperty('sort_soon', String(Date.now() + ms));
}

function mbsSoon() {
  removeTriggers_(TRG.soon);
  if (!quietOk_()) return;
  const dp = PropertiesService.getDocumentProperties();
  dp.deleteProperty('sort_soon');
  const quiet = Date.now() - Number(dp.getProperty('lastEdit') || 0);
  if (quiet < IDLE_MS) { soonPlan_(IDLE_MS - quiet + 5000); return; }
  mbsTick();
}

/**
 * Кожні 15 хв і після правки: нова версія таблиці (кроки), модулі, порядок журналу (коли 45 с тиші), закриття місяця
 * в «Історії місяців», плани ФОП і боргів, календар платежів, чекліст, знімок для радника (раз на день), перевірка
 * нової версії шаблону (раз на день). Нічого не змінилось: лише кілька властивостей, без читання журналу.
 */
function mbsTick() {
  if (!quietOk_()) return;
  beat_('sort', true);
  const ss = SpreadsheetApp.getActive();
  try { if (schemaTick_(ss)) return; } catch (e) { /* звіт: «🩺 Стан таблиці» */ }
  try { modulesIfChanged_(ss); } catch (e) { /* наступного разу */ }
  const dp = PropertiesService.getDocumentProperties();
  const last = Number(dp.getProperty('lastEdit') || 0);
  if (Date.now() - last < IDLE_MS) return;
  const today = todayString_();
  if (last > Number(dp.getProperty('sort_seen') || 0) || dp.getProperty('sort_day') !== today) {
    const r = sortJournal_(ss);
    if (r !== null) dp.setProperties({ sort_seen: String(last), sort_day: today });
  }
  try { monthCloseIfDue_(ss); } catch (e) { /* наступного разу */ }
  try { fopIfDue_(ss); } catch (e) { /* плани ФОП: наступного разу */ }
  try { debtsIfDue_(ss); } catch (e) { /* плани боргів: наступного разу */ }
  try { calIfDue_(ss); } catch (e) { /* помилка в cal_err */ }
  try { startIfDue_(ss); } catch (e) { /* чекліст: наступного разу */ }
  try { advisorIfDue_(ss); } catch (e) { /* знімок: помилка в adv_err */ }
  try { verIfDue_(ss); } catch (e) { /* перевірка версії: завтра */ }
}

function sortNow() {
  if (!menuOk_('🔃 Порядок журналу')) return;
  const r = sortJournal_(SpreadsheetApp.getActive());
  if (r === null) show_('🔃 Порядок журналу', ['Скрипт зайнятий іншою дією, спробуйте за хвилину.']);
}

/* ---------- сортування ---------- */

function sortJournal_(ss) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20 * 1000)) return null;
  try { return sortCore_(ss); } finally { lock.releaseLock(); }
}

/** Саме сортування (блокування вже взяте). Пише A…H, Y…AE; формули I, J, L…X лишаються на місці (вони рядкові). */
function sortCore_(ss) {
  const sh = ss.getSheetByName(JOURNAL);
  if (!sh) return false;
  const n = J_LAST - J_FIRST + 1, y = colNum_(JC.bankId), tz = ss.getSpreadsheetTimeZone();
  const ah = sh.getRange(J_FIRST, 1, n, J_INPUT_N).getValues(), bk = sh.getRange(J_FIRST, y, n, J_MOVE_N).getValues();
  const rows = ah.map((v, i) => ({ v: v, b: bk[i].slice(0, BANK_N), s: bk[i][BANK_N], a: bk[i][BANK_N + 1], key: dateKey_(v[0], tz) }));
  const waitInc = Date.now() - Number(PropertiesService.getDocumentProperties().getProperty('lastEdit') || 0) < INCOMPLETE_WAIT_MS;
  const plan = planOrder_(rows, todayString_(), { now: Date.now(), waitIncomplete: waitInc });
  if (!plan) return false;
  sh.getRange(J_FIRST, 1, plan.length, J_INPUT_N).setValues(plan.map(x => x.v));
  const val = x => (x === undefined || x === null ? '' : x);
  sh.getRange(J_FIRST, y, plan.length, J_MOVE_N).setValues(plan.map(x => (x.b || emptyBank_()).concat([val(x.s), val(x.a)])));
  return true;
}

function emptyBank_() { return ['', '', '', '', '', '']; }

/**
 * Чиста функція: новий порядок або null (міняти нічого не треба, або є недописаний рядок і ще чекаємо).
 * rows: [{v: A…H, b: банк Y…AD, s: черговість, key}]. Черговість s (мс): у рядків з банку: час операції,
 * у ручних: коли сортування вперше їх побачило. У межах дня: майбутнє у порядку запису, минуле: найновіше зверху.
 */
function planOrder_(rows, today, opt) {
  const empty = c => c === '' || c === null || c === undefined;
  const blank = x => x.v.every(empty);
  const hasBank = x => !!(x.b && x.b.some(c => !empty(c))) || !empty(x.a);
  const used = rows.map((x, i) => ({ x: x, i: i })).filter(o => !blank(o.x));
  if ((!opt || opt.waitIncomplete !== false) && used.some(o => o.x.key === '' || (!hasBank(o.x) && (empty(o.x.v[1]) || empty(o.x.v[2]))))) return null;
  let fresh = false;
  used.forEach(o => {
    o.k = o.x.key || today;
    if (empty(o.x.s) || isNaN(Number(o.x.s))) { o.x.s = opt.now + o.i; fresh = true; }
  });
  const byKey = (a, b) => (a.k < b.k ? -1 : a.k > b.k ? 1 : 0);
  const seq = (a, b) => (Number(a.x.s) - Number(b.x.s)) || (a.i - b.i);
  const future = used.filter(o => o.k > today).sort((a, b) => byKey(a, b) || seq(a, b));
  const past = used.filter(o => o.k <= today).sort((a, b) => byKey(b, a) || seq(b, a));
  const gap = Math.min(1 + INPUT_ROWS, rows.length - used.length);
  const none = () => ({ v: ['', '', '', '', '', '', '', ''], b: null, s: '', a: '' });
  const order = future.map(o => o.x);
  for (let k = 0; k < gap; k++) order.push(none());
  past.forEach(o => order.push(o.x));
  let lastIdx = -1;
  rows.forEach((x, i) => { if (!blank(x) || hasBank(x) || !empty(x.s)) lastIdx = i; });
  while (order.length < Math.max(lastIdx + 1, used.length + gap)) order.push(none());
  const clean = x => blank(x) && !hasBank(x) && empty(x.s);
  const same = !fresh && order.every((x, i) => x === rows[i] || (blank(x) && clean(rows[i])));
  return same ? null : order;
}
