/** 32_Reports.gs: Бюджети категорій, тижневий звіт (понеділок), закриття місяця й підсумок місяця (1-го числа). */

/*
 * Обидва звіти йдуть у Telegram одразу після ранкового підсумку, кожен окремим повідомленням:
 * тижневий у понеділок (за минулий пн…нд), місячний 1-го числа (за минулий місяць). Не вийшло: наступного ранку.
 * Закриття місяця: 1-го числа після 05:00 скрипт записує минулий місяць в «Історію місяців» значеннями
 * (пізні правки в журналі історію вже не змінюють). Перезаписати свідомо: меню «📚 Записати місяць в «Історію»…».
 * Бюджети: категорії типу «Витрати» з планом на місяць у «Довіднику» (колонка F).
 */

const MONTH_CLOSE_HOUR = 5;

/* ---------- бюджети ---------- */

/** Плани категорій з «Довідника»: { категорія: { plan, who } } (лише тип «Витрати», план > 0; who: порожньо = спільний бюджет). */
function categoryPlans_(ss) {
  const d = ss.getSheetByName(DICT_SHEET), out = {};
  if (!d) return out;
  d.getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 7).getValues().forEach(r => {
    if (String(r[0]).trim() && String(r[1]) === T_SPEND && Number(r[5]) > 0) out[String(r[0]).trim()] = { plan: Number(r[5]), who: String(r[6] || '').trim() };
  });
  return out;
}

/**
 * Бюджети місяця: витрачено (факти + чернетки з банку) проти плану категорії і норма на сьогодні.
 * rows: readJournal_; cats: лише ці категорії (null: усі). → [{ over, fast, text, cat, owner, budget, spent, draft, norm }]
 * Модуль «💰 Бюджети категорій» вимкнено: порожньо (без тривог). force: рахувати все одно (знімок для радника).
 */
function budgetLines_(ss, rows, cats, force) {
  if (!force && !modOn_('budgets')) return [];
  const plans = categoryPlans_(ss), today = todayKey_(ss), ym = today.slice(0, 7), out = [];
  const dayN = Number(today.slice(8)), monthN = Number(lastDayOf_(ym).slice(8));
  Object.keys(plans).filter(c => !cats || cats.indexOf(c) >= 0).forEach(c => {
    let fact = 0, draft = 0;
    const owner = plans[c].who;
    (rows || []).forEach(o => {
      if (o.c !== c || String(o.key).slice(0, 7) !== ym || o.kind === K_FOP || (owner && o.g !== owner)) return;
      if (o.e === STATUS_DRAFT) draft += Math.max(0, -Number(o.bank[2]) || Number(o.b) || 0);
      else if (Number(o.m) === 0 && !isPlan_(o.e)) fact += Number(o.o) || Number(o.b) || 0;
    });
    const budget = plans[c].plan, spent = round2_(fact + draft), norm = budget * dayN / monthN;
    const over = spent > budget + 0.005, fast = spent > norm * 1.1;
    out.push({ cat: c, owner: owner, budget: budget, spent: spent, draft: round2_(draft), norm: round2_(norm), over: over, fast: fast, text: '💰 ' + esc_(c) + (owner ? ' (' + esc_(owner) + ')' : '') + ': ' + (over ? '<b>перевищено на ' + money0_(spent - budget) + '</b>' : 'лишилось <b>' + money0_(budget - spent) + '</b>') +
      ' з ' + money0_(budget) + ' (витрачено ' + money0_(spent) + (draft ? ', з них у чернетках ' + money0_(draft) : '') + ') · норма на сьогодні ' + money0_(norm) +
      (fast ? ' <b>▲</b>' : ' ✓') });
  });
  return out;
}

/* ---------- з журналу ---------- */

function addTo_(o, k, v) { o[k] = (o[k] || 0) + v; }
function topOf_(o, n) { return Object.keys(o).map(k => [k, o[k]]).filter(x => Math.abs(x[1]) > 0.005).sort((a, b) => b[1] - a[1]).slice(0, n || 5); }
function pairsText_(pairs) { return pairs.map(x => esc_(x[0]) + ' ' + money0_(x[1])).join('; '); }

/** Факти журналу за from…to: витрати (тип «Витрати») і за категоріями, непомітні, з ліміту й чому, чернетки-витрати окремо. */
function factStats_(rows, from, to) {
  const s = { spent: 0, cats: {}, holes: {}, holeSum: 0, limit: 0, reasons: {}, drafts: 0, draftSum: 0 };
  rows.forEach(o => {
    if (o.key < from || o.key > to || o.kind === K_FOP) return;
    if (o.e === STATUS_DRAFT) { if (Number(o.bank[2]) < 0 && (o.i === T_SPEND || o.i === '')) { s.drafts++; s.draftSum -= Number(o.bank[2]); } return; }
    if (Number(o.m) !== 0 || isPlan_(o.e)) return;
    const v = Number(o.o) || 0;
    if (o.i === T_SPEND) { s.spent += v; addTo_(s.cats, o.c || 'без категорії', v); }
    if (o.p) { s.holeSum += v; addTo_(s.holes, o.p, v); }
    if (Number(o.w) > 0.005) { s.limit += Number(o.w); addTo_(s.reasons, o.x || o.h || 'без причини', Number(o.w)); }
  });
  return s;
}

/** Тиждень: key: понеділок поточного тижня; from…to: минулий пн…нд; prevFrom…prevTo: позаминулий. */
function weekOf_(today) {
  const mon = addDays_(today, -((dowOf_(today) + 6) % 7));
  return { key: mon, from: addDays_(mon, -7), to: addDays_(mon, -1), prevFrom: addDays_(mon, -14), prevTo: addDays_(mon, -8) };
}

function weeklyText_(ss, today) {
  const w = weekOf_(today), rows = readJournal_(ss).filter(o => !o.blank && o.key);
  const s = factStats_(rows, w.from, w.to), p = factStats_(rows, w.prevFrom, w.prevTo);
  const diff = s.spent - p.spent;
  const lines = ['📊 <b>Тиждень ' + dd_(w.from) + '–' + dd_(w.to) + '</b>',
    'Витрачено: <b>' + money0_(s.spent) + '</b>' + (p.spent ? ' (тиждень до того ' + money0_(p.spent) + ', ' + (diff >= 0 ? '+' : '−') + money0_(Math.abs(diff)) + ')' : '')];
  const top = topOf_(s.cats, 5);
  if (top.length) lines.push('Найбільше: ' + pairsText_(top));
  if (s.holeSum > 0.005) lines.push('🕳 Непомітні: ' + money0_(s.holeSum) + ' (' + pairsText_(topOf_(s.holes, 4)) + ')');
  if (s.limit > 0.005 && modOn_('credit')) lines.push('💳 З ліміту: ' + money0_(s.limit) + ' (' + pairsText_(topOf_(s.reasons, 4)) + ')');
  if (s.drafts) lines.push('🏦 Ще в чернетках: ' + s.drafts + ' на ' + money0_(s.draftSum) + ' (не враховано вище)');
  return lines.join('\n');
}

/* ---------- місяць ---------- */

/** Рядок «Історії місяців» за ym як { колонка: значення } або null. */
function historyRow_(ss, ym) {
  const sh = ss.getSheetByName(HISTORY_SHEET);
  if (!sh || sh.getLastRow() < H_FIRST) return null;
  const v = sh.getRange(H_FIRST, 1, sh.getLastRow() - H_FIRST + 1, H_COLS.length).getValues();
  const row = v.find(r => String(r[0]) === ym || (Object.prototype.toString.call(r[0]) === '[object Date]' && dateKey_(r[0], ss.getSpreadsheetTimeZone()).slice(0, 7) === ym));
  if (!row) return null;
  const o = {};
  H_COLS.forEach((h, i) => { o[h] = row[i]; });
  return o;
}

function monthText_(ss, ym) {
  const h = historyRow_(ss, ym), p = historyRow_(ss, prevYm_(ym));
  if (!h) return '📅 <b>' + ymName_(ym) + '</b>: в «Історії місяців» ще немає рядка.';
  const n = k => Number(h[k]) || 0, cmp = k => (p && p[k] !== '' && p[k] !== null ? ' (місяць до того ' + money0_(p[k]) + ')' : ''), credit = modOn_('credit');
  const lines = ['📅 <b>Підсумок: ' + ymName_(ym) + '</b>',
    'Витрати: <b>' + money0_(n(T_SPEND)) + '</b>' + cmp(T_SPEND) + ' · усі списання ' + money0_(n('Усі списання')),
    'Надходження: ' + money0_(n('Всього надходжень')) + ' · мої ' + money0_(n('Мої надходження')),
    '💚 З моїх ' + money0_(n('Витрачено з моїх')) + (credit ? ' · 💳 з ліміту ' + money0_(n('Витрачено з ліміту')) : ''),
    credit ? '🎯 Погашено за бюджетний місяць: ' + money0_(n('Погашено за бюджетний місяць (відкладено + невитрачене)')) +
      (h['Борг за лімітом на кінець місяця'] !== '' ? ' · борг на кінець ' + money0_(n('Борг за лімітом на кінець місяця')) : '') : ''];
  if (h['Топ-5 категорій']) lines.push('Найбільше: ' + esc_(h['Топ-5 категорій']));
  if (n('Непомітні разом')) lines.push('🕳 Непомітні: ' + money0_(n('Непомітні разом')) + (h['Непомітні: з чого'] ? ' (' + esc_(h['Непомітні: з чого']) + ')' : ''));
  if (credit && h['З ліміту: причини']) lines.push('Чому з ліміту: ' + esc_(h['З ліміту: причини']));
  const out = lines.filter(x => x !== ''), sup = supportLine_(ss, ym);
  if (sup) out.push('', sup);
  return out.join('\n');
}

/**
 * «❤️ Підтримати автора» у підсумку місяця: не частіше разу на квартал (1-й місяць кварталу), якщо не вимкнено
 * на «Налаштуваннях» (B24). Жодних функцій «за донат».
 */
function supportLine_(ss, ym) {
  const set = ss.getSheetByName(SETTINGS_SHEET);
  if (set && String(set.getRange(SET.support).getValue()).trim() === 'ні') return '';
  const dp = PropertiesService.getDocumentProperties(), q = ym.slice(0, 4) + '-Q' + Math.ceil(Number(ym.slice(5, 7)) / 3);
  if (dp.getProperty('sup_q') === q) return '';
  dp.setProperty('sup_q', q);
  return '❤️ Шаблон безкоштовний. Корисний? Підтримати автора: ' + SUPPORT_URL + ' (вимкнути цей рядок: «' + SETTINGS_SHEET + '», B24 «ні»)';
}

/** Старт обліку з «Налаштувань» як 'yyyy-MM-dd' ('' якщо не задано). */
function startKey_(ss) {
  const s = ss.getSheetByName(SETTINGS_SHEET);
  return s ? dateKey_(s.getRange(SET.start).getValue(), ss.getSpreadsheetTimeZone()) : '';
}

/** 1-го числа після 05:00 (і далі, якщо пропустили): минулий місяць в «Історію місяців». Місяці до старту обліку не пишуться. */
function monthCloseIfDue_(ss) {
  const tz = ss.getSpreadsheetTimeZone(), dp = PropertiesService.getDocumentProperties(), today = todayKey_(ss), ym = prevYm_(today.slice(0, 7));
  if (dp.getProperty('mh_last') === ym || !ss.getSheetByName(HISTORY_SHEET)) return null;
  if (ym < startKey_(ss).slice(0, 7)) { dp.setProperty('mh_last', ym); return null; }
  if (today.slice(8) === '01' && Number(Utilities.formatDate(new Date(), tz, 'HH')) < MONTH_CLOSE_HOUR) return null;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return null;
  try {
    const r = historyWrite_(ss, ym);
    dp.setProperty('mh_last', ym);
    return r;
  } finally { lock.releaseLock(); }
}

/** Після ранкового підсумку: тиждень (понеділок) і місяць (1-го), кожен окремим повідомленням; не вийшло: наступного ранку. */
function extraReports_(ss) {
  const dp = PropertiesService.getDocumentProperties(), today = todayKey_(ss), w = weekOf_(today), ym = prevYm_(today.slice(0, 7));
  const out = [], start = startKey_(ss);
  if (dp.getProperty('rep_week') !== w.key) {
    if (w.to >= start) { tgSend_(weeklyText_(ss, today)); out.push('тиждень'); }        // тиждень до старту обліку: без звіту
    dp.setProperty('rep_week', w.key);
  }
  if (dp.getProperty('rep_month') !== ym && ym < start.slice(0, 7)) dp.setProperty('rep_month', ym);
  if (dp.getProperty('rep_month') !== ym) {
    try { monthCloseIfDue_(ss); } catch (e) { /* звіт напише, що рядка ще немає */ }
    tgSend_(monthText_(ss, ym)); dp.setProperty('rep_month', ym); out.push('місяць');
  }
  return out;
}

/* ---------- меню ---------- */

function weeklyNow() {
  if (!menuOk_('📊 Тиждень')) return;
  const ss = SpreadsheetApp.getActive();
  try { tgSend_(weeklyText_(ss, todayKey_(ss))); show_('📊 Тиждень', ['Надіслано ✓.']); } catch (e) { show_('📊 Тиждень', ['⚠️ ' + esc_(e.message)]); }
}

function monthNow() {
  if (!menuOk_('📅 Підсумок місяця')) return;
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi(), prev = prevYm_(todayKey_(ss).slice(0, 7));
  const a = ui.prompt('📅 Підсумок місяця', 'Який місяць (рррр-мм)? Порожньо: ' + prev + '. Місяць спершу перезапишеться в «Історії місяців».', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const ym = String(a.getResponseText() || '').trim() || prev;
  if (!/^\d{4}-\d{2}$/.test(ym)) { show_('📅 Підсумок місяця', ['⚠️ Потрібно рррр-мм, наприклад ' + prev + '.']); return; }
  try { historyWrite_(ss, ym); tgSend_(monthText_(ss, ym)); show_('📅 Підсумок місяця', ['Записано в «Історію місяців» і надіслано ✓.']); }
  catch (e) { show_('📅 Підсумок місяця', ['⚠️ ' + esc_(e.message)]); }
}
