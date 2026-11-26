/** 02_Helpers.gs: Дрібні допоміжні функції. */

function colNum_(s) { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function colStr_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/*
 * Роздільник аргументів у формулах. Таблиця з регіоном «Україна» (і більшість європейських) чекає «;» між аргументами
 * і «,» у дробових числах; регіон «США»: «,» і «.». Код пише формули з «,» і перед записом перекладає їх під таблицю.
 * Який роздільник, визначає проба: формула =SUM(1,2) у порожній клітинці дає 3 лише там, де роздільник «,».
 */
let FSEP_ = null;
function fsep_(ss) {
  if (FSEP_) return FSEP_;
  const saved = PropertiesService.getDocumentProperties().getProperty('fsep');
  if (saved === ',' || saved === ';') return (FSEP_ = saved);
  return (FSEP_ = fsepProbe_(ss || SpreadsheetApp.getActive()));
}

function fsepProbe_(ss) {
  const sh = ss.getSheetByName(SETTINGS_SHEET) || ss.getSheets()[0], c = sh.getRange('Z50');
  let sep = ',';
  try { c.setFormula('=SUM(1,2)'); SpreadsheetApp.flush(); sep = Number(c.getValue()) === 3 ? ',' : ';'; }
  catch (e) { sep = ';'; }
  finally { try { c.clearContent(); } catch (e) { /* ignore */ } }
  PropertiesService.getDocumentProperties().setProperty('fsep', sep);
  FSEP_ = sep;
  return sep;
}

/** Формула з «,» → у форматі таблиці: «,» між аргументами → «;», 0.5 → 0,5 (поза текстом у лапках). */
function loc_(formula, sep) {
  if (sep !== ';') return formula;
  let out = '', inStr = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula.charAt(i);
    if (ch === '"') inStr = !inStr;
    if (!inStr && ch === ',') { out += ';'; continue; }
    if (!inStr && ch === '.' && /\d/.test(formula.charAt(i - 1)) && /\d/.test(formula.charAt(i + 1))) { out += ','; continue; }
    out += ch;
  }
  return out;
}

/** Масив формул (рядки × колонки) → у форматі таблиці; порожні клітинки лишаються порожніми. */
function locRows_(rows, ss) {
  const sep = fsep_(ss);
  return rows.map(r => r.map(f => (String(f).charAt(0) === '=' ? loc_(String(f), sep) : '')));
}

/** Записати формули (у форматі таблиці). */
function setF_(range, rows) { return range.setFormulas(locRows_(rows, range.getSheet().getParent())); }

/** Формула з {r} → номер рядка. */
function rowF_(f, r) { return f.split('{r}').join(String(r)); }

/**
 * Аркуш має щонайменше rows рядків і cols колонок. Нова вкладка в Таблицях: 1000 × 26 (A…Z); запис за межами
 * падає з помилкою «outside the dimensions of the sheet», тож спершу додаємо рядки й колонки.
 */
function ensureSize_(sh, rows, cols) {
  const r = sh.getMaxRows(), c = sh.getMaxColumns();
  if (rows && r < rows) sh.insertRowsAfter(r, rows - r);
  if (cols && c < cols) sh.insertColumnsAfter(c, cols - c);
  return sh;
}

/** Аркуш: знайти або створити (позиція: в кінці). */
function sheetOf_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function todayKey_(ss) { return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'); }
function todayString_() { return todayKey_(SpreadsheetApp.getActive()); }

function round2_(x) { return Math.round(x * 100) / 100; }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function addDays_(key, n) {
  const d = new Date(Date.UTC(Number(key.slice(0, 4)), Number(key.slice(5, 7)) - 1, Number(key.slice(8, 10))) + n * 86400000);
  return d.toISOString().slice(0, 10);
}

function dayDiff_(a, b) {
  const t = k => Date.UTC(Number(k.slice(0, 4)), Number(k.slice(5, 7)) - 1, Number(k.slice(8, 10)));
  return Math.round((t(a) - t(b)) / 86400000);
}

function lastDayOf_(ym) { return addDays_(addDays_(ym + '-28', 4).slice(0, 7) + '-01', -1); }
function prevYm_(ym) { return addDays_(ym + '-01', -1).slice(0, 7); }
function dowOf_(k) { return new Date(Date.UTC(+k.slice(0, 4), +k.slice(5, 7) - 1, +k.slice(8, 10))).getUTCDay(); }
function dd_(k) { return k.slice(8, 10) + '.' + k.slice(5, 7); }

/** Дата з клітинки → 'yyyy-MM-dd' (рядок як є). */
function dateKey_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v === null || v === undefined ? '' : v);
}

function money0_(v) {
  const n = Math.round(Number(v) || 0), s = String(Math.abs(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  return (n < 0 ? '−' : '') + s + ' ₴';
}

function money_(x) {
  const s = Math.abs(x).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ' ').replace('.', ',');
  return (x < 0 ? '−' : '') + s + ' ₴';
}

const MONTHS_UA = ['січень', 'лютий', 'березень', 'квітень', 'травень', 'червень', 'липень', 'серпень', 'вересень', 'жовтень', 'листопад', 'грудень'];
const WEEKDAYS = ['нд', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
function ymName_(ym) { return MONTHS_UA[Number(ym.slice(5, 7)) - 1] + ' ' + ym.slice(0, 4); }

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function link_(url, text) { return '<a href="' + esc_(url) + '" target="_blank">' + esc_(text) + '</a>'; }

function jsonProp_(dp, k) { try { return JSON.parse(dp.getProperty(k) || 'null'); } catch (e) { return null; } }

function hasTrigger_(fn) { return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === fn); }

function removeTriggers_(fn) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === fn).forEach(t => ScriptApp.deleteTrigger(t));
}

/** «Пульс» розкладу (для самоперевірки): коли востаннє відпрацював і чи без помилки. */
function beat_(name, ok, msg) {
  try {
    PropertiesService.getDocumentProperties().setProperty('hb_' + name,
      JSON.stringify({ t: new Date().toISOString(), ok: !!ok, msg: ok ? '' : String(msg || '').slice(0, 300) }));
  } catch (e) { /* пульс не має ламати основну роботу */ }
}
function readBeat_(name) { return jsonProp_(PropertiesService.getDocumentProperties(), 'hb_' + name); }

/** Повідомлення в маленькому вікні (у тестах: у calls.dialogs). Без інтерфейсу (розклад): у журнал скрипта. */
function show_(title, lines) {
  const body = lines.filter(Boolean).map(l => '<p>' + l + '</p>').join('');
  try {
    const html = HtmlService.createHtmlOutput('<div style="font:14px/1.5 Arial,sans-serif">' + body + '</div>')
      .setWidth(560).setHeight(Math.min(600, 110 + 34 * lines.length));
    SpreadsheetApp.getUi().showModalDialog(html, title);
  } catch (e) {
    Logger.log(title + '\n' + body.replace(/<[^>]+>/g, ' '));
  }
}

/** Значення з аркуша «Головне» за підписом у колонці A (стійко до зсуву рядків). null: немає аркуша чи рядка. */
function mainByLabel_(ss) {
  const s = ss.getSheetByName(MAIN_SHEET);
  if (!s) return () => null;
  const v = s.getRange(1, 1, Math.min(Math.max(s.getLastRow(), 1), 120), 2).getValues();
  return label => {
    const row = v.find(r => String(r[0]) === label);
    return row && row[1] !== '' && row[1] !== null ? row[1] : null;
  };
}
