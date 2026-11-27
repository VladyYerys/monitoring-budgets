/** 37_Debts.gs: Аркуш «Борги» (модуль «🤝 Борги й розстрочки»): хто кому винен і щомісячні платежі як плани в журналі. */

/*
 * Рядок на борг: назва, напрям («я винен» / «мені винні»), сума, платіж на місяць, день платежу, перший платіж, хто, категорія.
 * Сплачено: факти журналу з коментарем «[Борг] Назва» або «[Борг] Назва: мм.рррр» (так їх пише скрипт; можна й вручну).
 * Скрипт ставить у журнал плани «🔁 план регулярний» з коментарем «[Борг] Назва: мм.рррр» до кінця наступного місяця,
 * поки є залишок; змінили суму чи день: плани оновлюються; борг закрито чи модуль вимкнено: майбутні плани зникають.
 * План настав: поставте «✅ факт», і він піде в «Сплачено». Ваші власні рядки скрипт не чіпає.
 */

const DEBT_HEAD = 4, DEBT_FIRST = 5, DEBT_LAST = 24;
const DEBT_TAG = '[Борг] ';
const DEBT_OWE = 'я винен', DEBT_LENT = 'мені винні';
const DEBT_CAT = { 'я винен': 'Повернення боргу', 'мені винні': 'Мені повернули борг' };
/** Підказка в рядку 3: як додати борг, на прикладах (рядків-прикладів тут немає: кожен рядок рахується в залишках і планах). */
const DEBT_HINT = 'Як додати: рядок на кожен борг. Приклади: «Кредит у банку» · я винен · 30 000 · 2 500 · 15 · дата першого платежу; ' +
  '«Розстрочка за телефон» · я винен · 12 000 · 1 000 · 3; «Позика Олегу» · мені винні · 5 000 · 1 000 · 20. Без суми боргу (підписки, оренда): аркуш «Регулярні».';
const DEBT_COLS = ['Назва (кому або хто; без * і ?)', 'Напрям', 'Сума боргу, ₴', 'Платіж на місяць, ₴', 'День платежу (1–28)', 'Перший платіж (дата)',
  'Хто', 'Категорія (порожньо: за напрямом)', 'Сплачено, ₴', 'Залишок, ₴', 'Лишилось платежів'];

function buildDebts_(ss) {
  const sh = ensureSize_(sheetOf_(ss, DEBTS_SHEET), DEBT_LAST + 5, DEBT_COLS.length);
  const O = J_RANGE(JC.fact), D = J_RANGE(JC.note);
  sh.getRange(1, 1).setValue('🤝 Борги й розстрочки').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Платежі стають планами в журналі (коментар «' + DEBT_TAG + 'Назва: мм.рррр»). План настав: «✅ факт», і він піде в «Сплачено». ' +
    'Сплатили інакше: коментар «' + DEBT_TAG + 'Назва» у своєму рядку.').setFontColor('#5f6368');
  sh.getRange(3, 1).setValue(DEBT_HINT).setFontColor('#1a73e8').setFontWeight('bold');
  sh.getRange(DEBT_HEAD, 1, 1, DEBT_COLS.length).setValues([DEBT_COLS]).setFontWeight('bold').setWrap(true);
  const f = [];
  for (let r = DEBT_FIRST; r <= DEBT_LAST; r++) {
    f.push(['=IF($A' + r + '="","",SUMIFS(' + O + ',' + D + ',"' + DEBT_TAG + '"&$A' + r + ')+SUMIFS(' + O + ',' + D + ',"' + DEBT_TAG + '"&$A' + r + '&":*"))',
      '=IF($A' + r + '="","",MAX(0,N($C' + r + ')-N($I' + r + ')))',
      '=IF(OR($A' + r + '="",N($D' + r + ')<=0,N($J' + r + ')<=0),"",ROUNDUP($J' + r + '/$D' + r + ',0))']);
  }
  setF_(sh.getRange(DEBT_FIRST, 9, f.length, 3), f);
  const n = DEBT_LAST - DEBT_FIRST + 1, rng = c => sh.getRange(c + DEBT_FIRST + ':' + c + DEBT_LAST);
  rng('B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([DEBT_OWE, DEBT_LENT]).setAllowInvalid(false).build());
  rng('E').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(1, 28).setAllowInvalid(false).build());
  const set = ss.getSheetByName(SETTINGS_SHEET), dict = ss.getSheetByName(DICT_SHEET);
  if (set) rng('G').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(set.getRange('B' + SET_MEMBERS_FIRST + ':B' + SET_SHARED_ROW)).setAllowInvalid(true).build());
  if (dict) rng('H').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(dict.getRange('A' + D_FIRST + ':A' + D_LAST)).setAllowInvalid(true).build());
  rng('F').setNumberFormat('dd.mm.yyyy');
  sh.getRange(DEBT_FIRST, 3, n, 2).setNumberFormat('#,##0');
  sh.getRange(DEBT_FIRST, 9, n, 2).setNumberFormat('#,##0');
  sh.setFrozenRows(DEBT_HEAD);
  return sh;
}

/** Борги з аркуша: [{ name, dir, total, monthly, day, from (ключ дати або ''), who, cat }]. */
function debtsList_(ss) {
  const sh = ss.getSheetByName(DEBTS_SHEET), tz = ss.getSpreadsheetTimeZone();
  if (!sh) return [];
  return sh.getRange(DEBT_FIRST, 1, DEBT_LAST - DEBT_FIRST + 1, 11).getValues().map(r => ({
    name: String(r[0]).trim(), dir: String(r[1]).trim() === DEBT_LENT ? DEBT_LENT : DEBT_OWE, total: Number(r[2]) || 0, monthly: Number(r[3]) || 0,
    day: Math.max(1, Math.min(28, Math.round(Number(r[4]) || 1))), from: dateKey_(r[5], tz) || '', who: String(r[6] || '').trim(),
    cat: String(r[7] || '').trim(), paid: Number(r[8]) || 0, left: r[9] === '' ? null : Number(r[9]) || 0,
  })).filter(d => d.name && d.total > 0).map(d => Object.assign(d, { cat: d.cat || DEBT_CAT[d.dir] }));
}

/** Чиста функція: сплачено за боргом (факти журналу з тегом «[Борг] Назва» або «[Борг] Назва: …»). */
function debtPaid_(rows, name) {
  const a = DEBT_TAG + name, b = a + ':';
  return round2_(rows.filter(o => Number(o.m) === 0 && !isPlan_(o.e) && o.e !== STATUS_DRAFT && (o.d === a || o.d.indexOf(b) === 0))
    .reduce((s, o) => s + (Number(o.o) || Number(o.b) || 0), 0));
}

/**
 * Чиста функція: які платежі за боргами мають стояти в журналі як плани від today до horizon.
 * → [{ key (коментар), date, amount, cat, who }]
 */
function debtWanted_(debts, rows, today, horizon) {
  const out = [];
  debts.forEach(d => {
    if (!(d.monthly > 0)) return;
    let left = round2_(d.total - debtPaid_(rows, d.name));
    let ym = (d.from && d.from > today ? d.from : today).slice(0, 7);
    for (let k = 0; k < 24 && left > 0.005; k++, ym = addDays_(lastDayOf_(ym), 1).slice(0, 7)) {
      const last = Number(lastDayOf_(ym).slice(8)), date = ym + '-' + pad2_(Math.min(d.day, last));
      if (date > horizon) break;
      const key = DEBT_TAG + d.name + ': ' + ym.slice(5) + '.' + ym.slice(0, 4);
      if (rows.some(o => (o.d === key || o.d.indexOf(key + ' · ') === 0) && !isPlan_(o.e) && (Number(o.m) === 0 || o.e === STATUS_DRAFT))) continue;   // сплачено (чи банк закрив план: чернетка)
      if (date < today || (d.from && date < d.from)) continue;
      const amount = round2_(Math.min(d.monthly, left));
      left = round2_(left - amount);
      out.push({ key: key, date: date, amount: amount, cat: d.cat, who: d.who });
    }
  });
  return out;
}

/** Плани боргів у журналі (додати, оновити, прибрати зайві майбутні). Модуль вимкнено: лише прибрати. */
function debtsPlansSync_(ss) {
  const st = { added: 0, updated: 0, removed: 0 };
  if (!ss.getSheetByName(JOURNAL) || !ss.getSheetByName(DEBTS_SHEET)) return st;
  const today = todayKey_(ss), rows = readJournal_(ss);
  const want = modOn_('debts') ? debtWanted_(debtsList_(ss), rows, today, lastDayOf_(addDays_(lastDayOf_(today.slice(0, 7)), 1).slice(0, 7))) : [];
  const r = tagPlansSync_(ss, DEBT_TAG, want, rows);
  PropertiesService.getDocumentProperties().setProperties({ debt_day: today, debt_t: String(Date.now()) });
  return r;
}

/** З mbsTick: раз на день або після правок. */
function debtsIfDue_(ss) {
  if (!modOn_('debts') || !ss.getSheetByName(DEBTS_SHEET)) return null;
  const dp = PropertiesService.getDocumentProperties();
  if (dp.getProperty('debt_day') === todayKey_(ss) && Number(dp.getProperty('lastEdit') || 0) < Number(dp.getProperty('debt_t') || 0)) return null;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return null;
  try { return debtsPlansSync_(ss); } finally { lock.releaseLock(); }
}

/** Ранковий підсумок: залишки боргів і нагадування за день і в день платежу. */
function debtsMorning_(m) {
  if (!modOn_('debts')) return;
  const ds = debtsList_(m.ss);
  if (!ds.length) return;
  const left = dir => ds.filter(d => d.dir === dir).reduce((s, d) => s + Math.max(0, d.left === null ? d.total - d.paid : d.left), 0);
  const owe = left(DEBT_OWE), lent = left(DEBT_LENT);
  if (owe > 0.005 || lent > 0.005) m.lines.push('🤝 Борги: ' + [owe > 0.005 ? 'я винен ' + money0_(owe) : '', lent > 0.005 ? 'мені винні ' + money0_(lent) : ''].filter(Boolean).join(' · '));
  m.rows.filter(o => isPlan_(o.e) && o.d.indexOf(DEBT_TAG) === 0).forEach(o => {
    const d = dayDiff_(o.key, m.today);
    if (d === 1 || d === 0) m.alerts.push('🤝 <b>' + esc_(o.d.slice(DEBT_TAG.length)) + '</b>: ' + money0_(o.b) + ' (' + esc_(o.c) + '), ' + (d === 0 ? 'сьогодні' : 'завтра') + '.');
  });
}
