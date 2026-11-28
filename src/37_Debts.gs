/** 37_Debts.gs: Аркуш «Борги» (модуль «🤝 Борги й розстрочки»): хто кому винен і щомісячні платежі як плани в журналі. */

/*
 * Рядок на борг: назва, напрям («я винен» / «мені винні»), сума, платіж на місяць, день платежу, перший платіж, хто, категорія.
 * Сплачено: факти журналу з коментарем «[Борг] Назва» або «[Борг] Назва: мм.рррр» (так їх пише скрипт; можна й вручну).
 * Скрипт ставить у журнал плани «🔁 план регулярний» з коментарем «[Борг] Назва: мм.рррр» до кінця наступного місяця,
 * поки є залишок; змінили суму чи день: плани оновлюються; борг закрито чи модуль вимкнено: майбутні плани зникають.
 * План настав: поставте «✅ факт», і він піде в «Сплачено». Ваші власні рядки скрипт не чіпає.
 * Версія 17 (покупки частинами й розстрочки):
 *   L «Вид» (покупка частинами, розстрочка, кредит готівкою, кредит, позика, інше): для частин, розстрочок і кредитів категорія платежу
 *     типово «Кредит / розстрочка» (платіж щомісяця: витрата місяця, «мої» падають поступово, як гроші на картці);
 *   M «Назва в банку»: магазин з опису Mono («Платіж TECHNOMIX - monomarket» → TECHNOMIX), за нею імпорт знаходить цей рядок;
 *   N «Закрито»: дата повного погашення (пише імпорт з «Повне погашення …» або ви); залишок 0, планів більше немає;
 *   O «Прогрес»: смужка «▰▰▱▱ 3 з 25» (сплачено з суми боргу); без суми: «⚠️ впишіть суму (C)».
 * Новий рядок з банку: перший платіж «Платіж МАГАЗИН - monomarket» (аванс = щомісячний платіж) чи перше «Погашення …» без рядка:
 *   назва й «Назва в банку» = магазин, платіж = сума, день і перший платіж = дата, сума боргу порожня (впишіть, тоді будуть плани).
 * Ліміти (рядки 27–29): загальний ліміт (ви) і «доступно в Mono» (ви, з застосунку «Кредити»); зайнято = залишки рядків цього виду;
 *   вільно = загальний − зайнято. Не сходиться з Mono: таблиця підказує загальний = доступно + зайнято; «так» у G: записати його в B.
 */

const DEBT_HEAD = 4, DEBT_FIRST = 5, DEBT_LAST = 24;
const DEBT_TAG = '[Борг] ';
const DEBT_OWE = 'я винен', DEBT_LENT = 'мені винні';
const DEBT_CAT = { 'я винен': 'Повернення боргу', 'мені винні': 'Мені повернули борг' };
/** Підказка в рядку 3: як додати борг, на прикладах (рядків-прикладів тут немає: кожен рядок рахується в залишках і планах). */
const DEBT_HINT = 'Як додати: рядок на кожен борг. Приклади: «Кредит у банку» · я винен · 30 000 · 2 500 · 15 · дата першого платежу; ' +
  '«Розстрочка за телефон» · я винен · 12 000 · 1 000 · 3; «Позика Олегу» · мені винні · 5 000 · 1 000 · 20. Без суми боргу (підписки, оренда): аркуш «Регулярні».';
const DEBT_COLS = ['Назва (кому або хто; без * і ?)', 'Напрям', 'Сума боргу, ₴', 'Платіж на місяць, ₴', 'День платежу (1–28)', 'Перший платіж (дата)',
  'Хто', 'Категорія (порожньо: за видом чи напрямом)', 'Сплачено, ₴', 'Залишок, ₴', 'Лишилось платежів',
  'Вид', 'Назва в банку (магазин з опису Mono)', 'Закрито (дата повного погашення)', 'Прогрес'];
const DEBT_KINDS = ['покупка частинами', 'розстрочка', 'кредит готівкою', 'кредит', 'позика', 'інше'];
const DEBT_PARTS = 'покупка частинами', DEBT_INST = 'розстрочка', CAT_PARTS = 'Кредит / розстрочка';
const DEBT_BANK_KINDS = [DEBT_PARTS, DEBT_INST, 'кредит готівкою', 'кредит'];       // платіж банку: типово «Кредит / розстрочка»
/** Ліміти Mono («Кредити»): рядок, підпис, які види займають ліміт. */
const DEBT_LIM_HEAD = 27, DEBT_LIMITS = [
  { row: 28, label: 'Розстрочка й покупка частинами', kinds: [DEBT_PARTS, DEBT_INST] },
  { row: 29, label: 'Кредит готівкою', kinds: ['кредит готівкою'] }];
const DEBT_LIM_COLS = ['💳 Ліміти Mono (застосунок → «Кредити»)', 'Загальний ліміт, ₴', 'Доступно в Mono зараз, ₴', 'Зайнято (залишки рядків цього виду), ₴',
  'Вільно за таблицею, ₴', 'Перевірка', 'Підтвердити (так)'];
const DEBT_HINT2 = 'Покупки частинами й розстрочки: «Вид», «Назва в банку» (магазин з опису Mono); перший платіж з банку додає рядок сам, ' +
  'впишіть лише суму покупки (C). Ліміти: рядки ' + DEBT_LIMITS[0].row + '–' + DEBT_LIMITS[1].row + '.';

function buildDebts_(ss) {
  const sh = ensureSize_(sheetOf_(ss, DEBTS_SHEET), DEBT_LAST + 7, DEBT_COLS.length);
  sh.getRange(1, 1).setValue('🤝 Борги й розстрочки').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Платежі стають планами в журналі (коментар «' + DEBT_TAG + 'Назва: мм.рррр»). План настав: «✅ факт», і він піде в «Сплачено». ' +
    'Сплатили інакше: коментар «' + DEBT_TAG + 'Назва» у своєму рядку.').setFontColor('#5f6368');
  sh.getRange(3, 1).setValue(DEBT_HINT).setFontColor('#1a73e8').setFontWeight('bold');
  debtsColumns_(ss, sh);
  debtsLimits_(sh);
  sh.setFrozenRows(DEBT_HEAD);
  return sh;
}

/** Формули рядка r: I сплачено, J залишок (закрито: 0), K лишилось платежів; O прогрес. */
function debtRowF_(r) {
  const O = J_RANGE(JC.fact), D = J_RANGE(JC.note);
  const part = 'ROUND(10*MIN(1,N($I' + r + ')/N($C' + r + ')),0)', total = 'ROUNDUP(N($C' + r + ')/N($D' + r + '),0)';
  return {
    ijk: ['=IF($A' + r + '="","",SUMIFS(' + O + ',' + D + ',"' + DEBT_TAG + '"&$A' + r + ')+SUMIFS(' + O + ',' + D + ',"' + DEBT_TAG + '"&$A' + r + '&":*"))',
      '=IF($A' + r + '="","",IF($N' + r + '<>"",0,MAX(0,N($C' + r + ')-N($I' + r + '))))',
      '=IF(OR($A' + r + '="",N($D' + r + ')<=0,N($J' + r + ')<=0),"",ROUNDUP($J' + r + '/$D' + r + ',0))'],
    o: '=IF($A' + r + '="","",IF($N' + r + '<>"","✓ закрито "&TEXT($N' + r + ',"dd.mm.yyyy"),IF(N($C' + r + ')<=0,"⚠️ впишіть суму (C)",IF(N($D' + r + ')<=0,"",' +
      'REPT("▰",' + part + ')&REPT("▱",10-' + part + ')&" "&MIN(' + total + ',ROUND(N($I' + r + ')/N($D' + r + '),0))&" з "&' + total + '))))',
  };
}

/** Заголовки, формули, списки й формати колонок A…O (нова таблиця й крок 17; ваші значення A…H, L…N не чіпає). */
function debtsColumns_(ss, sh) {
  ensureSize_(sh, DEBT_LIMITS[DEBT_LIMITS.length - 1].row + 2, DEBT_COLS.length);
  sh.getRange(DEBT_HEAD, 1, 1, DEBT_COLS.length).setValues([DEBT_COLS]).setFontWeight('bold').setWrap(true);
  const ijk = [], o = [];
  for (let r = DEBT_FIRST; r <= DEBT_LAST; r++) { const f = debtRowF_(r); ijk.push(f.ijk); o.push([f.o]); }
  setF_(sh.getRange(DEBT_FIRST, 9, ijk.length, 3), ijk);
  setF_(sh.getRange(DEBT_FIRST, 15, o.length, 1), o);
  const rng = c => sh.getRange(c + DEBT_FIRST + ':' + c + DEBT_LAST);
  rng('B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([DEBT_OWE, DEBT_LENT]).setAllowInvalid(false).build());
  rng('E').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(1, 28).setAllowInvalid(false).build());
  rng('L').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(DEBT_KINDS).setAllowInvalid(true).build());
  const set = ss.getSheetByName(SETTINGS_SHEET), dict = ss.getSheetByName(DICT_SHEET);
  if (set) rng('G').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(set.getRange('B' + SET_MEMBERS_FIRST + ':B' + SET_SHARED_ROW)).setAllowInvalid(true).build());
  if (dict) rng('H').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(dict.getRange('A' + D_FIRST + ':A' + D_LAST)).setAllowInvalid(true).build());
  rng('F').setNumberFormat('dd.mm.yyyy'); rng('N').setNumberFormat('dd.mm.yyyy');
  sh.getRange(DEBT_FIRST, 3, DEBT_LAST - DEBT_FIRST + 1, 2).setNumberFormat('#,##0.00');
  sh.getRange(DEBT_FIRST, 9, DEBT_LAST - DEBT_FIRST + 1, 2).setNumberFormat('#,##0');
  rng('O').setFontColor('#5b4ea0');
  sh.getRange(3, 12).setValue(DEBT_HINT2).setFontColor('#5b4ea0').setWrap(false);
}

/** Блок лімітів (рядки 27–29, A…G): ваші B і C не перезаписує. */
function debtsLimits_(sh) {
  const J = c => '$' + c + '$' + DEBT_FIRST + ':$' + c + '$' + DEBT_LAST;
  sh.getRange(DEBT_LIM_HEAD, 1, 1, DEBT_LIM_COLS.length).setValues([DEBT_LIM_COLS]).setFontWeight('bold').setWrap(true).setBackground('#f3e8fd');
  DEBT_LIMITS.forEach(l => {
    const r = l.row, used = l.kinds.map(k => 'SUMIFS(' + J('J') + ',' + J('L') + ',"' + k + '")').join('+');
    sh.getRange(r, 1).setValue(l.label);
    setF_(sh.getRange(r, 4, 1, 3), [['=' + used, '=IF($B' + r + '="","",N($B' + r + ')-$D' + r + ')',
      '=IF($C' + r + '="",IF($B' + r + '="","Впишіть загальний ліміт (B) і доступно в Mono (C)","Впишіть «доступно в Mono» (C), щоб звірити"),' +
      'IF($B' + r + '="","Загальний ≈ "&TEXT(N($C' + r + ')+$D' + r + ',"#,##0")&"? «так» у G запише його в B",' +
      'IF(ABS($E' + r + '-N($C' + r + '))<1,"✓ сходиться з Mono","⚠️ не сходиться на "&TEXT($E' + r + '-N($C' + r + '),"#,##0")&": якщо в Mono правда, загальний ≈ "&' +
      'TEXT(N($C' + r + ')+$D' + r + ',"#,##0")&". «так» у G запише його в B (або перевірте, чи всі покупки є вище)")))']]);
    sh.getRange(r, 7).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['так']).setAllowInvalid(true).build());
  });
  const first = DEBT_LIMITS[0].row, n = DEBT_LIMITS.length;
  sh.getRange(first, 2, n, 4).setNumberFormat('#,##0');
  sh.getRange(first, 2, n, 2).setBackground('#fff2cc');
  sh.getRange(first, 6, n, 1).setWrap(true);
}

/** Чиста функція: типова категорія платежу. */
function debtCat_(cat, kind, dir) { return cat || (DEBT_BANK_KINDS.indexOf(kind) >= 0 ? CAT_PARTS : DEBT_CAT[dir]); }

/** Усі рядки «Боргів» з назвою (і без суми, і закриті): [{ row, name, dir, total, monthly, day, from, who, cat, paid, left, kind, bank, closed }]. */
function debtsAll_(ss) {
  const sh = ss.getSheetByName(DEBTS_SHEET), tz = ss.getSpreadsheetTimeZone();
  if (!sh) return [];
  const n = Math.min(sh.getMaxColumns(), DEBT_COLS.length);
  return sh.getRange(DEBT_FIRST, 1, DEBT_LAST - DEBT_FIRST + 1, n).getValues().map((r, i) => {
    const dir = String(r[1]).trim() === DEBT_LENT ? DEBT_LENT : DEBT_OWE, kind = String(r[11] || '').trim();
    return { row: DEBT_FIRST + i, name: String(r[0]).trim(), dir: dir, total: Number(r[2]) || 0, monthly: Number(r[3]) || 0,
      day: Math.max(1, Math.min(28, Math.round(Number(r[4]) || 1))), from: dateKey_(r[5], tz) || '', who: String(r[6] || '').trim(),
      cat: debtCat_(String(r[7] || '').trim(), kind, dir), paid: Number(r[8]) || 0, left: r[9] === '' || r[9] === undefined ? null : Number(r[9]) || 0,
      kind: kind, bank: String(r[12] || '').trim(), closed: r[13] === '' || r[13] === undefined || r[13] === null ? '' : (dateKey_(r[13], tz) || String(r[13])) };
  }).filter(d => d.name);
}

/** Стан для «⚡ Швидкого підключення»: рядків, частинами, без суми, ліміт. */
function debtsState_(ss) {
  const all = debtsAll_(ss), open = all.filter(d => !d.closed), sh = ss.getSheetByName(DEBTS_SHEET);
  const parts = open.filter(d => d.kind === DEBT_PARTS || d.kind === DEBT_INST).length, noSum = open.filter(d => !(d.total > 0)).length;
  let lim = '';
  if (sh && String(sh.getRange(DEBT_LIM_HEAD, 1).getValue()) === DEBT_LIM_COLS[0]) {
    const b = sh.getRange(DEBT_LIMITS[0].row, 2).getValue(), f = String(sh.getRange(DEBT_LIMITS[0].row, 6).getValue());
    lim = b === '' ? 'ліміт частинами не вписано' : f.charAt(0) === '⚠' ? 'ліміт не сходиться з Mono' : 'ліміт вписано';
  }
  return ['відкритих: ' + open.length, parts ? 'частинами й розстрочок: ' + parts : '', noSum ? '⚠️ без суми боргу: ' + noSum : '', lim].filter(Boolean).join(', ');
}

/** Борги з сумою, не закриті: для планів і підсумків. */
function debtsList_(ss) { return debtsAll_(ss).filter(d => d.total > 0 && !d.closed); }

/** Чиста функція: залишок боргів цих видів (для ліміту): сума − сплачено з журналу, закриті й без суми: 0. */
function debtsUsed_(debts, rows, kinds) {
  return round2_(debts.filter(d => kinds.indexOf(d.kind) >= 0 && !d.closed && d.total > 0)
    .reduce((s, d) => s + Math.max(0, d.total - debtPaid_(rows, d.name)), 0));
}

/** Правка «так» у G рядка ліміту? */
function isDebtLimitEdit_(range) {
  const sh = range.getSheet();
  return sh.getName() === DEBTS_SHEET && range.getColumn() <= 7 && range.getLastColumn() >= 7 &&
    DEBT_LIMITS.some(l => range.getRow() <= l.row && range.getLastRow() >= l.row);
}

/** «так» у G: загальний ліміт (B) = доступно в Mono (C) + зайнято за «Боргами»; G очищується. → [повідомлення] */
function debtLimitConfirm_(ss) {
  const sh = ss.getSheetByName(DEBTS_SHEET), out = [];
  if (!sh) return out;
  const debts = debtsAll_(ss), rows = ss.getSheetByName(JOURNAL) ? readJournal_(ss) : [];
  DEBT_LIMITS.forEach(l => {
    if (String(sh.getRange(l.row, 7).getValue()).trim().toLowerCase() !== 'так') return;
    const avail = sh.getRange(l.row, 3).getValue();
    sh.getRange(l.row, 7).setValue('');
    if (avail === '' || avail === null || isNaN(Number(avail))) { out.push(l.label + ': спершу впишіть «доступно в Mono» (C)'); return; }
    const total = round2_(Number(avail) + debtsUsed_(debts, rows, l.kinds));
    sh.getRange(l.row, 2).setValue(total);
    out.push(l.label + ': загальний ліміт ' + money0_(total));
  });
  return out;
}

/** Після імпорту з банку: нові рядки (перший платіж частинами чи погашення без рядка) і «Закрито» з «Повне погашення». → { added, closed, full } */
function debtsFromBank_(ss, plan) {
  const sh = ss.getSheetByName(DEBTS_SHEET), tz = ss.getSpreadsheetTimeZone(), out = { added: [], closed: [], full: false };
  if (!sh || (!plan.newDebts.length && !plan.closeDebts.length)) return out;
  ensureSize_(sh, DEBT_LAST, DEBT_COLS.length);
  const D = k => Utilities.parseDate(k, tz, 'yyyy-MM-dd');
  const names = sh.getRange(DEBT_FIRST, 1, DEBT_LAST - DEBT_FIRST + 1, 1).getValues().map(r => String(r[0]).trim());
  plan.newDebts.forEach(d => {
    if (names.indexOf(d.name) >= 0) return;
    const i = names.findIndex(n => !n);
    if (i < 0) { out.full = true; return; }
    sh.getRange(DEBT_FIRST + i, 1, 1, 8).setValues([[d.name, DEBT_OWE, '', d.monthly, d.day, D(d.from), d.who || '', '']]);
    sh.getRange(DEBT_FIRST + i, 12, 1, 2).setValues([[d.kind, d.bank]]);
    names[i] = d.name; out.added.push(d.name);
  });
  plan.closeDebts.forEach(c => {
    const i = c.row ? c.row - DEBT_FIRST : names.indexOf(c.name);
    if (i < 0 || i > DEBT_LAST - DEBT_FIRST) return;
    const cell = sh.getRange(DEBT_FIRST + i, 14);
    if (cell.getValue() !== '' && cell.getValue() !== null) return;
    cell.setValue(D(c.key));
    out.closed.push(c.name);
  });
  if (out.added.length || out.closed.length) { try { debtsPlansSync_(ss); } catch (e) { /* плани оновить розклад */ } }
  return out;
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
    if (!(d.monthly > 0) || d.closed) return;
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
