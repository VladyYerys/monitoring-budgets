/** 36_Fop.gs: «ФОП»: групи 1, 2, 3; ставки в таблиці; податки кожного ФОП окремо; ліміт доходу; плани податків у журналі. */

/*
 * Ставки, строки й ліміти: у таблиці «Ставки» (рядок на рік). Змінився закон: поправте чи додайте рядок, код не чіпаєте.
 * ФОП у сім'ї (до 5): ім'я як у колонці «Хто» журналу, група, своя ставка ЄП (рішення місцевої ради), чи платить ВЗ і ЄСВ.
 * Дохід ФОП: факти журналу з категорією «Дохід ФОП» цього учасника (на будь-якому рахунку: ФОП чи картці).
 * Якщо ФОП один, «Хто» не потрібне. Податки по кварталах рахують формули; «✅» або «Сплачено вручну» закривають квартал.
 * Плани в журналі («🔁 план регулярний», коментар «[ФОП] …»): ЄП і ВЗ 1 і 2 групи щомісяця до 20-го, ЄСВ до 20-го числа після
 * кварталу, ЄП і ВЗ 3 групи: через 50 днів після кварталу. Сплатили (факт тієї ж категорії поруч зі строком): план зникає.
 * Це підказка, не податкова консультація: строки й ставки перевіряйте в ДПС.
 */

const FOP_RATE_HEAD = 5, FOP_RATE_FIRST = 6, FOP_RATE_LAST = 10;
const FOP_HEAD = 13, FOP_FIRST = 14, FOP_LAST = 18, FOP_YEAR = 'B20', FOP_Q_HEAD = 22, FOP_Q_FIRST = 23;
const FOP_TAG = '[ФОП] ';
const FOP_RATES_2026 = [2026, 8647, 3328, 332.8, 1729.4, 5, 1902.34, 864.7, 1, 1444049, 7211598, 10091049,
  'Закон про Держбюджет-2026 і ПКУ (ДПС). ЄП 1 і 2 групи: максимум, ставку вашої громади впишіть у «ФОП у сім\'ї».'];
const FOP_RATE_COLS = ['Рік', 'Мінімальна зарплата', 'Прожитковий мінімум', 'ЄП 1 група, ₴/міс (макс.)', 'ЄП 2 група, ₴/міс (макс.)', 'ЄП 3 група, %',
  'ЄСВ, ₴/міс', 'ВЗ 1 і 2 група, ₴/міс', 'ВЗ 3 група, %', 'Ліміт доходу 1 група', 'Ліміт 2 група', 'Ліміт 3 група', 'Джерело'];
const QN = ['I', 'II', 'III', 'IV'];

function fopRate_(col) {
  const R = c => '$' + c + '$' + FOP_RATE_FIRST + ':$' + c + '$' + FOP_RATE_LAST;
  return 'IFERROR(INDEX(' + R(col) + ',MATCH($B$20,' + R('A') + ',0)),INDEX(' + R(col) + ',MATCH(MAX(' + R('A') + '),' + R('A') + ',0)))';
}

function buildFop_(ss) {
  const sh = ensureSize_(sheetOf_(ss, FOP_SHEET), FOP_Q_FIRST + 20, 14), tz = ss.getSpreadsheetTimeZone();
  const O = J_RANGE(JC.fact), C = J_RANGE(JC.cat), G = J_RANGE(JC.who), A = J_RANGE(JC.date);
  sh.getRange(1, 1).setValue('🧾 ФОП').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Податки кожного ФОП у сім\'ї окремо. Ставки міняйте рядком у таблиці «Ставки». Це підказка, не податкова консультація: ' +
    'строки й ставки перевіряйте в ДПС.').setFontColor('#5f6368');
  sh.getRange(4, 1).setValue('Ставки (рядок на рік)').setFontWeight('bold');
  sh.getRange(FOP_RATE_HEAD, 1, 1, FOP_RATE_COLS.length).setValues([FOP_RATE_COLS]).setFontWeight('bold').setWrap(true);
  sh.getRange(FOP_RATE_FIRST, 1, 1, FOP_RATES_2026.length).setValues([FOP_RATES_2026]);
  sh.getRange(12, 1).setValue('ФОП у сім\'ї').setFontWeight('bold');
  sh.getRange(FOP_HEAD, 1, 1, 11).setValues([['Ім\'я (як у «Хто»)', 'Група (1, 2, 3)', 'Своя ставка ЄП: ₴/міс (1, 2 група) або % (3 група)', 'ВЗ (так / ні)',
    'ЄСВ (так / ні)', 'Дохід з початку року', 'Ліміт групи', 'Використано ліміту', 'Стан', 'Сплачено податків за рік (журнал)', 'Хто (службове)']]).setFontWeight('bold').setWrap(true);
  const yr = 'DATE($B$20,1,1)', yr1 = 'DATE($B$20+1,1,1)', inYear = A + ',">="&' + yr + ',' + A + ',"<"&' + yr1;
  const fops = [];
  for (let r = FOP_FIRST; r <= FOP_LAST; r++) {
    const paid = Object.keys(CAT_TAX).map(k => 'SUMIFS(' + O + ',' + C + ',"' + CAT_TAX[k] + '",' + G + ',$K' + r + ',' + inYear + ')').join('+');
    fops.push(['=IF($A' + r + '="","",SUMIFS(' + O + ',' + C + ',"' + CAT_FOP_INCOME + '",' + G + ',$K' + r + ',' + inYear + '))',
      '=IF($A' + r + '="","",CHOOSE(MAX(1,MIN(3,IFERROR(VALUE($B' + r + '&""),1))),' + fopRate_('J') + ',' + fopRate_('K') + ',' + fopRate_('L') + '))',
      '=IF(OR($A' + r + '="",N($G' + r + ')=0),"",$F' + r + '/$G' + r + ')',
      '=IF($A' + r + '="","",IF(N($H' + r + ')>=1,"⛔ понад ліміт",IF(N($H' + r + ')>=0.9,"⚠️ понад 90%",IF(N($H' + r + ')>=0.8,"⚠️ понад 80%","OK"))))',
      '=IF($A' + r + '="","",' + paid + ')',
      '=IF(COUNTA($A$' + FOP_FIRST + ':$A$' + FOP_LAST + ')<=1,"<>✱",$A' + r + ')']);
  }
  setF_(sh.getRange(FOP_FIRST, 6, FOP_LAST - FOP_FIRST + 1, 6), fops);
  sh.getRange('B' + FOP_FIRST + ':B' + FOP_LAST).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['1', '2', '3']).setAllowInvalid(false).build());
  ['D', 'E'].forEach(c => sh.getRange(c + FOP_FIRST + ':' + c + FOP_LAST).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['так', 'ні']).setAllowInvalid(false).build()));
  const set = ss.getSheetByName(SETTINGS_SHEET);
  if (set) sh.getRange('A' + FOP_FIRST + ':A' + FOP_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(set.getRange('B' + SET_MEMBERS_FIRST + ':B' + (SET_MEMBERS_FIRST + SET_MEMBERS_N - 1))).setAllowInvalid(true).build());
  sh.getRange('H' + FOP_FIRST + ':H' + FOP_LAST).setNumberFormat('0%');
  sh.getRange('F' + FOP_FIRST + ':G' + FOP_LAST).setNumberFormat('#,##0');
  sh.getRange('J' + FOP_FIRST + ':J' + FOP_LAST).setNumberFormat('#,##0');
  sh.getRange('K' + FOP_FIRST + ':K' + FOP_LAST).setFontColor('#9aa0a6');
  sh.getRange('A20').setValue('Рік:');
  sh.getRange(FOP_YEAR).setValue(Number(Utilities.formatDate(new Date(), tz, 'yyyy')));
  sh.getRange(21, 1).setValue('Квартали').setFontWeight('bold');
  sh.getRange(FOP_Q_HEAD, 1, 1, 14).setValues([['ФОП', 'Квартал', 'Дохід', 'ЄП', 'ВЗ', 'ЄСВ', 'Разом', 'Строк ЄП і ВЗ', 'Строк ЄСВ', 'Сплачено вручну, ₴', '✅',
    'Залишок', 'Початок (службове)', 'Кінець (службове)']]).setFontWeight('bold').setWrap(true);
  const q = [];
  for (let f = 0; f < FOP_LAST - FOP_FIRST + 1; f++) {
    const fr = FOP_FIRST + f, name = '$A$' + fr, grp = 'IFERROR(VALUE($B$' + fr + '&""),0)', own = '$C$' + fr;   // група з меню може бути текстом «3»
    for (let k = 1; k <= 4; k++) {
      const n = FOP_Q_FIRST + f * 4 + k - 1, M = '$M' + n, N = '$N' + n, on = '$A' + n + '=""';
      q.push(['=IF(' + name + '="","",' + name + ')', '="' + QN[k - 1] + ' кв."',
        '=IF(' + on + ',"",SUMIFS(' + O + ',' + C + ',"' + CAT_FOP_INCOME + '",' + G + ',$K$' + fr + ',' + A + ',">="&' + M + ',' + A + ',"<"&' + N + '))',
        '=IF(' + on + ',"",IF(' + grp + '=3,$C' + n + '*IF(' + own + '="",' + fopRate_('F') + ',' + own + ')/100,3*IF(' + own + '="",IF(' + grp + '=1,' + fopRate_('D') + ',' + fopRate_('E') + '),' + own + ')))',
        '=IF(OR(' + on + ',$D$' + fr + '="ні"),0,IF(' + grp + '=3,$C' + n + '*' + fopRate_('I') + '/100,3*' + fopRate_('H') + '))',
        '=IF(OR(' + on + ',$E$' + fr + '="ні"),0,3*' + fopRate_('G') + ')',
        '=IF(' + on + ',"",N($D' + n + ')+N($E' + n + ')+N($F' + n + '))',
        '=IF(' + on + ',"",IF(' + grp + '=3,' + N + '-1+50,DATE(YEAR(' + M + '),MONTH(' + M + ')+2,20)))',
        '=IF(' + on + ',"",DATE(YEAR(' + N + '),MONTH(' + N + '),20))', '', '',
        '=IF(' + on + ',"",IF($K' + n + '="✅",0,MAX(0,N($G' + n + ')-N($J' + n + '))))',
        '=DATE($B$20,' + ((k - 1) * 3 + 1) + ',1)', '=EDATE($M' + n + ',3)']);
    }
  }
  setF_(sh.getRange(FOP_Q_FIRST, 1, q.length, 14), q);
  sh.getRange('K' + FOP_Q_FIRST + ':K' + (FOP_Q_FIRST + q.length - 1)).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['✅']).setAllowInvalid(true).build());
  sh.getRange('C' + FOP_Q_FIRST + ':G' + (FOP_Q_FIRST + q.length - 1)).setNumberFormat('#,##0.00');
  sh.getRange('L' + FOP_Q_FIRST + ':L' + (FOP_Q_FIRST + q.length - 1)).setNumberFormat('#,##0.00');
  sh.getRange('H' + FOP_Q_FIRST + ':I' + (FOP_Q_FIRST + q.length - 1)).setNumberFormat('dd.mm.yyyy');
  sh.getRange('M' + FOP_Q_FIRST + ':N' + (FOP_Q_FIRST + q.length - 1)).setNumberFormat('dd.mm.yyyy').setFontColor('#9aa0a6');
  return sh;
}

/* ---------- ФОП у JS (для планів і Telegram: ті самі правила, що й формули) ---------- */

/** ФОП з аркуша: [{name, group, own, vz, esv, row}] і ставки року. */
function fopList_(ss) {
  const sh = ss.getSheetByName(FOP_SHEET);
  if (!sh) return [];
  return sh.getRange(FOP_FIRST, 1, FOP_LAST - FOP_FIRST + 1, 5).getValues().map((r, i) => ({ row: FOP_FIRST + i, name: String(r[0]).trim(),
    group: Number(r[1]) || 0, own: r[2] === '' || r[2] === null ? null : Number(r[2]), vz: String(r[3]).trim() !== 'ні', esv: String(r[4]).trim() !== 'ні' }))
    .filter(f => f.name && f.group >= 1 && f.group <= 3);
}

function fopRates_(ss, year) {
  const sh = ss.getSheetByName(FOP_SHEET);
  if (!sh) return null;
  const v = sh.getRange(FOP_RATE_FIRST, 1, FOP_RATE_LAST - FOP_RATE_FIRST + 1, 12).getValues().filter(r => Number(r[0]) > 0);
  const row = v.find(r => Number(r[0]) === year) || v.sort((a, b) => b[0] - a[0])[0];
  if (!row) return null;
  return { ep1: Number(row[3]), ep2: Number(row[4]), ep3: Number(row[5]), esv: Number(row[6]), vz12: Number(row[7]), vz3: Number(row[8]),
    limit: [0, Number(row[9]), Number(row[10]), Number(row[11])] };
}

/** Чиста функція: дохід ФОП за період із журналу (факти «Дохід ФОП» цього учасника; якщо ФОП один: будь-чиї). */
function fopIncome_(rows, name, single, from, toExcl) {
  return round2_(rows.filter(o => o.c === CAT_FOP_INCOME && Number(o.m) === 0 && !isPlan_(o.e) && o.e !== STATUS_DRAFT && o.key >= from && o.key < toExcl &&
    (single || o.g === name)).reduce((s, o) => s + (Number(o.o) || Number(o.b) || 0), 0));
}

/**
 * Чиста функція: які податкові платежі мають стояти в журналі як плани від today до horizon.
 * → [{ key (коментар), date, amount, cat, who }]
 */
function fopWanted_(fops, rates, rows, today, horizon, closed) {
  const out = [], single = fops.length <= 1;
  const add = (f, what, period, date, amount, cat) => {
    if (date < today || date > horizon || !(amount > 0.005)) return;
    const key = FOP_TAG + f.name + ': ' + what + ' за ' + period;
    if (closed && closed[f.name + '|' + period]) return;
    const paid = rows.some(o => o.c === cat && Number(o.m) === 0 && !isPlan_(o.e) && (single || o.g === f.name) &&
      o.key >= addDays_(date, -35) && o.key <= addDays_(date, 10) && Number(o.b) >= amount * 0.9 - 0.005);
    if (!paid) out.push({ key: key, date: date, amount: round2_(amount), cat: cat, who: single ? '' : f.name });
  };
  const y0 = Number(today.slice(0, 4));
  fops.forEach(f => {
    for (let y = y0 - 1; y <= y0 + 1; y++) for (let q = 1; q <= 4; q++) {
      const qs = y + '-' + pad2_((q - 1) * 3 + 1) + '-01', qe = addDays_(lastDayOf_(y + '-' + pad2_(q * 3)), 1), period = QN[q - 1] + ' кв. ' + y;
      const esvDue = qe.slice(0, 7) + '-20';
      if (f.esv) add(f, 'ЄСВ', period, esvDue, 3 * rates.esv, CAT_TAX.esv);
      if (f.group === 3) {
        const due = addDays_(qe, 49), inc = fopIncome_(rows, f.name, single, qs, qe);
        add(f, 'ЄП', period, due, inc * (f.own === null ? rates.ep3 : f.own) / 100, CAT_TAX.ep);
        if (f.vz) add(f, 'ВЗ', period, due, inc * rates.vz3 / 100, CAT_TAX.vz);
      } else {
        for (let k = 0; k < 3; k++) {
          const ym = y + '-' + pad2_((q - 1) * 3 + 1 + k), mname = MONTHS_UA[Number(ym.slice(5)) - 1] + ' ' + y;
          add(f, 'ЄП', mname, ym + '-20', f.own === null ? (f.group === 1 ? rates.ep1 : rates.ep2) : f.own, CAT_TAX.ep);
          if (f.vz) add(f, 'ВЗ', mname, ym + '-20', rates.vz12, CAT_TAX.vz);
        }
      }
    }
  });
  return out;
}

/** Закриті вручну квартали («✅»): { 'Ім'я|I кв. 2026': true }. */
function fopClosed_(ss) {
  const sh = ss.getSheetByName(FOP_SHEET), out = {};
  if (!sh) return out;
  const y = Number(sh.getRange(FOP_YEAR).getValue()) || 0;
  sh.getRange(FOP_Q_FIRST, 1, 20, 11).getValues().forEach(r => { if (String(r[0]).trim() && String(r[10]).trim() === '✅') out[String(r[0]).trim() + '|' + String(r[1]).trim() + ' ' + y] = true; });
  return out;
}

/**
 * Плани з тегом у журналі (ФОП, борги): оновлює суми й дати своїх рядків («🔁 план регулярний», коментар «<тег>…»), додає нові,
 * прибирає зайві майбутні. Ваші рядки не чіпає. want: [{ key, date, amount, cat, who }]. → { added, updated, removed }
 */
function tagPlansSync_(ss, tag, want, rows) {
  const sh = ss.getSheetByName(JOURNAL), tz = ss.getSpreadsheetTimeZone(), st = { added: 0, updated: 0, removed: 0 };
  if (!sh) return st;
  const today = todayKey_(ss), byKey = {};
  rows = rows || readJournal_(ss);
  want.forEach(w => { byKey[w.key] = w; });
  const D = k => Utilities.parseDate(k, tz, 'yyyy-MM-dd');
  rows.filter(o => o.e === STATUS_REG && o.d.indexOf(tag) === 0).forEach(o => {
    const w = byKey[o.d];
    if (!w) {
      if (o.key >= today) { sh.getRange(o.row, 1, 1, J_INPUT_N).setValues([['', '', '', '', '', '', '', '']]); sh.getRange(o.row, colNum_(JC.bankId), 1, J_MOVE_N).setValues([emptyBank_().concat(['', ''])]); st.removed++; }
      return;
    }
    delete byKey[o.d];
    if (o.key !== w.date || Math.abs(Number(o.b) - w.amount) >= 0.005 || o.c !== w.cat || o.g !== (w.who || '')) {
      sh.getRange(o.row, 1, 1, 3).setValues([[D(w.date), w.amount, w.cat]]);
      if (o.g !== (w.who || '')) sh.getRange(o.row, 7).setValue(w.who || '');
      st.updated++;
    }
  });
  const add = Object.keys(byKey).map(k => byKey[k]);
  if (add.length) {
    let lastUsed = -1;
    rows.forEach((r, i) => { if (!r.blank || r.bank.some(x => x !== '' && x !== null) || (r.ord !== '' && r.ord !== null) || r.acct) lastUsed = i; });
    const fit = add.slice(0, Math.max(0, rows.length - lastUsed - 1));
    if (fit.length) sh.getRange(J_FIRST + lastUsed + 1, 1, fit.length, J_INPUT_N).setValues(fit.map(w => [D(w.date), w.amount, w.cat, w.key, STATUS_REG, '', w.who || '', '']));
    st.added = fit.length;
  }
  if (st.added || st.removed) sortCore_(ss);                       // нові плани: угору, до майбутнього
  return st;
}

/** Плани податків ФОП у журналі. Модуль «🧾 ФОП» вимкнено: лише прибрати майбутні. */
function fopPlansSync_(ss) {
  const st = { added: 0, updated: 0, removed: 0 };
  if (!ss.getSheetByName(JOURNAL) || !ss.getSheetByName(FOP_SHEET)) return st;
  const fops = modOn_('fop') ? fopList_(ss) : [], today = todayKey_(ss), rows = readJournal_(ss);
  const rates = fopRates_(ss, Number(today.slice(0, 4)));
  const want = fops.length && rates ? fopWanted_(fops, rates, rows, today, lastDayOf_(addDays_(lastDayOf_(today.slice(0, 7)), 1).slice(0, 7)), fopClosed_(ss)) : [];
  const r = tagPlansSync_(ss, FOP_TAG, want, rows);
  PropertiesService.getDocumentProperties().setProperties({ fop_day: today, fop_t: String(Date.now()) });
  return r;
}

/** З mbsTick: раз на день або після правок. */
function fopIfDue_(ss) {
  if (!modOn_('fop') || !ss.getSheetByName(FOP_SHEET) || !fopList_(ss).length) return null;
  const dp = PropertiesService.getDocumentProperties();
  if (dp.getProperty('fop_day') === todayKey_(ss) && Number(dp.getProperty('lastEdit') || 0) < Number(dp.getProperty('fop_t') || 0)) return null;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return null;
  try { return fopPlansSync_(ss); } finally { lock.releaseLock(); }
}

/** Для ранкового підсумку: ліміт доходу (тривога, коли поріг зріс: 80 → 90 → 100 %) і податки за 5 днів, за день і в день строку. */
function fopMorning_(m) {
  if (!modOn_('fop')) return;
  const fops = fopList_(m.ss);
  if (!fops.length) return;
  const rates = fopRates_(m.ss, Number(m.today.slice(0, 4)));
  if (!rates) return;
  const single = fops.length <= 1, y = m.today.slice(0, 4);
  fops.forEach(f => {
    const inc = fopIncome_(m.rows, f.name, single, y + '-01-01', (Number(y) + 1) + '-01-01'), lim = rates.limit[f.group], used = lim ? inc / lim : 0;
    const lvl = used >= 1 ? 100 : used >= 0.9 ? 90 : used >= 0.8 ? 80 : 0, k = 'fop_lvl:' + f.name + ':' + y, was = Number(m.dp.getProperty(k) || 0);
    if (lvl > was) {
      m.alerts.push((lvl >= 100 ? '⛔' : '⚠️') + ' <b>ФОП ' + esc_(f.name) + '</b>: дохід з початку року ' + money0_(inc) + ', це ' + Math.round(used * 100) + ' % ліміту ' +
        f.group + ' групи (' + money0_(lim) + ').' + (lvl >= 100 ? ' Перевищення: зверніться до бухгалтера щодо переходу.' : ''));
      if (!m.opts.test) m.dp.setProperty(k, String(lvl));
    }
  });
  m.rows.filter(o => isPlan_(o.e) && o.d.indexOf(FOP_TAG) === 0).forEach(o => {
    const left = dayDiff_(o.key, m.today);
    if (left === 5 || left === 1 || left === 0) m.alerts.push('🧾 <b>' + esc_(o.d.slice(FOP_TAG.length)) + '</b>: ' + money0_(o.b) + ', строк ' + dd_(o.key) + ' (' +
      (left === 0 ? 'сьогодні' : 'через ' + left + (left === 1 ? ' день' : ' днів')) + ').');
  });
}
