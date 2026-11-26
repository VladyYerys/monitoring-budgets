/** 24_Check.gs: «Перевірка»: чи сходяться суми й чи все в порядку з журналом і банком. */

/*
 * Рядок: що перевіряємо · значення · має бути · стан (OK / ПЕРЕВІРИТИ / інфо) · що робити.
 * Угорі підсумок. Ранковий звіт у Telegram читає стан і пише про рядки «ПЕРЕВІРИТИ».
 */

const CHECK_FIRST = 5;

/** [що, формула, має бути ('' = інфо), що робити]. */
function checkRows_() {
  const A = J_RANGE(JC.date), B = J_RANGE(JC.sum), C = J_RANGE(JC.cat), E = J_RANGE(JC.status), I = J_RANGE(JC.type), Y = J_RANGE(JC.bankId);
  const cpA = CP_RANGE('A'), cpB = CP_RANGE('B'), M = J_RANGE(JC.state), L = J_RANGE(JC.signed), AA = J_RANGE(JC.bankAmt), O = J_RANGE(JC.fact);
  const last = 'MAXIFS(' + cpA + ',' + cpB + ',"<>")', prev = 'MAXIFS(' + cpA + ',' + cpA + ',"<"&' + last + ',' + cpB + ',"<>")';
  const yr = "'" + MONTHS_SHEET + "'!$B$2", yearCol = colStr_(M_COL0 + 37);
  return [
    ['Невідомі категорії в журналі (немає в «Довіднику»)', '=COUNTIF(' + I + ',"⚠ невідома категорія")', 0, 'Додайте категорію в «Довідник» або виправте назву в журналі.'],
    ['Червоні рядки поза чернетками (переказ або без категорії)', '=COUNTIFS(' + C + ',"' + CAT_TRANSFER + '",' + E + ',"<>' + STATUS_DRAFT + '")+COUNTIFS(' +
      B + ',"<>",' + C + ',"",' + E + ',"<>' + STATUS_DRAFT + '")', 0, 'Оберіть справжню категорію: поки червоний, рядок не в статистиці.'],
    ['Недописані рядки (є сума без дати чи дата без суми)', '=COUNTIFS(' + A + ',"",' + B + ',"<>")+COUNTIFS(' + A + ',"<>",' + B + ',"")', 0,
      'Допишіть рядок або очистіть його.'],
    ['Операції банку, записані двічі (однаковий номер)', '=SUMPRODUCT((' + Y + '<>"")*(COUNTIF(' + Y + ',' + Y + ')>1))', 0, 'Видаліть дубль.'],
    ['Чернетки з банку, старші за 2 дні (за останній тиждень)', '=COUNTIFS(' + E + ',"' + STATUS_DRAFT + '",' + A + ',"<"&(TODAY()-2),' + A + ',">="&(TODAY()-7))', 0,
      'Поставте «✅ факт» (меню «✅ Підтвердити чернетки з категорією») або видаліть.'],
    ['Розбіжність з банком на останню контрольну точку, ₴', '=IFERROR(IF(' + prev + '=0,0,SUMIFS(' + cpB + ',' + cpA + ',' + prev + ')+SUMIFS(' + L + ',' + A + ',">="&' + prev +
      ',' + A + ',"<"&' + last + ')+SUMIFS(' + AA + ',' + M + ',2,' + A + ',">="&' + prev + ',' + A + ',"<"&' + last + ')-SUMIFS(' + cpB + ',' + cpA + ',' + last + ')),0)', 0,
      'Між двома останніми точками в журналі бракує операцій (або є зайві). Чернетки й перекази враховано.'],
    ['Усі факти року потрапили в «Місяці» (журнал − «Місяці», ₴)', '=ROUND(SUMIFS(' + O + ',' + A + ',">="&DATE(' + yr + ',1,1),' + A + ',"<"&DATE(' + yr + '+1,1,1),' + C +
      ',"<>",' + J_RANGE(JC.acctKind) + ',"<>' + K_FOP + '",' + J_RANGE(JC.who) + ',\'' + MONTHS_SHEET + '\'!$A$3)-SUM(\'' + MONTHS_SHEET + '\'!$' + yearCol + '$' + M_FIRST + ':$' + yearCol + '$' + (M_FIRST + D_LAST - D_FIRST) + '),2)', 0,
      'Факт з категорією, якої немає в «Довіднику», у «Місяці» не потрапляє.'],
    ['Є місце в журналі (щонайменше 100 вільних рядків)', '=IF(COUNTBLANK(' + B + ')>=100,1,0)', 1, 'Перенесіть старі роки в архів (з T9) або збільште журнал.'],
    ['Є місце для контрольних точок (щонайменше 30)', '=IF(COUNTBLANK(' + cpA + ')>=30,1,0)', 1, 'Очистіть найстаріші рядки на «' + CP_SHEET + '».'],
    ['«Дні» бачать сьогоднішній день', '=COUNTIF(\'' + DAYS_SHEET + '\'!$A$' + DAYS_FIRST + ':$A$' + (DAYS_FIRST + DAYS_ROWS - 1) + ',TODAY())', 1, ''],
    ['ФОП: дохід понад ліміт групи (скільки ФОП)', '=COUNTIF(\'' + FOP_SHEET + '\'!$I$' + FOP_FIRST + ':$I$' + FOP_LAST + ',"⛔*")', 0,
      'Перевищення ліміту: зверніться до бухгалтера щодо переходу в іншу групу.'],
    ['Рахунки в журналі, яких немає на «' + ACC_SHEET + '»', '=SUMPRODUCT((' + J_RANGE(JC.acct) + '<>"")*(COUNTIF(' + ACC_RANGE('A') + ',' + J_RANGE(JC.acct) + ')=0))', 0,
      'Ви перейменували рахунок? Поверніть назву на «' + ACC_SHEET + '» (так вона записана в журналі).'],
    ['Чернетки, старші за тиждень (інфо)', '=COUNTIFS(' + E + ',"' + STATUS_DRAFT + '",' + A + ',"<"&(TODAY()-7))', '', 'Без тривоги; поки чернетка, у статистику не входить.'],
    ['Планів з минулою датою (інфо)', '=COUNTIFS(' + M + ',1,' + A + ',"<"&TODAY())', '', 'Підтвердіть, перенесіть або видаліть.'],
  ];
}

function buildCheck_(ss) {
  const sh = sheetOf_(ss, CHECK_SHEET), rows = checkRows_(), last = CHECK_FIRST + rows.length - 1;
  sh.getRange(1, 1).setValue('✔️ Перевірка').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Підсумок:');
  setF_(sh.getRange(2, 2), [['=IF(COUNTIF($D$' + CHECK_FIRST + ':$D$' + last + ',"ПЕРЕВІРИТИ")=0,"✅ Усе OK","⚠️ Потребують уваги: "&COUNTIF($D$' +
    CHECK_FIRST + ':$D$' + last + ',"ПЕРЕВІРИТИ"))']]);
  sh.getRange(CHECK_FIRST - 1, 1, 1, 5).setValues([['Що перевіряємо', 'Значення', 'Має бути', 'Стан', 'Що робити']]).setFontWeight('bold');
  sh.getRange(CHECK_FIRST, 1, rows.length, 1).setValues(rows.map(r => [r[0]]));
  sh.getRange(CHECK_FIRST, 3, rows.length, 1).setValues(rows.map(r => [r[2]]));
  sh.getRange(CHECK_FIRST, 5, rows.length, 1).setValues(rows.map(r => [r[3]]));
  setF_(sh.getRange(CHECK_FIRST, 2, rows.length, 1), rows.map(r => [r[1]]));
  setF_(sh.getRange(CHECK_FIRST, 4, rows.length, 1), rows.map((r, i) => {
    const n = CHECK_FIRST + i;
    return [r[2] === '' ? '="інфо"' : '=IF(ISERROR(B' + n + '),"ПЕРЕВІРИТИ",IF(ABS(N(B' + n + ')-C' + n + ')<0.005,"OK","ПЕРЕВІРИТИ"))'];
  }));
  sh.getRange(CHECK_FIRST, 2, rows.length, 1).setNumberFormat('#,##0.##');
  const cf = SpreadsheetApp.newConditionalFormatRule().whenTextContains('ПЕРЕВІРИТИ').setBackground('#f8cfcb')
    .setRanges([sh.getRange('D' + CHECK_FIRST + ':D' + last)]).build();
  const ok = SpreadsheetApp.newConditionalFormatRule().whenTextContains('OK').setBackground('#d9ead3')
    .setRanges([sh.getRange('D' + CHECK_FIRST + ':D' + last)]).build();
  sh.setConditionalFormatRules([cf, ok]);
  sh.setFrozenRows(CHECK_FIRST - 1);
  return sh;
}

/** Рядки «ПЕРЕВІРИТИ» (для Telegram і «🩺 Стан таблиці»). → { bad: [текст], ok: число } */
function checkState_(ss) {
  const sh = ss.getSheetByName(CHECK_SHEET);
  if (!sh) return { bad: [], ok: 0, none: true };
  const v = sh.getRange(CHECK_FIRST, 1, Math.max(1, sh.getLastRow() - CHECK_FIRST + 1), 4).getValues();
  const bad = [];
  v.forEach(r => { if (String(r[3]).trim() === 'ПЕРЕВІРИТИ') bad.push('«' + String(r[0]).slice(0, 70) + '»: ' + r[1] + ' замість ' + r[2]); });
  return { bad: bad, ok: v.filter(r => String(r[3]).trim() === 'OK').length };
}
