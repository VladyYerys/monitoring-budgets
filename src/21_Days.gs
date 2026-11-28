/** 21_Days.gs: Аркуш «Дні»: баланси по днях від початку минулого місяця до кінця наступного. */

/*
 * Дні рядками (зручно на телефоні), ≈ 90 рядків, вікно зсувається саме з датою.
 * Колонки: надходження і списання дня, контрольна точка банку, на картці на кінець дня, борг за лімітом,
 * кредитний ліміт, доступно з лімітом, «мої» на кінець дня, плани дня, прогноз за планом і «як зараз».
 * На картці: остання контрольна точка банку + усі рухи журналу після неї (як «Залишок» в особистій моделі).
 */

const DAYS_ROWS = 93;                                     // 31 + 31 + 31
const DAYS_HEAD = ['Дата', 'Надходження', 'Списання', 'Контрольна точка (на початок дня)', 'На картці (кінець дня)', 'Борг за лімітом',
  'Кредитний ліміт', 'Доступно з лімітом', 'Мої (кінець дня)', 'План: надходження', 'План: списання', 'Прогноз за планом', 'Як зараз'];

function daysFormulas_(r, first) {
  const L = J_RANGE(JC.signed), A = J_RANGE(JC.date), N = J_RANGE(JC.plan), I = J_RANGE(JC.type);
  const cpA = CP_RANGE('A'), cpB = CP_RANGE('B'), cpC = CP_RANGE('C');
  const start = 'DATE(YEAR(' + SET_REF('start') + '),MONTH(' + SET_REF('start') + '),1)', m1 = 'DATE(YEAR($A' + r + '),MONTH($A' + r + '),1)';
  const blank = f => '=IF($A' + r + '="","",' + f + ')';
  const prev = r - 1;
  return [
    r === first ? '=EOMONTH(TODAY(),-2)+1' : '=IF($A' + prev + '="","",IF($A' + prev + '+1>EOMONTH(TODAY(),1),"",$A' + prev + '+1))',
    blank('SUMIFS(' + L + ',' + A + ',$A' + r + ',' + L + ',">0")'),
    blank('-SUMIFS(' + L + ',' + A + ',$A' + r + ',' + L + ',"<0")'),
    blank('IF(COUNTIFS(' + cpA + ',$A' + r + ',' + cpB + ',"<>")=0,"",SUMIFS(' + cpB + ',' + cpA + ',$A' + r + '))'),
    blank('IF($D' + r + '<>"",$D' + r + ',' + (r === first ? cardAt_('($A' + r + '-1)') : 'N($E' + prev + ')') + ')+$B' + r + '-$C' + r),
    blank('IF(N($E' + r + ')<0,-$E' + r + ',0)'),
    blank('IFERROR(SUMIFS(' + cpC + ',' + cpA + ',MAXIFS(' + cpA + ',' + cpA + ',"<="&$A' + r + ',' + cpC + ',"<>")),0)'),
    blank('$E' + r + '+$G' + r),
    blank('IF(' + m1 + '<' + start + ',"",SUMIFS(' + J_RANGE(JC.mineIn) + ',' + J_RANGE(JC.bmonth) + ',' + m1 + ',' + A + ',"<="&$A' + r + ')-SUMIFS(' +
      J_RANGE(JC.out) + ',' + J_RANGE(JC.bmonth) + ',' + m1 + ',' + A + ',"<="&$A' + r + '))'),
    blank('SUMIFS(' + N + ',' + A + ',$A' + r + ',' + I + ',"' + T_IN + '")+SUMIFS(' + N + ',' + A + ',$A' + r + ',' + I + ',"' + T_CORR_IN + '")'),
    blank('SUMIFS(' + N + ',' + A + ',$A' + r + ')-$J' + r + '-SUMIFS(' + N + ',' + A + ',$A' + r + ',' + I + ',"' + T_RESERVE + '")'),
    blank('IF($A' + r + '<TODAY(),$E' + r + ',' + (r === first ? '$E' + r : 'IF($A' + r + '=TODAY(),$E' + r + ',N($L' + prev + '))') + '+$J' + r + '-$K' + r + ')'),
    blank('IF($A' + r + '<=TODAY(),$L' + r + ',$L' + r + '-$B$3*($A' + r + '-TODAY()))'),
  ];
}

/**
 * Підсумки «Днів» (B2:C4). Найнижча точка лише від сьогодні (версія 19): минулі дні до першої контрольної точки банку
 * відновлюються назад лише з фактів, тож там бувають неможливі числа.
 */
function daysTop_(sh) {
  const lastRow = DAYS_FIRST + DAYS_ROWS - 1, Lr = '$L$' + DAYS_FIRST + ':$L$' + lastRow, Mr = '$M$' + DAYS_FIRST + ':$M$' + lastRow, Ar = '$A$' + DAYS_FIRST + ':$A$' + lastRow;
  const A = J_RANGE(JC.date), O = J_RANGE(JC.fact), I = J_RANGE(JC.type);
  sh.getRange(2, 1, 3, 1).setValues([['Найнижча точка за планом (від сьогодні)'], ['Темп витрат за 30 днів, ₴/день'], ['Найнижча точка «як зараз» (від сьогодні)']]);
  setF_(sh.getRange(2, 2, 3, 2), [
    ['=MINIFS(' + Lr + ',' + Ar + ',">="&TODAY())', '=IFERROR(INDEX(FILTER(' + Ar + ',' + Lr + '=B2,' + Ar + '>=TODAY()),1),"")'],
    ['=SUMIFS(' + O + ',' + I + ',"' + T_SPEND + '",' + A + ',">="&(TODAY()-30),' + A + ',"<"&TODAY())/30', ''],
    ['=MINIFS(' + Mr + ',' + Ar + ',">="&TODAY())', '=IFERROR(INDEX(FILTER(' + Ar + ',' + Mr + '=B4,' + Ar + '>=TODAY()),1),"")'],
  ]);
}

function buildDays_(ss) {
  const sh = ensureSize_(sheetOf_(ss, DAYS_SHEET), DAYS_FIRST + DAYS_ROWS, DAYS_HEAD.length), A = J_RANGE(JC.date), O = J_RANGE(JC.fact), I = J_RANGE(JC.type);
  const lastRow = DAYS_FIRST + DAYS_ROWS - 1, Lr = 'L' + DAYS_FIRST + ':L' + lastRow, Mr = 'M' + DAYS_FIRST + ':M' + lastRow, Ar = 'A' + DAYS_FIRST + ':A' + lastRow;
  sh.getRange(1, 1).setValue('📆 Дні').setFontWeight('bold').setFontSize(14);
  daysTop_(sh);
  sh.getRange(DAYS_FIRST - 1, 1, 1, DAYS_HEAD.length).setValues([DAYS_HEAD]).setFontWeight('bold').setWrap(true);
  const rows = [];
  for (let r = DAYS_FIRST; r <= lastRow; r++) rows.push(daysFormulas_(r, DAYS_FIRST));
  setF_(sh.getRange(DAYS_FIRST, 1, DAYS_ROWS, DAYS_HEAD.length), rows);
  sh.getRange(Ar).setNumberFormat('dd.mm.yyyy');
  sh.getRange(DAYS_FIRST, 2, DAYS_ROWS, DAYS_HEAD.length - 1).setNumberFormat('#,##0');
  sh.getRange('C2').setNumberFormat('dd.mm.yyyy'); sh.getRange('C4').setNumberFormat('dd.mm.yyyy');
  sh.setFrozenRows(DAYS_FIRST - 1); sh.setFrozenColumns(1);
  return sh;
}
