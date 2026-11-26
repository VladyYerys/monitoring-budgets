/** 20_Months.gs: Аркуш «Місяці»: категорії й підсумки × 12 місяців обраного року, «план | факт | різниця». */

/*
 * Замість «полотна» з кожним днем (у особистій моделі ~1 900 колонок і десятки тисяч формул) тут 12 місяців:
 * факт рахується одразу з журналу за датою (SUMIFS по місяцю), ≈ 3 000 формул на весь рік.
 * Рік вибирається списком у B2; новий рік додавати не треба.
 * Різниця = факт − план. Рядки «мої гроші», «картка», «чому з ліміту», «непомітні» рахуються лише в колонках «Факт».
 * Рядок 5 прихований: перше число кожного місяця (на нього посилаються формули).
 */

const M_CATS = D_LAST - D_FIRST + 1;                     // 80 рядків категорій
const M_CAT_LAST = M_FIRST + M_CATS - 1;                 // 86
const M_YEAR_COL = M_COL0 + 36;                          // «Рік: план | факт | різниця»
const M_LAST_COL = M_YEAR_COL + 2;

/** Колонки місяця k (1…12): план, факт, різниця (номери). */
function mCols_(k) { const p = M_COL0 + (k - 1) * 3; return { plan: p, fact: p + 1, diff: p + 2 }; }

/** Рядки підсумків: [підпис, вид, параметр]. Вид визначає формулу (див. monthsSummaryFormula_). */
function monthsSummaryRows_() {
  const out = [['ПІДСУМКИ', 'head']];
  [T_SPEND, T_SAVE, T_TAX, T_DEBT, T_CORR_OUT].forEach(t => out.push([t === T_CORR_OUT ? 'Нерозподілені списання (Корекція −)' : t, 'type', t]));
  out.push(['Усі списання', 'sumUp', 5]);
  out.push([T_IN, 'type', T_IN], ['Нерозподілені надходження (Корекція +)', 'type', T_CORR_IN], ['Всього надходжень', 'sumUp', 2]);
  out.push(['Відкладено (тип «Резерв»)', 'type', T_RESERVE]);
  out.push(['МОЇ ГРОШІ', 'head'],
    ['Мої надходження', 'mineIn'], ['Витрачено з моїх', 'fromMine'], ['Витрачено з ліміту', 'fromLimit'],
    ['Мої гроші бюджетного місяця (усього)', 'mineBudget'], ['Відкладено на погашення (бюджетний місяць)', 'setAside'],
    ['Мої на кінець місяця', 'mineEnd'], ['Погашено за бюджетний місяць (відкладено + невитрачене)', 'repaid'], ['Орієнтир погашення', 'goal']);
  out.push(['КАРТКА', 'head'], ['На картці на кінець місяця', 'cardEnd'], ['Борг за лімітом на кінець місяця', 'debtEnd']);
  out.push(['ЧОМУ З ЛІМІТУ', 'head']);
  for (let i = 0; i <= D_REASONS_LAST - D_REASONS_FIRST; i++) out.push(['', 'reason', D_REASONS_FIRST + i]);
  out.push(['З ліміту: без причини', 'noReason'], ['Разом з ліміту', 'sumBlock', 'ЧОМУ З ЛІМІТУ']);
  out.push(['НЕПОМІТНІ ВИТРАТИ', 'head']);
  for (let i = 0; i <= D_HOLES_LAST - D_HOLES_FIRST; i++) out.push(['', 'hole', D_HOLES_FIRST + i]);
  out.push(['Непомітні разом', 'sumBlock', 'НЕПОМІТНІ ВИТРАТИ']);
  return out;
}

/**
 * Картка на кінець дня x (формула-рядок): остання контрольна точка ≤ x + усі рухи журналу від неї до x.
 * До першої контрольної точки: назад від найближчої наступної (точка мінус рухи між x і нею), як в особистій таблиці;
 * так історія, завантажена з банку до першої точки, має справжній залишок, а не відлік від нуля. Точок немає: рухи від нуля.
 */
function cardAt_(x) {
  const cpA = CP_RANGE('A'), cpB = CP_RANGE('B'), L = J_RANGE(JC.signed), A = J_RANGE(JC.date);
  const last = 'MAXIFS(' + cpA + ',' + cpA + ',"<="&' + x + ',' + cpB + ',"<>")', next = 'MINIFS(' + cpA + ',' + cpA + ',">"&' + x + ',' + cpB + ',"<>")';
  return 'IF(' + last + '=0,IF(' + next + '=0,SUMIFS(' + L + ',' + A + ',"<="&' + x + '),SUMIFS(' + cpB + ',' + cpA + ',' + next + ')-SUMIFS(' + L + ',' + A + ',">"&' + x +
    ',' + A + ',"<"&' + next + ')),SUMIFS(' + cpB + ',' + cpA + ',' + last + ')+SUMIFS(' + L + ',' + A + ',">="&' + last + ',' + A + ',"<="&' + x + '))';
}

function monthsSummaryFormula_(kind, param, colL, r, rows, firstSumRow) {
  const A = J_RANGE(JC.date), m0 = colL + '$' + M_MONTH_ROW, inMonth = A + ',">="&' + m0 + ',' + A + ',"<"&EDATE(' + m0 + ',1)';
  const end = 'EOMONTH(' + m0 + ',0)', start = 'DATE(YEAR(' + SET_REF('start') + '),MONTH(' + SET_REF('start') + '),1)';
  const cats = (c) => '$' + c + '$' + M_FIRST + ':$' + c + '$' + M_CAT_LAST;
  switch (kind) {
    case 'type': return '=SUMIF(' + cats('B') + ',"' + param + '",' + colL + '$' + M_FIRST + ':' + colL + '$' + M_CAT_LAST + ')';
    case 'sumUp': return '=SUM(' + colL + (r - param) + ':' + colL + (r - 1) + ')';
    case 'mineIn': return '=SUMIFS(' + J_RANGE(JC.mineIn) + ',' + inMonth + ',' + J_RANGE(JC.mineIn) + ',">0")';
    case 'fromMine': return '=SUMIFS(' + J_RANGE(JC.fromMine) + ',' + inMonth + ')';
    case 'fromLimit': return '=SUMIFS(' + J_RANGE(JC.fromLimit) + ',' + inMonth + ')';
    case 'mineBudget': return '=SUMIFS(' + J_RANGE(JC.mineIn) + ',' + J_RANGE(JC.bmonth) + ',' + m0 + ')';
    case 'setAside': return '=SUMIFS(' + J_RANGE(JC.fact) + ',' + J_RANGE(JC.mineKind) + ',"' + MINE_SET_ASIDE + '",' + J_RANGE(JC.bmonth) + ',' + m0 + ')';
    case 'mineEnd': return '=IF(' + m0 + '<' + start + ',"",SUMIFS(' + J_RANGE(JC.mineIn) + ',' + J_RANGE(JC.bmonth) + ',' + m0 + ',' + A + ',"<="&' + end + ')-SUMIFS(' +
      J_RANGE(JC.out) + ',' + J_RANGE(JC.bmonth) + ',' + m0 + ',' + A + ',"<="&' + end + '))';
    case 'repaid': return '=IF(' + colL + (r - 1) + '="","",' + colL + (r - 2) + '+' + colL + (r - 1) + ')';
    case 'goal': return '=' + SET_REF('goal');
    case 'cardEnd': return '=' + cardAt_(end);
    case 'debtEnd': return '=IF(N(' + colL + (r - 1) + ')<0,-' + colL + (r - 1) + ',0)';
    case 'reason': return '=IF($A' + r + '="","",SUMIFS(' + J_RANGE(JC.fromLimit) + ',' + J_RANGE(JC.limitWhy) + ',$A' + r + ',' + inMonth + mWho_() + '))';
    case 'noReason': return '=SUMIFS(' + J_RANGE(JC.fromLimit) + ',' + J_RANGE(JC.limitWhy) + ',"без причини",' + inMonth + mWho_() + ')';
    case 'hole': return '=IF($A' + r + '="","",SUMIFS(' + J_RANGE(JC.fact) + ',' + J_RANGE(JC.holeEff) + ',$A' + r + ',' + inMonth + mWho_() + '))';
    case 'sumBlock': { const h = firstSumRow + rows.findIndex(x => x[0] === param); return '=SUM(' + colL + (h + 1) + ':' + colL + (r - 1) + ')'; }
  }
  return '';
}

/** Додаткові умови SUMIFS для «Місяців»: не рахунок ФОП і «Хто» з вибору в D2 (службовий критерій у A3). */
function mWho_() { return ',' + J_RANGE(JC.acctKind) + ',"<>' + K_FOP + '",' + J_RANGE(JC.who) + ',' + M_WHO_CRIT; }

/** Рядки-підсумки, у яких «план» і «різниця» теж мають сенс (суми категорій за типом). */
const M_PLAN_KINDS = ['type', 'sumUp'];
/** Рядки-залишки (на кінець місяця): у «Рік» не сумуються. */
const M_BALANCE_KINDS = ['mineEnd', 'repaid', 'goal', 'cardEnd', 'debtEnd'];

function buildMonths_(ss) {
  const sh = ensureSize_(sheetOf_(ss, MONTHS_SHEET), M_CAT_LAST + 2 + monthsSummaryRows_().length + 2, M_LAST_COL), tz = ss.getSpreadsheetTimeZone();
  const year = Number(Utilities.formatDate(new Date(), tz, 'yyyy'));
  sh.getRange(1, 1).setValue('📅 Місяці').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Рік:');
  const keepYear = Number(sh.getRange(M_YEAR).getValue()) || year;
  sh.getRange(M_YEAR).setValue(keepYear).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Array.from({ length: 11 }, (x, i) => String(year - 5 + i))).setAllowInvalid(true).build());
  sh.getRange(2, 3).setValue('Хто:');
  const set = ss.getSheetByName(SETTINGS_SHEET), keepWho = String(sh.getRange(M_WHO).getValue() || '') || WHO_ALL;
  sh.getRange(M_WHO).setValue(keepWho);
  if (set) sh.getRange(M_WHO).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(set.getRange('B' + SET_MEMBERS_FIRST + ':B' + SET_ALL_ROW)).setAllowInvalid(true).build());
  setF_(sh.getRange(M_WHO_CRIT.replace(/\$/g, '')), [['=IF(OR($D$2="",$D$2="' + WHO_ALL + '"),"<>✱",$D$2)']]);
  sh.getRange(M_WHO_CRIT.replace(/\$/g, '')).setFontColor('#ffffff');
  sh.getRange(2, 5).setValue('Факт з журналу за датою операції (без рахунків ФОП); «Різниця» = факт − план. «Хто» фільтрує категорії, причини й непомітні; ' +
    '«мої гроші» і «картка»: уся сім\'я.').setFontColor('#5f6368');
  // заголовки
  const head3 = [], head4 = ['Категорія / показник', 'Тип', 'Група'], row5 = ['', '', ''];
  for (let k = 1; k <= 12; k++) {
    head3.push(MONTHS_UA[k - 1], '', '');
    head4.push('План', 'Факт', 'Різниця');
    const f = '=DATE($B$2,' + k + ',1)';
    row5.push(f, f, f);
  }
  head3.push('Рік', '', ''); head4.push('Рік: план', 'Рік: факт', 'Рік: різниця'); row5.push('', '', '');
  sh.getRange(3, M_COL0, 1, head3.length).setValues([head3]).setFontWeight('bold');
  sh.getRange(M_HEAD, 1, 1, head4.length).setValues([head4]).setFontWeight('bold');
  setF_(sh.getRange(M_MONTH_ROW, 1, 1, row5.length), [row5]);
  for (let k = 1; k <= 13; k++) sh.getRange(3, M_COL0 + (k - 1) * 3, 1, 3).merge();
  // категорії
  const C = J_RANGE(JC.cat), A = J_RANGE(JC.date), O = J_RANGE(JC.fact), d = "'" + DICT_SHEET + "'!";
  const catRows = [];
  for (let i = 0; i < M_CATS; i++) {
    const r = M_FIRST + i, dr = D_FIRST + i, row = [
      '=IF(' + d + '$A$' + dr + '="","",' + d + '$A$' + dr + ')', '=IF($A' + r + '="","",' + d + '$B$' + dr + '&"")', '=IF($A' + r + '="","",' + d + '$C$' + dr + '&"")'];
    for (let k = 1; k <= 12; k++) {
      const c = mCols_(k), P = colStr_(c.plan), F = colStr_(c.fact), m0 = F + '$' + M_MONTH_ROW;
      row.push('=IF($A' + r + '="","",IF(OR($D$2="",$D$2="' + WHO_ALL + '",' + d + '$G$' + dr + '=$D$2),N(' + d + '$F$' + dr + '),0))',
        '=IF($A' + r + '="","",SUMIFS(' + O + ',' + C + ',$A' + r + ',' + A + ',">="&' + m0 + ',' + A + ',"<"&EDATE(' + m0 + ',1)' + mWho_() + '))',
        '=IF($A' + r + '="","",' + F + r + '-' + P + r + ')');
    }
    row.push(...monthsYearCells_(r, true));
    catRows.push(row);
  }
  setF_(sh.getRange(M_FIRST, 1, M_CATS, M_LAST_COL), catRows);
  // підсумки
  const rows = monthsSummaryRows_(), first = M_CAT_LAST + 2, out = [];
  rows.forEach((x, i) => {
    const r = first + i, kind = x[1], line = [];
    if (kind === 'head') { line.push(x[0]); while (line.length < M_LAST_COL) line.push(''); out.push(line); return; }
    line.push(kind === 'reason' ? '=IF(' + d + '$K$' + x[2] + '="","",' + d + '$K$' + x[2] + ')' : kind === 'hole' ? '=IF(' + d + '$I$' + x[2] + '="","",' + d + '$I$' + x[2] + ')' : x[0], '', '');
    for (let k = 1; k <= 12; k++) {
      const c = mCols_(k);
      if (M_PLAN_KINDS.indexOf(kind) >= 0) {
        line.push(monthsSummaryFormula_(kind, x[2], colStr_(c.plan), r, rows, first), monthsSummaryFormula_(kind, x[2], colStr_(c.fact), r, rows, first),
          monthsSummaryFormula_(kind, x[2], colStr_(c.diff), r, rows, first));
      } else line.push('', monthsSummaryFormula_(kind, x[2], colStr_(c.fact), r, rows, first), '');
    }
    line.push(...(M_BALANCE_KINDS.indexOf(kind) >= 0 ? ['', '', ''] : monthsYearCells_(r, M_PLAN_KINDS.indexOf(kind) >= 0)));
    out.push(line);
  });
  setF_(sh.getRange(first, 1, out.length, M_LAST_COL), out);
  out.forEach((l, i) => { if (String(l[0]).charAt(0) !== '=' && l[0]) sh.getRange(first + i, 1).setValue(l[0]); });
  rows.forEach((x, i) => { if (x[1] === 'head') sh.getRange(first + i, 1).setFontWeight('bold'); });
  sh.getRange(M_FIRST, M_COL0, first + rows.length - M_FIRST, M_LAST_COL - M_COL0 + 1).setNumberFormat('#,##0');
  sh.hideRows(M_MONTH_ROW, 1);
  sh.setFrozenRows(M_HEAD); sh.setFrozenColumns(1);
  return sh;
}

/** «Рік: план | факт | різниця» для рядка r: суми колонок «План» і «Факт» за 12 місяців. */
function monthsYearCells_(r, withPlan) {
  const hdr = '$' + colStr_(M_COL0) + '$' + M_HEAD + ':$' + colStr_(M_YEAR_COL - 1) + '$' + M_HEAD;
  const rowR = colStr_(M_COL0) + r + ':' + colStr_(M_YEAR_COL - 1) + r, P = colStr_(M_YEAR_COL), F = colStr_(M_YEAR_COL + 1);
  return [withPlan ? '=IF($A' + r + '="","",SUMIF(' + hdr + ',"План",' + rowR + '))' : '',
    '=IF($A' + r + '="","",SUMIF(' + hdr + ',"Факт",' + rowR + '))',
    withPlan ? '=IF($A' + r + '="","",' + F + r + '-' + P + r + ')' : ''];
}

/** Номер рядка підсумку за підписом (для «Історії місяців» і звітів). */
function monthsSummaryRow_(label) {
  const rows = monthsSummaryRows_(), i = rows.findIndex(x => x[0] === label);
  return i < 0 ? 0 : M_CAT_LAST + 2 + i;
}
