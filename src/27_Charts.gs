/** 27_Charts.gs: «Дані графіків»: схований аркуш з формулами для «Charts» і «charts_mob» (28_Dash.gs). */

/*
 * Лише формули (журнал, «Дні», «Довідник», «Залишки банку»); правити не треба, аркуш завжди схований.
 * «Цей місяць»: з 1-го по сьогодні; порівняння з середнім місяцем попереднього кварталу (квартал ÷ 3), як в особистій таблиці.
 *   A:B   ключові цифри (сьогодні, місяць, квартал, залишок, ліміт, мої, прогноз, погашення, «чорні діри»)
 *   D:L   категорії: цей місяць, минулий, квартал, середній місяць, «більше на», ключ рейтингу
 *   N:S   «Що зросло»: топ-10 категорій витрат понад середнє      U:Z  групи (назви пише скрипт)   AC:AH групи за спаданням
 *   AJ:AN типи                AQ:AU залишок за 60 днів (плюс, мінус, межа ліміту)      AX:BA цей місяць по днях: витрачено, темп, мої
 *   BC:BF залишок: факт, прогноз, межа (від сьогодні −30 до кінця наступного місяця)     BH:BJ з моїх чи з ліміту, 6 місяців
 *   BL:BO борг за лімітом: факт, прогноз, орієнтир    BQ:CA «чорні діри» (ліміт: «Charts», колонка «Ліміт на місяць ✎»)
 *   CC:CK «чорні діри» за 6 місяців                    CM:CQ причини «з ліміту»        CS:CX причини за спаданням
 * Порожні точки графіків: NA(), щоб лінії мали розрив, а не падали в нуль.
 */

const CH_SHEET = 'Дані графіків', DASH_SHEET = 'Charts', MOB_SHEET = 'charts_mob', PHONE_SHEET = '📱 Телефон';
const CH_ROWS = 94, CH_COLS = 102;                                          // A…CX
const CH_GROUPS = 12, CH_HOLES = 8, CH_UP = 10, CH_CATS = D_LAST - D_FIRST + 1, CH_WHY = D_REASONS_LAST - D_REASONS_FIRST + 2;
/** Перша колонка кожного блоку «Даних графіків». */
const CHD = { cat: 4, up: 14, grp: 21, grpS: 29, typ: 36, b60: 43, pace: 50, fc: 55, ml: 60, debt: 64, hole: 69, h6: 81, why: 91, whyS: 97 };
const CH_MON = ['січ', 'лют', 'бер', 'квіт', 'трав', 'черв', 'лип', 'серп', 'вер', 'жовт', 'лист', 'груд'];
const CH_MONTHS = ['Січень', 'Лютий', 'Березень', 'Квітень', 'Травень', 'Червень', 'Липень', 'Серпень', 'Вересень', 'Жовтень', 'Листопад', 'Грудень'];
const CH_TYPES = [T_SPEND, T_DEBT, T_SAVE, T_TAX];
const CH_TOTAL_OUT = 'Разом списано (4 типи)', CH_WHY_NONE = 'без причини', CH_HOLES_ALL = 'Усі непомітні витрати', CH_WHY_ALL = 'Разом з ліміту';
const H_OVER = '▲▲ понад', H_WARN = '▲ увага', H_OK = '● норма', H_NONE = '○ без бази';

/** Колонка блоку: A1 з $ (r2 → діапазон). */
function chCol_(c, r1, r2) { const l = colStr_(c); return '$' + l + '$' + r1 + (r2 ? ':$' + l + '$' + r2 : ''); }
/** Посилання на «Дані графіків» з іншого аркуша. */
function chRef_(a1) { return "'" + CH_SHEET + "'!" + a1; }

/** Ключові цифри: [ім'я, підпис, формула з {{ім'я}}, формат]. */
function chartsKeys_() {
  const A = J_RANGE(JC.date), O = J_RANGE(JC.fact), S = J_RANGE(JC.mineIn), T = J_RANGE(JC.out), Q = J_RANGE(JC.bmonth), N = J_RANGE(JC.plan), R = J_RANGE(JC.mineKind),
    I = J_RANGE(JC.type), W = J_RANGE(JC.fromLimit), X = J_RANGE(JC.limitWhy), noFop = ',' + J_RANGE(JC.acctKind) + ',"<>' + K_FOP + '"';
  const wMonth = A + ',">="&{{m0}},' + A + ',"<"&EDATE({{m0}},1)', days = c => "'" + DAYS_SHEET + "'!$" + c, H = (c, r) => chCol_(CHD.hole + c, r || 2, r ? 0 : 1 + CH_HOLES);
  const name = x => 'CHOOSE(MONTH(' + x + '),"' + CH_MONTHS.join('","') + '")&" "&YEAR(' + x + ')';
  const planOut = TYPES_OUT.map(t => 'SUMIFS(' + N + ',' + I + ',"' + t + '",' + Q + ',{{m0}}' + noFop + ')').join('+');
  const D = 'dd.mm.yyyy', M = '#,##0', dLast = DAYS_FIRST + DAYS_ROWS - 1, dA = days('A$' + DAYS_FIRST + ':$A$' + dLast);
  return [
    ['today', 'Сьогодні', '=TODAY()', D],
    ['m0', 'Початок цього місяця', '=DATE(YEAR({{today}}),MONTH({{today}}),1)', D],
    ['pm0', 'Початок минулого місяця', '=EDATE({{m0}},-1)', D],
    ['pmEnd', 'Кінець минулого місяця', '={{m0}}-1', D],
    ['q0', 'Початок цього кварталу', '=DATE(YEAR({{today}}),INT((MONTH({{today}})-1)/3)*3+1,1)', D],
    ['pq0', 'Початок попереднього кварталу', '=EDATE({{q0}},-3)', D],
    ['pqEnd', 'Кінець попереднього кварталу', '={{q0}}-1', D],
    ['qLabel', 'Попередній квартал', '="Q"&(INT((MONTH({{pq0}})-1)/3)+1)&" "&YEAR({{pq0}})', ''],
    ['mName', 'Цей місяць', '=' + name('{{today}}'), ''],
    ['pmName', 'Минулий місяць', '=' + name('{{pm0}}'), ''],
    ['daysPassed', 'Днів цього місяця минуло', '=DAY({{today}})', '0'],
    ['daysIn', 'Днів у місяці', '=DAY(EOMONTH({{today}},0))', '0'],
    ['daysLeft', 'Днів до кінця місяця (разом із сьогодні)', '=EOMONTH({{today}},0)-{{today}}+1', '0'],
    ['win', 'Початок вікна «Днів»', '=EOMONTH({{today}},-2)+1', D],
    ['bal', 'Залишок на кінець сьогодні', '=N(' + dayAt_('E', '{{today}}') + ')', M],
    ['balPm', 'Залишок на кінець минулого місяця', '=N(' + dayAt_('E', '{{pmEnd}}') + ')', M],
    ['change', 'Зміна за місяць', '={{bal}}-{{balPm}}', M],
    ['limit', 'Кредитний ліміт', '=N(' + dayAt_('G', '{{today}}') + ')', M],
    ['avail', 'Доступно з лімітом', '={{bal}}+{{limit}}', M],
    ['debt', 'Борг за лімітом зараз', '=MAX(0,-{{bal}})', M],
    ['mine', 'Доступно моїх сьогодні', '=N(' + dayAt_('I', '{{today}}') + ')', M],
    ['perDay', 'Можна на день', '=IF({{mine}}>0,{{mine}}/{{daysLeft}},0)', M],
    ['mineMonth', 'Мої гроші місяця (факт + план)', '=SUMIFS(' + S + ',' + Q + ',{{m0}})+SUMIFS(' + N + ',' + R + ',"' + MINE_IN + '",' + Q + ',{{m0}}' + noFop +
      ')-SUMIFS(' + N + ',' + R + ',"' + MINE_SET_ASIDE + '",' + Q + ',{{m0}}' + noFop + ')', M],
    ['spentMine', 'Списано в бюджетному місяці', '=SUMIFS(' + T + ',' + Q + ',{{m0}})', M],
    ['planOut', 'Ще заплановано списань у бюджетному місяці', '=' + planOut, M],
    ['mineEnd', 'Мої на кінець місяця (прогноз)', '={{mineMonth}}-{{spentMine}}-{{planOut}}', M],
    ['mineEndText', 'Текст: мої до кінця місяця', '=IF({{mineEnd}}<0,"▲▲ мінус "&FIXED(-{{mineEnd}},0),"● лишиться "&FIXED({{mineEnd}},0))', ''],
    ['fromLimit', 'З ліміту цього місяця', '=SUMIFS(' + W + ',' + wMonth + ')', M],
    ['limitOps', 'Операцій з ліміту цього місяця', '=COUNTIFS(' + W + ',">0",' + wMonth + ')', '0'],
    ['noReason', 'З ліміту без причини цього місяця', '=SUMIFS(' + W + ',' + X + ',"' + CH_WHY_NONE + '",' + wMonth + ')', M],
    ['setAside', 'Відкладено на погашення (бюджетний місяць)', '=SUMIFS(' + O + ',' + R + ',"' + MINE_SET_ASIDE + '",' + Q + ',{{m0}})', M],
    ['goal', 'Орієнтир погашення на місяць', '=N(' + SET_REF('goal') + ')', M],
    ['goalText', 'Текст: погашення ліміту', '=IF({{goal}}<=0,"○ орієнтир не задано",IF({{mine}}<0,"▲▲ у ліміті "&FIXED(-{{mine}},0),IF({{setAside}}<{{goal}},' +
      '"▲ відкладено ","● відкладено ")&FIXED({{setAside}},0)&" з "&FIXED({{goal}},0)))', ''],
    ['debtBase', 'Борг на кінець минулого місяця', '=MAX(0,-{{balPm}})', M],
    ['debt4', 'Борг чотири місяці тому', '=MAX(0,-(' + cardAt_('EOMONTH({{today}},-4)') + '))', M],
    ['repay3', 'Погашено в середньому за 3 місяці', '=MAX(0,({{debt4}}-{{debtBase}})/3)', M],
    ['repayPace', 'Темп погашення (або орієнтир)', '=IF({{repay3}}>0,{{repay3}},{{goal}})', M],
    ['repayMonths', 'Місяців до повного погашення', '=IF({{debt}}<=0,0,IF(N({{repayPace}})<=0,"",ROUNDUP({{debt}}/{{repayPace}},0)))', '0'],
    ['repayText', 'Текст: повне погашення', '=IF({{debt}}<=0,"● боргу немає",IF({{repayMonths}}="","▲▲ не зменшується","≈ "&{{repayMonths}}&" міс. · "&' +
      'TEXT(EDATE({{m0}},{{repayMonths}}),"mm.yyyy")&IF({{repay3}}>0,"","*")))', ''],
    ['fcEnd', 'Прогноз залишку на кінець місяця', '=N(' + dayAt_('L', 'EOMONTH({{today}},0)') + ')', M],
    ['low', 'Найнижча точка за планом (від сьогодні)', '=MINIFS(' + days('L$' + DAYS_FIRST + ':$L$' + dLast) + ',' + dA + ',">="&{{today}})', M],
    ['lowDate', 'Дата найнижчої точки', '=IFERROR(1/(1/MINIFS(' + dA + ',' + days('L$' + DAYS_FIRST + ':$L$' + dLast) + ',{{low}},' + dA + ',">="&{{today}})),"")', D],
    ['lowText', 'Текст: найнижча точка', '=IF({{lowDate}}="","○ немає",FIXED({{low}},0)&" · "&TEXT({{lowDate}},"dd.mm"))', ''],
    ['nowEnd', 'Як зараз: на кінець місяця', '=N(' + dayAt_('M', 'EOMONTH({{today}},0)') + ')', M],
    ['pace30', 'Витрачаєте ≈ на день', '=N(' + days('B$3') + ')', M],
    ['lowNow', 'Найнижча точка «як зараз» (від сьогодні)', '=MINIFS(' + days('M$' + DAYS_FIRST + ':$M$' + dLast) + ',' + dA + ',">="&{{today}})', M],
    ['lowNowDate', 'Дата найнижчої «як зараз»', '=IFERROR(1/(1/MINIFS(' + dA + ',' + days('M$' + DAYS_FIRST + ':$M$' + dLast) + ',{{lowNow}},' + dA + ',">="&{{today}})),"")', D],
    ['lowNowText', 'Текст: найнижча «як зараз»', '=IF({{lowNowDate}}="","○ немає",FIXED({{lowNow}},0)&" · "&TEXT({{lowNowDate}},"dd.mm"))', ''],
    ['hBig', 'Найбільша «чорна діра» цього місяця', '=IF(MAX(' + H(1) + ')>0,INDEX(' + H(0) + ',MATCH(MAX(' + H(10) + '),' + H(10) + ',0)),"○ немає")', ''],
    ['hBigSum', 'Її сума', '=MAX(' + H(1) + ')', M],
    ['hOver', '«Чорних дір» понад ліміт', '=COUNTIF(' + H(9) + ',"' + H_OVER + '")', '0'],
    ['hWarn', '«Чорних дір» «увага»', '=COUNTIF(' + H(9) + ',"' + H_WARN + '")', '0'],
    ['hOk', '«Чорних дір» у нормі', '=COUNTIF(' + H(9) + ',"' + H_OK + '")', '0'],
    ['hUsed', 'Використано лімітів «чорних дір»', '=IF(N(' + H(6, 2 + CH_HOLES) + ')>0,SUMIFS(' + H(1) + ',' + H(6) + ',">0")/' + H(6, 2 + CH_HOLES) + ',"")', '0%'],
    ['hLeft', 'Лишилось у межах лімітів', '=IF(N(' + H(6, 2 + CH_HOLES) + ')>0,' + H(6, 2 + CH_HOLES) + '-SUMIFS(' + H(1) + ',' + H(6) + ',">0"),"")', M],
  ];
}
const CH_KEYS = ['today', 'm0', 'pm0', 'pmEnd', 'q0', 'pq0', 'pqEnd', 'qLabel', 'mName', 'pmName', 'daysPassed', 'daysIn', 'daysLeft', 'win', 'bal', 'balPm', 'change',
  'limit', 'avail', 'debt', 'mine', 'perDay', 'mineMonth', 'spentMine', 'planOut', 'mineEnd', 'mineEndText', 'fromLimit', 'limitOps', 'noReason', 'setAside', 'goal',
  'goalText', 'debtBase', 'debt4', 'repay3', 'repayPace', 'repayMonths', 'repayText', 'fcEnd', 'low', 'lowDate', 'lowText', 'nowEnd', 'pace30', 'lowNow', 'lowNowDate',
  'lowNowText',
  'hBig', 'hBigSum', 'hOver', 'hWarn', 'hOk', 'hUsed', 'hLeft'];
/** Клітинка ключа на «Даних графіків» ($B$n). */
function chK_(name) { const i = CH_KEYS.indexOf(name); if (i < 0) throw new Error('chK_: немає «' + name + '»'); return '$B$' + (i + 2); }
/** Ключ з іншого аркуша. */
function chD_(name) { return chRef_(chK_(name)); }
/** {{ім'я}} → $B$n. */
function chFill_(f) { return f.replace(/\{\{(\w+)\}\}/g, (x, n) => chK_(n)); }

/** Сьогодні ± n днів (формула з {{today}}). */
function chDay_(n) { return n ? '{{today}}' + (n > 0 ? '+' : '-') + Math.abs(n) : '{{today}}'; }

/** Баланс на кінець дня x: з «Днів», якщо x у вікні, інакше з контрольної точки й журналу. */
function chBalAt_(x) {
  return 'IF(' + x + '>={{win}},N(INDEX(\'' + DAYS_SHEET + '\'!$E$' + DAYS_FIRST + ':$E$' + (DAYS_FIRST + DAYS_ROWS - 1) + ',' + x + '-{{win}}+1)),' + cardAt_(x) + ')';
}
/** Кредитний ліміт на день x. */
function chLimAt_(x) {
  const cpA = CP_RANGE('A'), cpC = CP_RANGE('C');
  return 'IF(' + x + '>={{win}},N(INDEX(\'' + DAYS_SHEET + '\'!$G$' + DAYS_FIRST + ':$G$' + (DAYS_FIRST + DAYS_ROWS - 1) + ',' + x + '-{{win}}+1)),IFERROR(SUMIFS(' +
    cpC + ',' + cpA + ',MAXIFS(' + cpA + ',' + cpA + ',"<="&' + x + ',' + cpC + ',"<>")),0))';
}

/** Формули блоків: [[рядок, колонка, 2D]]. */
function chartsDataCells_() {
  const A = J_RANGE(JC.date), O = J_RANGE(JC.fact), C = J_RANGE(JC.cat), P = J_RANGE(JC.holeEff), T = J_RANGE(JC.out), Q = J_RANGE(JC.bmonth),
    V = J_RANGE(JC.fromMine), W = J_RANGE(JC.fromLimit), X = J_RANGE(JC.limitWhy), noFop = ',' + J_RANGE(JC.acctKind) + ',"<>' + K_FOP + '"';
  const wThis = A + ',">="&{{m0}},' + A + ',"<="&{{today}}', wLast = A + ',">="&{{pm0}},' + A + ',"<="&{{pmEnd}}', wQ = A + ',">="&{{pq0}},' + A + ',"<="&{{pqEnd}}';
  const wK = k => A + ',">="&EDATE({{m0}},' + k + '),' + A + ',"<"&EDATE({{m0}},' + (k + 1) + ')', w6 = A + ',">="&EDATE({{m0}},-5),' + A + ',"<"&EDATE({{m0}},1)';
  const label = k => '=CHOOSE(MONTH(EDATE({{m0}},' + k + ')),"' + CH_MON.join('","') + '")&" "&RIGHT(YEAR(EDATE({{m0}},' + k + ')),2)';
  const dict = (c, r) => "'" + DICT_SHEET + "'!$" + c + '$' + r, days = (c, x) => 'N(INDEX(\'' + DAYS_SHEET + '\'!$' + c + '$' + DAYS_FIRST + ':$' + c + '$' +
    (DAYS_FIRST + DAYS_ROWS - 1) + ',' + x + '-{{win}}+1))';
  const L = (b, i) => colStr_(CHD[b] + i), col = (b, i, n) => chCol_(CHD[b] + i, 2, 1 + n);
  const out = [];
  // A:B ключі
  out.push([2, 2, chartsKeys_().map(k => [k[2]])]);
  // D:L категорії
  const cats = [], cLast = 1 + CH_CATS;
  for (let i = 0; i < CH_CATS; i++) {
    const r = i + 2, d = D_FIRST + i, nm = '$D' + r, sum = w => '=IF(' + nm + '="","",SUMIFS(' + O + ',' + C + ',' + nm + ',' + w + noFop + '))';
    cats.push(['=' + dict('A', d) + '&""', '=' + dict('B', d) + '&""', '=' + dict('C', d) + '&""', sum(wThis), sum(wLast), sum(wQ),
      '=IF(' + nm + '="","",$I' + r + '/3)', '=IF(' + nm + '="","",$G' + r + '-$J' + r + ')',
      '=IF(AND($E' + r + '="' + T_SPEND + '",N($K' + r + ')>0.005),$K' + r + '+ROW()/1E9,"")']);
  }
  out.push([2, CHD.cat, cats]);
  // N:S «Що зросло»
  const up = [], key = '$L$2:$L$' + cLast, at = (c, r) => 'INDEX($' + c + '$2:$' + c + '$' + cLast + ',$N' + r + ')';
  for (let k = 1; k <= CH_UP; k++) {
    const r = k + 1, no = '$N' + r + '=""';
    up.push(['=IFERROR(MATCH(LARGE(' + key + ',' + k + '),' + key + ',0),"")',
      '=IF(' + no + ',' + (k === 1 ? '"Поки що жодна категорія не перевищила середнє"' : '""') + ',' + at('D', r) + ')',
      '=IF(' + no + ',"",' + at('G', r) + ')', '=IF(' + no + ',"",' + at('J', r) + ')', '=IF(' + no + ',"",' + at('K', r) + ')',
      '=IF(' + no + ',"",IF(N($Q' + r + ')=0,"нова",$P' + r + '/$Q' + r + '-1))']);
  }
  out.push([2, CHD.up, up]);
  // V:Z групи (назви в U пише скрипт), рядок 14: разом
  const grp = [], gL = 1 + CH_GROUPS, cr = c => '$' + c + '$2:$' + c + '$' + cLast;
  for (let i = 0; i < CH_GROUPS; i++) {
    const r = i + 2, g = '$U' + r, s = c => '=IF(' + g + '="","",SUMIFS(' + cr(c) + ',' + cr('F') + ',' + g + ',' + cr('E') + ',"' + T_SPEND + '"))';
    grp.push([s('G'), s('H'), s('I'), '=IF(' + g + '="","",$X' + r + '/3)', '=IF(' + g + '="","",N($V' + r + ')+ROW()/1E9)']);
  }
  out.push([2, CHD.grp + 1, grp], [gL + 1, CHD.grp + 1, [['V', 'W', 'X', 'Y'].map(c => '=SUM(' + c + '2:' + c + gL + ')')]]);
  // AC:AH групи за спаданням цього місяця
  const gs = [], gk = '$Z$2:$Z$' + gL;
  for (let k = 1; k <= CH_GROUPS; k++) {
    const r = k + 1, no = '$AC' + r + '=""', ix = c => '=IF(' + no + ',"",INDEX($' + c + '$2:$' + c + '$' + gL + ',$AC' + r + '))';
    gs.push(['=IFERROR(MATCH(LARGE(' + gk + ',' + k + '),' + gk + ',0),"")', ix('U'), ix('V'), ix('W'), ix('X'), ix('Y')]);
  }
  out.push([2, CHD.grpS, gs]);
  // AK:AN типи (назви в AJ: значення)
  const typ = [];
  [0, 1, 2, 3, -1, 4].forEach((t, i) => {
    const r = i + 2, own = { G: 'AK', H: 'AL', I: 'AM' }, s = c => t < 0 ? '=SUM(' + own[c] + '2:' + own[c] + '5)' : '=SUMIFS(' + cr(c) + ',' + cr('E') + ',$AJ' + r + ')';
    typ.push([s('G'), s('H'), s('I'), '=$AM' + r + '/3']);
  });
  out.push([2, CHD.typ + 1, typ]);
  // AQ:AU залишок за 60 днів
  const b60 = [];
  for (let k = 0; k < 60; k++) {
    const r = k + 2, x = '$AQ' + r;
    b60.push(['=' + chDay_(k - 59), '=' + chBalAt_(x), '=MAX(0,$AR' + r + ')', '=MIN(0,$AR' + r + ')', '=IF(N(' + chLimAt_(x) + ')=0,NA(),-' + chLimAt_(x) + ')']);
  }
  out.push([2, CHD.b60, b60]);
  // AX:BA цей місяць по днях
  const pace = [];
  for (let k = 1; k <= 31; k++) {
    const no = '$AX' + (k + 1) + '=""';
    pace.push(['=IF(' + k + '>{{daysIn}},"",' + k + ')', '=IF(' + no + ',NA(),IF(' + k + '>{{daysPassed}},NA(),SUMIFS(' + T + ',' + Q + ',{{m0}},' + A + ',"<="&({{m0}}+' + (k - 1) + '))))',
      '=IF(' + no + ',NA(),{{mineMonth}}*' + k + '/{{daysIn}})', '=IF(' + no + ',NA(),{{mineMonth}})']);
  }
  out.push([2, CHD.pace, pace]);
  // BC:BF залишок: факт і прогноз
  const fc = [];
  for (let k = 0; k < 92; k++) {
    const r = k + 2, x = '$BC' + r, no = x + '=""', d = chDay_(k - 30);
    fc.push(['=IF(' + d + '>EOMONTH({{today}},1),"",' + d + ')', '=IF(' + no + ',NA(),IF(' + x + '>{{today}},NA(),' + chBalAt_(x) + '))',
      '=IF(' + no + ',NA(),IF(' + x + '<{{today}},NA(),' + days('L', x) + '))', '=IF(' + no + ',NA(),IF(N(' + chLimAt_(x) + ')=0,NA(),-' + chLimAt_(x) + '))']);
  }
  out.push([2, CHD.fc, fc]);
  // BH:BJ з моїх чи з ліміту; BL:BO борг
  const ml = [], debt = [];
  for (let k = -5; k <= 0; k++) ml.push([label(k), '=SUMIFS(' + V + ',' + wK(k) + ')', '=SUMIFS(' + W + ',' + wK(k) + ')']);
  for (let k = -3; k <= 6; k++) {
    debt.push([label(k), k < 0 ? '=MAX(0,-(' + cardAt_('EOMONTH({{today}},' + k + ')') + '))' : k === 0 ? '={{debt}}' : '=NA()',
      k === 0 || k === 1 ? '=MAX(0,-N(' + dayAt_('L', 'EOMONTH({{today}},' + k + ')') + '))' : '=NA()',
      k >= -1 ? '=IF({{goal}}<=0,NA(),MAX(0,{{debtBase}}-{{goal}}*' + (k + 1) + '))' : '=NA()']);
  }
  out.push([2, CHD.ml, ml], [2, CHD.debt, debt]);
  // BQ:CA «чорні діри» (BW: ліміт з «Charts», пише chartsHoleLimits_), рядок 10: разом
  const hole = [], hL = 1 + CH_HOLES, hc = i => L('hole', i);
  const status = p => '=IF(' + p + '="","' + H_NONE + '",IF(' + p + '>1,"' + H_OVER + '",IF(' + p + '>=0.8,"' + H_WARN + '","' + H_OK + '")))';
  for (let i = 0; i < CH_HOLES; i++) {
    const r = i + 2, nm = '$BQ' + r, sum = w => '=IF(' + nm + '="","",SUMIFS(' + O + ',' + P + ',' + nm + ',' + w + noFop + '))';
    hole.push(['=' + dict(DC.holes, D_HOLES_FIRST + i) + '&""', sum(wThis), sum(wLast), sum(wQ), '=IF(' + nm + '="","",$BT' + r + '/3)',
      '=IF(' + nm + '="","",COUNTIFS(' + P + ',' + nm + ',' + O + ',">0",' + wThis + noFop + '))']);
  }
  out.push([2, CHD.hole, hole]);
  const hole2 = [];
  for (let i = 0; i < CH_HOLES; i++) {
    const r = i + 2, nm = '$BQ' + r;
    hole2.push(['=IF(' + nm + '="","",IF(N($BW' + r + ')>0,$BW' + r + ',IF(N($BU' + r + ')>0,$BU' + r + ',"")))', '=IF(OR(' + nm + '="",$BX' + r + '=""),"",$BR' + r + '/$BX' + r + ')',
      '=IF(' + nm + '="","",' + status('$BY' + r).slice(1) + ')', '=IF(' + nm + '="","",N($BR' + r + ')+ROW()/1E9)']);
  }
  out.push([2, CHD.hole + 7, hole2]);
  const tot = c => '=SUM(' + c + '2:' + c + hL + ')', T10 = hL + 1;
  out.push([T10, CHD.hole + 1, [[tot('BR'), tot('BS'), tot('BT'), tot('BU'), tot('BV'), '=IF(SUM(BW2:BW' + hL + ')>0,SUM(BW2:BW' + hL + '),"")',
    '=IF(N(BW' + T10 + ')>0,BW' + T10 + ',IF(N(BU' + T10 + ')>0,BU' + T10 + ',""))',
    '=IF(N(BW' + T10 + ')>0,SUMIFS(BR2:BR' + hL + ',BW2:BW' + hL + ',">0")/BW' + T10 + ',IF(N(BU' + T10 + ')>0,BR' + T10 + '/BU' + T10 + ',""))', status('BY' + T10)]]]);
  // CC:CK за 6 місяців (рядок 1: назви «чорних дір»)
  const h6h = [], h6 = [];
  for (let j = 0; j < CH_HOLES; j++) h6h.push('=$BQ$' + (j + 2) + '&""');
  for (let k = -5; k <= 0; k++) {
    const row = [label(k)];
    for (let j = 0; j < CH_HOLES; j++) { const h = L('h6', j + 1); row.push('=IF(' + h + '$1="",NA(),SUMIFS(' + O + ',' + P + ',' + h + '$1,' + wK(k) + noFop + '))'); }
    h6.push(row);
  }
  out.push([1, CHD.h6 + 1, [h6h]], [2, CHD.h6, h6]);
  // CM:CQ причини, рядок 23: разом; CS:CX за спаданням
  const why = [], wL = 1 + CH_WHY;
  for (let i = 0; i < CH_WHY; i++) {
    const r = i + 2, nm = '$CM' + r, no = nm + '=""';
    why.push([i < CH_WHY - 1 ? '=' + dict(DC.reasons, D_REASONS_FIRST + i) + '&""' : '="' + CH_WHY_NONE + '"',
      '=IF(' + no + ',"",SUMIFS(' + W + ',' + X + ',' + nm + ',' + A + ',">="&{{m0}},' + A + ',"<"&EDATE({{m0}},1)))',
      '=IF(' + no + ',"",COUNTIFS(' + X + ',' + nm + ',' + W + ',">0",' + A + ',">="&{{m0}},' + A + ',"<"&EDATE({{m0}},1)))',
      '=IF(' + no + ',"",SUMIFS(' + W + ',' + X + ',' + nm + ',' + w6 + '))',
      '=IF(OR(' + no + ',N($CP' + r + ')+N($CN' + r + ')<=0),"",N($CP' + r + ')+N($CN' + r + ')/1E6+ROW()/1E9)']);
  }
  out.push([2, CHD.why, why], [wL + 1, CHD.why + 1, [['CN', 'CO', 'CP'].map(c => '=SUM(' + c + '2:' + c + wL + ')')]]);
  const ws = [], wk = '$CQ$2:$CQ$' + wL;
  for (let k = 1; k <= CH_WHY; k++) {
    const r = k + 1, no = '$CS' + r + '=""', ix = c => '=IF(' + no + ',"",INDEX($' + c + '$2:$' + c + '$' + wL + ',$CS' + r + '))';
    ws.push(['=IFERROR(MATCH(LARGE(' + wk + ',' + k + '),' + wk + ',0),"")', ix('CM'), ix('CN'), ix('CO'), ix('CP'),
      '=IF(OR(' + no + ',N($CP$' + (wL + 1) + ')=0),"",$CW' + r + '/$CP$' + (wL + 1) + ')']);
  }
  out.push([2, CHD.whyS, ws]);
  return out.map(b => [b[0], b[1], b[2].map(row => row.map(chFill_))]);
}

/** Заголовки блоків (рядок 1) і сталі підписи. */
function chartsHeads_() {
  return [[1, ['Показник', 'Значення']],
    [CHD.cat, ['Категорія', 'Тип', 'Група', 'Цей місяць (з 1-го по сьогодні)', 'Минулий місяць', 'Попередній квартал', 'Сер. місяць кварталу', 'Більше за середнє на', 'Ключ рейтингу']],
    [CHD.up, ['Рядок', 'Категорія', 'Цей місяць', 'Сер. місяць кварталу', 'Більше на', 'Зміна']],
    [CHD.grp, ['Група', 'Цей місяць', 'Минулий місяць', 'Квартал', 'Сер. місяць кварталу', 'Ключ']],
    [CHD.grpS, ['Рядок', 'Група', 'Цей місяць', 'Минулий місяць', 'Квартал', 'Сер. місяць кварталу']],
    [CHD.typ, ['Тип', 'Цей місяць', 'Минулий місяць', 'Квартал', 'Сер. місяць кварталу']],
    [CHD.b60, ['Дата', 'Залишок', 'Залишок у плюсі', 'Залишок у мінусі (кредит)', 'Межа кредитного ліміту']],
    [CHD.pace, ['День', 'Витрачено наростом', 'Рівний темп', 'Мої гроші місяця (факт + план)']],
    [CHD.fc, ['Дата', 'Залишок (факт)', 'Прогноз', 'Межа кредитного ліміту']],
    [CHD.ml, ['Місяць', 'З моїх', 'З ліміту']],
    [CHD.debt, ['Місяць', 'Борг (факт)', 'Борг (прогноз)', 'Орієнтир погашення']],
    [CHD.hole, ['«Чорна діра»', 'Цей місяць', 'Минулий місяць', 'Квартал', 'Сер. місяць кварталу', 'Операцій', 'Ліміт на місяць', 'База', '% від бази', 'Статус', 'Ключ']],
    [CHD.h6, ['Місяць']],
    [CHD.why, ['Причина', 'Цей місяць', 'Операцій', 'За 6 місяців', 'Ключ']],
    [CHD.whyS, ['Рядок', 'Причина', 'Цей місяць', 'Операцій', 'За 6 місяців', 'Частка']]];
}

/** Групи витрат з «Довідника» (тип «Витрати»), без повторів, у порядку появи. */
function chartsGroups_(ss) {
  const d = ss.getSheetByName(DICT_SHEET), out = [];
  if (!d) return out;
  d.getRange(D_FIRST, 1, CH_CATS, 3).getValues().forEach(r => {
    const g = String(r[2]).trim();
    if (g && String(r[1]).trim() === T_SPEND && out.indexOf(g) < 0 && out.length < CH_GROUPS) out.push(g);
  });
  return out;
}

/** Скільки «чорних дір» показувати (перші CH_HOLES рядків «Довідника», колонка I): за останньою непорожньою. */
function chartsHolesN_(ss) {
  const d = ss.getSheetByName(DICT_SHEET);
  if (!d) return 0;
  const v = d.getRange(D_HOLES_FIRST, colNum_(DC.holes), CH_HOLES, 1).getValues();
  let n = 0;
  v.forEach((r, i) => { if (String(r[0]).trim()) n = i + 1; });
  return n;
}

/** Назви груп на «Даних графіків» (U2:U13). true: змінились. */
function chartsGroupsWrite_(ss) {
  const sh = ss.getSheetByName(CH_SHEET);
  if (!sh) return false;
  const g = chartsGroups_(ss), want = [];
  for (let i = 0; i < CH_GROUPS; i++) want.push([g[i] || '']);
  const rng = sh.getRange(2, CHD.grp, CH_GROUPS, 1), have = rng.getValues().map(r => [String(r[0])]);
  if (JSON.stringify(have) === JSON.stringify(want)) return false;
  rng.setValues(want);
  return true;
}

/** «Дані графіків» старої будови (до версії 9): інші заголовки. */
function chartsDataOld_(ss) {
  const sh = ss.getSheetByName(CH_SHEET);
  return !!sh && String(sh.getRange(1, 1).getValue()) !== 'Показник';
}

function buildChartsData_(ss) {
  const sh = ensureSize_(sheetOf_(ss, CH_SHEET), CH_ROWS, CH_COLS);
  if (chartsDataOld_(ss)) sh.clear();
  chartsHeads_().forEach(h => sh.getRange(1, h[0], 1, h[1].length).setValues([h[1]]).setFontWeight('bold'));
  sh.getRange(2, 1, CH_KEYS.length, 1).setValues(chartsKeys_().map(k => [k[1]]));
  sh.getRange(2, CHD.typ, 6, 1).setValues([[T_SPEND], [T_DEBT], [T_SAVE], [T_TAX], [CH_TOTAL_OUT], [T_IN]]);
  sh.getRange(2 + CH_GROUPS, CHD.grp).setValue('Усі витрати');
  sh.getRange(2 + CH_HOLES, CHD.hole).setValue(CH_HOLES_ALL);
  sh.getRange(2 + CH_WHY, CHD.why).setValue(CH_WHY_ALL);
  chartsDataCells_().forEach(b => setF_(sh.getRange(b[0], b[1], b[2].length, b[2][0].length), b[2]));
  chartsHoleLimits_(ss);
  chartsGroupsWrite_(ss);
  chartsKeys_().forEach((k, i) => { if (k[3]) sh.getRange(i + 2, 2).setNumberFormat(k[3]); });
  sh.getRange(2, CHD.cat + 3, CH_ROWS - 1, CH_COLS - CHD.cat - 2).setNumberFormat('#,##0');
  [CHD.b60, CHD.fc].forEach(c => sh.getRange(2, c, CH_ROWS - 1, 1).setNumberFormat('dd.mm'));
  [CHD.up + 5, CHD.hole + 8, CHD.whyS + 5].forEach(c => sh.getRange(2, c, CH_ROWS - 1, 1).setNumberFormat('0%'));
  sh.getRange(1, 1).setNote('Службовий аркуш для «' + DASH_SHEET + '» і «' + MOB_SHEET + '»: лише формули, правити не треба.');
  sh.setFrozenRows(1);
  sheetShow_(ss, CH_SHEET, false);
  return sh;
}

/** BW2:BW9: ліміти «чорних дір» з «Charts» (колонка «Ліміт на місяць ✎»); без «Charts»: порожньо. */
function chartsHoleLimits_(ss) {
  const sh = ss.getSheetByName(CH_SHEET);
  if (!sh) return;
  const at = dashHoleRow_(ss), f = [];
  for (let i = 0; i < CH_HOLES; i++) {
    const c = at ? "'" + DASH_SHEET + "'!$" + colStr_(DASH_LIMIT_COL) + '$' + (at + i) : '';
    f.push([c ? '=IF(N(' + c + ')>0,' + c + ',"")' : '=""']);
  }
  setF_(sh.getRange(2, CHD.hole + 6, CH_HOLES, 1), f);
}

/** З mbsTick раз на день: назви груп і кількість «чорних дір» (могли змінитись у «Довіднику»); змінились: графіки заново. */
function chartsIfDue_(ss) {
  if (!ss.getSheetByName(CH_SHEET)) return null;
  const dp = PropertiesService.getDocumentProperties(), today = todayKey_(ss);
  if (dp.getProperty('ch_day') === today) return null;
  const was = dp.getProperty('ch_shape'), shape = chartsShape_(ss);
  const r = chartsGroupsWrite_(ss) || was !== shape;
  if (was !== shape && ss.getSheetByName(DASH_SHEET)) dashRefresh_(ss);
  dp.setProperty('ch_day', today);
  return r;
}

/** «Форма» графіків: скільки груп і «чорних дір» (від неї залежать діапазони графіків). */
function chartsShape_(ss) { return chartsGroups_(ss).length + '/' + chartsHolesN_(ss); }
