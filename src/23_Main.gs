/** 23_Main.gs: «Головне»: шапка з головними цифрами, графіки; коротка шапка над журналом. */

/*
 * Колонка A: підпис, B: значення (формула), C: що це. Telegram, звіти й «Перевірка» читають значення за підписом
 * (mainByLabel_), тож рядки можна переставляти, а підписи міняти не варто.
 * Звідки: «Дні» (картка, борг, ліміт, «мої», прогнози) і журнал (місяць, чернетки, червоні рядки).
 * Графіки (27_Charts.gs): сітка у дві колонки від рядка 18 праворуч від таблиці; таблиця E3:G15 живить графік «Витрати по місяцях».
 */

const MAIN_FIRST = 4;
const MAIN_TABLE_COL = 5;                                 // E: таблиця для графіка місяців (місяць, план, факт)

function dayAt_(col, x) {
  const last = DAYS_FIRST + DAYS_ROWS - 1;
  return 'IFERROR(INDEX(\'' + DAYS_SHEET + '\'!$' + col + '$' + DAYS_FIRST + ':$' + col + '$' + last + ',MATCH(' + x + ',\'' + DAYS_SHEET + '\'!$A$' +
    DAYS_FIRST + ':$A$' + last + ',0)),"")';
}

/** [підпис, формула, формат, пояснення]. */
function mainRows_() {
  const A = J_RANGE(JC.date), m0 = '(EOMONTH(TODAY(),-1)+1)', m1 = 'EOMONTH(TODAY(),0)';
  const inMonth = A + ',">="&' + m0 + ',' + A + ',"<="&' + m1 + ',' + J_RANGE(JC.acctKind) + ',"<>' + K_FOP + '"';
  const money = '#,##0', date = 'dd.mm.yyyy';
  return [
    ['Сьогодні', '=TODAY()', date, 'дата, на яку рахує таблиця'],
    ['На картці зараз', '=' + dayAt_('E', 'TODAY()'), money, 'мінус: у кредитному ліміті'],
    ['Кредитний ліміт', '=' + dayAt_('G', 'TODAY()'), money, 'з останньої контрольної точки банку'],
    ['Борг за лімітом зараз', '=' + dayAt_('F', 'TODAY()'), money, ''],
    ['Доступно з лімітом', '=' + dayAt_('H', 'TODAY()'), money, ''],
    ['Доступно моїх сьогодні', '=' + dayAt_('I', 'TODAY()'), money, 'мої гроші бюджетного місяця мінус уже витрачене'],
    ['Днів до кінця місяця', '=EOMONTH(TODAY(),0)-TODAY()+1', '0', 'разом із сьогоднішнім'],
    ['Можна на день', '=IF(N(B' + (MAIN_FIRST + 5) + ')>0,B' + (MAIN_FIRST + 5) + '/B' + (MAIN_FIRST + 6) + ',0)', money, 'мої ÷ днів до кінця місяця'],
    ['Прогноз на кінець місяця (за планом)', '=' + dayAt_('L', 'EOMONTH(TODAY(),0)'), money, 'на картці, якщо все піде за планами журналу'],
    ['Прогноз на кінець наступного місяця (за планом)', '=' + dayAt_('L', 'EOMONTH(TODAY(),1)'), money, ''],
    ['Витрачаєте ≈ на день', "='" + DAYS_SHEET + "'!B3", money, 'середнє за 30 днів (тип «Витрати»)'],
    ['Як зараз: на картці на кінець місяця', '=' + dayAt_('M', 'EOMONTH(TODAY(),0)'), money, 'план мінус звичний темп витрат'],
    ['Найнижча точка за планом', "='" + DAYS_SHEET + "'!B2", money, 'у вікні «Днів»'],
    ['Дата найнижчої точки', "='" + DAYS_SHEET + "'!C2", date, ''],
    ['Найнижча точка «як зараз»', "='" + DAYS_SHEET + "'!B4", money, ''],
    ['Дата найнижчої «як зараз»', "='" + DAYS_SHEET + "'!C4", date, ''],
    ['Витрати цього місяця: факт', '=SUMIFS(' + J_RANGE(JC.fact) + ',' + J_RANGE(JC.type) + ',"' + T_SPEND + '",' + inMonth + ')', money, 'тип «Витрати», лише ✅ факт'],
    ['Витрати цього місяця: план', '=SUMIFS(' + D_RANGE('F') + ',' + D_RANGE('B') + ',"' + T_SPEND + '")', money, 'сума планів категорій у «Довіднику»'],
    ['Непомітні цього місяця', '=SUMIFS(' + J_RANGE(JC.fact) + ',' + J_RANGE(JC.holeEff) + ',"?*",' + inMonth + ')', money, '«чорні діри»'],
    ['З ліміту цього місяця', '=SUMIFS(' + J_RANGE(JC.fromLimit) + ',' + inMonth + ')', money, 'витрачено понад мої гроші'],
    ['З ліміту без причини (рядків)', '=COUNTIF(' + J_RANGE(JC.limitWhy) + ',"без причини")', '0', 'вкажіть причину в колонці H'],
    ['Відкладено на погашення (бюджетний місяць)', '=SUMIFS(' + J_RANGE(JC.fact) + ',' + J_RANGE(JC.mineKind) + ',"' + MINE_SET_ASIDE + '",' + J_RANGE(JC.bmonth) + ',' + m0 + ')', money, ''],
    ['Орієнтир погашення на місяць', '=N(' + SET_REF('goal') + ')', money, '«Налаштування»'],
    ['Чернеток чекають підтвердження', '=COUNTIF(' + J_RANGE(JC.status) + ',"' + STATUS_DRAFT + '")', '0', '«🏦 чернетка» не йде в статистику'],
    ['Червоних рядків', '=COUNTIFS(' + J_RANGE(JC.cat) + ',"' + CAT_TRANSFER + '",' + J_RANGE(JC.status) + ',"<>' + STATUS_DRAFT + '")+COUNTIFS(' +
      J_RANGE(JC.sum) + ',"<>",' + J_RANGE(JC.cat) + ',"",' + J_RANGE(JC.status) + ',"<>' + STATUS_DRAFT + '")', '0', 'переказ або без категорії: не в статистиці'],
    ['Невідомих категорій', '=COUNTIF(' + J_RANGE(JC.type) + ',"⚠ невідома категорія")', '0', 'категорії немає в «Довіднику»'],
    ['Планів з минулою датою', '=COUNTIFS(' + J_RANGE(JC.state) + ',1,' + A + ',"<"&TODAY())', '0', 'підтвердіть або перенесіть'],
    ['Остання контрольна точка банку', '=IFERROR(1/(1/MAXIFS(' + CP_RANGE('A') + ',' + CP_RANGE('B') + ',"<>")),"")', date, '«Залишки банку»'],
    ['Межа «близько до ліміту»', '=N(' + SET_REF('margin') + ')', money, 'для тривоги в Telegram'],
  ].concat(mainGraceRows_(), mainDebtRows_()).map((r, i, all) => {
    // {{підпис}} у формулі → клітинка B цього рядка «Головного» (рядки можна переставляти)
    const f = r[1].replace(/\{\{([^}]+)\}\}/g, (x, l) => { const k = all.findIndex(y => y[0] === l); if (k < 0) throw new Error('mainRows_: немає «' + l + '»'); return 'B' + (MAIN_FIRST + k); });
    return [r[0], f, r[2], r[3]];
  });
}

/*
 * Грейс-період Mono (модуль «💳 Кредитний ліміт і грейс»): борг, що виник у місяці, без відсотків до кінця НАСТУПНОГО місяця,
 * якщо його погасити до цього строку. «Внести»: борг на кінець минулого місяця мінус те, що вже надійшло на картку цього місяця;
 * якщо минулий місяць закрито, то борг на кінець цього місяця (як зараз), до кінця наступного.
 * Переказами: «внести» мінус надходження за планом до строку; комісія = переказами × ставка («Налаштування», B10).
 */
const MAIN_GRACE = ['Грейс: ставка комісії за переказ з ліміту', 'Грейс: борг на кінець минулого місяця', 'Грейс: уже надійшло на картку цього місяця',
  'Грейс: лишилось внести до кінця місяця', 'Грейс: борг на кінець місяця (як зараз)', 'Грейс: строк', 'Грейс: внести до строку',
  'Грейс: надійде за планом до строку', 'Грейс: переказами (оцінка)', 'Грейс: комісія (оцінка)', 'Грейс: комісія, якщо все переказами'];
/** Рядки «Головного», які ховає вимкнений модуль «💳 Кредитний ліміт і грейс». */
const MAIN_CREDIT = ['Кредитний ліміт', 'Борг за лімітом зараз', 'Доступно з лімітом', 'З ліміту цього місяця', 'З ліміту без причини (рядків)',
  'Відкладено на погашення (бюджетний місяць)', 'Орієнтир погашення на місяць'].concat(MAIN_GRACE);
const MAIN_DEBTS = ['Борги: я винен (залишок)', 'Борги: мені винні (залишок)'];

function mainGraceRows_() {
  const A = J_RANGE(JC.date), L = J_RANGE(JC.signed), m0 = '(EOMONTH(TODAY(),-1)+1)', money = '#,##0', g = MAIN_GRACE;
  const planIn = 'SUMIFS(' + J_RANGE(JC.plan) + ',' + J_RANGE(JC.type) + ',"' + T_IN + '",' + J_RANGE(JC.acctKind) + ',"' + K_MAIN + '",' + A + ',">"&IF({{' + g[3] +
    '}}>0,TODAY(),EOMONTH(TODAY(),0)),' + A + ',"<="&{{' + g[5] + '}})';
  return [
    [g[0], '=N(' + SET_REF('rate') + ')', '0.0%', '«Налаштування», B10: ставка вашої картки'],
    [g[1], '=MAX(0,N(' + dayAt_('F', 'EOMONTH(TODAY(),-1)') + '))', money, 'без відсотків, якщо погасити до кінця цього місяця'],
    [g[2], '=SUMIFS(' + L + ',' + A + ',">="&' + m0 + ',' + A + ',"<="&TODAY(),' + L + ',">0")', money, 'факти на основну картку'],
    [g[3], '=MAX(0,{{' + g[1] + '}}-{{' + g[2] + '}})', money, ''],
    [g[4], '=MAX(0,-N(' + dayAt_('M', 'EOMONTH(TODAY(),0)') + '))', money, 'строк для нього: кінець наступного місяця'],
    [g[5], '=IF({{' + g[3] + '}}>0,EOMONTH(TODAY(),0),EOMONTH(TODAY(),1))', 'dd.mm.yyyy', 'найближчий'],
    [g[6], '=IF({{' + g[3] + '}}>0,{{' + g[3] + '}},{{' + g[4] + '}})', money, 'щоб не платити відсотки за користування лімітом'],
    [g[7], '=' + planIn, money, 'плани «Надходження» на основну картку'],
    [g[8], '=MAX(0,{{' + g[6] + '}}-{{' + g[7] + '}})', money, 'що не покриють планові надходження'],
    [g[9], '=ROUND({{' + g[8] + '}}*{{' + g[0] + '}},0)', money, ''],
    [g[10], '=ROUND({{' + g[6] + '}}*{{' + g[0] + '}},0)', money, ''],
  ];
}

/** Борги (модуль «🤝 Борги й розстрочки»): залишки з аркуша «Борги». */
function mainDebtRows_() {
  const R = c => "'" + DEBTS_SHEET + "'!$" + c + '$' + DEBT_FIRST + ':$' + c + '$' + DEBT_LAST;
  return [
    [MAIN_DEBTS[0], '=SUMIFS(' + R('J') + ',' + R('B') + ',"' + DEBT_OWE + '")', '#,##0', '«' + DEBTS_SHEET + '»: скільки ще віддати'],
    [MAIN_DEBTS[1], '=SUMIFS(' + R('J') + ',' + R('B') + ',"' + DEBT_LENT + '")', '#,##0', '«' + DEBTS_SHEET + '»: скільки ще мають повернути'],
  ];
}

function buildMain_(ss) {
  const sh = sheetOf_(ss, MAIN_SHEET), rows = mainRows_();
  sh.getRange(1, 1).setValue('🧭 Головне').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Головні цифри на сьогодні. Рахують формули з «Днів» і журналу; нічого вводити не треба.').setFontColor('#5f6368');
  sh.getRange(MAIN_FIRST - 1, 1, 1, 3).setValues([['Показник', 'Значення', 'Що це']]).setFontWeight('bold');
  sh.getRange(MAIN_FIRST, 1, rows.length, 1).setValues(rows.map(r => [r[0]]));
  setF_(sh.getRange(MAIN_FIRST, 2, rows.length, 1), rows.map(r => [r[1]]));
  sh.getRange(MAIN_FIRST, 3, rows.length, 1).setValues(rows.map(r => [r[3]])).setFontColor('#5f6368');
  rows.forEach((r, i) => sh.getRange(MAIN_FIRST + i, 2).setNumberFormat(r[2]));
  // таблиця для графіка «Витрати по місяцях»: рік з «Місяців»
  const spend = monthsSummaryRow_(T_SPEND), t0 = MAIN_TABLE_COL, head = MAIN_FIRST - 1;
  sh.getRange(head, t0, 1, 3).setValues([['Місяць', 'Витрати: план', 'Витрати: факт']]).setFontWeight('bold');
  const tbl = [];
  for (let k = 1; k <= 12; k++) {
    const c = mCols_(k), M = "'" + MONTHS_SHEET + "'!";
    tbl.push(['=DATE(' + M + '$B$2,' + k + ',1)', '=N(' + M + colStr_(c.plan) + spend + ')', '=N(' + M + colStr_(c.fact) + spend + ')']);
  }
  setF_(sh.getRange(head + 1, t0, 12, 3), tbl);
  sh.getRange(head + 1, t0, 12, 1).setNumberFormat('mmm yyyy');
  sh.getRange(head + 1, t0 + 1, 12, 2).setNumberFormat('#,##0');
  mainCharts_(ss);
  journalHeader_(ss);
  sh.setFrozenRows(MAIN_FIRST - 1);
  return sh;
}

/** Графіки на «Головному»: 27_Charts.gs (замінює свої, чужі не чіпає). */
function mainCharts_(ss) { return chartsMain_(ss); }
