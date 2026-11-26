/** 27_Charts.gs: Графіки: сітка на «Головному», стовпчик на «📱 Телефоні», дані для них на схованому аркуші «Дані графіків». */

/*
 * «Дані графіків»: лише формули (журнал, «Дні», «Довідник»), правити не треба; аркуш завжди схований.
 *   A:E  на картці по днях: факт (до сьогодні), за планом і «як зараз» (від сьогодні), межа ліміту
 *   G:J  цей місяць по днях: витрачено з моїх наростом, рівний темп, «як зараз»
 *   L:M  цей місяць за типами        O:P  «чорні діри» за 6 місяців     R:S  «чорні діри» цього місяця
 *   U:W  з моїх чи з ліміту, 6 міс.   Y:AB борг за лімітом: факт, прогноз, орієнтир (від боргу на кінець минулого місяця
 *        мінус «Орієнтир погашення» щомісяця; база в Z10)
 *   AD:AF групи: цей місяць проти середнього за 3 (назви груп пише скрипт)   AH:AM причини «з ліміту», топ-8
 *   BA:BD категорії з групами: допоміжне для груп
 * Порожні точки: NA(), щоб лінії мали розрив, а не падали в нуль.
 * Свої графіки скрипт впізнає за назвою (CH_TITLES) і перебудовує; ваші графіки не чіпає.
 * Модуль «💳 Кредитний ліміт» вимкнено: без межі ліміту, «з ліміту», боргу й причин.
 */

const CH_SHEET = 'Дані графіків', PHONE_SHEET = '📱 Телефон';
const CH_ROWS = 94, CH_COLS = 56;                                          // A…BD
const CH_MAIN_ROW = 18, CH_MAIN_STEP = 16, CH_W = 490, CH_H = 300;         // «Головне»: дві колонки графіків від рядка 18
const PHONE_CH_ROW = 21, PHONE_STEP = 14, PHONE_W = 340, PHONE_H = 270;
const CH_GROUPS = 12;
const CH_MON = ['січ', 'лют', 'бер', 'квіт', 'трав', 'черв', 'лип', 'серп', 'вер', 'жовт', 'лист', 'груд'];
/** Кольори перевірено на розрізнення (зокрема для дальтонізму); сірий: лише для орієнтирів (план, темп, середнє). */
const CH_C = { blue: '#1a73e8', green: '#188038', violet: '#8430ce', red: '#c5221f', gray: '#80868b' };
const CH_LEGACY = ['📈 На картці по днях: факт, за планом, як зараз', '📊 Витрати по місяцях: план і факт'];

/** Формули «Даних графіків»: [[рядок, колонка, 2D]]. */
function chartsDataCells_() {
  const A = J_RANGE(JC.date), O = J_RANGE(JC.fact), I = J_RANGE(JC.type), P = J_RANGE(JC.holeEff), Q = J_RANGE(JC.bmonth), S = J_RANGE(JC.mineIn),
    V = J_RANGE(JC.fromMine), W = J_RANGE(JC.fromLimit), X = J_RANGE(JC.limitWhy), C = J_RANGE(JC.cat);
  const M0 = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)', EOM = 'EOMONTH(TODAY(),0)', noFop = ',' + J_RANGE(JC.acctKind) + ',"<>' + K_FOP + '"';
  const ms = k => 'EDATE(' + M0 + ',' + k + ')';
  const inM = k => A + ',">="&' + ms(k) + ',' + A + ',"<"&' + ms(k + 1);
  const label = k => '=CHOOSE(MONTH(' + ms(k) + '),"' + CH_MON.join('","') + '")&" "&RIGHT(YEAR(' + ms(k) + '),2)';
  const day = (c, r) => "'" + DAYS_SHEET + "'!$" + c + '$' + r, dict = (c, r) => "'" + DICT_SHEET + "'!$" + c + '$' + r;
  const out = [];
  // A:E на картці по днях
  const bal = [];
  for (let k = 0; k < DAYS_ROWS; k++) {
    const r = 2 + k, dr = DAYS_FIRST + k, a = '$A' + r;
    bal.push(['=IF(' + day('A', dr) + '="","",' + day('A', dr) + ')',
      '=IF(' + a + '="",NA(),IF(' + a + '<=TODAY(),N(' + day('E', dr) + '),NA()))',
      '=IF(' + a + '="",NA(),IF(' + a + '>=TODAY(),N(' + day('L', dr) + '),NA()))',
      '=IF(' + a + '="",NA(),IF(' + a + '>=TODAY(),N(' + day('M', dr) + '),NA()))',
      '=IF(' + a + '="",NA(),-N(' + day('G', dr) + '))']);
  }
  out.push([2, 1, bal]);
  // G:J цей місяць по днях
  const pace = [];
  for (let k = 1; k <= 31; k++) {
    const g = '$G' + (k + 1), spent = t => 'SUMIFS(' + V + ',' + Q + ',' + M0 + ',' + A + ',"<="&' + t + ')';
    pace.push(['=IF(' + k + '>DAY(' + EOM + '),"","' + k + '")',
      '=IF(' + g + '="",NA(),IF(' + k + '>DAY(TODAY()),NA(),' + spent('(' + M0 + '+' + (k - 1) + ')') + '))',
      '=IF(' + g + '="",NA(),SUMIFS(' + S + ',' + Q + ',' + M0 + ')*' + k + '/DAY(' + EOM + '))',
      '=IF(' + g + '="",NA(),IF(' + k + '<DAY(TODAY()),NA(),' + spent('TODAY()') + '/DAY(TODAY())*' + k + '))']);
  }
  out.push([2, 7, pace]);
  // L:M цей місяць за типами
  out.push([2, 13, [T_SPEND, T_SAVE, T_TAX, T_DEBT].map((t, i) => ['=SUMIFS(' + O + ',' + I + ',$L' + (i + 2) + ',' + inM(0) + noFop + ')'])]);
  // O:P «чорні діри» за 6 місяців; U:W з моїх чи з ліміту
  const h6 = [], l6 = [];
  for (let k = -5; k <= 0; k++) {
    h6.push([label(k), '=SUMIFS(' + O + ',' + P + ',"?*",' + inM(k) + noFop + ')']);
    l6.push([label(k), '=SUMIFS(' + V + ',' + inM(k) + ')', '=SUMIFS(' + W + ',' + inM(k) + ')']);
  }
  out.push([2, 15, h6], [2, 21, l6]);
  // R:S «чорні діри» цього місяця
  const hn = [];
  for (let i = 0; i <= D_HOLES_LAST - D_HOLES_FIRST; i++) {
    const r = i + 2;
    hn.push(['=' + dict('I', D_HOLES_FIRST + i) + '&""', '=IF($R' + r + '="","",SUMIFS(' + O + ',' + P + ',$R' + r + ',' + inM(0) + noFop + '))']);
  }
  out.push([2, 18, hn]);
  // Y:AB борг за лімітом: 5 місяців тому … наступний
  const debt = [];
  for (let k = -5; k <= 1; k++) {
    debt.push([label(k),
      k <= 0 ? '=MAX(0,-(' + cardAt_('MIN(EOMONTH(TODAY(),' + k + '),TODAY())') + '))' : '=NA()',
      k >= 0 ? '=MAX(0,-N(' + dayAt_('L', 'EOMONTH(TODAY(),' + k + ')') + '))' : '=NA()',
      k >= -1 ? '=MAX(0,$Z$10-N(' + SET_REF('goal') + ')*' + (k + 1) + ')' : '=NA()']);
  }
  out.push([2, 25, debt], [10, 26, [['=MAX(0,-(' + cardAt_('EOMONTH(TODAY(),-1)') + '))']]]);
  // AE:AF групи (назви в AD пише скрипт)
  const grp = [];
  for (let i = 0; i < CH_GROUPS; i++) {
    const r = i + 2, g = '$AD' + r;
    grp.push(['=IF(' + g + '="","",SUMIFS($BC$2:$BC$81,$BB$2:$BB$81,' + g + '))', '=IF(' + g + '="","",SUMIFS($BD$2:$BD$81,$BB$2:$BB$81,' + g + ')/3)']);
  }
  out.push([2, 31, grp]);
  // AH:AJ причини «з ліміту» за 6 місяців; AL:AM топ-8
  const why = [], nWhy = D_REASONS_LAST - D_REASONS_FIRST + 1;
  for (let i = 0; i <= nWhy; i++) {
    const r = i + 2;
    why.push([i < nWhy ? '=' + dict('K', D_REASONS_FIRST + i) + '&""' : '="без причини"',
      '=IF($AH' + r + '="","",SUMIFS(' + W + ',' + X + ',$AH' + r + ',' + A + ',">="&' + ms(-5) + ',' + A + ',"<"&' + ms(1) + '))',
      '=IF(N($AI' + r + ')>0,$AI' + r + '+ROW()/1E6,"")']);
  }
  out.push([2, 34, why]);
  const lastWhy = nWhy + 2, rk = '$AJ$2:$AJ$' + lastWhy, top = [];
  for (let k = 1; k <= 8; k++) {
    const at = c => 'IFERROR(INDEX($' + c + '$2:$' + c + '$' + lastWhy + ',MATCH(LARGE(' + rk + ',' + k + '),' + rk + ',0)),"")';
    top.push(['=' + at('AH'), '=' + at('AI')]);
  }
  out.push([2, 38, top]);
  // BA:BD категорії з групами
  const cats = [];
  for (let i = 0; i <= D_LAST - D_FIRST; i++) {
    const r = i + 2, d = D_FIRST + i, skip = 'OR($BA' + r + '="",' + dict('B', d) + '<>"' + T_SPEND + '")';
    cats.push(['=' + dict('A', d) + '&""', '=' + dict('C', d) + '&""',
      '=IF(' + skip + ',0,SUMIFS(' + O + ',' + C + ',$BA' + r + ',' + inM(0) + noFop + '))',
      '=IF(' + skip + ',0,SUMIFS(' + O + ',' + C + ',$BA' + r + ',' + A + ',">="&' + ms(-3) + ',' + A + ',"<"&' + M0 + noFop + '))']);
  }
  out.push([2, 53, cats]);
  return out;
}

/** Заголовки блоків (рядок 1): [колонка, підписи]. */
function chartsHeads_() {
  const cr = modOn_('credit');
  return [[1, ['Дата', 'Факт', 'За планом', 'Як зараз', 'Межа ліміту']], [7, ['День', 'Витрачено з моїх', 'Рівний темп', 'Як зараз']], [12, ['Тип', 'Сума']],
    [15, ['Місяць', '«Чорні діри»']], [18, ['«Чорна діра»', 'Сума']], [21, ['Місяць', 'З моїх', cr ? 'З ліміту' : 'Понад мої']],
    [25, ['Місяць', 'Борг (факт)', 'Прогноз', 'Орієнтир']], [30, ['Група', 'Цей місяць', 'Середнє за 3 міс.']], [34, ['Причина', '6 міс.', 'Ранг']],
    [38, ['Причина', 'З ліміту за 6 міс.']], [53, ['Категорія', 'Група', 'Цей місяць', '3 міс. до нього']]];
}

/** Групи витрат з «Довідника» (тип «Витрати»), без повторів, у порядку появи. */
function chartsGroups_(ss) {
  const d = ss.getSheetByName(DICT_SHEET), out = [];
  if (!d) return out;
  d.getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 3).getValues().forEach(r => {
    const g = String(r[2]).trim();
    if (g && String(r[1]).trim() === T_SPEND && out.indexOf(g) < 0 && out.length < CH_GROUPS) out.push(g);
  });
  return out;
}

/** Назви груп на «Даних графіків» (AD2:AD13). true: змінились. */
function chartsGroupsWrite_(ss) {
  const sh = ss.getSheetByName(CH_SHEET);
  if (!sh) return false;
  const g = chartsGroups_(ss), want = [];
  for (let i = 0; i < CH_GROUPS; i++) want.push([g[i] || '']);
  const rng = sh.getRange(2, 30, CH_GROUPS, 1), have = rng.getValues().map(r => [String(r[0])]);
  if (JSON.stringify(have) === JSON.stringify(want)) return false;
  rng.setValues(want);
  return true;
}

function buildChartsData_(ss) {
  const sh = ensureSize_(sheetOf_(ss, CH_SHEET), CH_ROWS, CH_COLS);
  chartsHeads_().forEach(h => sh.getRange(1, h[0], 1, h[1].length).setValues([h[1]]).setFontWeight('bold'));
  sh.getRange(2, 12, 4, 1).setValues([[T_SPEND], [T_SAVE], [T_TAX], [T_DEBT]]);
  sh.getRange(10, 25).setValue('База орієнтира: борг на кінець минулого місяця');
  chartsDataCells_().forEach(b => setF_(sh.getRange(b[0], b[1], b[2].length, b[2][0].length), b[2]));
  chartsGroupsWrite_(ss);
  sh.getRange(2, 1, DAYS_ROWS, 1).setNumberFormat('dd.mm');
  sh.getRange(2, 2, CH_ROWS - 1, CH_COLS - 1).setNumberFormat('#,##0');
  sh.getRange(1, 1).setNote('Службовий аркуш для графіків «' + MAIN_SHEET + '» і «' + PHONE_SHEET + '»: лише формули, правити не треба.');
  sh.setFrozenRows(1);
  sheetShow_(ss, CH_SHEET, false);
  return sh;
}

/** Графіки: [{ id, title, type, ranges (A1 на «Даних графіків» або { sheet, a1 }), colors, pie, stacked, credit, desk }]. */
function chartSpecs_() {
  const cr = modOn_('credit'), T = Charts.ChartType;
  const all = [
    { id: 'bal', title: cr ? '📈 На картці: факт, прогноз і межа ліміту' : '📈 На картці: факт і прогноз', type: T.LINE,
      ranges: ['A1:A' + CH_ROWS, 'B1:B' + CH_ROWS, 'C1:C' + CH_ROWS, 'D1:D' + CH_ROWS].concat(cr ? ['E1:E' + CH_ROWS] : []),
      colors: [CH_C.blue, CH_C.green, CH_C.violet].concat(cr ? [CH_C.red] : []) },
    { id: 'mine', title: '📈 Цей місяць: витрачено з моїх і рівний темп', type: T.LINE, ranges: ['G1:G32', 'H1:H32', 'I1:I32', 'J1:J32'],
      colors: [CH_C.blue, CH_C.gray, CH_C.violet] },
    { id: 'types', title: '🍩 Цей місяць за типами', type: T.PIE, ranges: ['L1:M5'], pie: true, colors: [CH_C.blue, CH_C.green, CH_C.violet, CH_C.red] },
    { id: 'holes6', title: '📊 «Чорні діри» за 6 місяців', type: T.COLUMN, ranges: ['O1:P7'], colors: [CH_C.violet] },
    { id: 'holesNow', title: '🍩 «Чорні діри» цього місяця', type: T.PIE, ranges: ['R1:S21'], pie: true },
    { id: 'limit6', title: '📊 Витрати: з моїх чи з ліміту, 6 місяців', type: T.COLUMN, ranges: ['U1:W7'], stacked: true, colors: [CH_C.blue, CH_C.red], credit: true },
    { id: 'debt', title: '📉 Борг за лімітом і орієнтир погашення', type: T.LINE, ranges: ['Y1:Y8', 'Z1:Z8', 'AA1:AA8', 'AB1:AB8'],
      colors: [CH_C.red, CH_C.violet, CH_C.gray], credit: true },
    { id: 'months', title: '📊 Витрати по місяцях: план і факт', type: T.COLUMN, ranges: [{ sheet: MAIN_SHEET, a1: 'E' + (MAIN_FIRST - 1) + ':G' + (MAIN_FIRST + 11) }],
      colors: [CH_C.gray, CH_C.blue] },
    { id: 'groups', title: '📊 Групи: цей місяць проти середнього за 3', type: T.BAR, ranges: ['AD1:AF' + (CH_GROUPS + 1)], colors: [CH_C.blue, CH_C.gray], desk: true },
    { id: 'reasons', title: '📊 Чому з ліміту: 6 місяців, топ-8', type: T.BAR, ranges: ['AL1:AM9'], colors: [CH_C.red], credit: true, desk: true },
  ];
  return all.filter(s => !s.credit || cr);
}

/** Усі назви наших графіків (обидва варіанти, старі назви): за ними скрипт прибирає свої. */
function chartTitles_() {
  return CH_LEGACY.concat(['📈 На картці: факт, прогноз і межа ліміту', '📈 На картці: факт і прогноз', '📈 Цей місяць: витрачено з моїх і рівний темп',
    '🍩 Цей місяць за типами', '📊 «Чорні діри» за 6 місяців', '🍩 «Чорні діри» цього місяця', '📊 Витрати: з моїх чи з ліміту, 6 місяців',
    '📉 Борг за лімітом і орієнтир погашення', '📊 Групи: цей місяць проти середнього за 3', '📊 Чому з ліміту: 6 місяців, топ-8']);
}

function chartsRemoveOurs_(sh) {
  const t = chartTitles_();
  sh.getCharts().forEach(c => { try { if (t.indexOf(String(c.getOptions().get('title') || '')) >= 0) sh.removeChart(c); } catch (e) { /* ignore */ } });
}

/** Побудувати графіки за specs на аркуші sh. place(i) → [рядок, колонка]. */
function chartsInsert_(ss, sh, specs, place, w, h, phone) {
  const data = ss.getSheetByName(CH_SHEET);
  specs.forEach((s, i) => {
    const p = place(i);
    let b = sh.newChart().setChartType(s.type);
    s.ranges.forEach(r => { b = b.addRange(typeof r === 'string' ? data.getRange(r) : ss.getSheetByName(r.sheet).getRange(r.a1)); });
    b = b.setNumHeaders(1).setPosition(p[0], p[1], 0, 0).setOption('title', s.title).setOption('width', w).setOption('height', h)
      .setOption('legend', { position: s.pie && !phone ? 'right' : 'bottom' });
    if (s.colors) b = b.setOption('colors', s.colors);
    if (s.pie) b = b.setOption('pieHole', 0.5);
    if (s.stacked) b = b.setOption('isStacked', true);
    sh.insertChart(b.build());
  });
}

/** «Головне»: сітка у дві колонки праворуч від таблиці, від рядка 18. */
function chartsMain_(ss) {
  const sh = ss.getSheetByName(MAIN_SHEET);
  if (!sh || !ss.getSheetByName(DAYS_SHEET) || !ss.getSheetByName(DICT_SHEET)) return 0;
  if (!ss.getSheetByName(CH_SHEET)) buildChartsData_(ss);
  chartsRemoveOurs_(sh);
  const specs = chartSpecs_();
  chartsInsert_(ss, sh, specs, i => [CH_MAIN_ROW + Math.floor(i / 2) * CH_MAIN_STEP, MAIN_TABLE_COL + (i % 2) * 5], CH_W, CH_H, false);
  return specs.length;
}

/** «📱 Телефон»: головні цифри, три бюджети, графіки один під одним. */
function phoneRows_() {
  const cr = modOn_('credit');
  return [['💚 Мої сьогодні', 'Доступно моїх сьогодні'], ['📅 Можна на день', 'Можна на день'], ['💳 На картці', 'На картці зараз']]
    .concat(cr ? [['🔴 Борг за лімітом', 'Борг за лімітом зараз']] : [])
    .concat([['🔮 Кінець місяця за планом', 'Прогноз на кінець місяця (за планом)'], ['🐢 Кінець місяця як зараз', 'Як зараз: на картці на кінець місяця'],
      ['🛒 Витрати місяця', 'Витрати цього місяця: факт'], ['🕳 Непомітні місяця', 'Непомітні цього місяця']])
    .concat(cr ? [['💳 Грейс: внести до строку', 'Грейс: внести до строку']] : [])
    .concat([['🏦 Чернеток', 'Чернеток чекають підтвердження']]);
}

function buildPhone_(ss) {
  const sh = ensureSize_(sheetOf_(ss, PHONE_SHEET), PHONE_CH_ROW + 10 * PHONE_STEP, 2);
  const g = l => '=IFERROR(N(INDEX(\'' + MAIN_SHEET + '\'!$B$1:$B$120,MATCH("' + l + '",\'' + MAIN_SHEET + '\'!$A$1:$A$120,0))),"")';
  sh.getRange(1, 1, PHONE_CH_ROW - 1, 2).clearContent();
  sh.getRange(1, 1).setValue('📱 Телефон').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Відкрийте цей аркуш у застосунку Таблиць: цифри й графіки рахуються самі.').setFontColor('#5f6368');
  const rows = phoneRows_();
  sh.getRange(4, 1, rows.length, 1).setValues(rows.map(r => [r[0]]));
  setF_(sh.getRange(4, 2, rows.length, 1), rows.map(r => [g(r[1])]));
  sh.getRange(4, 2, rows.length, 1).setNumberFormat('#,##0').setFontWeight('bold');
  let r = 4 + rows.length + 1;
  if (modOn_('budgets')) {
    const D = c => "'" + DICT_SHEET + "'!$" + c + '$' + D_FIRST + ':$' + c + '$' + D_LAST, T = x => 'TEXT(' + x + ',"#,##0")';
    sh.getRange(r, 1).setValue('💰 Бюджети: найгірші три').setFontWeight('bold');
    const f = [];
    for (let k = 1; k <= 3; k++) {
      const idx = 'MATCH(LARGE(' + D('P') + ',' + k + '),' + D('P') + ',0)', at = c => 'INDEX(' + D(c) + ',' + idx + ')';
      f.push(['=IFERROR(' + at('O') + '&" "&' + at('A') + ',"")', '=IFERROR(' + T(at('M')) + '&" / "&' + T(at('F')) + ',"")']);
    }
    setF_(sh.getRange(r + 1, 1, 3, 2), f);
    sh.getRange(r + 1, 2, 3, 1).setHorizontalAlignment('right');
  }
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 140);
  chartsRemoveOurs_(sh);
  if (!ss.getSheetByName(CH_SHEET)) buildChartsData_(ss);
  const specs = chartSpecs_().filter(s => !s.desk);
  chartsInsert_(ss, sh, specs, i => [PHONE_CH_ROW + i * PHONE_STEP, 1], PHONE_W, PHONE_H, true);
  return sh;
}

/** Після зміни модулів (ліміт, бюджети) чи коду: заголовки, графіки «Головного» й «Телефону». */
function chartsRefresh_(ss) {
  if (!ss.getSheetByName(MAIN_SHEET)) return 0;
  const d = ss.getSheetByName(CH_SHEET);
  if (d) chartsHeads_().forEach(h => d.getRange(1, h[0], 1, h[1].length).setValues([h[1]]));
  const n = chartsMain_(ss);
  if (ss.getSheetByName(PHONE_SHEET)) buildPhone_(ss);
  return n;
}

/** З mbsTick раз на день: назви груп (могли змінитись у «Довіднику»). */
function chartsIfDue_(ss) {
  if (!ss.getSheetByName(CH_SHEET)) return null;
  const dp = PropertiesService.getDocumentProperties(), today = todayKey_(ss);
  if (dp.getProperty('ch_day') === today) return null;
  const r = chartsGroupsWrite_(ss);
  dp.setProperty('ch_day', today);
  return r;
}
