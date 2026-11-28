/** 25_Header.gs: Шапка журналу (рядки 1…3, колонки A…J) і кольоровий рядок заголовків (рядок 4). */

/*
 * Як у вашій таблиці, але в межах видимих колонок A…J (на телефоні не треба гортати вбік):
 *   A:B  💚 Мої зараз · 💳 На картці · ✍️ Писати у рядок № … ↓ (посилання на рядок для вводу)
 *   C:F  три найгірші бюджети: стан (⛔ перевищено, ▲ швидше норми, ✓), смужка (синє: витрачено в межах плану,
 *        червоне: понад план), «витрачено / план · +понад». Модуль «💰 Бюджети» вимкнено: витрати місяця, непомітні, на день.
 *   G:H  наступний місяць: вільно ≈ на день · «як зараз» на кінець наступного місяця · найнижча точка з датою
 *   I:J  грейс: внести до строку · комісія, якщо все переказами · чернетки. Модуль «💳 Ліміт» вимкнено: чернетки, на день, днів.
 *   K    (модуль «🤝 Борги й розстрочки», версія 17) борги місяця ще до сплати · найближчий платіж · ліміт частинами (з «Боргів», рядок 28).
 * Усе формулами (оновлюється одразу після правки). Бюджети рахують службові колонки «Довідника» M…R (сховані), смужку
 * малює SPARKLINE за параметрами T5:U8 (без масивів у формулах: так працює в будь-якому регіоні таблиці).
 * Вмикання й вимикання модулів переписує лише формули шапки; ваші дані не змінюються.
 */

const HDR_COLS = 10;                                   // A…J
const HDR_HELP = { first: 'M', last: 'R', opt: 'T5:U8' };
const HDR_HEAD_BG = '#5b4ea0', HDR_HEAD_FG = '#ffffff';

/** Службові колонки «Довідника» M…R для бюджетів шапки і параметри смужки T5:U8. */
function headerHelpers_(ss) {
  const d = ss.getSheetByName(DICT_SHEET);
  if (!d) return;
  ensureSize_(d, D_LAST, colNum_('U'));
  const J = c => J_RANGE(c), m0 = '(EOMONTH(TODAY(),-1)+1)', m1 = 'EOMONTH(TODAY(),0)';
  const base = r => J(JC.cat) + ',$A' + r + ',' + J(JC.date) + ',">="&' + m0 + ',' + J(JC.date) + ',"<="&' + m1 + ',' + J(JC.acctKind) + ',"<>' + K_FOP + '"';
  const who = r => ',' + J(JC.who) + ',$G' + r;
  const spent = r => 'IF($G' + r + '="",SUMIFS(' + J(JC.fact) + ',' + base(r) + ')+SUMIFS(' + J(JC.sum) + ',' + base(r) + ',' + J(JC.status) + ',"' + STATUS_DRAFT + '"),' +
    'SUMIFS(' + J(JC.fact) + ',' + base(r) + who(r) + ')+SUMIFS(' + J(JC.sum) + ',' + base(r) + who(r) + ',' + J(JC.status) + ',"' + STATUS_DRAFT + '"))';
  const rows = [];
  for (let r = D_FIRST; r <= D_LAST; r++) {
    rows.push([
      '=IF(AND($A' + r + '<>"",$B' + r + '="' + T_SPEND + '",N($F' + r + ')>0),' + spent(r) + ',"")',
      '=IF($M' + r + '="","",N($F' + r + ')*DAY(TODAY())/DAY(EOMONTH(TODAY(),0)))',
      '=IF($M' + r + '="","",IF($M' + r + '>N($F' + r + ')+0.005,"⛔",IF($M' + r + '>$N' + r + '*1.1,"▲","✓")))',
      '=IF($M' + r + '="","",$M' + r + '/N($F' + r + ')+IF($O' + r + '="⛔",2,IF($O' + r + '="▲",1,0))+ROW()/1E6)',
      '=IF($M' + r + '="","",MIN($M' + r + ',N($F' + r + '))/MAX($M' + r + ',N($F' + r + ')))',
      '=IF($M' + r + '="","",MAX(0,$M' + r + '-N($F' + r + '))/MAX($M' + r + ',N($F' + r + ')))',
    ]);
  }
  d.getRange('M4:R4').setValues([['Шапка: витрачено', 'норма', 'стан', 'порядок', 'смужка: в межах', 'смужка: понад']]).setFontColor('#9aa0a6');
  setF_(d.getRange(D_FIRST, colNum_('M'), rows.length, 6), rows);
  d.getRange('T4:U4').setValues([['Шапка: смужка', '']]).setFontColor('#9aa0a6');
  d.getRange(HDR_HELP.opt).setValues([['charttype', 'bar'], ['max', 1], ['color1', '#4a86e8'], ['color2', '#e06666']]);
  d.hideColumns(colNum_('M'), colNum_('U') - colNum_('M') + 1);
}

/** Формули шапки (рядки 1…3 × A…J) з урахуванням модулів. → [[10 формул або значень] × 3] */
function headerFormulas_(ss) {
  const j = ss.getSheetByName(JOURNAL), gid = j ? j.getSheetId() : 0;
  const g = l => 'N(INDEX(\'' + MAIN_SHEET + '\'!$B$1:$B$120,MATCH("' + l + '",\'' + MAIN_SHEET + '\'!$A$1:$A$120,0)))';
  const D = c => "'" + DICT_SHEET + "'!$" + c + '$' + D_FIRST + ':$' + c + '$' + D_LAST;
  const A = '$A$' + J_FIRST + ':$A$' + J_LAST;
  const first = 'MATCH(TRUE,INDEX(' + A + '="",0),0)+' + (J_FIRST - 1);
  const row = '(' + first + '+IF(' + first + '>' + J_FIRST + ',N(INDEX($A:$A,' + first + '-1))>TODAY(),0))';
  const months = '"' + MONTHS_UA.map(m => m[0].toUpperCase() + m.slice(1)).join('","') + '"';
  const T = x => 'TEXT(' + x + ',"#,##0")';
  const out = [[], [], []];
  // A:B
  out[0].push('💚 Мої зараз', '=IFERROR(' + g('Доступно моїх сьогодні') + ',"")');
  out[1].push('💳 На картці', '=IFERROR(' + g('На картці зараз') + ',"")');
  out[2].push('✍️ Писати у рядок', '=IFERROR(HYPERLINK("#gid=' + gid + '&range=A"&' + row + ',"№ "&' + row + '&" ↓"),"")');
  // C:F
  if (modOn_('budgets')) {
    for (let k = 1; k <= 3; k++) {
      const idx = 'MATCH(LARGE(' + D('P') + ',' + k + '),' + D('P') + ',0)', at = c => 'INDEX(' + D(c) + ',' + idx + ')';
      out[k - 1].push(
        '=IFERROR(' + at('O') + '&" "&' + at('A') + ',' + (k === 1 ? '"Бюджети: «' + DICT_SHEET + '», колонка F"' : '""') + ')',
        "=IFERROR(SPARKLINE(INDEX('" + DICT_SHEET + "'!$Q$" + D_FIRST + ':$R$' + D_LAST + ',' + idx + ",0),'" + DICT_SHEET + "'!$T$5:$U$8),\"\")", '',
        '=IFERROR(' + T(at('M')) + '&" / "&' + T(at('F')) + '&IF(' + at('O') + '="⛔"," · +"&' + T(at('M') + '-' + at('F')) + ',IF(' + at('O') + '="▲"," · швидше норми","")),"")');
    }
  } else {
    out[0].push('Витрати місяця', '', '', '=IFERROR(' + T(g('Витрати цього місяця: факт')) + '&" / "&' + T(g('Витрати цього місяця: план')) + ',"")');
    out[1].push('Непомітні', '', '', '=IFERROR(' + T(g('Непомітні цього місяця')) + ',"")');
    out[2].push('Витрачаєте ≈ на день', '', '', '=IFERROR(' + T(g('Витрачаєте ≈ на день')) + ',"")');
  }
  // G:H
  const nx = 'EOMONTH(TODAY(),1)';
  out[0].push('=CHOOSE(MONTH(' + nx + '),' + months + ')&": вільно ≈ "&' + T('MAX(0,' + g('Прогноз на кінець наступного місяця (за планом)') + ')/DAY(' + nx + ')') + '&" ₴/день"',
    '=IFERROR(' + g('Прогноз на кінець наступного місяця (за планом)') + ',"")');
  out[1].push('="Як зараз · "&TEXT(' + nx + ',"dd.mm")', '=IFERROR(N(' + dayAt_('M', nx) + '),"")');
  out[2].push('=IFERROR("Мін. · "&TEXT(' + g('Дата найнижчої точки') + ',"dd.mm"),"Мін.")', '=IFERROR(' + g('Найнижча точка за планом') + ',"")');
  // I:J
  if (modOn_('credit')) {
    out[0].push('=IFERROR("💳 Грейс до "&TEXT(' + g('Грейс: строк') + ',"dd.mm"),"💳 Грейс")', '=IFERROR(' + g('Грейс: внести до строку') + ',"")');
    out[1].push('=IFERROR("Комісія "&TEXT(' + g('Грейс: ставка комісії за переказ з ліміту') + '*100,"0.0")&" %","Комісія")', '=IFERROR(' + g('Грейс: комісія, якщо все переказами') + ',"")');
  } else {
    out[0].push('Можна на день', '=IFERROR(' + g('Можна на день') + ',"")');
    out[1].push('Днів до кінця місяця', '=IFERROR(' + g('Днів до кінця місяця') + ',"")');
  }
  out[2].push('🏦 Чернеток', '=IFERROR(' + g('Чернеток чекають підтвердження') + ',"")');
  return out;
}

/** Підказка на «Журналі» (A3): як писати рядок. */
function headerNote_() {
  return 'Пишіть у рядки нижче жовтої лінії: дата, ' + J_SUM_HINT + ', категорія зі списку, статус. Решту рахують формули. ' +
    'Майбутнє (плани) скрипт ставить угору, минуле вниз; сховані колонки праворуч службові.';
}

/** Побудувати або переписати шапку журналу (рядки 1…3) і рядок заголовків. Дані журналу не чіпає. */
function journalHeader_(ss) {
  const j = ss.getSheetByName(JOURNAL);
  if (!j) return;
  headerHelpers_(ss);
  const rows = headerFormulas_(ss), money = '#,##0" ₴"';
  ['D1:E1', 'D2:E2', 'D3:E3'].forEach(a => { try { j.getRange(a).breakApart(); } catch (e) { /* не було */ } });
  j.getRange(1, 1, 3, HDR_COLS).clearContent();
  setF_(j.getRange(1, 1, 3, HDR_COLS), rows.map(r => r.map(x => (String(x).charAt(0) === '=' ? x : ''))));
  rows.forEach((r, i) => r.forEach((x, k) => { if (x !== '' && String(x).charAt(0) !== '=') j.getRange(i + 1, k + 1).setValue(x); }));
  ['D1:E1', 'D2:E2', 'D3:E3'].forEach(a => j.getRange(a).merge());
  const credit = modOn_('credit');
  j.getRange('A1:J3').setFontSize(10).setVerticalAlignment('middle').setFontWeight('normal').setFontColor('#202124');
  j.getRange('B1:B2').setNumberFormat(money).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right');
  j.getRange('B3').setFontWeight('bold').setHorizontalAlignment('right');
  j.getRange('H1:H3').setNumberFormat(money).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right');
  j.getRange('J1:J3').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right');
  j.getRange('J1:J2').setNumberFormat(credit ? money : '#,##0');
  if (!credit) j.getRange('J1').setNumberFormat(money);
  j.getRange('J3').setNumberFormat('0');
  j.getRange('F1:F3').setHorizontalAlignment('right');
  j.getRange('A1:B1').setBackground('#d9ead3'); j.getRange('A2:B2').setBackground('#f1f3f4'); j.getRange('A3:B3').setBackground('#fff2cc');
  j.getRange('C1:F3').setBackground(null); j.getRange('G1:H3').setBackground('#e8f0fe'); j.getRange('I1:J3').setBackground('#fef7e0');
  j.getRange('A3').setNote(headerNote_());
  headerParts_(ss);
  for (let r = 1; r <= 3; r++) j.setRowHeight(r, 24);
  // рядок заголовків
  j.getRange(J_HEAD, 1, 1, HDR_COLS).setBackground(HDR_HEAD_BG).setFontColor(HDR_HEAD_FG).setFontWeight('bold').setWrap(true)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  headerColors_(j);
  j.setFrozenRows(J_HEAD);
}

/** Колонка K шапки: борги й покупки частинами (модуль «🤝 Борги й розстрочки»); модуль вимкнено: порожньо. */
function headerParts_(ss) {
  const j = ss.getSheetByName(JOURNAL);
  if (!j) return false;
  ensureSize_(j, 3, colNum_('K'));
  const k = j.getRange('K1:K3');
  k.clearContent().setBackground(null);
  if (!modOn_('debts') || !ss.getSheetByName(DEBTS_SHEET)) return false;
  const R = c => '$' + c + '$' + J_FIRST + ':$' + c + '$' + J_LAST, tag = '"' + DEBT_TAG + '*"', notIn = R(JC.cat) + ',"<>' + DEBT_CAT[DEBT_LENT] + '"';
  const m0 = '(EOMONTH(TODAY(),-1)+1)', m1 = 'EOMONTH(TODAY(),0)', T = x => 'TEXT(' + x + ',"#,##0")';
  const next = 'MINIFS(' + R(JC.date) + ',' + R(JC.note) + ',' + tag + ',' + R(JC.state) + ',1,' + R(JC.date) + ',">="&TODAY(),' + notIn + ')';
  const B = c => "'" + DEBTS_SHEET + "'!$" + c + '$' + DEBT_LIMITS[0].row;
  setF_(k, [
    ['="💳 Борги цього місяця: ще "&' + T('SUMIFS(' + R(JC.plan) + ',' + R(JC.note) + ',' + tag + ',' + R(JC.date) + ',">="&' + m0 + ',' + R(JC.date) + ',"<="&' + m1 + ',' + notIn + ')') + '&" ₴"'],
    ['=IF(' + next + '=0,"Найближчого платежу немає","Найближчий "&TEXT(' + next + ',"dd.mm")&": "&' + T('SUMIFS(' + R(JC.plan) + ',' + R(JC.note) + ',' + tag + ',' + R(JC.date) + ',' + next + ',' + notIn + ')') + '&" ₴")'],
    ['=IF(' + B('B') + '="","Ліміт частинами: впишіть на «' + DEBTS_SHEET + '»",IF(LEFT(' + B('F') + ',1)="⚠","⚠️ Ліміт частинами не сходиться з Mono: «' + DEBTS_SHEET + '»",' +
      '"Ліміт частинами: вільно "&' + T(B('E')) + '&" з "&' + T(B('B')) + '))'],
  ]);
  k.setFontSize(10).setVerticalAlignment('middle').setBackground('#f3e8fd').setFontColor('#202124').setWrap(false);
  try { j.setColumnWidth(colNum_('K'), 300); } catch (e) { /* ignore */ }
  return true;
}

/** Кольори шапки: свої правила (діапазони лише в рядках 1…3) замінюються, решта лишається. */
function headerColors_(j) {
  const ours = r => { try { return r.getRanges().every(g => g.getLastRow() <= 3); } catch (e) { return false; } };
  const keep = j.getConditionalFormatRules().filter(r => !ours(r));
  const R = a => j.getRange(a), B = () => SpreadsheetApp.newConditionalFormatRule();
  const rules = [
    B().whenFormulaSatisfied('=$B$1<0').setBackground('#fce8e6').setFontColor('#a50e0e').setRanges([R('A1:B1')]).build(),
    B().whenFormulaSatisfied('=$B$2<0').setBackground('#fce8e6').setFontColor('#a50e0e').setRanges([R('A2:B2')]).build(),
    B().whenTextStartsWith('⛔').setBackground('#f8cfcb').setFontColor('#a50e0e').setRanges([R('C1:C3')]).build(),
    B().whenTextStartsWith('▲').setBackground('#fde3cf').setFontColor('#8a4b08').setRanges([R('C1:C3')]).build(),
    B().whenTextStartsWith('✓').setBackground('#d9f2e3').setFontColor('#146c2e').setRanges([R('C1:C3')]).build(),
    B().whenTextContains(' · +').setFontColor('#c5221f').setRanges([R('F1:F3')]).build(),
    B().whenNumberLessThan(0).setFontColor('#c5221f').setRanges([R('H1:H3')]).build(),
  ];
  j.setConditionalFormatRules(rules.concat(keep));
}

/** Шапка вже нова? (для кроку версії 7) */
function headerIsNew_(ss) {
  const j = ss.getSheetByName(JOURNAL), d = ss.getSheetByName(DICT_SHEET);
  return !!j && !!d && String(j.getRange('A1').getValue()) === '💚 Мої зараз' && d.getMaxColumns() >= colNum_('U') && String(d.getRange('T5').getValue()) === 'charttype';
}

/** З mbsTick раз на день: посилання «№ … ↓» веде на цей аркуш (після копіювання таблиці id аркуша міг змінитись). */
function headerIfDue_(ss) {
  const dp = PropertiesService.getDocumentProperties(), today = todayKey_(ss);
  if (dp.getProperty('hdr_day') === today) return null;
  dp.setProperty('hdr_day', today);
  const j = ss.getSheetByName(JOURNAL);
  if (!j || !headerIsNew_(ss)) return null;
  if (String(j.getRange('B3').getFormula()).indexOf('#gid=' + j.getSheetId() + '&') >= 0) return null;
  journalHeader_(ss);
  return true;
}
