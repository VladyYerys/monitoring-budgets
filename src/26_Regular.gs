/** 26_Regular.gs: Аркуш «Регулярні»: правила платежів і надходжень → плани «🔁 план регулярний» у журналі. */

/*
 * Рядок на правило: назва, категорія, сума, день, як часто (щомісяця / щокварталу / щороку), з якої дати, до якої,
 * хто, увімкнено. Скрипт ставить у журнал плани до кінця наступного місяця з коментарем «[Регулярний] Назва: мм.рррр».
 * Місяць уже сплачено, якщо в журналі є:
 *   • факт чи чернетка з таким самим коментарем (ви поставили плану «✅ факт»), або
 *   • факт чи чернетка тієї ж категорії в межах ±5 днів від дати платежу чи плану і ±25 % суми (або з назвою правила в коментарі),
 *     наприклад чернетка з банку. Тоді план цього місяця прибирається, навіть якщо його дата вже минула: платіж не
 *     рахується двічі.
 * Змінили суму чи день: плани оновлюються. Вимкнули чи видалили правило: його плани зникають. Ваші рядки не чіпає.
 * Минулий план без факту лишається («⚠️ прострочено»); перенесли його на майбутню дату: лишається з вашою датою й сумою.
 * Колонка «Стан» (J) пише скрипт: наступний платіж, сплачено цього місяця, прострочено.
 * Приклади (REG_EXAMPLES): підписки, оренда, комуналка, бензин, спорт, садок, страховка, подарунки, зарплата, кредит, розстрочка,
 * борг другу й вам, заощадження, донат. Усі з «Увімкнено: ні», тож планів не ставлять; позначка в примітці колонки A
 * («💡 Приклад…»), підказка в «Стані». Свій платіж: новий рядок або поправте приклад і поставте «так». Додаються один раз.
 * Шапка (рядки 1–3, версія 18): темні плитки в стилі Grafana, як «Charts»: 💰 собі й інвестиції (першою), 🔁 підключено,
 * 💵 лишається з доходу, 📅 цей місяць, ⏭ найближчі 30 днів (графік по днях), 💳 борги (K:L, з модулем). Усе формулами;
 * службові колонки N…U сховані (на місяць, тип, дні, підсумки, параметри графіків). Пояснення: примітки на заголовках (рядок 4).
 */

const REG_SHEET = 'Регулярні';
const REG_HEAD = 4, REG_FIRST = 5, REG_LAST = 54;
const REG_TAG = '[Регулярний] ';
const REG_PER = ['щомісяця', 'щокварталу', 'щороку'];
const REG_CLOSE_DAYS = 5, REG_TOL = 0.25;
const REG_COLS = ['Назва', 'Категорія', 'Сума, ₴', 'День (1–31)', 'Як часто', 'З (перший платіж)', 'До (порожньо: без кінця)', 'Хто', 'Увімкнено', 'Стан (пише скрипт)'];
const REG_EX_MARK = '💡 Приклад';
const REG_HINT = 'Нижче приклади з «Увімкнено: ні»: вони нічого не планують. Свій платіж: новий рядок або поправте приклад (сума, день) і поставте «так». ' +
  'Непотрібні приклади можна видалити.';
/** [назва, категорія, сума, день, як часто, місяць «З» від поточного (для щокварталу й щороку), місяців до «До» (0: без кінця), підказка]. */
const REG_EXAMPLES = [
  ['Netflix', 'Підписки', 350, 12, 'щомісяця', 0, 0, 'підписка: банк спише, факт закриє план сам'],
  ['YouTube Premium', 'Підписки', 150, 3, 'щомісяця', 0, 0, 'кожна підписка окремим рядком: видно, скільки всі разом'],
  ['Хмара для фото', 'Підписки', 100, 20, 'щомісяця', 0, 0, 'дрібні підписки теж: за рік набігає'],
  ['Мобільний і домашній інтернет', 'Інтернет і зв\'язок', 500, 7, 'щомісяця', 0, 0, ''],
  ['Оренда квартири', 'Оренда житла', 15000, 1, 'щомісяця', 0, 0, 'велика сума: прогноз у «Днях» покаже, чи вистачить до зарплати'],
  ['Комуналка', 'Комуналка', 2500, 20, 'щомісяця', 0, 0, 'сума щоразу інша: ставте середню, факт ±25 % закриє план'],
  ['Бензин', 'Пальне', 3000, 1, 'щомісяця', 0, 0, 'кілька заправок на місяць: краще план «Пальне» в «Довіднику» (F); тут: якщо раз на місяць'],
  ['Спортзал', 'Спорт', 900, 5, 'щомісяця', 0, 0, ''],
  ['Садок або гурток', 'Діти', 2000, 10, 'щомісяця', 0, 0, 'у сім\'ї: колонка «Хто» (чий це платіж)'],
  ['Басейн на 3 місяці', 'Спорт', 2400, 1, 'щокварталу', 0, 0, 'щокварталу: рахується від місяця в колонці «З»'],
  ['Страховка авто', 'Транспорт', 2500, 15, 'щороку', 2, 0, 'щороку: місяць платежу береться з колонки «З»'],
  ['День народження мами', 'Подарунки', 2000, 10, 'щороку', 4, 0, 'великі річні витрати не застануть зненацька'],
  ['Зарплата', 'Зарплата', 40000, 25, 'щомісяця', 0, 0, 'надходження теж: план доходу з\'явиться в «Днях» і прогнозі'],
  ['Кредит у банку', 'Кредит / розстрочка', 2500, 15, 'щомісяця', 0, 0, 'борг із сумою й залишком краще на «' + DEBTS_SHEET + '» (модуль «🤝 Борги й розстрочки»)'],
  ['Розстрочка за телефон', 'Кредит / розстрочка', 1500, 3, 'щомісяця', 0, 10, 'з датою «До»: після неї планів не буде'],
  ['Повертаю борг другу', 'Повернення боргу', 1000, 10, 'щомісяця', 0, 0, 'позика з залишком: аркуш «' + DEBTS_SHEET + '»'],
  ['Мені повертають борг', 'Мені повернули борг', 1000, 20, 'щомісяця', 0, 0, 'вам винні: це надходження'],
  ['Відкладаю на заощадження', 'Заощадження', 3000, 26, 'щомісяця', 0, 0, 'відкладати після зарплати: переказ на банку чи депозит'],
  ['Заплати собі', CAT_SELF, 4000, 25, 'щомісяця', 0, 0, 'першим ділом у день зарплати: собі, решта на життя (шапка покаже % доходу)'],
  ['Інвестиції', CAT_INVEST, 2000, 26, 'щомісяця', 0, 0, 'ОВДП, фонд чи інше: окремо від заощаджень'],
  ['Щомісячний донат', 'Благодійність', 500, 1, 'щомісяця', 0, 0, ''],
];

function buildRegular_(ss) {
  const sh = ensureSize_(sheetOf_(ss, REG_SHEET), REG_LAST + 3, REG_COLS.length);
  sh.getRange(REG_HEAD, 1, 1, REG_COLS.length).setValues([REG_COLS]).setFontWeight('bold').setWrap(true)
    .setBackground(HDR_HEAD_BG).setFontColor(HDR_HEAD_FG).setHorizontalAlignment('center').setVerticalAlignment('middle');
  const rng = c => sh.getRange(c + REG_FIRST + ':' + c + REG_LAST), V = () => SpreadsheetApp.newDataValidation();
  const dict = ss.getSheetByName(DICT_SHEET), set = ss.getSheetByName(SETTINGS_SHEET);
  if (dict) rng('B').setDataValidation(V().requireValueInRange(dict.getRange('A' + D_FIRST + ':A' + D_LAST)).setAllowInvalid(true).build());
  rng('D').setDataValidation(V().requireNumberBetween(1, 31).setAllowInvalid(false).build());
  rng('E').setDataValidation(V().requireValueInList(REG_PER).setAllowInvalid(false).build());
  if (set) rng('H').setDataValidation(V().requireValueInRange(set.getRange('B' + SET_MEMBERS_FIRST + ':B' + SET_SHARED_ROW)).setAllowInvalid(true).build());
  rng('I').setDataValidation(V().requireValueInList(['так', 'ні']).setAllowInvalid(false).build());
  rng('C').setNumberFormat('#,##0.00');
  sh.getRange('F' + REG_FIRST + ':G' + REG_LAST).setNumberFormat('dd.mm.yyyy');
  rng('J').setFontColor('#5f6368');
  sh.setFrozenRows(REG_HEAD);
  sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 160); sh.setColumnWidth(10, 320);
  regExamplesAdd_(ss);
  regHeader_(ss);
  return sh;
}

/**
 * Приклади на «Регулярних» один раз (властивість reg_ex): у перші вільні рядки, «Увімкнено: ні», примітка «💡 Приклад…».
 * Категорії лише ті, що є в «Довіднику». Місця на всі немає: не додає. → скільки додано.
 */
function regExamplesAdd_(ss) {
  const sh = ss.getSheetByName(REG_SHEET), dp = PropertiesService.getDocumentProperties();
  if (!sh || dp.getProperty('reg_ex')) return 0;
  const dict = ss.getSheetByName(DICT_SHEET), cats = new Set(dict ? dict.getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 1).getValues().map(r => String(r[0]).trim()) : []);
  const ex = REG_EXAMPLES.filter(x => cats.has(x[1])), n = REG_LAST - REG_FIRST + 1;
  const v = sh.getRange(REG_FIRST, 1, n, 1).getValues(), free = [];
  v.forEach((r, i) => { if (String(r[0]).trim() === '') free.push(i); });
  if (!ex.length || free.length < ex.length) return 0;
  const tz = ss.getSpreadsheetTimeZone(), ym0 = todayKey_(ss).slice(0, 7);
  const ymAdd = k => { const t = Number(ym0.slice(0, 4)) * 12 + Number(ym0.slice(5, 7)) - 1 + k; return Math.floor(t / 12) + '-' + pad2_(t % 12 + 1); };
  const at = (k, day) => { const ym = ymAdd(k); return Utilities.parseDate(ym + '-' + pad2_(Math.min(day, Number(lastDayOf_(ym).slice(8)))), tz, 'yyyy-MM-dd'); };
  ex.forEach((x, j) => {
    const i = free[j], per = x[4], from = per === REG_PER[0] && !x[6] ? '' : at(x[5], x[3]), to = x[6] ? at(x[5] + x[6], x[3]) : '';
    sh.getRange(REG_FIRST + i, 1, 1, 9).setValues([[x[0], x[1], x[2], x[3], per, from, to, '', MOD_NO]]);
    sh.getRange(REG_FIRST + i, 1).setNote(REG_EX_MARK + ' шаблону' + (x[7] ? ': ' + x[7] : '') + '. Свій: поправте й поставте «так».');
  });
  dp.setProperty('reg_ex', '1');
  if (ss.getSheetByName(JOURNAL)) regPlansSync_(ss);                // «Стан»: підказки прикладів
  return ex.length;
}

/** Нові приклади для таблиць, де приклади вже додавали (версія 18): «Заплати собі» й «Інвестиції», якщо таких рядків ще немає. → додано */
function regExamplesMore_(ss) {
  const sh = ss.getSheetByName(REG_SHEET), dp = PropertiesService.getDocumentProperties();
  if (!sh || dp.getProperty('reg_ex2')) return 0;
  const dict = ss.getSheetByName(DICT_SHEET), cats = new Set(dict ? dict.getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 1).getValues().map(r => String(r[0]).trim()) : []);
  const v = sh.getRange(REG_FIRST, 1, REG_LAST - REG_FIRST + 1, 2).getValues(), names = v.map(r => String(r[0]).trim()), used = v.map(r => String(r[1]).trim());
  const add = REG_EXAMPLES.filter(x => (x[1] === CAT_SELF || x[1] === CAT_INVEST) && cats.has(x[1]) && names.indexOf(x[0]) < 0 && used.indexOf(x[1]) < 0);
  let n = 0;
  add.forEach(x => {
    const i = names.findIndex(s => !s);
    if (i < 0) return;
    sh.getRange(REG_FIRST + i, 1, 1, 9).setValues([[x[0], x[1], x[2], x[3], x[4], '', '', '', MOD_NO]]);
    sh.getRange(REG_FIRST + i, 1).setNote(REG_EX_MARK + ' шаблону: ' + x[7] + '. Свій: поправте й поставте «так».');
    names[i] = x[0]; n++;
  });
  dp.setProperty('reg_ex2', '1');
  if (n && ss.getSheetByName(JOURNAL)) regPlansSync_(ss);
  return n;
}

/* ---------- шапка: плитки в стилі Grafana (рядки 1–3) ---------- */

const RH = { per: 'N', type: 'O', days: 'Q', lab: 'R', val: 'S', opt: 'T', last: 'U' };   // службові колонки (сховані)
const RH_DAYS = 30, RH_TILES = 6;
const REG_NOTE_A = 'Кожен рядок стає планом «🔁 план регулярний» у журналі до кінця наступного місяця. Прийшов факт тієї ж категорії ' +
  '(±5 днів, ±25 % суми), наприклад з банку: план місяця закривається сам. Щокварталу й щороку: від місяця з колонки «З». ' + REG_HINT;
/** Підсумки шапки (S1…S27): [рядок, підпис, формула]. Колонки: N на місяць, O тип, J стан. */
function regHeaderVals_(debtsOn) {
  const R = c => '$' + c + '$' + REG_FIRST + ':$' + c + '$' + REG_LAST, N = R(RH.per), O = R(RH.type), JR = J_RANGE;
  const m0 = '(EOMONTH(TODAY(),-1)+1)', m1 = 'EOMONTH(TODAY(),0)', tag = '"' + REG_TAG + '*"';
  const B = c => "'" + DEBTS_SHEET + "'!$" + c + '$' + DEBT_FIRST + ':$' + c + '$' + DEBT_LAST, S = k => '$' + RH.val + '$' + k;
  return [
    [1, 'доходи на місяць', '=SUMIFS(' + N + ',' + O + ',"' + T_IN + '")'],
    [2, 'собі й заощадження', '=SUMIFS(' + N + ',' + R('B') + ',"' + CAT_SELF + '")+SUMIFS(' + N + ',' + R('B') + ',"Заощадження")'],
    [3, 'інвестиції', '=SUMIFS(' + N + ',' + R('B') + ',"' + CAT_INVEST + '")'],
    [4, 'собі й інвестиції', '=' + S(2) + '+' + S(3)],
    [5, 'регулярні витрати', '=SUM(' + N + ')-' + S(1) + '-' + S(4)],
    [6, 'платежі боргів на місяць', debtsOn ? '=SUMIFS(' + B('D') + ',' + B('B') + ',"' + DEBT_OWE + '",' + B('J') + ',">0")' : 0],
    [7, 'обов\'язкові (витрати + борги)', '=' + S(5) + '+' + S(6)],
    [8, 'лишається', '=' + S(1) + '-' + S(7) + '-' + S(4)],
    [9, 'частка обов\'язкових', '=IF(' + S(1) + '>0,' + S(7) + '/' + S(1) + ',0)'],
    [10, 'увімкнено', '=COUNTIFS(' + R('A') + ',"<>",' + R('I') + ',"так")'],
    [11, 'з них витрат', '=COUNTIFS(' + R('A') + ',"<>",' + R('I') + ',"так",' + O + ',"<>' + T_IN + '")'],
    [12, 'з них доходів', '=' + S(10) + '-' + S(11)],
    [13, 'прикладів', '=COUNTIF(' + R('J') + ',"💡*")'],
    [14, 'сплачено цього місяця', '=COUNTIF(' + R('J') + ',"✓*")'],
    [15, 'ще планів цього місяця', '=COUNTIFS(' + JR(JC.note) + ',' + tag + ',' + JR(JC.state) + ',1,' + JR(JC.date) + ',">="&' + m0 + ',' + JR(JC.date) + ',"<="&' + m1 + ')'],
    [16, 'ще сума цього місяця', '=SUMIFS(' + JR(JC.plan) + ',' + JR(JC.note) + ',' + tag + ',' + JR(JC.date) + ',">="&' + m0 + ',' + JR(JC.date) + ',"<="&' + m1 + ')'],
    [17, 'прострочено', '=COUNTIF(' + R('J') + ',"*прострочено*")'],
    [18, 'найближча дата', '=MINIFS(' + JR(JC.date) + ',' + JR(JC.note) + ',' + tag + ',' + JR(JC.state) + ',1,' + JR(JC.date) + ',">="&TODAY())'],
    [19, 'найближчий: коментар', '=IF(' + S(18) + '=0,"",IFERROR(INDEX(FILTER(' + JR(JC.note) + ',' + JR(JC.date) + '=' + S(18) + ',' + JR(JC.state) + '=1,LEFT(' + JR(JC.note) + ',' + REG_TAG.length + ')="' + REG_TAG + '"),1),""))'],
    [20, 'найближчий: назва', '=IFERROR(MID(' + S(19) + ',' + (REG_TAG.length + 1) + ',FIND(":",' + S(19) + ')-' + (REG_TAG.length + 1) + '),"")'],
    [21, 'найближчий: сума', '=IF(' + S(18) + '=0,0,SUMIFS(' + JR(JC.plan) + ',' + JR(JC.note) + ',' + tag + ',' + JR(JC.date) + ',' + S(18) + '))'],
    [22, '30 днів: сума', '=SUM($' + RH.days + '$' + REG_FIRST + ':$' + RH.days + '$' + (REG_FIRST + RH_DAYS - 1) + ')'],
    [23, 'залишок боргів', debtsOn ? '=SUMIFS(' + B('J') + ',' + B('B') + ',"' + DEBT_OWE + '")' : 0],
    [24, 'ліміт частинами: вільно', debtsOn ? "=IF('" + DEBTS_SHEET + "'!$B$" + DEBT_LIMITS[0].row + '="","",\'' + DEBTS_SHEET + "'!$E$" + DEBT_LIMITS[0].row + ')' : ''],
    [26, 'смужка: обов\'язкові', '=' + S(7)],
    [27, 'смужка: решта доходу', '=MAX(0,' + S(1) + '-' + S(7) + ')'],
  ];
}

/** Шапка «Регулярних»: рядки 1–3 (A…J, з модулем боргів ще K:L), службові N…U. Ваші рядки з 5-го не чіпає. */
function regHeader_(ss) {
  const sh = ss.getSheetByName(REG_SHEET);
  if (!sh) return false;
  const last = colNum_(RH.last), debtsOn = modOn_('debts') && !!ss.getSheetByName(DEBTS_SHEET);
  ensureSize_(sh, REG_LAST, last);
  const top = sh.getRange(1, 1, 3, colNum_('L'));
  top.breakApart();
  top.setValues([0, 1, 2].map(() => Array(colNum_('L')).fill('')));
  [1, 2, 3].forEach(r => sh.getRange(r, 1).setNote(''));
  top.setBackground(DK.bg).setFontColor(DK.text).setFontWeight('normal').setFontSize(10).setWrap(false).setVerticalAlignment('middle').setHorizontalAlignment('left');
  // службові: на місяць (N), тип (O), 30 днів (Q), підсумки (R:S), параметри графіків (T:U)
  const per = [], typ = [], days = [];
  for (let r = REG_FIRST; r <= REG_LAST; r++) {
    per.push(['=IF(OR($A' + r + '="",$I' + r + '<>"так",N($C' + r + ')<=0,AND($G' + r + '<>"",N($G' + r + ')<TODAY())),0,N($C' + r + ')*IF($E' + r + '="' + REG_PER[1] + '",1/3,IF($E' + r + '="' + REG_PER[2] + '",1/12,1)))']);
    typ.push(['=IF($B' + r + '="","",IFERROR(VLOOKUP($B' + r + ',\'' + DICT_SHEET + '\'!$A$' + D_FIRST + ':$B$' + D_LAST + ',2,FALSE),""))']);
  }
  for (let k = 0; k < RH_DAYS; k++) {
    const day = 'TODAY()+' + k;
    days.push(['=SUMIFS(' + J_RANGE(JC.plan) + ',' + J_RANGE(JC.note) + ',"' + REG_TAG + '*",' + J_RANGE(JC.date) + ',' + day + ')+SUMIFS(' + J_RANGE(JC.plan) + ',' +
      J_RANGE(JC.note) + ',"' + DEBT_TAG + '*",' + J_RANGE(JC.date) + ',' + day + ')']);
  }
  setF_(sh.getRange(REG_FIRST, colNum_(RH.per), per.length, 1), per);
  setF_(sh.getRange(REG_FIRST, colNum_(RH.type), typ.length, 1), typ);
  setF_(sh.getRange(REG_FIRST, colNum_(RH.days), days.length, 1), days);
  sh.getRange(REG_HEAD, colNum_(RH.per), 1, 4).setValues([['на місяць', 'тип', '', 'платежі по днях (30)']]);
  sh.getRange(1, colNum_(RH.lab), 27, 2).clearContent();
  regHeaderVals_(debtsOn).forEach(v => {
    sh.getRange(v[0], colNum_(RH.lab)).setValue(v[1]);
    dashPut_(sh.getRange(v[0], colNum_(RH.val)), v[2]);
  });
  const S = k => '$' + RH.val + '$' + k, O = (r, rows) => sh.getRange(r, colNum_(RH.opt), rows.length, 2).setValues(rows);
  const T = (r, n) => '$' + RH.opt + '$' + r + ':$' + RH.last + '$' + (r + n - 1);
  O(1, [['charttype', 'bar'], ['max', ''], ['color1', DKC.violet], ['color2', DKC.blue]]);
  dashPut_(sh.getRange(2, colNum_(RH.last)), '=MAX(' + S(1) + ',' + S(4) + ',1)');
  O(6, [['charttype', 'bar'], ['max', ''], ['color1', DKC.red], ['color2', DKC.green]]);
  dashPut_(sh.getRange(7, colNum_(RH.last)), '=MAX(' + S(1) + ',' + S(7) + ',1)');
  O(11, [['charttype', 'bar'], ['max', ''], ['color1', DKC.green], ['color2', DK.lineHead]]);
  dashPut_(sh.getRange(12, colNum_(RH.last)), '=MAX(' + S(14) + '+' + S(15) + ',1)');
  O(16, [['charttype', 'column'], ['color', DKC.orange], ['ymin', 0]]);
  sh.hideColumns(colNum_(RH.per), last - colNum_(RH.per) + 1);
  // плитки
  const money = x => 'TEXT(' + x + ',"#,##0")&" ₴"', tiles = [
    { c: 1, color: DKC.violet,
      t: '="💰 СОБІ Й ІНВЕСТИЦІЇ / МІС · СОБІ "&TEXT(' + S(2) + ',"#,##0")&" · ІНВЕСТ. "&TEXT(' + S(3) + ',"#,##0")',
      big: '=' + money(S(4)) + '&IF(' + S(1) + '>0,"  ·  "&TEXT(' + S(4) + '/' + S(1) + ',"0%")&" доходу","")',
      sub: '=IF(' + S(4) + '=0,"Додайте «' + CAT_SELF + '» чи «' + CAT_INVEST + '» нижче (є приклади з «ні»)",SPARKLINE($' + RH.val + '$2:$' + RH.val + '$3,' + T(1, 4) + '))' },
    { c: 3, color: DKC.blue, t: '🔁 ПІДКЛЮЧЕНО РЕГУЛЯРНИХ', big: '=' + S(10),
      sub: '="💸 "&' + S(11) + '&" витрат · 💵 "&' + S(12) + '&" доходів"&IF(' + S(13) + '>0," · 💡 "&' + S(13) + '&" прикладів","")' },
    { c: 5, color: DKC.green, t: '="💵 ЛИШАЄТЬСЯ / МІС"&IF(' + S(1) + '>0," · ОБОВ\'ЯЗКОВІ "&TEXT(' + S(9) + ',"0%"),"")', big: '=' + S(8), fmt: F_MONEY,
      sub: '=IF(' + S(1) + '=0,"Додайте дохід («Зарплата») нижче",SPARKLINE($' + RH.val + '$26:$' + RH.val + '$27,' + T(6, 4) + '))' },
    { c: 7, color: DKC.amber, t: '="📅 ЦЕЙ МІСЯЦЬ · ЩЕ "&' + money(S(16)) + '&IF(' + S(17) + '>0," · ⚠️ "&' + S(17) + '&" ПРОСТРОЧЕНО","")',
      big: '=' + S(14) + '&" / "&(' + S(14) + '+' + S(15) + ')&" ✓"',
      sub: '=IF(' + S(14) + '+' + S(15) + '=0,"Цього місяця регулярних немає",SPARKLINE($' + RH.val + '$14:$' + RH.val + '$15,' + T(11, 4) + '))' },
    { c: 9, color: DKC.orange, t: '="⏭ 30 ДНІВ · НАЙБЛИЖЧИЙ: "&IF(' + S(18) + '=0,"НЕМАЄ",' + S(20) + '&" "&TEXT(' + S(18) + ',"dd.mm")&" · "&' + money(S(21)) + ')',
      big: '=' + S(22), fmt: F_MONEY,
      sub: '=SPARKLINE($' + RH.days + '$' + REG_FIRST + ':$' + RH.days + '$' + (REG_FIRST + RH_DAYS - 1) + ',' + T(16, 3) + ')' },
  ];
  if (debtsOn) tiles.push({ c: 11, color: DKC.pink, t: '💳 БОРГИ Й ЧАСТИНАМИ / МІС', big: '=' + S(6), fmt: F_MONEY,
    sub: '="залишок "&' + money(S(23)) + '&IF(' + S(24) + '="",""," · ліміт вільно "&' + money(S(24)) + ')' });
  tiles.forEach(x => {
    const box = sh.getRange(1, x.c, 3, 2);
    box.setBackground(DK.tile).setBorder(true, true, true, true, null, null, DK.bg, dashThick_());
    [1, 2, 3].forEach(r => sh.getRange(r, x.c, 1, 2).merge());
    dashPut_(sh.getRange(1, x.c), x.t);
    dashPut_(sh.getRange(2, x.c), x.big);
    dashPut_(sh.getRange(3, x.c), x.sub);
    sh.getRange(1, x.c).setFontSize(8).setFontWeight('bold').setFontColor(DK.muted);
    sh.getRange(2, x.c).setFontSize(18).setFontWeight('bold').setFontColor(x.color).setNumberFormat(x.fmt || F_NUM);
    sh.getRange(3, x.c).setFontSize(9).setFontColor(DK.muted);
  });
  sh.setRowHeight(1, 22); sh.setRowHeight(2, 40); sh.setRowHeight(3, 26);
  // колір «лишається» нижче нуля: червоний (наші правила лише в рядках 1–3)
  const ours = r => { try { return r.getRanges().every(g => g.getLastRow() <= 3); } catch (e) { return false; } };
  const keep = sh.getConditionalFormatRules().filter(r => !ours(r));
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(DKC.red).setRanges([sh.getRange('E2:F2')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('⚠️').setFontColor(DKC.red).setRanges([sh.getRange('G1:H1')]).build()].concat(keep));
  // пояснення: примітки на заголовках
  sh.getRange(REG_HEAD, 1).setNote(REG_NOTE_A);
  sh.getRange(REG_HEAD, 9).setNote('«так»: правило ставить плани. «ні»: нічого не планує (так стоять приклади шаблону).');
  sh.getRange(REG_HEAD, 10).setNote('Пише скрипт: наступний платіж, сплачено цього місяця, прострочено. Шапка вгорі рахує з цієї колонки.');
  return true;
}

/** Шапка «Регулярних» нового вигляду? (крок 18) */
function regHeaderNew_(ss) {
  const sh = ss.getSheetByName(REG_SHEET);
  return !!sh && sh.getMaxColumns() >= colNum_(RH.last) && String(sh.getRange(10, colNum_(RH.lab)).getValue()) === 'увімкнено' &&
    String(sh.getRange(REG_FIRST, colNum_(RH.per)).getFormula()).indexOf('=IF(OR($A' + REG_FIRST) === 0;
}

/** Ваші правила: без вимкнених прикладів шаблону. */
function regOwn_(ss) { return regList_(ss).filter(r => r.on || !r.ex); }

/** Правила з аркуша: [{ i, name, cat, amt, day, per, from, to, who, on }]. */
function regList_(ss) {
  const sh = ss.getSheetByName(REG_SHEET), tz = ss.getSpreadsheetTimeZone();
  if (!sh) return [];
  const notes = sh.getRange(REG_FIRST, 1, REG_LAST - REG_FIRST + 1, 1).getNotes();
  return sh.getRange(REG_FIRST, 1, REG_LAST - REG_FIRST + 1, 9).getValues().map((r, i) => ({
    i: i, name: String(r[0]).trim(), cat: String(r[1]).trim(), amt: Number(r[2]) || 0, day: Math.max(1, Math.min(31, Math.round(Number(r[3]) || 1))),
    per: REG_PER.indexOf(String(r[4]).trim()) >= 0 ? String(r[4]).trim() : REG_PER[0], from: dateKey_(r[5], tz) || '', to: dateKey_(r[6], tz) || '',
    who: String(r[7] || '').trim(), on: String(r[8]).trim() !== 'ні',
    ex: String(notes[i][0] || '').indexOf(REG_EX_MARK) === 0 && REG_EXAMPLES.some(x => x[0] === String(r[0]).trim()),   // перейменували: уже ваш
    hint: String(notes[i][0] || '').replace(/^💡 Приклад шаблону:?\s*/, '').replace(/\s*Свій: поправте й поставте «так»\.$/, '').replace(/\.$/, ''),
  })).filter(x => x.name && x.cat && x.amt > 0);
}

/** Чиста функція: дата платежу правила в місяці ym або '' (не той місяць за періодичністю, поза «з» / «до»). */
function regDate_(rule, ym, today) {
  const anchor = Number((rule.from || today).slice(5, 7)), mon = Number(ym.slice(5, 7));
  if (rule.per === 'щокварталу' && (((mon - anchor) % 3) + 3) % 3 !== 0) return '';
  if (rule.per === 'щороку' && mon !== anchor) return '';
  const d = ym + '-' + pad2_(Math.min(rule.day, Number(lastDayOf_(ym).slice(8))));
  if (rule.from && d < rule.from) return '';
  if (rule.to && d > rule.to) return '';
  return d;
}

/** Факт або чернетка (не план, не порожній рядок, не майбутній рядок без статусу). */
const regFact_ = (o, today) => !o.blank && !!o.key && Number(o.b) > 0 && (o.e === STATUS_FACT || o.e === STATUS_DRAFT || (o.e === '' && o.key <= today));

/**
 * Чиста функція: які плани мають стояти в журналі (want) і які прибрати навіть з минулою датою (paid: місяць уже
 * сплачено, правило видалено чи вимкнено). Від місяця з датою «сьогодні − 5 днів» до horizon.
 * → { want: [{ key, date, amount, cat, who, rule }], paid: [key], state: { rule.i: { next, paid, over } } }
 */
function regWanted_(rules, rows, today, horizon) {
  const want = [], paid = [], state = {}, claimed = {};
  const facts = rows.filter(o => regFact_(o, today));
  rules.forEach(rule => {
    const st = state[rule.i] = { next: '', paidYm: '', over: '' };
    if (!rule.on) return;
    let ym = addDays_(today, -REG_CLOSE_DAYS).slice(0, 7);
    rows.forEach(o => {                                                         // старі невиконані плани цього правила теж перевіряємо (місяць з коментаря)
      const m = isPlan_(o.e) && o.d.indexOf(REG_TAG + rule.name + ':') === 0 ? /: (\d\d)\.(\d{4})$/.exec(o.d) : null;
      const pm = m ? m[2] + '-' + m[1] : '';
      if (pm && pm < ym && pm >= addDays_(today, -366).slice(0, 7)) ym = pm;
    });
    if (rule.from && rule.from.slice(0, 7) > ym) ym = rule.from.slice(0, 7);
    for (let k = 0; k < 15 && ym <= horizon.slice(0, 7); k++, ym = addDays_(lastDayOf_(ym), 1).slice(0, 7)) {
      const date = regDate_(rule, ym, today);
      if (!date || date > horizon) continue;
      const key = REG_TAG + rule.name + ': ' + ym.slice(5) + '.' + ym.slice(0, 4);
      const plan = rows.find(o => isPlan_(o.e) && o.d === key && o.key);              // план уже стоїть (дату могли перенести)
      const near = x => Math.abs(dayDiff_(x, date)) <= REG_CLOSE_DAYS || (!!plan && Math.abs(dayDiff_(x, plan.key)) <= REG_CLOSE_DAYS);
      let hit = facts.find(o => !claimed[o.row] && (o.d === key || o.d.indexOf(key + ' · ') === 0));   // « · …»: план закрила операція з банку
      if (!hit) {
        hit = facts.find(o => !claimed[o.row] && o.c === rule.cat && near(o.key) &&
          (Math.abs(Number(o.b) - rule.amt) <= rule.amt * REG_TOL + 0.005 || (o.d && o.d.indexOf(rule.name) >= 0)));
      }
      if (hit) { claimed[hit.row] = true; paid.push(key); st.paidYm = ym; continue; }
      if (date < today) {                                                        // минуле без факту: план (якщо був) лишається
        if (plan && plan.key >= today) {                                         // ви перенесли прострочений платіж на майбутнє: так і лишаємо
          want.push({ key: key, date: plan.key, amount: round2_(Number(plan.b) || rule.amt), cat: rule.cat, who: rule.who, rule: rule.i });
          if (!st.next || plan.key < st.next) st.next = plan.key;
        }
        continue;
      }
      want.push({ key: key, date: date, amount: round2_(rule.amt), cat: rule.cat, who: rule.who, rule: rule.i });
      if (!st.next || date < st.next) st.next = date;
    }
  });
  rows.filter(o => isPlan_(o.e) && o.d.indexOf(REG_TAG) === 0 && o.key < today && paid.indexOf(o.d) < 0).forEach(o => {
    const r = rules.find(x => o.d.indexOf(REG_TAG + x.name + ':') === 0);
    if (!r || !r.on) { paid.push(o.d); return; }                               // правило видалено чи вимкнено: старий план теж геть
    if (!state[r.i].over || o.key < state[r.i].over) state[r.i].over = o.key;
  });
  return { want: want, paid: paid, state: state };
}

/** Плани з «Регулярних» у журналі й колонка «Стан». → { added, updated, removed } */
function regPlansSync_(ss) {
  const st = { added: 0, updated: 0, removed: 0 };
  const sh = ss.getSheetByName(REG_SHEET);
  if (!ss.getSheetByName(JOURNAL) || !sh) return st;
  const today = todayKey_(ss), rows = readJournal_(ss), rules = regList_(ss);
  const horizon = lastDayOf_(addDays_(lastDayOf_(today.slice(0, 7)), 1).slice(0, 7));
  const w = regWanted_(rules, rows, today, horizon);
  const r = tagPlansSync_(ss, REG_TAG, w.want, rows, w.paid);
  const dm = k => k.slice(8) + '.' + k.slice(5, 7);
  const col = [];
  for (let i = 0; i <= REG_LAST - REG_FIRST; i++) {
    const rule = rules.find(x => x.i === i), s = rule && w.state[i];
    col.push([!rule ? '' : !rule.on ? (rule.ex ? '💡 приклад, «ні»: ' + (rule.hint || 'поставте «так», і він стане планом') : 'вимкнено') : (s.over ? '⚠️ прострочено ' + dm(s.over) + '; ' : '') +
      (s.paidYm === today.slice(0, 7) ? '✓ цього місяця сплачено' + (s.next ? '; ' : '') : '') + (s.next ? 'наступний ' + dm(s.next) : s.paidYm === today.slice(0, 7) ? '' : 'планів до ' + dm(horizon) + ' немає')]);
  }
  sh.getRange(REG_FIRST, 10, col.length, 1).setValues(col);
  PropertiesService.getDocumentProperties().setProperties({ reg_day: today, reg_t: String(Date.now()) });
  return r;
}

/** З mbsTick: раз на день або після правок у журналі чи «Регулярних» (reg_edit пише onEdit). */
function regIfDue_(ss) {
  if (!ss.getSheetByName(REG_SHEET)) return null;
  const dp = PropertiesService.getDocumentProperties();
  const edit = Math.max(Number(dp.getProperty('lastEdit') || 0), Number(dp.getProperty('reg_edit') || 0));
  if (dp.getProperty('reg_day') === todayKey_(ss) && edit < Number(dp.getProperty('reg_t') || 0)) return null;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return null;
  try { return regPlansSync_(ss); } finally { lock.releaseLock(); }
}

/** Меню: оновити плани зараз. */
function regularNow() {
  if (!menuOk_('🔁 Регулярні')) return;
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(REG_SHEET)) { try { schemaRun_(ss); } catch (e) { show_('🔁 Регулярні', ['⚠️ ' + esc_(e.message)]); return; } }
  const r = regPlansSync_(ss);
  show_('🔁 Регулярні', ['Плани в журналі: +' + r.added + ', оновлено ' + r.updated + ', прибрано ' + r.removed + '.', 'Стан кожного правила: колонка J на «' + REG_SHEET + '».']);
}
