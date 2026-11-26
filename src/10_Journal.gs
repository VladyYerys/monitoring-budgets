/** 10_Journal.gs: Аркуш «Журнал»: заголовки, списки, формули службових колонок, кольори, читання рядків. */

/*
 * Людина вводить A…H: дата, сума (завжди додатна), категорія, коментар, статус, непомітна (вручну), хто, причина з ліміту.
 * Формули: I «Тип» (з довідника), J «Мої після операції»; службові L…X приховані.
 * Логіка «моїх грошей» та сама, що в особистій таблиці: гроші місяця = мої надходження бюджетного місяця
 * (дохід з дня B5 іде в наступний місяць) мінус відкладене на погашення ліміту; списання спершу з «моїх»,
 * що не вмістилось: «з ліміту» (з причиною).
 * Y…AE (приховані) пише скрипт: що прийшло з банку і черговість запису.
 * Порядок рядків (13_Order.gs): майбутнє зверху, жовтий рядок-лінія, порожні рядки для вводу, минуле від нового до старого.
 */

const J_HEADERS = {
  A: 'Дата', B: 'Сума, ₴', C: 'Категорія', D: 'Коментар', E: 'Статус', F: 'Непомітна (вручну)', G: 'Хто', H: 'Причина (якщо з ліміту)',
  I: 'Тип', J: 'Мої після операції',
  L: 'Сума зі знаком (службовий)', M: 'Стан: 0 факт, 1 план, 2 чернетка чи переказ (службовий)', N: 'Сума план (службовий)',
  O: 'Сума факт (службовий)', P: 'Непомітна (службовий)', Q: 'Бюджетний місяць (службовий)', R: 'Мої гроші: вид (службовий)',
  S: 'Мої: внесок (службовий)', T: 'Списання (службовий)', U: 'Мої після операції (службовий)', V: 'З моїх (службовий)',
  W: 'З ліміту (службовий)', X: 'Причина ліміту (службовий)',
  Y: 'Банк: номер операції', Z: 'Банк: опис', AA: 'Банк: сума зі знаком', AB: 'Банк: MCC', AC: 'Банк: час', AD: 'Банк: залишок після операції',
  AE: 'Черговість запису (службовий)', AF: 'Рахунок (порожньо: основна картка)', AG: 'Вид рахунку (службовий)', AH: 'Сума зі знаком, будь-який рахунок (службовий)',
};

/** Формули рядка {r} (порядок не важливий: Таблиці рахують залежності самі). */
function journalFormulas_() {
  const dict = (c) => D_RANGE(c);
  const start = 'DATE(YEAR(' + SET_REF('start') + '),MONTH(' + SET_REF('start') + '),1)';
  const inList = (cell, list) => 'OR(' + list.map(t => cell + '="' + t + '"').join(',') + ')';
  return {
    I: '=IF(C{r}="","",IF(C{r}="' + CAT_TRANSFER + '","↔ уточніть",IFERROR(VLOOKUP(C{r},\'' + DICT_SHEET + '\'!$A$' + D_FIRST + ':$B$' + D_LAST + ',2,FALSE),"⚠ невідома категорія")))',
    J: '=IF(M{r}>=1,"",IF(U{r}="","",U{r}))',
    AH: '=IF(OR(B{r}="",I{r}=""),"",IF(I{r}="' + T_RESERVE + '",0,IF(' + inList('I{r}', TYPES_IN) + ',O{r},-O{r})))',
    AG: '=IF(AF{r}="","' + K_MAIN + '",IFERROR(VLOOKUP(AF{r},\'' + ACC_SHEET + '\'!$A$' + ACC_FIRST + ':$D$' + ACC_LAST + ',4,FALSE)&"","' + K_MAIN + '"))',
    L: '=IF(AH{r}="","",IF(OR(AG{r}="' + K_MAIN + '",AG{r}=""),AH{r},0))',
    M: '=IF(OR(A{r}="",B{r}=""),0,IF(OR(E{r}="' + STATUS_DRAFT + '",C{r}="' + CAT_TRANSFER + '",C{r}=""),2,IF(OR(E{r}="' + STATUS_PLAN + '",E{r}="' + STATUS_REG +
      '"),1,IF(AND(E{r}="",N(A{r})>TODAY()),1,0))))',
    N: '=IF(M{r}=1,N(B{r}),0)',
    O: '=IF(B{r}="","",IF(M{r}>=1,0,B{r}))',
    P: '=IF(OR(A{r}="",B{r}=""),"",IF(F{r}<>"",F{r},IFERROR(INDEX(' + dict('D') + ',MATCH(C{r},' + dict('A') + ',0))&"","")))',
    R: '=IF(C{r}="","",IFERROR(INDEX(' + dict('E') + ',MATCH(C{r},' + dict('A') + ',0))&"",""))',
    Q: '=IF(OR(A{r}="",B{r}=""),"",IF(AND(OR(R{r}="' + MINE_IN + '",R{r}="' + MINE_SET_ASIDE + '"),DAY(A{r})>=' + SET_REF('bday') +
      '),EDATE(DATE(YEAR(A{r}),MONTH(A{r}),1),1),DATE(YEAR(A{r}),MONTH(A{r}),1)))',
    S: '=IF(OR(Q{r}="",N(Q{r})<' + start + ',AG{r}="' + K_FOP + '"),0,IF(R{r}="' + MINE_IN + '",O{r},IF(R{r}="' + MINE_SET_ASIDE + '",-O{r},0)))',
    T: '=IF(OR(Q{r}="",N(Q{r})<' + start + ',AG{r}="' + K_FOP + '"),0,IF(N(AH{r})<0,O{r},0))',
    U: '=IF(OR(Q{r}="",N(Q{r})<' + start + '),"",SUMIFS($S$' + J_FIRST + ':$S$' + J_LAST + ',$Q$' + J_FIRST + ':$Q$' + J_LAST + ',Q{r},$A$' + J_FIRST +
      ':$A$' + J_LAST + ',"<"&A{r})-SUMIFS($T$' + J_FIRST + ':$T$' + J_LAST + ',$Q$' + J_FIRST + ':$Q$' + J_LAST + ',Q{r},$A$' + J_FIRST + ':$A$' + J_LAST +
      ',"<"&A{r})+SUMIFS($S{r}:$S$' + J_LAST + ',$Q{r}:$Q$' + J_LAST + ',Q{r},$A{r}:$A$' + J_LAST + ',A{r})-SUMIFS($T{r}:$T$' + J_LAST + ',$Q{r}:$Q$' + J_LAST +
      ',Q{r},$A{r}:$A$' + J_LAST + ',A{r}))',
    V: '=IF(N(T{r})=0,0,MIN(T{r},MAX(0,N(U{r})+T{r})))',
    W: '=N(T{r})-N(V{r})',
    X: '=IF(N(W{r})>0,IF(H{r}="","без причини",H{r}),"")',
  };
}

function buildJournal_(ss) {
  const sh = ensureSize_(sheetOf_(ss, JOURNAL), J_LAST, colNum_(J_SERVICE_LAST));
  if (!headerIsNew_(ss)) {                                         // нова шапка (25_Header.gs) замінює заголовок і підказку
    sh.getRange(1, 1).setValue('📒 Журнал').setFontWeight('bold').setFontSize(14);
    sh.getRange(2, 1).setValue('Пишіть у рядки нижче: дата, сума (завжди додатна), категорія зі списку, статус. Решту рахують формули. ' +
      'Сірі колонки праворуч службові: не правте.').setFontColor('#5f6368');
  }
  const last = colNum_(J_SERVICE_LAST);
  const heads = [];
  for (let c = 1; c <= last; c++) heads.push(J_HEADERS[colStr_(c)] || '');
  sh.getRange(J_HEAD, 1, 1, last).setValues([heads]).setFontWeight('bold').setWrap(true);
  const fs = journalFormulas_(), n = J_LAST - J_FIRST + 1;
  Object.keys(fs).forEach(col => {
    const rows = [];
    for (let r = J_FIRST; r <= J_LAST; r++) rows.push([rowF_(fs[col], r)]);
    setF_(sh.getRange(J_FIRST, colNum_(col), n, 1), rows);
  });
  const dv = (rng, rule) => sh.getRange(rng).setDataValidation(rule.setAllowInvalid(true).build());
  dv('E' + J_FIRST + ':E' + J_LAST, SpreadsheetApp.newDataValidation().requireValueInList(STATUSES));
  dv('C' + J_FIRST + ':C' + J_LAST, SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(DICT_SHEET).getRange('A' + D_FIRST + ':A' + D_LAST)));
  dv('F' + J_FIRST + ':F' + J_LAST, SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(DICT_SHEET).getRange('I' + D_HOLES_FIRST + ':I' + D_HOLES_LAST)));
  dv('H' + J_FIRST + ':H' + J_LAST, SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(DICT_SHEET).getRange('K' + D_REASONS_FIRST + ':K' + D_REASONS_LAST)));
  const set = ss.getSheetByName(SETTINGS_SHEET);
  if (set) dv('G' + J_FIRST + ':G' + J_LAST, SpreadsheetApp.newDataValidation().requireValueInRange(set.getRange('B' + SET_MEMBERS_FIRST + ':B' + SET_SHARED_ROW)));
  sh.getRange('A' + J_FIRST + ':A' + J_LAST).setNumberFormat('dd.mm.yyyy');
  sh.getRange('B' + J_FIRST + ':B' + J_LAST).setNumberFormat('#,##0.00');
  sh.getRange('J' + J_FIRST + ':J' + J_LAST).setNumberFormat('#,##0');
  sh.getRange(JC.bankId + J_FIRST + ':' + JC.order + J_LAST).setNumberFormat('@');
  sh.getRange(JC.bankAmt + J_FIRST + ':' + JC.bankAmt + J_LAST).setNumberFormat('#,##0.00');
  sh.getRange(JC.bankBal + J_FIRST + ':' + JC.bankBal + J_LAST).setNumberFormat('#,##0.00');
  sh.getRange(JC.order + J_FIRST + ':' + JC.order + J_LAST).setNumberFormat('0');
  journalColors_(sh);
  sh.setFrozenRows(J_HEAD);
  sh.hideColumns(colNum_(J_SERVICE_FIRST), last - colNum_(J_SERVICE_FIRST) + 1);
  return sh;
}

/*
 * Кольори рядків. Формули умовного форматування без ком (лише * і +): так вони працюють у таблиці з будь-яким
 * регіоном (в українському між аргументами «;»).
 */
const J_CF = [
  { f: '=($B5<>"")*(($C5="' + CAT_TRANSFER + '")+($C5=""))*($E5<>"' + STATUS_DRAFT + '")', bg: '#f8cfcb', fc: '#a50e0e', cols: 'A:H' },
  { f: '=($A5="")*(N($A4)>TODAY())', bg: '#fbbc04', cols: 'A:H' },
  { f: '=$E5="' + STATUS_DRAFT + '"', bg: '#fff4c2', cols: 'A:H' },
  { f: '=$M5=1', bg: '#e8f0fe', fc: '#1a56b0', cols: 'A:H' },
  { f: '=($W5>0)*($H5="")', bg: '#fde3cf', cols: 'H:H' },
];

function journalColors_(sh) {
  const mine = J_CF.map(x => x.f);
  const keep = sh.getConditionalFormatRules().filter(r => {
    const b = r.getBooleanCondition && r.getBooleanCondition();
    return !(b && mine.indexOf(String(b.getCriteriaValues()[0])) >= 0);
  });
  const rules = J_CF.map(x => {
    const c = x.cols.split(':');
    let b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(x.f).setBackground(x.bg);
    if (x.fc) b = b.setFontColor(x.fc);
    return b.setRanges([sh.getRange(c[0] + J_FIRST + ':' + c[1] + J_LAST)]).build();
  });
  sh.setConditionalFormatRules(rules.concat(keep));
}

/**
 * Журнал рядками (значення A…AE) одним читанням. key: дата 'yyyy-MM-dd'; blank: A…H порожні;
 * bank: [номер, опис, сума зі знаком, MCC, час, залишок після]; ord: черговість запису.
 */
function readJournal_(ss) {
  const sh = ss.getSheetByName(JOURNAL), tz = ss.getSpreadsheetTimeZone(), kinds = accKinds_(ss);
  const v = sh.getRange(J_FIRST, 1, J_LAST - J_FIRST + 1, J_READ_N).getValues();
  const empty = x => x === '' || x === null || x === undefined;
  return v.map((r, i) => ({
    row: J_FIRST + i, key: dateKey_(r[0], tz), a: r[0], b: r[1], c: String(r[2] === null ? '' : r[2]), d: String(r[3] || ''), e: String(r[4] || ''),
    f: String(r[5] || ''), g: String(r[6] || ''), h: String(r[7] || ''), i: String(r[8] || ''),
    l: r[11], m: r[12], o: r[14], p: String(r[15] || ''), q: r[16], r: String(r[17] || ''), t: r[19], v: r[21], w: r[22], x: String(r[23] || ''),
    bank: r.slice(24, 24 + BANK_N), ord: r[30], acct: String(r[31] || ''), kind: r[31] ? kinds[String(r[31])] || K_MAIN : K_MAIN, all: r[33],
    blank: r.slice(0, J_INPUT_N).every(empty),
  }));
}

/** «Рахунки»: назва → вид (порожній вид: основна картка). */
function accKinds_(ss) {
  const a = ss.getSheetByName(ACC_SHEET), out = {};
  if (!a) return out;
  a.getRange(ACC_FIRST, 1, ACC_LAST - ACC_FIRST + 1, 4).getValues().forEach(r => { if (String(r[0]).trim()) out[String(r[0]).trim()] = String(r[3] || '') || K_MAIN; });
  return out;
}
