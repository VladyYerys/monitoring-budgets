/** 28_Dash.gs: «Charts» (комп'ютер) і «charts_mob» (телефон): темні дашборди в стилі Grafana, як в особистій таблиці. */

/*
 * Плитки з головними цифрами, таблиці й графіки; цифри з «Даних графіків» (27_Charts.gs). Вводити нічого не треба,
 * крім колонки «Ліміт на місяць ✎» у таблиці «чорних дір» на «Charts» (порожньо: порівняння з середнім місяцем кварталу).
 * Будуються наново після зміни модуля «💳 Кредитний ліміт» чи кількості груп і «чорних дір»; ліміти «чорних дір» зберігаються.
 * Свої графіки скрипт впізнає за назвою (dashTitles_) і перебудовує; ваші графіки не чіпає.
 * «charts_mob»: три вузькі стовпчики (B:D, G:I, L:N) під екран телефона; гортайте вбік.
 * Рядки під графіками мають висоту 21 px: графік займає рівно стільки рядків, скільки треба (висота ÷ 21).
 */

const DK = { bg: '#111217', tile: '#181B1F', head: '#22252B', text: '#D8D9DA', muted: '#9FA7B3', line: '#2C3235', lineHead: '#3D4250', input: '#1F2633' };
/** Кольори серій (перевірено на розрізнення, зокрема для дальтонізму, на темному тлі). */
const DKC = { blue: '#3987E5', orange: '#D95926', green: '#199E70', amber: '#C98500', pink: '#D55181', lime: '#0CA30C', violet: '#9085E9', red: '#E66767',
  navy: '#184F95', gray: '#9FA7B3' };
const DK_TYPES = [DKC.blue, DKC.orange, DKC.green, DKC.amber];
const DK_HOLES = [DKC.blue, DKC.orange, DKC.green, DKC.amber, DKC.pink, DKC.lime, DKC.violet, DKC.red];
const F_MONEY = '#,##0" ₴";"−"#,##0" ₴";0" ₴"', F_NUM = '#,##0', F_PCT = '"+"0%;"−"0%;0%', F_P = '0%';
/** Для клітинок з текстом формули: числовий формат на текст не впливає (формат «@» міг би завадити формулі). */
const F_TEXT = F_NUM;
const ROW_PX = 21;
const DASH_WIDTHS = [20].concat(Array(16).fill(88), [20]);                  // A, B…Q, R
const MOB_WIDTHS = [10, 160, 84, 92, 10, 20, 160, 84, 92, 10, 20, 160, 84, 92, 10];
const DASH_HOLE_HEAD = 105, DASH_LIMIT_COL = 8;                            // таблиця «чорних дір»: заголовок у рядку 105, «Ліміт на місяць ✎» у H
const DASH_LIMIT_HEAD = 'Ліміт на місяць ✎';
/** Назви графіків версії 8 («Головне», «📱 Телефон»): прибрати при переході на версію 9. */
const CH_OLD_TITLES = ['📈 На картці по днях: факт, за планом, як зараз', '📊 Витрати по місяцях: план і факт', '📈 На картці: факт, прогноз і межа ліміту',
  '📈 На картці: факт і прогноз', '📈 Цей місяць: витрачено з моїх і рівний темп', '🍩 Цей місяць за типами', '📊 «Чорні діри» за 6 місяців',
  '🍩 «Чорні діри» цього місяця', '📊 Витрати: з моїх чи з ліміту, 6 місяців', '📉 Борг за лімітом і орієнтир погашення',
  '📊 Групи: цей місяць проти середнього за 3', '📊 Чому з ліміту: 6 місяців, топ-8'];

/* ---------- будівельні блоки ---------- */

/** Значення чи формула (рядок з «=») у клітинку. */
function dashPut_(rng, v) { if (typeof v === 'string' && v.charAt(0) === '=') setF_(rng, [[v]]); else rng.setValue(v); return rng; }
/** Текст як формула (для стовпчиків формул). */
function dashQ_(s) { return '="' + String(s).replace(/"/g, '""') + '"'; }
/** Клітинка «Даних графіків» з іншого аркуша: колонка (номер) і рядок. */
function dashD_(c, r) { return chRef_(chCol_(c, r)); }

/** Об'єднати nr × nc від (r, c) і записати текст чи формулу в лівий верхній кут. Повертає об'єднаний діапазон. */
function dashText_(sh, r, c, nr, nc, v) {
  const rg = sh.getRange(r, c, nr, nc);
  if (nr > 1 || nc > 1) rg.merge();
  dashPut_(sh.getRange(r, c), v);
  return rg;
}

/** Чистий темний аркуш потрібного розміру: без наших графіків, об'єднань, сітки; ширини колонок, висота рядків 21 px. */
function dashFrame_(ss, name, rows, widths) {
  const sh = sheetOf_(ss, name), cols = widths.length;
  dashRemoveOurs_(sh);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  ensureSize_(sh, rows, cols);
  if (sh.getMaxRows() > rows) sh.deleteRows(rows + 1, sh.getMaxRows() - rows);
  if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setRowHeights(1, rows, ROW_PX);
  sh.getRange(1, 1, rows, cols).setBackground(DK.bg).setFontColor(DK.text).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle').setWrap(false);
  sh.setHiddenGridlines(true);
  sh.setTabColor(DK.tile);
  return sh;
}

/** Заголовок розділу (рядок r) і пояснення під ним (note: необов'язково), на w колонок від c. */
function dashSection_(sh, r, c, w, title, note) {
  sh.getRange(r, c, 1, w).merge();
  dashPut_(sh.getRange(r, c), title);
  sh.getRange(r, c).setFontSize(13).setFontWeight('bold').setVerticalAlignment('bottom');
  sh.setRowHeight(r, 32);
  if (!note) return;
  sh.getRange(r + 1, c, 1, w).merge();
  dashPut_(sh.getRange(r + 1, c), note);
  sh.getRange(r + 1, c, 1, w).setFontColor(DK.muted).setWrap(true).setVerticalAlignment('top');
  sh.setRowHeight(r + 1, 36);
}

/**
 * Плитка: n рядків × w колонок від (r, c). t = { title, big, bigFmt, bigSize, rows: [[підпис, значення, формат, ширина значення]] }.
 * Підпис і заголовок можуть бути формулами.
 */
function dashTile_(sh, r, c, w, n, t, vs0) {
  sh.getRange(r, c, n, w).setBackground(DK.tile).setBorder(null, true, null, true, null, null, DK.bg, dashThick_());
  sh.getRange(r, c, 1, w).merge();
  dashPut_(sh.getRange(r, c), t.title);
  sh.getRange(r, c, 1, w).setFontSize(9).setFontWeight('bold').setFontColor(DK.muted).setWrap(true);
  let y = r + 1;
  if (t.big) {
    const b = sh.getRange(y, c, 2, w).merge();
    dashPut_(sh.getRange(y, c), t.big);
    b.setFontSize(t.bigSize || 24).setFontWeight('bold').setNumberFormat(t.bigFmt || F_MONEY).setHorizontalAlignment('left').setWrap(!!t.bigSize);
    y += 2;
  }
  (t.rows || []).forEach(x => {
    const vs = x[3] || vs0 || 1, ls = w - vs;
    if (ls > 1) sh.getRange(y, c, 1, ls).merge();
    dashPut_(sh.getRange(y, c), x[0]);
    sh.getRange(y, c, 1, ls).setFontColor(DK.muted);
    const v = sh.getRange(y, c + ls, 1, vs);
    if (vs > 1) v.merge();
    dashPut_(sh.getRange(y, c + ls), x[1]);
    v.setFontWeight('bold').setHorizontalAlignment('right').setNumberFormat(x[2] || F_MONEY);
    y++;
  });
}

/** Товста межа кольору тла: проміжок між сусідніми плитками й таблицями (у Таблицях немає відступів у клітинках). */
function dashThick_() { return SpreadsheetApp.BorderStyle ? SpreadsheetApp.BorderStyle.SOLID_THICK : null; }

/** Рядок плиток: висоти (заголовок, велике число, підписи). */
function dashTileRows_(sh, r, n, big) {
  sh.setRowHeight(r, big ? 40 : 30);
  if (big) { sh.setRowHeight(r + 1, 30); sh.setRowHeight(r + 2, 30); }
  for (let y = r + (big ? 3 : 1); y < r + n; y++) sh.setRowHeight(y, 25);
}

/**
 * Таблиця з рядка r: заголовок і n рядків. cols = [{ head, span, f: i → формула, fmt, input }]; total: останній рядок жирний.
 * Стовпчик input: ваші числа (не формули), типово порожньо.
 */
function dashTable_(sh, r, c, cols, n, total) {
  const w = cols.reduce((a, k) => a + (k.span || 1), 0), bs = SpreadsheetApp.BorderStyle ? SpreadsheetApp.BorderStyle.SOLID : null;
  let x = c;
  cols.forEach((k, j) => {
    const s = k.span || 1;
    if (s > 1) sh.getRange(r, x, n + 1, s).mergeAcross();
    dashPut_(sh.getRange(r, x), k.head);
    if (!k.input) {
      const fs = [];
      for (let i = 0; i < n; i++) fs.push([k.f(i)]);
      setF_(sh.getRange(r + 1, x, n, 1), fs);
    }
    sh.getRange(r, x, 1, s).setHorizontalAlignment(j ? 'center' : 'left');
    sh.getRange(r + 1, x, n, s).setHorizontalAlignment(j ? 'right' : 'left');
    if (k.fmt) sh.getRange(r + 1, x, n, s).setNumberFormat(k.fmt);
    if (k.input) sh.getRange(r + 1, x, n - (total ? 1 : 0), s).setBackground(DK.input).setFontColor(DKC.blue);
    x += s;
  });
  sh.getRange(r, c, 1, w).setBackground(DK.head).setFontWeight('bold').setWrap(true).setBorder(null, null, true, null, null, null, DK.lineHead, bs);
  sh.setRowHeight(r, 34);
  const body = sh.getRange(r + 1, c, n, w);
  body.setBorder(null, null, true, null, null, true, DK.line, bs);
  cols.forEach((k, j) => { if (!k.input) { let xx = c; for (let q = 0; q < j; q++) xx += cols[q].span || 1; sh.getRange(r + 1, xx, n, k.span || 1).setBackground(DK.tile); } });
  if (total) sh.getRange(r + n, c, 1, w).setFontWeight('bold').setBackground(DK.tile).setFontColor(DK.text);
  sh.getRange(r, c, n + 1, w).setBorder(null, true, null, true, null, null, DK.bg, dashThick_());
}

/* ---------- графіки ---------- */

/** Параметри графіка в стилі Grafana: темне тло, приглушені осі, тонкі лінії. */
function dashChartOpts_(s) {
  const txt = (color, size, bold) => { const o = { color: color, fontName: 'Arial', fontSize: size }; if (bold) o.bold = true; return o; };
  const axis = fmt => { const a = { textStyle: txt(DK.muted, 9), titleTextStyle: txt(DK.muted, 9), gridlines: { color: DK.line }, minorGridlines: { color: DK.tile },
    baselineColor: DK.lineHead }; if (fmt) a.format = fmt; return a; };
  const o = { title: s.title, width: s.w, height: s.h, backgroundColor: DK.tile, fontName: 'Arial', titleTextStyle: txt(DK.text, 12, true),
    legend: { position: s.legend || (s.pie ? 'right' : 'bottom'), textStyle: txt(DK.muted, 9) }, chartArea: { backgroundColor: DK.tile }, colors: s.colors };
  if (s.pie) Object.assign(o, { pieHole: 0.55, pieSliceText: 'percentage', pieSliceBorderColor: DK.tile, pieSliceTextStyle: txt('#FFFFFF', 9) });
  else Object.assign(o, { hAxis: axis(s.bar ? F_NUM : s.dates ? 'dd.MM' : ''), vAxis: axis(s.bar ? '' : F_NUM), lineWidth: 2, interpolateNulls: false });
  if (s.series) o.series = s.series;
  if (s.seriesType) o.seriesType = s.seriesType;
  if (s.area) o.areaOpacity = 0.3;
  if (s.stacked) o.isStacked = true;
  return o;
}

/** Вставити графік s ({ title, type, ranges на «Даних графіків», row, col, w, h, … }) на аркуш sh. */
function dashChart_(ss, sh, s) {
  const data = ss.getSheetByName(CH_SHEET), o = dashChartOpts_(s);
  let b = sh.newChart().setChartType(s.type);
  s.ranges.forEach(a => { b = b.addRange(data.getRange(a)); });
  b = b.setNumHeaders(1).setPosition(s.row, s.col, 0, 0);
  Object.keys(o).forEach(k => { b = b.setOption(k, o[k]); });
  sh.insertChart(b.build());
}

/** Діапазон колонки блоку «Даних графіків» з рядка 1 до рядка last. */
function dashR_(c, last) { const l = colStr_(c); return l + '1:' + l + last; }

/** Графіки обох дашбордів: k → { title, type, ranges, colors, … } (без місця й розміру). nG груп, nH «чорних дір». */
function dashSpecs_(nG, nH, cr) {
  const T = Charts.ChartType, dash = [6, 4], dot = [2, 3];
  const g = 1 + Math.max(nG, 1), h = 1 + Math.max(nH, 1);
  return {
    b60: { title: cr ? 'Залишок на кінець дня: останні 60 днів (реальні гроші, мінус = кредит)' : 'Залишок на кінець дня: останні 60 днів', type: T.COMBO,
      ranges: [dashR_(CHD.b60, 61), dashR_(CHD.b60 + 2, 61), dashR_(CHD.b60 + 3, 61)].concat(cr ? [dashR_(CHD.b60 + 4, 61)] : []),
      colors: [DKC.lime, DKC.red, DKC.violet], seriesType: 'area', area: true, dates: true,
      series: cr ? { 2: { type: 'line', lineDashStyle: dash, lineWidth: 2 } } : null },
    typesNow: { title: 'Цей місяць за типами', type: T.PIE, pie: true, ranges: [dashR_(CHD.typ, 5), dashR_(CHD.typ + 1, 5)], colors: DK_TYPES },
    typesQ: { title: 'Попередній квартал за типами', type: T.PIE, pie: true, ranges: [dashR_(CHD.typ, 5), dashR_(CHD.typ + 3, 5)], colors: DK_TYPES },
    groups: { title: 'Витрати за групами: цей місяць проти середнього місяця попереднього кварталу', type: T.BAR, bar: true,
      ranges: [dashR_(CHD.grpS + 1, g), dashR_(CHD.grpS + 2, g), dashR_(CHD.grpS + 5, g)], colors: [DKC.blue, DKC.navy], skip: !nG },
    holesBar: { title: 'Непомітні витрати: цей місяць проти середнього місяця попереднього кварталу', type: T.BAR, bar: true,
      ranges: [dashR_(CHD.hole, h), dashR_(CHD.hole + 1, h), dashR_(CHD.hole + 4, h)], colors: [DKC.blue, DKC.navy], skip: !nH },
    holes6: { title: 'Непомітні витрати по місяцях (останні 6, поточний по сьогодні)', type: T.COLUMN, stacked: true,
      ranges: [colStr_(CHD.h6) + '1:' + colStr_(CHD.h6 + Math.max(nH, 1)) + '7'], colors: DK_HOLES, skip: !nH },
    holesPie: { title: 'Цей місяць за «чорними дірами»', type: T.PIE, pie: true, ranges: [dashR_(CHD.hole, h), dashR_(CHD.hole + 1, h)], colors: DK_HOLES, skip: !nH },
    pace: { title: 'Мої гроші цього місяця: витрачено наростом проти рівного темпу', type: T.LINE,
      ranges: [dashR_(CHD.pace, 32), dashR_(CHD.pace + 1, 32), dashR_(CHD.pace + 2, 32), dashR_(CHD.pace + 3, 32)], colors: [DKC.orange, DKC.gray, DKC.green],
      series: { 1: { lineDashStyle: dash } } },
    fc: { title: cr ? 'Залишок: факт (суцільна) і прогноз (пунктир), межа кредитного ліміту' : 'Залишок: факт (суцільна) і прогноз (пунктир)', type: T.LINE, dates: true,
      ranges: [dashR_(CHD.fc, 93), dashR_(CHD.fc + 1, 93), dashR_(CHD.fc + 2, 93)].concat(cr ? [dashR_(CHD.fc + 3, 93)] : []), colors: [DKC.blue, DKC.blue, DKC.red],
      series: { 1: { lineDashStyle: dash }, 2: { lineDashStyle: dot, lineWidth: 1 } } },
    ml: { title: 'Витрати: з моїх чи з ліміту, 6 місяців', type: T.COLUMN, stacked: true, ranges: [colStr_(CHD.ml) + '1:' + colStr_(CHD.ml + 2) + '7'],
      colors: [DKC.blue, DKC.orange], skip: !cr },
    debt: { title: 'Борг за лімітом: факт, прогноз і орієнтир погашення', type: T.LINE, ranges: [colStr_(CHD.debt) + '1:' + colStr_(CHD.debt + 3) + '11'],
      colors: [DKC.orange, DKC.orange, DKC.green], series: { 1: { lineDashStyle: dash } }, skip: !cr },
    why: { title: 'Чому брав з ліміту: за 6 місяців (топ-8)', type: T.BAR, bar: true, legend: 'none', ranges: [dashR_(CHD.whyS + 1, 9), dashR_(CHD.whyS + 4, 9)],
      colors: [DKC.orange], skip: !cr },
  };
}

/** Короткі назви графіків «charts_mob». */
const MOB_TITLES = { b60: 'Залишок на кінець дня: 60 днів', typesNow: 'Цей місяць за типами', holes6: '«Чорні діри»: 6 місяців', holesPie: 'Цей місяць за «чорними дірами»',
  pace: 'Мої гроші: витрачено проти темпу', fc: 'Залишок: факт і прогноз', ml: 'З моїх чи з ліміту: 6 міс.', debt: 'Борг за лімітом і орієнтир' };

/** Усі назви наших графіків (обидва варіанти модуля ліміту): за ними скрипт прибирає свої. */
function dashTitles_() {
  const out = [];
  [true, false].forEach(cr => { const s = dashSpecs_(1, 1, cr); Object.keys(s).forEach(k => { if (out.indexOf(s[k].title) < 0) out.push(s[k].title); }); });
  Object.keys(MOB_TITLES).forEach(k => { if (out.indexOf(MOB_TITLES[k]) < 0) out.push(MOB_TITLES[k]); });
  return out;
}

function dashRemoveOurs_(sh, titles) {
  const t = titles || dashTitles_();
  sh.getCharts().forEach(c => { try { if (t.indexOf(String(c.getOptions().get('title') || '')) >= 0) sh.removeChart(c); } catch (e) { /* ignore */ } });
}

/** Поставити графіки: [[ключ, рядок, колонка, ширина, висота]]; mob: короткі назви. */
function dashCharts_(ss, sh, specs, place, mob) {
  let n = 0;
  place.forEach(p => {
    const s = specs[p[0]];
    if (!s || s.skip) return;
    dashChart_(ss, sh, Object.assign({}, s, { title: mob ? MOB_TITLES[p[0]] : s.title, row: p[1], col: p[2], w: p[3], h: p[4], legend: mob && s.pie ? 'bottom' : s.legend }));
    n++;
  });
  return n;
}

/* ---------- «Charts» ---------- */

/** Рядок заголовка таблиці «чорних дір» на «Charts» (дані з наступного рядка); 0: аркуша немає. */
function dashHoleRow_(ss) { return ss.getSheetByName(DASH_SHEET) ? DASH_HOLE_HEAD + 1 : 0; }

/** Ліміти «чорних дір», уписані на «Charts»: [{ name, v }] за порядком рядків. */
function dashLimitsRead_(ss) {
  const sh = ss.getSheetByName(DASH_SHEET);
  if (!sh || sh.getMaxRows() < DASH_HOLE_HEAD + CH_HOLES) return [];
  if (String(sh.getRange(DASH_HOLE_HEAD, DASH_LIMIT_COL).getValue()) !== DASH_LIMIT_HEAD) return [];
  const v = sh.getRange(DASH_HOLE_HEAD + 1, 2, CH_HOLES, DASH_LIMIT_COL - 1).getValues();
  return v.map(r => ({ name: String(r[0]), v: r[DASH_LIMIT_COL - 2] }));
}

/** Повернути ліміти: за назвою «чорної діри»; назви немає серед нових (перейменували) чи не прочиталась: за рядком. */
function dashLimitsWrite_(ss, sh, old) {
  if (!old.length) return;
  const names = ss.getSheetByName(DICT_SHEET).getRange(D_HOLES_FIRST, colNum_(DC.holes), CH_HOLES, 1).getValues().map(r => String(r[0]));
  const out = names.map((n, i) => {
    const hit = old.find(o => o.name && o.name === n), same = old[i];
    return [hit ? hit.v : same && (!same.name || names.indexOf(same.name) < 0) ? same.v : ''];
  });
  sh.getRange(DASH_HOLE_HEAD + 1, DASH_LIMIT_COL, CH_HOLES, 1).setValues(out);
}

function buildDash_(ss) {
  if (!ss.getSheetByName(CH_SHEET)) buildChartsData_(ss);
  const cr = modOn_('credit'), K = chD_, typ = (r, i) => dashD_(CHD.typ + i, r), hole = (i, r) => dashD_(CHD.hole + i, r || 2 + CH_HOLES);
  const old = dashLimitsRead_(ss), rows = cr ? 221 : 177;
  const sh = dashFrame_(ss, DASH_SHEET, rows, DASH_WIDTHS);
  // шапка
  sh.getRange('B1:L1').merge();
  sh.getRange('B1').setValue('Фінанси зараз').setFontSize(18).setFontWeight('bold');
  dashText_(sh, 1, 13, 1, 3, 'станом на').setFontColor(DK.muted).setHorizontalAlignment('right');
  sh.getRange('P1:Q1').merge();
  dashPut_(sh.getRange('P1'), '=' + K('today'));
  sh.getRange('P1').setNumberFormat('dd.mm.yyyy').setFontWeight('bold');
  dashText_(sh, 2, 2, 1, 16, cr ? 'Залишок: ваші реальні гроші на кінець дня (з мінусом, якщо ви в кредиті). «Доступно з лімітом», як показує банк, окремо, в останній плитці.'
    : 'Залишок: ваші гроші на картці на кінець дня. Цифри рахують формули з журналу, «Днів» і «Довідника»: нічого вводити не треба.').setFontColor(DK.muted);
  sh.setRowHeight(1, 45); sh.setRowHeight(2, 24); sh.setRowHeight(3, 10);
  // плитки
  dashTileRows_(sh, 4, 5, true);
  dashTile_(sh, 4, 2, 4, 5, { title: 'ЗАЛИШОК НА КІНЕЦЬ ДНЯ', big: '=' + K('bal'), rows: [['на кінець минулого місяця', '=' + K('balPm')], ['зміна за місяць', '=' + K('change')]] });
  dashTile_(sh, 4, 6, 4, 5, { title: 'ВИТРАЧЕНО ЦЬОГО МІСЯЦЯ', big: '=' + typ(2, 1), rows: [['минулий місяць', '=' + typ(2, 2)], ['="сер. місяць "&' + K('qLabel'), '=' + typ(2, 4)]] });
  dashTile_(sh, 4, 10, 4, 5, { title: 'ДОХІД ЦЬОГО МІСЯЦЯ', big: '=' + typ(7, 1), rows: [['заощаджено цього місяця', '=' + typ(4, 1)],
    ['податки і борги цього місяця', '=' + typ(5, 1) + '+' + typ(3, 1)]] });
  dashTile_(sh, 4, 14, 4, 5, cr ? { title: 'ДОСТУПНО З КРЕДИТНИМ ЛІМІТОМ (як у банку)', big: '=' + K('avail'), rows: [['кредитний ліміт', '=' + K('limit')],
    ['використано кредитних грошей', '=' + K('debt')]] }
    : { title: '≈ МОЖНА НА ДЕНЬ ДО КІНЦЯ МІСЯЦЯ', big: '=' + K('perDay'), rows: [['доступно моїх сьогодні', '=' + K('mine')], ['днів до кінця місяця', '=' + K('daysLeft'), F_NUM]] });
  sh.setRowHeight(9, 16);
  // цей місяць проти попереднього кварталу
  dashSection_(sh, 51, 2, 16, '="Цей місяць ("&' + K('mName') + '&", з 1-го по сьогодні) проти попереднього кварталу ("&' + K('qLabel') + '&")"',
    '="Місяць ще триває: минуло "&' + K('daysPassed') + '&" дн. Порівняння з середнім місяцем попереднього кварталу (квартал ÷ 3)."');
  sh.getRange(53, 10, 2, 8).merge();
  dashPut_(sh.getRange(53, 10), '="Що зросло: цього місяця вже більше, ніж у середньому за місяць "&' + K('qLabel'));
  sh.getRange(53, 10).setFontSize(12).setFontWeight('bold').setWrap(true);
  const up = (i, k) => '=' + dashD_(CHD.up + k, 2 + i);
  dashTable_(sh, 55, 10, [{ head: 'Категорія витрат', span: 4, f: i => up(i, 1) }, { head: 'Цей місяць', f: i => up(i, 2), fmt: F_NUM },
    { head: 'Сер. місяць кварталу', f: i => up(i, 3), fmt: F_NUM }, { head: 'Більше на', f: i => up(i, 4), fmt: F_NUM }, { head: 'Зміна', f: i => up(i, 5), fmt: F_PCT }], CH_UP);
  // зведення: групи й типи
  dashSection_(sh, 80, 2, 16, 'Зведення: групи витрат і типи');
  const gs = (i, k) => '=' + (i < CH_GROUPS ? dashD_(CHD.grpS + k, 2 + i) : k === 1 ? dashD_(CHD.grp, 2 + CH_GROUPS) : dashD_(CHD.grp + k - 1, 2 + CH_GROUPS));
  dashTable_(sh, 81, 2, [{ head: 'Група витрат (за спаданням цього місяця)', span: 4, f: i => gs(i, 1) }, { head: 'Цей місяць', f: i => gs(i, 2), fmt: F_NUM },
    { head: 'Минулий місяць', f: i => gs(i, 3), fmt: F_NUM }, { head: '="Квартал "&' + K('qLabel'), f: i => gs(i, 4), fmt: F_NUM },
    { head: '="Сер. місяць "&' + K('qLabel'), f: i => gs(i, 5), fmt: F_NUM }], CH_GROUPS + 1, true);
  dashTable_(sh, 81, 10, [{ head: 'Тип', span: 4, f: i => '=' + typ(2 + i, 0) }, { head: 'Цей місяць', f: i => '=' + typ(2 + i, 1), fmt: F_NUM },
    { head: 'Минулий місяць', f: i => '=' + typ(2 + i, 2), fmt: F_NUM }, { head: '="Квартал "&' + K('qLabel'), f: i => '=' + typ(2 + i, 3), fmt: F_NUM },
    { head: '="Сер. місяць "&' + K('qLabel'), f: i => '=' + typ(2 + i, 4), fmt: F_NUM }], 6);
  // «чорні діри»
  dashSection_(sh, 96, 2, 16, 'Непомітні витрати: «чорні діри»', 'Кава, доставка, підписки, таксі, комісії: за категоріями («' + DICT_SHEET +
    '», колонка D) або міткою в журналі (колонка F). «' + DASH_LIMIT_HEAD + '»: впишіть свою суму; без ліміту діру порівнюємо із середнім місяцем кварталу.');
  sh.setRowHeight(98, 8);
  dashTileRows_(sh, 99, 5, true);
  dashTile_(sh, 99, 2, 4, 5, { title: 'НЕПОМІТНІ ВИТРАТИ ЦЬОГО МІСЯЦЯ', big: '=' + hole(1), rows: [['операцій у журналі', '=' + hole(5), F_NUM], ['минулий місяць', '=' + hole(2)]] });
  dashTile_(sh, 99, 6, 4, 5, { title: 'СЕРЕДНІЙ МІСЯЦЬ ПОПЕРЕДНЬОГО КВАРТАЛУ', big: '=' + hole(4), rows: [['квартал усього', '=' + hole(3)],
    ['цей місяць мінус середній', '=' + hole(1) + '-' + hole(4)]] });
  dashTile_(sh, 99, 10, 4, 5, { title: 'ВИКОРИСТАНО ЛІМІТІВ (ДІРИ З ЛІМІТОМ)', big: '=IF(' + K('hUsed') + '="","○ не задано",' + K('hUsed') + ')', bigFmt: F_P, rows: [['сума лімітів', '=N(' + hole(6) + ')'],
    ['лишилось у межах лімітів', '=' + K('hLeft')]] });
  dashTile_(sh, 99, 14, 4, 5, { title: 'НАЙБІЛЬША ДІРА ЦЬОГО МІСЯЦЯ', big: '=' + K('hBig'), bigFmt: F_TEXT, bigSize: 14, rows: [['її сума', '=' + K('hBigSum')],
    ['дір понад ліміт / увага', '=' + K('hOver') + '&" / "&' + K('hWarn'), F_TEXT]] });
  const ht = (i, k) => '=' + (i < CH_HOLES ? hole(k, 2 + i) : hole(k));
  dashTable_(sh, DASH_HOLE_HEAD, 2, [{ head: 'Непомітна витрата', span: 3, f: i => ht(i, 0) }, { head: 'Цей місяць', f: i => ht(i, 1), fmt: F_NUM },
    { head: 'Операцій', f: i => ht(i, 5), fmt: F_NUM }, { head: 'Сер. місяць кварталу', f: i => ht(i, 4), fmt: F_NUM },
    { head: DASH_LIMIT_HEAD, input: true, fmt: F_NUM }, { head: 'Статус', f: i => ht(i, 9) }], CH_HOLES + 1, true);
  setF_(sh.getRange(DASH_HOLE_HEAD + 1 + CH_HOLES, DASH_LIMIT_COL), [['=' + hole(6)]]);
  dashLimitsWrite_(ss, sh, old);
  dashText_(sh, DASH_HOLE_HEAD + 2 + CH_HOLES, 2, 2, 8, '✎ Ліміт: ваша сума на місяць. ▲ увага: від 80 %, ▲▲ понад: більше за ліміт (без ліміту: за середній місяць кварталу).')
    .setFontColor(DK.muted).setWrap(true).setVerticalAlignment('top');
  // мої гроші, ліміт і прогноз
  dashSection_(sh, 147, 2, 16, cr ? 'Мої гроші, ліміт і прогноз' : 'Мої гроші й прогноз', cr
    ? '«Мої»: гроші з доходу мінус відкладене на погашення ліміту; коли їх не вистачає, витрата йде з кредитного ліміту. Прогноз: реальний залишок плюс планові операції журналу (зокрема регулярні платежі й податки ФОП).'
    : '«Мої»: гроші з доходу бюджетного місяця мінус уже списане. Прогноз: залишок плюс планові операції журналу (зокрема регулярні платежі й податки ФОП).');
  dashTileRows_(sh, 150, 5, true);
  const fcTile = { title: 'ПРОГНОЗ: ЗАЛИШОК НА КІНЕЦЬ МІСЯЦЯ', big: '=' + K('fcEnd'), rows: [['найнижча точка', '=' + K('lowText'), F_TEXT, 2],
    ['мої до кінця місяця', '=' + K('mineEndText'), F_TEXT, 2]] };
  dashTile_(sh, 150, 2, 4, 5, { title: '💚 ДОСТУПНО МОЇХ', big: '=' + K('mine'), rows: [['≈ можна на день до кінця місяця', '=' + K('perDay')],
    ['мої гроші місяця (факт + план)', '=' + K('mineMonth')]] });
  if (cr) {
    dashTile_(sh, 150, 6, 4, 5, { title: 'З ЛІМІТУ ЦЬОГО МІСЯЦЯ', big: '=' + K('fromLimit'), rows: [['операцій з ліміту', '=' + K('limitOps'), F_NUM],
      ['з них без причини', '=' + K('noReason')]] });
    dashTile_(sh, 150, 10, 4, 5, { title: 'ПОГАШЕННЯ ЛІМІТУ (ОРІЄНТИР: «' + SETTINGS_SHEET.toUpperCase() + '»)', big: '=' + K('goalText'), bigFmt: F_TEXT, bigSize: 14,
      rows: [['борг за лімітом зараз', '=' + K('debt')], ['повне погашення*', '=' + K('repayText'), F_TEXT, 2]] });
    dashTile_(sh, 150, 14, 4, 5, fcTile);
  } else {
    dashTile_(sh, 150, 6, 4, 5, fcTile);
    dashTile_(sh, 150, 10, 4, 5, { title: '«ЯК ЗАРАЗ»: НА КІНЕЦЬ МІСЯЦЯ', big: '=' + K('nowEnd'), rows: [['витрачаєте ≈ на день', '=' + K('pace30')],
      ['найнижча «як зараз»', '=' + K('lowNowText'), F_TEXT, 2]] });
    dashTile_(sh, 150, 14, 4, 5, { title: 'СПИСАНО З МОЇХ ЦЬОГО МІСЯЦЯ', big: '=' + K('spentMine'), rows: [['ще заплановано списань', '=' + K('planOut')],
      ['у середньому за день', '=' + K('spentMine') + '/' + K('daysPassed')]] });
  }
  // чому з ліміту
  if (cr) {
    dashSection_(sh, 198, 2, 16, 'Чому брав з ліміту: 6 місяців');
    const ws = (i, k) => '=' + (i < 16 ? dashD_(CHD.whyS + k, 2 + i) : dashD_(CHD.why + (k === 1 ? 0 : k === 4 ? 3 : k - 1), 2 + CH_WHY));
    dashTable_(sh, 199, 10, [{ head: 'Причина', span: 4, f: i => ws(i, 1) }, { head: 'Цей місяць', f: i => ws(i, 2), fmt: F_NUM },
      { head: 'Операцій', f: i => ws(i, 3), fmt: F_NUM }, { head: 'За 6 місяців', f: i => ws(i, 4), fmt: F_NUM },
      { head: 'Частка', f: i => i < 16 ? ws(i, 5) : '=""', fmt: F_P }], 17, true);
    dashText_(sh, 217, 10, 2, 8, '* Повне погашення: борг ÷ середнє погашення за 3 місяці (або орієнтир з «' + SETTINGS_SHEET + '», якщо погашень ще не було).')
      .setFontColor(DK.muted).setWrap(true).setVerticalAlignment('top');
  }
  // графіки
  const specs = dashSpecs_(chartsGroups_(ss).length, chartsHolesN_(ss), cr);
  const W2 = 690;
  const n = dashCharts_(ss, sh, specs, [['b60', 10, 2, 1408, 420], ['typesNow', 31, 2, W2, 399], ['typesQ', 31, 10, W2, 399], ['groups', 53, 2, W2, 546],
    ['holesBar', DASH_HOLE_HEAD, 10, W2, 441], ['holes6', 127, 2, 1408, 399], ['pace', 156, 2, W2, 420], ['fc', 156, 10, W2, 420],
    ['ml', 177, 2, W2, 420], ['debt', 177, 10, W2, 420], ['why', 199, 2, W2, 441]], false);
  chartsHoleLimits_(ss);
  return n;
}

/* ---------- «charts_mob» ---------- */

function buildMob_(ss) {
  if (!ss.getSheetByName(CH_SHEET)) buildChartsData_(ss);
  const cr = modOn_('credit'), K = chD_, typ = (r, i) => dashD_(CHD.typ + i, r), hole = (i, r) => dashD_(CHD.hole + i, r || 2 + CH_HOLES);
  const sh = dashFrame_(ss, MOB_SHEET, cr ? 107 : 90, MOB_WIDTHS);
  const title = (c, t) => { sh.getRange(1, c, 1, 2).merge(); sh.getRange(1, c).setValue(t).setFontSize(16).setFontWeight('bold'); };
  const sub = (r, c, t) => { sh.getRange(r, c, 1, 3).merge(); dashPut_(sh.getRange(r, c), t); sh.getRange(r, c).setFontSize(12).setFontWeight('bold'); };
  sh.setRowHeight(1, 36); sh.setRowHeight(2, 8); sh.setRowHeight(10, 8); sh.setRowHeight(17, 8);
  dashTileRows_(sh, 3, 7, true); dashTileRows_(sh, 11, 6, false);
  // 1: фінанси зараз
  title(2, 'Фінанси зараз');
  dashPut_(sh.getRange(1, 4), '=' + K('today'));
  sh.getRange(1, 4).setNumberFormat('dd.mm.yyyy').setFontWeight('bold').setHorizontalAlignment('right');
  dashTile_(sh, 3, 2, 3, 7, { title: 'ЗАЛИШОК НА КІНЕЦЬ ДНЯ', big: '=' + K('bal'), rows: [['на кінець минулого місяця', '=' + K('balPm')], ['зміна за місяць', '=' + K('change')]]
    .concat(cr ? [['доступно з лімітом', '=' + K('avail')], ['кредитний ліміт', '=' + K('limit')]] : []) }, 2);
  dashTile_(sh, 11, 2, 3, 6, { title: '="ЦЕЙ МІСЯЦЬ: "&UPPER(' + K('mName') + ')', rows: [['Витрачено', '=' + typ(2, 1)], ['Дохід', '=' + typ(7, 1)],
    ['Заощаджено', '=' + typ(4, 1)], ['Податки і борги', '=' + typ(5, 1) + '+' + typ(3, 1)], ['="витрати, сер. "&' + K('qLabel'), '=' + typ(2, 4)]] }, 2);
  sub(35, 2, 'Групи витрат: цей місяць');
  const gs = (i, k) => i < CH_GROUPS ? dashD_(CHD.grpS + k, 2 + i) : k === 1 ? dashD_(CHD.grp, 2 + CH_GROUPS) : dashD_(CHD.grp + k - 1, 2 + CH_GROUPS);
  dashTable_(sh, 36, 2, [{ head: 'Група', f: i => '=' + gs(i, 1) }, { head: 'Цей місяць', f: i => '=' + gs(i, 2), fmt: F_NUM },
    { head: 'Проти сер.', f: i => '=IF(OR(' + gs(i, 1) + '="",N(' + gs(i, 2) + ')=0),"",IF(N(' + gs(i, 5) + ')=0,"нова",' + gs(i, 2) + '/' + gs(i, 5) + '-1))', fmt: F_PCT }],
  CH_GROUPS + 1, true);
  sub(68, 2, 'За типами');
  dashTable_(sh, 69, 2, [{ head: 'Тип', f: i => '=' + typ(2 + i, 0) }, { head: 'Цей місяць', f: i => '=' + typ(2 + i, 1), fmt: F_NUM },
    { head: '="Сер. "&' + K('qLabel'), f: i => '=' + typ(2 + i, 4), fmt: F_NUM }], 5, true);
  sub(76, 2, 'Що зросло');
  const up = (i, k) => '=' + dashD_(CHD.up + k, 2 + i);
  dashTable_(sh, 77, 2, [{ head: 'Категорія', f: i => up(i, 1) }, { head: 'Цей місяць', f: i => up(i, 2), fmt: F_NUM }, { head: 'Більше на', f: i => up(i, 4), fmt: F_NUM }], CH_UP);
  sh.getRange(88, 2, 2, 3).merge();
  dashPut_(sh.getRange(88, 2), '="Більше, ніж у середньому за місяць "&' + K('qLabel') + '&". Місяць ще триває: минуло "&' + K('daysPassed') + '&" дн."');
  sh.getRange(88, 2).setFontColor(DK.muted).setWrap(true).setVerticalAlignment('top');
  // 2: непомітні витрати
  title(7, 'Непомітні витрати');
  dashTile_(sh, 3, 7, 3, 7, { title: '«ЧОРНІ ДІРИ» ЦЬОГО МІСЯЦЯ', big: '=' + hole(1), rows: [['операцій у журналі', '=' + hole(5), F_NUM], ['минулий місяць', '=' + hole(2)],
    ['="сер. місяць "&' + K('qLabel'), '=' + hole(4)], ['використано лімітів', '=' + K('hUsed'), F_P]] }, 2);
  dashTile_(sh, 11, 7, 3, 6, { title: 'СТАТУС ДІР', rows: [['▲▲ понад ліміт', '=' + K('hOver'), F_NUM], ['▲ увага (80…100 %)', '=' + K('hWarn'), F_NUM],
    ['● у нормі', '=' + K('hOk'), F_NUM], ['найбільша діра', '=' + K('hBig'), F_TEXT], ['її сума', '=' + K('hBigSum')]] }, 2);
  const ht = (i, k) => '=' + (i < CH_HOLES ? hole(k, 2 + i) : k ? hole(k) : dashQ_('Усі'));
  dashTable_(sh, 18, 7, [{ head: 'Діра', f: i => ht(i, 0) }, { head: 'Цей місяць', f: i => ht(i, 1), fmt: F_NUM }, { head: 'Статус', f: i => ht(i, 9) }], CH_HOLES + 1, true);
  // 3: мої гроші
  title(12, cr ? 'Мої гроші й ліміт' : 'Мої гроші й прогноз');
  dashTile_(sh, 3, 12, 3, 7, { title: '💚 ДОСТУПНО МОЇХ', big: '=' + K('mine'), rows: [['≈ можна на день', '=' + K('perDay')], ['мої гроші місяця', '=' + K('mineMonth')]]
    .concat(cr ? [['з ліміту цього місяця', '=' + K('fromLimit')], ['відкладено на погашення', '=' + K('setAside')]]
      : [['списано з моїх', '=' + K('spentMine')], ['ще заплановано списань', '=' + K('planOut')]]) }, 2);
  dashTile_(sh, 11, 12, 3, 6, { title: cr ? '📅 ПРОГНОЗ І БОРГ' : '📅 ПРОГНОЗ', rows: [['залишок на кінець місяця', '=' + K('fcEnd')], ['найнижча точка', '=' + K('lowText'), F_TEXT],
    ['мої до кінця місяця', '=' + K('mineEndText'), F_TEXT]].concat(cr ? [['борг за лімітом зараз', '=' + K('debt')], ['повне погашення*', '=' + K('repayText'), F_TEXT]]
    : [['«як зараз» на кінець місяця', '=' + K('nowEnd')], ['витрачаєте ≈ на день', '=' + K('pace30')]]) }, 2);
  if (cr) {
    const ws = (i, k) => '=' + (i < 8 ? dashD_(CHD.whyS + k, 2 + i) : dashD_(CHD.why + (k === 1 ? 0 : k === 4 ? 3 : k - 1), 2 + CH_WHY));
    dashTable_(sh, 18, 12, [{ head: 'Чому з ліміту', f: i => ws(i, 1) }, { head: 'Цей міс.', f: i => ws(i, 2), fmt: F_NUM }, { head: '6 міс.', f: i => ws(i, 4), fmt: F_NUM }], 9, true);
    dashText_(sh, 28, 12, 1, 3, '* ≈ борг ÷ середнє погашення за 3 місяці (або орієнтир)').setFontColor(DK.muted).setFontSize(9);
  }
  const specs = dashSpecs_(chartsGroups_(ss).length, chartsHolesN_(ss), cr), c0 = cr ? 30 : 18, W = 336;
  return dashCharts_(ss, sh, specs, [['b60', 18, 2, W, 336], ['typesNow', 51, 2, W, 336], ['holes6', 29, 7, W, 391], ['holesPie', 49, 7, W, 336],
    ['pace', c0, 12, W, 350], ['fc', c0 + 18, 12, W, 370], ['ml', c0 + 37, 12, W, 384], ['debt', c0 + 57, 12, W, 384]], true);
}

/** Обидва дашборди наново (модуль ліміту, групи чи «чорні діри» змінились). Кількість графіків. */
function dashRefresh_(ss) {
  if (!ss.getSheetByName(DAYS_SHEET) || !ss.getSheetByName(DICT_SHEET)) return 0;
  if (!ss.getSheetByName(CH_SHEET) || chartsDataOld_(ss)) buildChartsData_(ss);
  const n = buildDash_(ss) + buildMob_(ss);
  PropertiesService.getDocumentProperties().setProperty('ch_shape', chartsShape_(ss));
  return n;
}
