/** 22_History.gs: «Історія місяців» і «Історія категорій»: закриті місяці значеннями з «Місяців». */

/*
 * Що записується: один рядок на місяць (підсумки) + рядки «місяць × категорія» (план і факт кожної категорії).
 * Значення, не формули: пізні правки в журналі історію не змінюють. Перезаписати місяць свідомо: меню.
 * Звідки: «Місяці» рахують обраний рік; якщо там інший рік, скрипт на мить ставить потрібний, читає й повертає назад.
 */

const H_CAT_SHEET = 'Історія категорій';                   // без «:» (Excel не дозволяє двокрапку в назві аркуша)
const H_COLS = ['Місяць', T_SPEND, T_SAVE, T_TAX, T_DEBT, 'Усі списання', 'Всього надходжень', 'Мої надходження', 'Витрачено з моїх',
  'Витрачено з ліміту', 'Відкладено на погашення (бюджетний місяць)', 'Погашено за бюджетний місяць (відкладено + невитрачене)', 'Мої на кінець місяця',
  'На картці на кінець місяця', 'Борг за лімітом на кінець місяця', 'Непомітні разом', 'Непомітні: з чого', 'З ліміту: причини', 'Топ-5 категорій', 'Записано'];
/** Колонка історії → підпис рядка на «Місяцях» (якщо відрізняється). */
const H_FROM = { 'Витрати': T_SPEND, 'Заощадження': T_SAVE, 'Податки': T_TAX, 'Борги': T_DEBT };

function buildHistory_(ss) {
  const sh = sheetOf_(ss, HISTORY_SHEET);
  sh.getRange(1, 1).setValue('📚 Історія місяців').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Підсумки місяців значеннями: пізні правки в журналі їх не змінюють. Записати чи перезаписати місяць: меню «📊 Monitoring Budgets».')
    .setFontColor('#5f6368');
  sh.getRange(3, 1, 1, H_COLS.length).setValues([H_COLS]).setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(3); sh.setFrozenColumns(1);
  const c = sheetOf_(ss, H_CAT_SHEET);
  c.getRange(1, 1).setValue('📚 Історія категорій').setFontWeight('bold').setFontSize(14);
  c.getRange(3, 1, 1, 6).setValues([['Місяць', 'Категорія', 'Тип', 'Група', 'План', 'Факт']]).setFontWeight('bold');
  c.setFrozenRows(3);
  return sh;
}

/** Значення «Місяців» за місяць ym ('yyyy-MM'): { sum: {підпис: число}, cats: [[назва, тип, група, план, факт]], reasons, holes }. */
function monthsRead_(ss, ym) {
  const sh = ss.getSheetByName(MONTHS_SHEET), y = Number(ym.slice(0, 4)), k = Number(ym.slice(5, 7));
  const was = sh.getRange(M_YEAR).getValue(), wasWho = String(sh.getRange(M_WHO).getValue() || '');
  const whoOff = wasWho !== '' && wasWho !== WHO_ALL;                                     // історія завжди за всю сім'ю
  if (whoOff) sh.getRange(M_WHO).setValue(WHO_ALL);
  if (Number(was) !== y || whoOff) { if (Number(was) !== y) sh.getRange(M_YEAR).setValue(y); SpreadsheetApp.flush(); }
  try {
    const c = mCols_(k), last = monthsSummaryRow_('Непомітні разом');
    const lab = sh.getRange(M_FIRST, 1, last - M_FIRST + 1, 3).getValues();
    const val = sh.getRange(M_FIRST, c.plan, last - M_FIRST + 1, 2).getValues();
    const num = v => (v === '' || v === null ? '' : Number(v) || 0);
    const out = { sum: {}, cats: [], reasons: [], holes: [] };
    const rows = monthsSummaryRows_(), first = M_CAT_LAST + 2;
    lab.forEach((l, i) => {
      const r = M_FIRST + i, name = String(l[0] || '').trim();
      if (!name) return;
      if (r <= M_CAT_LAST) { out.cats.push([name, String(l[1] || ''), String(l[2] || ''), num(val[i][0]), num(val[i][1])]); return; }
      const kind = (rows[r - first] || [])[1];
      if (kind === 'reason' || kind === 'noReason') out.reasons.push([kind === 'noReason' ? 'без причини' : name, num(val[i][1])]);
      else if (kind === 'hole') out.holes.push([name, num(val[i][1])]);
      else if (kind && kind !== 'head') out.sum[name] = num(val[i][1]);
    });
    return out;
  } finally {
    if (Number(was) !== y) sh.getRange(M_YEAR).setValue(was);
    if (whoOff) sh.getRange(M_WHO).setValue(wasWho);
    if (Number(was) !== y || whoOff) SpreadsheetApp.flush();
  }
}

function topText_(pairs, n) {
  return pairs.filter(x => Math.abs(Number(x[1]) || 0) > 0.005).sort((a, b) => b[1] - a[1]).slice(0, n || 5).map(x => x[0] + ' ' + money0_(x[1])).join('; ');
}

/** Записує (перезаписує) місяць ym у обидві історії. → { ym, cats } */
function historyWrite_(ss, ym) {
  guard_(ss);
  const m = monthsRead_(ss, ym), sh = ss.getSheetByName(HISTORY_SHEET), cs = ss.getSheetByName(H_CAT_SHEET);
  const sumOf = h => {
    const label = { 'Усі списання': 'Усі списання', 'Всього надходжень': 'Всього надходжень' }[h] || H_FROM[h] || h;
    return m.sum[label] === undefined ? '' : m.sum[label];
  };
  const row = H_COLS.map(h => {
    if (h === 'Місяць') return ym;
    if (h === 'Непомітні: з чого') return topText_(m.holes, 8);
    if (h === 'З ліміту: причини') return topText_(m.reasons, 8);
    if (h === 'Топ-5 категорій') return topText_(m.cats.filter(c => c[1] === T_SPEND).map(c => [c[0], c[4]]), 5);
    if (h === 'Записано') return new Date();
    return sumOf(h);
  });
  upsertRows_(sh, H_FIRST, H_COLS.length, [row], r => String(r[0]) === ym);
  const cats = m.cats.map(c => [ym, c[0], c[1], c[2], c[3], c[4]]);
  upsertRows_(cs, H_FIRST, 6, cats, r => String(r[0]) === ym);
  return { ym: ym, cats: cats.length };
}

/** Прибирає рядки, що підходять під same(), додає нові, сортує за першою колонкою (місяць) і пише значеннями. */
function upsertRows_(sh, first, n, add, same) {
  const last = sh.getLastRow(), have = last >= first ? sh.getRange(first, 1, last - first + 1, n).getValues() : [];
  ensureSize_(sh, first + have.length + add.length + 1, n);
  const keep = have.filter(r => String(r[0]) !== '' && !same(r)).concat(add).sort((a, b) => (String(a[0]) < String(b[0]) ? -1 : String(a[0]) > String(b[0]) ? 1 : 0));
  if (have.length) sh.getRange(first, 1, have.length, n).clearContent();
  if (!keep.length) return;
  sh.getRange(first, 1, keep.length, 1).setNumberFormat('@');                    // «2026-10» лишається текстом
  sh.getRange(first, 1, keep.length, n).setValues(keep);
}
