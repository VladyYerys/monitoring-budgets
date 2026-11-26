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
 */

const REG_SHEET = 'Регулярні';
const REG_HEAD = 4, REG_FIRST = 5, REG_LAST = 54;
const REG_TAG = '[Регулярний] ';
const REG_PER = ['щомісяця', 'щокварталу', 'щороку'];
const REG_CLOSE_DAYS = 5, REG_TOL = 0.25;
const REG_COLS = ['Назва', 'Категорія', 'Сума, ₴', 'День (1–31)', 'Як часто', 'З (перший платіж)', 'До (порожньо: без кінця)', 'Хто', 'Увімкнено', 'Стан (пише скрипт)'];

function buildRegular_(ss) {
  const sh = ensureSize_(sheetOf_(ss, REG_SHEET), REG_LAST + 3, REG_COLS.length);
  sh.getRange(1, 1).setValue('🔁 Регулярні платежі й надходження').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Кожен рядок стає планом «🔁 план регулярний» у журналі до кінця наступного місяця. Прийшов факт тієї ж категорії ' +
    '(±5 днів, ±25 % суми), наприклад з банку: план місяця закривається сам. Щокварталу й щороку: від місяця з колонки «З».').setFontColor('#5f6368').setWrap(true);
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
  sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 160); sh.setColumnWidth(10, 220);
  return sh;
}

/** Правила з аркуша: [{ i, name, cat, amt, day, per, from, to, who, on }]. */
function regList_(ss) {
  const sh = ss.getSheetByName(REG_SHEET), tz = ss.getSpreadsheetTimeZone();
  if (!sh) return [];
  return sh.getRange(REG_FIRST, 1, REG_LAST - REG_FIRST + 1, 9).getValues().map((r, i) => ({
    i: i, name: String(r[0]).trim(), cat: String(r[1]).trim(), amt: Number(r[2]) || 0, day: Math.max(1, Math.min(31, Math.round(Number(r[3]) || 1))),
    per: REG_PER.indexOf(String(r[4]).trim()) >= 0 ? String(r[4]).trim() : REG_PER[0], from: dateKey_(r[5], tz) || '', to: dateKey_(r[6], tz) || '',
    who: String(r[7] || '').trim(), on: String(r[8]).trim() !== 'ні',
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
      let hit = facts.find(o => !claimed[o.row] && o.d === key);
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
    col.push([!rule ? '' : !rule.on ? 'вимкнено' : (s.over ? '⚠️ прострочено ' + dm(s.over) + '; ' : '') +
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
