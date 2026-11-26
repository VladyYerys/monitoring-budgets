/** 29_Bank.gs: Спільний вхід для всіх банків: «Рахунки», операція → чернетка, захист від дублів, журнал «Банк». */

/*
 * Кожне джерело (Mono, ПриватБанк «Автоклієнт», виписка CSV чи Google Таблиця) готує операції в одному вигляді:
 *   { id, time (с), amount (коп., мінус: списання), description, comment, mcc, hold, balance?, commissionRate? }
 * і віддає їх у bankIngest_(). Далі все однаково: правила категорій, ваш рядок з тією ж сумою й датою ±1 день
 * не дублюється, та сама операція з іншого джерела (API і виписка) теж, банк скасував: чернетка зникає.
 * Номер операції має приставку джерела: Mono без приставки, «pb:» ПриватБанк, «csv:» виписка.
 * * «Рахунки»: назва · банк · чий · вид (основна картка / інша картка / рахунок ФОП) · ФОП · ключ (відбиток, без номерів) · останній імпорт.
 * Основна картка: з неї «Дні» й контрольні точки. Інші картки сім'ї: у «мої» й «Місяці». Рахунок ФОП: лише «ФОП».
 */

const CAT_FEE = 'Комісії й відсотки банку', HOLE_FEE = '🏦 Банківські комісії';
const ACC_HEAD = ['Назва (не перейменовуйте: так записано в журналі)', 'Банк', 'Чий', 'Вид', 'ФОП (для рахунку ФОП)', 'Ключ (службовий)', 'Останній імпорт', 'Примітка'];

/* ---------- «Рахунки» ---------- */

function buildAccounts_(ss) {
  const sh = ensureSize_(sheetOf_(ss, ACC_SHEET), ACC_LAST, 8), set = ss.getSheetByName(SETTINGS_SHEET);
  sh.getRange(1, 1).setValue('💳 Рахунки').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('З\'являються самі, коли підключаєте банк чи імпортуєте виписку. Ви можете змінити «Чий», «Вид» і «ФОП». ' +
    'Основна картка: з неї «Дні» й залишки; інші картки: у «мої» й «Місяці»; рахунок ФОП: лише на аркуші «ФОП».').setFontColor('#5f6368');
  sh.getRange(4, 1, 1, ACC_HEAD.length).setValues([ACC_HEAD]).setFontWeight('bold').setWrap(true);
  sh.getRange('D' + ACC_FIRST + ':D' + ACC_LAST).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ACC_KINDS).setAllowInvalid(false).build());
  if (set) sh.getRange('C' + ACC_FIRST + ':C' + ACC_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(set.getRange('B' + SET_MEMBERS_FIRST + ':B' + SET_SHARED_ROW)).setAllowInvalid(true).build());
  sh.getRange('G' + ACC_FIRST + ':G' + ACC_LAST).setNumberFormat('dd.mm.yyyy hh:mm');
  sh.getRange('F' + ACC_FIRST + ':F' + ACC_LAST).setFontColor('#9aa0a6');
  sh.setFrozenRows(4);
  return sh;
}

function accList_(ss) {
  const sh = ss.getSheetByName(ACC_SHEET);
  if (!sh) return [];
  return sh.getRange(ACC_FIRST, 1, ACC_LAST - ACC_FIRST + 1, 8).getValues().map((r, i) => ({ row: ACC_FIRST + i, name: String(r[0]).trim(), bank: String(r[1]),
    who: String(r[2] || ''), kind: String(r[3] || '') || K_MAIN, fop: String(r[4] || ''), key: String(r[5] || ''), last: r[6] })).filter(a => a.name);
}

/**
 * Рахунок за ключем; немає: новий рядок. Перший рахунок-картка стає основною, рахунок ФОП: «рахунок ФОП», решта: «інша картка».
 * Ваші зміни «Чий», «Вид», «ФОП» не перезаписуються. → { name, kind, who, fop }
 */
function accEnsure_(ss, o) {
  const sh = ss.getSheetByName(ACC_SHEET) || buildAccounts_(ss), list = accList_(ss);
  const have = list.find(a => a.key === o.key);
  if (have) return have;
  const all = sh.getRange(ACC_FIRST, 1, ACC_LAST - ACC_FIRST + 1, 1).getValues(), free = all.findIndex(r => !String(r[0]).trim());
  if (free < 0) throw new Error('На аркуші «' + ACC_SHEET + '» немає вільних рядків.');
  let name = o.name, n = 2;
  while (list.some(a => a.name === name)) name = o.name + ' (' + (n++) + ')';
  const kind = o.kind || (list.some(a => a.kind === K_MAIN) ? K_CARD : K_MAIN);
  sh.getRange(ACC_FIRST + free, 1, 1, 6).setValues([[name, o.bank, o.who || '', kind, o.fop || '', o.key]]);
  return { row: ACC_FIRST + free, name: name, bank: o.bank, who: o.who || '', kind: kind, fop: o.fop || '', key: o.key };
}

function accTouch_(ss, acc) {
  const sh = ss.getSheetByName(ACC_SHEET);
  if (sh && acc && acc.row) sh.getRange(acc.row, 7).setValue(new Date());
}

/** У сімейному режимі: чиї рахунки (вікно); інакше порожньо. null: скасовано. */
function askWho_(title, what) {
  const ss = SpreadsheetApp.getActive();
  if (!familyMode_(ss)) return '';
  const ms = members_(ss), ui = SpreadsheetApp.getUi();
  if (!ms.length) return '';
  const a = ui.prompt(title, 'Чиї ' + what + '? ' + ms.map((m, i) => (i + 1) + ') ' + m).join(', ') + ', ' + (ms.length + 1) + ') ' + WHO_SHARED +
    '.\nВведіть номер:', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return null;
  const k = Number(String(a.getResponseText()).trim());
  return k >= 1 && k <= ms.length ? ms[k - 1] : WHO_SHARED;
}

/* ---------- спільний вхід ---------- */

function srcOf_(id) { id = String(id || ''); return id.indexOf('csv:') === 0 ? 'csv' : id.indexOf('pb:') === 0 ? 'pb' : 'mono'; }

/**
 * Операції одного рахунку → журнал. src: { acc: {name, kind, who}, items, dateOf, timeOf, balShift?, coverFrom?, coverTo? }
 * → { plan, written, lost }
 */
function bankIngest_(ss, src) {
  const sh = ss.getSheetByName(JOURNAL), bankSh = ss.getSheetByName(BANK_SHEET), tz = ss.getSpreadsheetTimeZone();
  const journal = readJournal_(ss);
  const lastLog = bankSh.getLastRow();
  const logged = new Set(lastLog >= BANK_LOG_FIRST ? bankSh.getRange(BANK_LOG_FIRST, 2, lastLog - BANK_LOG_FIRST + 1, 1).getValues().map(r => String(r[0])) : []);
  const plan = planImport_({ items: src.items, journal: journal, logged: logged, rules: readRules_(ss), dateOf: src.dateOf, timeOf: src.timeOf,
    balShift: src.balShift, coverFrom: src.coverFrom, coverTo: src.coverTo, acc: src.acc, src: src.src || 'mono' });
  (src.extraLog || []).forEach(l => { if (!logged.has(l[0])) plan.log.push(l); });
  const w = writeDrafts_(sh, journal, plan, tz, src.acc);
  plan.vanish.forEach(row => {
    sh.getRange(row, 1, 1, J_INPUT_N).setValues([['', '', '', '', '', '', '', '']]);
    sh.getRange(row, colNum_(JC.bankId), 1, J_MOVE_N).setValues([emptyBank_().concat(['', ''])]);
  });
  const stamp = new Date(), logRows = plan.log.map(l => [stamp].concat(l).concat([src.acc.name]));
  if (logRows.length) {
    const at = Math.max(bankSh.getLastRow() + 1, BANK_LOG_FIRST);
    ensureSize_(bankSh, at + logRows.length, BANK_LOG_N).getRange(at, 1, logRows.length, BANK_LOG_N).setValues(logRows);
  }
  accTouch_(ss, src.acc);
  return { plan: plan, written: w.written, lost: w.lost };
}

function isPendingReversal_(it) {
  return !!it.hold && Number(it.amount) > 0 && /скасуван|cancel|reversal/i.test(String(it.description || ''));
}

/** Чиста функція: категорія операції: комісія, правило; на рахунку ФОП: надходження → «Дохід ФОП», переказ → між своїми; інакше переказ → червоний. */
function categorize_(it, amt, rules, acc) {
  if (it._fee) return { cat: CAT_FEE, hole: HOLE_FEE, rule: null };
  const rule = matchRule_(rules, { desc: it.description, comment: it.comment, mcc: it.mcc }, amt, false);
  if (rule) return { cat: rule.category, hole: rule.hole, rule: rule };
  if (acc && acc.kind === K_FOP) return { cat: isTransfer_(it) && amt < 0 ? CAT_OWN_MOVE : amt > 0 ? CAT_FOP_INCOME : '', hole: '', rule: null };
  return { cat: isTransfer_(it) ? CAT_TRANSFER : '', hole: '', rule: null };
}

function noteFor_(it, amt) {
  return [String(it.description || ''), it.comment ? '«' + it.comment + '»' : '',
    it._feeSplit ? 'комісія ' + round2_(it._feeSplit / 100) + ' ₴, окремим рядком' : '',
    it.hold ? '⏳ в обробці' : '', amt > 0 ? 'надходження' : ''].filter(Boolean).join(' · ');
}

/** Чиста функція: ваш рядок журналу з тією ж сумою, напрямом і датою ±1 день (без номера банку, не план і не чернетка). */
function findManual_(journal, claimed, target, amt, key) {
  let best = null;
  journal.forEach(r => {
    if (r.blank || claimed[r.row] || (r.bank && r.bank[0])) return;
    if (r.e === STATUS_DRAFT || Number(r.m) === 1 || isPlan_(r.e) || r.key === '') return;
    if (Math.abs(Number(r.b) - target) >= 0.005) return;
    const l = Number(r.l);
    if (r.l !== '' && r.l !== null && r.l !== undefined && l !== 0 && (l < 0) !== (amt < 0)) return;
    const d = Math.abs(dayDiff_(r.key, key));
    if (d > 1) return;
    if (!best || d < best.d) best = { r: r, d: d };
  });
  return best;
}

/** Чиста функція: та сама операція, що вже прийшла з іншого джерела (API і виписка): та сама сума зі знаком, дата ±1 день, той самий рахунок. */
function findOtherSource_(journal, claimed, id, amt, key, accName) {
  const src = srcOf_(id);
  return journal.find(r => r.bank && r.bank[0] && !claimed[r.row] && srcOf_(r.bank[0]) !== src && r.key &&
    Math.abs(Number(r.bank[2]) - amt) < 0.005 && Math.abs(dayDiff_(r.key, key)) <= 1 && (r.acct === accName || (!r.acct && !accName))) || null;
}

/**
 * Чиста функція: що зробити з операціями рахунку.
 * ctx: {items, journal (readJournal_), logged:Set, rules, dateOf, timeOf, balShift, coverFrom, coverTo, acc}
 * → {drafts, marks, updates, vanish, log, stats}
 */
function planImport_(ctx) {
  const acc = ctx.acc || { name: '', kind: K_MAIN, who: '' };
  const out = { drafts: [], marks: [], updates: [], log: [], vanish: [], stats: { transfers: 0, uncategorized: 0, changed: 0, resized: 0, redated: 0, other: 0 } };
  const mine = r => r.acct === acc.name || (!r.acct && acc.kind === K_MAIN);
  if (ctx.coverFrom && ctx.coverTo) {                     // чернетки цього рахунку, яких у виписці за повністю охоплені дні вже немає: банк скасував
    const ids = {};
    ctx.items.forEach(it => { ids[String(it.id)] = 1; });
    ctx.journal.forEach(r => {
      const id = r.bank && r.bank[0] ? String(r.bank[0]) : '';
      if (!id || r.e !== STATUS_DRAFT || ids[id] || !mine(r) || srcOf_(id) !== (ctx.src || 'mono') ||
        r.key < ctx.coverFrom || r.key > ctx.coverTo) return;
      out.vanish.push(r.row);
      out.log.push([id, r.key, '', String(r.bank[1] || ''), '', '', '', '', 'банк скасував операцію: чернетку прибрано', r.row]);
    });
  }
  const byId = {};
  ctx.journal.forEach(r => { if (r.bank && r.bank[0]) byId[String(r.bank[0])] = r; });
  const claimed = {}, feeInRow = {};
  ctx.items.slice().sort((a, b) => a.time - b.time).forEach(it => {
    const id = String(it.id), amt = round2_(it.amount / 100), abs = Math.abs(amt), key = ctx.dateOf(it);
    const own = typeof it._bal === 'number' && typeof ctx.balShift === 'number' ? round2_((it._bal - ctx.balShift) / 100) : '';
    const bank = [id, String(it.description || ''), amt, Number(it.mcc || 0), ctx.timeOf(it), own];
    const base = [id, key, amt, String(it.description || ''), Number(it.mcc || 0), String(it.comment || ''),
      round2_((it._feeSplit || it.commissionRate || 0) / 100), it.hold ? 'так' : ''];
    const row = byId[id];
    if (row) {                                            // уже в журналі: без дубля; чернетку уточнюємо
      const draft = row.e === STATUS_DRAFT, sumDiff = Math.abs(Number(row.b) - abs) >= 0.005, keyDiff = draft && row.key !== key;
      const old = row.bank || [];
      const balDiff = own !== '' && (old[5] === '' || old[5] === null || old[5] === undefined || Math.abs(Number(old[5]) - own) >= 0.005);
      if (sumDiff && !draft) { out.stats.changed++; out.log.push(base.concat(['⚠️ сума в банку ' + abs + ', у журналі ' + row.b + ', перевірте', row.row])); }
      if ((sumDiff && draft) || keyDiff || balDiff) {
        out.updates.push({ row: row.row, b: sumDiff && draft ? abs : null, key: keyDiff ? key : null, bank: bank });
        if (sumDiff && draft) { out.stats.resized++; out.log.push(base.concat(['сума уточнилась: ' + row.b + ' → ' + abs, row.row])); }
        if (keyDiff) { out.stats.redated++; out.log.push(base.concat(['дату уточнено: ' + row.key + ' → ' + key, row.row])); }
      }
      return;
    }
    if (ctx.logged.has(id)) return;                       // уже оброблена раніше (могли видалити чернетку)
    if (it._fee && feeInRow[it._parent]) { out.log.push(base.concat(['комісія вже у вашому рядку', feeInRow[it._parent]])); return; }
    const other = findOtherSource_(ctx.journal, claimed, id, amt, key, acc.name);
    if (other) { claimed[other.row] = 1; out.stats.other++; out.log.push(base.concat(['✓ уже є (прийшла з іншого джерела: ' + other.bank[0] + ')', other.row])); return; }
    const findRow = target => findManual_(ctx.journal, claimed, target, amt, key);
    let best = findRow(abs), withFee = false;
    if (!best && it._feeSplit) { best = findRow(round2_(abs + it._feeSplit / 100)); withFee = !!best; }
    if (best) {
      claimed[best.r.row] = 1;
      if (withFee) feeInRow[id] = best.r.row;
      out.marks.push({ row: best.r.row, bank: bank });
      out.log.push(base.concat(['✓ уже є у журналі (ваш рядок' + (withFee ? ', разом із комісією' : '') + ')', best.r.row]));
      return;
    }
    const cz = categorize_(it, amt, ctx.rules, acc), rule = cz.rule, cat = cz.cat;
    if (cat === CAT_TRANSFER) out.stats.transfers++;
    if (!cat) out.stats.uncategorized++;
    const log = base.concat(['чернетка' + (cat ? ' → ' + cat : ' (без категорії)') + (rule ? ' (правило: ' + rule.text.slice(0, 40) + ')' : ''), '']);
    out.drafts.push({ key: key, b: abs, c: cat, d: noteFor_(it, amt), f: cz.hole, g: acc.who || '', bank: bank, log: log, ts: Number(it.time) * 1000 || '' });
    out.log.push(log);
  });
  return out;
}

/** Пише план: чернетки у вільні рядки після останнього заповненого; номер банку й рахунок вашим рядкам; уточнення. */
function writeDrafts_(sh, journal, plan, tz, acc) {
  const name = acc && acc.name || '';
  let lastUsed = -1;
  journal.forEach((r, i) => { if (!r.blank || r.bank.some(x => x !== '' && x !== null) || (r.ord !== '' && r.ord !== null) || r.acct) lastUsed = i; });
  const room = (J_LAST - J_FIRST + 1) - (lastUsed + 1), fit = plan.drafts.slice(0, Math.max(0, room)), y = colNum_(JC.bankId), af = colNum_(JC.acct);
  if (fit.length) {
    const r0 = J_FIRST + lastUsed + 1;
    sh.getRange(r0, 1, fit.length, J_INPUT_N).setValues(fit.map(d => [Utilities.parseDate(d.key, tz, 'yyyy-MM-dd'), d.b, d.c, d.d, STATUS_DRAFT, d.f, d.g || '', '']));
    sh.getRange(r0, y, fit.length, J_MOVE_N).setValues(fit.map(d => d.bank.concat([d.ts || Date.now(), name])));
    fit.forEach((d, i) => { d.log[9] = r0 + i; });
  }
  plan.drafts.slice(fit.length).forEach(d => { d.log[0] = 'немає місця:' + d.log[0]; d.log[8] = '⚠️ не вмістилось у журнал (немає вільних рядків)'; });
  plan.marks.forEach(mk => { sh.getRange(mk.row, y, 1, BANK_N).setValues([mk.bank]); if (name) sh.getRange(mk.row, af).setValue(name); });
  plan.updates.forEach(u => {
    if (u.b !== null) sh.getRange(u.row, 2).setValue(u.b);
    if (u.key) sh.getRange(u.row, 1).setValue(Utilities.parseDate(u.key, tz, 'yyyy-MM-dd'));
    sh.getRange(u.row, y, 1, BANK_N).setValues([u.bank]);
  });
  return { written: fit.length, lost: plan.drafts.length - fit.length };
}

/** Підтвердити всі чернетки, де вже є категорія (крім червоних переказів). */
function confirmDrafts() {
  if (!menuOk_('✅ Чернетки')) return;
  const ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActive().getSheetByName(JOURNAL);
  const ce = sh.getRange(J_FIRST, 3, J_LAST - J_FIRST + 1, 3).getValues();
  const rows = [];
  ce.forEach((v, i) => { if (v[2] === STATUS_DRAFT && v[0] !== '' && v[0] !== CAT_TRANSFER) rows.push(J_FIRST + i); });
  const red = ce.filter(v => v[0] === CAT_TRANSFER).length;
  if (!rows.length) { show_('✅ Чернетки', ['Немає чернеток з категорією.' + (red ? ' Червоних переказів: ' + red + '.' : '')]); return; }
  const ok = ui.alert('✅ Підтвердити чернетки', 'Поставити «✅ факт» у ' + rows.length + ' чернетках з категорією?' +
    (red ? '\nЧервоні перекази (' + red + ') не чіпаються.' : ''), ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  rows.forEach(r => sh.getRange(r, 5).setValue(STATUS_FACT));
}

/* ---------- аркуш «Банк» ---------- */

function buildBankLog_(ss) {
  const sh = sheetOf_(ss, BANK_SHEET);
  sh.getRange(1, 1).setValue('🏦 Банк: що прийшло з банку').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1, 2, 6).setValues([['Останній імпорт', '', 'Результат', '', 'Рахунки', ''], ['Telegram', '', '', '', '', '']]);
  sh.getRange('B2').setNumberFormat('dd.mm.yyyy hh:mm');
  sh.getRange(4, 1, 1, BANK_LOG_N).setValues([['Коли', 'Номер операції', 'Дата', 'Сума', 'Опис', 'MCC', 'Коментар', 'Комісія', 'В обробці', 'Що зроблено',
    'Рядок журналу', 'Рахунок']]).setFontWeight('bold');
  sh.getRange('A' + BANK_LOG_FIRST + ':A' + sh.getMaxRows()).setNumberFormat('dd.mm.yyyy hh:mm');
  sh.getRange('B' + BANK_LOG_FIRST + ':B' + sh.getMaxRows()).setNumberFormat('@');
  sh.setFrozenRows(4);
  return sh;
}
