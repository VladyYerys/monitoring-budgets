/** 29_Bank.gs: Спільний вхід для всіх банків: «Рахунки», операція → чернетка, захист від дублів, журнал «Банк». */

/*
 * Кожне джерело (Mono, ПриватБанк «Автоклієнт», виписка CSV чи Google Таблиця) готує операції в одному вигляді:
 *   { id, time (с), amount (коп., мінус: списання), description, comment, mcc, hold, balance?, commissionRate?,
 *     counterName?, counterIban?, counterEdrpou? (отримувач чи відправник, як дає Mono) }
 * і віддає їх у bankIngest_(). Далі все однаково: правила категорій, ваш рядок з тією ж сумою й датою ±1 день
 * не дублюється, та сама операція з іншого джерела (API і виписка) теж, банк скасував: чернетка зникає.
 * Без правила (переказ чи порожня категорія), по черзі (версія 0.4.9):
 *   1) між своїми: така ж сума з протилежним знаком на іншому вашому рахунку ±1 день → обидві «між своїми»;
 *   2) план: план з «Регулярних» чи «Боргів» ±5 днів, ±10 % суми, той самий напрям → рядок плану стає цією чернеткою
 *      (категорія плану, коментар плану + опис банку), тож план закрито;
 *   3) історія: цьому опису ви ≥ 3 рази ставили ту саму категорію (≥ 80 %) → вона; інакше підказка «зазвичай: …».
 * Категорію визначено (правило, комісія, ФОП, між своїми, план, історія) → одразу «✅ факт» («Налаштування», B35: так).
 * Чернетками лишаються: червоний переказ, порожня категорія, «⏳ в обробці» (сума ще може змінитися, банк може скасувати);
 * коли «в обробці» завершилось і категорія є, наступний імпорт сам ставить «✅ факт».
 * Повернення за покупку (версія 0.5.3): надходження з описом «Скасування. МАГАЗИН» (так пише Mono; також «Повернення», refund, reversal)
 * або без слова, але той самий магазин і MCC, що в покупці. Шукаємо покупку цього рахунку в журналі: той самий магазин, сума не
 * менша за повернення, до 60 днів тому. Рядок: категорія покупки, сума з мінусом («Одяг −1 200»), коментар «↩ повернення за покупку 12.11».
 * Немає покупки: категорія з правила для магазину; немає й правила: чернетка без категорії. Без слова: завжди чернетка (перевірте).
 * «Скасування … в обробці» за вже списаною покупкою: чернетка «⏳ в обробці» (сама стане фактом); за ще не списаною чи невідомою: пропуск, як і раніше.
 * Покупка частинами й розстрочка (версія 0.5.4, модуль «🤝 Борги й розстрочки»): «Платіж МАГАЗИН - monomarket» (перший платіж),
 * «Погашення МАГАЗИН - monomarket», «Повне погашення МАГАЗИН - monomarket» чи «Повне погашення» (розстрочка на картку).
 * Рядок «Боргів» шукаємо за «Назвою в банку» чи назвою; план цього боргу ±10 днів закривається незалежно від категорії й суми;
 * без плану: факт з категорією боргу й коментарем «[Борг] Назва: мм.рррр» (іде в «Сплачено»). Рядка немає: перший платіж
 * (чи перше погашення) додає його на «Борги» (сума боргу порожня: її впише людина). «Повне погашення»: «Закрито» = дата.
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

/** «Довідник»: категорія → тип (для напряму плану). */
function catTypes_(ss) {
  const d = ss.getSheetByName(DICT_SHEET), out = {};
  if (!d) return out;
  d.getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 2).getValues().forEach(r => { if (String(r[0]).trim()) out[String(r[0]).trim()] = String(r[1] || ''); });
  return out;
}

function srcOf_(id) { id = String(id || ''); return id.indexOf('csv:') === 0 ? 'csv' : id.indexOf('pb:') === 0 ? 'pb' : 'mono'; }

/**
 * Операції одного рахунку → журнал. src: { acc: {name, kind, who}, items, pending? («Скасування … в обробці»), dateOf, timeOf, balShift?, coverFrom?, coverTo? }
 * → { plan, written, lost }
 */
function bankIngest_(ss, src) {
  const sh = ss.getSheetByName(JOURNAL), bankSh = ss.getSheetByName(BANK_SHEET), tz = ss.getSpreadsheetTimeZone();
  const journal = readJournal_(ss), debtsOn = modOn_('debts') && !!ss.getSheetByName(DEBTS_SHEET);
  const lastLog = bankSh.getLastRow();
  const logged = new Set(lastLog >= BANK_LOG_FIRST ? bankSh.getRange(BANK_LOG_FIRST, 2, lastLog - BANK_LOG_FIRST + 1, 1).getValues().map(r => String(r[0])) : []);
  const plan = planImport_({ items: src.items, pending: src.pending, journal: journal, logged: logged, rules: readRules_(ss), dateOf: src.dateOf, timeOf: src.timeOf,
    balShift: src.balShift, coverFrom: src.coverFrom, coverTo: src.coverTo, acc: src.acc, src: src.src || 'mono', types: catTypes_(ss), autoFact: autoFact_(ss),
    debts: debtsOn ? debtsAll_(ss) : null });
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
  const debts = debtsOn ? debtsFromBank_(ss, plan) : { added: [], closed: [] };
  return { plan: plan, written: w.written, lost: w.lost, taken: w.taken, debts: debts };
}

function isPendingReversal_(it) {
  return !!it.hold && Number(it.amount) > 0 && /скасуван|cancel|reversal/i.test(String(it.description || ''));
}

/** Чиста функція: категорія операції: комісія, правило; на рахунку ФОП: надходження → «Дохід ФОП», переказ → між своїми; інакше переказ → червоний. */
function categorize_(it, amt, rules, acc) {
  if (it._fee) return { cat: CAT_FEE, hole: HOLE_FEE, rule: null };
  const rule = matchRule_(rules, { desc: it.description, comment: it.comment, mcc: it.mcc, cname: it.counterName, ciban: it.counterIban, cedrpou: it.counterEdrpou }, amt, false);
  if (rule) return { cat: rule.category, hole: rule.hole, rule: rule };
  if (acc && acc.kind === K_FOP) return { cat: isTransfer_(it) && amt < 0 ? CAT_OWN_MOVE : amt > 0 ? CAT_FOP_INCOME : '', hole: '', rule: null };
  return { cat: isTransfer_(it) ? CAT_TRANSFER : '', hole: '', rule: null };
}

function noteFor_(it, amt) {
  const who = String(it.counterName || '').trim(), desc = String(it.description || '');
  return [desc, who && norm_(desc).indexOf(norm_(who)) < 0 && norm_(who).indexOf(norm_(desc)) < 0 ? who : '', it.comment ? '«' + it.comment + '»' : '',
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

/* ---------- покупка частинами й розстрочка ---------- */

const PARTS_DAYS = 10;
/** Чиста функція: платіж частинами з опису Mono. → null | { kind: 'first' | 'pay' | 'final', shop, mono } */
function partsOf_(desc) {
  const s = String(desc || '').replace(/\s+/g, ' ').trim(), mono = /\s-\s*monomarket$/i.test(s), body = s.replace(/\s-\s*monomarket$/i, '').trim();
  let m = body.match(/^повне погашення(?:\s+(.+))?$/i);
  if (m) return { kind: 'final', shop: (m[1] || '').trim(), mono: mono };
  m = body.match(/^погашення(?:\s+(.+))?$/i);
  if (m) return { kind: 'pay', shop: (m[1] || '').trim(), mono: mono };
  m = body.match(/^платіж\s+(.+)$/i);
  if (m && mono) return { kind: 'first', shop: m[1].trim(), mono: true };
  return null;
}

/** Чиста функція: рядок «Боргів» для платежу: за «Назвою в банку» чи назвою; без магазину: єдина відкрита розстрочка з таким платежем. */
function partsDebt_(debts, pt, abs) {
  const list = (debts || []).filter(d => !d.closed && d.dir === DEBT_OWE);
  if (pt.shop) {
    const s = refundShop_(pt.shop), same = x => { const t = refundShop_(x); return !!t && (t === s || (t.length >= 3 && s.length >= 3 && (t.indexOf(s) === 0 || s.indexOf(t) === 0))); };
    return list.find(d => d.bank && same(d.bank)) || list.find(d => same(d.name)) || null;
  }
  const c = list.filter(d => d.kind === DEBT_INST && d.monthly > 0 && Math.abs(d.monthly - abs) <= Math.max(1, abs * 0.05));
  return c.length === 1 ? c[0] : null;
}

/** Чиста функція: план боргу в журналі. Відомий борг: його план ±10 днів (будь-яка сума); невідомий: план «[Борг] …» ±5 днів і ±1 % суми. */
function partsPlan_(journal, claimed, abs, key, debt) {
  let best = null, bestS = 0;
  journal.forEach(r => {
    if (claimed[r.row] || !isPlan_(r.e) || !r.key || String(r.d).indexOf(DEBT_TAG) !== 0) return;
    const dd = Math.abs(dayDiff_(r.key, key)), diff = Math.abs(Number(r.b) - abs);
    if (debt) { if (String(r.d).indexOf(DEBT_TAG + debt.name + ':') !== 0 || dd > PARTS_DAYS) return; }
    else if (dd > PLAN_DAYS || diff > Math.max(1, abs * 0.01)) return;
    const sc = dd * 1e6 + diff;
    if (!best || sc < bestS) { best = r; bestS = sc; }
  });
  return best;
}

/* ---------- повернення за покупку ---------- */

const REFUND_RE = /^\s*(?:скасування|повернення|refund|reversal|cancel(?:lation)?)(?=[\s.:,-]|$)[\s.:,-]*/i;
const REFUND_DAYS = 60, REFUND_NOTE = '↩ повернення';

/** Чиста функція: назва магазину для порівняння (без регістру, лапок і розділових знаків). */
function refundShop_(s) {
  return String(s || '').toLowerCase().replace(/[«»"'ʼ’`]/g, '').replace(/[^0-9a-zа-яіїєґ]+/gi, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * Чиста функція: чи це повернення за покупку і за яку. → null | { word, shop (текст після «Скасування.»), buy: рядок покупки | null }
 * Покупка: рядок з банку цього рахунку (mine), списання, той самий магазин, не дешевша за повернення, від key−60 днів до key,
 * з категорією (не переказ). Без слова «Скасування»: лише не переказ (MCC, опис), покупку знайдено й MCC той самий. Спершу ще не «зайняті» іншим
 * поверненням, точна сума, найновіша.
 */
function refundOf_(it, journal, key, claimed, mine) {
  const amt = round2_(Number(it.amount) / 100);
  if (!(amt > 0) || it._fee) return null;
  const desc = String(it.description || ''), m = desc.match(REFUND_RE), word = !!m, raw = word ? desc.slice(m[0].length).trim() : desc;
  if (!word && isTransfer_(it)) return null;                       // переказ від людини з тією ж сумою: не повернення
  const shop = refundShop_(raw), from = addDays_(key, -REFUND_DAYS);
  let best = null, bestScore = null;
  if (shop.length >= 2) journal.forEach(r => {
    if (!r.bank || !r.bank[0] || !r.key || r.key > key || r.key < from || !(Number(r.bank[2]) < 0) || isPlan_(r.e)) return;
    if (looseCat_(r.c) || r.c === CAT_OWN_MOVE || (mine && !mine(r))) return;
    const s = refundShop_(r.bank[1]);
    if (!(s === shop || (s.length >= 3 && shop.length >= 3 && (s.indexOf(shop) === 0 || shop.indexOf(s) === 0)))) return;
    const paid = -Number(r.bank[2]);
    if (amt > paid + 0.005 || (!word && Number(r.bank[3]) !== Number(it.mcc || 0))) return;
    const score = [claimed && claimed[r.row] ? 1 : 0, Math.abs(paid - amt) < 0.005 ? 0 : 1, -dayDiff_(r.key, from)];
    if (!bestScore || score[0] < bestScore[0] || (score[0] === bestScore[0] && (score[1] < bestScore[1] || (score[1] === bestScore[1] && score[2] < bestScore[2])))) {
      best = r; bestScore = score;
    }
  });
  if (!word && !best) return null;
  return { word: word, shop: raw, buy: best };
}

/* ---------- без правила: між своїми, план, історія ---------- */

const PLAN_DAYS = 5, PLAN_TOL = 0.10, HIST_MIN = 3, HIST_SHARE = 0.8;
const looseCat_ = c => !c || c === CAT_TRANSFER;

/** Чиста функція: рядок іншого вашого рахунку з тією ж сумою протилежного знаку, ±1 день, ще без категорії чи вже «між своїми». */
function findOwnPair_(journal, claimed, amt, key, accName) {
  let best = null;
  journal.forEach(r => {
    if (claimed[r.row] || !r.bank || !r.bank[0] || !r.key || (r.acct || '') === (accName || '')) return;
    if (Math.abs(Number(r.bank[2]) + amt) >= 0.005) return;
    if (!(looseCat_(r.c) || r.c === CAT_OWN_MOVE) || isPlan_(r.e)) return;
    const d = Math.abs(dayDiff_(r.key, key));
    if (d > 1) return;
    if (!best || d < best.d) best = { r: r, d: d };
  });
  return best && best.r;
}

/** Чиста функція: план (регулярний чи борг) ±5 днів, ±10 % суми, той самий напрям; cat: лише план цієї категорії (порожньо: будь-який). */
function findPlanFor_(journal, claimed, amt, key, cat, types) {
  let best = null;
  journal.forEach(r => {
    if (claimed[r.row] || !isPlan_(r.e) || !r.key || !r.c || (r.bank && r.bank[0])) return;
    if (r.d.indexOf(REG_TAG) !== 0 && r.d.indexOf(DEBT_TAG) !== 0) return;
    if (cat && r.c !== cat) return;
    const inc = (types || {})[r.c] === T_IN;
    if (inc !== (amt > 0)) return;
    const plan = Number(r.b) || 0, diff = Math.abs(plan - Math.abs(amt)), days = Math.abs(dayDiff_(r.key, key));
    if (plan <= 0 || diff > plan * PLAN_TOL + 0.005 || days > PLAN_DAYS) return;
    const score = diff / plan + days / 100;
    if (!best || score < best.score) best = { r: r, score: score };
  });
  return best && best.r;
}

/** Ключ історії: напрям і опис банку без цифр. */
const histKey_ = (amt, desc) => (amt < 0 ? '-' : '+') + norm_(desc);

/** Чиста функція: що ви ставили операціям з таким описом. → { ключ: { n, cats: { категорія: скільки } } } */
function histMap_(journal) {
  const m = {};
  journal.forEach(r => {
    if (!r.bank || !r.bank[0] || r.e !== STATUS_FACT || looseCat_(r.c) || r.c === CAT_FEE) return;
    const k = histKey_(Number(r.bank[2]) || 0, r.bank[1]);
    if (k.length < 4) return;
    const x = m[k] || (m[k] = { n: 0, cats: {} });
    x.n++; x.cats[r.c] = (x.cats[r.c] || 0) + 1;
  });
  return m;
}

/** Чиста функція: { cat (впевнено: ≥ 3 і ≥ 80 %), hint: «зазвичай: Оренда (5 з 6), Борги (1)» } або null. */
function histGuess_(hist, amt, desc) {
  const x = hist[histKey_(amt, desc)];
  if (!x) return null;
  const top = Object.keys(x.cats).sort((a, b) => x.cats[b] - x.cats[a]), k0 = x.cats[top[0]];
  const sure = x.n >= HIST_MIN && k0 / x.n >= HIST_SHARE;
  return { cat: sure ? top[0] : '', text: top[0] + ' (' + k0 + ' з ' + x.n + ')' + (top[1] ? ', ' + top[1] + ' (' + x.cats[top[1]] + ')' : '') };
}

/**
 * Чиста функція: що зробити з операціями рахунку.
 * ctx: {items, journal (readJournal_), logged:Set, rules, dateOf, timeOf, balShift, coverFrom, coverTo, acc, types: {категорія: тип}, autoFact}
 * → {drafts, marks, takes (план стає чернеткою), recats (інший рахунок → «між своїми»), updates, vanish, log, stats}
 */
function planImport_(ctx) {
  const acc = ctx.acc || { name: '', kind: K_MAIN, who: '' };
  const out = { drafts: [], marks: [], takes: [], recats: [], updates: [], log: [], vanish: [],
    newDebts: [], closeDebts: [],
    stats: { transfers: 0, uncategorized: 0, changed: 0, resized: 0, redated: 0, other: 0, own: 0, plans: 0, learned: 0, facts: 0, refunds: 0, skipped: 0, parts: 0 } };
  const HOLD_NOTE = ' · ⏳ в обробці', statusOf = (cat, hold) => ctx.autoFact && !hold && !looseCat_(cat) ? STATUS_FACT : STATUS_DRAFT;
  const hist = histMap_(ctx.journal), pairClaimed = {}, planClaimed = {};
  const mine = r => r.acct === acc.name || (!r.acct && acc.kind === K_MAIN);
  if (ctx.coverFrom && ctx.coverTo) {                     // чернетки цього рахунку, яких у виписці за повністю охоплені дні вже немає: банк скасував
    const ids = {};
    ctx.items.concat(ctx.pending || []).forEach(it => { ids[String(it.id)] = 1; });
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
  const claimed = {}, feeInRow = {}, refundClaimed = {}, pendTake = [];
  (ctx.pending || []).forEach(it => {                     // «Скасування … в обробці»: повернення за вже списану покупку беремо, решту пропускаємо
    const id = String(it.id), key = ctx.dateOf(it), rf = byId[id] ? null : refundOf_(it, ctx.journal, key, null, mine);
    if (byId[id] || (rf && rf.buy && String(rf.buy.d || '').indexOf('⏳') < 0)) { pendTake.push(it); return; }
    out.stats.skipped++;
    if (!ctx.logged.has('skip:' + id)) out.log.push(['skip:' + id, key, round2_(it.amount / 100), String(it.description || ''), Number(it.mcc || 0), '', 0, 'так',
      'пропущено: скасування ще не списаної оплати (у залишку банку його немає)', '']);
  });
  ctx.items.concat(pendTake).sort((a, b) => a.time - b.time).forEach(it => {
    const id = String(it.id), amt = round2_(it.amount / 100), abs = Math.abs(amt), key = ctx.dateOf(it);
    const own = typeof it._bal === 'number' && typeof ctx.balShift === 'number' ? round2_((it._bal - ctx.balShift) / 100) : '';
    const bank = [id, String(it.description || ''), amt, Number(it.mcc || 0), ctx.timeOf(it), own];
    const base = [id, key, amt, String(it.description || ''), Number(it.mcc || 0), String(it.comment || ''),
      round2_((it._feeSplit || it.commissionRate || 0) / 100), it.hold ? 'так' : ''];
    const row = byId[id];
    if (row) {                                            // уже в журналі: без дубля; чернетку уточнюємо
      const want = Number(row.b) < 0 ? -abs : abs;        // повернення записане мінусом
      const draft = row.e === STATUS_DRAFT, sumDiff = Math.abs(Number(row.b) - want) >= 0.005, keyDiff = draft && row.key !== key;
      const old = row.bank || [];
      const balDiff = own !== '' && (old[5] === '' || old[5] === null || old[5] === undefined || Math.abs(Number(old[5]) - own) >= 0.005);
      if (sumDiff && !draft) { out.stats.changed++; out.log.push(base.concat(['⚠️ сума в банку ' + abs + ', у журналі ' + row.b + ', перевірте', row.row])); }
      const promote = draft && ctx.autoFact && !it.hold && !looseCat_(row.c) && row.d.indexOf('⏳ в обробці') >= 0;   // «в обробці» завершилось
      if (promote) { out.stats.facts++; out.log.push(base.concat(['«в обробці» завершилось: ✅ факт', row.row])); }
      if ((sumDiff && draft) || keyDiff || balDiff || promote) {
        out.updates.push({ row: row.row, b: sumDiff && draft ? want : null, key: keyDiff ? key : null, bank: bank,
          e: promote ? STATUS_FACT : null, d: promote ? row.d.replace(HOLD_NOTE, '').replace('⏳ в обробці', '').replace(/ · $/, '') : null });
        if (sumDiff && draft) { out.stats.resized++; out.log.push(base.concat(['сума уточнилась: ' + row.b + ' → ' + want, row.row])); }
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
    if (!best && amt > 0) best = findRow(-abs);                         // ваш рядок повернення з мінусом
    if (!best && it._feeSplit) { best = findRow(round2_(abs + it._feeSplit / 100)); withFee = !!best; }
    if (best) {
      claimed[best.r.row] = 1;
      if (withFee) feeInRow[id] = best.r.row;
      out.marks.push({ row: best.r.row, bank: bank });
      out.log.push(base.concat(['✓ уже є у журналі (ваш рядок' + (withFee ? ', разом із комісією' : '') + ')', best.r.row]));
      return;
    }
    const ts = Number(it.time) * 1000 || '';
    const pt = amt < 0 && !it._fee ? partsOf_(it.description) : null;
    if (pt) {                                                             // платіж частинами чи за розстрочкою: рядок «Боргів», його план
      const known = partsDebt_((ctx.debts || []).concat(out.newDebts), pt, abs), note0 = noteFor_(it, amt);
      const pl = partsPlan_(ctx.journal, planClaimed, abs, key, known);
      out.stats.parts++;
      if (pl) {
        planClaimed[pl.row] = 1; out.stats.plans++;
        const e = statusOf(pl.c, it.hold);
        if (e === STATUS_FACT) out.stats.facts++;
        const log = base.concat(['💳 платіж частинами: план закрито (' + pl.d.slice(0, 40) + ')', pl.row]);
        out.takes.push({ row: pl.row, key: key, b: abs, c: pl.c, d: pl.d + ' · ' + note0, e: e, bank: bank, ts: ts, log: log });
        out.log.push(log);
      } else {
        let debt = known;
        if (!debt && ctx.debts && pt.shop && pt.kind !== 'final') {          // нова покупка частинами чи розстрочка: рядок на «Борги»
          debt = { row: null, name: pt.shop.slice(0, 60).replace(/[*?]/g, ''), bank: pt.shop, dir: DEBT_OWE, monthly: abs, day: Math.min(28, Number(key.slice(8, 10))),
            from: key, kind: pt.mono ? DEBT_PARTS : DEBT_INST, cat: CAT_PARTS, closed: '', who: acc.who || '' };
          out.newDebts.push(debt);
        }
        const c = debt ? debt.cat : CAT_PARTS, e = statusOf(c, it.hold);
        if (e === STATUS_FACT) out.stats.facts++;
        const d = (debt && ctx.debts ? DEBT_TAG + debt.name + ': ' + key.slice(5, 7) + '.' + key.slice(0, 4) + ' · ' : '') + note0;
        const log = base.concat([(e === STATUS_FACT ? '✅ факт' : 'чернетка') + ' 💳 платіж частинами → ' + c + (debt ? ' («' + debt.name + '»' + (debt.row ? '' : ', новий рядок на «' + DEBTS_SHEET + '»') + ')' : ''), '']);
        out.drafts.push({ key: key, b: abs, c: c, d: d, e: e, f: '', g: acc.who || '', bank: bank, log: log, ts: ts });
        out.log.push(log);
        if (pt.kind === 'final' && debt) out.closeDebts.push({ row: debt.row, name: debt.name, key: key });
        return;
      }
      if (pt.kind === 'final' && known) out.closeDebts.push({ row: known.row, name: known.name, key: key });
      return;
    }
    const rf = refundOf_(it, ctx.journal, key, refundClaimed, mine);
    if (rf) {                                                             // повернення за покупку: категорія покупки, сума з мінусом
      let cat = rf.buy ? rf.buy.c : '', why = rf.buy ? 'за покупку ' + dd_(rf.buy.key) : '';
      if (!cat) {
        const rr = matchRule_(ctx.rules, { desc: rf.shop, comment: it.comment, mcc: it.mcc, cname: it.counterName, ciban: it.counterIban, cedrpou: it.counterEdrpou }, -amt, false);
        if (rr) { cat = rr.category; why = 'правило: ' + rr.text.slice(0, 40); }
      }
      if (rf.buy) refundClaimed[rf.buy.row] = 1;
      const e = rf.word ? statusOf(cat, it.hold) : STATUS_DRAFT;
      out.stats.refunds++;
      if (e === STATUS_FACT) out.stats.facts++;
      if (!cat) out.stats.uncategorized++;
      const note = [REFUND_NOTE + (rf.buy ? ' за покупку ' + dd_(rf.buy.key) : cat ? '' : ': оберіть категорію покупки') + (rf.word ? '' : ' (перевірте: схоже на повернення)'),
        noteFor_(it, -amt)].filter(Boolean).join(' · ');
      const log = base.concat([(e === STATUS_FACT ? '✅ факт' : 'чернетка') + ' ' + REFUND_NOTE + (cat ? ' → ' + cat + ', мінусом' : ' (без категорії)') + (why ? ' (' + why + ')' : ''), '']);
      out.drafts.push({ key: key, b: -abs, c: cat, d: note, e: e, f: '', g: acc.who || '', bank: bank, log: log, ts: ts });
      out.log.push(log);
      return;
    }
    const cz = categorize_(it, amt, ctx.rules, acc), rule = cz.rule;
    let cat = cz.cat, why = rule ? 'правило: ' + rule.text.slice(0, 40) : '', note = noteFor_(it, amt);
    if (looseCat_(cat) && !it._fee) {
      const pair = findOwnPair_(ctx.journal, pairClaimed, amt, key, acc.name);
      if (pair) {                                                         // 1) між своїми: обидва рядки
        pairClaimed[pair.row] = 1; cat = CAT_OWN_MOVE; why = 'між своїми: ' + (pair.acct || 'основна картка'); out.stats.own++;
        if (pair.c !== CAT_OWN_MOVE || pair.e === STATUS_DRAFT) {
          const e = pair.e === STATUS_DRAFT && pair.d.indexOf('⏳') < 0 ? statusOf(CAT_OWN_MOVE, false) : pair.e;
          if (e === STATUS_FACT && pair.e !== STATUS_FACT) out.stats.facts++;
          out.recats.push({ row: pair.row, c: CAT_OWN_MOVE, e: e });
        }
      }
    }
    if (!it._fee && acc.kind !== K_FOP && (looseCat_(cat) || rule)) {    // 2) план тієї ж категорії (з правилом) чи будь-який (без)
      const pl = findPlanFor_(ctx.journal, planClaimed, amt, key, looseCat_(cat) ? '' : cat, ctx.types);
      if (pl) {
        planClaimed[pl.row] = 1; out.stats.plans++;
        const log = base.concat(['план закрито: ' + pl.c + ' (' + pl.d.slice(0, 40) + ')', pl.row]);
        const e = statusOf(pl.c, it.hold);
        if (e === STATUS_FACT) out.stats.facts++;
        out.takes.push({ row: pl.row, key: key, b: abs, c: pl.c, d: pl.d + ' · ' + note, e: e, bank: bank, ts: ts, log: log });
        out.log.push(log);
        return;
      }
    }
    if (looseCat_(cat) && !it._fee) {                                     // 3) історія
      const h = histGuess_(hist, amt, it.description);
      if (h && h.cat) { cat = h.cat; why = 'за історією: ' + h.text; note += ' · за історією: ' + h.text; out.stats.learned++; }
      else if (h) note += ' · зазвичай: ' + h.text;
    }
    if (cat === CAT_TRANSFER) out.stats.transfers++;
    if (!cat) out.stats.uncategorized++;
    const e = statusOf(cat, it.hold);
    if (e === STATUS_FACT) out.stats.facts++;
    const log = base.concat([(e === STATUS_FACT ? '✅ факт' : 'чернетка') + (cat ? ' → ' + cat : ' (без категорії)') + (why ? ' (' + why + ')' : ''), '']);
    out.drafts.push({ key: key, b: abs, c: cat, d: note, e: e, f: cz.hole, g: acc.who || '', bank: bank, log: log, ts: ts });
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
    sh.getRange(r0, 1, fit.length, J_INPUT_N).setValues(fit.map(d => [Utilities.parseDate(d.key, tz, 'yyyy-MM-dd'), d.b, d.c, d.d, d.e || STATUS_DRAFT, d.f, d.g || '', '']));
    sh.getRange(r0, y, fit.length, J_MOVE_N).setValues(fit.map(d => d.bank.concat([d.ts || Date.now(), name])));
    fit.forEach((d, i) => { d.log[9] = r0 + i; });
  }
  plan.drafts.slice(fit.length).forEach(d => { d.log[0] = 'немає місця:' + d.log[0]; d.log[8] = '⚠️ не вмістилось у журнал (немає вільних рядків)'; });
  plan.marks.forEach(mk => { sh.getRange(mk.row, y, 1, BANK_N).setValues([mk.bank]); if (name) sh.getRange(mk.row, af).setValue(name); });
  (plan.takes || []).forEach(tk => {                                     // план → чернетка з банку (категорія, «Хто», «чорна діра» лишаються)
    sh.getRange(tk.row, 1, 1, 2).setValues([[Utilities.parseDate(tk.key, tz, 'yyyy-MM-dd'), tk.b]]);
    sh.getRange(tk.row, 4, 1, 2).setValues([[tk.d, tk.e || STATUS_DRAFT]]);
    sh.getRange(tk.row, y, 1, J_MOVE_N).setValues([tk.bank.concat([tk.ts || Date.now(), name])]);
  });
  (plan.recats || []).forEach(rc => { sh.getRange(rc.row, 3).setValue(rc.c); if (rc.e) sh.getRange(rc.row, 5).setValue(rc.e); });
  plan.updates.forEach(u => {
    if (u.b !== null) sh.getRange(u.row, 2).setValue(u.b);
    if (u.key) sh.getRange(u.row, 1).setValue(Utilities.parseDate(u.key, tz, 'yyyy-MM-dd'));
    if (u.d !== null && u.d !== undefined) sh.getRange(u.row, 4).setValue(u.d);
    if (u.e) sh.getRange(u.row, 5).setValue(u.e);
    sh.getRange(u.row, y, 1, BANK_N).setValues([u.bank]);
  });
  return { written: fit.length, lost: plan.drafts.length - fit.length, taken: (plan.takes || []).length };
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
