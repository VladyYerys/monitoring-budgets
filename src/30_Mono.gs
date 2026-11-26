/** 30_Mono.gs: Mono (особистий токен): картки й рахунки ФОП; імпорт з усіх банків за розкладом; повідомлення. */

/*
 * Токен Mono людина створює сама на api.monobank.ua і вводить через меню. Він зберігається лише в налаштуваннях
 * скрипта в її Google-акаунті (UserProperties): не в таблиці, не в копіях, не в репозиторії. Код лише читає виписку й баланс.
 * Сім'я: кожен учасник відкриває ту саму таблицю у своєму акаунті й підключає свій Mono: токени окремо, розклади окремо.
 * Можна вести кілька рахунків одного токена (наприклад, картку й рахунок ФОП). Контрольні точки: лише для основної картки.
 * Розклад (mbsImport) імпортує все, що підключив цей користувач: Mono і ПриватБанк.
 */

const MONO_API = 'https://api.monobank.ua';
const MONO_MAX_SPAN = 2682000;                        // 31 доба + 1 година, с
const IMPORT_HOURS = [12, 18, 23], IMPORT_MINUTE = 45; // за часом таблиці (Google запускає в межах ~15 хв від цієї хвилини)
const IMPORT_RETRY_MIN = [10, 30];
const IMPORT_SLOTS = IMPORT_HOURS.map(h => pad2_(h) + ':' + pad2_(IMPORT_MINUTE)).join(', ');

function monoFetch_(token, path) {
  for (let attempt = 0; attempt < 2; attempt++) {
    const r = UrlFetchApp.fetch(MONO_API + path, { headers: { 'X-Token': token }, muteHttpExceptions: true });
    const code = r.getResponseCode();
    if (code === 200) return JSON.parse(r.getContentText());
    if (code === 429 && attempt === 0) { Utilities.sleep(61 * 1000); continue; }
    if (code === 401 || code === 403) throw new Error('Mono не прийняв токен (код ' + code + '). Створіть новий на api.monobank.ua і підключіть знову.');
    throw new Error('Mono відповів кодом ' + code + ': ' + String(r.getContentText()).slice(0, 200));
  }
  throw new Error('Mono: забагато запитів, спробуйте за хвилину.');
}

/* ---------- рахунки Mono ---------- */

const CARD_TYPES = { black: 'Чорна', white: 'Біла', platinum: 'Platinum', iron: 'Iron', yellow: 'Жовта', fop: 'ФОП',
  eAid: 'єПідтримка', madeInUkraine: 'Зроблено в Україні' };
function cardName_(type) { return CARD_TYPES[type] || type || 'картка'; }
function cardLast4_(pan) { return (String(pan || '').match(/(\d{4})\s*$/) || [])[1] || ''; }
function cardLabel_(type, last4) { const l = cardLast4_(last4); return 'Mono ' + cardName_(type) + (l ? ' *' + l : ''); }
function cardLabelOf_(acc) { return cardLabel_(acc.type, (acc.maskedPan || [])[0] || (acc.type === 'fop' ? acc.iban : '')); }

/** Гривневі рахунки: спершу з кредитним лімітом (для кредитної картки), далі решта, рахунки ФОП у кінці. */
function cardCandidates_(accounts, credit) {
  const uah = (accounts || []).filter(a => Number(a.currencyCode) === 980);
  const score = a => (a.type === 'fop' ? 0 : 2) + (credit && Number(a.creditLimit) > 0 ? 1 : 0);
  return uah.slice().sort((a, b) => (score(b) - score(a)) || (Number(b.creditLimit) - Number(a.creditLimit)) || (Number(b.balance) - Number(a.balance)));
}

/** Збережений рахунок у відповіді банку: за id, потім за IBAN (картку перевипустили). */
function findCard_(accounts, saved) {
  const list = accounts || [];
  let a = list.filter(x => saved.id && x.id === saved.id)[0];
  if (a) return { acc: a, changed: false };
  a = list.filter(x => saved.iban && x.iban === saved.iban && Number(x.currencyCode) === 980)[0];
  return a ? { acc: a, changed: true } : { acc: null, changed: false };
}

function accountsText_(accounts) {
  const cur = { 980: '₴', 840: '$', 978: '€' };
  return (accounts || []).map(a => cardName_(a.type) + '/' + (cur[a.currencyCode] || a.currencyCode) +
    (Number(a.creditLimit) > 0 ? ' (ліміт ' + money_(a.creditLimit / 100) + ')' : '')).join(', ');
}

/** Збережені рахунки Mono цього користувача: [{id, iban, type, last4, name}]. Стара форма (одна картка) читається теж. */
function monoAccs_() {
  const up = PropertiesService.getUserProperties();
  const list = jsonProp_(up, 'mono_accs');
  if (list && list.length) return list;
  const id = up.getProperty('mono_account');
  return id ? [{ id: id, iban: up.getProperty('mono_iban') || '', type: up.getProperty('mono_type') || '', last4: up.getProperty('mono_pan') || '',
    name: cardLabel_(up.getProperty('mono_type'), up.getProperty('mono_pan')) }] : [];
}
function savedCardLabel_() { return monoAccs_().map(a => a.name).join(', '); }

function cardIsCredit_(ss) {
  const s = ss.getSheetByName(SETTINGS_SHEET);
  return !s || String(s.getRange(SET.card).getValue()) !== CARD_DEBIT;
}

/* ---------- меню ---------- */

function connectMono() {
  if (!menuOk_('🏦 Mono')) return;
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  const r = ui.prompt('🏦 Підключити Mono', 'Вставте особистий токен з api.monobank.ua.\n' +
    'Він зберігається лише в налаштуваннях скрипта у вашому Google-акаунті: не в таблиці й не в копіях.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const token = String(r.getResponseText() || '').trim();
  if (!/^[A-Za-z0-9_\-]{20,}$/.test(token)) { show_('🏦 Mono', ['Це не схоже на токен (довгий рядок латинських літер і цифр). Спробуйте ще раз.']); return; }
  let info;
  try { info = monoFetch_(token, '/personal/client-info'); } catch (e) { show_('🏦 Mono', ['⚠️ ' + esc_(e.message)]); return; }
  const cands = cardCandidates_(info.accounts, cardIsCredit_(ss));
  if (!cands.length) { show_('🏦 Mono', ['⚠️ Гривневих рахунків не знайдено. Рахунки: ' + esc_(accountsText_(info.accounts))]); return; }
  let chosen = [cands[0]];
  if (cands.length > 1) {
    const q = ui.prompt('🏦 Які рахунки вести?', cands.map((a, i) => (i + 1) + ') ' + cardLabelOf_(a) + (a.type === 'fop' ? ' (рахунок ФОП)' : '') +
      (Number(a.creditLimit) > 0 ? ', ліміт ' + money_(a.creditLimit / 100) : '')).join('\n') + '\n\nНомери через кому (порожньо: 1):', ui.ButtonSet.OK_CANCEL);
    if (q.getSelectedButton() !== ui.Button.OK) return;
    const nums = String(q.getResponseText() || '1').split(/[,\s]+/).map(Number).filter(n => n >= 1 && n <= cands.length);
    chosen = (nums.length ? nums : [1]).filter((n, i, a) => a.indexOf(n) === i).map(n => cands[n - 1]);
  }
  const who = askWho_('🏦 Mono', 'це рахунки');
  if (who === null) return;
  const saved = chosen.map(a => {
    const fop = a.type === 'fop', acc = accEnsure_(ss, { key: 'mono:' + md5Short_(a.id), name: fop ? 'Mono ФОП *' + cardLast4_(a.iban) : cardLabelOf_(a), bank: 'Mono',
      who: who, kind: fop ? K_FOP : undefined });
    return { id: a.id, iban: a.iban || '', type: a.type || '', last4: cardLast4_(fop ? a.iban : (a.maskedPan || [])[0]), name: acc.name, kind: acc.kind };
  });
  const up = PropertiesService.getUserProperties();
  up.setProperties({ mono_token: token, mono_accs: JSON.stringify(saved.map(s => ({ id: s.id, iban: s.iban, type: s.type, last4: s.last4, name: s.name }))) });
  ['mono_account', 'mono_iban', 'mono_type', 'mono_pan'].forEach(k => up.deleteProperty(k));
  const main = chosen.find((a, i) => saved[i].kind === K_MAIN);
  show_('🏦 Mono підключено', ['Рахунки: <b>' + saved.map(s => esc_(s.name) + ' (' + esc_(s.kind) + ')').join(', ') + '</b>.',
    main ? 'Зараз на основній картці <b>' + money_((main.balance - main.creditLimit) / 100) + '</b>' + (Number(main.creditLimit) ? ', кредитний ліміт <b>' +
      money_(main.creditLimit / 100) + '</b>' : '') + '.' : '',
    'Вид рахунку можна змінити на аркуші «' + ACC_SHEET + '». Далі: «🏦 Імпорт з банку зараз» (перший імпорт: сьогодні й 2 попередні дні).',
    'Щодня самі: «🚀 Підключення», крок 1, «Імпорт з банку щодня» (' + IMPORT_SLOTS + ').']);
}

function disconnectMono() {
  if (!menuOk_('🏦 Mono')) return;
  ['mono_token', 'mono_accs', 'mono_account', 'mono_pan', 'mono_iban', 'mono_type'].forEach(k => PropertiesService.getUserProperties().deleteProperty(k));
  if (!PropertiesService.getUserProperties().getProperty('pb_token')) { removeTriggers_(TRG.imp); removeTriggers_(TRG.impRetry); }
  show_('🏦 Mono відключено', ['Токен стерто. Сам токен можна відкликати й на api.monobank.ua. Рахунки на аркуші «' + ACC_SHEET + '» лишились (там історія).']);
}

function enableDailyImport() {
  if (!menuOk_('🏦 Імпорт за розкладом')) return;
  const up = PropertiesService.getUserProperties();
  if (!up.getProperty('mono_token') && !up.getProperty('pb_token')) { show_('🏦 Імпорт за розкладом', ['Спершу підключіть Mono або ПриватБанк («🚀 Підключення», крок 1).']); return; }
  importTriggers_(SpreadsheetApp.getActive().getSpreadsheetTimeZone());
  show_('🏦 Імпорт за розкладом', ['Увімкнено: щодня близько <b>' + IMPORT_SLOTS + '</b> (для ваших рахунків; кожен учасник сім\'ї вмикає свій).',
    tgReady_() ? 'Після кожного імпорту: повідомлення в Telegram.' : 'Telegram не підключено: повідомлень не буде.',
    'Не вдалося: повтор сам через ' + IMPORT_RETRY_MIN.join(' і ще через ') + ' хв. Історія запусків: «🩺 Стан таблиці».']);
}

function disableDailyImport() {
  if (!menuOk_('🏦 Імпорт за розкладом')) return;
  removeTriggers_(TRG.imp); removeTriggers_(TRG.impRetry);
  show_('🏦 Імпорт за розкладом', ['Вимкнено. Вручну: «🏦 Імпорт з банку зараз».']);
}

function importTriggers_(tz) {
  removeTriggers_(TRG.imp); removeTriggers_(TRG.impRetry);
  IMPORT_HOURS.forEach(h => ScriptApp.newTrigger(TRG.imp).timeBased().atHour(h).nearMinute(IMPORT_MINUTE).everyDays(1).inTimezone(tz).create());
}

function importNow() {
  if (!menuOk_('🏦 Імпорт з банку')) return;
  let r;
  try { r = importBank_(SpreadsheetApp.getActive()); } catch (e) { show_('🏦 Імпорт з банку', ['⚠️ ' + esc_(e.message)]); return; }
  show_('🏦 Імпорт з банку', importReport_(r));
}

/* ---------- за розкладом ---------- */

function mbsImport() { if (quietOk_()) return runImport_(0); }
function mbsImportRetry() { if (quietOk_()) return runImport_(Number(PropertiesService.getDocumentProperties().getProperty('imp_retry') || 1)); }

function runImport_(attempt) {
  const ss = SpreadsheetApp.getActive(), tz = ss.getSpreadsheetTimeZone(), dp = PropertiesService.getDocumentProperties();
  const at = Utilities.formatDate(new Date(), tz, 'HH:mm'), wait = IMPORT_RETRY_MIN[attempt];
  removeTriggers_(TRG.impRetry);
  if (wait) { dp.setProperty('imp_retry', String(attempt + 1)); ScriptApp.newTrigger(TRG.impRetry).timeBased().after(wait * 60 * 1000).create(); }
  else dp.deleteProperty('imp_retry');
  let r = null, err = '';
  try { r = importBank_(ss); } catch (e) { err = String(e && e.message || e); }
  const rec = { t: new Date().toISOString(), ok: !err, try: attempt + 1, msg: err.slice(0, 200), n: r ? r.drafts : 0, tg: '' };
  if (!err) {
    beat_('import', true);
    removeTriggers_(TRG.impRetry); dp.deleteProperty('imp_retry');
    if (tgReady_()) { try { tgSend_(importText_(ss, r, { at: at, attempt: attempt })); rec.tg = 'ok'; } catch (e) { rec.tg = String(e && e.message || e).slice(0, 120); } }
    impHist_(rec);
    return r;
  }
  beat_('import', false, err);
  if (tgReady_()) {
    try {
      tgSend_('⚠️ <b>Імпорт з банку не вдався</b> · ' + at + (attempt ? ' (спроба ' + (attempt + 1) + ')' : '') + '\nПричина: ' + esc_(err) + '\n' +
        (wait ? '🔁 Повторю сам за ' + wait + ' хв.' : '❗ Повтори теж не вдались. Запустіть вручну: «🏦 Імпорт з банку зараз».'));
      rec.tg = 'ok';
    } catch (e) { rec.tg = String(e && e.message || e).slice(0, 120); }
  }
  impHist_(rec);
  if (!wait) throw new Error('Імпорт з банку не вдався (і повтори теж): ' + err);
  return null;
}

function impHist_(rec) {
  const dp = PropertiesService.getDocumentProperties();
  const h = jsonProp_(dp, 'imp_hist') || [];
  if (rec) h.push(rec);
  dp.setProperty('imp_hist', JSON.stringify(h.slice(-15)));
  return h;
}

/* ---------- імпорт з усіх банків цього користувача ---------- */

/** Mono і ПриватБанк (що підключено). Помилка одного джерела не зупиняє інше; обидва впали: помилка. */
function importBank_(ss) {
  guard_(ss);
  if (!ss.getSheetByName(JOURNAL) || !ss.getSheetByName(BANK_SHEET) || !ss.getSheetByName(RULES_SHEET) || !ss.getSheetByName(ACC_SHEET))
    throw new Error('Немає аркушів журналу чи банку: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».');
  const up = PropertiesService.getUserProperties(), dp = PropertiesService.getDocumentProperties();
  const jobs = [];
  if (up.getProperty('mono_token') && monoAccs_().length) jobs.push(['Mono', importMono_]);
  if (up.getProperty('pb_token')) jobs.push(['ПриватБанк', importPrivat_]);
  if (!jobs.length) throw new Error('Банк не підключено: «🚀 Підключення», крок 1 (Mono, ПриватБанк для ФОП або виписка).');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60 * 1000)) throw new Error('Інша дія скрипта ще триває, спробуйте за хвилину.');
  const rep = emptyReport_(), errs = [];
  try {
    jobs.forEach(j => {
      try { mergeReport_(rep, j[1](ss)); } catch (e) { errs.push(j[0] + ': ' + String(e && e.message || e)); }
    });
    if (errs.length === jobs.length) throw new Error(errs.join('; '));
    rep.errors = errs;
    rep.sorted = sortCore_(ss);
    const bankSh = ss.getSheetByName(BANK_SHEET), stamp = new Date();
    bankSh.getRange('B2').setValue(stamp);
    bankSh.getRange('D2').setValue(importSummary_(rep));
    bankSh.getRange('F2').setValue(accList_(ss).map(a => a.name).join(', '));
    dp.setProperties({ bank_last_ok: stamp.toISOString(), bank_last_error: errs.join('; ') });
    if (errs.length) dp.setProperty('bank_last_error_at', stamp.toISOString());
    return rep;
  } catch (e) {
    dp.setProperties({ bank_last_error: String(e && e.message || e), bank_last_error_at: new Date().toISOString() });
    try { ss.getSheetByName(BANK_SHEET).getRange('D2').setValue('⚠️ ' + String(e && e.message || e)); } catch (x) { /* ignore */ }
    throw e;
  } finally { lock.releaseLock(); }
}

function emptyReport_() {
  return { drafts: 0, lost: 0, transfers: 0, uncategorized: 0, matched: 0, other: 0, updated: 0, redated: 0, changed: 0, seen: 0, fees: 0, skipped: 0, vanished: 0,
    checkpoints: [], own: null, limit: null, newOps: [], todayOps: { out: 0, inc: 0, n: 0 }, notes: [], accounts: [] };
}

function mergeReport_(rep, r) {
  ['drafts', 'lost', 'transfers', 'uncategorized', 'matched', 'other', 'updated', 'redated', 'changed', 'seen', 'fees', 'skipped', 'vanished']
    .forEach(k => { rep[k] += r[k] || 0; });
  ['checkpoints', 'newOps', 'notes', 'accounts'].forEach(k => { rep[k] = rep[k].concat(r[k] || []); });
  ['out', 'inc', 'n'].forEach(k => { rep.todayOps[k] = round2_(rep.todayOps[k] + ((r.todayOps || {})[k] || 0)); });
  if (r.own !== null && r.own !== undefined) { rep.own = r.own; rep.limit = r.limit; }
  return rep;
}

/** Звіт одного рахунку з результату bankIngest_. */
function accReport_(acc, items, res, dateOf, today) {
  const p = res.plan, r = emptyReport_();
  Object.assign(r, { drafts: res.written, lost: res.lost, transfers: p.stats.transfers, uncategorized: p.stats.uncategorized, matched: p.marks.length,
    other: p.stats.other, updated: p.stats.resized, redated: p.stats.redated, changed: p.stats.changed, seen: items.length, vanished: p.vanish.length, accounts: [acc.name] });
  r.newOps = p.drafts.slice(0, res.written).map(d => ({ key: d.key, time: String(d.bank[4] || '').slice(11, 16), amt: Number(d.bank[2]) || 0, cat: d.c || '',
    desc: String(d.bank[1] || ''), acc: acc.name }));
  items.forEach(it => { if (dateOf(it) !== today) return; const a = Number(it.amount) / 100; r.todayOps.n++;
    if (a < 0) r.todayOps.out = round2_(r.todayOps.out - a); else r.todayOps.inc = round2_(r.todayOps.inc + a); });
  return r;
}

/** Усі рахунки Mono цього користувача. Між виписками різних рахунків: 61 с (обмеження Mono). */
function importMono_(ss) {
  const tz = ss.getSpreadsheetTimeZone(), up = PropertiesService.getUserProperties(), dp = PropertiesService.getDocumentProperties();
  const token = up.getProperty('mono_token'), saved = monoAccs_(), accs = accList_(ss), rep = emptyReport_();
  const today = todayKey_(ss), nowSec = Math.floor(Date.now() / 1000), days = [addDays_(today, -2), addDays_(today, -1), today];
  const startSec = Math.floor(Utilities.parseDate(days[0] + ' 00:00:00', tz, 'yyyy-MM-dd HH:mm:ss').getTime() / 1000);
  const dateOf = it => Utilities.formatDate(new Date(it.time * 1000), tz, 'yyyy-MM-dd');
  const timeOf = it => Utilities.formatDate(new Date(it.time * 1000), tz, 'dd.MM.yyyy HH:mm');
  const info = monoFetch_(token, '/personal/client-info');
  saved.forEach((s, i) => {
    const found = findCard_(info.accounts, s), a = found.acc;
    if (!a) { rep.notes.push(s.name + ': рахунку більше немає в банку (перевипустили?): «🏦 Підключити Mono» ще раз'); return; }
    const k = 'mono:' + md5Short_(s.id), acc = accList_(ss).find(x => x.key === k) || accEnsure_(ss, { key: k, name: s.name, bank: 'Mono' });
    if (found.changed) { dp.deleteProperty('bank_last_to:' + k); rep.notes.push(s.name + ': картку перевипущено, веду нову'); }
    if (i > 0) Utilities.sleep(61 * 1000);
    const fromSec = found.changed ? startSec : importWindowFrom_(nowSec, startSec, Number(dp.getProperty('bank_last_to:' + k) || (i === 0 ? dp.getProperty('bank_last_to') : 0) || 0));
    const items = fetchStatement_(token, a.id, fromSec, nowSec);
    const k0 = Utilities.formatDate(new Date(fromSec * 1000), tz, 'yyyy-MM-dd');
    const coverFrom = fromSec <= Utilities.parseDate(k0 + ' 00:00:00', tz, 'yyyy-MM-dd HH:mm:ss').getTime() / 1000 ? k0 : addDays_(k0, 1);
    const exp = expandOps_(items, dp.getProperty('mono_fee_included:' + k) || dp.getProperty('mono_fee_included'));
    if (exp.learned) dp.setProperty('mono_fee_included:' + k, exp.learned);
    const bs = balanceShift_(items, a, dp.getProperty('mono_balance_mode:' + k) || dp.getProperty('mono_balance_mode'));
    if (bs.learned) dp.setProperty('mono_balance_mode:' + k, bs.learned);
    const skipLog = exp.skipped.map(it => ['skip:' + it.id, dateOf(it), round2_(it.amount / 100), String(it.description || ''), Number(it.mcc || 0), '', 0, 'так',
      'пропущено: скасування ще не списаної оплати (у залишку банку його немає)', '']);
    const res = bankIngest_(ss, { acc: acc, items: exp.ops, dateOf: dateOf, timeOf: timeOf, balShift: bs.shift, src: 'mono',
      coverFrom: found.changed ? null : coverFrom, coverTo: found.changed ? null : today, extraLog: skipLog });
    const r = accReport_(acc, exp.ops, res, dateOf, today);
    r.fees = exp.fees; r.skipped = exp.skipped.length;
    if (acc.kind === K_MAIN) {
      const own = (a.balance - a.creditLimit) / 100, limit = a.creditLimit / 100;
      r.checkpoints = upsertCheckpoints_(ss, checkpointPlan_(exp.ops, own, limit, days, dateOf));
      r.own = own; r.limit = limit;
    }
    mergeReport_(rep, r);
    dp.setProperty('bank_last_to:' + k, String(nowSec));
  });
  return rep;
}

/** Чиста функція: початок вікна виписки (с). Перекриття 1 доба; не довше 31 доби. */
function importWindowFrom_(nowSec, startSec, lastTo) {
  const from = lastTo ? Math.min(startSec, lastTo - 86400) : startSec;
  return Math.max(from, nowSec - MONO_MAX_SPAN + 60);
}

function fetchStatement_(token, accountId, fromSec, toSec) {
  const seen = {}, all = [];
  let to = toSec;
  for (let page = 0; page < 6; page++) {
    const items = monoFetch_(token, '/personal/statement/' + accountId + '/' + fromSec + '/' + to);
    items.forEach(it => { if (!seen[it.id]) { seen[it.id] = 1; all.push(it); } });
    if (items.length < 500) break;
    to = items[items.length - 1].time;
    Utilities.sleep(61 * 1000);
  }
  return all;
}

/**
 * Чиста функція: операції банку → рядки для журналу. Операція з комісією стає двома: сума і окремий рядок комісії
 * (номер «…#fee»). Чи входить комісія в amount, визначаємо за залишком після операції; висновок запам'ятовується.
 */
function expandOps_(items, savedMode) {
  const skipped = items.filter(isPendingReversal_);
  const sorted = items.filter(it => !isPendingReversal_(it)).sort((a, b) => a.time - b.time);
  let mode = savedMode === 'так' || savedMode === 'ні' ? savedMode : '', learned = '';
  sorted.forEach((it, i) => {
    const fee = Number(it.commissionRate || 0), prev = sorted[i - 1];
    if (!fee || it.amount >= 0 || !prev || typeof it.balance !== 'number' || typeof prev.balance !== 'number') return;
    const delta = it.balance - prev.balance;
    if (Math.abs(delta - it.amount) <= 1) learned = 'так';
    else if (Math.abs(delta - (it.amount - fee)) <= 1) learned = 'ні';
  });
  if (learned) mode = learned;
  const included = mode !== 'ні', ops = [];
  let fees = 0;
  sorted.forEach(it => {
    const fee = Number(it.commissionRate || 0), bal = typeof it.balance === 'number' ? it.balance : null;
    if (!fee || it.amount >= 0) { ops.push(Object.assign({}, it, { _bal: bal })); return; }
    const total = included ? -it.amount : -it.amount + fee;
    ops.push(Object.assign({}, it, { amount: -(total - fee), commissionRate: 0, _feeSplit: fee, _bal: bal === null ? null : bal + fee }));
    ops.push({ id: it.id + '#fee', time: it.time, description: 'Комісія · ' + String(it.description || ''), mcc: it.mcc, amount: -fee,
      hold: it.hold, comment: '', _fee: true, _parent: it.id, _total: total, _bal: bal });
    fees++;
  });
  return { ops: ops, learned: learned && learned !== savedMode ? learned : '', mode: mode || '?', fees: fees, skipped: skipped };
}

/** Чиста функція: як перевести «залишок після операції» з виписки у ваш залишок (мінус: у кредиті). */
function balanceShift_(items, acc, saved) {
  const withBal = items.filter(it => typeof it.balance === 'number').sort((a, b) => a.time - b.time);
  const last = withBal[withBal.length - 1];
  let mode = saved === 'own' || saved === 'limit' ? saved : '', learned = '';
  if (last) {
    if (Math.abs(last.balance - acc.balance) <= 1) learned = 'limit';
    else if (Math.abs(last.balance - (acc.balance - acc.creditLimit)) <= 1) learned = 'own';
  }
  if (learned) mode = learned;
  return { shift: mode === 'own' ? 0 : Number(acc.creditLimit || 0), learned: learned && learned !== saved ? learned : '' };
}

/** Чиста функція: залишок на початок кожного з днів = залишок зараз − операції від цього дня. */
function checkpointPlan_(items, ownNow, limit, days, dateOf) {
  return days.map(day => {
    const after = items.filter(it => dateOf(it) >= day).reduce((s, it) => s + it.amount, 0) / 100;
    return { key: day, own: round2_(ownNow - after), limit: round2_(limit) };
  });
}

/** Контрольні точки на «Залишках банку»: оновлює наявну дату або пише в перший вільний рядок. */
function upsertCheckpoints_(ss, cps) {
  const sh = ss.getSheetByName(CP_SHEET), tz = ss.getSpreadsheetTimeZone();
  const vals = sh.getRange(CP_FIRST, 1, CP_LAST - CP_FIRST + 1, 3).getValues(), keys = vals.map(v => dateKey_(v[0], tz));
  return cps.map(cp => {
    let i = keys.indexOf(cp.key);
    const old = i >= 0 ? vals[i] : null;
    if (old && old[1] !== '' && old[2] !== '' && Math.abs(Number(old[1]) - cp.own) < 0.005 && Math.abs(Number(old[2]) - cp.limit) < 0.005)
      return { key: cp.key, own: cp.own, limit: cp.limit, state: 'без змін' };
    if (i < 0) {
      i = keys.findIndex((k, j) => k === '' && vals[j][1] === '' && vals[j][2] === '');
      if (i < 0) return { key: cp.key, state: '⚠️ немає місця' };
      keys[i] = cp.key;
    }
    sh.getRange(CP_FIRST + i, 1, 1, 3).setValues([[Utilities.parseDate(cp.key, tz, 'yyyy-MM-dd'), cp.own, cp.limit]]);
    vals[i] = [cp.key, cp.own, cp.limit];
    return { key: cp.key, own: cp.own, limit: cp.limit, state: old ? 'оновлено (було ' + old[1] + ')' : 'додано' };
  });
}

/* ---------- звіти про імпорт ---------- */

function importSummary_(r) {
  return 'нових чернеток ' + r.drafts + (r.transfers ? ' (переказів ' + r.transfers + ')' : '') + (r.uncategorized ? ', без категорії ' + r.uncategorized : '') +
    ' · ваших рядків знайдено ' + r.matched + (r.other ? ' · з іншого джерела ' + r.other : '') + (r.lost ? ' · ⚠️ не вмістилось ' + r.lost : '') +
    (r.changed ? ' · ⚠️ розбіжностей сум ' + r.changed : '') + ((r.errors || []).length ? ' · ⚠️ ' + r.errors.join('; ') : '');
}

function importReport_(r) {
  const lines = ['Рахунки: ' + esc_(r.accounts.filter((a, i, x) => x.indexOf(a) === i).join(', ') || 'немає') + '. Операцій у виписках: ' + r.seen + '.',
    '<b>Нових чернеток: ' + r.drafts + '</b>' + (r.transfers ? ': з них переказів (червоні): <b>' + r.transfers + '</b>' : '') +
      (r.uncategorized ? ', без категорії: <b>' + r.uncategorized + '</b>' : '') + '.',
    'Ваших рядків знайдено у виписці (без дубля): <b>' + r.matched + '</b>.' + (r.other ? ' Уже були з іншого джерела: <b>' + r.other + '</b>.' : '')];
  r.notes.forEach(n => lines.push('ℹ️ ' + esc_(n) + '.'));
  (r.errors || []).forEach(n => lines.push('⚠️ ' + esc_(n)));
  if (r.fees) lines.push('Комісій окремими рядками «' + CAT_FEE + '»: ' + r.fees + '.');
  if (r.updated) lines.push('Уточнено суму в чернетках (операції «в обробці»): ' + r.updated + '.');
  if (r.redated) lines.push('Уточнено дату в чернетках: ' + r.redated + '.');
  if (r.vanished) lines.push('Прибрано чернеток, які банк скасував: ' + r.vanished + ' (деталі на аркуші «' + BANK_SHEET + '»).');
  if (r.skipped) lines.push('Пропущено «Скасування … в обробці»: ' + r.skipped + '.');
  if (r.changed) lines.push('⚠️ Розбіжність суми з підтвердженим рядком: ' + r.changed + ': деталі на аркуші «' + BANK_SHEET + '».');
  if (r.lost) lines.push('⚠️ Не вмістилось у журнал: ' + r.lost + ': потрібні вільні рядки.');
  if (r.checkpoints.length) lines.push('<b>Контрольні точки основної картки (на початок дня / ліміт):</b><br>' + r.checkpoints.map(c =>
    dd_(c.key) + ': ' + (c.own !== undefined ? money_(c.own) + ' / ' + money_(c.limit) : '') + ': ' + esc_(c.state)).join('<br>'));
  if (r.own !== null) lines.push('Зараз на основній картці: ' + money_(r.own) + (r.limit ? ' (ліміт ' + money_(r.limit) + ')' : '') + '.');
  return lines;
}

/** Повідомлення в Telegram після імпорту за розкладом. */
function importText_(ss, r, meta) {
  const ops = r.newOps || [], t = r.todayOps || { out: 0, inc: 0, n: 0 }, today = todayKey_(ss);
  const out = ops.filter(o => o.amt < 0).reduce((x, o) => x - o.amt, 0), inc = ops.filter(o => o.amt > 0).reduce((x, o) => x + o.amt, 0);
  let own = null, rows = [];
  try { SpreadsheetApp.flush(); own = mainByLabel_(ss)('Доступно моїх сьогодні'); rows = readJournal_(ss).filter(o => !o.blank); } catch (e) { /* без цього */ }
  const signed = x => (x > 0 ? '+' : '') + money0_(x), many = r.accounts.filter((a, i, x) => x.indexOf(a) === i).length > 1;
  const lines = ['🏦 <b>Банк · ' + meta.at + '</b> ✓' + (meta.attempt ? ' (з ' + (meta.attempt + 1) + '-ї спроби)' : '')];
  lines.push(ops.length ? 'Нових операцій: <b>' + ops.length + '</b>' + (out ? ' · списано <b>' + money0_(out) + '</b>' : '') +
    (inc ? ' · надійшло <b>' + money0_(inc) + '</b>' : '') : 'Нових операцій немає.');
  ops.slice(0, 10).forEach(o => lines.push('• ' + (o.key && o.key !== today ? dd_(o.key) + ' ' : '') + esc_(o.time) + ' ' + esc_(o.cat || 'без категорії') +
    ' <b>' + signed(o.amt) + '</b>' + (o.desc ? ': ' + esc_(String(o.desc).slice(0, 32)) : '') + (many ? ' (' + esc_(o.acc) + ')' : '')));
  if (ops.length > 10) lines.push('• … ще ' + (ops.length - 10));
  lines.push('📆 Сьогодні: списано <b>' + money0_(t.out) + '</b>' + (t.inc ? ', надійшло ' + money0_(t.inc) : '') + ' (' + t.n + ' опер.)');
  if (r.own !== null) lines.push((own !== null ? '💚 Мої: <b>' + money0_(own) + '</b> · ' : '') + '💳 на основній картці <b>' + money0_(r.own) + '</b>' +
    (r.limit ? ' (з лімітом доступно ' + money0_(Number(r.own) + Number(r.limit)) + ')' : ''));
  const bud = budgetLines_(ss, rows, ops.map(o => o.cat).filter(Boolean));
  if (bud.length) lines.push(bud.map(x => x.text).join('\n'));
  const drafts = rows.filter(o => o.e === STATUS_DRAFT).length;
  if (drafts) lines.push('🏦 Чернетки чекають «✅ факт»: <b>' + drafts + '</b>');
  if (r.transfers) lines.push('↔ Нових переказів: ' + r.transfers + ': оберіть категорію.');
  (r.errors || []).forEach(e => lines.push('⚠️ ' + esc_(e)));
  if (r.lost) lines.push('⚠️ Не вмістилось у журнал: ' + r.lost + '.');
  if (r.changed) lines.push('⚠️ Сума в банку не збігається з вашим рядком: ' + r.changed + ': деталі на аркуші «' + BANK_SHEET + '».');
  return lines.join('\n');
}

/** Стан імпорту для «🩺 Стан таблиці». */
function bankStatusLines_(ss) {
  const tz = ss.getSpreadsheetTimeZone(), up = PropertiesService.getUserProperties(), dp = PropertiesService.getDocumentProperties();
  const fmt = iso => iso ? Utilities.formatDate(new Date(iso), tz, 'dd.MM.yyyy HH:mm') : 'ще не було';
  const sched = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === TRG.imp).length;
  const lines = ['Mono (ваш акаунт): ' + (up.getProperty('mono_token') ? '<b>підключено</b>: ' + esc_(savedCardLabel_()) : 'не підключено'),
    'ПриватБанк для ФОП (ваш акаунт): ' + (up.getProperty('pb_token') ? '<b>підключено</b>: ' + esc_(privatAccs_().map(a => a.name).join(', ')) : 'не підключено'),
    'Рахунки в таблиці: ' + esc_(accList_(ss).map(a => a.name + ' (' + a.kind + (a.who ? ', ' + a.who : '') + ')').join('; ') || 'ще немає'),
    'Імпорт за розкладом (ваш): <b>' + (sched ? IMPORT_SLOTS : 'вимкнено') + '</b>' + (sched ? (tgReady_() ? ', після кожного: Telegram' : ', Telegram не підключено') : ''),
    'Останній успішний імпорт: <b>' + fmt(dp.getProperty('bank_last_ok')) + '</b>'];
  const hist = impHist_(null).slice(-8).reverse();
  if (hist.length) lines.push('<b>Запуски за розкладом</b> (новіші зверху):<br>' + hist.map(h => fmt(h.t) + ': ' +
    (h.ok ? '✓ нових ' + (h.n || 0) : '⚠️ ' + esc_(h.msg || 'помилка')) + (h.try > 1 ? ' (спроба ' + h.try + ')' : '') +
    (h.tg && h.tg !== 'ok' ? ' · Telegram: ' + esc_(h.tg) : '')).join('<br>'));
  const errAt = dp.getProperty('bank_last_error_at');
  if (dp.getProperty('bank_last_error') && errAt && errAt >= (dp.getProperty('bank_last_ok') || '')) lines.push('⚠️ Остання помилка (' + fmt(errAt) + '): ' + esc_(dp.getProperty('bank_last_error')));
  return lines;
}
