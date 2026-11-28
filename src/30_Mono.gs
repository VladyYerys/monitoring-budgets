/** 30_Mono.gs: Mono (особистий токен): картки й рахунки ФОП; імпорт з усіх банків за розкладом; повідомлення. */

/*
 * Токен Mono людина створює сама на api.monobank.ua і вводить через меню. Він зберігається лише в налаштуваннях
 * скрипта в її Google-акаунті (UserProperties): не в таблиці, не в копіях, не в репозиторії. Код лише читає виписку й баланс.
 * Сім'я: кожен учасник відкриває ту саму таблицю у своєму акаунті й підключає свій Mono: токени окремо, розклади окремо.
 * Можна вести кілька рахунків одного токена (наприклад, картку й рахунок ФОП). Контрольні точки: лише для основної картки.
 * Розклад (mbsImport) імпортує все, що підключив цей користувач: Mono і ПриватБанк.
 * «🗓 Імпорт з банку за період…»: місяці назад, по 31 добі у фоні (mbsPeriod), див. розділ нижче.
 */

const MONO_API = 'https://api.monobank.ua';
const MONO_MAX_SPAN = 2682000;                        // 31 доба + 1 година, с
const IMPORT_HOURS = [12, 18, 23], IMPORT_MINUTE = 45; // за часом таблиці (Google запускає в межах ~15 хв від цієї хвилини)
const IMPORT_RETRY_MIN = [10, 30];
const IMPORT_SLOTS = IMPORT_HOURS.map(h => pad2_(h) + ':' + pad2_(IMPORT_MINUTE)).join(', ');

/** Паузи перед повтором, коли Mono приймає сервер Google за VPN (наступний запит часто йде з іншої адреси), мс. */
const MONO_VPN_WAIT = [5000, 20000];
function monoVpnText_() {
  return 'Mono тимчасово заблокував з\'єднання: сервери Google, з яких працює скрипт таблиці, для Mono іноді схожі на VPN. ' +
    'Це не у вас і не з токеном. Зачекайте 5 хвилин, оновіть сторінку таблиці й запустіть «🏦 Імпорт з банку зараз» ще раз. ' +
    'За розкладом скрипт сам повторить за 10 і 30 хв.';
}

/** Помилка Mono з тіла відповіді: { code: errCode, text: errText / errTitle / errorDescription }. */
function monoErr_(body) {
  try {
    const o = JSON.parse(body);
    return { code: String(o.errCode || ''), text: String(o.errText || o.errTitle || o.errorDescription || o.message || '').slice(0, 200) };
  } catch (e) { return { code: '', text: String(body || '').slice(0, 200) }; }
}

function monoFetch_(token, path) {
  let busy = 0, vpn = 0;
  for (;;) {
    const r = UrlFetchApp.fetch(MONO_API + path, { headers: { 'X-Token': token }, muteHttpExceptions: true });
    const code = r.getResponseCode();
    if (code === 200) return JSON.parse(r.getContentText());
    const e = monoErr_(r.getContentText());
    if (code === 429) { if (busy++ === 0) { Utilities.sleep(61 * 1000); continue; } throw new Error('Mono: забагато запитів, спробуйте за хвилину.'); }
    if (e.code === 'VPN_BLOCKED') { if (vpn < MONO_VPN_WAIT.length) { Utilities.sleep(MONO_VPN_WAIT[vpn++]); continue; } throw new Error(monoVpnText_()); }
    if (code === 401 || code === 403) throw new Error('Mono не прийняв токен (код ' + code + '). Створіть новий на api.monobank.ua і підключіть знову.');
    throw new Error('Mono відповів кодом ' + code + (e.text ? ': ' + e.text : '') + (e.code ? ' (' + e.code + ')' : ''));
  }
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
  removeTriggers_(TRG.period); PropertiesService.getUserProperties().deleteProperty('period_job');
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

/* ---------- 🔎 сирі дані Mono: що банк реально віддає (нічого не пише в таблицю) ---------- */

const RAW_DAYS = 31;
const rawIban_ = x => { x = String(x || ''); return x.length > 10 ? x.slice(0, 4) + '…' + x.slice(-4) : x; };
const rawCode_ = x => { x = String(x || ''); return x.length === 10 ? x.slice(0, 2) + '…' + x.slice(-2) : x; };   // ІПН людини: середина прихована; ЄДРПОУ (8): повністю
const RAW_SALARY = /заробітн|зарплат|з\/п|salary/i, RAW_PARTS = /частин|погашенн|purchase in parts/i;

/**
 * Чиста функція: операції Mono → рядки вікна й перевірки припущень (де «заробітна плата»; чи є окремий рядок покупки частинами).
 * → { rows: [{ t, amt, desc, comment, mcc, omcc, name, code, iban, hold, cat, why, flag }], sum: { … } }
 */
function rawRows_(items, rules, acc, timeOf, hist) {
  const rows = items.slice().sort((a, b) => b.time - a.time).map(it => {
    const amt = round2_(it.amount / 100), refund = amt > 0 && REFUND_RE.test(String(it.description || ''));
    const cz = refund ? { cat: REFUND_NOTE + ' (мінусом)', rule: null } : categorize_(it, amt, rules, acc), h = !cz.cat || cz.cat === CAT_TRANSFER ? histGuess_(hist || {}, amt, it.description) : null;
    const where = ['description', 'comment', 'counterName'].filter(k => RAW_SALARY.test(String(it[k] || '')));
    const pt = partsOf_(it.description), parts = !!pt || RAW_PARTS.test(String(it.description || '') + ' ' + String(it.comment || ''));
    return { t: timeOf(it), amt: amt, desc: String(it.description || ''), comment: String(it.comment || ''), mcc: Number(it.mcc || 0), omcc: Number(it.originalMcc || 0),
      name: String(it.counterName || ''), code: rawCode_(it.counterEdrpou), iban: rawIban_(it.counterIban), hold: !!it.hold,
      cat: cz.cat || (h && h.cat) || '', why: refund ? 'категорія покупки з журналу (той самий магазин, до 60 днів)' : cz.rule ? 'правило: ' + cz.rule.text.slice(0, 30) + (cz.rule.text.length > 30 ? '…' : '') : h && h.cat ? 'історія: ' + h.text : isTransfer_(it) ? 'схоже на переказ' : cz.cat ? '' : 'немає правила',
      salary: where, parts: pt ? { first: 'перший платіж', pay: 'погашення', final: 'повне погашення' }[pt.kind] + (pt.shop ? ': ' + pt.shop : '') :
        parts ? (/погашенн/i.test(String(it.description || '')) ? 'погашення' : 'покупка?') : '' };
  });
  const cnt = f => rows.filter(f).length, sal = rows.filter(r => r.salary.length);
  return { rows: rows, sum: {
    n: rows.length, comment: cnt(r => r.comment), name: cnt(r => r.name), code: cnt(r => r.code), iban: cnt(r => r.iban), omcc: cnt(r => r.omcc && r.omcc !== r.mcc),
    salary: sal.length, salaryIn: ['comment', 'description', 'counterName'].map(k => [k, sal.filter(r => r.salary.indexOf(k) >= 0).length]),
    pay: cnt(r => !!r.parts && r.parts !== 'покупка?'), buy: cnt(r => r.parts === 'покупка?'), none: cnt(r => !r.cat || r.cat === CAT_TRANSFER) } };
}

/** HTML вікна: підсумок перевірок, пошук, таблиця. */
function rawHtml_(name, days, data) {
  const s = data.sum, F = { comment: 'коментар (comment)', description: 'опис (description)', counterName: 'отримувач (counterName)' };
  const head = ['Час', 'Сума', 'Опис', 'Коментар', 'MCC', 'Отримувач', 'ЄДРПОУ / ІПН', 'IBAN', 'Шаблон поставить', 'Чому'];
  const td = x => '<td>' + esc_(x) + '</td>';
  const rows = data.rows.map(r => '<tr' + (r.salary.length || r.parts ? ' class="hl"' : !r.cat || r.cat === CAT_TRANSFER ? ' class="no"' : '') + '>' + td(r.t + (r.hold ? ' ⏳' : '')) +
    '<td class="n">' + esc_(money_(r.amt)) + '</td>' + td(r.desc) + td(r.comment) + td(r.mcc + (r.omcc && r.omcc !== r.mcc ? ' (було ' + r.omcc + ')' : '')) + td(r.name) + td(r.code) + td(r.iban) +
    td(r.cat || '(порожньо)') + td(r.why + (r.parts ? ' · частинами: ' + r.parts : '') + (r.salary.length ? ' · «зарплата» в: ' + r.salary.map(k => F[k]).join(', ') : '')) + '</tr>').join('');
  return '<div style="font:13px/1.45 Arial,sans-serif">' +
    '<p><b>' + esc_(name) + '</b>, за ' + days + ' дн.: операцій <b>' + s.n + '</b>; з коментарем ' + s.comment + ', з отримувачем ' + s.name + ', з кодом ЄДРПОУ / ІПН ' + s.code +
    ', з IBAN ' + s.iban + (s.omcc ? ', MCC уточнено банком ' + s.omcc : '') + '. Без категорії чи «переказ»: <b>' + s.none + '</b>.</p>' +
    '<p>🔎 «Заробітна плата»: ' + (s.salary ? '<b>' + s.salary + '</b> опер.; знайдено в: ' + s.salaryIn.filter(x => x[1]).map(x => F[x[0]] + ' ' + x[1]).join(', ') : 'немає за ці дні') + '.<br>' +
    '🔎 Покупки частинами: погашень <b>' + s.pay + '</b>, інших рядків «частинами» (можлива сама покупка) <b>' + s.buy + '</b>.</p>' +
    '<p style="color:#5f6368">Нічого не записано в таблицю. IBAN і ІПН людей з прихованою серединою. Правило за отримувачем: аркуш «' + RULES_SHEET +
    '», «Де шукати: отримувач», текст: назва чи код. Жовті: зарплата чи частинами; рожеві: без категорії.</p>' +
    '<input id="q" placeholder="Пошук: опис, коментар, отримувач…" style="width:100%;padding:6px;margin:4px 0" oninput="f(this.value)">' +
    '<div style="max-height:430px;overflow:auto"><table id="t" style="border-collapse:collapse;width:100%">' +
    '<tr>' + head.map(h => '<th>' + h + '</th>').join('') + '</tr>' + rows + '</table></div></div>' +
    '<style>th,td{border-bottom:1px solid #eee;padding:3px 6px;text-align:left;vertical-align:top}th{position:sticky;top:0;background:#f8f9fa}' +
    'td.n{text-align:right;white-space:nowrap}tr.hl td{background:#fff8e1}tr.no td{background:#fdecea}</style>' +
    '<script>function f(q){q=q.toLowerCase();var r=document.getElementById("t").rows;for(var i=1;i<r.length;i++)' +
    'r[i].style.display=!q||r[i].textContent.toLowerCase().indexOf(q)>=0?"":"none"}</script>';
}

/** Меню «🔎 Сирі дані Mono»: рахунок, кількість днів → вікно. Таблицю не змінює. */
function monoRawNow() {
  if (!menuOk_('🔎 Сирі дані Mono')) return;
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi(), token = PropertiesService.getUserProperties().getProperty('mono_token');
  const T = '🔎 Сирі дані Mono';
  if (!token) { show_(T, ['Mono ще не підключено: «🚀 Підключення», крок 1.']); return; }
  const accs = monoAccs_();
  if (!accs.length) { show_(T, ['Немає збережених рахунків Mono: «🚀 Підключення», крок 1.']); return; }
  let s = accs[0];
  if (accs.length > 1) {
    const a = ui.prompt(T, 'Який рахунок? ' + accs.map((x, i) => (i + 1) + ') ' + x.name).join(', ') + '.\nВведіть номер (порожньо: 1):', ui.ButtonSet.OK_CANCEL);
    if (a.getSelectedButton() !== ui.Button.OK) return;
    s = accs[Math.max(1, Math.min(accs.length, Number(String(a.getResponseText()).trim()) || 1)) - 1];
  }
  const d = ui.prompt(T, 'За скільки днів показати (1–' + RAW_DAYS + ')? Порожньо: ' + RAW_DAYS + '.\nMono дає 1 запит на хвилину: якщо щойно був імпорт, зачекайте до хвилини.', ui.ButtonSet.OK_CANCEL);
  if (d.getSelectedButton() !== ui.Button.OK) return;
  const days = Math.max(1, Math.min(RAW_DAYS, Math.round(Number(String(d.getResponseText()).trim()) || RAW_DAYS)));
  const tz = ss.getSpreadsheetTimeZone(), now = Math.floor(Date.now() / 1000), from = Math.max(now - days * 86400, now - MONO_MAX_SPAN + 60);
  const timeOf = it => Utilities.formatDate(new Date(it.time * 1000), tz, 'dd.MM.yyyy HH:mm');
  let items;
  try { ss.toast('Завантажую з Mono…', T, 10); items = fetchStatement_(token, s.id, from, now); } catch (e) { show_(T, ['⚠️ ' + esc_(e.message)]); return; }
  const acc = accList_(ss).find(x => x.key === 'mono:' + md5Short_(s.id)) || { name: s.name, kind: K_MAIN, who: '' };
  const data = rawRows_(items, readRules_(ss), acc, timeOf, histMap_(readJournal_(ss)));
  const html = HtmlService.createHtmlOutput(rawHtml_(s.name, days, data)).setWidth(1150).setHeight(700);
  try { ui.showModalDialog(html, T); } catch (e) { Logger.log(JSON.stringify(data.sum)); }
}

/* ---------- 🗓 імпорт за період: місяці назад, по 31 добі у фоні ---------- */

/*
 * «N місяців» = з того ж числа N місяців тому, наступного дня, по сьогодні (1 місяць: 30 днів і сьогодні, одне вікно).
 * Mono віддає виписку максимум за 31 добу й 1 запит на хвилину. Тому період ділимо на вікна по 31 добі, новіші спершу:
 * перше вікно одразу з меню, далі розклад mbsPeriod раз на хвилину, по вікну за запуск. Таблицю можна закрити.
 * Дублів немає: ті самі номери операцій, що й у щоденному імпорті (уже в журналі чи в журналі «Банк»: пропуск);
 * ваш рядок з тією ж сумою ±1 день отримує номер операції, як завжди. Категорії, «✅ факт», між своїми, історія: як у щоденному.
 * Не чіпає: «Старт обліку», контрольні точки, час останнього імпорту; старих чернеток не прибирає (вікно без «банк скасував»).
 * Зупиняється: усе пройдено; у журналі закінчились вільні рядки; Mono не віддає (код 4xx); 3 невдачі поспіль (ліміт, VPN, збій); вручну.
 * Стан у налаштуваннях скрипта цього користувача (UserProperties, як і токен): 'period_job', підсумок 'period_last'.
 */
const PERIOD_MAX_MONTHS = 12, PERIOD_DEF_MONTHS = 3, PERIOD_SPAN = 31 * 86400, PERIOD_WAIT_MS = 65 * 1000, PERIOD_FAILS = 3;
const PERIOD_T = '🗓 Імпорт з банку за період';
const PERIOD_SUM = ['seen', 'drafts', 'facts', 'transfers', 'uncategorized', 'matched', 'other', 'pairs', 'plans', 'learned', 'fees', 'lost', 'refunds', 'parts'];

/** Чиста функція: та сама дата n місяців тому ('yyyy-MM-dd'); 31-ше в коротшому місяці → його останній день. */
function monthsBack_(key, n) {
  let y = Number(key.slice(0, 4)), m = Number(key.slice(5, 7)) - n;
  while (m < 1) { m += 12; y--; }
  const ym = y + '-' + pad2_(m);
  return ym + '-' + pad2_(Math.min(Number(key.slice(8, 10)), Number(lastDayOf_(ym).slice(8))));
}

/** Чиста функція: скільки вікон по 31 добі в [from, to] (с). */
function periodWindows_(fromSec, toSec) { return Math.max(1, Math.ceil((toSec - fromSec) / PERIOD_SPAN)); }

/** Помилка Mono, після якої далі не йдемо (токен, код 4xx: найчастіше «так далеко банк не віддає»). Ліміт, VPN, 5xx, мережа: повторити. */
function periodFatal_(msg) { return /токен|кодом 4\d\d/.test(String(msg || '')); }

function periodLater_() {
  removeTriggers_(TRG.period);
  ScriptApp.newTrigger(TRG.period).timeBased().after(PERIOD_WAIT_MS).create();
}

/** Одне вікно: виписка → bankIngest_ (як щоденний імпорт, без контрольних точок і без «банк скасував»). → звіт вікна */
function periodStep_(ss, job) {
  const tz = ss.getSpreadsheetTimeZone(), up = PropertiesService.getUserProperties(), dp = PropertiesService.getDocumentProperties();
  const token = up.getProperty('mono_token'), k = job.key;
  if (!token || !monoAccs_().some(x => 'mono:' + md5Short_(x.id) === k)) throw new Error('Mono відключено або цієї картки більше немає серед підключених');
  const acc = accList_(ss).find(x => x.key === k) || accEnsure_(ss, { key: k, name: job.name, bank: 'Mono' });
  const to = job.next, from = Math.max(job.from, to - PERIOD_SPAN);
  const dateOf = it => Utilities.formatDate(new Date(it.time * 1000), tz, 'yyyy-MM-dd');
  const timeOf = it => Utilities.formatDate(new Date(it.time * 1000), tz, 'dd.MM.yyyy HH:mm');
  const items = fetchStatement_(token, job.id, from, to);
  const exp = expandOps_(items, dp.getProperty('mono_fee_included:' + k) || dp.getProperty('mono_fee_included'));
  if (exp.learned) dp.setProperty('mono_fee_included:' + k, exp.learned);
  const res = bankIngest_(ss, { acc: acc, items: exp.ops, pending: exp.skipped, dateOf: dateOf, timeOf: timeOf, balShift: job.shift, src: 'mono' });
  const r = accReport_(acc, exp.ops, res, dateOf, todayKey_(ss));
  r.fees = exp.fees;
  r.win = [Utilities.formatDate(new Date(from * 1000), tz, 'yyyy-MM-dd'), Utilities.formatDate(new Date(to * 1000), tz, 'yyyy-MM-dd')];
  PERIOD_SUM.forEach(x => { job.rep[x] = (job.rep[x] || 0) + (r[x] || 0); });
  job.done++; job.next = from; job.fails = 0;
  try { sortCore_(ss); } catch (e) { /* порядок наведе розклад */ }
  return r;
}

/**
 * Крок роботи з обробкою помилок і плануванням наступного. → { state: 'more' | 'done' | 'stop', job, r?, err? }
 * Запуск рахується спробою ще до запиту, і наступний запуск планується заздалегідь: якщо Google обірве виконання
 * (6 хв, дуже велика виписка), робота не зависне, а третя така спроба поспіль зупинить її з поясненням.
 */
function periodRun_(ss, job) {
  const up = PropertiesService.getUserProperties();
  job.fails = (job.fails || 0) + 1;
  if (job.fails > PERIOD_FAILS) return { state: 'stop', job: job, err: job.err || 'три спроби поспіль обірвались', last: periodEnd_(ss, job, job.err || 'три спроби поспіль обірвались') };
  up.setProperty('period_job', JSON.stringify(job));
  periodLater_();
  let r = null;
  try { r = periodStep_(ss, job); }
  catch (e) {
    const msg = String(e && e.message || e);
    if (!up.getProperty('period_job')) { removeTriggers_(TRG.period); return { state: 'stop', job: job, err: 'зупинено вручну' }; }
    if (periodFatal_(msg) || job.fails >= PERIOD_FAILS || /відключено/.test(msg)) return { state: 'stop', job: job, err: msg, last: periodEnd_(ss, job, msg) };
    job.err = msg.slice(0, 200);
    up.setProperty('period_job', JSON.stringify(job));
    return { state: 'more', job: job, err: msg };
  }
  if (!up.getProperty('period_job')) { removeTriggers_(TRG.period); return { state: 'stop', job: job, r: r, err: 'зупинено вручну' }; }   // зупинили з меню, поки йшло вікно
  if (r.lost) return { state: 'stop', job: job, r: r, err: 'у журналі закінчились вільні рядки (не вмістилось ' + r.lost + ')', last: periodEnd_(ss, job, 'у журналі закінчились вільні рядки') };
  if (job.next <= job.from) return { state: 'done', job: job, r: r, last: periodEnd_(ss, job, '') };
  delete job.err;
  up.setProperty('period_job', JSON.stringify(job));
  periodLater_();
  return { state: 'more', job: job, r: r };
}

/** Кінець роботи: прибрати розклад і стан, зберегти підсумок; у Telegram, якщо працювало у фоні. → підсумок */
function periodEnd_(ss, job, err) {
  const up = PropertiesService.getUserProperties();
  removeTriggers_(TRG.period);
  up.deleteProperty('period_job');
  const last = { t: new Date().toISOString(), ok: !err, err: String(err || '').slice(0, 200), name: job.name, months: job.months, fromKey: job.fromKey,
    done: job.done, total: job.total, rep: job.rep };
  up.setProperty('period_last', JSON.stringify(last));
  if (job.total > 1 && tgReady_()) { try { tgSend_(periodText_(last)); } catch (e) { /* підсумок є в «Стані таблиці» */ } }
  return last;
}

/** Підсумок одним рядком (без HTML-розмітки, з екрануванням): для «Стану таблиці», вікна й Telegram. */
function periodSummary_(rep) {
  rep = rep || {};
  return 'нових операцій ' + (rep.drafts || 0) + (rep.facts ? ', з них одразу «✅ факт» ' + rep.facts : '') + (rep.transfers ? ', переказів ' + rep.transfers : '') +
    (rep.uncategorized ? ', без категорії ' + rep.uncategorized : '') + (rep.refunds ? ', повернень ' + rep.refunds : '') + (rep.parts ? ', частинами ' + rep.parts : '') + '; уже були в журналі: ' + ((rep.matched || 0) + (rep.other || 0)) +
    '; операцій у виписках: ' + (rep.seen || 0) + (rep.lost ? '; ⚠️ не вмістилось: ' + rep.lost : '');
}

function periodText_(last) {
  return '🗓 <b>Імпорт за період</b> ' + (last.ok ? '✓' : '⚠️') + '\n' + esc_(last.name) + ' · з ' + dd_(last.fromKey) + '.' + last.fromKey.slice(0, 4) +
    ' · вікон ' + last.done + ' з ' + last.total + '\n' + esc_(periodSummary_(last.rep)) + (last.ok ? '' : '\nЗупинено: ' + esc_(last.err)) +
    ((last.rep || {}).transfers || (last.rep || {}).uncategorized ? '\nПерекази й операції без категорії: чернетки в «' + JOURNAL + '», оберіть категорію.' : '');
}

/** Меню: почати (картка, скільки місяців) або, якщо вже триває, показати стан і зупинити. */
function monoPeriodNow() {
  if (!menuOk_(PERIOD_T)) return;
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi(), up = PropertiesService.getUserProperties(), T = PERIOD_T;
  const busy = jsonProp_(up, 'period_job');
  if (busy) {
    const a = ui.alert(T, 'Уже триває: ' + busy.name + ', з ' + dd_(busy.fromKey) + '.' + busy.fromKey.slice(0, 4) + '. Пройдено вікон ' + busy.done + ' з ' + busy.total +
      ' (по одному за хвилину).\n' + periodSummary_(busy.rep) + (busy.err ? '\nОстання спроба: ' + busy.err : '') + '\n\nЗупинити? Що вже завантажилось, лишиться в журналі.', ui.ButtonSet.YES_NO);
    if (a === ui.Button.YES) { periodEnd_(ss, busy, 'зупинено вручну'); show_(T, ['Зупинено. Що встигло завантажитись, лишилось у журналі.']); }
    return;
  }
  const token = up.getProperty('mono_token'), accs = monoAccs_();
  if (!token || !accs.length) {
    show_(T, ['Mono ще не підключено: «🚀 Підключення», крок 1. Інші банки: виписка CSV і «📥 Імпорт виписок з папки».']); return;
  }
  if (!ss.getSheetByName(JOURNAL) || !ss.getSheetByName(BANK_SHEET) || !ss.getSheetByName(RULES_SHEET) || !ss.getSheetByName(ACC_SHEET)) {
    show_(T, ['Немає аркушів журналу чи банку: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».']); return;
  }
  let s = accs[0];
  if (accs.length > 1) {
    const a = ui.prompt(T, 'Яку картку? ' + accs.map((x, i) => (i + 1) + ') ' + x.name).join(', ') + '.\nВведіть номер (порожньо: 1):', ui.ButtonSet.OK_CANCEL);
    if (a.getSelectedButton() !== ui.Button.OK) return;
    s = accs[Math.max(1, Math.min(accs.length, Number(String(a.getResponseText()).trim()) || 1)) - 1];
  }
  const q = ui.prompt(T, 'За скільки місяців назад завантажити (1–' + PERIOD_MAX_MONTHS + ')? Порожньо: ' + PERIOD_DEF_MONTHS + '.\n' +
    'Mono віддає 31 день за раз і 1 запит на хвилину: перший місяць прийде одразу, решта сама у фоні, по місяцю за хвилину. Таблицю можна закрити. ' +
    'Що вже є в журналі, не задвоїться.', ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() !== ui.Button.OK) return;
  const months = Math.max(1, Math.min(PERIOD_MAX_MONTHS, Math.round(Number(String(q.getResponseText()).trim()) || PERIOD_DEF_MONTHS)));
  show_(T, periodBegin_(ss, s, months, addDays_(monthsBack_(todayKey_(ss), months), 1)));
}

/** Почати імпорт за період з дати fromKey по сьогодні: перше вікно одразу, решта у фоні. Також з «🧹 Зачистити дані». → рядки для вікна */
function periodBegin_(ss, s, months, fromKey) {
  const T = PERIOD_T, up = PropertiesService.getUserProperties(), token = up.getProperty('mono_token');
  let info;
  try { ss.toast('Перевіряю картку в Mono…', T, 10); info = monoFetch_(token, '/personal/client-info'); } catch (e) { return ['⚠️ ' + esc_(e.message)]; }
  const found = findCard_(info.accounts, s);
  if (!found.acc) return ['⚠️ ' + esc_(s.name) + ': рахунку більше немає в банку (перевипустили?): «🏦 Підключити Mono» ще раз.'];
  const tz = ss.getSpreadsheetTimeZone(), dp = PropertiesService.getDocumentProperties(), k = 'mono:' + md5Short_(s.id), nowSec = Math.floor(Date.now() / 1000);
  const fromSec = Math.floor(Utilities.parseDate(fromKey + ' 00:00:00', tz, 'yyyy-MM-dd HH:mm:ss').getTime() / 1000);
  const job = { key: k, id: found.acc.id, name: s.name, months: months, fromKey: fromKey, from: fromSec, next: nowSec, total: periodWindows_(fromSec, nowSec), done: 0, fails: 0,
    shift: balanceShift_([], found.acc, dp.getProperty('mono_balance_mode:' + k) || dp.getProperty('mono_balance_mode')).shift, rep: {}, started: new Date().toISOString() };
  up.setProperty('period_job', JSON.stringify(job));
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    periodLater_();
    return ['Інша дія скрипта ще триває. Почну сам за хвилину, у фоні: ' + esc_(s.name) + ', з ' + dd_(fromKey) + '.' + fromKey.slice(0, 4) + ', вікон ' + job.total + '.'];
  }
  let res;
  try { ss.toast('Завантажую перший місяць…', T, 30); res = periodRun_(ss, job); } finally { lock.releaseLock(); }
  return periodReport_(res);
}

/** Вікно після першого кроку з меню. */
function periodReport_(res) {
  const j = res.job, r = res.r, lines = ['Картка: <b>' + esc_(j.name) + '</b>. Період: з <b>' + dd_(j.fromKey) + '.' + j.fromKey.slice(0, 4) + '</b> по сьогодні (' + j.months +
    ' міс.), вікон по 31 дню: ' + j.total + '.'];
  if (r) lines.push('Вікно ' + dd_(r.win[0]) + '–' + dd_(r.win[1]) + ': операцій у виписці ' + r.seen + '; нових <b>' + r.drafts + '</b>' +
    (r.facts ? ' (одразу «✅ факт» ' + r.facts + ')' : '') + '; уже були в журналі: ' + (r.matched + r.other) + '.');
  if (res.state === 'more') {
    const left = j.total - j.done;
    if (res.err) lines.push('⚠️ Mono зараз не відповів: ' + esc_(res.err) + '. Спробую сам за хвилину.');
    lines.push('Решта (' + left + ' ' + (left === 1 ? 'вікно' : left < 5 ? 'вікна' : 'вікон') + ') сама у фоні, по вікну за хвилину: приблизно ' + (left + 1) + ' хв. Таблицю можна закрити.',
      'Підсумок: ' + (tgReady_() ? 'повідомлення в Telegram і ' : '') + '«🩺 Стан таблиці». Зупинити: цей самий пункт меню.');
  } else if (res.state === 'done') {
    lines.push('<b>Готово.</b> Разом: ' + esc_(periodSummary_(j.rep)) + '.');
  } else {
    lines.push('⚠️ Зупинено: ' + esc_(res.err) + '. Що встигло завантажитись, лишилось у журналі. ' + (res.job.done ? 'Разом: ' + esc_(periodSummary_(j.rep)) + '.' : ''));
  }
  if (j.rep.transfers || j.rep.uncategorized) lines.push('Перекази й операції без категорії лишились чернетками в «' + JOURNAL + '»: оберіть категорію.');
  return lines;
}

/** Рядки для «🩺 Стан таблиці». */
function periodStatusLines_(ss) {
  const up = PropertiesService.getUserProperties(), tz = ss.getSpreadsheetTimeZone(), job = jsonProp_(up, 'period_job'), last = jsonProp_(up, 'period_last');
  if (job) return ['Імпорт за період: <b>триває</b>, ' + esc_(job.name) + ', з ' + dd_(job.fromKey) + '.' + job.fromKey.slice(0, 4) + ', вікон ' + job.done + ' з ' + job.total +
    (hasTrigger_(TRG.period) ? '' : ' ⚠️ розкладу немає: запустіть «' + PERIOD_T + '…» ще раз') + (job.err ? '; остання спроба: ' + esc_(job.err) : '')];
  if (!last) return [];
  return ['Імпорт за період (' + Utilities.formatDate(new Date(last.t), tz, 'dd.MM.yyyy HH:mm') + ', ' + esc_(last.name) + ', з ' + dd_(last.fromKey) + '.' + last.fromKey.slice(0, 4) + '): ' +
    (last.ok ? '✓ ' : '⚠️ ' + esc_(last.err) + '; ') + esc_(periodSummary_(last.rep))];
}

function mbsPeriod() { if (quietOk_()) return periodTick_(); }

/** З розкладу: одне вікно; зайнято іншою дією: ще через хвилину. */
function periodTick_() {
  const ss = SpreadsheetApp.getActive(), up = PropertiesService.getUserProperties();
  removeTriggers_(TRG.period);
  if (!up.getProperty('period_job')) return null;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) { periodLater_(); return null; }
  try {
    const job = jsonProp_(up, 'period_job');                         // читаємо під замком: попередній запуск міг щойно зберегти поступ
    return job ? periodRun_(ss, job) : null;
  } finally { lock.releaseLock(); }
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
    try { rep.recat = draftsRecat_(ss); } catch (e) { rep.recat = 0; }             // старі чернетки, для яких тепер є правило
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
    pairs: 0, plans: 0, learned: 0, facts: 0, refunds: 0, parts: 0, partsNew: [], partsClosed: [],
    checkpoints: [], own: null, limit: null, newOps: [], todayOps: { out: 0, inc: 0, n: 0 }, notes: [], accounts: [] };
}

function mergeReport_(rep, r) {
  ['drafts', 'lost', 'transfers', 'uncategorized', 'matched', 'other', 'updated', 'redated', 'changed', 'seen', 'fees', 'skipped', 'vanished', 'pairs', 'plans', 'learned', 'facts', 'refunds', 'parts']
    .forEach(k => { rep[k] = (rep[k] || 0) + (r[k] || 0); });
  ['checkpoints', 'newOps', 'notes', 'accounts', 'partsNew', 'partsClosed'].forEach(k => { rep[k] = (rep[k] || []).concat(r[k] || []); });
  ['out', 'inc', 'n'].forEach(k => { rep.todayOps[k] = round2_(rep.todayOps[k] + ((r.todayOps || {})[k] || 0)); });
  if (r.own !== null && r.own !== undefined) { rep.own = r.own; rep.limit = r.limit; }
  return rep;
}

/** Звіт одного рахунку з результату bankIngest_. */
function accReport_(acc, items, res, dateOf, today) {
  const p = res.plan, r = emptyReport_();
  Object.assign(r, { drafts: res.written + (res.taken || 0), lost: res.lost, transfers: p.stats.transfers, uncategorized: p.stats.uncategorized, matched: p.marks.length,
    other: p.stats.other, updated: p.stats.resized, redated: p.stats.redated, changed: p.stats.changed, seen: items.length, vanished: p.vanish.length, accounts: [acc.name],
    pairs: p.stats.own || 0, plans: p.stats.plans || 0, learned: p.stats.learned || 0, facts: p.stats.facts || 0, refunds: p.stats.refunds || 0, skipped: p.stats.skipped || 0,
    parts: p.stats.parts || 0, partsNew: (res.debts || {}).added || [], partsClosed: (res.debts || {}).closed || [] });
  if ((res.debts || {}).full) r.notes.push('на «' + DEBTS_SHEET + '» немає вільних рядків для нової покупки частинами');
  const op = (d, cat) => ({ key: d.key, time: String(d.bank[4] || '').slice(11, 16), amt: Number(d.bank[2]) || 0, cat: cat || '', desc: String(d.bank[1] || ''), acc: acc.name });
  r.newOps = p.drafts.slice(0, res.written).map(d => op(d, d.c)).concat((p.takes || []).map(t => op(t, t.c)));
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
    const res = bankIngest_(ss, { acc: acc, items: exp.ops, pending: exp.skipped, dateOf: dateOf, timeOf: timeOf, balShift: bs.shift, src: 'mono',
      coverFrom: found.changed ? null : coverFrom, coverTo: found.changed ? null : today });
    const r = accReport_(acc, exp.ops, res, dateOf, today);
    r.fees = exp.fees;
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
  return 'нових операцій ' + r.drafts + (r.transfers ? ' (переказів ' + r.transfers + ')' : '') + (r.uncategorized ? ', без категорії ' + r.uncategorized : '') +
    (r.facts ? ' · одразу «✅ факт» ' + r.facts : '') + (r.plans ? ' · закрито планів ' + r.plans : '') + (r.pairs ? ' · між своїми ' + r.pairs : '') + (r.learned ? ' · за історією ' + r.learned : '') + (r.refunds ? ' · повернень ' + r.refunds : '') + (r.parts ? ' · частинами ' + r.parts : '') + (r.recat ? ' · за правилами ' + r.recat : '') +
    ' · ваших рядків знайдено ' + r.matched + (r.other ? ' · з іншого джерела ' + r.other : '') + (r.lost ? ' · ⚠️ не вмістилось ' + r.lost : '') +
    (r.changed ? ' · ⚠️ розбіжностей сум ' + r.changed : '') + ((r.errors || []).length ? ' · ⚠️ ' + r.errors.join('; ') : '');
}

function importReport_(r) {
  const lines = ['Рахунки: ' + esc_(r.accounts.filter((a, i, x) => x.indexOf(a) === i).join(', ') || 'немає') + '. Операцій у виписках: ' + r.seen + '.',
    '<b>Нових операцій: ' + r.drafts + '</b>' + (r.transfers ? ': з них переказів (червоні): <b>' + r.transfers + '</b>' : '') +
      (r.uncategorized ? ', без категорії: <b>' + r.uncategorized + '</b>' : '') + '.',
    'Ваших рядків знайдено у виписці (без дубля): <b>' + r.matched + '</b>.' + (r.other ? ' Уже були з іншого джерела: <b>' + r.other + '</b>.' : '')];
  if (r.facts) lines.push('Одразу «✅ факт» (категорію визначено): <b>' + r.facts + '</b>. Вимкнути: «' + SETTINGS_SHEET + '», B35.');
  if (r.plans) lines.push('Закрито планів («Регулярні», «Борги»): <b>' + r.plans + '</b>: рядок плану став операцією з банку.');
  if (r.pairs) lines.push('Між своїми рахунками (пара сум на двох рахунках): <b>' + r.pairs + '</b>.');
  if (r.learned) lines.push('Категорія за історією (≥ 3 рази й ≥ 80 % однаково): <b>' + r.learned + '</b>.');
  if (r.refunds) lines.push('↩ Повернень за покупки: <b>' + r.refunds + '</b>: мінусом у категорії покупки (витрати категорії зменшились).');
  if (r.recat) lines.push('🏷 Старі чернетки отримали категорію за правилами: <b>' + r.recat + '</b>.');
  if (r.parts) lines.push('💳 Платежів частинами й за розстрочками: <b>' + r.parts + '</b>' + (modOn_('debts') ? ' (рахуються в «Сплачено» на «' + DEBTS_SHEET + '»).' :
    '. Щоб вести покупки частинами, залишки й ліміт: модуль «🤝 Борги й розстрочки» (меню «🧩 Модулі»).'));
  if ((r.partsNew || []).length) lines.push('💳 Нова покупка частинами на «' + DEBTS_SHEET + '»: <b>' + esc_(r.partsNew.join(', ')) +
    '</b>. Впишіть суму покупки в колонку C: тоді будуть плани платежів, прогрес і ліміт.');
  if ((r.partsClosed || []).length) lines.push('✓ Повністю погашено: <b>' + esc_(r.partsClosed.join(', ')) + '</b> (на «' + DEBTS_SHEET + '» «Закрито»).');
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
  if (r.facts) lines.push('✅ Одразу факт: ' + r.facts);
  if (r.plans) lines.push('✓ Закрито планів: ' + r.plans);
  if (r.refunds) lines.push('↩ Повернень за покупки: ' + r.refunds);
  if ((r.partsNew || []).length) lines.push('💳 Нова покупка частинами: ' + esc_(r.partsNew.join(', ')) + ': впишіть суму на «' + DEBTS_SHEET + '» (C).');
  if ((r.partsClosed || []).length) lines.push('✓ Повністю погашено: ' + esc_(r.partsClosed.join(', ')));
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
  return lines.concat(periodStatusLines_(ss));
}
