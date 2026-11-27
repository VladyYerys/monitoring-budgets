/** 31_Telegram.gs: Telegram: підключення бота, ранковий підсумок о 05:00 з тривогами й самоперевіркою. */

/*
 * Бота людина створює сама в @BotFather. Токен бота і номер чату: лише в налаштуваннях скрипта (UserProperties),
 * не в таблиці. Токен ніколи не потрапляє в текст помилки. Повідомлення: ранковий підсумок (раз на добу),
 * після імпорту з банку, тижневий (понеділок) і місячний (1-го) звіти, збої.
 * Окремий бот для кожної таблиці: так повідомлення не переплутаються.
 * Сім'я: додайте бота в сімейну групу Telegram і напишіть у групі /start: підсумок піде в групу.
 */

const TG_HOUR = 5;
const TG_API = 'https://api.telegram.org/bot';
const TG_MAX = 3900;

function tgCall_(token, method, payload) {
  const hide = s => String(s).split(token).join('•••');
  let res;
  try {
    res = UrlFetchApp.fetch(TG_API + token + '/' + method, { method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload || {}), muteHttpExceptions: true });
  } catch (e) { throw new Error('немає зв\'язку з Telegram: ' + hide(e && e.message || e)); }
  let body = null;
  try { body = JSON.parse(res.getContentText()); } catch (e) { /* не JSON */ }
  if (!body || !body.ok) throw new Error('Telegram відповів ' + res.getResponseCode() + (body && body.description ? ': ' + hide(body.description) : ''));
  return body.result;
}

/** Бот підключено (токен і чат є). */
function tgConnected_() {
  const up = PropertiesService.getUserProperties();
  return !!(up.getProperty('tg_token') && up.getProperty('tg_chat'));
}
/** Можна надсилати: бот підключено і модуль «📨 Telegram» увімкнено. */
function tgReady_() { return tgConnected_() && modOn_('tg'); }

function tgSend_(text) {
  const up = PropertiesService.getUserProperties(), token = up.getProperty('tg_token'), chat = up.getProperty('tg_chat');
  if (!token || !chat) throw new Error('бот не підключено («🚀 Підключення», крок 2)');
  if (!modOn_('tg')) throw new Error('модуль «📨 Telegram» вимкнено (' + MOD_WHERE + ')');
  if (text.length > TG_MAX) text = text.slice(0, TG_MAX - 20).replace(/\n[^\n]*$/, '') + '\n…(скорочено)';
  return tgCall_(token, 'sendMessage', { chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true });
}

/* ---------- меню ---------- */

/** Перший раз: токен від @BotFather; далі: /start у боті, і скрипт запам'ятає чат. */
function connectTelegram() {
  if (!menuOk_('📨 Telegram')) return;
  const ui = SpreadsheetApp.getUi(), up = PropertiesService.getUserProperties();
  let token = up.getProperty('tg_token');
  if (token) {
    const a = ui.alert('📨 Telegram', 'Бот @' + (up.getProperty('tg_bot') || '?') + ' уже збережено. Ввести НОВИЙ токен?\n«Ні»: лише знайти ваш чат (після /start у боті).',
      ui.ButtonSet.YES_NO_CANCEL);
    if (a === ui.Button.CANCEL || a === ui.Button.CLOSE) return;
    if (a === ui.Button.YES) token = null;
  }
  if (!token) {
    const r = ui.prompt('📨 Підключити Telegram-бота', 'Вставте токен бота від @BotFather (вигляд: 123456789:AAE…).\n' +
      'Він збережеться лише в налаштуваннях скрипта, не в таблиці. Нікому його не пересилайте.', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    const t = String(r.getResponseText() || '').trim();
    if (!/^\d{5,15}:[A-Za-z0-9_-]{30,}$/.test(t)) { show_('📨 Telegram', ['Це не схоже на токен бота. Скопіюйте його з @BotFather повністю.']); return; }
    let me;
    try { me = tgCall_(t, 'getMe'); } catch (e) { show_('📨 Telegram', ['⚠️ Токен не підійшов: ' + esc_(e.message)]); return; }
    up.setProperties({ tg_token: t, tg_bot: String(me.username || '') });
    up.deleteProperty('tg_chat');
    token = t;
  }
  const bot = up.getProperty('tg_bot') || '';
  let upd = [];
  try { upd = tgCall_(token, 'getUpdates', { limit: 100, allowed_updates: ['message'] }) || []; } catch (e) { show_('📨 Telegram', ['⚠️ ' + esc_(e.message)]); return; }
  const starts = upd.map(u => u.message).filter(m => m && m.chat && ['private', 'group', 'supergroup'].indexOf(m.chat.type) >= 0 &&
    /^\/start(@\w+)?\b/.test(String(m.text || '')));
  if (!starts.length) {
    show_('📨 Telegram: крок 2', ['Токен збережено ✓ (бот @' + esc_(bot) + ').',
      'Тепер у Telegram знайдіть @' + esc_(bot) + ', натисніть <b>Start</b> (або напишіть /start) і знову запустіть цей пункт меню: токен вводити вже не треба.',
      'Для сімейної групи: додайте бота в групу і напишіть там /start@' + esc_(bot) + '.']);
    return;
  }
  const m = starts[starts.length - 1];
  up.setProperty('tg_chat', String(m.chat.id));
  modSet_(SpreadsheetApp.getActive(), 'tg', true);                 // підключили: модуль увімкнено
  ensureMorningTrigger_();
  try { tgSend_('✅ <b>' + APP_NAME + '</b> підключено. Підсумок: щодня о ' + pad2_(TG_HOUR) + ':00.'); }
  catch (e) { show_('📨 Telegram', ['⚠️ Чат знайдено, але надіслати не вдалося: ' + esc_(e.message)]); return; }
  show_('📨 Telegram підключено', [m.chat.type === 'private' ? 'Чат: особисті повідомлення.' : 'Чат: група «' + esc_(m.chat.title || '') + '».',
    'У Telegram уже має бути повідомлення «підключено».',
    'Підсумок з тривогами й самоперевіркою: щодня о ' + pad2_(TG_HOUR) + ':00. Подивитись зараз: «📨 Telegram зараз → 🌅 Ранковий підсумок».']);
}

function disconnectTelegram() {
  if (!menuOk_('📨 Telegram')) return;
  ['tg_token', 'tg_chat', 'tg_bot'].forEach(k => PropertiesService.getUserProperties().deleteProperty(k));
  removeTriggers_(TRG.morning); removeTriggers_(TRG.morningRetry);
  show_('📨 Telegram вимкнено', ['Токен і чат стерто, повідомлення вимкнено. Самого бота можна видалити в @BotFather (/deletebot).']);
}

function ensureMorningTrigger_() {
  removeTriggers_(TRG.morning);
  ScriptApp.newTrigger(TRG.morning).timeBased().atHour(TG_HOUR).everyDays(1).inTimezone(SpreadsheetApp.getActive().getSpreadsheetTimeZone()).create();
}

function telegramTestNow() {
  if (!menuOk_('📨 Підсумок зараз')) return;
  const r = sendMorning_(SpreadsheetApp.getActive(), true);
  show_('📨 Підсумок зараз', r.ok ? ['Надіслано ✓, гляньте Telegram.', 'Тривог: ' + r.alerts + ', зауважень самоперевірки: ' + r.issues + '.']
    : ['⚠️ Не надіслано: ' + esc_(r.error)]);
}

/* ---------- за розкладом ---------- */

function mbsMorning() {
  if (!quietOk_() || !modOn_('tg')) return;
  const ss = SpreadsheetApp.getActive(), dp = PropertiesService.getDocumentProperties(), today = todayKey_(ss);
  if (dp.getProperty('tg_sent_day') === today) return;
  const r = sendMorning_(ss, false);
  if (r.ok) return;
  if (dp.getProperty('tg_retry_day') !== today) {
    dp.setProperty('tg_retry_day', today);
    removeTriggers_(TRG.morningRetry);
    ScriptApp.newTrigger(TRG.morningRetry).timeBased().after(20 * 60 * 1000).create();
  }
}

function mbsMorningRetry() {
  removeTriggers_(TRG.morningRetry);
  if (!quietOk_()) return;
  const ss = SpreadsheetApp.getActive();
  if (PropertiesService.getDocumentProperties().getProperty('tg_sent_day') === todayKey_(ss)) return;
  const r = sendMorning_(ss, false);
  if (!r.ok) throw new Error('Підсумок у Telegram не надіслано: ' + r.error);
}

function sendMorning_(ss, test) {
  const dp = PropertiesService.getDocumentProperties(), tz = ss.getSpreadsheetTimeZone(), bank = ss.getSheetByName(BANK_SHEET);
  const stamp = () => Utilities.formatDate(new Date(), tz, 'dd.MM HH:mm');
  const status = t => { try { if (bank) bank.getRange('B3').setValue(t); } catch (e) { /* ignore */ } };
  let rep;
  try {
    status('готую підсумок… ' + stamp());
    SpreadsheetApp.flush();
    rep = buildMorning_(ss, { test: test });
    tgSend_(rep.text);
  } catch (e) {
    const msg = String(e && e.message || e);
    status('⚠️ не надіслано ' + stamp() + ': ' + msg);
    return { ok: false, error: msg };
  }
  status('✅ надіслано ' + stamp() + (test ? ' (вручну)' : ''));
  if (!test) {
    dp.setProperty('tg_sent_day', rep.today); dp.setProperty('tg_prev_day', rep.today);
    try { extraReports_(ss); } catch (e) { /* тиждень / місяць: наступного ранку */ }
  }
  return { ok: true, alerts: rep.alerts.length, issues: rep.issues.length, text: rep.text };
}

/* ---------- зміст ---------- */

function mList_(arr, max) { return arr.slice(0, max).join('; ') + (arr.length > max ? '; … ще ' + (arr.length - max) : ''); }
function mLabel_(o) { return dd_(o.key) + ' ' + esc_(o.c || 'без категорії') + ' ' + money0_(o.b); }

/**
 * { text, alerts[], issues[], today }: усе, що піде в ранкове повідомлення. Блоки по черзі дописують
 * рядки підсумку (m.lines), тривоги (m.alerts) і зауваження самоперевірки (m.issues).
 */
function buildMorning_(ss, opts) {
  opts = opts || {};
  const tz = ss.getSpreadsheetTimeZone(), today = opts.today || todayKey_(ss);
  const D = mainByLabel_(ss);
  const m = { ss: ss, tz: tz, opts: opts, today: today, yday: addDays_(today, -1), D: D, rows: readJournal_(ss).filter(o => !o.blank && o.key),
    dp: PropertiesService.getDocumentProperties(), up: PropertiesService.getUserProperties(), lines: [], alerts: [], issues: [], okJobs: [], checkLine: '' };
  m.need = lbl => { const v = D(lbl); if (v === null) m.issues.push('на «' + MAIN_SHEET + '» немає значення «' + lbl + '»'); return v; };
  MORNING_BLOCKS.forEach(block => block(m));
  return { text: morningText_(m), alerts: m.alerts, issues: m.issues, today: today };
}

const MORNING_BLOCKS = [morningSummary_, morningFamily_, morningMineAlerts_, morningGrace_, morningSoon_, morningBudgets_, morningDue_, morningBank_, morningIncome_,
  morningForecast_, m => debtsMorning_(m), m => fopMorning_(m), morningSelfCheck_, m => morningVersion_(m)];

/** Сім'я: учора по учасниках (витрати, факти). */
function morningFamily_(m) {
  if (!familyMode_(m.ss)) return;
  const by = {};
  m.rows.filter(o => o.key === m.yday && Number(o.m) === 0 && o.i === T_SPEND && o.kind !== K_FOP).forEach(o => { const k = o.g || WHO_SHARED; by[k] = (by[k] || 0) + (Number(o.o) || 0); });
  const parts = Object.keys(by).filter(k => by[k] > 0.005).map(k => esc_(k) + ' ' + money0_(by[k]));
  if (parts.length) m.lines.push('👥 Вчора: ' + parts.join(' · '));
}

function morningSummary_(m) {
  const D = m.D, y = m.rows.filter(o => o.key === m.yday);
  const spent = y.filter(o => Number(o.m) === 0 && o.i === T_SPEND && o.kind !== K_FOP).reduce((s, o) => s + (Number(o.o) || 0), 0);
  const fromLimit = y.reduce((s, o) => s + (Number(o.m) === 0 ? Number(o.w) || 0 : 0), 0);
  const yDrafts = y.filter(o => o.e === STATUS_DRAFT && Number(o.bank[2]) < 0 && (o.i === T_SPEND || o.i === ''));
  const yDraftSum = -yDrafts.reduce((s, o) => s + Number(o.bank[2]), 0);
  m.lastOk = m.dp.getProperty('bank_last_ok') || '';
  m.stale = !!m.up.getProperty('mono_token') && hasTrigger_(TRG.imp) && (!m.lastOk ||
    new Date(m.lastOk).getTime() < Utilities.parseDate(m.yday + ' 21:00:00', m.tz, 'yyyy-MM-dd HH:mm:ss').getTime());
  const own = m.need('Доступно моїх сьогодні'), perDay = D('Можна на день'), daysLeft = D('Днів до кінця місяця');
  m.credit = modOn_('credit');
  m.own = own; m.limit = m.credit ? D('Кредитний ліміт') : 0; m.margin = Number(D('Межа «близько до ліміту»') || 0);
  const drafts = m.rows.filter(o => o.e === STATUS_DRAFT);
  m.staleDrafts = drafts.filter(o => o.key < addDays_(m.today, -2) && o.key >= addDays_(m.today, -7));
  m.old = drafts.filter(o => o.key < addDays_(m.today, -7));
  const num = x => (x === null ? 'немає даних' : money0_(x));
  const wd = WEEKDAYS[dowOf_(m.today)];
  const asNow = D('Як зараз: на картці на кінець місяця'), pace = D('Витрачаєте ≈ на день');
  m.lines.push((m.opts.test ? '🧪 <i>Тест (надіслано вручну)</i>\n' : '') + '🌅 <b>Підсумок · ' + wd + ' ' + dd_(m.today) + '</b>',
    'Вчора витрачено: <b>' + money0_(spent + yDraftSum) + '</b>' + (yDrafts.length ? ' (з них у чернетках ' + yDrafts.length + ' на ' + money0_(yDraftSum) + ', підтвердіть)' : '') +
      (fromLimit > 0.005 ? ', з ліміту ' + money0_(fromLimit) : '') + (m.stale ? ', ⚠️ <b>неповні дані</b> (див. тривоги)' : ''),
    '💚 Доступно «моїх»: <b>' + num(own) + '</b>' + (perDay !== null && Number(own) > 0 ? ' · ≈ <b>' + money0_(perDay) + '</b> на день' + (daysLeft ? ' (' + daysLeft + ' дн.)' : '') : ''),
    '📅 На картці на кінець місяця за планом: <b>' + num(D('Прогноз на кінець місяця (за планом)')) + '</b>',
    asNow !== null && pace !== null ? '🧭 Якщо тратити як зараз (≈ ' + money0_(pace) + ' на день): на кінець місяця ' + money0_(asNow) : '',
    D('Прогноз на кінець наступного місяця (за планом)') !== null ? '📆 Кінець наступного місяця за планом: ' + money0_(D('Прогноз на кінець наступного місяця (за планом)')) : '',
    m.limit ? '💳 Борг за лімітом: <b>' + num(D('Борг за лімітом зараз')) + '</b> (ліміт ' + money0_(m.limit) + ')' : '💳 На картці: <b>' + num(D('На картці зараз')) + '</b>',
    '🏦 Чернетки чекають підтвердження: <b>' + drafts.length + '</b>' + (m.old.length ? ' (з них старших за тиждень ' + m.old.length + ')' : ''));
}

function morningMineAlerts_(m) {
  if (m.own !== null && Number(m.own) < 0) m.alerts.push('▲▲ <b>«Мої» закінчились</b>: мінус ' + money0_(-m.own) + (m.credit ? ': далі витрати йдуть з ліміту.' :
    ': витрачено більше, ніж надійшло цього бюджетного місяця.'));
  if (!m.credit) return;
  const noReason = m.rows.filter(o => Number(o.m) === 0 && Number(o.w) > 0.005 && !o.h);
  if (noReason.length) m.alerts.push('❓ <b>З ліміту без причини: ' + noReason.length + '</b>: ' +
    mList_(noReason.map(o => dd_(o.key) + ' ' + esc_(o.c) + ' ' + money0_(o.w)), 5) + '. Вкажіть причину в колонці H.');
}

function morningSoon_(m) {
  const soon = m.rows.filter(o => (isPlan_(o.e) || Number(o.m) === 1) && (o.key === m.today || o.key === addDays_(m.today, 1)));
  if (soon.length) m.lines.push('🗓 Сьогодні/завтра за планом: ' + mList_(soon.map(mLabel_), 6));
}

function morningBudgets_(m) {
  const bud = budgetLines_(m.ss, m.rows, null);
  bud.filter(x => x.over).forEach(x => m.alerts.push('💰 <b>Бюджет перевищено</b>: ' + x.text.replace(/^💰 /, '') + '.'));
  bud.filter(x => x.fast && !x.over).forEach(x => m.alerts.push('💰 <b>Швидше норми</b>: ' + x.text.replace(/^💰 /, '') + '.'));
}

function morningDue_(m) {
  const due = m.rows.filter(o => isPlan_(o.e) && o.key < m.today);
  if (due.length) m.alerts.push('🗓 <b>План настав, підтвердіть: ' + due.length + '</b>: ' + mList_(due.map(mLabel_), 5) + '.');
}

function morningBank_(m) {
  if (m.stale) m.alerts.push('🏦 <b>Імпорт з банку не відпрацював</b>, останній вдалий ' +
    (m.lastOk ? Utilities.formatDate(new Date(m.lastOk), m.tz, 'dd.MM HH:mm') : 'ніколи') + ': цифри за вчора й «мої» можуть бути неповні. «🏦 Імпорт з банку зараз»; причина: «🩺 Стан таблиці».');
  if (m.staleDrafts.length) m.alerts.push('🏦 <b>Чернетки з банку старші за 2 дні: ' + m.staleDrafts.length + '</b> (найстаріша ' +
    dd_(m.staleDrafts.map(o => o.key).sort()[0]) + '): підтвердіть («✅ факт») або видаліть.');
  const transfers = m.rows.filter(o => o.c === CAT_TRANSFER && o.e !== STATUS_DRAFT), noCat = m.rows.filter(o => o.c === '' && o.e !== STATUS_DRAFT);
  if (transfers.length || noCat.length) m.alerts.push('↔ <b>Червоні рядки</b>' + (transfers.length ? ': переказ (уточніть): ' + transfers.length : '') +
    (noCat.length ? (transfers.length ? '; ' : ': ') + 'без категорії: ' + noCat.length : '') + '. Поки червоні, не в статистиці.');
}

/** Грейс Mono: скільки внести до строку; тривога за 5 днів і ближче. */
function morningGrace_(m) {
  if (!m.credit || !m.limit) return;
  const due = Number(m.D('Грейс: внести до строку') || 0), dl = m.D('Грейс: строк');
  if (!(due > 0.5) || dl === null) return;
  const key = dateKey_(dl, m.tz), left = dayDiff_(key, m.today), fee = Number(m.D('Грейс: комісія, якщо все переказами') || 0);
  const est = Number(m.D('Грейс: переказами (оцінка)') || 0), feeEst = Number(m.D('Грейс: комісія (оцінка)') || 0);
  const line = '💳 Грейс до ' + dd_(key) + ': внести <b>' + money0_(due) + '</b>' + (est > 0.5 ? '; переказами ≈ ' + money0_(est) + ' (комісія ≈ ' + money0_(feeEst) + ')' :
    ' (планові надходження покривають)') + (fee > 0.5 ? '; якщо все переказами: комісія ' + money0_(fee) : '');
  if (left >= 0 && left <= 5) m.alerts.push('⏳ <b>Грейс закінчується</b> через ' + left + ' дн.: ' + line.replace(/^💳 /, ''));
  else m.lines.push(line);
}

/** Прийшли «мої» надходження (зарплата, дохід ФОП): нагадати відкласти на погашення ліміту до орієнтира. */
function morningIncome_(m) {
  if (!m.credit) return;
  const goal = Number(m.D('Орієнтир погашення на місяць') || 0), saved = Number(m.D('Відкладено на погашення (бюджетний місяць)') || 0);
  if (!(goal > 0) || saved >= goal) return;
  const prev = m.dp.getProperty('tg_prev_day'), from = prev && prev < m.today && prev >= addDays_(m.today, -7) ? prev : m.yday;
  const inc = m.rows.filter(o => o.key >= from && o.key < m.today && o.r === MINE_IN && !isPlan_(o.e) && Number(o.b) >= 1000);
  if (!inc.length) return;
  m.alerts.push('💰 <b>Прийшли надходження</b> ' + money0_(inc.reduce((s, o) => s + Number(o.b), 0)) + ' → відкладіть <b>' + money0_(goal - saved) +
    '</b> на погашення ліміту (цього місяця вже ' + money0_(saved) + ' з ' + money0_(goal) + ').');
}

function morningForecast_(m) {
  const low = m.D('Найнижча точка «як зараз»') !== null ? m.D('Найнижча точка «як зараз»') : m.D('Найнижча точка за планом');
  const lowD = m.D('Найнижча точка «як зараз»') !== null ? m.D('Дата найнижчої «як зараз»') : m.D('Дата найнижчої точки');
  if (low === null) return;
  const avail = Number(low) + Number(m.limit || 0);
  if (avail < m.margin) m.alerts.push('📉 <b>Прогноз</b>: ' + (lowD ? dd_(dateKey_(lowD, m.tz)) + ' ' : '') +
    (avail < 0 ? (m.limit ? 'вийде за ліміт на ' : 'картка піде в мінус на ') + money0_(-avail) : 'доступно лише ' + money0_(avail)) + ' (межа ' + money0_(m.margin) + ').');
}

/** Самоперевірка: «сьогодні» в таблиці, розклади (імпорт, копія, порядок), «Перевірка», календар. */
function morningSelfCheck_(m) {
  const seen = m.D('Сьогодні');
  if (seen !== null && dateKey_(seen, m.tz) !== m.today) m.issues.push('таблиця ще рахує «сьогодні» = ' + esc_(dateKey_(seen, m.tz)) + ': цифри можуть бути на вчора');
  const ago = b => Utilities.formatDate(new Date(b.t), m.tz, 'dd.MM HH:mm');
  const job = (fn, name, title, maxH) => {
    if (!hasTrigger_(fn)) return title + ': розклад вимкнено';
    const b = readBeat_(name);
    if (!b) return null;
    if (!b.ok) return title + ': помилка ' + ago(b) + ': ' + esc_(b.msg);
    if (Date.now() - new Date(b.t).getTime() > maxH * 3600 * 1000) return title + ': не відпрацював, останній раз ' + ago(b);
    return null;
  };
  [[TRG.imp, 'import', 'Імпорт з банку', 14], [TRG.backup, 'backup', 'Нічна копія', 26], [TRG.tick, 'sort', 'Порядок журналу', 2]].forEach(j => {
    if (j[0] === TRG.imp && !m.up.getProperty('mono_token')) return;
    if (j[0] === TRG.backup && !modOn_('backup')) return;
    const r = job(j[0], j[1], j[2], j[3]);
    if (r) m.issues.push(r); else m.okJobs.push(j[2]);
  });
  try { calIssues_(m.ss).forEach(x => m.issues.push(x)); } catch (e) { /* ignore */ }
  const cs = checkState_(m.ss);
  if (cs.bad.length) m.issues.push('«' + CHECK_SHEET + '»: потребують уваги ' + cs.bad.length + ': ' + mList_(cs.bad.map(esc_), 5));
  else if (!cs.none) m.checkLine = '«' + CHECK_SHEET + '» ' + cs.ok + ' OK';
}

function morningText_(m) {
  const tail = [];
  if (m.alerts.length) tail.push('', '⚠️ <b>Тривоги</b>');
  m.alerts.forEach(a => tail.push(a));
  tail.push('', '🔧 <b>Самоперевірка</b>: ' + (m.issues.length ? 'є зауваження' : 'усе OK') +
    (m.okJobs.length || m.checkLine ? ' (' + m.okJobs.concat(m.checkLine ? [m.checkLine] : []).map(x => x + ' ✓').join(', ') + ')' : ''));
  m.issues.forEach(x => tail.push('⚠️ ' + x));
  if (m.old.length) tail.push('', 'ℹ️ Старіші чернетки (понад тиждень): ' + m.old.length + ', без тривоги.');
  return m.lines.filter(x => x !== '').concat(tail).join('\n');
}

function tgStatusLines_(ss) {
  const up = PropertiesService.getUserProperties(), dp = PropertiesService.getDocumentProperties();
  if (!modOn_('tg')) return ['Модуль вимкнено (' + MOD_WHERE + ')' + (tgConnected_() ? ' (бот @' + esc_(up.getProperty('tg_bot') || '?') + ' лишився підключеним)' : '') + '.'];
  return [tgReady_() ? 'Бот @' + esc_(up.getProperty('tg_bot') || '?') + ' підключено ✓; ранковий підсумок: ' + (hasTrigger_(TRG.morning) ? 'щодня о ' + pad2_(TG_HOUR) + ':00' : '<b>вимкнено</b>') +
    (dp.getProperty('tg_sent_day') ? ', останній: ' + dp.getProperty('tg_sent_day') : '') : 'Не підключено: «🚀 Підключення», крок 2.'];
}
