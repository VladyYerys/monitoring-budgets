/** 34_Privat.gs: ПриватБанк «Автоклієнт» (Приват24 для бізнесу): виписка рахунків ФОП. */

/*
 * Токен людина створює сама: «Приват24 для бізнесу» → Каталог послуг → Інтеграція (Автоклієнт) → підключити → токен.
 * Він зберігається лише в налаштуваннях скрипта в її Google-акаунті (UserProperties). Код лише читає виписку
 * (GET /api/statements/transactions); платежів не створює. Список рахунків банк через API не дає: IBAN вводите самі.
 * Беремо лише проведені гривневі операції. Номер операції: «pb:» + ID банку. Рахунок за замовчуванням: «рахунок ФОП».
 */

const PB_API = 'https://acp.privatbank.ua/api';
const PB_FIRST_DAYS = 7;                                  // перший імпорт: тиждень назад

function privatFetch_(token, path) {
  const hide = s => String(s).split(token).join('•••');
  let r;
  try {
    r = UrlFetchApp.fetch(PB_API + path, { method: 'get', muteHttpExceptions: true,
      headers: { 'User-Agent': 'monitoring-budgets', token: token, 'Content-Type': 'application/json;charset=utf8' } });
  } catch (e) { throw new Error('немає зв\'язку з ПриватБанком: ' + hide(e && e.message || e)); }
  const code = r.getResponseCode();
  let body = null;
  try { body = JSON.parse(r.getContentText()); } catch (e) { body = null; }
  if (code === 401 || code === 403) throw new Error('ПриватБанк не прийняв токен (код ' + code + '). Створіть новий в «Приват24 для бізнесу» → Автоклієнт.');
  if (code !== 200 || !body || String(body.status).toUpperCase() !== 'SUCCESS')
    throw new Error('ПриватБанк відповів кодом ' + code + (body && (body.message || body.description) ? ': ' + hide(body.message || body.description) : ''));
  return body;
}

function privatAccs_() { return jsonProp_(PropertiesService.getUserProperties(), 'pb_accs') || []; }

function connectPrivat() {
  if (!menuOk_('🏦 ПриватБанк')) return;
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  const a = ui.prompt('🏦 Підключити ПриватБанк (ФОП)', 'Вставте токен «Автоклієнта» з «Приват24 для бізнесу».\n' +
    'Він зберігається лише в налаштуваннях скрипта у вашому Google-акаунті. Код лише читає виписку.', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const token = String(a.getResponseText() || '').trim();
  if (token.length < 20 || /\s/.test(token)) { show_('🏦 ПриватБанк', ['Це не схоже на токен «Автоклієнта». Скопіюйте його повністю.']); return; }
  const b = ui.prompt('🏦 Рахунки ФОП', 'IBAN рахунків ФОП у ПриватБанку через кому (UA + 27 цифр).', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  const ibans = String(b.getResponseText() || '').toUpperCase().replace(/\s/g, '').split(',').filter(Boolean);
  if (!ibans.length || ibans.some(x => !/^UA\d{27}$/.test(x))) { show_('🏦 ПриватБанк', ['IBAN має вигляд UA і 27 цифр. Спробуйте ще раз.']); return; }
  const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd-MM-yyyy');
  try { ibans.forEach(iban => privatFetch_(token, '/statements/transactions?acc=' + iban + '&startDate=' + today + '&limit=20')); }
  catch (e) { show_('🏦 ПриватБанк', ['⚠️ ' + esc_(e.message)]); return; }
  const who = askWho_('🏦 ПриватБанк', 'це рахунки ФОП');
  if (who === null) return;
  const saved = ibans.map(iban => {
    const acc = accEnsure_(ss, { key: 'pb:' + md5Short_(iban), name: 'ПриватБанк ФОП *' + iban.slice(-4), bank: 'ПриватБанк', who: who, kind: K_FOP });
    return { iban: iban, name: acc.name };
  });
  PropertiesService.getUserProperties().setProperties({ pb_token: token, pb_accs: JSON.stringify(saved) });
  show_('🏦 ПриватБанк підключено', ['Рахунки: <b>' + saved.map(s => esc_(s.name)).join(', ') + '</b> (вид «' + K_FOP + '»: ідуть на аркуш «' + FOP_SHEET + '»).',
    'Далі: «🏦 Імпорт з банку зараз» (перший імпорт: ' + PB_FIRST_DAYS + ' днів). За розкладом: разом з Mono.']);
}

function disconnectPrivat() {
  if (!menuOk_('🏦 ПриватБанк')) return;
  ['pb_token', 'pb_accs'].forEach(k => PropertiesService.getUserProperties().deleteProperty(k));
  if (!PropertiesService.getUserProperties().getProperty('mono_token')) { removeTriggers_(TRG.imp); removeTriggers_(TRG.impRetry); }
  show_('🏦 ПриватБанк відключено', ['Токен стерто. Відкликати його можна і в «Приват24 для бізнесу» → Автоклієнт.']);
}

/** Чиста функція: операція «Автоклієнта» → загальний вигляд. null: не проведена або не в гривні. */
function privatItem_(t, tz) {
  if (String(t.PR_PR || 'r').toLowerCase() !== 'r') return null;
  if (t.CCY && ['UAH', '980'].indexOf(String(t.CCY).toUpperCase()) < 0) return null;
  const sum = Math.round(Number(String(t.SUM === undefined ? t.SUM_E : t.SUM).replace(',', '.').replace(/\s/g, '')) * 100);
  if (!isFinite(sum) || !sum) return null;
  const d = String(t.DAT_OD || '').match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  if (!d) return null;
  const tm = (String(t.TIM_P || '').match(/^(\d{2}):(\d{2})/) || [null, '12', '00']);
  const time = Math.floor(Utilities.parseDate(d[3] + '-' + d[2] + '-' + d[1] + ' ' + tm[1] + ':' + tm[2] + ':00', tz, 'yyyy-MM-dd HH:mm:ss').getTime() / 1000);
  return { id: 'pb:' + (t.ID || (String(t.REF || '') + '.' + String(t.REFN || ''))), time: time, amount: String(t.TRANTYPE).toUpperCase() === 'D' ? -Math.abs(sum) : Math.abs(sum),
    description: String(t.OSND || ''), comment: String(t.AUT_CNTR_NAM || ''), mcc: 0, hold: false };
}

/** Усі рахунки ПриватБанку цього користувача (сторінки по 100, поки банк каже, що є ще). */
function importPrivat_(ss) {
  const tz = ss.getSpreadsheetTimeZone(), up = PropertiesService.getUserProperties(), dp = PropertiesService.getDocumentProperties();
  const token = up.getProperty('pb_token'), today = todayKey_(ss), rep = emptyReport_();
  const dateOf = it => Utilities.formatDate(new Date(it.time * 1000), tz, 'yyyy-MM-dd');
  const timeOf = it => Utilities.formatDate(new Date(it.time * 1000), tz, 'dd.MM.yyyy HH:mm');
  privatAccs_().forEach(s => {
    const k = 'pb:' + md5Short_(s.iban), acc = accList_(ss).find(x => x.key === k) || accEnsure_(ss, { key: k, name: s.name, bank: 'ПриватБанк', kind: K_FOP });
    const last = dp.getProperty('bank_last_day:' + k), from = last ? addDays_(last, -1) : addDays_(today, -PB_FIRST_DAYS);
    const ddmm = key => key.slice(8, 10) + '-' + key.slice(5, 7) + '-' + key.slice(0, 4);
    const items = [];
    let follow = '';
    for (let page = 0; page < 30; page++) {
      const body = privatFetch_(token, '/statements/transactions?acc=' + s.iban + '&startDate=' + ddmm(from) + '&endDate=' + ddmm(today) + '&limit=100' +
        (follow ? '&followId=' + encodeURIComponent(follow) : ''));
      (body.transactions || []).forEach(t => { const it = privatItem_(t, tz); if (it) items.push(it); });
      if (!body.exist_next_page || !body.next_page_id) break;
      follow = body.next_page_id;
    }
    const res = bankIngest_(ss, { acc: acc, items: items, dateOf: dateOf, timeOf: timeOf, src: 'pb' });
    mergeReport_(rep, accReport_(acc, items, res, dateOf, today));
    dp.setProperty('bank_last_day:' + k, today);
  });
  return rep;
}
