/** 33_Calendar.gs: Плани з журналу → окремий Google Календар з нагадуваннями (за бажанням). */

/*
 * Окремий календар «Monitoring Budgets: платежі» (не основний: його легко сховати чи видалити).
 * Події на цілий день: рядки-плани журналу зі списанням (статус «план», «🔁 план регулярний» або майбутня дата без статусу),
 * від сьогодні. Нагадування: за 5 днів і напередодні о 09:00. Змінилась сума чи дата: подія оновлюється.
 * План став «✅ факт»: «✅» у назві, без нагадувань. Рядок прибрали: майбутня подія видаляється.
 * Скрипт чіпає лише події зі своєю позначкою (тег mbs) у своєму календарі: ваші власні події та інші календарі не чіпає.
 */

const CAL_NAME = 'Monitoring Budgets: платежі';
const CAL_TAG = 'mbs';
const CAL_REMIND = [6660, 900];                       // хвилин до початку дня: за 5 днів о 09:00 і напередодні о 09:00
const CAL_BACK_DAYS = 45, CAL_AHEAD_DAYS = 400, CAL_RETRY_MS = 3600 * 1000;

/** Модуль «📆 Календар платежів» увімкнено («🚀 Підключення», крок 5). */
function calOn_() { return modOn_('cal'); }

/** Лише календар, який створив сам шаблон (id збережено). Чужий календар з такою ж назвою не береться. */
function calGet_(create) {
  const dp = PropertiesService.getDocumentProperties(), id = dp.getProperty('cal_id');
  let cal = null;
  if (id) { try { cal = CalendarApp.getCalendarById(id); } catch (e) { cal = null; } }
  if (!cal && create) {
    cal = CalendarApp.createCalendar(CAL_NAME, { summary: 'Плани з таблиці ' + APP_NAME + '. Оновлюється сама: правки тут перезапишуться.' });
    dp.setProperty('cal_id', cal.getId());
  }
  return cal;
}

function calTitle_(o) { return '💳 ' + o.c + ' · ' + money0_(o.b); }

/** Що має бути в календарі. → { want: {ключ: {date, title, desc}}, paid: {ключ: дата} }. Ключ: P|черговість запису рядка. */
function calWanted_(ss, today) {
  const want = {}, paid = {}, horizon = addDays_(today, CAL_AHEAD_DAYS);
  readJournal_(ss).forEach(o => {
    if (o.blank || !o.key || o.ord === '' || o.ord === null || o.ord === undefined) return;
    const k = 'P|' + o.ord, planned = isPlan_(o.e) || (o.e === '' && o.key > today);
    if (!planned) { if (o.e === STATUS_FACT) paid[k] = o.key; return; }
    if (o.key < today || o.key > horizon || TYPES_IN.indexOf(o.i) >= 0 || o.i === T_RESERVE || !(Number(o.b) > 0)) return;
    want[k] = { date: o.key, title: calTitle_(o), desc: 'Сума: ' + money0_(o.b) + '\nКатегорія: ' + o.c + (o.d ? '\nКоментар: ' + o.d : '') +
      '\n\nЗ таблиці «' + APP_NAME + '», оновлюється сама: правки тут перезапишуться.' };
  });
  return { want: want, paid: paid };
}

function calSync_(ss) {
  const t0 = Date.now(), dp = PropertiesService.getDocumentProperties(), today = todayKey_(ss);
  const cal = calGet_(true), w = calWanted_(ss, today), stz = Session.getScriptTimeZone();
  const day = k => Utilities.parseDate(k + ' 12:00:00', stz, 'yyyy-MM-dd HH:mm:ss');
  const st = { created: 0, updated: 0, paid: 0, removed: 0, open: Object.keys(w.want).length }, seen = {};
  cal.getEvents(day(addDays_(today, -CAL_BACK_DAYS)), day(addDays_(today, CAL_AHEAD_DAYS))).forEach(ev => {
    const k = String(ev.getTag(CAL_TAG) || '');
    if (!k) return;                                                     // ваша власна подія
    if (seen[k]) { ev.deleteEvent(); st.removed++; return; }
    seen[k] = 1;
    const evDay = Utilities.formatDate(ev.getAllDayStartDate(), stz, 'yyyy-MM-dd');
    if (w.paid[k] !== undefined) {
      if (String(ev.getTitle()).indexOf('✅') !== 0) { ev.setTitle('✅ ' + ev.getTitle()); ev.removeAllReminders(); st.paid++; }
      return;
    }
    const x = w.want[k];
    if (!x) { if (evDay >= today) { ev.deleteEvent(); st.removed++; } return; }
    let ch = false;
    if (ev.getTitle() !== x.title) { ev.setTitle(x.title); ch = true; }
    if (evDay !== x.date) { ev.setAllDayDate(day(x.date)); ch = true; }
    if (ev.getDescription() !== x.desc) { ev.setDescription(x.desc); ch = true; }
    if (ch) st.updated++;
  });
  Object.keys(w.want).sort().forEach(k => {
    if (seen[k]) return;
    const x = w.want[k], ev = cal.createAllDayEvent(x.title, day(x.date), { description: x.desc });
    ev.setTag(CAL_TAG, k);
    ev.removeAllReminders();
    CAL_REMIND.forEach(m => ev.addPopupReminder(m));
    st.created++;
  });
  dp.setProperties({ cal_day: today, cal_t: String(t0), cal_stat: JSON.stringify(Object.assign({ t: Date.now() }, st)) });
  dp.deleteProperty('cal_err');
  return st;
}

/** З mbsTick: після правок журналу або раз на день. */
function calIfDue_(ss) {
  if (!calOn_()) return null;
  const dp = PropertiesService.getDocumentProperties(), err = jsonProp_(dp, 'cal_err');
  if (err && Date.now() - Number(err.t || 0) < CAL_RETRY_MS) return null;
  const dirty = Number(dp.getProperty('lastEdit') || 0) >= Number(dp.getProperty('cal_t') || 0);
  if (!dirty && dp.getProperty('cal_day') === todayKey_(ss)) return null;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return null;
  try { return calSync_(ss); }
  catch (e) { dp.setProperty('cal_err', JSON.stringify({ t: Date.now(), msg: String(e && e.message || e).slice(0, 300) })); return null; }
  finally { lock.releaseLock(); }
}

function calStatText_(st) {
  return 'відкритих планів ' + st.open + (st.created ? ', нових подій ' + st.created : '') + (st.updated ? ', оновлено ' + st.updated : '') +
    (st.paid ? ', сплачено ' + st.paid : '') + (st.removed ? ', прибрано ' + st.removed : '');
}

function calendarOn() {
  if (!menuOk_('📆 Календар платежів')) return;
  const ss = SpreadsheetApp.getActive();
  modSet_(ss, 'cal', true);
  let st;
  try { st = calSync_(ss); } catch (e) { show_('📆 Календар платежів', ['⚠️ Не вдалося: ' + esc_(e && e.message || e)]); return; }
  show_('📆 Календар платежів', ['Календар «' + esc_(CAL_NAME) + '» ✓: ' + calStatText_(st) + '.',
    'Нагадування: за 5 днів і напередодні о 09:00. Сплачене отримує «✅» у назві.',
    'Оновлюється сам: після правок журналу і раз на день.',
    'На iPhone: якщо календаря не видно, увімкніть його в застосунку Google Календар (або на calendar.google.com/calendar/syncselect).']);
}

function calendarOff() {
  if (!menuOk_('📆 Календар платежів')) return;
  const ui = SpreadsheetApp.getUi(), dp = PropertiesService.getDocumentProperties();
  const a = ui.alert('📆 Вимкнути календар платежів', 'Скрипт перестане оновлювати календар «' + CAL_NAME + '».\n«Так»: ще й видалити цей календар разом з подіями.\n«Ні»: лишити як є.',
    ui.ButtonSet.YES_NO_CANCEL);
  if (a === ui.Button.CANCEL || a === ui.Button.CLOSE) return;
  modSet_(SpreadsheetApp.getActive(), 'cal', false);
  let gone = false;
  if (a === ui.Button.YES) {
    try { const cal = calGet_(false); if (cal) { cal.deleteCalendar(); gone = true; } dp.deleteProperty('cal_id'); }
    catch (e) { show_('📆 Календар вимкнено', ['Оновлення вимкнено.', '⚠️ Сам календар видалити не вдалося: ' + esc_(e && e.message || e)]); return; }
  }
  show_('📆 Календар вимкнено', ['Скрипт більше не оновлює календар.' + (gone ? ' Календар видалено.' : ' Календар лишився: його можна сховати чи видалити в Google Календарі.')]);
}

function calLines_(ss) {
  const dp = PropertiesService.getDocumentProperties(), tz = ss.getSpreadsheetTimeZone(), st = jsonProp_(dp, 'cal_stat'), err = jsonProp_(dp, 'cal_err');
  const fmt = t => Utilities.formatDate(new Date(Number(t)), tz, 'dd.MM HH:mm');
  if (!calOn_()) return ['Вимкнено. Увімкнути: модуль «📆 Календар платежів», ' + MOD_WHERE + '.'];
  const out = ['Увімкнено: «' + esc_(CAL_NAME) + '»' + (st ? ', оновлено ' + fmt(st.t) + ': ' + calStatText_(st) : '') + '.'];
  if (err) out.push('⚠️ Помилка ' + fmt(err.t) + ': ' + esc_(err.msg));
  return out;
}

function calIssues_(ss) {
  if (!calOn_()) return [];
  const dp = PropertiesService.getDocumentProperties(), err = jsonProp_(dp, 'cal_err'), day = dp.getProperty('cal_day');
  if (err) return ['календар платежів: помилка: ' + esc_(err.msg)];
  if (day && day < addDays_(todayKey_(ss), -2)) return ['календар платежів не оновлювався з ' + dd_(day)];
  return [];
}
