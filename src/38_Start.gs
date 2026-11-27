/** 38_Start.gs: аркуш «Старт»: кроки з номерами кроків панелі «🚀 Підключення», демо-місяць. Підключення: 44_Connect.gs. */

/*
 * У копії шаблону розклади й дозволи не переносяться, тож усе починається з меню «📊 Monitoring Budgets → 🚀 Підключення»:
 * скрипт добудовує, чого бракує, і відкриває бічну панель з п'ятьма кроками (44_Connect.gs).
 * «Старт» оновлює скрипт (раз на день, після «Готово», галочок модулів і з меню). Рядок: «✅ 1» зроблено, «⬜ 2» ще ні;
 * число: крок у панелі, «·»: поза панеллю. Непотрібне (вимкнений модуль, «я сам» без сім'ї) не показується.
 * Модулі: «🚀 Підключення», крок 5 (до версії 12 блок «🧩 Модулі» стояв тут, рядки 24…).
 * Демо-місяць: вигадані операції з позначкою «[демо]» у коментарі, щоб одразу побачити «Charts» і «charts_mob»; «🧹 Очистити демо» прибирає все.
 */

const START_FIRST = 6, START_ROWS = 30;                                  // кроки з рядка 6; місце під кроки й низ: 30 рядків
const DEMO_TAG = '[демо] ';
const P_STEP = 'крок ';                                                 // «Де»: крок панелі «🚀 Підключення»

function buildStart_(ss) {
  const sh = ensureSize_(sheetOf_(ss, START_SHEET), START_FIRST + START_ROWS, 4);
  sh.getRange(1, 1).setValue('🚀 Старт').setFontWeight('bold').setFontSize(14);
  if (!startHasModules_(ss) && ss.getSheetByName(SETTINGS_SHEET)) modBlockBuild_(ss);
  startRefresh_(ss);
  return sh;
}

/** Кроки: { ok: true (✅) / false (⬜) / null (не потрібно: не показується), n: крок панелі (0: поза панеллю), what, where }. */
function startSteps_(ss) {
  const dp = PropertiesService.getDocumentProperties(), accs = accList_(ss), fam = familyMode_(ss);
  const has = fn => hasTrigger_(fn), rows = ss.getSheetByName(JOURNAL) ? readJournal_(ss).filter(o => !o.blank) : [];
  const cp = ss.getSheetByName(CP_SHEET) ? ss.getSheetByName(CP_SHEET).getRange(CP_FIRST, 1, 30, 2).getValues().some(r => r[0] && r[1] !== '') : false;
  const cs = checkState_(ss), up = PropertiesService.getUserProperties(), fops = fopList_(ss), mono = !!up.getProperty('mono_token');
  const hid = modOn_('service') ? '' : ' (видно з модулем «' + modName_('service') + '», крок 5)';
  const P = P_STEP;
  return [
    { n: 0, ok: fam ? members_(ss).length >= 2 : null, what: 'Учасники сім\'ї (до 5)', where: 'угорі панелі: «Сім\'я»' },
    { n: 1, ok: accs.length > 0, what: 'Картка й банк: Mono, ПриватБанк для ФОП або виписка', where: P + '1' },
    { n: 1, ok: mono || up.getProperty('pb_token') ? has(TRG.imp) : null, what: 'Імпорт з банку щодня', where: P + '1: «Імпорт з банку щодня»' },
    { n: 1, ok: cp, what: 'Залишок на картці (контрольна точка)', where: mono ? 'сам з Mono' : 'аркуш «' + CP_SHEET + '»' + hid },
    { n: 1, ok: modOn_('fop') ? fops.length > 0 : null, what: 'ФОП: група й ставки', where: P + '1: «Рахунок ФОП»' },
    { n: 2, ok: modOn_('tg') ? tgConnected_() : null, what: 'Telegram-бот: ранковий підсумок і звіти', where: P + '2' },
    { n: 3, ok: ss.getSheetByName(REG_SHEET) ? regOwn_(ss).length > 0 : null, what: 'Регулярні платежі й доходи (оренда, підписки, зарплата)', where: P + '3 або аркуш «' + REG_SHEET + '»' },
    { n: 4, ok: modOn_('advisor') ? !!dp.getProperty('adv_doc_id') : null, what: 'Радник Gemini: знімок і навичка', where: P + '4' },
    { n: 5, ok: !!dp.getProperty('mods_seen') || !!dp.getProperty('wiz_done'), what: 'Модулі: що увімкнути', where: P + '5 (меню «🧩 Модулі»)' },
    { n: 0, ok: rows.length > 0, what: 'Перші записи в журналі', where: 'аркуш «' + JOURNAL + '» або «Демо-місяць» у панелі' },
    { n: 0, ok: modOn_('cal') ? !!dp.getProperty('cal_stat') : null, what: 'Календар платежів', where: 'оновлюється сам' },
    { n: 0, ok: modOn_('debts') ? debtsList_(ss).length > 0 : null, what: 'Борги й розстрочки', where: 'аркуш «' + DEBTS_SHEET + '»' + hid },
    { n: 0, ok: cs.none ? null : cs.bad.length === 0, what: '«' + CHECK_SHEET + '»: усе OK',
      where: cs.none || !cs.bad.length ? 'перевіряється сам' : 'аркуш «' + CHECK_SHEET + '»: ' + cs.bad.length + ' ⚠️' + hid },
  ].filter(x => x.ok !== null);
}

/** Оновити «Старт»: кроки, смужка «готово», низ з «Підтримати». */
function startRefresh_(ss) {
  const sh = ss.getSheetByName(START_SHEET);
  if (!sh) return null;
  const steps = startSteps_(ss), done = steps.filter(x => x.ok).length, need = steps.length;
  ensureSize_(sh, START_FIRST + START_ROWS, 4);
  const lim = modOldOnStart_(ss) ? MOD_OLD_HEAD - START_FIRST : START_ROWS;    // старий блок модулів (до версії 12) не чіпаємо
  sh.getRange(START_FIRST, 1, lim, 3).clearContent();
  sh.getRange(START_FIRST, 1, lim, 3).setFontWeight('normal').setFontColor('#000000');
  sh.getRange(2, 1).setValue('Меню «📊 ' + APP_NAME + ' → 🚀 Підключення» на комп\'ютері (на телефоні меню скриптів немає). ' +
    'Число: крок у панелі. Галочки ставляться самі.').setWrap(false);
  sh.getRange(3, 1).setValue('▰'.repeat(done) + '▱'.repeat(need - done) + '  ' + done + ' з ' + need + (done === need ? ' ✅' : '')).setFontWeight('bold');
  sh.getRange(START_FIRST - 1, 1, 1, 3).setValues([['', 'Крок', 'Де']]).setFontWeight('bold');
  sh.getRange(START_FIRST, 1, steps.length, 3).setValues(steps.map(x => [(x.ok ? '✅ ' : '⬜ ') + (x.n || '·'), x.what, x.where.indexOf(P_STEP) === 0 ? '«🚀 Підключення», ' + x.where : x.where]));
  const foot = START_FIRST + steps.length + 1;
  if (foot + 1 <= START_FIRST + lim - 1) {
    setF_(sh.getRange(foot, 1), [['=HYPERLINK("' + SUPPORT_URL + '","❤️ Підтримати автора")']]);
    sh.getRange(foot, 2).setValue('Шаблон безкоштовний, код відкритий: github.com/VladyYerys/monitoring-budgets');
    sh.getRange(foot, 2).setFontColor('#5f6368');
  }
  PropertiesService.getDocumentProperties().setProperty('start_day', todayKey_(ss));
  return { done: done, need: need };
}

/** З mbsTick: раз на день. */
function startIfDue_(ss) {
  if (!ss.getSheetByName(START_SHEET) || PropertiesService.getDocumentProperties().getProperty('start_day') === todayKey_(ss)) return null;
  return startRefresh_(ss);
}

/* ---------- майстер ---------- */

/* ---------- демо-місяць ---------- */

/** Вигадані операції за минулий і цей місяць (до вчора), два плани наперед і контрольна точка. Детерміновано (без випадковості). */
function demoAdd_(ss) {
  const tz = ss.getSpreadsheetTimeZone(), today = todayKey_(ss), ym = today.slice(0, 7), prev = prevYm_(ym), dp = PropertiesService.getDocumentProperties();
  const D = k => Utilities.parseDate(k, tz, 'yyyy-MM-dd'), fam = familyMode_(ss), ms = members_(ss);
  const who = i => (fam && ms.length ? (i % 3 === 2 ? WHO_SHARED : ms[i % ms.length]) : '');
  const cats = new Set(ss.getSheetByName(DICT_SHEET).getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 1).getValues().map(r => String(r[0]).trim()));
  const ops = [], add = (key, b, c, d, i, status) => { if (key < today && cats.has(c)) ops.push([D(key), b, c, DEMO_TAG + d, status || STATUS_FACT, '', who(i), '']); };
  const plan = [['Продукти', 420, 'супермаркет', 3], ['Кафе й ресторани', 380, 'обід', 7], ['Кава та перекуси', 75, 'кава', 2], ['Таксі', 160, 'поїздка', 9],
    ['Транспорт', 30, 'метро', 4], ['Здоров\'я й аптека', 260, 'аптека', 11], ['Дрібні покупки', 199, 'маркетплейс', 13]];
  [prev, ym].forEach(m => {
    for (let day = 1; day <= Number(lastDayOf_(m).slice(8)); day++) {
      const key = m + '-' + pad2_(day);
      plan.forEach((p, j) => { if ((day + j) % p[3] === 0) add(key, p[1] + ((day * 37 + j * 11) % 9) * 10, p[0], p[2], day + j); });
    }
    add(m + '-05', 2400, 'Комуналка', 'комунальні послуги', 1);
    add(m + '-10', 350, 'Інтернет і зв\'язок', 'інтернет і телефон', 2);
    add(m + '-12', 299, 'Підписки', 'стрімінг', 0);
    add(m + '-25', 42000, 'Зарплата', 'зарплата', 0);
    add(m + '-26', 6000, '🎯 На погашення ліміту', 'відкладено на ліміт', 0);
  });
  const next = addDays_(lastDayOf_(ym), 1).slice(0, 7), reg = ss.getSheetByName(REG_SHEET);
  // регулярні: правила на «Регулярних» (плани поставить скрипт); старіша таблиця без аркуша: два плани в журналі
  const regRules = [[DEMO_TAG + 'Комуналка', 'Комуналка', 2400, 5], [DEMO_TAG + 'Підписка', 'Підписки', 299, 12]].filter(r => cats.has(r[1]));
  const plans = reg ? [] : [[D(next + '-05'), 2400, 'Комуналка', DEMO_TAG + 'комунальні послуги', STATUS_PLAN, '', '', ''],
    [D(next + '-12'), 299, 'Підписки', DEMO_TAG + 'стрімінг', STATUS_PLAN, '', '', '']].filter(r => cats.has(r[2]));
  const all = ops.concat(plans), sh = ss.getSheetByName(JOURNAL), rows = readJournal_(ss);
  let lastUsed = -1;
  rows.forEach((r, i) => { if (!r.blank || r.bank.some(x => x !== '' && x !== null) || (r.ord !== '' && r.ord !== null) || r.acct) lastUsed = i; });
  const fit = all.slice(0, Math.max(0, rows.length - lastUsed - 1));
  if (fit.length) sh.getRange(J_FIRST + lastUsed + 1, 1, fit.length, J_INPUT_N).setValues(fit);
  // контрольна точка на початок минулого місяця, старт обліку з минулого місяця, плани двох категорій
  const set = ss.getSheetByName(SETTINGS_SHEET), state = { start: dateKey_(set.getRange(SET.start).getValue(), tz), cp: prev + '-01', plans: [] };
  if (!state.start || state.start > prev + '-01') set.getRange(SET.start).setValue(D(prev + '-01'));
  upsertCheckpoints_(ss, [{ key: prev + '-01', own: -15000, limit: cardIsCredit_(ss) ? 50000 : 0 }]);
  const d = ss.getSheetByName(DICT_SHEET), dv = d.getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 6).getValues();
  [['Продукти', 9000], ['Кафе й ресторани', 3000], ['Таксі', 1500]].forEach(p => {
    const i = dv.findIndex(r => String(r[0]).trim() === p[0]);
    if (i >= 0 && (dv[i][5] === '' || dv[i][5] === null)) { d.getRange(D_FIRST + i, 6).setValue(p[1]); state.plans.push(D_FIRST + i); }
  });
  if (reg && regRules.length) {
    const v = reg.getRange(REG_FIRST, 1, REG_LAST - REG_FIRST + 1, 1).getValues();
    let k = 0;
    v.forEach((r, i) => { if (k < regRules.length && String(r[0]).trim() === '') { const x = regRules[k++]; reg.getRange(REG_FIRST + i, 1, 1, 9).setValues([[x[0], x[1], x[2], x[3], REG_PER[0], '', '', '', MOD_YES]]); } });
  }
  dp.setProperty('demo', JSON.stringify(state));
  sortCore_(ss);
  if (reg) regPlansSync_(ss);
  return { rows: fit.length };
}

/** Прибрати демо: рядки з «[демо]», контрольну точку, плани категорій, старт обліку як був. */
function demoClear_(ss) {
  const sh = ss.getSheetByName(JOURNAL), tz = ss.getSpreadsheetTimeZone(), dp = PropertiesService.getDocumentProperties(), state = jsonProp_(dp, 'demo') || {};
  let n = 0;
  readJournal_(ss).forEach(o => {
    if (o.d.indexOf(DEMO_TAG) !== 0) return;
    sh.getRange(o.row, 1, 1, J_INPUT_N).setValues([['', '', '', '', '', '', '', '']]);
    sh.getRange(o.row, colNum_(JC.bankId), 1, J_MOVE_N).setValues([emptyBank_().concat(['', ''])]);
    n++;
  });
  if (state.cp) {
    const cp = ss.getSheetByName(CP_SHEET), v = cp.getRange(CP_FIRST, 1, CP_LAST - CP_FIRST + 1, 3).getValues();
    v.forEach((r, i) => { if (dateKey_(r[0], tz) === state.cp && Number(r[1]) === -15000) cp.getRange(CP_FIRST + i, 1, 1, 3).setValues([['', '', '']]); });
  }
  (state.plans || []).forEach(r => ss.getSheetByName(DICT_SHEET).getRange(r, 6).setValue(''));
  const reg = ss.getSheetByName(REG_SHEET);
  if (reg) {
    reg.getRange(REG_FIRST, 1, REG_LAST - REG_FIRST + 1, 1).getValues().forEach((r, i) => {
      if (String(r[0]).indexOf(DEMO_TAG) === 0) reg.getRange(REG_FIRST + i, 1, 1, REG_COLS.length).setValues([REG_COLS.map(() => '')]);
    });
  }
  if (state.start) ss.getSheetByName(SETTINGS_SHEET).getRange(SET.start).setValue(Utilities.parseDate(state.start, tz, 'yyyy-MM-dd'));
  dp.deleteProperty('demo');
  sortCore_(ss);
  if (reg) regPlansSync_(ss);                                       // плани демо-правил теж геть
  return n;
}

function demoAdd() {
  if (!menuOk_('🎬 Демо-місяць')) return;
  const ss = SpreadsheetApp.getActive();
  if (jsonProp_(PropertiesService.getDocumentProperties(), 'demo')) { show_('🎬 Демо-місяць', ['Демо вже додано. Прибрати: «🧹 Очистити демо».']); return; }
  const r = demoAdd_(ss);
  show_('🎬 Демо-місяць', ['Додано ' + r.rows + ' вигаданих операцій (коментар «[демо]»). Гляньте «' + DASH_SHEET + '» і «' + MOB_SHEET + '».',
    'Прибрати: «🧹 Очистити демо».']);
}

function demoClear() {
  if (!menuOk_('🧹 Очистити демо')) return;
  const n = demoClear_(SpreadsheetApp.getActive());
  show_('🧹 Демо прибрано', ['Прибрано рядків: ' + n + '. Ваші записи не чіпались.']);
}

function startNow() {
  if (!menuOk_('🚀 Старт')) return;
  const r = startRefresh_(SpreadsheetApp.getActive());
  show_('🚀 Старт', [r ? 'Чекліст оновлено: ' + r.done + ' з ' + r.need + '.' : 'Аркуша «' + START_SHEET + '» немає: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».']);
}
