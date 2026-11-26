/** 38_Start.gs: аркуш «Старт» з чеклістом і модулями, демо-місяць. Підключення (картка, Telegram, регулярні, Gemini): 44_Connect.gs. */

/*
 * У копії шаблону розклади й дозволи не переносяться, тож усе починається з меню «📊 Monitoring Budgets → 🚀 Підключення»:
 * скрипт добудовує, чого бракує, і відкриває бічну панель з чотирма кроками (44_Connect.gs).
 * Чекліст на «Старті» оновлює скрипт (раз на день, після «Готово» в панелі «🚀 Підключення» і з меню): ✓ зроблено, • ще ні, «—» не потрібно.
 * Демо-місяць: вигадані операції з позначкою «[демо]» у коментарі, щоб одразу побачити «Charts» і «charts_mob»; «🧹 Очистити демо» прибирає все.
 */

const START_FIRST = 6;
const DEMO_TAG = '[демо] ';

function buildStart_(ss) {
  const sh = ensureSize_(sheetOf_(ss, START_SHEET), MOD_FIRST + MODS.length + 5, 4);
  sh.getRange(1, 1).setValue('🚀 Старт').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Почніть з меню «📊 ' + APP_NAME + ' → 🚀 Підключення» (на комп\'ютері: на телефоні меню скриптів немає). ' +
    'Кроки нижче позначаються самі.').setWrap(false);
  sh.getRange(START_FIRST - 1, 1, 1, 3).setValues([['', 'Крок', 'Як']]).setFontWeight('bold');
  startModulesBlock_(ss, sh);
  startRefresh_(ss);
  return sh;
}

/** [перевірка, крок, як]. Перевірка → true (✓), false (•), null (не потрібно). */
function startSteps_(ss) {
  const dp = PropertiesService.getDocumentProperties(), set = ss.getSheetByName(SETTINGS_SHEET), accs = accList_(ss), fam = familyMode_(ss);
  const has = fn => hasTrigger_(fn), rows = ss.getSheetByName(JOURNAL) ? readJournal_(ss).filter(o => !o.blank) : [];
  const cp = ss.getSheetByName(CP_SHEET) ? ss.getSheetByName(CP_SHEET).getRange(CP_FIRST, 1, 30, 2).getValues().some(r => r[0] && r[1] !== '') : false;
  const cs = checkState_(ss), up = PropertiesService.getUserProperties(), fops = fopList_(ss);
  const hid = modOn_('service') ? '' : ' (сховано: «' + modName_('service') + '»: так)';
  return [
    [!!dp.getProperty('wiz_done'), '«🚀 Підключення» пройдено (кнопка «Готово»)', 'меню «📊 ' + APP_NAME + ' → 🚀 Підключення»'],
    [!!(set && set.getRange(SET.start).getValue()), 'Старт обліку «моїх грошей», вид картки, день доходу', '«🚀 Підключення» або аркуш «' + SETTINGS_SHEET + '»' + hid],
    [fam ? members_(ss).length >= 2 : null, 'Учасники сім\'ї (до 5)', '«🚀 Підключення» («Сім\'я») або «' + SETTINGS_SHEET + '»' + hid],
    [accs.length > 0, 'Банк: Mono, ПриватБанк для ФОП або виписка', '«🚀 Підключення», крок 1'],
    [up.getProperty('mono_token') || up.getProperty('pb_token') ? has(TRG.imp) : null, 'Імпорт з банку за розкладом (ваш)', '«🚀 Підключення», крок 1: «Імпорт з банку щодня»'],
    [cp, 'Залишок на картці (контрольна точка)', 'сам з Mono або вручну на «' + CP_SHEET + '»' + hid],
    [rows.length > 0, 'Перші записи в журналі', 'аркуш «' + JOURNAL + '» або демо-місяць'],
    [ss.getSheetByName(REG_SHEET) ? regOwn_(ss).length > 0 : null, 'Регулярні платежі й доходи (оренда, підписки, зарплата)', '«🚀 Підключення», крок 3, або аркуш «' + REG_SHEET + '»'],
    [modOn_('tg') ? tgConnected_() : null, 'Telegram-бот (ранковий підсумок, звіти)', '«🚀 Підключення», крок 2'],
    [modOn_('cal') ? !!dp.getProperty('cal_stat') : null, 'Календар платежів', 'увімкнено модулем; оновлюється сам'],
    [modOn_('fop') ? fops.length > 0 : null, 'ФОП: група, ставки', 'аркуш «' + FOP_SHEET + '»' + hid],
    [modOn_('debts') ? debtsList_(ss).length > 0 : null, 'Борги й розстрочки', 'аркуш «' + DEBTS_SHEET + '»' + hid],
    [modOn_('advisor') ? !!dp.getProperty('adv_doc_id') : null, 'Радник Gemini: знімок і навичка', '«🚀 Підключення», крок 4'],
    [cs.none ? null : cs.bad.length === 0, '«' + CHECK_SHEET + '»: усе OK', 'аркуш «' + CHECK_SHEET + '»' + hid],
  ];
}

/** Оновити чекліст «Старту» і нижній блок «Підтримати». */
function startRefresh_(ss) {
  const sh = ss.getSheetByName(START_SHEET);
  if (!sh) return null;
  const steps = startSteps_(ss);
  sh.getRange(START_FIRST, 1, steps.length, 3).setValues(steps.map(s => [s[0] === null ? '—' : s[0] ? '✓' : '•', s[1], s[2]]));
  const done = steps.filter(s => s[0] === true).length, need = steps.filter(s => s[0] !== null).length;
  sh.getRange(3, 1).setValue('Готово: ' + done + ' з ' + need + (done === need ? ' ✅' : '')).setFontWeight('bold');
  const foot = START_FIRST + steps.length + 2;
  if (MOD_HEAD > START_FIRST + steps.length) sh.getRange(START_FIRST + steps.length, 1, MOD_HEAD - START_FIRST - steps.length, 3).clearContent();   // старий низ (крок додався)
  sh.getRange(foot, 1, 2, 1).setValues([['Шаблон безкоштовний і з відкритим кодом (github.com/VladyYerys/monitoring-budgets).'], ['Корисний?']]);
  setF_(sh.getRange(foot + 1, 2), [['=HYPERLINK("' + SUPPORT_URL + '","❤️ Підтримати автора")']]);
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
