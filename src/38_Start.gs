/** 38_Start.gs: аркуш «⚡ Швидке підключення»: стан таблиці й покрокова інструкція для ШІ (Gemini чи іншого); демо-місяць. */

/*
 * Перша вкладка. Людина відкриває ШІ з доступом до таблиці (Gemini справа вгорі в Таблицях) і пише:
 * «Прочитай аркуш «⚡ Швидке підключення» і налаштуй зі мною таблицю». Далі ШІ веде сам за цим аркушем:
 *   1. СТАН: що вже підключено (пише скрипт: раз на день, після «Готово» в панелі, галочок модулів і з меню);
 *   2. ЯК ВЕСТИ ЛЮДИНУ: що за чим; банк, Telegram, модулі й радника робить людина в панелі «🚀 Підключення»
 *      (меню, дозволи Google, токени: ШІ цього не може і не повинен);
 *   3. КАРТА: завдання → аркуш → колонки → приклад (регулярні, бюджети, борги, «Мої правила», правила, журнал);
 *   4. ЯК ВІДПОВІДАТИ про гроші: загальне з шапки «Журналу» й «Головного», конкретні витрати з «Журналу»;
 *   5. НЕ МОЖНА: формули, службові аркуші, токени, сортування, «✅ факт» без прохання;
 *   6. ГОТОВІ ЗАПИТИ для людини.
 * У «Журнал» ШІ пише лише на пряме прохання людини, у рядок для вводу, колонки A–E (G для сім'ї).
 * До версії 15 аркуш звався «Старт» (кроки з галочками).
 * Демо-місяць: вигадані операції з позначкою «[демо]» у коментарі, щоб одразу побачити «Charts» і «charts_mob»; «🧹 Очистити демо» прибирає все.
 */

const START_OLD = 'Старт';                                               // назва аркуша до версії 15
const QUICK_TITLE = '⚡ Швидке підключення';
const START_FIRST = 7, START_ROWS = 110;                                 // стан з рядка 7; усе разом: до 110 рядків
const DEMO_TAG = '[демо] ';
const P_STEP = 'крок ';

function buildStart_(ss) {
  const sh = ensureSize_(sheetOf_(ss, START_SHEET), START_ROWS + 5, 3);
  if (!startHasModules_(ss) && ss.getSheetByName(SETTINGS_SHEET)) modBlockBuild_(ss);
  startRefresh_(ss);
  return sh;
}

/** Стан для ШІ: [що, стан, хто й де]. done/need: скільки з основного вже готово. */
function quickState_(ss) {
  const dp = PropertiesService.getDocumentProperties(), up = PropertiesService.getUserProperties(), accs = accList_(ss), fam = familyMode_(ss);
  const rows = ss.getSheetByName(JOURNAL) ? readJournal_(ss).filter(o => !o.blank) : [];
  const cp = ss.getSheetByName(CP_SHEET) ? ss.getSheetByName(CP_SHEET).getRange(CP_FIRST, 1, 30, 2).getValues().some(r => r[0] && r[1] !== '') : false;
  const cs = checkState_(ss), mono = !!up.getProperty('mono_token'), bank = mono || !!up.getProperty('pb_token') || accs.length > 0;
  const imp = bank ? hasTrigger_(TRG.imp) : false, tg = modOn_('tg') ? tgConnected_() : null;
  const regs = ss.getSheetByName(REG_SHEET) ? regOwn_(ss).length : 0, d = ss.getSheetByName(DICT_SHEET);
  const budgets = d ? d.getRange(D_FIRST, 6, D_LAST - D_FIRST + 1, 1).getValues().filter(r => Number(r[0]) > 0).length : 0;
  const drafts = rows.filter(o => o.e === STATUS_DRAFT), red = drafts.filter(o => o.c === CAT_TRANSFER || !o.c).length;
  const off = 'модуль вимкнено (людина: ' + P_STEP + '5)', H = 'людина: «🚀 Підключення», ', AI = 'ШІ: ';
  const adv = modOn_('advisor') ? (dp.getProperty('adv_doc_id') ? 'знімок є' : 'знімка ще немає') : off;
  const list = [
    ['Режим', fam ? 'сім\'я: ' + members_(ss).join(', ') : 'я сам', 'людина: угорі панелі «🚀 Підключення»'],
    ['Банк', accs.length ? accs.map(a => a.name).join(', ') : 'не підключено', H + P_STEP + '1'],
    ['Імпорт з банку щодня', !bank ? 'після банку' : imp ? 'так' : 'ні', H + P_STEP + '1'],
    ['Залишок на картці (контрольна точка)', cp ? 'є' : 'немає', mono ? 'сам з Mono' : 'людина: аркуш «' + CP_SHEET + '»'],
    ['Telegram-бот', tg === null ? off : tg ? 'підключено' : 'не підключено', H + P_STEP + '2'],
    ['Регулярні платежі й доходи', 'своїх: ' + regs, AI + 'аркуш «' + REG_SHEET + '» (див. КАРТУ)'],
    ['Бюджети категорій на місяць', modOn_('budgets') ? 'з планом: ' + budgets : off, AI + '«' + DICT_SHEET + '», колонка F'],
    ['Борги й розстрочки', modOn_('debts') ? 'рядків: ' + debtsList_(ss).length : off, AI + 'аркуш «' + DEBTS_SHEET + '»'],
    ['ФОП', modOn_('fop') ? (fopList_(ss).length ? 'ФОП: ' + fopList_(ss).length : 'не вказано') : off, H + P_STEP + '1 («Рахунок ФОП»)'],
    ['Мої правила для радника', advRules_(ss).filled ? 'заповнено' : 'не заповнено', AI + '«' + SETTINGS_SHEET + '», B27–B32'],
    ['Радник Gemini', adv, H + P_STEP + '4'],
    ['Модулі увімкнено', MODS.filter(m => modOn_(m.id)).map(m => m.name).join(', ') || 'жодного', H + P_STEP + '5 (меню «🧩 Модулі»)'],
    ['Операції в журналі', 'фактів: ' + rows.filter(o => o.e === STATUS_FACT).length + '; чернеток: ' + drafts.length + (red ? ', з них без категорії: ' + red : ''), 'аркуш «' + JOURNAL + '»'],
    ['«' + CHECK_SHEET + '»', cs.none ? 'ще не рахувалась' : cs.bad.length ? cs.bad.length + ' ⚠️' : 'усе OK', 'перевіряється сам'],
    ['Оновлено', Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm'), 'меню «⚡ Оновити «Швидке підключення»»'],
  ];
  const need = [accs.length > 0, !bank || imp, tg !== false, regs > 0, !modOn_('advisor') || !!dp.getProperty('adv_doc_id'), !!dp.getProperty('mods_seen') || !!dp.getProperty('wiz_done')];
  return { rows: list, done: need.filter(Boolean).length, need: need.length };
}

/** Інструкція для ШІ (не залежить від стану): розділи 2–6. [A, B, C]; рядок з одним A і '#': заголовок розділу. */
function quickGuide_() {
  const J = '«' + JOURNAL + '»', R = '«' + REG_SHEET + '»', D = '«' + DICT_SHEET + '»', S = '«' + SETTINGS_SHEET + '»';
  return [
    ['# 2. ДЛЯ ШІ: ЯК ВЕСТИ ЛЮДИНУ (по черзі, по одному питанню)'],
    ['1', 'Прочитай розділ 1 «СТАН». Скажи коротко, що вже готово і що далі.', ''],
    ['2', 'Банк, Telegram, модулі й радника робить людина в панелі: меню «📊 ' + APP_NAME + ' → 🚀 Підключення» на комп\'ютері (на телефоні меню скриптів немає). ' +
      'Токени, номери карток, IBAN не проси й не приймай: їх вставляють лише в панель.', 'людина'],
    ['3', 'Регулярні: розпитай про оренду, комуналку, інтернет, підписки, зарплату, кредити й розстрочки. Кожен платіж чи дохід: рядок у ' + R + ' (див. КАРТУ). ' +
      'Рядки «💡 Приклад» з «ні» можна поправити й поставити «так».', 'ШІ'],
    ['4', 'Бюджети: запропонуй план на місяць для 5–7 головних категорій витрат. Якщо є «Місяці» за 2–3 місяці, спирайся на них. Пиши в ' + D + ', колонка F.', 'ШІ'],
    ['5', 'Борги й розстрочки (якщо модуль увімкнено): рядок на кожен у «' + DEBTS_SHEET + '».', 'ШІ'],
    ['6', '«Мої правила» для радника: пріоритет, подушка, запас, розстрочки, ФОП, свої правила: ' + S + ', B27–B32.', 'ШІ'],
    ['7', 'Наприкінці: дивитись на «' + DASH_SHEET + '» (комп\'ютер) і «' + MOB_SHEET + '» (телефон); операції з банку приходять самі; питати можна тебе.', 'ШІ'],
    ['# 3. КАРТА: ЩО ДЕ ЗАПИСУВАТИ', 'Аркуш і колонки', 'Приклад'],
    ['Регулярний платіж чи дохід', R + ', перший порожній рядок з 5-го: A назва, B категорія (лише з аркуша ' + D + ', колонка A), C сума, D день 1–31, ' +
      'E як часто (' + REG_PER.join(' / ') + '), F перший платіж (дата, для щокварталу й щороку), G до (порожньо: без кінця), H хто (сім\'я), I «так»',
      'Оренда · Оренда житла · 15000 · 5 · щомісяця · · · · так'],
    ['План категорії на місяць', D + ': рядок категорії (колонка A), колонка F: сума ₴', 'Продукти → F: 9000'],
    ['Борг чи розстрочка', '«' + DEBTS_SHEET + '», з рядка 5: A назва, B напрям (' + DEBT_OWE + ' / ' + DEBT_LENT + '), C сума боргу, D платіж на місяць, E день 1–28, ' +
      'F перший платіж, G хто, H категорія (можна порожньо)', 'Розстрочка за телефон · я винен · 12000 · 1000 · 15'],
    ['Мої правила для радника', S + ': B27 пріоритет (зі списку), B28 подушка % (0,1 = 10 %), B29 запас ₴, B30 розстрочки обов\'язкові (так / ні), ' +
      'B31 ФОП % , B32 свої правила (до ' + ADV_RULES_MAX + ' знаків)', 'B29: 3000'],
    ['Правило категорії для банку', '«' + RULES_SHEET + '», вільний рядок: A «так», B слова через кому, C де шукати (' + RULE_WHERE.join(' / ') + '), ' +
      'D напрям (' + RULE_DIR.join(' / ') + '), E категорія з аркуша ' + D, 'так · затак · опис банку · списання · Продукти'],
    ['Витрата чи дохід вручну (ЛИШЕ на пряме прохання людини)', J + ': рядок з шапки «✍️ Писати у рядок № …» (порожній). A дата, B сума (завжди додатна), ' +
      'C категорія з аркуша ' + D + ', D коментар, E «' + STATUS_FACT + '» (чи «' + STATUS_PLAN + '» для майбутнього), G хто (сім\'я). Колонки з I і далі не чіпати.',
      '27.11.2026 · 85 · Кава та перекуси · кава · ' + STATUS_FACT],
    ['# 4. ЯК ВІДПОВІДАТИ НА ПИТАННЯ ПРО ГРОШІ'],
    ['Загальне (скільки є, скільки можна, як справи)', 'шапка аркуша ' + J + ' (рядки 1–3: «Мої зараз», «На картці», бюджети) і «' + MAIN_SHEET + '» (схований аркуш: підпис у A, число в B: ' +
      '«Доступно моїх сьогодні», «Можна на день», «Прогноз на кінець місяця (за планом)»…). Готові цифри не перераховуй.', ''],
    ['Конкретні витрати (за категорією, місяць, магазин)', J + ' з рядка 5: A дата, B сума, C категорія, D коментар, E статус, I тип. Рахуй «' + STATUS_FACT + '»; ' +
      'чернетки «' + STATUS_DRAFT + '» і плани показуй окремо. Підсумки місяців: «' + MONTHS_SHEET + '».', ''],
    ['# 5. НЕ МОЖНА'],
    ['Формули', 'Не змінюй формули й колонки аркуша ' + J + ' з I і далі; аркуші «' + DAYS_SHEET + '», «' + MAIN_SHEET + '», «' + CH_SHEET + '», «' + DASH_SHEET + '», «' + MOB_SHEET + '», «' + CHECK_SHEET +
      '», «' + BANK_SHEET + '», «' + ACC_SHEET + '» пише лише скрипт.', ''],
    ['Журнал', 'Не видаляй і не сортуй рядки (порядок веде скрипт); не став «' + STATUS_FACT + '» чернеткам і не змінюй чужі рядки без прохання людини.', ''],
    ['Категорії', 'Лише з аркуша ' + D + ' (колонка A). Нова категорія: спершу рядок на аркуші ' + D + ' (A назва, B тип, C група).', ''],
    ['Секрети', 'Токени банку й бота, номери карток, IBAN, паролі, коди з SMS: не проси й не вписуй нікуди.', ''],
    ['# 6. ГОТОВІ ЗАПИТИ (для людини)'],
    ['«Прочитай аркуш «' + QUICK_TITLE + '» і налаштуй зі мною таблицю»', '', ''],
    ['«Додай мої підписки: Netflix 350 ₴ 12-го, YouTube 150 ₴ 3-го»', '', ''],
    ['«Постав бюджети на місяць за моїми витратами за 3 місяці»', '', ''],
    ['«Скільки я витратив на кафе цього місяця?»', '', ''],
    ['«Запиши витрату готівкою: кава 85 ₴ сьогодні»', '', ''],
    ['«Що ще не налаштовано?»', '', ''],
  ];
}

/** Оновити «⚡ Швидке підключення»: стан і інструкція. → { done, need } */
function startRefresh_(ss) {
  const sh = ss.getSheetByName(START_SHEET);
  if (!sh) return null;
  const st = quickState_(ss), guide = quickGuide_(), tz = ss.getSpreadsheetTimeZone();
  ensureSize_(sh, START_ROWS + 5, 3);
  const lim = modOldOnStart_(ss) ? MOD_OLD_HEAD - 1 : START_ROWS;         // старий блок модулів (до версії 12) не чіпаємо
  const all = sh.getRange(1, 1, lim, 3);
  all.clearContent();
  all.setFontWeight('normal').setFontColor('#000000').setFontSize(10).setWrap(true).setBackground(null);
  sh.getRange(1, 1).setValue(QUICK_TITLE).setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('З ШІ: відкрийте Gemini в цій таблиці (зірочка справа вгорі) чи інший ШІ з доступом до неї й напишіть: «Прочитай аркуш «' + QUICK_TITLE +
    '» і налаштуй зі мною таблицю». Без ШІ: меню «📊 ' + APP_NAME + ' → 🚀 Підключення» на комп\'ютері.').setWrap(false);
  sh.getRange(3, 1).setValue('Токени банку й бота, номери карток і IBAN: лише в панелі «🚀 Підключення», ніколи в чат ШІ чи в клітинки.').setFontColor('#5f6368').setWrap(false);
  sh.getRange(5, 1).setValue('# 1. СТАН (оновлює скрипт; ШІ читає звідси). Готово основного: ' + st.done + ' з ' + st.need).setFontWeight('bold').setWrap(false);
  sh.getRange(6, 1, 1, 3).setValues([['Що', 'Стан', 'Хто й де']]).setFontWeight('bold');
  sh.getRange(START_FIRST, 1, st.rows.length, 3).setValues(st.rows);
  let r = START_FIRST + st.rows.length + 1;
  const g = guide.map(x => [x[0], x[1] || '', x[2] || '']);
  if (r + g.length + 2 <= lim) {
    sh.getRange(r, 1, g.length, 3).setValues(g);
    g.forEach((x, i) => { if (String(x[0]).indexOf('# ') === 0) sh.getRange(r + i, 1, 1, 3).setFontWeight('bold').setBackground('#f1f3f4'); });
    r += g.length + 1;
    setF_(sh.getRange(r, 1), [['=HYPERLINK("' + SUPPORT_URL + '","❤️ Підтримати автора")']]);
    sh.getRange(r, 2).setValue('Шаблон безкоштовний, код відкритий: github.com/VladyYerys/monitoring-budgets').setFontColor('#5f6368');
  }
  try { sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 520); sh.setColumnWidth(3, 300); } catch (e) { /* ignore */ }
  PropertiesService.getDocumentProperties().setProperty('start_day', todayKey_(ss));
  return { done: st.done, need: st.need, tz: tz };
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
  if (!menuOk_(QUICK_TITLE)) return;
  const r = startRefresh_(SpreadsheetApp.getActive());
  show_(QUICK_TITLE, [r ? 'Оновлено. Готово основного: ' + r.done + ' з ' + r.need + '.' : 'Аркуша «' + START_SHEET + '» немає: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».']);
}
