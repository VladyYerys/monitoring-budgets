/** 38_Start.gs: «🚀 Почати»: аркуш «Старт» з чеклістом і модулями, майстер (режим, картка, доходи, категорії, модулі), демо-місяць. */

/*
 * У копії шаблону розклади й дозволи не переносяться, тож усе починається з меню «📊 Monitoring Budgets → 🚀 Почати»:
 * скрипт добудовує, чого бракує, і ставить кілька запитань. Відповіді можна змінити пізніше на «Налаштуваннях», «Довіднику», «ФОП».
 * Чекліст на «Старті» оновлює скрипт (раз на день, після майстра і з меню): ✓ зроблено, • ще ні, «—» не потрібно.
 * Демо-місяць: вигадані операції з позначкою «[демо]» у коментарі, щоб одразу побачити «Головне» й графіки; «🧹 Очистити демо» прибирає все.
 */

const START_FIRST = 6;
const DEMO_TAG = '[демо] ';
const CAT_SETS = [
  ['Діти', ['Діти']], ['Тварини', ['Тварини']], ['Авто', ['Пальне']], ['Оренда житла', ['Оренда житла']],
  ['ФОП і податки', ['Єдиний податок', 'ЄСВ', 'Військовий збір', 'Дохід ФОП', 'Резерв на податки']],
];

function buildStart_(ss) {
  const sh = ensureSize_(sheetOf_(ss, START_SHEET), MOD_FIRST + MODS.length + 5, 4);
  sh.getRange(1, 1).setValue('🚀 Старт').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Почніть з меню «📊 ' + APP_NAME + ' → 🚀 Почати» (на комп\'ютері: на телефоні меню скриптів немає). ' +
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
  return [
    [!!dp.getProperty('wiz_done'), 'Майстер «🚀 Почати» пройдено', 'меню «📊 ' + APP_NAME + ' → 🚀 Почати»'],
    [!!(set && set.getRange(SET.start).getValue()), 'Старт обліку «моїх грошей», вид картки, день доходу', 'аркуш «' + SETTINGS_SHEET + '»'],
    [fam ? members_(ss).length >= 2 : null, 'Учасники сім\'ї (до 5)', '«' + SETTINGS_SHEET + '», блок «Сім\'я»'],
    [accs.length > 0, 'Банк: Mono, ПриватБанк для ФОП або виписка', 'меню «⚙️ Налаштування» або «📥 Імпорт виписок з папки»'],
    [up.getProperty('mono_token') || up.getProperty('pb_token') ? has(TRG.imp) : null, 'Імпорт з банку за розкладом (ваш)', '«⚙️ Налаштування → 🏦 Імпорт за розкладом: увімкнути»'],
    [cp, 'Залишок на картці (контрольна точка)', 'сам з Mono або вручну на «' + CP_SHEET + '»'],
    [rows.length > 0, 'Перші записи в журналі', 'аркуш «' + JOURNAL + '» або демо-місяць'],
    [modOn_('tg') ? tgConnected_() : null, 'Telegram-бот (ранковий підсумок, звіти)', '«⚙️ Налаштування → 📨 Підключити Telegram-бота»'],
    [modOn_('cal') ? !!dp.getProperty('cal_stat') : null, 'Календар платежів', 'увімкнено модулем; оновлюється сам'],
    [modOn_('fop') ? fops.length > 0 : null, 'ФОП: група, ставки', 'аркуш «' + FOP_SHEET + '»'],
    [modOn_('debts') ? debtsList_(ss).length > 0 : null, 'Борги й розстрочки', 'аркуш «' + DEBTS_SHEET + '»'],
    [modOn_('advisor') ? !!dp.getProperty('adv_doc_id') : null, 'Радник Gemini: знімок і навичка', 'меню «🤖 Радник Gemini»'],
    [cs.none ? null : cs.bad.length === 0, '«' + CHECK_SHEET + '»: усе OK', 'аркуш «' + CHECK_SHEET + '»'],
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

/** «🚀 Почати»: добудувати таблицю (у копії й порожній), потім запитання. */
function startWizard() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  try { guard_(ss, true); } catch (e) { show_('🚀 Почати', ['⚠️ ' + esc_(e.message)]); return; }
  let built;
  try { built = schemaRun_(ss, { allowEmpty: true }); } catch (e) { show_('🚀 Почати', ['⚠️ ' + esc_(e.message)]); return; }
  const ask = (title, text, def) => {
    const a = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
    if (a.getSelectedButton() !== ui.Button.OK) return null;
    const v = String(a.getResponseText() || '').trim();
    return v === '' ? def : v;
  };
  const set = ss.getSheetByName(SETTINGS_SHEET), out = [], wiz = {};
  // 1. режим і учасники
  const mode = ask('🚀 1/6 · Хто веде облік', '1: одна людина\n2: сім\'я (до 5 учасників)\n\nНомер (порожньо: 1):', '1');
  if (mode === null) return;
  if (mode === '2') {
    const names = ask('🚀 1/6 · Учасники', 'Імена учасників через кому, до 5 (наприклад: Олена, Іван).\nТак вони будуть у колонці «Хто».', '');
    if (names === null) return;
    const list = names.split(',').map(s => s.trim()).filter(Boolean).slice(0, SET_MEMBERS_N);
    set.getRange(SET.mode).setValue(MODE_FAMILY);
    set.getRange(SET_MEMBERS_FIRST, 2, SET_MEMBERS_N, 1).setValues(Array.from({ length: SET_MEMBERS_N }, (x, i) => [list[i] || '']));
    out.push('Режим: сім\'я (' + esc_(list.join(', ') || 'учасників впишіть на «' + SETTINGS_SHEET + '»') + ').');
  } else { set.getRange(SET.mode).setValue(MODE_SOLO); out.push('Режим: одна людина.'); }
  // 2. картка
  const card = ask('🚀 2/6 · Картка', 'Основна картка:\n1: кредитна з лімітом (рахуємо «мої» окремо від ліміту)\n2: дебетова\n\nНомер (порожньо: 1):', '1');
  if (card === null) return;
  set.getRange(SET.card).setValue(card === '2' ? CARD_DEBIT : CARD_CREDIT);
  out.push('Картка: ' + (card === '2' ? CARD_DEBIT : CARD_CREDIT) + '.');
  // 3. доходи
  const bday = ask('🚀 3/6 · Доходи', 'З якого числа дохід (зарплата, ФОП) іде в бюджет наступного місяця?\n' +
    'Наприклад, 25: зарплата 25-го наповнює наступний місяць. 1: дохід завжди в поточний місяць.\n\nЧисло (порожньо: 25):', '25');
  if (bday === null) return;
  set.getRange(SET.bday).setValue(Math.max(1, Math.min(31, Number(bday) || 25)));
  const fopAns = ask('🚀 3/6 · ФОП', 'Є ФОП у сім\'ї? Введіть «ім\'я:група» через кому, наприклад «Олена:3, Іван:2».\n' +
    'Ім\'я: як у колонці «Хто». Порожньо: ФОП немає.', '');
  if (fopAns === null) return;
  const fops = fopAns.split(',').map(s => s.split(':').map(x => x.trim())).filter(x => x[0] && ['1', '2', '3'].indexOf(x[1]) >= 0).slice(0, FOP_LAST - FOP_FIRST + 1);
  wiz.fop = fops.length > 0;
  if (fops.length) {
    const fsh = ss.getSheetByName(FOP_SHEET);
    fsh.getRange(FOP_FIRST, 1, fops.length, 5).setValues(fops.map(f => [f[0], f[1], '', 'так', 'так']));
    out.push('ФОП: ' + esc_(fops.map(f => f[0] + ' (' + f[1] + ' група)').join(', ')) + '. Свою ставку ЄП громади можна вписати на «' + FOP_SHEET + '».');
  }
  out.push('Дохід з ' + set.getRange(SET.bday).getValue() + '-го числа: у наступний місяць.');
  // 4. категорії
  const keep = ask('🚀 4/6 · Категорії', 'Які набори категорій лишити? ' + CAT_SETS.map((c, i) => (i + 1) + ') ' + c[0]).join(', ') +
    '.\nНомери через кому; порожньо: усі. Інші категорії лишаються завжди.', '');
  if (keep === null) return;
  if (keep) {
    const nums = keep.split(/[,\s]+/).map(Number);
    const drop = CAT_SETS.filter((c, i) => nums.indexOf(i + 1) < 0 && !(c[0] === 'ФОП і податки' && fops.length)).reduce((a, c) => a.concat(c[1]), []);
    const removed = dictRemove_(ss, drop);
    if (removed.length) out.push('Прибрано категорії: ' + esc_(removed.join(', ')) + '.');
  }
  // 5. модулі й підключення
  const rec = { tg: true, credit: card !== '2', debts: false, budgets: true, cal: false, backup: true, fop: fops.length > 0, advisor: false };
  const modsAns = ask('🚀 5/6 · Модулі', 'Що ввімкнути? ' + MODS.map((x, i) => (i + 1) + ') ' + x.name).join('  ') +
    '\nНомери через кому; порожньо: рекомендовані (' + MODS.map((x, i) => (rec[x.id] ? i + 1 : 0)).filter(Boolean).join(', ') + ').\n' +
    'Змінити потім: аркуш «' + START_SHEET + '», блок «🧩 Модулі» (так / ні).', '');
  if (modsAns === null) return;
  const picked = modsAns ? modsAns.split(/[,\s]+/).map(Number).filter(x => x >= 1 && x <= MODS.length) : null;
  const st = ss.getSheetByName(START_SHEET);
  MODS.forEach((x, i) => { wiz[x.id] = picked ? picked.indexOf(i + 1) >= 0 : rec[x.id]; });
  if (st && startHasModules_(ss)) st.getRange(MOD_FIRST, 2, MODS.length, 1).setValues(MODS.map(x => [wiz[x.id] ? MOD_YES : MOD_NO]));
  out.push('Модулі: ' + MODS.filter(x => wiz[x.id]).map(x => x.name).join(', ') + '.');
  const conn = ask('🚀 5/6 · Що підключити зараз', '1) Mono  2) ПриватБанк для ФОП («Автоклієнт»)' + (wiz.tg ? '  3) Telegram-бот' : '') +
    '\nНомери через кому; порожньо: нічого (можна пізніше в «⚙️ Налаштування»).', '');
  if (conn === null) return;
  const m = conn.split(/[,\s]+/).map(Number);
  // 6. демо
  const journalEmpty = !readJournal_(ss).some(o => !o.blank);
  const demo = journalEmpty ? ask('🚀 6/6 · Демо-місяць', 'Додати вигадані операції за минулий і цей місяць, щоб одразу побачити «Головне» й графіки?\n' +
    'Прибрати потім: «🧹 Очистити демо». (так / ні; порожньо: так)', 'так') : 'ні';
  if (demo === null) return;
  const dp = PropertiesService.getDocumentProperties();
  dp.setProperties({ wiz: JSON.stringify(wiz), wiz_done: new Date().toISOString() });
  if (/^т/i.test(demo)) { const d = demoAdd_(ss); out.push('Демо: ' + d.rows + ' операцій ✓ (коментар «[демо]»).'); }
  let modLines = [];
  try { modLines = modulesApply_(ss, true).filter(l => l.indexOf('⚠️') >= 0); } catch (e) { modLines = ['⚠️ Модулі: ' + esc_(e && e.message || e)]; }
  startRefresh_(ss);
  const wantConn = m.indexOf(1) >= 0 || m.indexOf(2) >= 0 || (wiz.tg && m.indexOf(3) >= 0);
  show_('🚀 Готово', (built.ran ? ['Таблицю добудовано ✓.'] : []).concat(out, modLines, ['Далі: аркуш «' + START_SHEET + '» (чекліст і модулі).',
    wantConn ? 'Зараз відкриються вікна підключення.' : '']));
  if (m.indexOf(1) >= 0) connectMono();
  if (m.indexOf(2) >= 0) connectPrivat();
  if (wiz.tg && m.indexOf(3) >= 0) connectTelegram();
}

/** Прибрати категорії з «Довідника», якщо їх немає в журналі. Правила з цими категоріями теж. → прибрані назви. */
function dictRemove_(ss, names) {
  const d = ss.getSheetByName(DICT_SHEET), used = new Set(readJournal_(ss).map(o => o.c)), out = [];
  const v = d.getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 7).getValues();
  v.forEach((r, i) => {
    const n = String(r[0]).trim();
    if (names.indexOf(n) < 0 || used.has(n)) return;
    d.getRange(D_FIRST + i, 1, 1, 7).setValues([['', '', '', '', '', '', '']]);
    out.push(n);
  });
  const rs = ss.getSheetByName(RULES_SHEET);
  if (rs && out.length) rs.getRange(RULES_FIRST, 1, RULES_LAST - RULES_FIRST + 1, 7).getValues().forEach((r, i) => {
    if (out.indexOf(String(r[4]).trim()) >= 0) rs.getRange(RULES_FIRST + i, 1, 1, 7).setValues([['', '', '', '', '', '', '']]);
  });
  return out;
}

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
  const next = addDays_(lastDayOf_(ym), 1).slice(0, 7);
  const plans = [[D(next + '-05'), 2400, 'Комуналка', DEMO_TAG + 'комунальні послуги', STATUS_PLAN, '', '', ''],
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
  dp.setProperty('demo', JSON.stringify(state));
  sortCore_(ss);
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
  if (state.start) ss.getSheetByName(SETTINGS_SHEET).getRange(SET.start).setValue(Utilities.parseDate(state.start, tz, 'yyyy-MM-dd'));
  dp.deleteProperty('demo');
  sortCore_(ss);
  return n;
}

function demoAdd() {
  if (!menuOk_('🎬 Демо-місяць')) return;
  const ss = SpreadsheetApp.getActive();
  if (jsonProp_(PropertiesService.getDocumentProperties(), 'demo')) { show_('🎬 Демо-місяць', ['Демо вже додано. Прибрати: «🧹 Очистити демо».']); return; }
  const r = demoAdd_(ss);
  show_('🎬 Демо-місяць', ['Додано ' + r.rows + ' вигаданих операцій (коментар «[демо]»). Гляньте «' + MAIN_SHEET + '», «' + DAYS_SHEET + '», «' + MONTHS_SHEET + '».',
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
