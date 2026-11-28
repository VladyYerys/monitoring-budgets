/** 90_Menu.gs: Меню і збирання шаблону. */

/*
 * У таблиці з аркушами «m» і «y» меню не з'являється взагалі. У непозначеній таблиці: лише «🚀 Почати» і «🩺 Стан шаблону».
 * У шаблоні та його копіях: повне меню. Кожен пункт ще раз перевіряє таблицю (01_Guard.gs).
 */

function onOpen() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  if (isPersonalModel_(ss)) return;
  let ours = false;
  try { ours = guard_(ss) === true; } catch (e) { ours = false; }
  if (!ours) {
    ui.createMenu('📊 ' + APP_NAME).addItem('🚀 Почати (у порожній таблиці)', 'startWizard').addItem('🩺 Стан шаблону', 'templateStatus').addToUi();
    return;
  }
  try { schemaToast_(ss); } catch (e) { /* підказка не критична */ }
  try { if (!PropertiesService.getDocumentProperties().getProperty('wiz_done')) ss.toast('Почніть з «📊 ' + APP_NAME + ' → 🚀 Підключення».', '🚀 Старт', 10); else supportToast_(ss); } catch (e) { /* ignore */ }
  ui.createMenu('📊 ' + APP_NAME)
    .addItem('🚀 Підключення', 'startWizard')
    .addItem('🔃 Упорядкувати журнал зараз', 'sortNow')
    .addItem('🏦 Імпорт з банку зараз', 'importNow')
    .addItem('🗓 Імпорт з банку за період…', 'monoPeriodNow')
    .addItem('📥 Імпорт виписок з папки', 'importStatements')
    .addItem('🔎 Сирі дані Mono', 'monoRawNow')
    .addItem('✅ Підтвердити чернетки з категорією', 'confirmDrafts')
    .addItem('🔁 Регулярні: оновити плани зараз', 'regularNow')
    .addItem('📚 Записати місяць в «Історію»…', 'historyMonthNow')
    .addItem('💾 Копія зараз', 'backupNow')
    .addItem('🧹 Зачистити дані…', 'wipeNow')
    .addSubMenu(ui.createMenu('🎬 Демо')
      .addItem('🎬 Додати демо-місяць', 'demoAdd')
      .addItem('🧹 Очистити демо', 'demoClear'))
    .addSubMenu(ui.createMenu('📨 Telegram зараз')
      .addItem('🌅 Ранковий підсумок', 'telegramTestNow')
      .addItem('📊 Тижневий звіт', 'weeklyNow')
      .addItem('📅 Підсумок місяця…', 'monthNow'))
    .addSubMenu(ui.createMenu('🤖 Радник Gemini')
      .addItem('🤖 Як підключити навичку', 'advisorNow')
      .addItem('🔄 Оновити знімок зараз', 'advisorSnapshotNow'))
    .addSeparator()
    .addItem('🩺 Стан таблиці', 'tableStatus')
    .addItem('🧩 Модулі', 'modulesNow')
    .addItem('⚡ Оновити «Швидке підключення»', 'startNow')
    .addSubMenu(ui.createMenu('🛠 Оновлення таблиці')
      .addItem('🔎 Чи вийшла нова версія шаблону', 'versionNow')
      .addItem('📦 Перенести дані зі старої копії', 'migrateFromOld')
      .addItem('⬆️ Оновити зараз (зазвичай відбувається само)', 'schemaUpdateNow'))
    .addItem('❤️ Підтримати автора', 'supportAuthor')
    .addToUi();
}

/** Порожня таблиця → шаблон (усі кроки). У копії шаблону: лише те, чого бракує (розклади не копіюються разом з файлом). */
function setupTemplate() {
  const ss = SpreadsheetApp.getActive();
  try { guard_(ss, true); } catch (e) { show_('🚀 Шаблон', ['⚠️ ' + esc_(e.message)]); return; }
  let r;
  try { r = schemaRun_(ss, { allowEmpty: true }); } catch (e) { show_('🚀 Шаблон', ['⚠️ ' + esc_(e.message)]); return; }
  try { const j = ss.getSheetByName(JOURNAL); if (j && ss.setActiveSheet) ss.setActiveSheet(j); } catch (e) { /* ignore */ }
  show_('🚀 Шаблон готовий', r.lines.concat(['Далі: «🚀 Підключення» (картка, Telegram, регулярні, Gemini, модулі); потім пишіть у «' + JOURNAL + '», а дивіться на «' +
    DASH_SHEET + '» і «' + MOB_SHEET + '». «' + SETTINGS_SHEET + '», «' + CP_SHEET + '» та інші аркуші працюють схованими: показати їх можна модулем «' +
    modName_('service') + '» (' + MOD_WHERE + ').']));
}

/** При відкритті: одна підказка, якщо оновлення зупинилось або чекає. */
function schemaToast_(ss) {
  const dp = PropertiesService.getDocumentProperties(), fail = jsonProp_(dp, 'schema_fail'), cur = Number(dp.getProperty('schema_v') || 0);
  if (fail && fail.to === schemaLast_()) { ss.toast(String(fail.msg || '').slice(0, 180) + ' Деталі: «🩺 Стан таблиці».', '⚠️ Оновлення таблиці зупинилось', 20); return; }
  if (cur < schemaLast_()) ss.toast('Нова версія таблиці (' + schemaLast_() + ') застосується сама протягом 15 хв, з копією перед змінами. Або: «🛠 Оновлення таблиці → ⬆️ Оновити зараз».',
    '🛠 Оновлення таблиці', 10);
}

function historyMonthNow() {
  if (!menuOk_('📚 Історія')) return;
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi(), prev = prevYm_(todayKey_(ss).slice(0, 7));
  const a = ui.prompt('📚 Записати місяць', 'Який місяць записати (рррр-мм)? Порожньо: ' + prev + '.', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const ym = String(a.getResponseText() || '').trim() || prev;
  if (!/^\d{4}-\d{2}$/.test(ym)) { show_('📚 Історія', ['⚠️ Потрібно рррр-мм, наприклад ' + prev + '.']); return; }
  const r = historyWrite_(ss, ym);
  show_('📚 Історія', ['Записано ' + ym + ' ✓ (категорій: ' + r.cats + ').']);
}

/**
 * Одне нагадування «❤️ Підтримати автора» після першого повного місяця (з'явився перший рядок «Історії місяців»).
 * Більше не показується; вимкнути зовсім: «Налаштування», B24 «ні».
 */
function supportToast_(ss) {
  const dp = PropertiesService.getDocumentProperties(), set = ss.getSheetByName(SETTINGS_SHEET), h = ss.getSheetByName(HISTORY_SHEET);
  if (dp.getProperty('sup_toast') || !h || h.getLastRow() < H_FIRST) return false;
  if (set && String(set.getRange(SET.support).getValue()).trim() === 'ні') return false;
  dp.setProperty('sup_toast', new Date().toISOString());
  ss.toast('Перший місяць в «Історії» ✓. Шаблон безкоштовний; якщо він вам корисний: меню «❤️ Підтримати автора». Це нагадування більше не з\'явиться.', '❤️ Дякую!', 15);
  return true;
}

function templateStatus() {
  const ss = SpreadsheetApp.getActive();
  let ok = false, why = '';
  try { ok = guard_(ss) === true; } catch (e) { why = e.message; }
  show_('🩺 Стан шаблону', [ok ? 'Це таблиця-шаблон ' + APP_NAME + ' ✓' : '⚠️ ' + esc_(why || 'Таблиця не позначена як шаблон.'),
    'Версія коду: ' + TEMPLATE_VERSION + '. Код відкритий: ' + link_(REPO_URL, 'github.com/VladyYerys/monitoring-budgets') + ' (ліцензія MIT).']);
}

function supportAuthor() {
  show_('❤️ Підтримати автора', ['Шаблон безкоштовний і з відкритим кодом. Якщо він вам корисний, можна підтримати розвиток:',
    link_(SUPPORT_URL, '☕ Банка Mono «підтримка проєкту»'), 'Дякую! Жодних функцій «за донат» немає і не буде.',
    'Питання й ідеї: ' + link_(REPO_URL + '/issues', 'розділ Issues на GitHub') + '.']);
}
