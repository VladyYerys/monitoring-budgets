/** 11_Dictionary.gs: «Довідник» (категорії, типи, групи, «чорні діри», причини з ліміту) і «Налаштування». */

/*
 * Довідник, колонки:
 *   A категорія · B тип (Витрати, Заощадження, Податки, Борги, Надходження, Резерв, Корекція +, Корекція −)
 *   C група (як рядки підсумків на «Місяцях») · D непомітна («чорна діра», якщо вся категорія така)
 *   E мої гроші: «надходження» (дохід, що наповнює «мої» гроші місяця) або «відкладення» (відкладено на погашення ліміту)
 *   F план на місяць, ₴ · G чий бюджет (порожньо: спільний; ім'я учасника: особистий)
 *   I «чорні діри» (список) · J що туди входить · K причини, чому з ліміту (список)
 * Стартовий набір нейтральний: людина змінює його під себе (назви можна перейменовувати, формули тримаються).
 */

const DICT_START = [
  // [категорія, тип, група, непомітна, мої гроші]
  ['Продукти', T_SPEND, '🍎 Їжа'], ['Кафе й ресторани', T_SPEND, '🍎 Їжа'], ['Доставка їжі', T_SPEND, '🍎 Їжа', '🛵 Доставка'],
  ['Кава та перекуси', T_SPEND, '🍎 Їжа', '☕ Кава та перекуси'],
  ['Комуналка', T_SPEND, '🏠 Дім'], ['Оренда житла', T_SPEND, '🏠 Дім'], ['Дім і побут', T_SPEND, '🏠 Дім'], ['Інтернет і зв\'язок', T_SPEND, '🏠 Дім'],
  ['Транспорт', T_SPEND, '🚗 Транспорт'], ['Пальне', T_SPEND, '🚗 Транспорт'], ['Таксі', T_SPEND, '🚗 Транспорт', '🚕 Таксі'],
  ['Здоров\'я й аптека', T_SPEND, '💊 Здоров\'я й догляд'], ['Краса й догляд', T_SPEND, '💊 Здоров\'я й догляд'], ['Спорт', T_SPEND, '💊 Здоров\'я й догляд'],
  ['Діти', T_SPEND, '👨‍👩‍👧 Діти й тварини'], ['Тварини', T_SPEND, '👨‍👩‍👧 Діти й тварини'],
  ['Одяг і взуття', T_SPEND, '🛍 Покупки'], ['Техніка', T_SPEND, '🛍 Покупки'], ['Дрібні покупки', T_SPEND, '🛍 Покупки', '🛒 Дрібні імпульсивні покупки'],
  ['Розваги', T_SPEND, '🎉 Розваги й подарунки'], ['Подарунки', T_SPEND, '🎉 Розваги й подарунки'], ['Подорожі', T_SPEND, '🎉 Розваги й подарунки'],
  ['Підписки', T_SPEND, '📱 Підписки й навчання', '🔁 Підписки'], ['Навчання', T_SPEND, '📱 Підписки й навчання'],
  ['Благодійність', T_SPEND, '🤝 Благодійність'], ['Інші витрати', T_SPEND, '📦 Інше'],
  ['Заощадження', T_SAVE, '💰 Заощадження'], ['Резерв на податки', T_SAVE, '💰 Заощадження'],
  ['Єдиний податок', T_TAX, '🧾 ФОП і податки'], ['ЄСВ', T_TAX, '🧾 ФОП і податки'], ['Військовий збір', T_TAX, '🧾 ФОП і податки'],
  ['Кредит / розстрочка', T_DEBT, '💳 Борги'], ['Повернення боргу', T_DEBT, '💳 Борги'], ['Комісії й відсотки банку', T_DEBT, '💳 Борги', '🏦 Банківські комісії'],
  ['Зарплата', T_IN, '💵 Надходження', '', MINE_IN], ['Дохід ФОП', T_IN, '💵 Надходження', '', MINE_IN],
  ['Мені повернули борг', T_IN, '💵 Надходження', '', MINE_IN], ['Інші надходження', T_IN, '💵 Надходження', '', MINE_IN],
  ['Переказ з власного рахунку', T_IN, '💵 Надходження', '', MINE_IN], ['Отримано в борг / кредит', T_IN, '💵 Надходження'],
  ['Нерозподілені надходження (до розбору)', T_CORR_IN, '❓ До розбору'], ['Нерозподілені списання (до розбору)', T_CORR_OUT, '❓ До розбору'],
  ['🎯 На погашення ліміту', T_RESERVE, '🎯 Погашення ліміту', '', MINE_SET_ASIDE],
  [CAT_TRANSFER, '', '↔ Перекази'], [CAT_OWN_MOVE, '', '↔ Перекази'],
];
const HOLES_START = [
  ['☕ Кава та перекуси', 'кава, снеки, дрібні перекуси дорогою'], ['🛵 Доставка', 'доставка їжі разом зі збором і чайовими'],
  ['🔁 Підписки', 'сервіси, що списуються щомісяця'], ['🚕 Таксі', 'поїздки за звичкою'],
  ['🛒 Дрібні імпульсивні покупки', 'дрібниці (маркетплейси тощо)'], ['🏦 Банківські комісії', 'комісії, плата за ліміт і картку'],
];
const REASONS_START = ['🚨 Непередбачене / терміново', '📅 Не вистачило до зарплати', '🛒 Імпульс / не втримався', '📦 Планова велика покупка',
  '🏥 Здоров\'я', '🛠️ Ремонт / поломка', '🎁 Подарунки / свята', '✈️ Відпочинок', '💳 Платіж за кредитом / розстрочкою', 'Інше'];

function buildDictionary_(ss) {
  const sh = sheetOf_(ss, DICT_SHEET);
  sh.getRange(1, 1).setValue('📚 Довідник').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Категорії: назва, тип, група, «чорна діра», чи наповнює «мої гроші», план на місяць. Змінюйте під себе.').setFontColor('#5f6368');
  sh.getRange(4, 1, 1, 7).setValues([['Категорія', 'Тип', 'Група', 'Непомітна («чорна діра»)', 'Мої гроші', 'План на місяць, ₴', 'Чий бюджет']]).setFontWeight('bold');
  sh.getRange(4, 9, 1, 3).setValues([['«Чорні діри»', 'Що сюди входить', 'Чому з ліміту']]).setFontWeight('bold');
  const n = D_LAST - D_FIRST + 1;
  const rows = [];
  for (let i = 0; i < n; i++) { const x = DICT_START[i] || []; rows.push([x[0] || '', x[1] || '', x[2] || '', x[3] || '', x[4] || '', '']); }
  sh.getRange(D_FIRST, 1, n, 6).setValues(rows);
  const nh = D_HOLES_LAST - D_HOLES_FIRST + 1, holes = [];
  for (let i = 0; i < nh; i++) holes.push(HOLES_START[i] || ['', '']);
  sh.getRange(D_HOLES_FIRST, 9, nh, 2).setValues(holes);
  const nr = D_REASONS_LAST - D_REASONS_FIRST + 1, rs = [];
  for (let i = 0; i < nr; i++) rs.push([REASONS_START[i] || '']);
  sh.getRange(D_REASONS_FIRST, 11, nr, 1).setValues(rs);
  sh.getRange('B' + D_FIRST + ':B' + D_LAST).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TYPES).setAllowInvalid(false).build());
  sh.getRange('E' + D_FIRST + ':E' + D_LAST).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([MINE_IN, MINE_SET_ASIDE]).setAllowInvalid(false).build());
  sh.getRange('D' + D_FIRST + ':D' + D_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getRange('I' + D_HOLES_FIRST + ':I' + D_HOLES_LAST)).setAllowInvalid(true).build());
  sh.getRange('F' + D_FIRST + ':F' + D_LAST).setNumberFormat('#,##0');
  dictWhoRule_(ss, sh);
  sh.setFrozenRows(4);
  return sh;
}

/** «Чий бюджет»: список учасників з «Налаштувань». */
function dictWhoRule_(ss, sh) {
  const set = ss.getSheetByName(SETTINGS_SHEET);
  if (set) sh.getRange('G' + D_FIRST + ':G' + D_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(set.getRange('B' + SET_MEMBERS_FIRST + ':B' + (SET_MEMBERS_FIRST + SET_MEMBERS_N - 1))).setAllowInvalid(true).build());
}

/** Оновлення наявного «Довідника» (версія 5): колонка «Чий бюджет» і нові службові категорії; ваші рядки не змінюються. */
function dictUpgrade_(ss) {
  const sh = ss.getSheetByName(DICT_SHEET), out = [];
  if (!sh) return out;
  if (String(sh.getRange('G4').getValue()) !== 'Чий бюджет') { sh.getRange('G4').setValue('Чий бюджет').setFontWeight('bold'); dictWhoRule_(ss, sh); out.push('колонка «Чий бюджет»'); }
  const names = sh.getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 1).getValues().map(r => String(r[0]).trim());
  [[CAT_OWN_MOVE, '', '↔ Перекази']].forEach(c => {
    if (names.indexOf(c[0]) >= 0) return;
    const free = names.findIndex(n => !n);
    if (free < 0) return;
    sh.getRange(D_FIRST + free, 1, 1, 3).setValues([c]); names[free] = c[0]; out.push('категорія «' + c[0] + '»');
  });
  return out;
}

function buildSettings_(ss) {
  const sh = sheetOf_(ss, SETTINGS_SHEET), tz = ss.getSpreadsheetTimeZone();
  const now = new Date(), first = Utilities.parseDate(Utilities.formatDate(now, tz, 'yyyy-MM') + '-01', tz, 'yyyy-MM-dd');
  sh.getRange(1, 1).setValue('⚙️ Налаштування').setFontWeight('bold').setFontSize(14);
  sh.getRange(3, 1, 1, 2).setValues([['Що', 'Значення']]).setFontWeight('bold');
  sh.getRange(4, 1, 6, 2).setValues([
    ['Старт обліку «моїх грошей» (перше число місяця)', first],
    ['Дохід з цього числа йде в бюджет наступного місяця', 25],
    ['Орієнтир погашення ліміту на місяць, ₴', 0],
    ['Картка (кредитна з лімітом / дебетова)', CARD_CREDIT],
    ['Тривога в Telegram, якщо прогноз «доступно з лімітом» нижчий за, ₴', 5000],
    ['Шаблон: версія', TEMPLATE_VERSION],
  ]);
  settingsRate_(sh);
  sh.getRange(SET.start).setNumberFormat('dd.mm.yyyy');
  sh.getRange(SET.card).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([CARD_CREDIT, CARD_DEBIT]).setAllowInvalid(false).build());
  sh.getRange(11, 1).setValue('Токени банку й Telegram тут не зберігаються: лише в налаштуваннях скрипта у вашому Google-акаунті (меню «⚙️ Налаштування»).')
    .setFontColor('#5f6368');
  settingsFamily_(sh);
  settingsExtra_(sh);
  const r = Number(SET_MARK_CELL.slice(1));
  sh.getRange(SET_MARK_CELL).setValue(TEMPLATE_MARK).setFontColor('#ffffff');   // мітка шаблону (див. 01_Guard.gs)
  sh.hideRows(r, 1);
  return sh;
}

/** Блок «Сім'я» на «Налаштуваннях»: режим, до 5 учасників, «спільне» і «уся сім'я» (службові пункти списків). */
function settingsFamily_(sh) {
  sh.getRange(12, 1).setValue('Сім\'я').setFontWeight('bold');
  sh.getRange(13, 1, 1, 2).setValues([['Режим (одна людина / сім\'я)', MODE_SOLO]]);
  sh.getRange(SET.mode).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([MODE_SOLO, MODE_FAMILY]).setAllowInvalid(false).build());
  const rows = [];
  for (let i = 0; i < SET_MEMBERS_N; i++) rows.push(['Учасник ' + (i + 1) + ' (як у колонці «Хто»)', '']);
  sh.getRange(SET_MEMBERS_FIRST, 1, SET_MEMBERS_N, 2).setValues(rows);
  sh.getRange(SET_SHARED_ROW, 1, 2, 2).setValues([['Спільне (для колонки «Хто»)', WHO_SHARED], ['Уся сім\'я (для «Місяців»)', WHO_ALL]]);
  sh.getRange(SET_SHARED_ROW, 1, 2, 2).setFontColor('#5f6368');
}

/** Рядок 10: ставка комісії за переказ з кредитного ліміту (для грейсу, модуль «💳 Кредитний ліміт і грейс»). */
function settingsRate_(sh) {
  sh.getRange(10, 1, 1, 2).setValues([['Комісія за переказ з кредитного ліміту (грейс; 0,035 = 3,5 %)', 0.035]]);
  sh.getRange(SET.rate).setNumberFormat('0.0%');
}

const ADV_LEVELS = ['підсумки', 'категорії', 'операції'];
/** Рядки 22…24: радник Gemini (рівень деталей знімка) і нагадування «Підтримати автора». */
function settingsExtra_(sh) {
  sh.getRange(22, 1).setValue('Радник і нагадування').setFontWeight('bold');
  sh.getRange(23, 1, 2, 2).setValues([['Знімок для радника Gemini: рівень деталей (підсумки / категорії / операції без коментарів)', ADV_LEVELS[1]],
    ['Нагадування «❤️ Підтримати автора» (раз після першого місяця й рядок у підсумку місяця раз на квартал)', 'так']]);
  sh.getRange(SET.advisor).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ADV_LEVELS).setAllowInvalid(false).build());
  sh.getRange(SET.support).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['так', 'ні']).setAllowInvalid(false).build());
}

/** Оновлення наявних «Налаштувань» (версія 6): ставка комісії, радник, нагадування. Ваші значення не змінюються. */
function settingsUpgrade6_(ss) {
  const sh = ss.getSheetByName(SETTINGS_SHEET), out = [];
  if (!sh) return out;
  if (String(sh.getRange('A10').getValue()) === '') { settingsRate_(sh); out.push('ставка комісії (грейс)'); }
  if (String(sh.getRange('A22').getValue()) !== 'Радник і нагадування') { settingsExtra_(sh); out.push('радник і нагадування'); }
  return out;
}

/** Оновлення наявних «Налаштувань» (версія 5): блок «Сім'я», якщо його ще немає. */
function settingsUpgrade_(ss) {
  const sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh || String(sh.getRange('A12').getValue()) === 'Сім\'я') return [];
  settingsFamily_(sh);
  return ['блок «Сім\'я» на «Налаштуваннях»'];
}

/** Учасники з «Налаштувань» (без порожніх). У режимі «одна людина»: порожньо. */
function members_(ss) {
  const sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) return [];
  return sh.getRange(SET_MEMBERS_FIRST, 2, SET_MEMBERS_N, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
}
function familyMode_(ss) { const sh = ss.getSheetByName(SETTINGS_SHEET); return !!sh && String(sh.getRange(SET.mode).getValue()) === MODE_FAMILY; }

function buildCheckpoints_(ss) {
  const sh = ensureSize_(sheetOf_(ss, CP_SHEET), CP_LAST, 3);
  sh.getRange(1, 1).setValue('🏦 Залишки банку').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Контрольні точки: залишок на картці на початок дня (мінус, якщо в кредиті) і кредитний ліміт. Пише імпорт з банку; можна й вручну.')
    .setFontColor('#5f6368');
  sh.getRange(4, 1, 1, 3).setValues([['Дата', 'Залишок на початок дня, ₴', 'Кредитний ліміт, ₴']]).setFontWeight('bold');
  sh.getRange('A' + CP_FIRST + ':A' + CP_LAST).setNumberFormat('dd.mm.yyyy');
  sh.getRange('B' + CP_FIRST + ':C' + CP_LAST).setNumberFormat('#,##0.00');
  sh.setFrozenRows(4);
  return sh;
}
