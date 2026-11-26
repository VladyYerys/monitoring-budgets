/** 39_Advisor.gs: Радник на Gemini (модуль «🤖 Радник Gemini»): знімок фінансів на Диску й готова навичка (skill) Gemini. */

/*
 * Знімок: Google Документ «Monitoring Budgets: знімок для радника» у папці «Monitoring Budgets: радник» на ВАШОМУ Диску.
 * Оновлюється сам щодня після 05:00 (і з меню). Що в ньому: головні цифри, «Мої правила», місяці, бюджети, плани на 2 місяці,
 * непомітні, підписки, борги, ліміт і грейс, ФОП, сім'я по учасниках; словничок; «Де що в таблиці». Рівень деталей: «Налаштування», B23.
 * Чого в ньому немає ніколи: ваших коментарів, токенів, номерів карток і рахунків, ключів рахунків. Описи з банку: лише для
 * категорії «Підписки» (назва сервісу), довгі числа в них замінено на «…». У «Моїх правилах» довгі числа й ключі теж стають «…».
 * Навичка Gemini (skills у застосунку Gemini): одне повідомлення для чату («Створи нову навичку…»), Gemini створює її сам.
 * У скопійованому тексті пряме посилання на ваш знімок (у SKILL.md репозиторію його немає: там пошук за назвою).
 * Роль, українська, лише цифри знімка з датою й розділом, сценарії типових питань, без інвестиційних порад.
 * «Мої правила» (Налаштування, B27…B32) живуть у таблиці й ідуть у знімок: навичку створюють один раз.
 * Версія навички (ADV_SKILL.v) є і в навичці, і в знімку: стара навичка сама підкаже оновитись.
 * Скрипт сам нікуди знімок не надсилає: до Gemini він потрапляє, лише коли ви самі питаєте навичку.
 */

const ADV_FOLDER = APP_NAME + ': радник';
const ADV_DOC = APP_NAME + ': знімок для радника';
const ADV_HOUR = 5, ADV_RETRY_MS = 3600 * 1000, ADV_OPS_DAYS = 60;
const ADV_SKILL = {
  v: 2,                                                                  // версія навички: знімок пише її, навичка порівнює зі своєю
  name: 'monitoring-budgets-advisor',
  title: 'Радник Monitoring Budgets',
  description: 'Сімейний фінансовий радник українською за знімком таблиці Monitoring Budgets з Google Диска. Використовуй, коли я питаю про наш бюджет і витрати, ' +
    'скільки можна витратити сьогодні чи до зарплати, чи вистачить грошей, як розподілити зарплату чи інший дохід, кредитний ліміт і грейс, підписки, борги й розстрочки, ' +
    'ФОП і податки, де ми переплачуємо.',
  starters: ['Скільки можна витратити сьогодні?', 'Чи вистачить до зарплати, якщо купити подарунки на 5 000 ₴?', 'Прийшла зарплата 40 000 ₴, як її розподілити?',
    'Де ми переплачуємо за останні 3 місяці?', 'Які підписки варто скасувати?', 'Скільки відкласти на податки ФОП до кінця кварталу?', 'Як швидше погасити кредитний ліміт без комісій?'],
};

/**
 * Текст інструкції навички (без заголовка SKILL.md). Назви цифр і розділів: як у знімку (snapshotText_).
 * url: пряме посилання на ваш знімок (лише в тексті, який ви копіюєте з таблиці); без нього (SKILL.md у репозиторії) Gemini шукає документ за назвою.
 */
function advSkillInstructions_(url) {
  const v = ADV_SKILL.v;
  return [
    'Ти сімейний фінансовий помічник для таблиці «' + APP_NAME + '». Версія навички: ' + v + '.',
    '',
    '## 1. Звідки дані',
    url ? '1. Перед кожною відповіддю відкрий мій знімок «' + ADV_DOC + '» за посиланням ' + url + ' (мій Google Диск) і прочитай найновішу версію. ' +
      'Якщо посилання не відкривається, знайди документ з такою назвою на моєму Google Диску (папка «' + ADV_FOLDER + '»). Цифри бери лише з нього.'
      : '1. Перед кожною відповіддю знайди на моєму Google Диску документ «' + ADV_DOC + '» (папка «' + ADV_FOLDER + '») і прочитай найновішу версію. Цифри бери лише з нього.',
    '2. Перший рядок відповіді: «За знімком від ДД.ММ». Якщо знімку понад 3 дні, скажи про це й підкажи оновити: у таблиці меню «🤖 Радник Gemini» → «🔄 Оновити знімок зараз».',
    '3. Готові цифри не перераховуй: «Можна на день», «Доступно моїх сьогодні», прогнози, найнижча точка й грейс уже пораховані таблицею з планами й бюджетним місяцем. ' +
      'Сам рахуй лише те, чого в знімку немає (наприклад «а якщо відкласти 5 000»), і покажи розрахунок одним рядком.',
    '4. Біля кожної цифри в дужках назви розділ знімка: (Головне), (Бюджети), (Місяці), (Плани), (Підписки), (Борги), (ФОП).',
    '5. Терміни (мої гроші, бюджетний місяць, факт, план, чернетка, непомітні, грейс) розумій так, як у розділі «Як читати».',
    '6. Чого немає в знімку, не вигадуй: скажи, чого бракує і де це заповнити (розділ «Де що в таблиці»: аркуш і колонка).',
    '',
    '## 2. Мої правила',
    '- Розділ знімка «Мої правила»: мій пріоритет, подушка, запас на картці, розстрочки, ФОП і вільні правила. Враховуй їх у кожній пораді й не суперечь їм.',
    '- Якщо правила не заповнені, один раз порадь заповнити: у таблиці «🚀 Підключення», крок 4.',
    '',
    '## 3. Як відповідати',
    '- Українською, коротко, суми округлюй до гривні.',
    '- Спершу відповідь на питання з цифрами, потім 2–3 конкретні кроки: яку категорію зменшити й на скільки, що перенести, що відкласти. Без моралізаторства.',
    '- Питання загальне («як у нас з грошима?»): почни з трьох цифр: доступно моїх сьогодні, можна на день, прогноз на кінець місяця.',
    '- Є чернетки («Чернеток чекають підтвердження» більше 0): попередь, що цифри зміняться, коли їх підтвердити в «Журналі».',
    '',
    '## 4. Типові питання',
    '- «Скільки можна витратити сьогодні чи до зарплати?»: «Можна на день» і «Доступно моїх сьогодні» (Головне), скільки днів лишилось. Є запас у «Моїх правилах»: нагадай про нього.',
    '- «Чи вистачить на X?»: порівняй X з «Найнижча точка за планом» і її датою (Головне) мінус запас з «Моїх правил». Не вистачає: скажи, на скільки, і які плани (Плани) можна посунути чи зменшити.',
    '- «Прийшла зарплата чи інший дохід, як розподілити?»: 1) грейс: «Грейс: внести до строку» до «Грейс: строк» (Головне); 2) податки ФОП за планом (ФОП) і відсоток ФОП з «Моїх правил»; ' +
      '3) обов\'язкові плани до наступного доходу (Плани), з розстрочками й кредитами, якщо так у «Моїх правилах»; 4) подушка: відсоток з «Моїх правил»; 5) решта на життя: скільки це на день. ' +
      'Пріоритет №1 з «Моїх правил» став першим.',
    '- «Де ми переплачуємо?»: бюджети «перевищено» і «швидше норми» (Бюджети), найбільші непомітні й категорії порівняно з трьома попередніми місяцями (Місяці).',
    '- «Які підписки скасувати?»: розділ «Підписки»: що повторюється, скільки на місяць і на рік.',
    '- «Як швидше закрити ліміт чи борги?»: «Борг за лімітом зараз», «Орієнтир погашення на місяць», грейс і комісія (Головне), залишки й платежі (Борги).',
    '- «Додай витрату» чи «зміни таблицю»: чесно скажи, що навичка лише читає знімок і нічого в таблиці не змінює. Внести: новий рядок у «Журналі» (A дата, B сума, C категорія, E статус) ' +
      'з комп\'ютера чи в застосунку Google Таблиць.',
    '',
    '## 5. Межі й безпека',
    '- Не давай інвестиційних порад: не радь конкретних акцій, облігацій, криптовалют, фондів чи банків.',
    '- ФОП: суми й строки зі знімка; остаточні податкові рішення: з бухгалтером. Це не податкова консультація.',
    '- Ніколи не проси номери карток, IBAN, токени, паролі чи коди з SMS. Якщо їх надіслали, попроси видалити.',
    '- Сімейний режим: говори і про кожного учасника, і про сім\'ю разом; без осуду.',
    '- Ти бачиш лише знімок, не банк і не живу таблицю: не обіцяй більшого.',
    '- У знімку є рядок «Навичка: версія N». Якщо N більше за ' + v + ', один раз скажи, що навичку варто оновити: «🚀 Підключення», крок 4.',
    '',
    'Стартові питання (можна запропонувати, якщо питання немає):',
  ].concat(ADV_SKILL.starters.map(q => '- ' + q)).join('\n');
}

/** SKILL.md для імпорту в Gemini («Skills → Create → Upload») і для репозиторію (gemini/…/SKILL.md). */
function advSkillMd_() {
  return '---\nname: ' + ADV_SKILL.name + '\ndescription: ' + ADV_SKILL.description + '\n---\n\n# ' + ADV_SKILL.title + '\n\n' + advSkillInstructions_() + '\n';
}

/** Одне повідомлення для чату Gemini: Gemini сам створює з нього навичку (без полів «Назва», «Опис», «Інструкції»). url: пряме посилання на знімок. */
function advSkillMsg_(url) {
  return ['Створи нову навичку (skill) з тексту нижче. Назву, опис та інструкції збережи дослівно: нічого не скорочуй і не переписуй. Коли збережеш, коротко підтверди.',
    '', 'Назва: ' + ADV_SKILL.title, '', 'Опис: ' + ADV_SKILL.description, '', 'Інструкції:', '', advSkillInstructions_(url)].join('\n');
}

/* ---------- мої правила (Налаштування, B27…B32) ---------- */

const ADV_RULES_KEYS = ['prio', 'cushion', 'floor', 'debtsMust', 'fopPct', 'rules'];
const advPct_ = v => { const n = Math.max(0, Number(v) || 0); return n > 1 ? n / 100 : n; };   // «10» і «10 %» (0,1): обидва 10 %
/** Довгі числа (картки, IBAN, рахунки) і довгі ключі (токени) не потрапляють у знімок. */
function advMask_(s) { return String(s || '').replace(/(?:\d[ -]?){12,}/g, '…').replace(/[A-Za-z0-9_:\-]{24,}/g, '…'); }
/** Схоже на номер картки, IBAN чи токен: такий текст у таблицю не пишемо. */
function advSecretLike_(s) { return advMask_(s) !== String(s || ''); }

/** «Мої правила» з «Налаштувань». → { prio, cushion (частка), floor, debtsMust, fopPct (частка), rules, filled } */
function advRules_(ss) {
  const sh = ss.getSheetByName(SETTINGS_SHEET), ok = sh && String(sh.getRange('A26').getValue()) === ADV_RULES_HEAD;
  const v = ok ? sh.getRange(27, 2, 6, 1).getValues().map(r => r[0]) : [ADV_PRIO[0], 0, 0, 'так', 0, ''];
  const r = { prio: ADV_PRIO.indexOf(String(v[0]).trim()) >= 0 ? String(v[0]).trim() : ADV_PRIO[0], cushion: advPct_(v[1]), floor: Math.max(0, Math.round(Number(v[2]) || 0)),
    debtsMust: String(v[3]).trim() !== 'ні', fopPct: advPct_(v[4]), rules: String(v[5] || '').trim().slice(0, ADV_RULES_MAX), block: !!ok };
  r.filled = r.prio !== ADV_PRIO[0] || r.cushion > 0 || r.floor > 0 || r.fopPct > 0 || !!r.rules;
  return r;
}

/** Зберегти «Мої правила» з панелі (відсотки: числа 0…100). Номер картки, IBAN чи токен у вільному тексті: відмова. */
function advRulesSave_(ss, x) {
  x = x || {};
  const text = String(x.rules || '').trim();
  if (text.length > ADV_RULES_MAX) throw new Error('«Ще мої правила»: до ' + ADV_RULES_MAX + ' знаків (зараз ' + text.length + ').');
  if (advSecretLike_(text)) throw new Error('У «Ще мої правила» схоже на номер картки, IBAN чи токен: приберіть його. Такого в таблиці й у знімку бути не повинно.');
  settingsRulesUpgrade_(ss);
  const pct = v => Math.min(100, Math.max(0, Number(v) || 0)) / 100;
  ss.getSheetByName(SETTINGS_SHEET).getRange(27, 2, 6, 1).setValues([[ADV_PRIO.indexOf(x.prio) >= 0 ? x.prio : ADV_PRIO[0]], [pct(x.cushion)],
    [Math.max(0, Math.round(Number(x.floor) || 0))], [x.debtsMust === false || x.debtsMust === 'ні' ? 'ні' : 'так'], [pct(x.fopPct)], [text]]);
  return advRules_(ss);
}

/** Розділ знімка «Мої правила». */
function advRulesLines_(ss) {
  const r = advRules_(ss), pc = n => String(Math.round(n * 1000) / 10).replace('.', ',') + ' %', out = ['## Мої правила'];
  if (r.prio !== ADV_PRIO[0]) out.push('- Пріоритет №1: ' + r.prio + '.');
  if (r.cushion > 0) out.push('- Подушка: відкладати ' + pc(r.cushion) + ' від кожного доходу.');
  if (r.floor > 0) out.push('- Не опускатися нижче: ' + advMoney_(r.floor) + ' на картці (запас).');
  out.push('- Розстрочки й кредити: ' + (r.debtsMust ? 'обов\'язкові платежі, не переносити.' : 'можна переносити, якщо бракує.'));
  if (modOn_('fop')) out.push('- ФОП: ' + (r.fopPct > 0 ? 'відкладати ' + pc(r.fopPct) + ' з доходу ФОП на податки.' : 'відкладати на податки за планом аркуша «' + FOP_SHEET + '».'));
  if (r.rules) out.push('- Ще: ' + advMask_(r.rules).replace(/\s*\n\s*/g, '; '));
  if (!r.filled) out.push('- Більше правил ще немає. Порадь заповнити: «🚀 Підключення», крок 4 (або «' + SETTINGS_SHEET + '», рядки 27–32).');
  return out.concat(['']);
}

/** Розділ знімка «Де що в таблиці»: куди підказати вписати те, чого бракує. */
function advMapLines_(ss) {
  const out = ['## Де що в таблиці',
    '- «' + JOURNAL + '»: кожна операція рядком: A дата, B сума (завжди додатна), C категорія, D коментар, E статус (' + STATUS_FACT + ', ' + STATUS_PLAN + ', ' + STATUS_DRAFT + '), ' +
      'F непомітна вручну, G хто' + (modOn_('credit') ? ', H причина, якщо з ліміту' : '') + '. Чернетки з банку підтверджують тут.',
    '- «' + DICT_SHEET + '»: категорії (A), групи (C), непомітні (D), F план на місяць (бюджет категорії)' + (familyMode_(ss) ? ', G чий бюджет' : '') + '.',
    '- «' + REG_SHEET + '»: повторювані платежі й доходи (назва, категорія, сума, день, як часто, увімкнено). З них самі стають плани в журналі.',
    '- «' + SETTINGS_SHEET + '»: B5 з якого числа дохід іде в наступний місяць, B6 орієнтир погашення ліміту, рядки 27–32 «Мої правила».'];
  if (modOn_('debts')) out.push('- «' + DEBTS_SHEET + '»: борг чи розстрочка рядком: назва, напрям, сума, платіж на місяць, день.');
  if (modOn_('fop')) out.push('- «' + FOP_SHEET + '»: група, ставки, ліміти, дохід ФОП і податки.');
  out.push('- Аркуші «' + DICT_SHEET + '», «' + SETTINGS_SHEET + '» та інші службові сховані: показати можна на «' + START_SHEET + '», модуль «' + modName_('service') + '»: так.', '');
  return out;
}

/* ---------- знімок ---------- */

function advLevel_(ss) {
  const s = ss.getSheetByName(SETTINGS_SHEET), v = s ? String(s.getRange(SET.advisor).getValue()).trim() : '';
  return ADV_LEVELS.indexOf(v) >= 0 ? v : ADV_LEVELS[1];
}

function advVal_(v, label, tz) {
  if (v === null) return 'немає даних';
  if (Object.prototype.toString.call(v) === '[object Date]') { const k = dateKey_(v, tz); return k.slice(8) + '.' + k.slice(5, 7) + '.' + k.slice(0, 4); }
  if (/ставка/.test(label)) return (Math.round(Number(v) * 1000) / 10).toString().replace('.', ',') + ' %';
  if (/^(Днів|Чернеток|Червоних|Невідомих|Планів|З ліміту без причини)/.test(label)) return String(Math.round(Number(v) || 0));
  return money0_(v).replace(' ₴', ' грн');
}
const advMoney_ = v => money0_(v).replace(' ₴', ' грн');
const advDate_ = k => k.slice(8, 10) + '.' + k.slice(5, 7) + '.' + k.slice(0, 4);
function advNoDigits_(s) { return String(s || '').replace(/\d{6,}/g, '…').slice(0, 40); }

/** Чиста функція над прочитаними даними: текст знімка (Markdown-заголовки «#»/«##»). */
function snapshotText_(ss, opt) {
  opt = opt || {};
  const tz = ss.getSpreadsheetTimeZone(), today = opt.today || todayKey_(ss), level = opt.level || advLevel_(ss), lv = ADV_LEVELS.indexOf(level);
  const D = mainByLabel_(ss), rows = readJournal_(ss).filter(o => !o.blank && o.key), fam = familyMode_(ss), ms = members_(ss);
  const set = ss.getSheetByName(SETTINGS_SHEET), bday = set ? Number(set.getRange(SET.bday).getValue()) || 25 : 25;
  const credit = modOn_('credit'), out = [];
  const at = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');
  out.push('# ' + ADV_DOC, 'Дата знімка: ' + at + ' (' + tz + '). Оновлюється сам щодня. Суми в гривнях, округлено до гривні.',
    'Рівень деталей: ' + level + '. Режим: ' + (fam ? 'сім\'я (' + (ms.join(', ') || 'учасників не вписано') + ')' : 'одна людина') + '.',
    'Це знімок, не жива таблиця: відповідай лише цими цифрами й називай дату знімка.',
    'Навичка: версія ' + ADV_SKILL.v + '. Якщо твоя навичка старіша, один раз підкажи оновити її: «🚀 Підключення», крок 4.', '');
  out.push('## Як читати',
    '- «Мої гроші»: надходження бюджетного місяця (зарплата тощо) мінус відкладене на погашення мінус уже витрачене.' + (credit ? ' Кредитний ліміт до них не входить.' : ''),
    '- Бюджетний місяць: дохід з ' + bday + '-го числа йде в бюджет наступного місяця.',
    '- Факт: що вже сталося. План: майбутній платіж чи надходження. Чернетка: операція з банку, ще не підтверджена (у підсумки не входить).',
    '- «Непомітні»: дрібні витрати за звичкою (кава, доставка, таксі, підписки, дрібні покупки, комісії).',
    credit ? '- Кредитний ліміт: гроші банку. «З ліміту»: витрачено понад «мої». Грейс Mono: борг місяця без відсотків до кінця наступного місяця, якщо внести до строку; перекриття переказами з ліміту коштує комісію.' : '',
    modOn_('fop') ? '- ФОП: дохід ФОП і податки рахуються окремо від сімейного бюджету; у сім\'ю йде лише переказ з рахунку ФОП на картку.' : '', '');
  advRulesLines_(ss).forEach(l => out.push(l));
  // головне
  const labels = ['Доступно моїх сьогодні', 'Можна на день', 'Днів до кінця місяця', 'На картці зараз', 'Прогноз на кінець місяця (за планом)',
    'Прогноз на кінець наступного місяця (за планом)', 'Витрачаєте ≈ на день', 'Як зараз: на картці на кінець місяця', 'Найнижча точка за планом', 'Дата найнижчої точки',
    'Витрати цього місяця: факт', 'Витрати цього місяця: план', 'Непомітні цього місяця', 'Чернеток чекають підтвердження']
    .concat(credit ? ['Кредитний ліміт', 'Борг за лімітом зараз', 'Доступно з лімітом', 'З ліміту цього місяця', 'Відкладено на погашення (бюджетний місяць)',
      'Орієнтир погашення на місяць', 'Грейс: внести до строку', 'Грейс: строк', 'Грейс: переказами (оцінка)', 'Грейс: комісія, якщо все переказами',
      'Грейс: ставка комісії за переказ з ліміту'] : [])
    .concat(modOn_('debts') ? MAIN_DEBTS : []);
  out.push('## Головне на ' + advDate_(today));
  labels.forEach(l => { const v = D(l); if (v !== null) out.push('- ' + l + ': ' + advVal_(v, l, tz)); });
  out.push('');
  // бюджети
  const bud = budgetLines_(ss, rows, null, true);
  if (bud.length) {
    out.push('## Бюджети категорій: ' + ymName_(today.slice(0, 7)));
    bud.forEach(b => out.push('- ' + b.cat + (b.owner ? ' (' + b.owner + ')' : '') + ': план ' + advMoney_(b.budget) + ', витрачено ' + advMoney_(b.spent) +
      (b.draft ? ' (з них у чернетках ' + advMoney_(b.draft) + ')' : '') + ', ' + (b.over ? 'перевищено на ' + advMoney_(b.spent - b.budget) : 'лишилось ' + advMoney_(b.budget - b.spent)) +
      '; норма на сьогодні ' + advMoney_(b.norm) + (b.fast ? ', швидше норми' : '')));
    out.push('');
  }
  // місяці
  const months = [], ym0 = today.slice(0, 7);
  let ym = ym0;
  for (let i = 0; i < 4; i++) { months.push(ym); ym = prevYm_(ym); }
  const start = startKey_(ss);
  out.push('## Місяці (факти)');
  months.forEach(m => {
    if (start && lastDayOf_(m) < start) return;
    const to = m === ym0 ? today : lastDayOf_(m), s = factStats_(rows, m + '-01', to);
    const inc = round2_(rows.filter(o => o.key >= m + '-01' && o.key <= to && o.i === T_IN && Number(o.m) === 0 && !isPlan_(o.e) && o.e !== STATUS_DRAFT && o.kind !== K_FOP)
      .reduce((a, o) => a + (Number(o.o) || 0), 0));
    out.push('### ' + ymName_(m) + (m === ym0 ? ' (до ' + advDate_(today) + ')' : ''),
      '- Витрати: ' + advMoney_(s.spent) + '; надходження: ' + advMoney_(inc) + '; непомітні: ' + advMoney_(s.holeSum) + (credit ? '; з ліміту: ' + advMoney_(s.limit) : '') +
      (s.drafts ? '; ще в чернетках ' + s.drafts + ' на ' + advMoney_(s.draftSum) : ''));
    if (lv >= 1) {
      const top = topOf_(s.cats, 40);
      if (top.length) out.push('- Категорії: ' + top.map(x => x[0] + ' ' + advMoney_(x[1])).join('; '));
      const holes = topOf_(s.holes, 20);
      if (holes.length) out.push('- Непомітні: ' + holes.map(x => x[0] + ' ' + advMoney_(x[1])).join('; '));
      if (credit && s.limit > 0.005) out.push('- З ліміту, чому: ' + topOf_(s.reasons, 10).map(x => x[0] + ' ' + advMoney_(x[1])).join('; '));
    }
    if (fam) {
      const by = {};
      rows.filter(o => o.key >= m + '-01' && o.key <= to && Number(o.m) === 0 && !isPlan_(o.e) && o.e !== STATUS_DRAFT && o.i === T_SPEND && o.kind !== K_FOP)
        .forEach(o => { const k = o.g || WHO_SHARED; by[k] = (by[k] || 0) + (Number(o.o) || 0); });
      const parts = topOf_(by, 10);
      if (parts.length) out.push('- По учасниках (витрати): ' + parts.map(x => x[0] + ' ' + advMoney_(x[1])).join('; '));
    }
  });
  out.push('');
  // плани на 2 місяці
  const horizon = lastDayOf_(addDays_(lastDayOf_(ym0), 1).slice(0, 7));
  const plans = rows.filter(o => (isPlan_(o.e) || (o.e === '' && o.key > today)) && o.key >= today && o.key <= horizon).sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
  out.push('## Плани до ' + advDate_(horizon));
  if (!plans.length) out.push('- Планів немає.');
  else if (lv === 0) {
    [ym0, horizon.slice(0, 7)].forEach(m => {
      const p = plans.filter(o => o.key.slice(0, 7) === m), inc = p.filter(o => TYPES_IN.indexOf(o.i) >= 0).reduce((a, o) => a + Number(o.b), 0);
      out.push('- ' + ymName_(m) + ': надходжень за планом ' + advMoney_(inc) + ', списань ' + advMoney_(p.reduce((a, o) => a + Number(o.b), 0) - inc));
    });
  } else {
    plans.slice(0, 80).forEach(o => out.push('- ' + advDate_(o.key) + ' · ' + (o.c || 'без категорії') + ' · ' + advMoney_(o.b) + (TYPES_IN.indexOf(o.i) >= 0 ? ' (надходження)' : '') +
      (o.g ? ' · ' + o.g : '') + (o.d.indexOf(FOP_TAG) === 0 || o.d.indexOf(DEBT_TAG) === 0 ? ' · ' + o.d : o.d.indexOf(REG_TAG) === 0 ? ' · регулярний' : '')));
    if (plans.length > 80) out.push('- … ще ' + (plans.length - 80));
  }
  out.push('');
  // підписки
  const from3 = prevYm_(prevYm_(ym0)) + '-01', subs = rows.filter(o => o.c === 'Підписки' && o.key >= from3 && o.key <= today && Number(o.m) === 0 && !isPlan_(o.e) && o.e !== STATUS_DRAFT);
  if (subs.length) {
    out.push('## Підписки (з ' + advDate_(from3) + ')', '- Разом: ' + advMoney_(subs.reduce((a, o) => a + (Number(o.o) || 0), 0)) + ' (платежів: ' + subs.length + ').');
    if (lv >= 1) {
      const by = {};
      subs.forEach(o => { const k = advNoDigits_(o.bank && o.bank[1]) || 'без опису з банку'; by[k] = (by[k] || 0) + (Number(o.o) || 0); });
      out.push('- За сервісами (опис з банку): ' + topOf_(by, 20).map(x => x[0] + ' ' + advMoney_(x[1])).join('; '));
    }
    out.push('');
  }
  // борги
  if (modOn_('debts')) {
    const ds = debtsList_(ss);
    if (ds.length) {
      out.push('## Борги й розстрочки');
      ds.forEach(d => out.push('- ' + d.name + ' (' + d.dir + '): залишок ' + advMoney_(d.left === null ? d.total - d.paid : d.left) + ' з ' + advMoney_(d.total) +
        (d.monthly > 0 ? ', платіж ' + advMoney_(d.monthly) + ' на місяць, ' + d.day + '-го' : '') + (d.who ? ' · ' + d.who : '')));
      out.push('');
    }
  }
  // ФОП
  if (modOn_('fop')) {
    const fops = fopList_(ss), rates = fopRates_(ss, Number(today.slice(0, 4)));
    if (fops.length && rates) {
      const y = today.slice(0, 4), single = fops.length <= 1;
      out.push('## ФОП (ставки й ліміти: аркуш «' + FOP_SHEET + '»)');
      fops.forEach(f => {
        const inc = fopIncome_(rows, f.name, single, y + '-01-01', (Number(y) + 1) + '-01-01'), lim = rates.limit[f.group];
        out.push('- ' + f.name + ': ' + f.group + ' група; дохід з початку року ' + advMoney_(inc) + (lim ? ', ліміт групи ' + advMoney_(lim) + ' (' + Math.round(inc / lim * 100) + ' %)' : ''));
      });
      const taxes = plans.filter(o => o.d.indexOf(FOP_TAG) === 0);
      if (taxes.length) out.push('- Податки за планом: ' + taxes.map(o => o.d.slice(FOP_TAG.length) + ' ' + advMoney_(o.b) + ' до ' + advDate_(o.key)).join('; '));
      out.push('');
    }
  }
  // операції
  if (lv >= 2) {
    const from = addDays_(today, -ADV_OPS_DAYS), ops = rows.filter(o => o.key >= from && o.key <= today && Number(o.m) === 0 && !isPlan_(o.e) && o.e !== STATUS_DRAFT && o.kind !== K_FOP);
    out.push('## Операції за ' + ADV_OPS_DAYS + ' днів (без коментарів)');
    ops.slice(0, 600).forEach(o => out.push('- ' + advDate_(o.key) + ' · ' + (o.c || 'без категорії') + ' · ' + (TYPES_IN.indexOf(o.i) >= 0 ? '+' : '−') + advMoney_(o.o || o.b) +
      (o.p ? ' · непомітна: ' + o.p : '') + (o.g ? ' · ' + o.g : '')));
    if (ops.length > 600) out.push('- … ще ' + (ops.length - 600));
    out.push('');
  }
  advMapLines_(ss).forEach(l => out.push(l));
  out.push('## Чого тут немає', '- Коментарів до операцій, номерів карток і рахунків, токенів банку й Telegram, описів з банку (крім назв підписок).');
  return out.filter((l, i, a) => !(l === '' && a[i - 1] === '')).join('\n');
}

/* ---------- документ на Диску ---------- */

function advFolder_(ss) {
  const dp = PropertiesService.getDocumentProperties(), id = dp.getProperty('adv_folder_id');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* видалили */ } }
  let parent;
  try { const it = DriveApp.getFileById(ss.getId()).getParents(); parent = it.hasNext() ? it.next() : DriveApp.getRootFolder(); } catch (e) { parent = DriveApp.getRootFolder(); }
  const folder = parent.createFolder(ADV_FOLDER);
  dp.setProperty('adv_folder_id', folder.getId());
  return folder;
}

/** Наш документ: за збереженим id (чужий документ з такою ж назвою не береться). */
function advDoc_(ss, create) {
  const dp = PropertiesService.getDocumentProperties(), id = dp.getProperty('adv_doc_id');
  if (id) {
    try { if (!DriveApp.getFileById(id).isTrashed()) return DocumentApp.openById(id); } catch (e) { /* видалили: створимо новий */ }
  }
  if (!create) return null;
  const doc = DocumentApp.create(ADV_DOC);
  try { DriveApp.getFileById(doc.getId()).moveTo(advFolder_(ss)); } catch (e) { /* лишиться в корені Диска */ }
  dp.setProperty('adv_doc_id', doc.getId());
  return doc;
}

/** Записати знімок у документ. → { url, chars } */
function snapshotWrite_(ss) {
  const text = snapshotText_(ss), doc = advDoc_(ss, true), body = doc.getBody(), dp = PropertiesService.getDocumentProperties();
  body.clear();
  const H = { '#': DocumentApp.ParagraphHeading.HEADING1, '##': DocumentApp.ParagraphHeading.HEADING2, '###': DocumentApp.ParagraphHeading.HEADING3 };
  text.split('\n').forEach((l, i) => {
    const h = /^(#{1,3}) (.*)$/.exec(l), t = h ? h[2] : l;
    const first = i === 0 ? body.getParagraphs()[0] : null;
    const p = first ? first.setText(t) : body.appendParagraph(t);
    p.setHeading(h ? H[h[1]] : DocumentApp.ParagraphHeading.NORMAL);
  });
  doc.saveAndClose();
  dp.setProperties({ adv_day: todayKey_(ss), adv_t: String(Date.now()), adv_chars: String(text.length) });
  dp.deleteProperty('adv_err');
  return { url: doc.getUrl(), chars: text.length };
}

/** З mbsTick: раз на день після 05:00; помилка: ще раз за годину. */
function advisorIfDue_(ss) {
  if (!modOn_('advisor')) return null;
  const dp = PropertiesService.getDocumentProperties(), today = todayKey_(ss), err = jsonProp_(dp, 'adv_err');
  if (dp.getProperty('adv_day') === today) return null;
  if (Number(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'HH')) < ADV_HOUR) return null;
  if (err && Date.now() - Number(err.t || 0) < ADV_RETRY_MS) return null;
  try { return snapshotWrite_(ss); }
  catch (e) { dp.setProperty('adv_err', JSON.stringify({ t: Date.now(), msg: String(e && e.message || e).slice(0, 300) })); return null; }
}

function advisorLines_(ss) {
  const dp = PropertiesService.getDocumentProperties(), tz = ss.getSpreadsheetTimeZone(), err = jsonProp_(dp, 'adv_err'), t = dp.getProperty('adv_t');
  if (!modOn_('advisor')) return ['Модуль вимкнено (на «' + START_SHEET + '»).'];
  const out = [t ? 'Знімок оновлено ' + Utilities.formatDate(new Date(Number(t)), tz, 'dd.MM HH:mm') + ' (' + (dp.getProperty('adv_chars') || '?') + ' знаків), рівень: ' + esc_(advLevel_(ss)) + '.'
    : 'Знімка ще немає: з\'явиться сам після 05:00 або «🤖 Радник Gemini → 🔄 Оновити знімок зараз».'];
  if (err) out.push('⚠️ Помилка ' + Utilities.formatDate(new Date(Number(err.t)), tz, 'dd.MM HH:mm') + ': ' + esc_(err.msg));
  return out;
}

/* ---------- меню ---------- */

function advisorSnapshotNow() {
  if (!menuOk_('🤖 Радник Gemini')) return;
  const ss = SpreadsheetApp.getActive();
  if (!modOn_('advisor')) modSet_(ss, 'advisor', true);
  let r;
  try { r = snapshotWrite_(ss); } catch (e) { show_('🤖 Радник Gemini', ['⚠️ Знімок не записано: ' + esc_(e && e.message || e)]); return; }
  show_('🤖 Знімок оновлено', [link_(r.url, ADV_DOC) + ' (' + r.chars + ' знаків).', 'Далі сам щодня після 0' + ADV_HOUR + ':00. Як підключити навичку: «🤖 Радник Gemini → 🤖 Як підключити навичку».']);
}

/** Вікно: посилання на знімок, одне повідомлення для чату Gemini (кнопка «Скопіювати») і запасні способи. */
function advisorNow() {
  if (!menuOk_('🤖 Радник Gemini')) return;
  const ss = SpreadsheetApp.getActive(), doc = advDoc_(ss, false), r = advRules_(ss), url = doc ? doc.getUrl() : '';
  const area = (id, t, hgt) => '<textarea id="' + id + '" readonly style="width:100%;height:' + hgt + 'px;font:12px/1.4 monospace" onclick="this.select()">' + esc_(t) + '</textarea>';
  const lines = [
    modOn_('advisor') ? '' : '⚠️ Модуль «🤖 Радник Gemini» вимкнено: увімкніть на «' + START_SHEET + '» або натисніть «🔄 Оновити знімок зараз».',
    doc ? 'Знімок: ' + link_(url, ADV_DOC) + ' (на вашому Диску, оновлюється щодня). У тексті для Gemini пряме посилання на нього: Gemini відкриє знімок одразу, без пошуку.'
      : 'Знімка ще немає: «🤖 Радник Gemini → 🔄 Оновити знімок зараз», тоді в тексті для Gemini буде пряме посилання на нього (без знімка Gemini шукатиме документ за назвою).',
    advSkillRelink_() ? '⚠️ Знімок на Диску створено заново: скопіюйте текст ще раз, щоб у навичці було нове посилання.' : '',
    '<b>Як підключити навичку</b> (особистий акаунт Google, 18+; навичка версії ' + ADV_SKILL.v + '):',
    '1. <button onclick="cp()">📋 Скопіювати для Gemini</button> <span id="ok"></span>',
    '2. Відкрийте ' + link_(CN_GEMINI, 'Gemini') + ', вставте в чат і надішліть: Gemini сам створить навичку «' + esc_(ADV_SKILL.title) + '».',
    '3. Далі просто питайте в будь-якому чаті Gemini, наприклад: «' + esc_(ADV_SKILL.starters[0]) + '». Google Диск у Gemini має бути підключений.',
    area('msg', advSkillMsg_(url), 150),
    '<details><summary>Інакше: створити вручну або з файла</summary>Skills → Create → Create manually: назва, опис та інструкції нижче. Або Upload: ' +
      link_(REPO_URL + '/blob/main/gemini/' + ADV_SKILL.name + '/SKILL.md', 'SKILL.md у репозиторії') + ' (у файлі посилання немає: Gemini шукатиме знімок за назвою).<br>Назва:' + area('t', ADV_SKILL.title, 44) +
      'Опис:' + area('d', ADV_SKILL.description, 90) + 'Інструкції:' + area('i', advSkillInstructions_(url), 170) + '</details>',
    '<b>Мої правила</b>: ' + (r.filled ? 'заповнено ✓' : 'ще не заповнено') + '. Пріоритет, подушка, запас на картці, розстрочки, ФОП і свої правила: «🚀 Підключення», крок 4 ' +
      '(або «' + SETTINGS_SHEET + '», рядки 27–32). Вони щодня потрапляють у знімок, тож навичку перестворювати не треба.',
    'Приватність: знімок лежить лише на вашому Диску; скрипт нікуди його не надсилає. Gemini читає його, лише коли ви питаєте. Коментарів, номерів карток і токенів у ньому немає.',
  ];
  const js = '<script>function cp(){var t=document.getElementById("msg"),ok=document.getElementById("ok");function done(){ok.textContent="✓ Скопійовано";' +
    'try{google.script.run.advSkillCopied()}catch(e){}}function old(){t.select();try{if(document.execCommand("copy")){done();return}}catch(e){}' +
    'ok.textContent="Виділено: натисніть Cmd+C або Ctrl+C"}try{navigator.clipboard.writeText(t.value).then(done,old)}catch(e){old()}}</script>';
  const html = HtmlService.createHtmlOutput('<div style="font:14px/1.5 Arial,sans-serif">' + lines.filter(Boolean).map(l => '<p>' + l + '</p>').join('') + '</div>' + js)
    .setWidth(640).setHeight(640);
  try { SpreadsheetApp.getUi().showModalDialog(html, '🤖 Радник Gemini'); } catch (e) { Logger.log(advSkillMd_()); }
}

/** Людина скопіювала повідомлення для Gemini: запам'ятати версію навички й знімок, на який вона посилається (панель підкаже, коли треба скопіювати ще раз). */
function advSkillCopied() {
  guard_(SpreadsheetApp.getActive());                                    // лише таблиця шаблону
  const dp = PropertiesService.getDocumentProperties();
  dp.setProperties({ adv_skill_v: String(ADV_SKILL.v), adv_skill_doc: dp.getProperty('adv_doc_id') || '' });   // з яким посиланням скопійовано
  return true;
}

/** Навичку скопіювали з посиланням на інший знімок (документ створено заново або тоді його ще не було): скопіювати ще раз. */
function advSkillRelink_() {
  const dp = PropertiesService.getDocumentProperties();
  return !!dp.getProperty('adv_skill_v') && (dp.getProperty('adv_skill_doc') || '') !== (dp.getProperty('adv_doc_id') || '');
}

/** Скопійована колись навичка старіша за цю версію коду. */
function advSkillOld_() {
  const v = Number(PropertiesService.getDocumentProperties().getProperty('adv_skill_v') || 0);
  return v > 0 && v < ADV_SKILL.v;
}
