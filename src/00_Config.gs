/** 00_Config.gs: Monitoring Budgets (шаблон). Назви аркушів, межі таблиць, колонки журналу. */

/*
 * Це код шаблону «Monitoring Budgets» (відкритий, ліцензія MIT). Він працює лише в таблиці-шаблоні або в її копії.
 * У таблиці, схожій на особисту «Monitoring Budget» (аркуші «m» і «y»), він нічого не змінює: див. 01_Guard.gs.
 * Жодних токенів, паролів чи адрес тут немає і бути не повинно: токени людина вводить сама,
 * вони зберігаються лише в налаштуваннях скрипта в її Google-акаунті (див. PRIVACY.md).
 */

const TEMPLATE_MARK = 'monitoring-budgets-template';   // мітка шаблону: її шукає CI перед заливкою (ci/guard.js)
const TEMPLATE_VERSION = '0.5.3';
const SUPPORT_URL = 'https://send.monobank.ua/jar/6ChdYHjkd9';   // «❤️ Підтримати автора»: публічна банка автора
const APP_NAME = 'Monitoring Budgets';
const REPO_URL = 'https://github.com/VladyYerys/monitoring-budgets';                 // відкритий код, «Що нового», питання (issues)
// Номер найновішої версії (публічний файл у репозиторії, без жодних даних людей): див. 42_Update.gs і PRIVACY.md
const VERSION_URL = 'https://raw.githubusercontent.com/VladyYerys/monitoring-budgets/main/version.json';

/* ---------- аркуші ---------- */
const JOURNAL = 'Журнал';
const DICT_SHEET = 'Довідник';
const SETTINGS_SHEET = 'Налаштування';
const CP_SHEET = 'Залишки банку';
const MONTHS_SHEET = 'Місяці';
const DAYS_SHEET = 'Дні';
const HISTORY_SHEET = 'Історія місяців';
const RULES_SHEET = 'Правила';
const BANK_SHEET = 'Банк';
const MAIN_SHEET = 'Головне';
const CHECK_SHEET = 'Перевірка';
const START_SHEET = '⚡ Швидке підключення';                          // до версії 15: «Старт»
const ACC_SHEET = 'Рахунки';
const FOP_SHEET = 'ФОП';
const DEBTS_SHEET = 'Борги';

/* ---------- журнал ---------- */
const J_HEAD = 4, J_FIRST = 5, J_LAST = 3004;        // 3 000 рядків
const J_RANGE = n => "'" + JOURNAL + "'!$" + n + '$' + J_FIRST + ':$' + n + '$' + J_LAST;   // «'Журнал'!$O$5:$O$3004»

/**
 * Колонки журналу (літери). Вводить людина: A…H. Рахують формули: I, J і службові L…X (приховані).
 * Y…AD: що прийшло з банку (номер операції, опис, сума зі знаком, MCC, час, залишок після операції); AE: черговість запису;
 * AF: рахунок (з «Рахунків»; порожньо: основна картка). Їх пише скрипт; при сортуванні вони їдуть разом зі своїм рядком.
 * AG: вид рахунку, AH: сума зі знаком для будь-якого рахунку (формули). L (сума зі знаком) рахує лише основну картку:
 * з неї «Дні» й контрольні точки; інші картки сім'ї йдуть у «мої» і «Місяці», рахунок ФОП: лише на аркуш «ФОП».
 */
const JC = {
  date: 'A', sum: 'B', cat: 'C', note: 'D', status: 'E', hole: 'F', who: 'G', reason: 'H',
  type: 'I', mineAfter: 'J',
  signed: 'L', state: 'M', plan: 'N', fact: 'O', holeEff: 'P', bmonth: 'Q', mineKind: 'R',
  mineIn: 'S', out: 'T', mineRun: 'U', fromMine: 'V', fromLimit: 'W', limitWhy: 'X',
  bankId: 'Y', bankDesc: 'Z', bankAmt: 'AA', bankMcc: 'AB', bankTime: 'AC', bankBal: 'AD', order: 'AE',
  acct: 'AF', acctKind: 'AG', signedAll: 'AH',
};
const J_INPUT_N = 8;                                      // A…H
const J_INPUT_LAST = 'H', J_SERVICE_FIRST = 'L', J_SERVICE_LAST = 'AH';
const BANK_N = 6;                                         // Y…AD
const J_READ_N = 34;                                      // A…AH одним читанням
const J_MOVE_N = BANK_N + 2;                              // Y…AF: їдуть разом із рядком при сортуванні

const STATUS_FACT = '✅ факт', STATUS_DRAFT = '🏦 чернетка', STATUS_PLAN = 'план', STATUS_REG = '🔁 план регулярний';
const STATUSES = [STATUS_FACT, STATUS_DRAFT, STATUS_PLAN, STATUS_REG];
function isPlan_(e) { return e === STATUS_PLAN || e === STATUS_REG; }
const CAT_TRANSFER = '↔ Переказ (уточніть)';
const CAT_OWN_MOVE = 'Переказ між своїми рахунками', CAT_FOP_INCOME = 'Дохід ФОП';
const CAT_TAX = { ep: 'Єдиний податок', esv: 'ЄСВ', vz: 'Військовий збір' };

/* ---------- типи категорій ---------- */
const T_SPEND = 'Витрати', T_SAVE = 'Заощадження', T_TAX = 'Податки', T_DEBT = 'Борги', T_IN = 'Надходження',
  T_RESERVE = 'Резерв', T_CORR_IN = 'Корекція +', T_CORR_OUT = 'Корекція −';
const TYPES = [T_SPEND, T_SAVE, T_TAX, T_DEBT, T_IN, T_RESERVE, T_CORR_IN, T_CORR_OUT];
const TYPES_IN = [T_IN, T_CORR_IN];                                   // гроші приходять на картку
const TYPES_OUT = [T_SPEND, T_SAVE, T_TAX, T_DEBT, T_CORR_OUT];       // гроші йдуть з картки (Резерв лишається на картці)
const MINE_IN = 'надходження', MINE_SET_ASIDE = 'відкладення';       // «Мої гроші» (Довідник, колонка E)

/* ---------- довідник ---------- */
const D_FIRST = 5, D_LAST = 84;                          // 80 категорій
const D_HOLES_FIRST = 5, D_HOLES_LAST = 24;              // «чорні діри» (колонка I)
const D_REASONS_FIRST = 5, D_REASONS_LAST = 24;          // «чому з ліміту» (колонка K)
const DC = { name: 'A', type: 'B', group: 'C', hole: 'D', mine: 'E', plan: 'F', holes: 'I', holesNote: 'J', reasons: 'K' };
const D_RANGE = (c, a, b) => "'" + DICT_SHEET + "'!$" + c + '$' + (a || D_FIRST) + ':$' + c + '$' + (b || D_LAST);

/* ---------- налаштування ---------- */
const SET = { start: 'B4', bday: 'B5', goal: 'B6', card: 'B7', margin: 'B8', version: 'B9', rate: 'B10', mode: 'B13', advisor: 'B23', support: 'B24',
  prio: 'B27', cushion: 'B28', floor: 'B29', debtsMust: 'B30', fopPct: 'B31', rules: 'B32',
  autoFact: 'B35' };                                                     // B35: операції з банку з категорією одразу «✅ факт»   // B27…B32: «Мої правила для радника Gemini»
const SET_MEMBERS_FIRST = 14, SET_MEMBERS_N = 5, SET_SHARED_ROW = 19, SET_ALL_ROW = 20;   // учасники B14:B18, «спільне» B19, «уся сім'я» B20
const MODE_SOLO = 'одна людина', MODE_FAMILY = 'сім\'я', WHO_SHARED = 'спільне', WHO_ALL = 'уся сім\'я';
const SET_REF = k => "'" + SETTINGS_SHEET + "'!$" + SET[k].charAt(0) + '$' + SET[k].slice(1);
const SET_MARK_CELL = 'A40';                             // прихована мітка шаблону (їде разом з копією таблиці; див. 01_Guard.gs)
const CARD_CREDIT = 'кредитна з лімітом', CARD_DEBIT = 'дебетова';

/* ---------- залишки банку (контрольні точки) ---------- */
const CP_FIRST = 5, CP_LAST = 2004;
const CP_RANGE = c => "'" + CP_SHEET + "'!$" + c + '$' + CP_FIRST + ':$' + c + '$' + CP_LAST;   // A дата, B на початок дня, C ліміт

/* ---------- «Місяці», «Дні», «Історія місяців» ---------- */
const M_YEAR = 'B2';                                     // рік на «Місяцях» (список)
const M_WHO = 'D2', M_WHO_CRIT = '$A$3';                  // хто: «уся сім'я» або учасник; A3: службовий критерій для SUMIFS
const M_HEAD = 4, M_MONTH_ROW = 5, M_FIRST = 7;          // рядок 5 (прихований): перше число кожного місяця
const M_COL0 = 4;                                        // D: січень «план»
const DAYS_FIRST = 6;
const H_FIRST = 4;

/* ---------- «Правила» і «Банк» ---------- */
const RULES_FIRST = 5, RULES_LAST = 204;
const BANK_LOG_FIRST = 5, BANK_LOG_N = 12;

/* ---------- «Рахунки» ---------- */
const ACC_FIRST = 5, ACC_LAST = 34;                       // A назва · B банк · C чий · D вид · E ФОП · F ключ · G останній імпорт · H примітка
const K_MAIN = 'основна картка', K_CARD = 'інша картка', K_FOP = 'рахунок ФОП';
const ACC_KINDS = [K_MAIN, K_CARD, K_FOP];
const ACC_RANGE = c => "'" + ACC_SHEET + "'!$" + c + '$' + ACC_FIRST + ':$' + c + '$' + ACC_LAST;

/* ---------- розклади (назви обробників з префіксом mbs: шаблон не відповідає на розклади іншого коду) ---------- */
const TRG = { tick: 'mbsTick', edit: 'mbsEdit', soon: 'mbsSoon', imp: 'mbsImport', impRetry: 'mbsImportRetry',
  backup: 'mbsBackup', morning: 'mbsMorning', morningRetry: 'mbsMorningRetry', period: 'mbsPeriod' };
