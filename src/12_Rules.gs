/** 12_Rules.gs: Аркуш «Правила»: яку категорію ставити операції з банку чи рядку з коментарем. */

/*
 * Рядок правила: Активне? · Текст (кілька слів через кому) · Де шукати (опис банку / коментар / будь-де / MCC) ·
 * Напрям (усі / списання / надходження) · Категорія · Непомітна · Примітка. Перше активне правило, що підходить, виграє.
 * MCC: числа або діапазони через кому («5812, 5813-5814»). Регістр і цифри в описі не мають значення.
 * Скрипт сам пропонує правило («✨ нове, підтвердіть»), коли ви змінюєте категорію рядка з банку.
 */

const NEW_RULE = '✨ нове, підтвердіть';
const RULE_WHERE = ['опис банку', 'коментар', 'будь-де', 'MCC'], RULE_DIR = ['усі', 'списання', 'надходження'];

/** Стартові правила: загальні назви мереж і коди MCC, жодних особистих даних. Людина змінює їх під себе. */
const RULES_TAX = [
  ['так', 'єдиний податок', 'будь-де', 'списання', 'Єдиний податок'],
  ['так', 'єдиний внесок, єсв', 'будь-де', 'списання', 'ЄСВ'],
  ['так', 'військовий збір', 'будь-де', 'списання', 'Військовий збір'],
];
const RULES_START = RULES_TAX.concat([
  ['так', 'атб, сільпо, novus, фора, ашан, auchan, metro, varus, близенько, eko market, космос', 'опис банку', 'списання', 'Продукти'],
  ['так', 'glovo, bolt food, rocket', 'опис банку', 'списання', 'Доставка їжі'],
  ['так', 'uklon, bolt, uber', 'опис банку', 'списання', 'Таксі'],
  ['так', 'wog, okko, socar, upg, klo, shell', 'опис банку', 'списання', 'Пальне'],
  ['так', 'netflix, spotify, youtube, google one, apple.com, icloud, megogo, sweet.tv', 'опис банку', 'списання', 'Підписки'],
  ['так', 'аптека, apteka, podorozhnyk, подорожник', 'опис банку', 'списання', 'Здоров\'я й аптека'],
  ['так', 'kyivstar, київстар, vodafone, lifecell', 'опис банку', 'списання', 'Інтернет і зв\'язок'],
  ['так', 'кава, кофе, coffee', 'будь-де', 'списання', 'Кава та перекуси'],
  ['так', 'відсотки, комісія за', 'опис банку', 'списання', 'Комісії й відсотки банку'],
  ['так', '5411, 5422, 5441, 5451, 5499', 'MCC', 'списання', 'Продукти'],
  ['так', '5812, 5813, 5814', 'MCC', 'списання', 'Кафе й ресторани'],
  ['так', '4121', 'MCC', 'списання', 'Таксі'],
  ['так', '5541, 5542', 'MCC', 'списання', 'Пальне'],
  ['так', '5912', 'MCC', 'списання', 'Здоров\'я й аптека'],
  ['так', '4111, 4112, 4131', 'MCC', 'списання', 'Транспорт'],
  ['так', '4814, 4899', 'MCC', 'списання', 'Інтернет і зв\'язок'],
  ['так', '5651, 5661, 5691, 5699', 'MCC', 'списання', 'Одяг і взуття'],
  ['так', '5732, 5734, 5045', 'MCC', 'списання', 'Техніка'],
]);

function buildRules_(ss) {
  const sh = ensureSize_(sheetOf_(ss, RULES_SHEET), RULES_LAST, 7), d = ss.getSheetByName(DICT_SHEET);
  sh.getRange(1, 1).setValue('🧭 Правила').setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setValue('Яку категорію ставити операції з банку (і рядку, де категорія порожня, а в коментарі є слово). ' +
    'Перше активне правило, що підходить, виграє. «' + NEW_RULE + '»: запропонував скрипт, поставте «так», якщо згодні.').setFontColor('#5f6368');
  sh.getRange(4, 1, 1, 7).setValues([['Активне?', 'Текст (через кому)', 'Де шукати', 'Напрям', 'Категорія', 'Непомітна', 'Примітка']]).setFontWeight('bold');
  const n = RULES_LAST - RULES_FIRST + 1, rows = [];
  for (let i = 0; i < n; i++) { const x = RULES_START[i]; rows.push(x ? [x[0], x[1], x[2], x[3], x[4], '', ''] : ['', '', '', '', '', '', '']); }
  sh.getRange(RULES_FIRST, 1, n, 7).setValues(rows);
  const dv = (col, rule) => sh.getRange(col + RULES_FIRST + ':' + col + RULES_LAST).setDataValidation(rule.build());
  dv('A', SpreadsheetApp.newDataValidation().requireValueInList(['так', 'ні', NEW_RULE]).setAllowInvalid(false));
  dv('C', SpreadsheetApp.newDataValidation().requireValueInList(RULE_WHERE).setAllowInvalid(false));
  dv('D', SpreadsheetApp.newDataValidation().requireValueInList(RULE_DIR).setAllowInvalid(false));
  if (d) {
    dv('E', SpreadsheetApp.newDataValidation().requireValueInRange(d.getRange('A' + D_FIRST + ':A' + D_LAST)).setAllowInvalid(true));
    dv('F', SpreadsheetApp.newDataValidation().requireValueInRange(d.getRange('I' + D_HOLES_FIRST + ':I' + D_HOLES_LAST)).setAllowInvalid(true));
  }
  sh.getRange('B' + RULES_FIRST + ':B' + RULES_LAST).setNumberFormat('@');
  sh.setFrozenRows(4);
  return sh;
}

function norm_(s) {
  return String(s === null || s === undefined ? '' : s).toLowerCase()
    .replace(/[ʼ’`]/g, '\'').replace(/[\d#*№]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function readRules_(ss) {
  const r = (ss || SpreadsheetApp.getActive()).getSheetByName(RULES_SHEET);
  if (!r) return [];
  return parseRules_(r.getRange(RULES_FIRST, 1, RULES_LAST - RULES_FIRST + 1, 7).getValues());
}

/** Чиста функція: рядки «Правил» → правила. */
function parseRules_(values) {
  return values.map((v, i) => ({
    row: RULES_FIRST + i, active: String(v[0]).trim(), text: String(v[1] || ''),
    where: String(v[2] || 'опис банку').trim() || 'опис банку', dir: String(v[3] || 'усі').trim() || 'усі',
    category: String(v[4] || '').trim(), hole: String(v[5] || '').trim(),
    terms: String(v[1] || '').split(',').map(t => t.trim()).filter(Boolean),
  })).filter(r => r.text && r.category);
}

/**
 * Чиста функція: перше активне правило, що підходить.
 * op: {desc, comment, mcc}; amount: зі знаком (0: напрям невідомий); commentOnly: лише правила для коментаря.
 */
function matchRule_(rules, op, amount, commentOnly) {
  for (const r of rules) {
    if (r.active !== 'так') continue;
    if (commentOnly && r.where !== 'коментар' && r.where !== 'будь-де') continue;
    if (amount < 0 && r.dir === 'надходження') continue;
    if (amount > 0 && r.dir === 'списання') continue;
    if (r.where === 'MCC') {
      const mcc = Number(op.mcc || 0);
      if (mcc && r.terms.some(t => {
        const m = t.match(/^(\d+)\s*-\s*(\d+)$/);
        return m ? mcc >= Number(m[1]) && mcc <= Number(m[2]) : Number(t) === mcc;
      })) return r;
      continue;
    }
    let hay;
    if (r.where === 'коментар') hay = norm_(op.comment);
    else if (r.where === 'будь-де') hay = norm_([op.desc, op.comment].join(' '));
    else hay = norm_(op.desc);
    if (r.terms.some(t => { const k = norm_(t); return k && hay.indexOf(k) >= 0; })) return r;
  }
  return null;
}

const TRANSFER_MCC = [4829, 6536, 6537, 6538, 6540];
const TRANSFER_RE = /(переказ|від:|^від |^на картку|з картки|на банку|з банки|card ?to ?card|p2p|transfer)/i;

function isTransfer_(item) {
  return TRANSFER_MCC.indexOf(Number(item.mcc)) >= 0 || TRANSFER_RE.test(String(item.description || ''));
}

/** Чиста функція: пропозиція правила після того, як ви змінили категорію рядка з банку. */
function suggestRule_(rules, desc, category) {
  const key = String(desc || '').replace(/[\d#*№]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (norm_(key).length < 3) return null;
  const hit = matchRule_(rules, { desc: desc, comment: '', mcc: 0 }, 0, false);
  if (hit && hit.category === category) return null;
  return { text: key, category: category };
}

function learnRule_(ss, desc, category, today) {
  const sh = ss.getSheetByName(RULES_SHEET);
  if (!sh) return;
  const values = sh.getRange(RULES_FIRST, 1, RULES_LAST - RULES_FIRST + 1, 7).getValues();
  const s = suggestRule_(parseRules_(values), desc, category);
  if (!s) return;
  const k = norm_(s.text);
  for (let i = 0; i < values.length; i++) {
    if (norm_(values[i][1]) !== k) continue;
    if (String(values[i][0]).indexOf('✨') === 0) sh.getRange(RULES_FIRST + i, 5).setValue(category);
    return;                                                 // ваше власне правило з таким текстом не чіпаємо
  }
  const free = values.findIndex(v => v.every(c => c === '' || c === null));
  if (free < 0) return;
  sh.getRange(RULES_FIRST + free, 1, 1, 7).setValues([[NEW_RULE, s.text, 'опис банку', 'усі', category, '', 'запропоновано ' + today]]);
}

/** Оновлення наявних «Правил» (версія 5): правила для податків ФОП, якщо таких ще немає (у вільні рядки). */
function rulesUpgrade_(ss) {
  const sh = ss.getSheetByName(RULES_SHEET), out = [];
  if (!sh) return out;
  const v = sh.getRange(RULES_FIRST, 1, RULES_LAST - RULES_FIRST + 1, 7).getValues();
  RULES_TAX.forEach(x => {
    if (v.some(r => norm_(r[1]) === norm_(x[1]))) return;
    const free = v.findIndex(r => r.every(c => c === '' || c === null));
    if (free < 0) return;
    v[free] = [x[0], x[1], x[2], x[3], x[4], '', ''];
    sh.getRange(RULES_FIRST + free, 1, 1, 7).setValues([v[free]]);
    out.push('правило «' + x[4] + '»');
  });
  return out;
}
