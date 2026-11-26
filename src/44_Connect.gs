/** 44_Connect.gs: «🚀 Підключення»: бічна панель з чотирма кроками замість анкети: картка, Telegram, регулярні, Gemini. */

/*
 * Панель (HTML у 45_ConnectUi.gs) викликає функції cn… через google.script.run. Кожна спершу перевіряє таблицю
 * (guard_): у чужій таблиці й у таблиці з «m» і «y» нічого не робить і повертає помилку.
 * Питань майже немає: відповіді беруться з банку.
 *   1. 💳 Картка: токен Mono (QR у застосунку на api.monobank.ua) → ваші картки й рахунки з галочками. З них: вид картки
 *      й ліміт (модуль «Ліміт»), рахунок ФОП (модуль «ФОП», лишається вказати групу), імпорт за розкладом. Інше: Приват
 *      для ФОП, виписка з папки, «без банку» (лише вид картки). День доходу: поле тут же.
 *   2. 📨 Telegram: токен від @BotFather (QR веде прямо в BotFather), далі QR «Відкрити бота» → Start; панель сама бачить
 *      натискання (getUpdates) і надсилає перше повідомлення. Окремий QR для сімейної групи (startgroup).
 *   3. 🔁 Регулярні: пошук повторюваних платежів і надходжень в операціях (журнал; з Mono ще й виписка за 2 місяці
 *      без запису в журнал) → галочки → аркуш «Регулярні». Свій платіж: форма.
 *   4. 🤖 Gemini: знімок на Диску, посилання на Gemini, назва, опис і інструкція навички для копіювання.
 * Токени, як і раніше, лише в UserProperties вашого акаунта; у панель вони назад не повертаються.
 * Режим «я сам / сім'я» і учасники: перемикач угорі панелі. Модулі вмикаються самі за підключеним (змінити: «Старт»).
 */

const CN_TITLE = '🚀 Підключення';
const CN_BOTFATHER = 'https://t.me/BotFather', CN_TG = 'https://t.me/', CN_MONO = 'https://api.monobank.ua/', CN_GEMINI = 'https://gemini.google.com/';
const CN_SCAN_DAYS = 124, CN_MONO_DAYS = 62;

/** Модуль так / ні через клітинку «Старту»: тоді modulesApply_ побачить зміну й застосує її (аркуші, рядки, розклади). */
function cnMod_(ss, id, on) {
  const i = MODS.findIndex(m => m.id === id);
  if (i >= 0 && startHasModules_(ss)) ss.getSheetByName(START_SHEET).getRange(MOD_FIRST + i, 2).setValue(on ? MOD_YES : MOD_NO);
  else modSet_(ss, id, on);
}

/** Таблиця для дій панелі: наша або помилка (панель покаже текст). */
function cnSs_() {
  const ss = SpreadsheetApp.getActive();
  guard_(ss);
  return ss;
}

/** Меню «🚀 Підключення» (і «🚀 Почати» в порожній таблиці): добудувати таблицю, якщо треба, і відкрити панель. */
function startWizard() {
  const ss = SpreadsheetApp.getActive();
  try { guard_(ss, true); } catch (e) { show_(CN_TITLE, ['⚠️ ' + esc_(e.message)]); return; }
  try { schemaRun_(ss, { allowEmpty: true }); } catch (e) { show_(CN_TITLE, ['⚠️ ' + esc_(e.message)]); return; }
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(cnHtml_()).setTitle(CN_TITLE));
}

/* ---------- стан для панелі ---------- */

function cnState() {
  const ss = cnSs_(), dp = PropertiesService.getDocumentProperties(), up = PropertiesService.getUserProperties();
  const set = ss.getSheetByName(SETTINGS_SHEET), accs = accList_(ss), fops = fopList_(ss);
  const reg = ss.getSheetByName(REG_SHEET) ? regOwn_(ss) : [];
  const regState = ss.getSheetByName(REG_SHEET) ? ss.getSheetByName(REG_SHEET).getRange(REG_FIRST, 10, REG_LAST - REG_FIRST + 1, 1).getValues() : [];
  const tgBot = up.getProperty('tg_bot') || '', tgCode = up.getProperty('tg_link') || '';
  const doc = advDoc_(ss, false);
  return {
    family: familyMode_(ss), members: members_(ss), who: familyMode_(ss) ? members_(ss).concat([WHO_SHARED]) : [],
    bank: {
      mono: up.getProperty('mono_token') ? monoAccs_().map(a => a.name) : [], privat: privatAccs_().map(a => a.name),
      accounts: accs.map(a => ({ name: a.name, kind: a.kind, who: a.who })), schedule: hasTrigger_(TRG.imp),
      credit: cardIsCredit_(ss), bday: Number(set.getRange(SET.bday).getValue()) || 25,
      fopAcc: accs.some(a => a.kind === K_FOP), fops: fops.map(f => f.name + ' (' + f.group + ' група)'),
      last: dp.getProperty('bank_last_ok') || '',
    },
    tg: { bot: tgBot, connected: tgConnected_(), on: modOn_('tg'), link: tgBot && tgCode && !tgConnected_() ? cnTgLinks_(tgBot, tgCode) : null },
    qrBotFather: qrSvg_(CN_BOTFATHER, 132),
    reg: { items: reg.map(r => ({ name: r.name, cat: r.cat, amount: r.amt, day: r.day, per: r.per, on: r.on, state: String((regState[r.i] || [''])[0]) })) },
    cats: cnCats_(ss),
    adv: { on: modOn_('advisor'), doc: doc ? doc.getUrl() : '', t: dp.getProperty('adv_t') || '' },
    demo: !!jsonProp_(dp, 'demo'), empty: !readJournal_(ss).some(o => !o.blank), done: !!dp.getProperty('wiz_done'),
    links: { mono: CN_MONO, botfather: CN_BOTFATHER, gemini: CN_GEMINI },
  };
}

/** Категорії з «Довідника» для списків: [{ name, type }]. */
function cnCats_(ss) {
  const d = ss.getSheetByName(DICT_SHEET);
  if (!d) return [];
  return d.getRange(D_FIRST, 1, D_LAST - D_FIRST + 1, 2).getValues().filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), type: String(r[1]).trim() }));
}

/** Режим: одна людина чи сім'я (учасники через кому, до 5). */
function cnMode(family, names) {
  const ss = cnSs_(), set = ss.getSheetByName(SETTINGS_SHEET);
  if (family === true) {
    const list = String(names || '').split(',').map(s => s.trim()).filter(Boolean).slice(0, SET_MEMBERS_N);
    set.getRange(SET.mode).setValue(MODE_FAMILY);
    if (list.length) set.getRange(SET_MEMBERS_FIRST, 2, SET_MEMBERS_N, 1).setValues(Array.from({ length: SET_MEMBERS_N }, (x, i) => [list[i] || '']));
  } else if (family === false) set.getRange(SET.mode).setValue(MODE_SOLO);
  return cnState();
}

/* ---------- 1. картка ---------- */

/** Перевірити токен Mono (нічого не зберігає): гривневі рахунки для галочок. */
function cnMonoCheck(token) {
  cnSs_();
  token = String(token || '').trim();
  if (!/^[A-Za-z0-9_\-]{20,}$/.test(token)) throw new Error('Це не схоже на токен Mono (довгий рядок латинських літер і цифр).');
  const info = monoFetch_(token, '/personal/client-info'), cands = cardCandidates_(info.accounts, true);
  if (!cands.length) throw new Error('Гривневих рахунків не знайдено.');
  return cands.map((a, i) => ({ id: a.id, label: a.type === 'fop' ? 'Mono ФОП *' + cardLast4_(a.iban) : cardLabelOf_(a), fop: a.type === 'fop',
    limit: Math.round(Number(a.creditLimit || 0) / 100), pick: i === 0 || a.type === 'fop' }));
}

/** Зберегти вибрані рахунки Mono: «Рахунки», токен у вашому акаунті, вид картки, модулі «Ліміт» і «ФОП», імпорт за розкладом. */
function cnMonoSave(token, ids, who) {
  const ss = cnSs_();
  token = String(token || '').trim();
  const info = monoFetch_(token, '/personal/client-info'), cands = cardCandidates_(info.accounts, true);
  const chosen = cands.filter(a => (ids || []).indexOf(a.id) >= 0);
  if (!chosen.length) throw new Error('Позначте хоча б один рахунок.');
  who = familyMode_(ss) && (members_(ss).indexOf(who) >= 0 || who === WHO_SHARED) ? who : '';
  const saved = chosen.map(a => {
    const fop = a.type === 'fop', acc = accEnsure_(ss, { key: 'mono:' + md5Short_(a.id), name: fop ? 'Mono ФОП *' + cardLast4_(a.iban) : cardLabelOf_(a), bank: 'Mono',
      who: who, kind: fop ? K_FOP : undefined });
    return { id: a.id, iban: a.iban || '', type: a.type || '', last4: cardLast4_(fop ? a.iban : (a.maskedPan || [])[0]), name: acc.name, kind: acc.kind, credit: Number(a.creditLimit) > 0 };
  });
  const up = PropertiesService.getUserProperties();
  up.setProperties({ mono_token: token, mono_accs: JSON.stringify(saved.map(s => ({ id: s.id, iban: s.iban, type: s.type, last4: s.last4, name: s.name }))) });
  ['mono_account', 'mono_iban', 'mono_type', 'mono_pan'].forEach(k => up.deleteProperty(k));
  const main = saved.find(s => s.kind === K_MAIN);
  if (main) { ss.getSheetByName(SETTINGS_SHEET).getRange(SET.card).setValue(main.credit ? CARD_CREDIT : CARD_DEBIT); cnMod_(ss, 'credit', main.credit); }
  if (saved.some(s => s.kind === K_FOP)) cnMod_(ss, 'fop', true);
  importTriggers_(ss.getSpreadsheetTimeZone());
  modulesApply_(ss, false);
  return { saved: saved.map(s => s.name + ' (' + s.kind + ')'), credit: main ? main.credit : null, fop: saved.some(s => s.kind === K_FOP) };
}

/** Перший імпорт з банку (сьогодні й 2 дні перед тим). */
function cnImport() {
  const ss = cnSs_(), r = importBank_(ss);
  return { text: importSummary_(r), drafts: r.drafts, own: r.own, limit: r.limit, errors: r.errors || [] };
}

/** Без банку чи вручну: вид картки (кредитна з лімітом / дебетова) і модуль «Ліміт». */
function cnCard(credit) {
  const ss = cnSs_();
  ss.getSheetByName(SETTINGS_SHEET).getRange(SET.card).setValue(credit ? CARD_CREDIT : CARD_DEBIT);
  cnMod_(ss, 'credit', !!credit);
  modulesApply_(ss, false);
  return cnState();
}

/** З якого числа дохід іде в наступний місяць. */
function cnBday(day) {
  const ss = cnSs_(), d = Math.round(Number(day));
  if (!(d >= 1 && d <= 31)) throw new Error('День: число від 1 до 31.');
  ss.getSheetByName(SETTINGS_SHEET).getRange(SET.bday).setValue(d);
  return cnState();
}

/** Рахунок ФОП знайдено: група (1, 2, 3) у «ФОП у сім'ї». */
function cnFop(group, who) {
  const ss = cnSs_(), g = String(group);
  if (['1', '2', '3'].indexOf(g) < 0) throw new Error('Група ФОП: 1, 2 або 3.');
  const sh = ss.getSheetByName(FOP_SHEET) || buildFop_(ss), name = String(who || '').trim() || (members_(ss)[0] || 'Я');
  const v = sh.getRange(FOP_FIRST, 1, FOP_LAST - FOP_FIRST + 1, 2).getValues();
  let i = v.findIndex(r => String(r[0]).trim() === name);
  if (i < 0) i = v.findIndex(r => !String(r[0]).trim());
  if (i < 0) throw new Error('На «' + FOP_SHEET + '» уже 5 ФОП.');
  sh.getRange(FOP_FIRST + i, 1, 1, 5).setValues([[name, g, v[i][0] ? sh.getRange(FOP_FIRST + i, 3).getValue() : '', 'так', 'так']]);
  cnMod_(ss, 'fop', true);
  modulesApply_(ss, false);
  return cnState();
}

/** Імпорт за розкладом: так / ні (для ваших підключених банків). */
function cnSchedule(on) {
  const ss = cnSs_(), up = PropertiesService.getUserProperties();
  if (on === true) {
    if (!up.getProperty('mono_token') && !up.getProperty('pb_token')) throw new Error('Спершу підключіть Mono або ПриватБанк.');
    importTriggers_(ss.getSpreadsheetTimeZone());
  } else if (on === false) { removeTriggers_(TRG.imp); removeTriggers_(TRG.impRetry); }
  return cnState();
}

/* ---------- 2. Telegram ---------- */

function cnTgLinks_(bot, code) {
  const p = CN_TG + bot + '?start=' + code, g = CN_TG + bot + '?startgroup=' + code;
  return { bot: bot, private: p, group: g, qr: qrSvg_(p, 160), qrGroup: qrSvg_(g, 160) };
}

/** Токен від @BotFather: перевірити (getMe), зберегти у вашому акаунті, дати посилання й QR «Відкрити бота». */
function cnTgToken(token) {
  cnSs_();
  token = String(token || '').trim();
  if (!/^\d{5,15}:[A-Za-z0-9_-]{30,}$/.test(token)) throw new Error('Це не схоже на токен бота. Скопіюйте його з @BotFather повністю (вигляд 123456789:AAE…).');
  const me = tgCall_(token, 'getMe');
  const code = Utilities.getUuid().replace(/-/g, '').slice(0, 12), up = PropertiesService.getUserProperties();
  up.setProperties({ tg_token: token, tg_bot: String(me.username || ''), tg_link: code, tg_link_t: String(Date.now()) });
  up.deleteProperty('tg_chat');
  return cnTgLinks_(String(me.username || ''), code);
}

/** Чистa функція: з оновлень бота знайти /start: спершу з нашим кодом, інакше свіжий /start в особистому чаті. */
function cnTgFind_(updates, code, sinceSec) {
  const msgs = (updates || []).map(u => u.message).filter(m => m && m.chat && ['private', 'group', 'supergroup'].indexOf(m.chat.type) >= 0);
  const start = m => /^\/start(@\w+)?(\s+(\S+))?/.exec(String(m.text || ''));
  const withCode = msgs.filter(m => { const s = start(m); return s && s[3] === code; });
  if (withCode.length) return withCode[withCode.length - 1];
  const fresh = msgs.filter(m => start(m) && m.chat.type === 'private' && Number(m.date || 0) >= sinceSec);
  return fresh.length ? fresh[fresh.length - 1] : null;
}

/** Панель питає кожні 3 с: натиснули Start? Так: чат збережено, модуль увімкнено, перше повідомлення надіслано. */
function cnTgPoll() {
  const ss = cnSs_(), up = PropertiesService.getUserProperties(), token = up.getProperty('tg_token');
  if (!token) throw new Error('Спершу вставте токен бота.');
  if (up.getProperty('tg_chat')) return { ok: true, chat: 'уже підключено' };
  const upd = tgCall_(token, 'getUpdates', { limit: 100, allowed_updates: ['message'] }) || [];
  const m = cnTgFind_(upd, up.getProperty('tg_link') || '', Math.floor(Number(up.getProperty('tg_link_t') || 0) / 1000) - 120);
  if (!m) return { ok: false };
  up.setProperty('tg_chat', String(m.chat.id));
  cnMod_(ss, 'tg', true);
  modulesApply_(ss, false);
  if (!hasTrigger_(TRG.morning)) ensureMorningTrigger_();
  try { tgCall_(token, 'getUpdates', { offset: upd[upd.length - 1].update_id + 1, limit: 1 }); } catch (e) { /* черга не критична */ }
  tgSend_('✅ <b>' + APP_NAME + '</b> підключено. Підсумок: щодня о ' + pad2_(TG_HOUR) + ':00.');
  return { ok: true, chat: m.chat.type === 'private' ? 'особисті повідомлення' : 'група «' + String(m.chat.title || '') + '»' };
}

/* ---------- 3. регулярні ---------- */

/** Ключ групи операцій: опис без цифр і зайвого (або категорія). */
function cnOpKey_(desc, cat) {
  const d = String(desc || '').toLowerCase().replace(/\d+/g, ' ').replace(/[^a-zа-яіїєґ' ]+/gi, ' ').replace(/\s+/g, ' ').trim().slice(0, 30);
  return d ? 'd:' + d : cat ? 'c:' + cat : '';
}

/**
 * Чиста функція: повторювані платежі й надходження. ops: [{ key, amt (зі знаком), desc, cat }]. Група: той самий опис
 * (або категорія), той самий напрям; щонайменше 2 різні місяці; не частіше разу на місяць; суми ±25 % від медіани;
 * дні ±5 від медіанного. Ті, що вже є на «Регулярних» (категорія й сума ±25 % або назва), пропускаються.
 * → [{ name, cat, amount, day, months, income }] від більших сум.
 */
function regSuggest_(ops, existing) {
  const groups = {};
  ops.forEach(o => {
    if (!o.key || !o.amt) return;
    const k = cnOpKey_(o.desc, o.cat);
    if (!k) return;
    const g = (o.amt > 0 ? '+' : '-') + k;
    (groups[g] = groups[g] || []).push(o);
  });
  const med = a => { const s = a.slice().sort((x, y) => x - y); return s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2; };
  const out = [];
  Object.keys(groups).forEach(g => {
    const list = groups[g], months = {};
    list.forEach(o => { (months[o.key.slice(0, 7)] = months[o.key.slice(0, 7)] || []).push(o); });
    const ms = Object.keys(months);
    if (ms.length < 2 || list.length > ms.length + 1) return;
    const m = med(list.map(o => Math.abs(o.amt)));
    const per = ms.map(k => months[k].slice().sort((a, b) => Math.abs(Math.abs(a.amt) - m) - Math.abs(Math.abs(b.amt) - m))[0]);
    if (per.some(o => Math.abs(Math.abs(o.amt) - m) > m * REG_TOL + 0.005)) return;
    const days = per.map(o => Number(o.key.slice(8))), dm = Math.round(med(days));
    if (days.some(d => Math.min(Math.abs(d - dm), 31 - Math.abs(d - dm)) > REG_CLOSE_DAYS)) return;
    const cats = {};
    list.forEach(o => { if (o.cat && o.cat !== CAT_TRANSFER) cats[o.cat] = (cats[o.cat] || 0) + 1; });
    const cat = Object.keys(cats).sort((a, b) => cats[b] - cats[a])[0] || '';
    const name = (String(list[list.length - 1].desc || '').replace(/\s+/g, ' ').trim() || cat).slice(0, 40);
    const amount = Math.round(m * 100) / 100;
    if ((existing || []).some(r => (r.name && r.name.toLowerCase() === name.toLowerCase()) ||
      (cat && r.cat === cat && Math.abs(r.amt - amount) <= r.amt * REG_TOL + 0.005))) return;
    out.push({ name: name, cat: cat, amount: amount, day: Math.max(1, Math.min(31, dm)), months: ms.length, income: g.charAt(0) === '+' });
  });
  return out.sort((a, b) => b.amount - a.amount).slice(0, 20);
}

/**
 * Знайти регулярні. deep: ще й виписка основної картки Mono за 2 місяці (лише читання, у журнал нічого не пишеться;
 * Mono дає 31 день за запит і не частіше разу на хвилину, тож ≈ 1 хв). → { items, days, canDeep, income }
 */
function cnRegSuggest(deep) {
  const ss = cnSs_(), today = todayKey_(ss), from = addDays_(today, -CN_SCAN_DAYS), seen = {};
  const ops = [];
  readJournal_(ss).forEach(o => {
    if (o.blank || !o.key || o.key < from || o.key > today || isPlan_(o.e) || !(Number(o.b) > 0) || o.kind === K_FOP) return;
    if (o.bank && o.bank[0]) seen[String(o.bank[0]).replace(/#fee$/, '')] = 1;
    ops.push({ key: o.key, amt: (TYPES_IN.indexOf(o.i) >= 0 ? 1 : -1) * Number(o.b), desc: String((o.bank && o.bank[1]) || o.d || ''), cat: o.c });
  });
  const token = PropertiesService.getUserProperties().getProperty('mono_token'), main = accList_(ss).find(a => a.kind === K_MAIN && a.bank === 'Mono');
  let monoDays = 0;
  if (deep === true && token && main) {
    const sv = monoAccs_().find(s => 'mono:' + md5Short_(s.id) === main.key);
    if (sv) {
      const tz = ss.getSpreadsheetTimeZone(), now = Math.floor(Date.now() / 1000), rules = readRules_(ss);
      const dateOf = it => Utilities.formatDate(new Date(it.time * 1000), tz, 'yyyy-MM-dd');
      const mid = now - 31 * 86400, start = now - CN_MONO_DAYS * 86400;
      const items = fetchStatement_(token, sv.id, mid, now);
      Utilities.sleep(61 * 1000);
      fetchStatement_(token, sv.id, start, mid).forEach(it => items.push(it));
      items.forEach(it => {
        if (seen[it.id] || isPendingReversal_(it)) return;
        seen[it.id] = 1;
        const amt = Number(it.amount) / 100;
        ops.push({ key: dateOf(it), amt: amt, desc: String(it.description || ''), cat: categorize_(it, amt, rules, main).cat });
      });
      monoDays = CN_MONO_DAYS;
    }
  }
  const keys = ops.map(o => o.key).sort(), days = keys.length ? dayDiff_(today, keys[0]) : 0;
  const items = regSuggest_(ops, ss.getSheetByName(REG_SHEET) ? regOwn_(ss) : []);
  const inc = items.filter(x => x.income).sort((a, b) => b.amount - a.amount)[0] || null;
  return { items: items, days: Math.max(days, monoDays), canDeep: !!(token && main) && deep !== true && days < 45,
    income: inc ? { day: inc.day, name: inc.name, amount: inc.amount } : null };
}

/** Додати на «Регулярні»: [{ name, cat, amount, day, per }] → рядки в першій вільній частині; плани в журналі одразу. */
function cnRegAdd(list) {
  const ss = cnSs_(), sh = ss.getSheetByName(REG_SHEET) || buildRegular_(ss), cats = cnCats_(ss).map(c => c.name), bad = [];
  const rows = (Array.isArray(list) ? list : []).filter(x => {
    const ok = x && String(x.name || '').trim() && cats.indexOf(String(x.cat)) >= 0 && Number(x.amount) > 0 && Number(x.day) >= 1 && Number(x.day) <= 31;
    if (!ok && x && x.name) bad.push(String(x.name));
    return ok;
  });
  const v = sh.getRange(REG_FIRST, 1, REG_LAST - REG_FIRST + 1, 1).getValues();
  let k = 0;
  v.forEach((r, i) => {
    if (k >= rows.length || String(r[0]).trim()) return;
    const x = rows[k++];
    sh.getRange(REG_FIRST + i, 1, 1, 9).setValues([[String(x.name).trim().slice(0, 60), x.cat, round2_(Number(x.amount)), Math.round(Number(x.day)),
      REG_PER.indexOf(x.per) >= 0 ? x.per : REG_PER[0], '', '', familyMode_(ss) && x.who ? x.who : '', MOD_YES]]);
  });
  const st = k ? regPlansSync_(ss) : { added: 0 };
  return { added: k, plans: st.added, skipped: bad, full: k < rows.length, state: cnState() };
}

/* ---------- 4. Gemini ---------- */

/** Знімок для радника на вашому Диску й текст навички. */
function cnAdvisor() {
  const ss = cnSs_();
  if (!modOn_('advisor')) modSet_(ss, 'advisor', true);                  // знімок пишемо нижче самі
  const r = snapshotWrite_(ss);
  return { doc: r.url, docName: ADV_DOC, chars: r.chars, title: ADV_SKILL.title, description: ADV_SKILL.description, instructions: advSkillInstructions_(),
    starters: ADV_SKILL.starters, gemini: CN_GEMINI, skillFile: REPO_URL + '/blob/main/gemini/' + ADV_SKILL.name + '/SKILL.md' };
}

/* ---------- кінець і демо ---------- */

/** «Готово»: позначка, модулі, чекліст «Старту». */
function cnFinish() {
  const ss = cnSs_(), dp = PropertiesService.getDocumentProperties();
  dp.setProperty('wiz_done', new Date().toISOString());
  let lines = [];
  try { lines = modulesApply_(ss, false).filter(l => l.indexOf('⚠️') >= 0); } catch (e) { lines = ['⚠️ ' + String(e && e.message || e)]; }
  const r = startRefresh_(ss);
  return { done: r ? r.done : 0, need: r ? r.need : 0, warnings: lines };
}

/** Демо-місяць: додати (true) чи прибрати (false). */
function cnDemo(on) {
  const ss = cnSs_();
  if (on === true) { if (!jsonProp_(PropertiesService.getDocumentProperties(), 'demo')) demoAdd_(ss); }
  else if (on === false) demoClear_(ss);
  return cnState();
}
