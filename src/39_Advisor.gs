/** 39_Advisor.gs: Радник на Gemini (модуль «🤖 Радник Gemini»): знімок фінансів на Диску й готова навичка (skill) Gemini. */

/*
 * Знімок: Google Документ «Monitoring Budgets: знімок для радника» у папці «Monitoring Budgets: радник» на ВАШОМУ Диску.
 * Оновлюється сам щодня після 05:00 (і з меню). Що в ньому: головні цифри, місяці, бюджети, плани на 2 місяці, непомітні,
 * підписки, борги, ліміт і грейс, ФОП, сім'я по учасниках; словничок. Рівень деталей: «Налаштування», B23.
 * Чого в ньому немає ніколи: ваших коментарів, токенів, номерів карток і рахунків, ключів рахунків. Описи з банку: лише для
 * категорії «Підписки» (назва сервісу), довгі числа в них замінено на «…».
 * Навичка Gemini (skills у застосунку Gemini): готовий текст (роль, українська, лише цифри знімка з датою, без інвестиційних порад)
 * і стартові питання. Gemini читає знімок з Диска через підключений застосунок Google Диска, коли ви ставите питання.
 * Скрипт сам нікуди знімок не надсилає: до Gemini він потрапляє, лише коли ви самі питаєте навичку.
 */

const ADV_FOLDER = APP_NAME + ': радник';
const ADV_DOC = APP_NAME + ': знімок для радника';
const ADV_HOUR = 5, ADV_RETRY_MS = 3600 * 1000, ADV_OPS_DAYS = 60;
const ADV_SKILL = {
  name: 'monitoring-budgets-advisor',
  title: 'Радник Monitoring Budgets',
  description: 'Сімейний фінансовий радник українською за знімком таблиці Monitoring Budgets з Google Диска: бюджети, витрати, плани, борги, ліміт, ФОП.',
  starters: ['Де ми переплачуємо за останні 3 місяці?', 'Чи вистачить на подарунки в грудні, якщо відкласти 5 000 ₴?', 'Скільки відкласти на податки ФОП до кінця кварталу?',
    'Які підписки варто скасувати?', 'Скільки можна витрачати на день до кінця місяця?', 'Як швидше погасити кредитний ліміт без комісій?'],
};

/** Текст інструкції навички (без заголовка SKILL.md). */
function advSkillInstructions_() {
  return [
    'Ти сімейний фінансовий помічник для таблиці «' + APP_NAME + '». Відповідаєш українською, коротко й по суті, у гривнях.',
    '',
    'Звідки дані:',
    '1. Перед кожною відповіддю знайди на моєму Google Диску документ «' + ADV_DOC + '» (папка «' + ADV_FOLDER + '») через Google Диск і прочитай найновішу версію.',
    '2. Відповідай лише цифрами з цього знімка. На початку відповіді назви дату знімка («за знімком від …»). Якщо знімку понад 3 дні, скажи про це.',
    '3. Чого немає в знімку, не вигадуй: скажи, чого бракує і де це заповнити в таблиці (аркуш і колонка).',
    '4. Розрахунки показуй коротко: звідки число і як порахував. Округлюй до гривні.',
    '',
    'Як радити:',
    '- Спершу головне: «мої гроші» на сьогодні, можна на день, прогноз на кінець місяця, бюджети, що перевищені або йдуть швидше норми.',
    '- Пропонуй 2–3 конкретні кроки (яку категорію зменшити й на скільки, що перенести, що відкласти), без моралізаторства.',
    '- Кредитний ліміт і грейс: пояснюй, скільки внести до строку, щоб не платити відсотки, і яка комісія, якщо перекривати переказами.',
    '- ФОП: суми й строки з знімка; остаточні податкові рішення: з бухгалтером. Це не податкова консультація.',
    '- Не давай інвестиційних порад: не радь конкретних акцій, облігацій, криптовалют, фондів чи банків.',
    '- Ніколи не проси номери карток, IBAN, токени, паролі чи коди з SMS. Якщо їх надіслали, попроси видалити.',
    '- Сімейний режим: говори і про кожного учасника, і про сім\'ю разом; без осуду.',
    '',
    'Стартові питання (можна запропонувати, якщо питання немає):',
  ].concat(ADV_SKILL.starters.map(q => '- ' + q)).join('\n');
}

/** SKILL.md для імпорту в Gemini («Skills → Create → Upload») і для репозиторію (gemini/…/SKILL.md). */
function advSkillMd_() {
  return '---\nname: ' + ADV_SKILL.name + '\ndescription: ' + ADV_SKILL.description + '\n---\n\n# ' + ADV_SKILL.title + '\n\n' + advSkillInstructions_() + '\n';
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
    'Це знімок, не жива таблиця: відповідай лише цими цифрами й називай дату знімка.', '');
  out.push('## Як читати',
    '- «Мої гроші»: надходження бюджетного місяця (зарплата тощо) мінус відкладене на погашення мінус уже витрачене.' + (credit ? ' Кредитний ліміт до них не входить.' : ''),
    '- Бюджетний місяць: дохід з ' + bday + '-го числа йде в бюджет наступного місяця.',
    '- Факт: що вже сталося. План: майбутній платіж чи надходження. Чернетка: операція з банку, ще не підтверджена (у підсумки не входить).',
    '- «Непомітні»: дрібні витрати за звичкою (кава, доставка, таксі, підписки, дрібні покупки, комісії).',
    credit ? '- Кредитний ліміт: гроші банку. «З ліміту»: витрачено понад «мої». Грейс Mono: борг місяця без відсотків до кінця наступного місяця, якщо внести до строку; перекриття переказами з ліміту коштує комісію.' : '',
    modOn_('fop') ? '- ФОП: дохід ФОП і податки рахуються окремо від сімейного бюджету; у сім\'ю йде лише переказ з рахунку ФОП на картку.' : '', '');
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
      (o.g ? ' · ' + o.g : '') + (o.d.indexOf(FOP_TAG) === 0 || o.d.indexOf(DEBT_TAG) === 0 ? ' · ' + o.d : '')));
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

/** Вікно: посилання на знімок, кроки створення навички й текст для копіювання. */
function advisorNow() {
  if (!menuOk_('🤖 Радник Gemini')) return;
  const ss = SpreadsheetApp.getActive(), doc = advDoc_(ss, false);
  const area = t => '<textarea readonly style="width:100%;height:' + (t.split('\n').length > 3 ? 170 : 44) + 'px;font:12px/1.4 monospace" onclick="this.select()">' + esc_(t) + '</textarea>';
  const lines = [
    modOn_('advisor') ? '' : '⚠️ Модуль «🤖 Радник Gemini» вимкнено: увімкніть на «' + START_SHEET + '» або натисніть «🔄 Оновити знімок зараз».',
    doc ? 'Знімок: ' + link_(doc.getUrl(), ADV_DOC) + ' (на вашому Диску, оновлюється щодня).' : 'Знімка ще немає: «🤖 Радник Gemini → 🔄 Оновити знімок зараз».',
    '<b>Як створити навичку</b> (gemini.google.com, особистий акаунт Google, 18+):',
    '1. Бокова панель → <b>Customize</b> (або More) → <b>Skills</b> → <b>Create</b> → <b>Create manually</b>.',
    '2. Назва:' + area(ADV_SKILL.title) + 'Опис:' + area(ADV_SKILL.description) + 'Інструкції:' + area(advSkillInstructions_()),
    '3. Або завантажте готовий файл: ' + link_(REPO_URL + '/blob/main/gemini/' + ADV_SKILL.name + '/SKILL.md', 'SKILL.md у репозиторії') + ' (Create → Upload; на Mac і в браузері).',
    '4. У чаті Gemini: «/» і назва навички, потім питання. Gemini знайде знімок на вашому Диску (застосунок Google Диска має бути підключений).',
    '<b>Стартові питання</b>: ' + ADV_SKILL.starters.map(esc_).join(' · '),
    'Приватність: знімок лежить лише на вашому Диску; скрипт нікуди його не надсилає. Gemini читає його, лише коли ви питаєте навичку. Коментарів, номерів карток і токенів у ньому немає.',
  ];
  const html = HtmlService.createHtmlOutput('<div style="font:14px/1.5 Arial,sans-serif">' + lines.filter(Boolean).map(l => '<p>' + l + '</p>').join('') + '</div>')
    .setWidth(640).setHeight(620);
  try { SpreadsheetApp.getUi().showModalDialog(html, '🤖 Радник Gemini'); } catch (e) { Logger.log(advSkillMd_()); }
}
