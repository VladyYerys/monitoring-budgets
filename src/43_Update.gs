/** 43_Update.gs: Чи вийшла нова версія шаблону, і «📦 Перенести дані зі старої копії» в нову. */

/*
 * Нова версія: раз на день скрипт читає публічний файл version.json з репозиторію (номер версії, посилання «Зробити копію»,
 * що нового). Це звичайне читання відкритого файла: таблиця нічого про себе не надсилає, автор не бачить, хто перевіряв.
 * Показується в «🩺 Стан таблиці» і один раз у ранковому Telegram.
 * Як оновитись: «Зробити копію» нового шаблону → у НОВІЙ копії «📦 Перенести дані зі старої копії» (посилання на стару).
 * Переносяться значення: журнал (з банківськими полями й рахунками), довідник, правила, налаштування, залишки банку, рахунки,
 * ФОП, борги, історія, модулі. Не переносяться: токени, розклади, підключення (їх робите заново: вони прив'язані до копії).
 * Стару копію скрипт лише читає. У таблиці з аркушами «m» і «y» (особистій) нічого не читає: одразу відмова.
 */

const VER_CACHE_MS = 20 * 3600 * 1000;

/** Порівняння версій «0.4.0»: -1, 0, 1. */
function verCmp_(a, b) {
  const p = s => String(s || '0').split('.').map(x => parseInt(x, 10) || 0), x = p(a), y = p(b);
  for (let i = 0; i < Math.max(x.length, y.length); i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d < 0 ? -1 : 1; }
  return 0;
}

/** Прочитати version.json. → { version, copy, notes[] } або виняток. */
function verFetch_() {
  let res;
  try { res = UrlFetchApp.fetch(VERSION_URL, { muteHttpExceptions: true, followRedirects: true }); }
  catch (e) { throw new Error('немає зв\'язку з GitHub: ' + String(e && e.message || e).slice(0, 120)); }
  if (res.getResponseCode() !== 200) throw new Error('GitHub відповів ' + res.getResponseCode());
  let j;
  try { j = JSON.parse(res.getContentText()); } catch (e) { throw new Error('version.json не читається'); }
  if (!j || !/^\d+\.\d+\.\d+$/.test(String(j.version || ''))) throw new Error('у version.json немає номера версії');
  const url = u => (/^https:\/\/docs\.google\.com\/spreadsheets\/d\/[\w-]+\/copy/.test(String(u || '')) ? String(u) : '');
  return { version: String(j.version), copy: url(j.copy), notes: (Array.isArray(j.notes) ? j.notes : []).slice(0, 8).map(n => String(n).slice(0, 200)) };
}

/** Кеш перевірки (властивість ver_info). force: прочитати зараз. → { t, version, copy, notes, err } */
function verCheck_(force) {
  const dp = PropertiesService.getDocumentProperties(), c = jsonProp_(dp, 'ver_info');
  if (!force && c && Date.now() - Number(c.t || 0) < VER_CACHE_MS) return c;
  let info;
  try { info = Object.assign({ t: Date.now(), err: '' }, verFetch_()); }
  catch (e) { info = Object.assign({}, c || {}, { t: Date.now(), err: String(e && e.message || e) }); }
  dp.setProperty('ver_info', JSON.stringify(info));
  return info;
}

/** З mbsTick: раз на день. */
function verIfDue_(ss) {
  const dp = PropertiesService.getDocumentProperties(), today = todayKey_(ss);
  if (dp.getProperty('ver_day') === today) return null;
  dp.setProperty('ver_day', today);
  return verCheck_(true);
}

function verNewer_(info) { return !!(info && info.version && verCmp_(info.version, TEMPLATE_VERSION) > 0); }

/** Рядки для «🩺 Стан таблиці». fetch: перевірити зараз, якщо кеш старий. */
function verLines_(ss, fetch) {
  let info = jsonProp_(PropertiesService.getDocumentProperties(), 'ver_info');
  if (fetch) { try { info = verCheck_(false); } catch (e) { /* лишаємо кеш */ } }
  if (!info || !info.version) return ['Нова версія: ще не перевірялось' + (info && info.err ? ' (⚠️ ' + esc_(info.err) + ')' : '') + '.'];
  if (!verNewer_(info)) return ['Нова версія: немає, у вас найновіша ✓' + (info.err ? ' (остання перевірка: ⚠️ ' + esc_(info.err) + ')' : '')];
  return ['🆕 <b>Вийшла версія ' + esc_(info.version) + '</b> (у вас ' + TEMPLATE_VERSION + ').']
    .concat(info.notes.length ? ['Що нового: ' + info.notes.map(esc_).join('; ')] : [])
    .concat(['Як оновитись: ' + (info.copy ? link_(info.copy, 'зробіть копію нового шаблону') : 'зробіть копію нового шаблону (посилання в README репозиторію)') +
      ', у НОВІЙ копії «🛠 Оновлення таблиці → 📦 Перенести дані зі старої копії». Стара копія лишиться як є.',
      'Або вставте новий Code.gs у «Розширення → Apps Script» цієї таблиці: ' + link_(REPO_URL + '/releases', 'випуски на GitHub') + '.']);
}

/** Ранковий Telegram: один раз про кожну нову версію (без додаткових запитів: з кешу). */
function morningVersion_(m) {
  const info = jsonProp_(m.dp, 'ver_info');
  if (!verNewer_(info) || m.dp.getProperty('ver_told') === info.version) return;
  m.lines.push('🆕 Вийшла версія шаблону ' + esc_(info.version) + ': «🩺 Стан таблиці» підкаже, як оновитись.');
  if (!m.opts.test) m.dp.setProperty('ver_told', info.version);
}

function versionNow() {
  if (!menuOk_('🔎 Нова версія')) return;
  let info;
  try { info = verCheck_(true); } catch (e) { info = { err: String(e && e.message || e) }; }
  show_('🔎 Нова версія', verLines_(SpreadsheetApp.getActive(), false).concat(info && info.err ? ['⚠️ Перевірити зараз не вдалося: ' + esc_(info.err)] : [],
    ['Перевірка: лише читання публічного файла ' + esc_(VERSION_URL.replace('https://', '')) + '; таблиця нічого про себе не надсилає.']));
}

/* ---------- перенесення даних ---------- */

/** Що переносимо: аркуш і діапазони значень (лише клітинки для вводу; формули нової копії лишаються). */
function migratePlan_() {
  const rows = (a, b) => a + ':' + b;
  return [
    { sheet: SETTINGS_SHEET, byLabel: [4, 5, 6, 7, 8, 10, 13, 14, 15, 16, 17, 18, 23, 24, 27, 28, 29, 30, 31, 32, 35] },
    { sheet: DICT_SHEET, ranges: [rows('A' + D_FIRST, 'G' + D_LAST), rows('I' + D_HOLES_FIRST, 'J' + D_HOLES_LAST), rows('K' + D_REASONS_FIRST, 'K' + D_REASONS_LAST)] },
    { sheet: RULES_SHEET, ranges: [rows('A' + RULES_FIRST, 'G' + RULES_LAST)] },
    { sheet: ACC_SHEET, ranges: [rows('A' + ACC_FIRST, 'H' + ACC_LAST)] },
    { sheet: CP_SHEET, ranges: [rows('A' + CP_FIRST, 'C' + CP_LAST)] },
    { sheet: JOURNAL, ranges: [rows('A' + J_FIRST, 'H' + J_LAST), rows(JC.bankId + J_FIRST, JC.order + J_LAST)], acct: true },
    { sheet: FOP_SHEET, ranges: [rows('A' + FOP_RATE_FIRST, 'L' + FOP_RATE_LAST), rows('A' + FOP_FIRST, 'E' + FOP_LAST), FOP_YEAR, rows('J' + FOP_Q_FIRST, 'K' + (FOP_Q_FIRST + 19))] },
    { sheet: DEBTS_SHEET, ranges: [rows('A' + DEBT_FIRST, 'H' + DEBT_LAST)] },
    { sheet: REG_SHEET, ranges: [rows('A' + REG_FIRST, 'I' + REG_LAST)] },
    { sheet: DASH_SHEET, ranges: [rows(colStr_(DASH_LIMIT_COL) + (DASH_HOLE_HEAD + 1), colStr_(DASH_LIMIT_COL) + (DASH_HOLE_HEAD + CH_HOLES))] },
    { sheet: SETTINGS_SHEET, mods: true },                               // модулі: блок на «Налаштуваннях» (у старих копіях ще на «Старті»)
    { sheet: HISTORY_SHEET, history: H_COLS.length },
    { sheet: H_CAT_SHEET, history: 6 },
  ];
}

/** Id таблиці з посилання або сам id. */
function sheetIdOf_(s) {
  const t = String(s || '').trim(), m = t.match(/\/d\/([\w-]{20,})/);
  return m ? m[1] : /^[\w-]{20,}$/.test(t) ? t : '';
}

/**
 * Перенести значення з src (стара копія шаблону) у dst (ця копія). Лише читає src. → { lines, n }
 * Аркуша чи діапазону немає в старій копії (старіша версія): пропускається з поясненням.
 */
function migrateCore_(src, dst) {
  if (isPersonalModel_(src)) throw new Error('Це схоже на особисту таблицю (є аркуші «m» і «y»): її шаблон не читає.');
  if (!hasSheetMark_(src)) throw new Error('Це не копія шаблону ' + APP_NAME + ' (немає мітки шаблону на «' + SETTINGS_SHEET + '»).');
  guard_(dst);
  const lines = [], fits = (sh, g) => g.r + g.nr - 1 <= sh.getMaxRows() && g.c + g.nc - 1 <= sh.getMaxColumns();
  let total = 0;
  migratePlan_().forEach(p => {
    const a = src.getSheetByName(p.sheet), b = dst.getSheetByName(p.sheet);
    if (!a) { lines.push('«' + p.sheet + '»: у старій копії немає, пропущено'); return; }
    if (!b) { lines.push('⚠️ «' + p.sheet + '»: у цій копії немає, пропущено'); return; }
    let cells = 0, skipped = 0;
    if (p.byLabel) {                                                    // «Налаштування»: рядок, лише якщо підпис той самий
      const la = a.getRange(1, 1, Math.min(a.getMaxRows(), 40), 2).getValues(), lb = b.getRange(1, 1, 40, 1).getValues();
      p.byLabel.forEach(r => {
        if (String(la[r - 1][0]) === '' || String(la[r - 1][0]) !== String(lb[r - 1][0])) { skipped++; return; }
        const cell = b.getRange(r, 2), dv = cell.getDataValidation();
        if (dv) cell.clearDataValidations();
        try { cell.setValue(la[r - 1][1]); } finally { if (dv) cell.setDataValidation(dv); }
        cells++;
      });
      lines.push('«' + p.sheet + '»: значень ' + cells + (skipped ? ' (немає в старій копії: ' + skipped + ')' : ''));
      total += cells;
      return;
    }
    if (p.mods) {                                                       // модулі: за назвою (порядок між версіями може змінитись)
      const sb = modBlock_(src);
      if (!sb) { lines.push('Модулі: модулів у старій копії ще не було, лишились як є'); return; }
      if (!modBlock_(dst) || modBlock_(dst).old) modBlockBuild_(dst);
      const db = modBlock_(dst), v = sb.sh.getRange(sb.first, 1, Math.min(MODS.length + 4, sb.sh.getMaxRows() - sb.first + 1), 2).getValues();
      MODS.forEach((m, i) => {
        const row = v.find(x => String(x[0]) === m.name);
        if (!row) return;
        db.sh.getRange(db.first + i, 2).setValue(String(row[1]).trim() === MOD_NO ? MOD_NO : MOD_YES);
        cells++;
      });
      lines.push('Модулі: перенесено ' + cells);
      total += cells;
      return;
    }
    if (p.history) {
      const last = a.getLastRow();
      if (last >= H_FIRST) {
        const v = a.getRange(H_FIRST, 1, last - H_FIRST + 1, p.history).getValues().filter(r => String(r[0]).trim() !== '');
        if (v.length) {
          ensureSize_(b, H_FIRST + v.length, p.history);
          const old = b.getLastRow();
          if (old >= H_FIRST) b.getRange(H_FIRST, 1, old - H_FIRST + 1, p.history).clearContent();
          b.getRange(H_FIRST, 1, v.length, 1).setNumberFormat('@');
          b.getRange(H_FIRST, 1, v.length, p.history).setValues(v.map(r => r.map((x, i) => (i === 0 && Object.prototype.toString.call(x) === '[object Date]' ?
            dateKey_(x, src.getSpreadsheetTimeZone()).slice(0, 7) : x))));
          cells += v.length;
        }
      }
      lines.push('«' + p.sheet + '»: рядків ' + cells);
      total += cells;
      return;
    }
    const ranges = p.ranges.slice();
    if (p.acct && a.getMaxColumns() >= colNum_(JC.acct) && String(a.getRange(JC.acct + J_HEAD).getValue()).indexOf('Рахунок') === 0) ranges.push(JC.acct + J_FIRST + ':' + JC.acct + J_LAST);
    ranges.forEach(r => {
      const g = a.getRange(r), w = { r: g.getRow(), c: g.getColumn(), nr: g.getNumRows(), nc: g.getNumColumns() };
      if (!fits(a, w)) { skipped++; return; }
      const v = a.getRange(w.r, w.c, w.nr, w.nc).getValues();
      ensureSize_(b, w.r + w.nr - 1, w.c + w.nc - 1);
      const dvOff = b.getRange(w.r, w.c, w.nr, w.nc).getDataValidations();     // списки з відмовою: на час запису знімаємо, потім повертаємо
      b.getRange(w.r, w.c, w.nr, w.nc).clearDataValidations();
      try { b.getRange(w.r, w.c, w.nr, w.nc).setValues(v); }
      finally { b.getRange(w.r, w.c, w.nr, w.nc).setDataValidations(dvOff); }
      cells += v.filter(x => x.some(c => c !== '' && c !== null)).length;
    });
    total += cells;
    lines.push('«' + p.sheet + '»: заповнених рядків ' + cells + (skipped ? ' (діапазонів, яких немає в старій копії: ' + skipped + ')' : ''));
  });
  return { lines: lines, n: total };
}

function migrateFromOld() {
  if (!menuOk_('📦 Перенести дані')) return;
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi(), title = '📦 Перенести дані зі старої копії';
  const a = ui.prompt(title, 'Вставте посилання на СТАРУ копію шаблону (Файл → Поділитися → Копіювати посилання або адреса з браузера).\n' +
    'Стару таблицю скрипт лише читає. Токени й розклади не переносяться: підключите заново.', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const id = sheetIdOf_(a.getResponseText());
  if (!id) { show_(title, ['⚠️ Це не схоже на посилання на Google Таблицю.']); return; }
  if (id === ss.getId()) { show_(title, ['⚠️ Це посилання на цю саму таблицю. Потрібне посилання на стару копію.']); return; }
  let src;
  try { src = SpreadsheetApp.openById(id); } catch (e) { show_(title, ['⚠️ Не вдалося відкрити: ' + esc_(e && e.message || e), 'Таблиця має бути вашою (або з доступом на перегляд).']); return; }
  if (isPersonalModel_(src)) { show_(title, ['⚠️ Це схоже на особисту таблицю (є аркуші «m» і «y»): шаблон її не читає.']); return; }
  if (!hasSheetMark_(src)) { show_(title, ['⚠️ Це не копія шаблону ' + APP_NAME + '.']); return; }
  try { schemaRun_(ss, { allowEmpty: false }); } catch (e) { show_(title, ['⚠️ ' + esc_(e.message)]); return; }
  const used = readJournal_(ss).some(o => !o.blank);
  let backup = null;
  if (used) {
    const ok = ui.alert(title, 'У цій копії журнал не порожній: його й інші дані буде замінено даними старої копії.\nСпершу зробимо копію цієї таблиці. Продовжити?', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
    try { backup = backup_(ss, 'ручна'); } catch (e) { show_(title, ['⚠️ Копію зробити не вдалося, нічого не змінено: ' + esc_(e.message)]); return; }
  }
  let r;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) { show_(title, ['⚠️ Скрипт зайнятий іншою дією, спробуйте за хвилину.']); return; }
  try { r = migrateCore_(src, ss); } catch (e) { show_(title, ['⚠️ ' + esc_(e.message)]); return; } finally { lock.releaseLock(); }
  const dp = PropertiesService.getDocumentProperties();
  dp.setProperties({ wiz_done: new Date().toISOString(), migrated_from: new Date().toISOString() });
  dp.deleteProperty('mods');                                           // модулі: з перенесених клітинок блоку «🧩 Модулі»
  MODS_CACHE_ = null;
  let mods = [];
  try { mods = modulesApply_(ss, true).filter(l => l.indexOf('⚠️') >= 0); } catch (e) { mods = ['⚠️ Модулі: ' + esc_(e.message)]; }
  try { sortJournal_(ss); } catch (e) { /* порядок: за 15 хв */ }
  try { startRefresh_(ss); } catch (e) { /* чекліст: завтра */ }
  SpreadsheetApp.flush();
  const cs = checkState_(ss);
  show_('📦 Дані перенесено', [backup ? '💾 Копія цієї таблиці перед перенесенням: ' + link_(backup.url, backup.name) : '']
    .concat(r.lines.map(esc_), mods, ['«' + CHECK_SHEET + '»: ' + (cs.none ? 'немає' : cs.bad.length ? '⚠️ ' + cs.bad.map(esc_).join('; ') : 'усе OK ✓'),
      '<b>Підключіть заново</b> (токени не переносяться): «🚀 Підключення», кроки 1 і 2 (банк, імпорт за розкладом, Telegram).',
      'Стара копія не змінилась. Коли все перевірите, її можна перейменувати («… стара») або видалити.']));
}
