/** 35_Statements.gs: Виписки CSV чи Google Таблиці з папки на Диску для банків без API (Приват24, Ощад, ПУМБ, Sense, А-Банк…). */

/*
 * Папка «Monitoring Budgets: виписки» з'являється поруч із таблицею (скрипт бере лише свою папку, за збереженим id).
 * Ви кладете туди виписку: CSV (будь-яке кодування, «,» «;» чи табуляція) або Google Таблицю. Excel (.xlsx): відкрийте його
 * в Google Таблицях і «Файл → Зберегти як Google Таблиці», або збережіть з банку як CSV.
 * Колонки розпізнаються за заголовками (дата, час, сума або списання / зарахування, опис, MCC); заголовок може бути не в першому рядку.
 * Скрипт лише читає файли з цієї папки й нічого в них не змінює. Таблиці з аркушами «m» і «y» не читає взагалі.
 * Та сама операція з виписки й з API (Mono, Приват) не задвоюється: однакова сума зі знаком і дата ±1 день на тому ж рахунку.
 */

const STMT_FOLDER = APP_NAME + ': виписки';
const STMT_COLS = {
  date: ['дата і час операції', 'дата i час операції', 'дата операції', 'дата транзакції', 'дата', 'date'],
  time: ['час операції', 'час', 'time'],
  amount: ['сума в валюті картки', 'сума в валюті рахунку', 'сума операції', 'сума, грн', 'сума (грн)', 'сума', 'amount'],
  debit: ['списання', 'дебет', 'витрата', 'debit'],
  credit: ['зарахування', 'надходження', 'кредит', 'credit'],
  desc: ['деталі операції', 'опис операції', 'призначення платежу', 'призначення', 'опис', 'деталі', 'description', 'контрагент'],
  mcc: ['mcc'],
  note: ['категорія', 'коментар', 'comment'],
};

/** Наша папка: за збереженим id; інакше нова поруч із таблицею (чужу папку з такою ж назвою не беремо). */
function stmtFolder_(ss) {
  const dp = PropertiesService.getDocumentProperties(), id = dp.getProperty('stmt_folder_id');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* видалили */ } }
  let parent;
  try { const it = DriveApp.getFileById(ss.getId()).getParents(); parent = it.hasNext() ? it.next() : DriveApp.getRootFolder(); }
  catch (e) { parent = DriveApp.getRootFolder(); }
  const folder = parent.createFolder(STMT_FOLDER);
  dp.setProperty('stmt_folder_id', folder.getId());
  return folder;
}

/** Чиста функція: число з тексту банку («−1 234,56», «1,234.56», «-12.5 грн»). NaN: не число. */
function stmtNum_(v) {
  if (typeof v === 'number') return v;
  let s = String(v === null || v === undefined ? '' : v).replace(/[\s ]/g, '').replace(/[−–]/g, '-').replace(/(грн|uah|₴)/gi, '');
  if (!s) return NaN;
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else s = s.replace(',', '.');
  return /^[+-]?\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
}

/** Чиста функція: дата (і час) з клітинки → { key: 'yyyy-MM-dd', hm: 'HH:mm' } або null. */
function stmtDate_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') return { key: Utilities.formatDate(v, tz, 'yyyy-MM-dd'), hm: Utilities.formatDate(v, tz, 'HH:mm') };
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})(?:[ T,]+(\d{1,2}):(\d{2}))?/);
  if (m) return { key: m[3] + '-' + pad2_(+m[2]) + '-' + pad2_(+m[1]), hm: m[4] ? pad2_(+m[4]) + ':' + m[5] : '' };
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
  if (m) return { key: m[1] + '-' + m[2] + '-' + m[3], hm: m[4] ? m[4] + ':' + m[5] : '' };
  return null;
}

/** Чиста функція: де заголовок і які колонки. → { row, cols: {date, amount|debit+credit, …} } або null. */
function stmtHeader_(rows) {
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const h = rows[i].map(c => String(c === null || c === undefined ? '' : c).trim().toLowerCase().replace(/\s+/g, ' '));
    const cols = {};
    Object.keys(STMT_COLS).forEach(k => {
      for (const word of STMT_COLS[k]) {
        const j = h.findIndex((c, jj) => c && (c === word || c.indexOf(word) === 0) && Object.keys(cols).every(x => cols[x] !== jj));
        if (j >= 0) { cols[k] = j; break; }
      }
    });
    if (cols.date !== undefined && (cols.amount !== undefined || (cols.debit !== undefined && cols.credit !== undefined))) return { row: i, cols: cols };
  }
  return null;
}

/**
 * Чиста функція: рядки виписки → операції загального вигляду. Номер: «csv:» + відбиток (рахунок, дата, час, сума, опис, № повтору),
 * тож та сама виписка вдруге нічого не додає. → { items, skipped, header }
 */
function parseStatement_(rows, tz, accName, digest) {
  const hd = stmtHeader_(rows);
  if (!hd) return { items: [], skipped: rows.length, header: null };
  const c = hd.cols, items = [], seen = {};
  let skipped = 0;
  rows.slice(hd.row + 1).forEach(r => {
    const d = stmtDate_(r[c.date], tz);
    let amt = c.amount !== undefined ? stmtNum_(r[c.amount]) : NaN;
    if (c.amount === undefined || isNaN(amt)) {
      const deb = c.debit !== undefined ? stmtNum_(r[c.debit]) : NaN, cred = c.credit !== undefined ? stmtNum_(r[c.credit]) : NaN;
      amt = (isNaN(cred) ? 0 : Math.abs(cred)) - (isNaN(deb) ? 0 : Math.abs(deb));
      if (isNaN(deb) && isNaN(cred)) amt = NaN;
    }
    if (!d || isNaN(amt) || !amt) { skipped++; return; }
    const hm = (c.time !== undefined ? String(r[c.time] || '').slice(0, 5) : '') || d.hm || '12:00';
    const desc = c.desc !== undefined ? String(r[c.desc] || '').trim() : '', cents = Math.round(amt * 100);
    const base = [accName, d.key, hm, cents, desc].join('|');
    seen[base] = (seen[base] || 0) + 1;
    const time = Math.floor(Utilities.parseDate(d.key + ' ' + (/^\d{2}:\d{2}$/.test(hm) ? hm : '12:00') + ':00', tz, 'yyyy-MM-dd HH:mm:ss').getTime() / 1000);
    items.push({ id: 'csv:' + digest(base + '|' + seen[base]), time: time, amount: cents, description: desc,
      comment: c.note !== undefined ? String(r[c.note] || '').trim() : '', mcc: c.mcc !== undefined ? Number(r[c.mcc]) || 0 : 0, hold: false });
  });
  return { items: items, skipped: skipped, header: hd };
}

function md5Short_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').slice(0, 20);
}

/** Чиста функція: текст CSV → рядки (роздільник за першими рядками). */
function csvRows_(text) {
  const t = String(text || '').replace(/^﻿/, '');
  const head = t.split(/\r?\n/).filter(l => l.trim()).slice(0, 5).join('\n');
  const count = ch => (head.match(new RegExp(ch === '\t' ? '\t' : '\\' + ch, 'g')) || []).length;
  const delim = [';', ',', '\t'].sort((a, b) => count(b) - count(a))[0];
  return Utilities.parseCsv(t, delim);
}

/** Рядки файла з папки: CSV (UTF-8 або Windows-1251) чи Google Таблиця (лише читання). null: формат не підтримується. */
function stmtFileRows_(file) {
  const mime = file.getMimeType();
  if (mime === SHEETS_MIME) {
    const other = SpreadsheetApp.openById(file.getId());
    if (other.getSheetByName('m') && other.getSheetByName('y')) return { err: 'це схоже на особисту таблицю (аркуші «m» і «y»): не читаю' };
    return { rows: other.getSheets()[0].getDataRange().getValues() };
  }
  if (/csv|text\/plain|text\/comma/.test(mime) || /\.csv$/i.test(file.getName())) {
    let text = file.getBlob().getDataAsString('UTF-8');
    if (text.indexOf('�') >= 0) text = file.getBlob().getDataAsString('windows-1251');
    return { rows: csvRows_(text) };
  }
  if (/sheet|excel|xlsx|xls/i.test(mime) || /\.xlsx?$/i.test(file.getName()))
    return { err: 'Excel: відкрийте файл у Google Таблицях → «Файл → Зберегти як Google Таблиці» (або збережіть виписку як CSV)' };
  return { err: 'невідомий формат (' + mime + ')' };
}

/** Меню: імпорт нових файлів з папки «виписки» на обраний рахунок. */
function importStatements() {
  if (!menuOk_('📥 Виписки')) return;
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi(), dp = PropertiesService.getDocumentProperties(), tz = ss.getSpreadsheetTimeZone();
  const folder = stmtFolder_(ss), done = jsonProp_(dp, 'stmt_done') || {}, files = [];
  const it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); const stamp = String(f.getLastUpdated ? f.getLastUpdated().getTime() : 0); if (done[f.getId()] !== stamp) files.push({ f: f, stamp: stamp }); }
  if (!files.length) {
    show_('📥 Виписки', ['Нових файлів немає. Покладіть виписку (CSV або Google Таблицю) в папку ' + link_(folder.getUrl(), STMT_FOLDER) + ' і запустіть цей пункт знову.']);
    return;
  }
  const accs = accList_(ss).filter(a => a.bank === 'Виписка');
  const q = ui.prompt('📥 На який рахунок?', 'Нових файлів: ' + files.length + ' (' + files.map(x => x.f.getName()).join(', ').slice(0, 300) + ').\n\n' +
    (accs.length ? accs.map((a, i) => (i + 1) + ') ' + a.name + ' (' + a.kind + ')').join('\n') + '\n\nНомер, або ' : '') +
    'нова назва рахунку (наприклад, «Ощад картка»):', ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() !== ui.Button.OK) return;
  const ans = String(q.getResponseText() || '').trim();
  let acc = accs[Number(ans) - 1];
  if (!acc) {
    if (!ans) { show_('📥 Виписки', ['Потрібна назва рахунку.']); return; }
    const who = askWho_('📥 Виписки', 'це рахунок');
    if (who === null) return;
    acc = accEnsure_(ss, { key: 'csv:' + md5Short_(ans), name: ans, bank: 'Виписка', who: who });   // немає основної картки: ця стає основною
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) { show_('📥 Виписки', ['Інша дія скрипта ще триває, спробуйте за хвилину.']); return; }
  const lines = [], dateOf = x => Utilities.formatDate(new Date(x.time * 1000), tz, 'yyyy-MM-dd'), timeOf = x => Utilities.formatDate(new Date(x.time * 1000), tz, 'dd.MM.yyyy HH:mm');
  let total = 0;
  try {
    files.forEach(x => {
      const got = stmtFileRows_(x.f);
      if (got.err) { lines.push('⚠️ ' + esc_(x.f.getName()) + ': ' + esc_(got.err)); done[x.f.getId()] = x.stamp; return; }
      const p = parseStatement_(got.rows, tz, acc.name, md5Short_);
      if (!p.header) { lines.push('⚠️ ' + esc_(x.f.getName()) + ': не знайшов заголовка з датою й сумою.'); done[x.f.getId()] = x.stamp; return; }
      const res = bankIngest_(ss, { acc: acc, items: p.items, dateOf: dateOf, timeOf: timeOf, src: 'csv' });
      total += res.written;
      lines.push('✓ ' + esc_(x.f.getName()) + ': операцій ' + p.items.length + ', нових чернеток ' + res.written + ', ваших рядків знайдено ' + res.plan.marks.length +
        (res.plan.stats.other ? ', уже були з банку ' + res.plan.stats.other : '') + (p.skipped ? ', пропущено рядків без дати чи суми ' + p.skipped : '') + '.');
      done[x.f.getId()] = x.stamp;
    });
    const keys = Object.keys(done);
    keys.slice(0, Math.max(0, keys.length - 200)).forEach(k => delete done[k]);
    dp.setProperty('stmt_done', JSON.stringify(done));
    sortCore_(ss);
  } finally { lock.releaseLock(); }
  show_('📥 Виписки → «' + acc.name + '»', lines.concat(['Разом нових чернеток: <b>' + total + '</b>. Підтвердіть їх у журналі («✅ факт»).']));
}
