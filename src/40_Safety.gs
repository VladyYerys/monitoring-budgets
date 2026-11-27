/** 40_Safety.gs: Нічні копії таблиці на Диску і захист аркушів «з попередженням». */

/*
 * Копії: щоночі (03:00…04:00) і вручну, у папку «Monitoring Budgets: копії» поруч із таблицею.
 * Зберігаються останні 30 нічних і 10 ручних; старіші йдуть у кошик Диска (звідти їх можна повернути 30 днів).
 * Скрипт прибирає лише власні копії ЦІЄЇ таблиці (назва «<назва таблиці>: копія … (авто|ручна)») у своїй папці.
 * Захист: аркуші з формулами захищені «з попередженням» (Google спитає «Ви впевнені?»), клітинки для вводу вільні.
 * Ваш власний захист скрипт не чіпає.
 */

const PROT_TAG = '🛡 ' + APP_NAME;
const BACKUP_FOLDER = APP_NAME + ': копії';
const KEEP = { 'авто': 30, 'ручна': 10 };
const BACKUP_HOUR = 3;
const COPY_MARK = ': копія ';
const SHEETS_MIME = 'application/vnd.google-apps.spreadsheet';

/** Що захищаємо: аркуш і вільні для вводу діапазони. */
function protectList_() {
  return [
    { sheet: JOURNAL, allow: ['A' + J_FIRST + ':H' + J_LAST] },
    { sheet: MONTHS_SHEET, allow: [M_YEAR] },
    { sheet: DAYS_SHEET, allow: [] },
    { sheet: MAIN_SHEET, allow: [] },
    { sheet: CHECK_SHEET, allow: [] },
    { sheet: HISTORY_SHEET, allow: [] },
    { sheet: H_CAT_SHEET, allow: [] },
    { sheet: BANK_SHEET, allow: [] },
    { sheet: START_SHEET, allow: [] },                                     // кроки пише скрипт; модулі: «🚀 Підключення», крок 5
    { sheet: REG_SHEET, allow: ['A' + REG_FIRST + ':I' + REG_LAST] },
    { sheet: CH_SHEET, allow: [] },
    { sheet: DASH_SHEET, allow: [colStr_(DASH_LIMIT_COL) + (DASH_HOLE_HEAD + 1) + ':' + colStr_(DASH_LIMIT_COL) + (DASH_HOLE_HEAD + CH_HOLES)] },
    { sheet: MOB_SHEET, allow: [] },
  ];
}

function isOurs_(p) { return String(p.getDescription() || '').indexOf(PROT_TAG) === 0; }

function applyProtection_(ss) {
  const done = [], skipped = [];
  protectList_().forEach(cfg => {
    const sh = ss.getSheetByName(cfg.sheet);
    if (!sh) return;
    const existing = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (existing.some(p => !isOurs_(p))) { skipped.push(cfg.sheet + ': уже має ваш власний захист, не чіпаю'); return; }
    const ours = existing.filter(isOurs_);
    ours.slice(1).forEach(p => p.remove());
    const p = ours.length ? ours[0] : sh.protect();
    p.setDescription(PROT_TAG + ' · ' + cfg.sheet);
    p.setWarningOnly(true);
    p.setUnprotectedRanges(cfg.allow.map(a => sh.getRange(a)));
    done.push(cfg.sheet);
  });
  PropertiesService.getDocumentProperties().setProperty('prot', 'on');
  return { done: done, skipped: skipped };
}

function protectionOn_(ss) {
  return protectList_().every(cfg => { const sh = ss.getSheetByName(cfg.sheet); return !sh || sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).some(isOurs_); });
}

function protectionOnMenu() {
  if (!menuOk_('🛡 Захист')) return;
  const r = applyProtection_(SpreadsheetApp.getActive());
  show_('🛡 Захист з попередженням', ['Захищено: ' + esc_(r.done.join(', ') || 'нічого') + '.', r.skipped.length ? 'Пропущено: ' + esc_(r.skipped.join('; ')) : '',
    'Клітинки для вводу (журнал A…H, рік на «Місяцях», «Довідник», «Правила», «Налаштування», «Залишки банку») вільні.']);
}

/** Зняти лише наш захист (ваш власний лишається). → скільки знято. */
function protectionRemove_(ss) {
  let n = 0;
  ss.getSheets().forEach(sh => [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE].forEach(t =>
    sh.getProtections(t).filter(isOurs_).forEach(p => { p.remove(); n++; })));
  PropertiesService.getDocumentProperties().setProperty('prot', 'off');
  return n;
}

function removeProtection() {
  if (!menuOk_('🛡 Захист')) return;
  const n = protectionRemove_(SpreadsheetApp.getActive());
  show_('🛡 Захист прибрано', ['Знято захистів: <b>' + n + '</b>. Ваш власний захист (якщо є) не чіпався.',
    'Повернути: «🚀 Підключення → Інше → Захист: увімкнути».']);
}

/* ---------- копії ---------- */

function ensureBackupTrigger_(ss) {
  removeTriggers_(TRG.backup);
  ScriptApp.newTrigger(TRG.backup).timeBased().atHour(BACKUP_HOUR).everyDays(1).inTimezone(ss.getSpreadsheetTimeZone()).create();
}

function mbsBackup() {
  if (!quietOk_() || !modOn_('backup')) return;
  try { backup_(SpreadsheetApp.getActive(), 'авто'); } catch (e) { beat_('backup', false, e && e.message); throw e; }
  beat_('backup', true);
}

function backupNow() {
  if (!menuOk_('💾 Копія')) return;
  let r;
  try { r = backup_(SpreadsheetApp.getActive(), 'ручна'); } catch (e) { show_('💾 Копія', ['⚠️ ' + esc_(e.message)]); return; }
  show_('💾 Копію зроблено', [link_(r.url, r.name), 'Папка: ' + link_(r.folderUrl, BACKUP_FOLDER), r.trashed ? 'Найстаріших ручних копій у кошик: ' + r.trashed : '']);
}

function backup_(ss, kind) {
  guard_(ss);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) throw new Error('Інша дія скрипта ще триває, спробуйте за хвилину.');
  const dp = PropertiesService.getDocumentProperties();
  try {
    const folder = backupFolder_(ss);
    const name = ss.getName() + COPY_MARK + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm') + ' (' + kind + ')';
    const copy = DriveApp.getFileById(ss.getId()).makeCopy(name, folder);
    const trashed = rotate_(folder, ss.getName(), kind);
    const now = new Date().toISOString(), upd = { backup_last_ok: now, backup_last_name: name, backup_last_url: copy.getUrl() };
    upd['backup_last_' + (kind === 'авто' ? 'auto' : 'manual')] = now;
    dp.setProperties(upd);
    return { name: name, url: copy.getUrl(), folderUrl: folder.getUrl(), trashed: trashed };
  } catch (e) {
    dp.setProperties({ backup_last_error: String(e && e.message || e), backup_last_error_at: new Date().toISOString() });
    throw e;
  } finally { lock.releaseLock(); }
}

/** Наша папка: за збереженим id; інакше нова поруч із таблицею (чужу папку з такою ж назвою не беремо). */
function backupFolder_(ss) {
  const dp = PropertiesService.getDocumentProperties(), id = dp.getProperty('backup_folder_id');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* видалили */ } }
  let parent;
  try { const it = DriveApp.getFileById(ss.getId()).getParents(); parent = it.hasNext() ? it.next() : DriveApp.getRootFolder(); }
  catch (e) { parent = DriveApp.getRootFolder(); }
  let folder;
  try { folder = parent.createFolder(BACKUP_FOLDER); } catch (e) { folder = DriveApp.getRootFolder().createFolder(BACKUP_FOLDER); }
  dp.setProperty('backup_folder_id', folder.getId());
  return folder;
}

function rotate_(folder, base, kind) {
  const extra = copiesOf_(listCopies_(folder), base, kind).slice(KEEP[kind]);
  extra.forEach(f => f.file.setTrashed(true));
  return extra.length;
}

function listCopies_(folder) {
  const out = [], it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); out.push({ file: f, name: f.getName(), mime: f.getMimeType(), created: f.getDateCreated().getTime() }); }
  return out;
}

/** Чиста функція: копії цієї таблиці певного виду, новіші спершу. */
function copiesOf_(files, base, kind) {
  const tail = '(' + kind + ')', head = base + COPY_MARK;
  return files.filter(f => f.mime === SHEETS_MIME && f.name.indexOf(head) === 0 && f.name.slice(-tail.length) === tail).sort((a, b) => b.created - a.created);
}

function backupStatusLines_(ss) {
  const p = PropertiesService.getDocumentProperties().getProperties(), tz = ss.getSpreadsheetTimeZone();
  const fmt = iso => iso ? Utilities.formatDate(new Date(iso), tz, 'dd.MM.yyyy HH:mm') : 'ще не було';
  const lines = ['Нічні копії: ' + (!modOn_('backup') ? '<b>модуль вимкнено</b> (' + MOD_WHERE + ')' : hasTrigger_(TRG.backup) ? '<b>увімкнено</b> (щодня ' + pad2_(BACKUP_HOUR) + ':00…' + pad2_(BACKUP_HOUR + 1) + ':00)' : '<b>вимкнено</b>'),
    'Остання успішна копія: <b>' + fmt(p.backup_last_ok) + '</b>'];
  if (p.backup_folder_id) {
    try {
      const folder = DriveApp.getFolderById(p.backup_folder_id), files = listCopies_(folder);
      lines.push('Папка: ' + link_(folder.getUrl(), BACKUP_FOLDER) + ' · нічних ' + copiesOf_(files, ss.getName(), 'авто').length + ' з ' + KEEP['авто'] +
        ', ручних ' + copiesOf_(files, ss.getName(), 'ручна').length + ' з ' + KEEP['ручна']);
    } catch (e) { lines.push('⚠️ Папку з копіями не знайдено: наступна копія створить нову.'); }
  }
  if (p.backup_last_error_at && (!p.backup_last_ok || p.backup_last_error_at > p.backup_last_ok)) lines.push('⚠️ Остання спроба з помилкою (' + fmt(p.backup_last_error_at) + '): ' + esc_(p.backup_last_error || ''));
  return lines;
}
