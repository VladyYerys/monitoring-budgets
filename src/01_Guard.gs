/** 01_Guard.gs: Запобіжник. Шаблон не змінює жодну таблицю, крім порожньої або вже позначеної як шаблон. */

/*
 * Навіщо: якщо цей код випадково потрапить в іншу таблицю (наприклад, особисту «Monitoring Budget»), він має
 * нічого не зламати. Тому кожна дія, що щось пише (меню, розклади, onEdit), спершу перевіряє таблицю:
 *   1) таблиця з аркушами «m» і «y» (особиста модель) → відмова, завжди, навіть якщо є мітка;
 *   2) таблиця позначена як шаблон → можна. Мітка двох видів: властивість документа mbs = TEMPLATE_MARK
 *      і прихована клітинка на «Налаштуваннях» (її копіює «Файл → Зробити копію», тож копія шаблону теж позначена);
 *   3) порожня таблиця (лише один порожній аркуш) → можна, і вона стає шаблоном (лише для «🚀 Почати» і збирання шаблону);
 *   4) будь-яка інша → відмова.
 * Розклади шаблону мають власні назви (mbsTick, mbsImport…), тож на розклади іншого коду шаблон не відповідає.
 * Якщо таблиця не підходить, розклад просто нічого не робить (і нічого не видаляє).
 */

function isPersonalModel_(ss) { return !!(ss.getSheetByName('m') && ss.getSheetByName('y')); }

function isEmptySpreadsheet_(ss) {
  const sh = ss.getSheets();
  return sh.length === 1 && sh[0].getLastRow() === 0 && sh[0].getLastColumn() === 0;
}

function hasSheetMark_(ss) {
  const s = ss.getSheetByName(SETTINGS_SHEET);
  try { return !!s && String(s.getRange(SET_MARK_CELL).getValue()) === TEMPLATE_MARK; } catch (e) { return false; }
}

/** Кидає помилку, якщо в цій таблиці шаблонові щось змінювати не можна. allowEmpty: можна позначити порожню. */
function guard_(ss, allowEmpty) {
  if (isPersonalModel_(ss)) throw new Error('Це схоже на особисту таблицю (є аркуші «m» і «y»). Шаблон Monitoring Budgets тут нічого не змінює.');
  const dp = PropertiesService.getDocumentProperties();
  if (dp.getProperty('mbs') === TEMPLATE_MARK) return true;
  if (hasSheetMark_(ss)) { dp.setProperty('mbs', TEMPLATE_MARK); return true; }      // копія шаблону: властивість відновлюємо
  if (allowEmpty && isEmptySpreadsheet_(ss)) { dp.setProperty('mbs', TEMPLATE_MARK); return true; }
  throw new Error('Ця таблиця не шаблон Monitoring Budgets. Відкрийте порожню таблицю або копію шаблону.');
}

/** Для розкладів і onEdit: true, якщо можна працювати; інакше тихо false (без вікон, без змін). */
function quietOk_() {
  try { return guard_(SpreadsheetApp.getActive()) === true; } catch (e) { return false; }
}

/** Для пунктів меню: true або вікно з причиною і false. */
function menuOk_(title) {
  try { return guard_(SpreadsheetApp.getActive()) === true; }
  catch (e) { try { show_(title, ['⚠️ ' + esc_(e.message)]); } catch (x) { /* без інтерфейсу */ } return false; }
}
