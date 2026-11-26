// Запобіжник перед заливкою: код шаблону йде лише в порожній скрипт або в скрипт, де вже є шаблон Monitoring Budgets.
// Порожній: лише appsscript.json і, можливо, стандартний Code з порожньою функцією. Шаблон: є мітка TEMPLATE_MARK.
// Будь-що інше (наприклад, скрипт вашої особистої таблиці) → зупинка, заливки не буде.
// Використання: node ci/guard.js <папка з кодом, стягнутим clasp pull>
const fs = require('fs'), path = require('path');
const MARK = 'monitoring-budgets-template';

function files(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap(e => {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) return files(p);
    return e.name === '.clasp.json' ? [] : [p];
  });
}

function check(dir) {
  const list = files(dir), code = list.filter(f => !/appsscript\.json$/.test(f));
  const text = f => fs.readFileSync(f, 'utf8');
  if (code.some(f => text(f).indexOf(MARK) >= 0)) return { ok: true, why: 'у скрипті вже шаблон Monitoring Budgets' };
  const isDefault = f => /^\s*function\s+myFunction\s*\(\s*\)\s*\{\s*\}\s*$/.test(text(f).replace(/\/\/.*$/gm, ''));
  if (code.length === 0 || (code.length === 1 && isDefault(code[0]))) return { ok: true, why: 'скрипт порожній (перша заливка)' };
  return { ok: false, why: 'у скрипті ' + code.length + ' файл(ів) чужого коду без мітки шаблону. Схоже, SCRIPT_ID вказує не на таблицю-шаблон ' +
    '(можливо, на вашу особисту). Заливку зупинено: перевірте секрет SCRIPT_ID.' };
}

if (require.main === module) {
  const dir = process.argv[2] || 'remote-before';
  const r = check(dir);
  console.log((r.ok ? '✓ ' : '✗ ') + r.why);
  process.exit(r.ok ? 0 : 1);
}
module.exports = { check, MARK };
