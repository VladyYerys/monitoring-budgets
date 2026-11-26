// «Весь код одним файлом»: усі файли скрипта в порядку завантаження + маніфест + готовий запит для перевірки через ШІ.
// Результат: dist/monitoring-budgets-all-code.txt (у git не йде; у GitHub Actions: артефакт кожного запуску).
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..');
const order = JSON.parse(fs.readFileSync(path.join(ROOT, '.clasp.json'), 'utf8')).filePushOrder;
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));

const PROMPT = [
  'Нижче весь код шаблону Google Таблиць «Monitoring Budgets» (Google Apps Script). Перевір його як аудитор безпеки:',
  '1) перелічи всі місця, де код зберігає чи передає токени, ключі, паролі, і куди саме;',
  '2) перелічи всі зовнішні адреси (URL), до яких код звертається;',
  '3) чи надсилає код щось на сервери автора або третіх осіб (аналітика, логи)?',
  '4) які дозволи Google потрібні цьому коду і навіщо кожен;',
  '5) чи може код змінити іншу таблицю, крім тієї, в якій він працює?',
  'Відповідай коротко, з назвами файлів і функцій.',
].join('\n');

/** Код одним файлом для Apps Script: файли в порядку завантаження (як у clasp push). */
function codeFile() {
  return ['// Monitoring Budgets ' + pkg.version + ': весь код одним файлом для Apps Script (github.com/VladyYerys/monitoring-budgets, MIT).',
    '// Вставте замість вмісту Code.gs у порожній таблиці, збережіть, оновіть таблицю: меню «📊 Monitoring Budgets → 🚀 Почати».', '']
    .concat(order.map(f => '// ===== ' + f + ' =====\n' + fs.readFileSync(path.join(ROOT, f), 'utf8'))).join('\n');
}
module.exports = { codeFile };
if (require.main !== module) return;

const parts = ['# Monitoring Budgets ' + pkg.version + ': весь код одним файлом', '', PROMPT, ''];
order.concat(['src/appsscript.json']).forEach(f => {
  parts.push('='.repeat(78), '# ' + f, '='.repeat(78), fs.readFileSync(path.join(ROOT, f), 'utf8'));
});
fs.mkdirSync(path.join(ROOT, 'dist'), { recursive: true });
const out = path.join(ROOT, 'dist', 'monitoring-budgets-all-code.txt');
fs.writeFileSync(out, parts.join('\n'));
console.log('✓ ' + path.relative(ROOT, out) + ' (' + order.length + ' файлів)');

// Той самий код одним файлом для Apps Script (без запиту для ШІ): можна вставити в «Розширення → Apps Script → Code.gs»
// замість заливки через GitHub Actions. Порядок файлів той самий, тож результат такий самий.
const code = codeFile();
const codeOut = path.join(ROOT, 'dist', 'monitoring-budgets-Code.gs');
fs.writeFileSync(codeOut, code);
console.log('✓ ' + path.relative(ROOT, codeOut) + ' (вставити в Apps Script одним файлом)');
