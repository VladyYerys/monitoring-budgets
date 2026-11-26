// Навичка Gemini «Радник Monitoring Budgets»: SKILL.md з тексту в коді (src/39_Advisor.gs → advSkillMd_), щоб файл у репозиторії
// і вікно «🤖 Радник Gemini» в таблиці завжди збігались. Тест test_advisor.js перевіряє, що файл свіжий.
const fs = require('fs'), path = require('path'), vm = require('vm');
const ROOT = path.join(__dirname, '..');
const order = JSON.parse(fs.readFileSync(path.join(ROOT, '.clasp.json'), 'utf8')).filePushOrder;
const ctx = vm.createContext({});
order.forEach(f => vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f }));
const out = path.join(ROOT, 'gemini', 'monitoring-budgets-advisor', 'SKILL.md');
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, vm.runInContext('advSkillMd_()', ctx));
console.log('✓ ' + path.relative(ROOT, out));
