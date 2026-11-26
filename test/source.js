// Код таблиці з src/ у тому ж порядку, що й у Apps Script (.clasp.json → filePushOrder).
// load(ctx) виконує файли по одному, як Apps Script: файл бачить лише те, що завантажилось раніше.
const fs = require('fs'), path = require('path'), vm = require('vm');
const ROOT = path.join(__dirname, '..');
const ORDER = JSON.parse(fs.readFileSync(path.join(ROOT, '.clasp.json'), 'utf8')).filePushOrder;
const FILES = ORDER.map(f => ({ name: f, text: fs.readFileSync(path.join(ROOT, f), 'utf8') }));
const SRC = FILES.map(f => f.text).join('\n');
function load(ctx) { FILES.forEach(f => vm.runInContext(f.text, ctx, { filename: f.name })); }
module.exports = { ORDER, FILES, SRC, load };
