// Перевірка приватності всього репозиторію. Падає (код 1), якщо знайде:
//   токени й ключі (Telegram, Google API, OAuth, приватні ключі), номери карток (з перевіркою Луна), IBAN, телефони,
//   пошти (крім noreply і example), схожі на Script ID / ID файлів Google рядки,
//   зовнішні адреси в коді, яких немає в PRIVACY.md,
//   і слова з вашого локального списку .private-denylist.txt (якщо він є; файл лише на вашому комп'ютері, у git не йде).
// Знайдене друкується замаскованим: сам секрет у журнал CI не потрапляє.
const fs = require('fs'), path = require('path'), { execSync } = require('child_process');
const ROOT = path.join(__dirname, '..');
const SKIP_DIRS = new Set(['.git', 'node_modules', 'dist', 'remote-before']);
const SKIP_FILES = new Set(['.private-denylist.txt', 'package-lock.json']);
const ALLOWED_EMAIL = /@(users\.noreply\.github\.com|example\.(com|org))$/i;

function walk(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap(e => {
    if (SKIP_DIRS.has(e.name)) return [];
    const p = path.join(dir, e.name);
    if (e.isDirectory()) return walk(p);
    return SKIP_FILES.has(e.name) ? [] : [p];
  });
}

function luhn(d) {
  let s = 0, alt = false;
  for (let i = d.length - 1; i >= 0; i--) { let n = +d[i]; if (alt) { n *= 2; if (n > 9) n -= 9; } s += n; alt = !alt; }
  return s % 10 === 0;
}

const mask = s => (s.length <= 6 ? '***' : s.slice(0, 3) + '…' + s.slice(-2) + ' (' + s.length + ' симв.)');

const RULES = [
  { name: 'токен Telegram-бота', re: /\b\d{8,10}:[A-Za-z0-9_-]{35}\b/g },
  { name: 'ключ Google API', re: /AIza[0-9A-Za-z_-]{35}/g },
  { name: 'OAuth / вхід clasp', re: /"(refresh_token|client_secret|access_token)"\s*:/g },
  { name: 'приватний ключ', re: /-----BEGIN [A-Z ]*PRIVATE KEY-----/g },
  { name: 'IBAN', re: /\bUA\d{27}\b/g },
  { name: 'телефон', re: /\+380\d{9}\b/g },
  { name: 'схоже на Script ID чи ID файлу Google', re: /\b1[A-Za-z0-9_-]{40,}\b/g },
  { name: 'пошта', re: /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, ok: m => ALLOWED_EMAIL.test(m) },
  { name: 'номер картки', re: /\b(?:\d[ -]?){12,18}\d\b/g, ok: m => { const d = m.replace(/\D/g, ''); return d.length < 13 || d.length > 19 || !luhn(d); } },
];

function denylist() {
  const f = path.join(ROOT, '.private-denylist.txt');
  if (!fs.existsSync(f)) return [];
  return fs.readFileSync(f, 'utf8').split('\n').map(s => s.trim()).filter(s => s && s[0] !== '#' && s.length >= 3);
}

function hostsInCode() {
  const out = new Map(), src = path.join(ROOT, 'src');
  walk(src).forEach(f => {
    const t = fs.readFileSync(f, 'utf8'), re = /https?:\/\/([A-Za-z0-9.-]+)/g;
    let m;
    while ((m = re.exec(t))) out.set(m[1].toLowerCase(), path.relative(ROOT, f));
  });
  return out;
}

function run() {
  const found = [], deny = denylist();
  walk(ROOT).forEach(f => {
    const rel = path.relative(ROOT, f);
    if (/\.(png|jpg|jpeg|gif|ico|pdf|zip)$/i.test(f)) return;
    const lines = fs.readFileSync(f, 'utf8').split('\n');
    lines.forEach((line, i) => {
      RULES.forEach(rule => {
        (line.match(rule.re) || []).forEach(m => { if (!rule.ok || !rule.ok(m)) found.push(rel + ':' + (i + 1) + ': ' + rule.name + ': ' + mask(m)); });
      });
      const low = line.toLowerCase();
      deny.forEach((w, k) => { if (low.indexOf(w.toLowerCase()) >= 0) found.push(rel + ':' + (i + 1) + ': слово № ' + (k + 1) + ' з вашого .private-denylist.txt'); });
    });
  });
  // зовнішні адреси в коді мають бути описані в PRIVACY.md
  const privacy = fs.existsSync(path.join(ROOT, 'PRIVACY.md')) ? fs.readFileSync(path.join(ROOT, 'PRIVACY.md'), 'utf8').toLowerCase() : '';
  hostsInCode().forEach((file, host) => { if (privacy.indexOf(host) < 0) found.push(file + ': адреса ' + host + ' не описана в PRIVACY.md'); });
  // локальний список слів не має потрапити в git
  try {
    const tracked = execSync('git ls-files', { cwd: ROOT, stdio: ['ignore', 'pipe', 'ignore'] }).toString().split('\n');
    if (tracked.indexOf('.private-denylist.txt') >= 0) found.push('.private-denylist.txt: файл у git! Приберіть його з репозиторію (git rm --cached).');
  } catch (e) { /* не git-репозиторій: пропускаємо */ }
  return { found, denyWords: deny.length };
}

if (require.main === module) {
  const r = run();
  if (r.found.length) { console.log('✗ Знайдено ' + r.found.length + ':\n' + r.found.join('\n')); process.exit(1); }
  console.log('✓ Приватність: нічого не знайдено' + (r.denyWords ? ' (ваш список: ' + r.denyWords + ' слів)' : '') + '.');
}
module.exports = { run, RULES, luhn, hostsInCode };
