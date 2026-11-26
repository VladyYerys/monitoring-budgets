// Готує clasp у GitHub Actions: вхід із секрету CLASPRC_B64, Script ID із секрету SCRIPT_ID (у репозиторії його немає).
// Нічого не друкує з секретів.
const fs = require('fs'), os = require('os'), path = require('path');
const b64 = process.env.CLASPRC_B64, scriptId = String(process.env.SCRIPT_ID || '').trim();
if (!b64 || !scriptId) { console.error('Немає секретів CLASPRC_B64 і SCRIPT_ID.'); process.exit(1); }
if (!/^[A-Za-z0-9_-]{30,}$/.test(scriptId)) { console.error('SCRIPT_ID не схожий на Script ID.'); process.exit(1); }
const creds = Buffer.from(b64.trim(), 'base64url').toString('utf8');
try { JSON.parse(creds); } catch (e) { console.error('Секрет CLASPRC_B64 пошкоджений: це не вхід clasp. Створіть його ще раз.'); process.exit(1); }
fs.mkdirSync(os.homedir(), { recursive: true });
fs.writeFileSync(path.join(os.homedir(), '.clasprc.json'), creds, { mode: 0o600 });
const conf = JSON.parse(fs.readFileSync('.clasp.json', 'utf8'));
conf.scriptId = scriptId;
fs.writeFileSync('.clasp.json', JSON.stringify(conf, null, 2));            // лише в робочій копії CI, не комітиться
fs.mkdirSync('remote-before', { recursive: true });
fs.writeFileSync(path.join('remote-before', '.clasp.json'), JSON.stringify({ scriptId, rootDir: '.' }));
console.log('clasp готовий.');
