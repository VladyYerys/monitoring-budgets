// Усі тести по черзі; код виходу 1, якщо хоч один упав (так GitHub Actions не пустить зламаний код у таблицю).
const { spawnSync } = require('child_process'), fs = require('fs'), path = require('path');
const files = fs.readdirSync(__dirname).filter(f => /^test_.*\.js$/.test(f)).sort();
let failed = 0;
for (const f of files) {
  const r = spawnSync(process.execPath, [path.join(__dirname, f)], { encoding: 'utf8' });
  const last = (r.stdout || '').trim().split('\n').pop();
  if (r.status === 0) console.log('✓ ' + f + ': ' + last);
  else { failed++; console.log('✗ ' + f + '\n' + (r.stdout || '').slice(-2000) + (r.stderr || '').slice(-4000)); }
}
console.log(failed ? '\n' + failed + ' з ' + files.length + ' наборів упали' : '\nУсі ' + files.length + ' наборів пройшли');
process.exit(failed ? 1 : 0);
